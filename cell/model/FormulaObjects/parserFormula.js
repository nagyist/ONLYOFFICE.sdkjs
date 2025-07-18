/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(
/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {
  // Import
  var CellValueType = AscCommon.CellValueType;
  var cBoolLocal = AscCommon.cBoolLocal;
  var cBoolOrigin = AscCommon.cBoolOrigin;
  var cErrorOrigin = AscCommon.cErrorOrigin;
  var cErrorLocal = AscCommon.cErrorLocal;
  var FormulaSeparators = AscCommon.FormulaSeparators;
  var parserHelp = AscCommon.parserHelp;
  var g_oFormatParser = AscCommon.g_oFormatParser;
  var CellAddress = AscCommon.CellAddress;
  var cDate = Asc.cDate;
  var bIsSupportArrayFormula = true;
  var bIsSupportDynamicArrays = false;

  var c_oAscError = Asc.c_oAscError;

	var TOK_TYPE_OPERAND = 1;
	var TOK_TYPE_FUNCTION = 2;
	var TOK_TYPE_SUBEXPR = 3;
	var TOK_TYPE_ARGUMENT = 4;
	var TOK_TYPE_OP_IN = 5;
	var TOK_TYPE_OP_POST = 6;
	var TOK_TYPE_WSPACE = 7;
	var TOK_TYPE_UNKNOWN = 8;

	var TOK_SUBTYPE_START = 9;
	var TOK_SUBTYPE_STOP = 10;

	var TOK_SUBTYPE_TEXT = 11;
	var TOK_SUBTYPE_LOGICAL = 12;
	var TOK_SUBTYPE_ERROR = 14;

	var TOK_SUBTYPE_UNION = 15;

	var arrayFunctionsMap = {"SUMPRODUCT": 1, "FILTER": 1, "SUM": 1, "LOOKUP": 1, "AGGREGATE": 1};

	var importRangeLinksState = {importRangeLinks: null, startBuildImportRangeLinks: null};
	const aExcludeRecursiveFormulas = ['ISFORMULA', 'SHEETS', 'AREAS', 'COLUMN', 'COLUMNS', 'ROW', 'ROWS', 'CELL', 'OFFSET'];

	const cReplaceFormulaType = {
		val: 1,
		formula: 2
	};

	function getArrayCopy(arr) {
		var newArray = [];
		for (var i = 0; i < arr.length; i++) {
			newArray[i] = arr[i].slice();
		}
		return newArray
	}

	function generate3DLink(externalPath, sheet, range) {
		let filePrefix = "file:///";
		if (externalPath && 0 === externalPath.indexOf(filePrefix)) {
			//"file:///C:\root\from1.xlsx"
			sheet = sheet.split(":");
			var wsFrom = sheet[0], wsTo = sheet[1] === undefined ? wsFrom : sheet[1];
			wsFrom = wsFrom.replace(/'/g, "''");
			wsTo = wsTo.replace(/'/g, "''");
			return "'" + externalPath + (wsFrom !== wsTo ? wsFrom + ":" + wsTo : wsFrom) + "'!" + range;
		} else {
			if (!externalPath) {
				externalPath = "";
			}
			return parserHelp.get3DRef(externalPath + sheet, range);
		}
	}
	
	function ParsedThing(value, type, subtype, pos, length) {
		this.value = value;
		this.type = type;
		this.subtype = subtype;
		this.pos = pos;
		this.length = length;
	}

	ParsedThing.prototype.getStop = function () {
		return new ParsedThing(this.value, this.type, TOK_SUBTYPE_STOP, this.pos, this.length);
	};

	var g_oCodeSpace = 32; // Code of space
	var g_oCodeNumberSign = 35; // Code of #
	var g_oCodeDQuote = 34; // Code of "
	var g_oCodePercent = 37; // Code of %
	var g_oCodeAmpersand = 38; // Code of &
	var g_oCodeQuote = 39; // Code of '
	var g_oCodeLeftParenthesis = 40; // Code of (
	var g_oCodeRightParenthesis = 41; // Code of )
	var g_oCodeMultiply = 42; // Code of *
	var g_oCodePlus = 43; // Code of +
	var g_oCodeComma = 44; // Code of ,
	var g_oCodeMinus = 45; // Code of -
	var g_oCodeDivision = 47; // Code of /
	var g_oCodeSemicolon = 59; // Code of ;
	var g_oCodeLessSign = 60; // Code of <
	var g_oCodeEqualSign = 61; // Code of =
	var g_oCodeGreaterSign = 62; // Code of >
	var g_oCodeLeftSquareBracked = 91; // Code of [
	var g_oCodeRightSquareBracked = 93; // Code of ]
	var g_oCodeAccent = 94; // Code of ^
	var g_oCodeLeftCurlyBracked = 123; // Code of {
	var g_oCodeRightCurlyBracked = 125; // Code of }

	function getTokens(formula) {

		var tokens = [];
		var tokenStack = [];

		var offset = 0;
		var length = formula.length;
		var currentChar, currentCharCode, nextCharCode, tmp;

		var token = "";

		var inString = false;
		var inPath = false;
		var inRange = false;
		var inError = false;

		var regexSN = /^[1-9]{1}(\.[0-9]+)?E{1}$/;

		nextCharCode = formula.charCodeAt(offset);
		while (offset < length) {

			// state-dependent character evaluation (order is important)

			// double-quoted strings
			// embeds are doubled
			// end marks token

			currentChar = formula[offset];
			currentCharCode = nextCharCode;
			nextCharCode = formula.charCodeAt(offset + 1);

			if (inString) {
				if (currentCharCode === g_oCodeDQuote) {
					if (nextCharCode === g_oCodeDQuote) {
						token += currentChar;
						offset += 1;
					} else {
						inString = false;
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_TEXT, offset, token.length));
						token = "";
					}
				} else {
					token += currentChar;
				}
				offset += 1;
				continue;
			} else if (inPath) {
				// single-quoted strings (links)
				// embeds are double
				// end does not mark a token
				if (currentCharCode === g_oCodeQuote) {
					if (nextCharCode === g_oCodeQuote) {
						token += currentChar;
						offset += 1;
					} else {
						inPath = false;
					}
				} else {
					token += currentChar;
				}
				offset += 1;
				continue;
			} else if (inRange) {
				// bracked strings (range offset or linked workbook name)
				// no embeds (changed to "()" by Excel)
				// end does not mark a token
				if (currentCharCode === g_oCodeRightSquareBracked) {
					inRange = false;
				}
				token += currentChar;
				offset += 1;
				continue;
			} else if (inError) {
				// error values
				// end marks a token, determined from absolute list of values
				token += currentChar;
				offset += 1;
				if ((",#NULL!,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#N/A,").indexOf("," + token + ",") != -1) {
					inError = false;
					tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_ERROR, offset, token.length));
					token = "";
				}
				continue;
			}

			// trim white-space
			if (currentCharCode === g_oCodeSpace) {
				if (token.length > 0) {
					tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
					token = "";
				}
				tokens.push(new ParsedThing("", TOK_TYPE_WSPACE, null, offset, token.length));
				offset += 1;

				while ((currentCharCode = formula.charCodeAt(offset)) === g_oCodeSpace) {
					offset += 1;
				}
				if (offset >= length) {
					break;
				}

				currentChar = formula[offset];
				nextCharCode = formula.charCodeAt(offset + 1);
			}

			// multi-character comparators (>= || <= || <>)
			if ((currentCharCode === g_oCodeLessSign &&
				(nextCharCode === g_oCodeEqualSign || nextCharCode === g_oCodeGreaterSign)) ||
				(currentCharCode === g_oCodeGreaterSign && nextCharCode === g_oCodeEqualSign)) {
				if (token.length > 0) {
					tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
					token = "";
				}
				tokens.push(new ParsedThing(formula.substr(offset, 2), TOK_TYPE_OP_IN, TOK_SUBTYPE_LOGICAL, offset, token.length));
				offset += 2;
				nextCharCode = formula.charCodeAt(offset);
				continue;
			}

			// scientific notation check
			if (currentCharCode === g_oCodePlus || currentCharCode === g_oCodeMinus) {
				if (token.length > 1) {
					if (token.match(regexSN)) {
						token += currentChar;
						offset += 1;
						continue;
					}
				}
			}

			// independent character evaulation (order not important)

			// establish state-dependent character evaluations
			switch (currentCharCode) {
				case g_oCodeDQuote:
				{
					if (token.length > 0) {
						// not expected
						tokens.push(new ParsedThing(token, TOK_TYPE_UNKNOWN, null, offset, token.length));
						token = "";
					}
					inString = true;
					break;
				}
				case g_oCodeQuote:
				{
					if (token.length > 0) {
						// not expected
						tokens.push(new ParsedThing(token, TOK_TYPE_UNKNOWN, null, offset, token.length));
						token = "";
					}
					inPath = true;
					break;
				}
				case g_oCodeLeftSquareBracked:
				{
					inRange = true;
					token += currentChar;
					break;
				}
				case g_oCodeNumberSign:
				{
					if (token.length > 0) {
						// not expected
						tokens.push(new ParsedThing(token, TOK_TYPE_UNKNOWN, null, offset, token.length));
						token = "";
					}
					inError = true;
					token += currentChar;
					break;
				}
				case g_oCodeLeftCurlyBracked:
				{
					// mark start and end of arrays and array rows
					if (token.length > 0) {
						// not expected
						tokens.push(new ParsedThing(token, TOK_TYPE_UNKNOWN, null, offset, token.length));
						token = "";
					}
					tmp = new ParsedThing('ARRAY', TOK_TYPE_FUNCTION, TOK_SUBTYPE_START, offset, token.length);
					tokens.push(tmp);
					tokenStack.push(tmp.getStop());
					tmp = new ParsedThing('ARRAYROW', TOK_TYPE_FUNCTION, TOK_SUBTYPE_START, offset, token.length);
					tokens.push(tmp);
					tokenStack.push(tmp.getStop());
					break;
				}
				case g_oCodeSemicolon:
				{
					if (token.length > 0) {
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
						token = "";
					}
					tmp = tokenStack.pop();
					if (tmp && 'ARRAYROW' !== tmp.value) {
						return null;
					}
					tokens.push(tmp);
					tokens.push(new ParsedThing(';', TOK_TYPE_ARGUMENT, null, offset, token.length));
					tmp = new ParsedThing('ARRAYROW', TOK_TYPE_FUNCTION, TOK_SUBTYPE_START, offset, token.length);
					tokens.push(tmp);
					tokenStack.push(tmp.getStop());
					break;
				}
				case g_oCodeRightCurlyBracked:
				{
					if (token.length > 0) {
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
						token = "";
					}
					tokens.push(tokenStack.pop());
					tokens.push(tokenStack.pop());
					break;
				}
				case g_oCodePlus:
				case g_oCodeMinus:
				case g_oCodeMultiply:
				case g_oCodeDivision:
				case g_oCodeAccent:
				case g_oCodeAmpersand:
				case g_oCodeEqualSign:
				case g_oCodeGreaterSign:
				case g_oCodeLessSign:
				{
					// standard infix operators
					if (token.length > 0) {
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
						token = "";
					}
					tokens.push(new ParsedThing(currentChar, TOK_TYPE_OP_IN, null, offset, token.length));
					break;
				}
				case g_oCodePercent:
				{
					// standard postfix operators
					if (token.length > 0) {
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
						token = "";
					}
					tokens.push(new ParsedThing(currentChar, TOK_TYPE_OP_POST, null, offset, token.length));
					break;
				}
				case g_oCodeLeftParenthesis:
				{
					// start subexpression or function
					if (token.length > 0) {
						tmp = new ParsedThing(token, TOK_TYPE_FUNCTION, TOK_SUBTYPE_START, offset, token.length);
						tokens.push(tmp);
						tokenStack.push(tmp.getStop());
						token = "";
					} else {
						tmp = new ParsedThing("", TOK_TYPE_SUBEXPR, TOK_SUBTYPE_START, offset, token.length);
						tokens.push(tmp);
						tokenStack.push(tmp.getStop());
					}
					break;
				}
				case g_oCodeComma:
				{
					// function, subexpression, array parameters
					if (token.length > 0) {
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
						token = "";
					}
					tmp = (0 !== tokenStack.length) ? (TOK_TYPE_FUNCTION === tokenStack[tokenStack.length - 1].type) : false;
					tokens.push(tmp ? new ParsedThing(currentChar, TOK_TYPE_ARGUMENT, null, offset, token.length) :
						new ParsedThing(currentChar, TOK_TYPE_OP_IN, TOK_SUBTYPE_UNION, offset, token.length));
					break;
				}
				case g_oCodeRightParenthesis:
				{
					// stop subexpression
					if (token.length > 0) {
						tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
						token = "";
					}
					if(tokenStack.length) {
						tokens.push(tokenStack.pop());
					}
					break;
				}
				default:
				{
					// token accumulation
					token += currentChar;
					break;
				}
			}

			++offset;
		}

		// dump remaining accumulation
		if (token.length > 0) {
			tokens.push(new ParsedThing(token, TOK_TYPE_OPERAND, null, offset, token.length));
		}

		return tokens;
	}

	function prepareTypedArrayUniversal(array, lookingElem, isByRangeCall) {
		const typedArr = [];

		for (let i = 0; i < array.length; i++) {
			let arrayElemValue = isByRangeCall ? array[i].v : array[i];
			// let elemType = isByRangeCall ? array[i].v.type : array[i].type;
			let elemType = arrayElemValue.type;
			let elemIndex = isByRangeCall ? array[i].i : i;

			if (lookingElem.type === cElementType.bool) {
				// return only bool
				if (lookingElem.type !== elemType) {
					continue
				}
				typedArr.push({i: elemIndex, v: arrayElemValue});
			} else if (lookingElem.type === cElementType.number) {
				// return only numbers or string.tocNumber
				if (elemType !== cElementType.string && elemType !== cElementType.number) {
					continue
				}
				let temp = arrayElemValue.tocNumber();
				if (temp.type !== cElementType.error) {
					typedArr.push({i: elemIndex, v: temp});
				}
			} else if (cElementType.string === lookingElem.type) {
				// return only strings
				if (lookingElem.type !== elemType) {
					continue
				}
				typedArr.push({i: elemIndex, v: new cString(arrayElemValue.toString().toLowerCase())});
			}
		}

		return typedArr;
	}


/** @enum */
var cElementType = {
		number      : 0,
		string      : 1,
		bool        : 2,
		error       : 3,
		empty       : 4,
		cellsRange  : 5,
		cell        : 6,
		date        : 7,
		func        : 8,
		operator    : 9,
		name        : 10,
		array       : 11,
		cell3D      : 12,
		cellsRange3D: 13,
		table       : 14,
		name3D      : 15,
		specialFunctionStart: 16,
		specialFunctionEnd  : 17,
		pivotTable  : 18

  };
/** @enum */
var cErrorType = {
		unsupported_function: 0,
		null_value          : 1,
		division_by_zero    : 2,
		wrong_value_type    : 3,
		bad_reference       : 4,
		wrong_name          : 5,
		not_numeric         : 6,
		not_available       : 7,
		getting_data        : 8,
		array_not_calc      : 9,
		cannot_be_spilled	: 10,
		busy                : 11
  };
//добавляю константу cReturnFormulaType для корректной обработки формул массива
// value - функция умеет возвращать только значение(не массив)
// в этом случае данная функция вызывается множество раз для каждого элемента внутренних массивов
// предварительно area и area3d преобразуются в массив
// value_convert_area - аналогично value, но area и area3d не преобразуются в массив
// array - умеет возвращать массив
// используоется в returnValueType у каждой формулы
// так же этот параметр у формул может быть массивом - массив индексов аргментов, которые являются входными array/area
// area_to_ref - заменяем area на массив ссылок на ячейку(REF)
// replace_only_array - в случае с Area - оставляем его в аргументах и рассчитываем только 1 значение(аналогично array)
// replace_only_array - в слуае с массивом - обрабатываем стандартно по элементам
// dynamic_array - в отличие от обычного массива такой тип будут использовать формулы которые могут не иметь в аргументах диапазонов/массивов, но при этом будут их возвращать(прим. SEQUENCE)

/** @enum */
var cReturnFormulaType = {
	value: 0,
	value_replace_area: 1,
	array: 2,
	area_to_ref: 3,
	replace_only_array: 4,
	setArrayRefAsArg: 5, //для row/column если нет аргументов
	dynamic_array: 6
};

/*
	arrayIndexesType - an supporting structure that shows what type of data (associated only with arrays) we expect to see in the argument
	There are functions whose arguments only work with a certain type
	For example, the argument can process an array, but a range with the same data will not be processed (as in the WORKDAY_INTL function for the first two arguments)
	Each type is used to obtain that type in its pure form in a function, for example:
	If the data type is array, we process only ranges, and pass the arrays unchanged into the formula
	If the data type is range, we process only arrays, and pass ranges unchanged to the formula
	If the data type is any - any data type is passed unchanged to the formula (analogous to the previous use of arrayIndex in the format {0: 1, 1: 1})
*/

/** @enum */
const arrayIndexesType = {
	array: 0,
	any: 1,
	range: 2,
};

var cExcelSignificantDigits = 15; //количество цифр в числе после запятой
var cExcelMaxExponent = 308;
var cExcelMinExponent = -308;
var c_Date1904Const = 24107; //разница в днях между 01.01.1970 и 01.01.1904 годами
var c_Date1900Const = 25568; //разница в днях между 01.01.1970 и 01.01.1900 годами
var rx_sFuncPref = /_xlfn\./i;
var rx_sFuncPrefXlWS = /_xlws\./i;// /_xlfn\.(_xlws\.)?/i;
var rx_sDefNamePref = /_xlnm\./i;
var rx_sFuncPrefXLUFD = /__xludf.DUMMYFUNCTION\./i;
var cNumFormatFirstCell = -1;
var cNumFormatNone = -2;
var cNumFormatNull = -3;
var g_nFormulaStringMaxLength = 255;
var c_nMaxDate1900 = 2958465;
var c_nMaxDate1904 = c_nMaxDate1900 - (c_Date1900Const - c_Date1904Const) + 1;

function getMaxDate () {
	return AscCommon.bDate1904 ? c_nMaxDate1904 : c_nMaxDate1900; 	// Maximum date used in calculations in ms (equivalent 31/12/9999)
}

let fIsPromise = function (val) {
	return val && val.promise && val.promise.then;
};

// set type weight of base types
let cElementTypeWeight =  new Map();
	cElementTypeWeight.set(cElementType.number, 0);
	cElementTypeWeight.set(cElementType.empty, 0);
	cElementTypeWeight.set(cElementType.string, 1);
	cElementTypeWeight.set(cElementType.bool, 2);
	cElementTypeWeight.set(cElementType.error, 3);




Math.fmod = function ( a, b ) {
	return Number( (a - (this.floor( a / b ) * b)).toPrecision( cExcelSignificantDigits ) );
};



parserHelp.setDigitSeparator(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);

	/** @constructor */
	function cBaseType(val) {
		this.numFormat = cNumFormatNull;
		this.value = val;
		this.hyperlink = null;
	}

	cBaseType.prototype.cloneTo = function (oRes) {
		oRes.numFormat = this.numFormat;
		oRes.value = this.value;
		oRes.hyperlink = this.hyperlink;
	};
	cBaseType.prototype.getValue = function () {
		return this.value;
	};
	cBaseType.prototype.getHyperlink = function () {
		return this.hyperlink;
	};
	cBaseType.prototype.toString = function () {
		return this.value.toString();
	};
	cBaseType.prototype.toLocaleString = function () {
		return this.toString();
	};
	cBaseType.prototype.toLocaleStringObj = function () {
		var localStr = this.toLocaleString();
		var localStrWithoutSheet;
		if(localStr) {
			var result = parserHelp.parse3DRef(localStr);
			if (result) {
				localStrWithoutSheet = result.range;
			} else {
				localStrWithoutSheet = localStr;
			}
		} else {
			localStr = this.value;
			localStrWithoutSheet = this.value;
		}
		return [localStr, localStrWithoutSheet];
	};
	cBaseType.prototype.getDimensions = function () {
		return {col: 1, row: 1};
	};
	cBaseType.prototype.isOneElement = function () {
		let dimensions = this.getDimensions();
		if (dimensions.col === 1 && dimensions.row === 1) {
			return true;
		}
		return false;
	};
	cBaseType.prototype.getExternalLinkStr = function (externalLink, locale, isShortLink) {
		let wb = Asc.editor && Asc.editor.wbModel;
		if (!wb) {
			return "";
		}
		return wb && wb.externalReferenceHelper && wb.externalReferenceHelper.getExternalLinkStr(externalLink, locale, isShortLink);
	};

	cBaseType.prototype.toArray = function (putValue, checkOnError, fPrepareElem, bSaveBoolean) {
		let arr = [];
		if (this.getMatrix) {
			arr = this.getMatrix();
			arr = getArrayCopy(arr);

			if (putValue || checkOnError) {
				for (let i = 0; i < arr.length; i++) {
					if (arr[i]) {
						for (let j = 0; j < arr[i].length; j++) {
							if (checkOnError) {
								if (arr[i][j].type === cElementType.error) {
									return arr[i][j];
								}
							}
							if (fPrepareElem) {
								arr[i][j] = fPrepareElem(arr[i][j]);
								if (checkOnError) {
									if (arr[i][j].type === cElementType.error) {
										return arr[i][j];
									}
								}
							}
							if (putValue) {
								if (bSaveBoolean && arr[i][j].type === cElementType.bool) {
									arr[i][j] = arr[i][j].toBool();
								} else {
									arr[i][j] = arr[i][j].getValue();
								}
							}
						}
					}
				}
			}
		} else {
			if (checkOnError) {
				if (this.type === cElementType.error) {
					return this;
				}
			}
			let _res = fPrepareElem ? fPrepareElem(this) : this;
			if (fPrepareElem && checkOnError) {
				if (this.type === cElementType.error) {
					return this;
				}
			}
			if (!arr[0]) {
				arr[0] = [];
			}
			arr[0][0] = putValue ? _res.getValue() : _res;
		}
		return arr;
	};

	/*Basic types of an elements used into formulas*/
	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cNumber(val) {
		cBaseType.call(this, parseFloat(val));
		var res;

		if (!isNaN(this.value) && Math.abs(this.value) !== Infinity) {
			res = this;
		} else if (val instanceof cError) {
			res = val;
		} else {
			res = new cError(cErrorType.not_numeric);
		}
		return res;
	}

	cNumber.prototype = Object.create(cBaseType.prototype);
	cNumber.prototype.constructor = cNumber;
	cNumber.prototype.type = cElementType.number;
	cNumber.prototype.tocString = function () {
		return new cString(("" + this.value).replace(FormulaSeparators.digitSeparatorDef,
			FormulaSeparators.digitSeparator));
	};
	cNumber.prototype.tocNumber = function () {
		return this;
	};
	cNumber.prototype.toNumber = function () {
		return this.value;
	};
	cNumber.prototype.tocBool = function () {
		return new cBool(this.value !== 0);
	};
	cNumber.prototype.toLocaleString = function (digitDelim) {
		var res = this.value.toString();
		if (digitDelim) {
			return res.replace(FormulaSeparators.digitSeparatorDef, FormulaSeparators.digitSeparator);
		} else {
			return res;
		}
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cString(val) {
		cBaseType.call(this, val);
	}

	cString.prototype = Object.create(cBaseType.prototype);
	cString.prototype.constructor = cString;
	cString.prototype.type = cElementType.string;
	cString.prototype.tocNumber = function (doNotParseNum) {
		var res, m = this.value;
		if (this.value === "") {
			res = new cNumber(0);
		}

		/*if ( this.value[0] === '"' && this.value[this.value.length - 1] === '"' ) {
		 m = this.value.substring( 1, this.value.length - 1 );
		 }*/

		if (g_oFormatParser.isLocaleNumber(this.value)) {
			var numberValue = g_oFormatParser.parseLocaleNumber(this.value);
			if (!isNaN(numberValue)) {
				res = new cNumber(numberValue);
			}
		} else {
			var parseRes = !doNotParseNum ? AscCommon.g_oFormatParser.parse(this.value) : null;
			if (null != parseRes) {
				res = new cNumber(parseRes.value);
			} else {
				res = new cError(cErrorType.wrong_value_type);
			}
		}

		return res;
	};
	cString.prototype.tocBool = function () {
		//TODO value === cBoolLocal.t || value === cBoolLocal.f || value === cBoolOrigin.t || value === cBoolOrigin.f
		var res;
		if (parserHelp.isBoolean(this.value, 0)) {
			res = new cBool(parserHelp.operand_str.toUpperCase() === cBoolLocal.t);
		} else {
			res = this;
		}
		return res;
	};
	cString.prototype.tocString = function () {
		return this;
	};
	cString.prototype.getValue = function (doNotReplace) {
		//TODO many function calls -> many calls indexOf/replaceAll - review and if only necessary to do the conversion
		if (!doNotReplace && -1 !== this.value.indexOf("\"\"")) {
			return this.value.replaceAll("\"\"", "\"");
		}
		return this.value;
	};
	cString.prototype.toString = function () {
		return this.value;
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cBool(val) {
		var v = false;
		if (val === true) {
			v = true;
		} else if (val === false) {
			v = false;
		} else {
			switch (val.toString().toUpperCase()) {
				case "TRUE":
				case cBoolLocal.t:
					v = true;
			}
		}

		cBaseType.call(this, v);
	}

	cBool.prototype = Object.create(cBaseType.prototype);
	cBool.prototype.constructor = cBool;
	cBool.prototype.type = cElementType.bool;
	cBool.prototype.toString = function () {
		return this.value ? cBoolOrigin.t : cBoolOrigin.f;
	};
	cBool.prototype.getValue = function () {
		return this.toString();
	};
	cBool.prototype.tocNumber = function () {
		return new cNumber(this.value ? 1.0 : 0.0);
	};
	cBool.prototype.tocString = function () {
		return new cString(this.value ? "TRUE" : "FALSE");
	};
	cBool.prototype.toLocaleString = function () {
		return this.value ? cBoolLocal.t : cBoolLocal.f;
	};
	cBool.prototype.tocBool = function () {
		return this;
	};
	cBool.prototype.toBool = function () {
		return this.value;
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cError(val) {
		cBaseType.call(this, val);

		this.errorType = -1;

		switch (val) {
			case cErrorLocal["value"]:
			case cErrorOrigin["value"]:
			case cErrorType.wrong_value_type: {
				this.value = "#VALUE!";
				this.errorType = cErrorType.wrong_value_type;
				break;
			}
			case cErrorLocal["nil"]:
			case cErrorOrigin["nil"]:
			case cErrorType.null_value: {
				this.value = "#NULL!";
				this.errorType = cErrorType.null_value;
				break;
			}
			case cErrorLocal["div"]:
			case cErrorOrigin["div"]:
			case cErrorType.division_by_zero: {
				this.value = "#DIV/0!";
				this.errorType = cErrorType.division_by_zero;
				break;
			}
			case cErrorLocal["ref"]:
			case cErrorOrigin["ref"]:
			case cErrorType.bad_reference: {
				this.value = "#REF!";
				this.errorType = cErrorType.bad_reference;
				break;
			}
			case cErrorLocal["name"]:
			case cErrorOrigin["name"]:
			case cErrorType.wrong_name: {
				this.value = "#NAME?";
				this.errorType = cErrorType.wrong_name;
				break;
			}
			case cErrorLocal["num"]:
			case cErrorOrigin["num"]:
			case cErrorType.not_numeric: {
				this.value = "#NUM!";
				this.errorType = cErrorType.not_numeric;
				break;
			}
			case cErrorLocal["na"]:
			case cErrorOrigin["na"]:
			case cErrorType.not_available: {
				this.value = "#N/A";
				this.errorType = cErrorType.not_available;
				break;
			}
			case cErrorLocal["getdata"]:
			case cErrorOrigin["getdata"]:
			case cErrorType.getting_data: {
				this.value = "#GETTING_DATA";
				this.errorType = cErrorType.getting_data;
				break;
			}
			case cErrorLocal["uf"]:
			case cErrorOrigin["uf"]:
			case cErrorType.unsupported_function: {
				this.value = "#UNSUPPORTED_FUNCTION!";
				this.errorType = cErrorType.unsupported_function;
				break;
			}
			case cErrorLocal["calc"]:
			case cErrorOrigin["calc"]:
			case cErrorType.array_not_calc: {
				this.value = "#CALC!";
				this.errorType = cErrorType.array_not_calc;
				break;
			}
			case cErrorLocal["spill"]:
			case cErrorOrigin["spill"]:
			case cErrorType.cannot_be_spilled: {
				this.value = "#SPILL!";
				this.errorType = cErrorType.cannot_be_spilled;
				break;
			}
			case cErrorLocal["busy"]:
			case cErrorOrigin["busy"]:
			case cErrorType.busy: {
				this.value = "#BUSY!";
				this.errorType = cErrorType.busy;
				break;
			}
		}

		return this;
	}

	cError.prototype = Object.create(cBaseType.prototype);
	cError.prototype.constructor = cError;
	cError.prototype.type = cElementType.error;
	cError.prototype.tocNumber = cError.prototype.tocString = cError.prototype.tocBool = function () {
		return this;
	};
	cError.prototype.toLocaleString = function () {
		var val = this.value ? this.value.toUpperCase() : this.value;
		switch (val) {
			case cErrorOrigin["value"]:
			case cErrorType.wrong_value_type: {
				return cErrorLocal["value"];
			}
			case cErrorOrigin["nil"]:
			case cErrorType.null_value: {
				return cErrorLocal["nil"];
			}
			case cErrorOrigin["div"]:
			case cErrorType.division_by_zero: {
				return cErrorLocal["div"];
			}

			case cErrorOrigin["ref"]:
			case cErrorType.bad_reference: {
				return cErrorLocal["ref"];
			}

			case cErrorOrigin["name"]:
			case cErrorType.wrong_name: {
				return cErrorLocal["name"];
			}

			case cErrorOrigin["num"]:
			case cErrorType.not_numeric: {
				return cErrorLocal["num"];
			}

			case cErrorOrigin["na"]:
			case cErrorType.not_available: {
				return cErrorLocal["na"];
			}

			case cErrorOrigin["getdata"]:
			case cErrorType.getting_data: {
				return cErrorLocal["getdata"];
			}

			case cErrorOrigin["uf"]:
			case cErrorType.unsupported_function: {
				return cErrorLocal["uf"];
			}

			case cErrorOrigin["calc"]:
			case cErrorType.array_not_calc: {
				return cErrorLocal["calc"];
			}
			case cErrorOrigin["spill"]:
			case cErrorType.cannot_be_spilled: {
				return cErrorLocal["spill"];
			}
			case cErrorOrigin["busy"]:
			case cErrorType.busy: {
				return cErrorLocal["busy"];
			}
		}
		return cErrorLocal["na"];
	};
	cError.prototype.getErrorTypeFromString = function(val) {
		var res;
		switch (val) {
			case cErrorOrigin["value"]: {
				res = cErrorType.wrong_value_type;
				break;
			}
			case cErrorOrigin["nil"]: {
				res = cErrorType.null_value;
				break;
			}
			case cErrorOrigin["div"]: {
				res = cErrorType.division_by_zero;
				break;
			}
			case cErrorOrigin["ref"]: {
				res = cErrorType.bad_reference;
				break;
			}
			case cErrorOrigin["name"]: {
				res = cErrorType.wrong_name;
				break;
			}
			case cErrorOrigin["num"]: {
				res = cErrorType.not_numeric;
				break;
			}
			case cErrorOrigin["na"]: {
				res = cErrorType.not_available;
				break;
			}
			case cErrorOrigin["getdata"]: {
				res = cErrorType.getting_data;
				break;
			}
			case cErrorOrigin["uf"]: {
				res = cErrorType.unsupported_function;
				break;
			}
			case cErrorOrigin["calc"]: {
				res = cErrorType.array_not_calc;
				break;
			}
			case cErrorOrigin["spill"]: {
				res = cErrorType.cannot_be_spilled;
				break;
			}
			case cErrorOrigin["busy"]: {
				res = cErrorType.busy;
				break;
			}
			default: {
				res = cErrorType.not_available;
				break;
			}
		}
		return res;
	};
	cError.prototype.getStringFromErrorType = function(type) {
		var res;
		switch (type) {
			case cErrorType.wrong_value_type: {
				res = cErrorOrigin["value"];
				break;
			}
			case cErrorType.null_value: {
				res = cErrorOrigin["nil"];
				break;
			}
			case cErrorType.division_by_zero: {
				res = cErrorOrigin["div"];
				break;
			}
			case cErrorType.bad_reference: {
				res = cErrorOrigin["ref"];
				break;
			}
			case cErrorType.wrong_name: {
				res = cErrorOrigin["name"];
				break;
			}
			case cErrorType.not_numeric: {
				res = cErrorOrigin["num"];
				break;
			}
			case cErrorType.not_available: {
				res = cErrorOrigin["na"];
				break;
			}
			case cErrorType.getting_data: {
				res = cErrorOrigin["getdata"];
				break;
			}
			case cErrorType.unsupported_function: {
				res = cErrorOrigin["uf"];
				break;
			}
			case cErrorType.array_not_calc: {
				res = cErrorOrigin["calc"];
				break;
			}
			case cErrorType.cannot_be_spilled: {
				res = cErrorOrigin["spill"];
				break;
			}
			case cErrorType.busy: {
				res = cErrorOrigin["busy"];
				break;
			}
			default:
				res = cErrorType.not_available;
				break;
		}
		return res;
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cArea(val, ws) {/*Area means "A1:E5" for example*/
		cBaseType.call(this, val);

		this.ws = ws;
		this.range = null;
		if (val) {
			AscCommonExcel.executeInR1C1Mode(false, function () {
				val = ws.getRange2(val);
			});
			this.range = val;
		}
	}

	cArea.prototype = Object.create(cBaseType.prototype);
	cArea.prototype.constructor = cArea;
	cArea.prototype.type = cElementType.cellsRange;
	cArea.prototype.clone = function (opt_ws) {
		var ws = opt_ws ? opt_ws : this.ws;
		var oRes = new cArea(null, ws);
		this.cloneTo(oRes);
		if (this.range) {
			oRes.range = this.range.clone(ws);
		}
		return oRes;
	};
	cArea.prototype.getWsId = function () {
		return this.ws.Id;
	};
	cArea.prototype.getValue = function (checkExclude, excludeHiddenRows, excludeErrorsVal, excludeNestedStAg) {
		var val = [], r = this.getRange();
		if (!r) {
			val.push(new cError(cErrorType.bad_reference));
		} else {
			if (checkExclude && !excludeHiddenRows) {
				excludeHiddenRows = this.ws.isApplyFilterBySheet();
			}
			r._foreachNoEmpty(function (cell) {
				if(!(excludeNestedStAg && cell.formulaParsed && cell.formulaParsed.isFoundNestedStAg())){
					var checkTypeVal = checkTypeCell(cell);
					if(!(excludeErrorsVal && CellValueType.Error === checkTypeVal.type)){
						val.push(checkTypeVal);
					}
				}

			}, undefined, excludeHiddenRows);
		}
		return val;
	};
	cArea.prototype.getValue2 = function (i, j) {
		var res = this.index(i + 1, j + 1), r;
		if (!res) {
			r = this.getRange();
			r.worksheet._getCellNoEmpty(r.bbox.r1 + i, r.bbox.c1 + j, function(cell) {
				res = checkTypeCell(cell);
			});
		}
		return res;
	};
	cArea.prototype.getValueByRowCol = function (i, j, checkEmpty) {
		let res, r;
		r = this.getRange();
		r.worksheet._getCellNoEmpty(r.bbox.r1 + i, r.bbox.c1 + j, function(cell) {
			if(cell) {
				res = checkTypeCell(cell);
			}
		});
		if (checkEmpty && res == null) {
			res = new cEmpty();
		}
		return res;
	};
	cArea.prototype.getRange = function () {
		if (!this.range) {
			this.range = this.ws.getRange2(this.value);
		}
		return this.range;
	};
	cArea.prototype.tocNumber = function () {
		var v = this.getValue()[0];
		if (!v) {
			v = new cNumber(0);
		} else {
			v = v.tocNumber();
		}
		return v;
	};
	cArea.prototype.tocString = function () {
		let val = this.getValue()[0];
		if (!val) {
			return new cString("");
		}
		return val.tocString();
	};
	cArea.prototype.tocBool = function () {
		return new cError(cErrorType.wrong_value_type);
	};
	cArea.prototype.to3D = function (opt_ws) {
		opt_ws = opt_ws || this.ws;
		var res = new cArea3D(null, opt_ws, opt_ws);
		this.cloneTo(res);
		if (this.range) {
			res.bbox = this.range.getBBox0().clone();
		}
		return res;
	};
	cArea.prototype.toString = function () {
		var _c;

		if (AscCommonExcel.g_ProcessShared && this.range) {
			_c = this.range.getName();
		} else {
			_c = this.value;
		}

		if (_c.indexOf(":") < 0) {
			_c = _c + ":" + _c;
		}
		return _c;
	};
	cArea.prototype.toLocaleString = function () {
		var _c;

		if (this.range) {
			_c = this.range.getName();
		} else {
			_c = this.value;
		}
		if (_c.indexOf(":") < 0) {
			_c = _c + ":" + _c;
		}
		return _c;
	};
	cArea.prototype.getWS = function () {
		return this.ws;
	};
	cArea.prototype.getBBox0 = function () {
		return this.getRange().getBBox0();
	};
	cArea.prototype.cross = function (arg) {
		var r = this.getRange(), cross;
		if (!r) {
			return new cError(cErrorType.wrong_name);
		}
		cross = r.cross(arg);
		if (cross) {
			if (undefined !== cross.r) {
				return this.getValue2(cross.r - this.getBBox0().r1, 0);
			} else if (undefined !== cross.c) {
				return this.getValue2(0, cross.c - this.getBBox0().c1);
			}
		}
		return new cError(cErrorType.wrong_value_type);
	};
	cArea.prototype.isValid = function () {
		return !!this.getRange();
	};
	cArea.prototype.countCells = function () {
		var r = this.getRange(), bbox = r.bbox, count = (Math.abs(bbox.c1 - bbox.c2) + 1) *
			(Math.abs(bbox.r1 - bbox.r2) + 1);
		r._foreachNoEmpty(function (cell) {
			if (!cell || !cell.isEmptyTextString()) {
				count--;
			}
		});
		return new cNumber(count);
	};
	cArea.prototype.foreach = function (action) {
		var r = this.getRange();
		if (r) {
			r._foreach2(action);
		}
	};
	cArea.prototype.foreach2 = function (action) {
		var r = this.getRange();
		if (r) {
			r._foreach2(function (cell, row, col) {
				action(checkTypeCell(cell), cell, row, col);
			});
		}
	};
	cArea.prototype.getMatrix = function (excludeHiddenRows, excludeErrorsVal, excludeNestedStAg) {
		var arr = [], r = this.getRange();

		var ws = r.worksheet;
		var oldExcludeHiddenRows = ws.bExcludeHiddenRows;
		ws.bExcludeHiddenRows = false;
		r._foreach2(function (cell, i, j, r1, c1) {
			if (!arr[i - r1]) {
				arr[i - r1] = [];
			}

			var resValue = new cEmpty();
			if(!(excludeNestedStAg && cell.formulaParsed && cell.formulaParsed.isFoundNestedStAg())){
				var checkTypeVal = checkTypeCell(cell);
				if(!(excludeErrorsVal && CellValueType.Error === checkTypeVal.type)){
					resValue = checkTypeVal;
				}
			}

			arr[i - r1][j - c1] = resValue;
		});
		ws.bExcludeHiddenRows = oldExcludeHiddenRows;

		return arr;
	};
	cArea.prototype.getFullArray = function (emptyReplaceOn, maxRowCount, maxColCount) {
		let arr = new cArray();
		let elemsNoEmpty = this.getMatrixNoEmpty();
		let bbox = this.getBBox0();
		if (!emptyReplaceOn) {
			emptyReplaceOn = new cEmpty();
		}
		for (let i = bbox.r1; i <= Math.min(bbox.r2, maxRowCount != null ? bbox.r1 + maxRowCount : bbox.r2); i++) {
			if ( !arr.array[i - bbox.r1] ) {
				arr.addRow();
			}
			for (let j = bbox.c1; j <= Math.min(bbox.c2, maxColCount != null ? bbox.c1 + maxColCount : bbox.c2); j++) {
				let elem = null;
				if (elemsNoEmpty && elemsNoEmpty[i - bbox.r1] && elemsNoEmpty[i - bbox.r1][j - bbox.c1]) {
					elem = elemsNoEmpty[i - bbox.r1][j - bbox.c1];
				}
				if (elem === null || elem.type === cElementType.empty) {
					elem = emptyReplaceOn;
				}

				arr.addElement(elem);
			}
		}

		return arr;
	};
	cArea.prototype.getMatrixNoEmpty = function () {
		var arr = [], r = this.getRange(), res;
		r._foreachNoEmpty(function (cell, i, j, r1, c1) {
			if (!arr[i - r1]) {
				arr[i - r1] = [];
			}

			arr[i - r1][j - c1] = checkTypeCell(cell);
		});
		return arr;
	};
	cArea.prototype.getValuesNoEmpty = function (checkExclude, excludeHiddenRows, excludeErrorsVal, excludeNestedStAg) {
		var arr = [], r = this.getRange();

		r._foreachNoEmpty(function (cell) {
			if(!(excludeNestedStAg && cell.formulaParsed && cell.formulaParsed.isFoundNestedStAg())){
				var checkTypeVal = checkTypeCell(cell);
				if(!(excludeErrorsVal && CellValueType.Error === checkTypeVal.type)){
					arr.push(checkTypeVal);
				}
			}

		}, undefined, excludeHiddenRows);

		return [arr];
	};
	cArea.prototype.index = function (r, c) {
		var bbox = this.getBBox0();
		bbox.normalize();
		var box = {c1: 1, c2: bbox.c2 - bbox.c1 + 1, r1: 1, r2: bbox.r2 - bbox.r1 + 1};

		if (r < box.r1 || r > box.r2 || c < box.c1 || c > box.c2) {
			return new cError(cErrorType.bad_reference);
		}
	};
	cArea.prototype.changeSheet = function (wsLast, wsNew) {
		if (this.ws === wsLast) {
			this.ws = wsNew;
			if (this.range) {
				this.range.worksheet = wsNew;
			}
		}
	};
	cArea.prototype.getDimensions = function () {
		var res = null;
		if (this.range && this.range.bbox) {
			var bbox = this.range.bbox;
			res =  {col: bbox.c2 - bbox.c1 + 1, row:  bbox.r2 - bbox.r1 + 1, bbox: bbox};
		}
		return res;
	};
	cArea.prototype.getFirstElement = function () {
		return this.getValueByRowCol(0, 0, true);
	};
	cArea.prototype._getCol = function (colIndex) {
		let dimensions = this.getDimensions();
		if (colIndex < 0 || colIndex > dimensions.col) {
			return null;
		}

		let col = [];
		for (let i = 0; i < dimensions.row; i++) {
			let elem = this.getValueByRowCol(i, colIndex);
			if (!elem) {
				elem = new cEmpty();
			}
			col[i] = [];
			col[i].push(elem);
		}
		return col;
	};
	cArea.prototype._getRow = function (rowIndex) {
		let dimensions = this.getDimensions();
		if (rowIndex < 0 || rowIndex > dimensions.row) {
			return null;
		}

		let row = [[]];
		for (let j = 0; j < this.getDimensions().col; j++) {
			let elem = this.getValueByRowCol(rowIndex, j);
			if (!elem) {
				elem = new cEmpty();
			}
			row[0].push(elem);
		}
		return row;
	};


	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cArea3D(val, wsFrom, wsTo, externalLink) {/*Area3D means "Sheat1!A1:E5" for example*/
		cBaseType.call(this, val);

		this.bbox = null;
		if (val) {
			AscCommonExcel.executeInR1C1Mode(false, function () {
				val = AscCommonExcel.g_oRangeCache.getAscRange(val);
			});
			if (val) {
				this.bbox = val.clone();
			}
		}
		this.wsFrom = wsFrom;
		this.wsTo = wsTo || this.wsFrom;
		this.externalLink = externalLink;
	}

	cArea3D.prototype = Object.create(cBaseType.prototype);
	cArea3D.prototype.constructor = cArea3D;
	cArea3D.prototype.type = cElementType.cellsRange3D;
	cArea3D.prototype.clone = function (opt_ws) {
		var oRes = new cArea3D(null, opt_ws ? opt_ws : this.wsFrom, opt_ws ? opt_ws : this.wsTo);
		this.cloneTo(oRes);
		if (this.bbox) {
			oRes.bbox = this.bbox.clone();
		}
		oRes.externalLink = this.externalLink;
		return oRes;
	};
	cArea3D.prototype.wsRange = function () {
		if (this.externalLink != null) {
			return [this.wsFrom];
		}
		var wb = this.wsFrom.workbook;
		var wsF = this.wsFrom.getIndex(), wsL = this.wsTo.getIndex(), r = [];
		for (var i = wsF; i <= wsL; i++) {
			r.push(wb.getWorksheet(i));
		}
		return r;
	};
	cArea3D.prototype.range = function (wsRange) {
		if (!wsRange) {
			return [null];
		}
		var r = [];
		for (var i = 0; i < wsRange.length; i++) {
			if (!wsRange[i]) {
				r.push(null);
			} else {
				r.push(AscCommonExcel.Range.prototype.createFromBBox(wsRange[i], this.bbox));
			}
		}
		return r;
	};
	cArea3D.prototype.getRange = function () {
		if (!this.isSingleSheet()) {
			return null;
		}
		return (this.range(this.wsRange()))[0];
	};
	cArea3D.prototype.getRanges = function () {
		return (this.range(this.wsRange()));
	};
	cArea3D.prototype.getValue = function (checkExclude, excludeHiddenRows, excludeErrorsVal, excludeNestedStAg) {
		var i, _wsA = this.wsRange();
		var _val = [];
		if (_wsA.length < 1) {
			_val.push(new cError(cErrorType.bad_reference));
			return _val;
		}
		for (i = 0; i < _wsA.length; i++) {
			if (!_wsA[i]) {
				_val.push(new cError(cErrorType.bad_reference));
				return _val;
			}

		}

		var _exclude;
		var _r = this.range(_wsA);
		for (i = 0; i < _r.length; i++) {
			if (!_r[i]) {
				_val.push(new cError(cErrorType.bad_reference));
				return _val;
			}
			if (checkExclude && !(_exclude = excludeHiddenRows)) {
				_exclude = _wsA[i].isApplyFilterBySheet();
			}

			_r[i]._foreachNoEmpty(function (cell) {
				if (!(excludeNestedStAg && cell.formulaParsed && cell.formulaParsed.isFoundNestedStAg())) {
					var checkTypeVal = checkTypeCell(cell);
					if (!(excludeErrorsVal && CellValueType.Error === checkTypeVal.type)) {
						_val.push(checkTypeVal);
					}
				}

			}, undefined, _exclude);
		}
		return _val;
	};
	cArea3D.prototype.getValue2 = function (cell) {
		var _wsA = this.wsRange(), _val = [], _r;
		if (_wsA.length < 1) {
			_val.push(new cError(cErrorType.bad_reference));
			return _val;
		}
		for (var i = 0; i < _wsA.length; i++) {
			if (!_wsA[i]) {
				_val.push(new cError(cErrorType.bad_reference));
				return _val;
			}

		}
		_r = this.range(_wsA);
		if (!_r[0]) {
			_val.push(new cError(cErrorType.bad_reference));
			return _val;
		}

		if (_r[0].worksheet) {
			_r[0].worksheet._getCellNoEmpty(cell.row - 1, cell.col - 1, function (_cell) {
				_val.push(checkTypeCell(_cell));
			});
		}

		return (null == _val[0]) ? new cEmpty() : _val[0];
	};
	cArea3D.prototype.getValueByRowCol = function (i, j, checkEmpty) {
		let r = this.getRanges(), res;

		if (r[0]) {
			r[0].worksheet._getCellNoEmpty(r[0].bbox.r1 + i, r[0].bbox.c1 + j, function (cell) {
				if (cell) {
					res = checkTypeCell(cell);
				}
			});
		}

		if (checkEmpty && res == null) {
			res = new cEmpty();
		}

		return res;
	};
	cArea3D.prototype.changeSheet = function (wsLast, wsNew) {
		if (this.wsFrom === wsLast) {
			this.wsFrom = wsNew;
		}
		if (this.wsTo === wsLast) {
			this.wsTo = wsNew;
		}
	};
	cArea3D.prototype.toString = function () {
		var wsFrom = this.wsFrom.getName();
		var wsTo = this.wsTo.getName();
		var name = AscCommonExcel.g_ProcessShared && this.bbox ? this.bbox.getName() : this.value;
		var exPath = this.getExternalLinkStr(this.externalLink);
		return parserHelp.get3DRef(wsFrom !== wsTo ? (exPath + wsFrom + ':' + wsTo) : (exPath + wsFrom), name);
	};
	cArea3D.prototype.toLocaleString = function () {
		var wsFrom = this.wsFrom.getName();
		var wsTo = this.wsTo.getName();
		var name = this.bbox ? this.bbox.getName() : this.value;
		var exPath = this.getExternalLinkStr(this.externalLink, true);
		return generate3DLink(exPath, wsFrom !== wsTo ? (wsFrom + ':' + wsTo) : wsFrom, name);
	};
	cArea3D.prototype.tocNumber = function () {
		return this.getValue()[0].tocNumber();
	};
	cArea3D.prototype.tocString = function () {
		let val = this.getValue()[0];
		if (!val) {
			return new cString("");
		}
		return val.tocString();
	};
	cArea3D.prototype.tocBool = function () {
		return new cError(cErrorType.wrong_value_type);
	};
	cArea3D.prototype.tocArea = function () {
		var wsR = this.wsRange();
		if (wsR.length === 1) {
			return new cArea(this.value, wsR[0]);
		}
		return false;
	};
	cArea3D.prototype.getWS = function () {
		return this.wsFrom;
	};
	cArea3D.prototype.getWsId = function () {
		return this.wsFrom && this.wsFrom.Id;
	};
	cArea3D.prototype.cross = function (arg, ws) {
		if (!this.isSingleSheet()) {
			return new cError(cErrorType.wrong_value_type);
		}
		/*if ( this.wsFrom !== ws ) {
		 return new cError( cErrorType.wrong_value_type );
		 }*/
		var r = this.getRange();
		if (!r) {
			return new cError(cErrorType.wrong_name);
		}
		var cross = r.cross(arg);
		if (cross) {
			if (undefined !== cross.r) {
				return this.getValue2(new CellAddress(cross.r, this.getBBox0().c1, 0));
			} else if (undefined !== cross.c) {
				return this.getValue2(new CellAddress(this.getBBox0().r1, cross.c, 0));
			}
		}
		return new cError(cErrorType.wrong_value_type);
	};
	cArea3D.prototype.getBBox0 = function () {
		var range = this.getRange();
		return range ? range.getBBox0() : range;
	};
	cArea3D.prototype.getBBox0NoCheck = function () {
		return this.bbox;
	};
	cArea3D.prototype.isValid = function () {
		var r = this.getRanges();
		for (var i = 0; i < r.length; ++i) {
			if (!r) {
				return false;
			}
		}
		return true;
	};
	cArea3D.prototype.countCells = function () {
		var _wsA = this.wsRange();
		var _val = [];
		if (_wsA.length < 1) {
			_val.push(new cError(cErrorType.bad_reference));
			return _val;
		}
		var i;
		for (i = 0; i < _wsA.length; i++) {
			if (!_wsA[i]) {
				_val.push(new cError(cErrorType.bad_reference));
				return _val;
			}

		}
		var _r = this.range(_wsA), bbox = _r[0].bbox, count = (Math.abs(bbox.c1 - bbox.c2) + 1) * (Math.abs(bbox.r1 - bbox.r2) + 1);
		count = _r.length * count;
		for (i = 0; i < _r.length; i++) {
			_r[i]._foreachNoEmpty(function (cell) {
				if (!cell || !cell.isEmptyTextString()) {
					count--;
				}
			});
		}
		return new cNumber(count);
	};
	cArea3D.prototype.getMatrix = function (excludeHiddenRows, excludeErrorsVal, excludeNestedStAg) {
		var arr = [], r = this.getRanges(), res;

		var ws = r[0] ? r[0].worksheet : null;
		if (ws) {
			var oldExcludeHiddenRows = ws.bExcludeHiddenRows;
			ws.bExcludeHiddenRows = false;
		}
		for (var k = 0; k < r.length; k++) {
			arr[k] = [];
			r[k]._foreach2(function (cell, i, j, r1, c1) {
				if (!arr[k][i - r1]) {
					arr[k][i - r1] = [];
				}

				var resValue = new cEmpty();
				if (!(excludeNestedStAg && cell.formulaParsed && cell.formulaParsed.isFoundNestedStAg())) {
					var checkTypeVal = checkTypeCell(cell);
					if (!(excludeErrorsVal && CellValueType.Error === checkTypeVal.type)) {
						resValue = checkTypeVal;
					}
				}

				arr[k][i - r1][j - c1] = resValue;
			});
		}
		return arr;
	};
	cArea3D.prototype.getMatrixAllRange = function () {
		var arr = [], r = this.getRanges(), res;
		for (var k = 0; k < r.length; k++) {
			arr[k] = [];
			r[k]._foreach(function (cell, i, j, r1, c1) {
				if (!arr[k][i - r1]) {
					arr[k][i - r1] = [];
				}
				res = checkTypeCell(cell);

				arr[k][i - r1][j - c1] = res;
			});
		}
		return arr;
	};
	cArea3D.prototype.getFullArray = function (emptyReplaceOn, maxRowCount, maxColCount) {
		let arr = new cArray();
		let elemsNoEmpty = this.getMatrixNoEmpty();
		let bbox = this.getBBox0();
		if (!emptyReplaceOn) {
			emptyReplaceOn = new cEmpty();
		}
		for (let i = bbox.r1; i <= Math.min(bbox.r2, maxRowCount != null ? bbox.r1 + maxRowCount : bbox.r2); i++) {
			if (!arr.array[i - bbox.r1]) {
				arr.addRow();
			}
			for (let j = bbox.c1; j <= Math.min(bbox.c2, maxColCount != null ? bbox.c1 + maxColCount : bbox.c2); j++) {
				let elem = null;
				if (elemsNoEmpty && elemsNoEmpty[0] && elemsNoEmpty[0][i - bbox.r1] && elemsNoEmpty[0][i - bbox.r1][j - bbox.c1]) {
					elem = elemsNoEmpty[0][i - bbox.r1][j - bbox.c1];
				}
				if (elem === null || elem.type === cElementType.empty) {
					elem = emptyReplaceOn;
				}

				arr.addElement(elem);
			}
		}

		return arr;
	};
	cArea3D.prototype.getMatrixNoEmpty = function () {
		var arr = [], r = this.getRanges(), res;

		var ws = r[0] ? r[0].worksheet : null;
		var oldExcludeHiddenRows = ws ? ws.bExcludeHiddenRows : null;

		for (var k = 0; k < r.length; k++) {
			arr[k] = [];
			r[k]._foreachNoEmpty(function (cell, i, j, r1, c1) {
				if (!arr[k][i - r1]) {
					arr[k][i - r1] = [];
				}
				res = checkTypeCell(cell);

				arr[k][i - r1][j - c1] = res;
			});
		}
		if (ws) {
			ws.bExcludeHiddenRows = oldExcludeHiddenRows;
		}

		return arr;
	};
	cArea3D.prototype.foreach2 = function (action) {
		var _wsA = this.wsRange();
		if (_wsA.length >= 1) {
			var _r = this.range(_wsA);
			for (var i = 0; i < _r.length; i++) {
				if (_r[i]) {
					_r[i]._foreach2(function (cell, row, col) {
						action(checkTypeCell(cell), cell, row, col);
					});
				}
			}
		}
	};
	cArea3D.prototype.isSingleSheet = function () {
		return this.wsFrom === this.wsTo;
	};
	cArea3D.prototype.isBetweenSheet = function (ws) {
		return ws && this.wsFrom.getIndex() <= ws.getIndex() && ws.getIndex() <= this.wsTo.getIndex();
	};
	cArea3D.prototype.getDimensions = function () {
		var res = null;
		if (this.bbox) {
			res = {col: this.bbox.c2 - this.bbox.c1 + 1, row: this.bbox.r2 - this.bbox.r1 + 1, bbox: this.bbox};
		}
		return res;
	};
	cArea3D.prototype.getFirstElement = function () {
		return this.getValueByRowCol(0, 0, true);
	};
	cArea3D.prototype._getCol = function (colIndex) {
		let dimensions = this.getDimensions();
		if (colIndex < 0 || colIndex > dimensions.col) {
			return null;
		}

		let col = [];
		for (let i = 0; i < dimensions.row; i++) {
			let elem = this.getValueByRowCol(i, colIndex);
			if (!elem) {
				elem = new cEmpty();
			}
			col[i] = [];
			col[i].push(elem);
		}
		return col;
	};
	cArea3D.prototype._getRow = function (rowIndex) {
		let dimensions = this.getDimensions();
		if (rowIndex < 0 || rowIndex > dimensions.row) {
			return null;
		}

		let row = [[]];
		for (let j = 0; j < this.getDimensions().col; j++) {
			let elem = this.getValueByRowCol(rowIndex, j);
			if (!elem) {
				elem = new cEmpty();
			}
			row[0].push(elem);
		}
		return row;
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cRef(val, ws) {/*Ref means A1 for example*/
		cBaseType.call(this, val);

		this.ws = ws;
		this.range = null;
		if (val) {
			AscCommonExcel.executeInR1C1Mode(false, function () {
				val = ws.getRange2(val.replace(AscCommon.rx_space_g, ""));
			});
			this.range = val;
		}
	}

	cRef.prototype = Object.create(cBaseType.prototype);
	cRef.prototype.constructor = cRef;
	cRef.prototype.type = cElementType.cell;
	cRef.prototype.clone = function (opt_ws) {
		var ws = opt_ws ? opt_ws : this.ws;
		var oRes = new cRef(null, ws);
		this.cloneTo(oRes);
		if (this.range) {
			oRes.range = this.range.clone(ws);
		}
		return oRes;
	};
	cRef.prototype.getWsId = function () {
		return this.ws.Id;
	};
	cRef.prototype.getValue = function () {
		if (!this.isValid()) {
			return new cError(cErrorType.bad_reference);
		}
		var res;
		this.range.getLeftTopCellNoEmpty(function (cell) {
			res = checkTypeCell(cell);
		});
		return res;
	};
	cRef.prototype.tocNumber = function () {
		return this.getValue().tocNumber();
	};
	cRef.prototype.tocString = function () {
		return this.getValue().tocString();
		/* new cString(""+this.range.getValueWithFormat()); */
	};
	cRef.prototype.tocBool = function () {
		return this.getValue().tocBool();
	};
	cRef.prototype.to3D = function (opt_ws) {
		var ws = opt_ws ? opt_ws : this.ws;
		var oRes = new cRef3D(null, null);
		this.cloneTo(oRes);
		oRes.ws = ws;
		if (this.range) {
			oRes.range = this.range.clone(ws);
		}
		return oRes;
	};
	cRef.prototype.toString = function () {
		if (AscCommonExcel.g_ProcessShared) {
			return this.range.getName();
		} else {
			return this.value;
		}
	};
	cRef.prototype.toLocaleString = function () {
		if (this.range) {
			return this.range.getName();
		} else {
			return this.value;
		}
	};
	cRef.prototype.getRange = function () {
		return this.range;
	};
	cRef.prototype.getWS = function () {
		return this.ws;
	};
	cRef.prototype.isValid = function () {
		return !!this.getRange();
	};
	cRef.prototype.getMatrix = function () {
		return [[this.getValue()]];
	};
	cRef.prototype.getBBox0 = function () {
		return this.getRange().getBBox0();
	};
	cRef.prototype.isHidden = function (excludeHiddenRows) {
		if (!excludeHiddenRows) {
			excludeHiddenRows = this.ws.isApplyFilterBySheet();
		}
		return excludeHiddenRows && this.isValid() && this.ws.getRowHidden(this.getRange().r1);
	};
	cRef.prototype.changeSheet = function (wsLast, wsNew) {
		if (this.ws === wsLast) {
			this.ws = wsNew;
			if (this.range) {
				this.range.worksheet = wsNew;
			}
		}
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cRef3D(val, ws, externalLink) {/*Ref means Sheat1!A1 for example*/
		cBaseType.call(this, val);

		this.ws = ws;
		this.range = null;
		if (val && this.ws) {
			AscCommonExcel.executeInR1C1Mode(false, function () {
				val = ws.getRange2(val);
			});
			this.range = val;
		}

		this.externalLink = externalLink;
	}

	cRef3D.prototype = Object.create(cBaseType.prototype);
	cRef3D.prototype.constructor = cRef3D;
	cRef3D.prototype.type = cElementType.cell3D;
	cRef3D.prototype.clone = function (opt_ws) {
		//TODO заливаю дополнительную проверку на вставку листа в другую книгу.
		//необходимо перепроверить и всегда, если приходит opt_ws, использовать только его.
		var isAddingSheet = Asc["editor"] && Asc["editor"].wb && Asc["editor"].wb.model && Asc["editor"].wb.model.addingWorksheet;
		var ws = opt_ws ? opt_ws : this.ws;
		var oRes = new cRef3D(null, null);
		this.cloneTo(oRes);
		if (opt_ws && (this.ws.getName() == opt_ws.getName() || isAddingSheet)) {
			oRes.ws = opt_ws;
		} else {
			oRes.ws = this.ws;
		}
		if (this.range) {
			oRes.range = this.range.clone(ws);
		}
		oRes.externalLink = this.externalLink;
		return oRes;
	};
	cRef3D.prototype.getWsId = function () {
		return this.ws && this.ws.Id;
	};
	cRef3D.prototype.getRange = function () {
		if (this.ws) {
			if (this.range) {
				return this.range;
			}
			return this.range = this.ws.getRange2 ? this.ws.getRange2(this.value) : null;
		} else {
			return this.range = null;
		}
	};
	cRef3D.prototype.isValid = function () {
		return !!this.getRange();
	};
	cRef3D.prototype.getValue = function () {
		const t = this;
		let _r = this.getRange();
		if (!_r) {
			return new cError(cErrorType.bad_reference);
		}
		var res;
		_r.getLeftTopCellNoEmpty(function (cell) {
			if (!cell && t.externalLink) {
				// if we refer to a non-existent cell in external data, return a #REF error
				res = new cError(cErrorType.bad_reference);
			} else {
				res = checkTypeCell(cell);
			}
		});
		return res;
	};
	cRef3D.prototype.tocBool = function () {
		return this.getValue().tocBool();
	};
	cRef3D.prototype.tocNumber = function () {
		return this.getValue().tocNumber();
	};
	cRef3D.prototype.tocString = function () {
		return this.getValue().tocString();
	};
	cRef3D.prototype.changeSheet = function (wsLast, wsNew) {
		//TODO обработать externalLink
		if (this.externalLink) {
			return;
		}
		if (this.ws === wsLast) {
			this.ws = wsNew;
			if (this.range) {
				this.range.worksheet = wsNew;
			}
		}
	};
	cRef3D.prototype.toString = function () {
		var exPath = this.getExternalLinkStr(this.externalLink);
		if (AscCommonExcel.g_ProcessShared) {
			return parserHelp.get3DRef(exPath + this.ws.getName(), this.range.getName());
		} else {
			return parserHelp.get3DRef(exPath + this.ws.getName(), this.value);
		}
	};
	cRef3D.prototype.toLocaleString = function () {
		var exPath = this.getExternalLinkStr(this.externalLink, true);
		return generate3DLink(exPath, this.ws.getName(), this.range.getName());
	};
	cRef3D.prototype.getWS = function () {
		return this.ws;
	};
	cRef3D.prototype.getMatrix = function () {
		return [[this.getValue()]];
	};
	cRef3D.prototype.getBBox0 = function () {
		var range = this.getRange();
		if (range) {
			return range.getBBox0();
		}
		return null;
	};
	cRef3D.prototype.isHidden = function (excludeHiddenRows) {
		if (!excludeHiddenRows) {
			excludeHiddenRows = this.ws.isApplyFilterBySheet();
		}
		var _r = this.getRange();
		return excludeHiddenRows && _r && this.ws.getRowHidden(_r.r1);
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cEmpty() {
		cBaseType.call(this, "");
	}

	cEmpty.prototype = Object.create(cBaseType.prototype);
	cEmpty.prototype.constructor = cEmpty;
	cEmpty.prototype.type = cElementType.empty;
	cEmpty.prototype.tocNumber = function () {
		return new cNumber(0);
	};
	cEmpty.prototype.tocBool = function () {
		return new cBool(false);
	};
	cEmpty.prototype.tocString = function () {
		return new cString("");
	};
	cEmpty.prototype.toString = function () {
		return "";
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cName(val, ws) {
		cBaseType.call(this, val);
		this.ws = ws;
	}

	cName.prototype = Object.create(cBaseType.prototype);
	cName.prototype.constructor = cName;
	cName.prototype.type = cElementType.name;
	cName.prototype.clone = function (opt_ws) {
		var ws = opt_ws ? opt_ws : this.ws;
		var oRes = new cName(this.value, ws);
		this.cloneTo(oRes);
		return oRes;
	};
	cName.prototype.toRef = function (opt_bbox, checkMultiSelect) {
		var defName = this.getDefName();
		if (!defName || !defName.ref) {
			return new cError(cErrorType.wrong_name);
		}
		return this.Calculate(undefined, opt_bbox, checkMultiSelect);
	};
	cName.prototype.toString = function () {
		var defName = this.getDefName();
		if (defName) {
			if (defName.isXLNM) {
				return new cString("_xlnm." + defName.name);
			}
			return defName.name;
		} else {
			return this.value;
		}
	};
	cName.prototype.toLocaleString = function () {
		var defName = this.getDefName();
		if (defName) {
			return defName.sheetId ? AscCommon.translateManager.getValue(defName.name) : defName.name;
		} else {
			//сделано для: создаем формулу со ссылкой на Область_печати, далее удаляем область печати с листа
			//поскольку в стеке лежит cName c именем "Print_Area", формула собиралась уже без учёта локали(мы попадали в текущую ветку и возвращали this.value)
			// - вместо области печати мы видим Print_Area
			//но с данной правкой есть проблема. если мы ссылаемся, допустим, в русской локали в формуле на именованный
			//диапазон Print_Area, то при сборке формулы он автоматически преобразуется в Область_Печати
			//аналогично тому, что если мы создаём в менеджере имен новое имя "Print_Area" - преоразуется с учетом локали
			return AscCommon.translateManager.getValue(this.value);
		}
	};
	cName.prototype.getValue = function () {
		return this.Calculate();
	};
	cName.prototype.getFormula = function () {
		var defName = this.getDefName();
		if (!defName || !defName.ref) {
			return new cError(cErrorType.wrong_name);
		}

		if (!defName.parsedRef) {
			return new cError(cErrorType.wrong_name);
		}
		return defName.parsedRef;
	};
	cName.prototype.Calculate = function () {
		var defName = this.getDefName();
		if (!defName || !defName.ref) {
			return new cError(cErrorType.wrong_name);
		}

		if (!defName.parsedRef) {
			return new cError(cErrorType.wrong_name);
		}

		//несмотря на то, что именованный диапазон ссылается на ошибку
		//при рассчётах с его участием необходимо возвращать пустую строку
		if (defName.type === Asc.c_oAscDefNameType.slicer) {
			return new cString("");
		}

		//defName not linked to cell, use inherit range
		var offset;
		var bbox = arguments[1];
		if (bbox) {
			//offset - to support relative references in def names
			offset = new AscCommon.CellBase(bbox.r1, bbox.c1);
		}
		return defName.parsedRef.calculate(this, bbox, offset, arguments[2]);
	};
	cName.prototype.getDefName = function () {
		return this.ws ? this.ws.workbook.getDefinesNames(this.value, this.ws.getId()) : null;
	};
	cName.prototype.changeDefName = function (from, to) {
		var sheetId = this.ws ? this.ws.getId() : null;
		if (AscCommonExcel.getDefNameIndex(this.value) == AscCommonExcel.getDefNameIndex(from.name)) {
			if (null == from.sheetId) {
				//in case of changes in workbook defname should not be sheet defname
				var defName = this.getDefName();
				if (!(defName && null != defName.sheetId)) {
					this.value = to.name;
				}
			} else if (sheetId == from.sheetId) {
				this.value = to.name;
			}
		}
	};
	cName.prototype.getWS = function () {
		return this.ws;
	};
	cName.prototype.changeSheet = function (wsLast, wsNew) {
		if (this.ws === wsLast) {
			this.ws = wsNew;
		}
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cStrucTable(val, wb, ws) {
		cBaseType.call(this, val);
		this.wb = wb;
		this.ws = ws;

		this.tableName = null;
		this.oneColumnIndex = null;
		this.colStartIndex = null;
		this.colEndIndex = null;
		this.reservedColumnIndex = null;
		this.hdtIndexes = null;
		this.hdtcstartIndex = null;
		this.hdtcendIndex = null;

		this.isDynamic = false;//#This row
		this.area = null;
	}

	cStrucTable.prototype = Object.create(cBaseType.prototype);
	cStrucTable.prototype.constructor = cStrucTable;
	cStrucTable.prototype.type = cElementType.table;
	cStrucTable.prototype.createFromVal = function (val, wb, ws, tablesMap) {
		var res = new cStrucTable(val[0], wb, ws);
		if (tablesMap && tablesMap[val["tableName"]]) {
			val["tableName"] = tablesMap[val["tableName"]];
		}
		if (res._parseVal(val)) {
			res._updateArea(null, false);
		}
		return (res.area && res.area.type != cElementType.error) ? res : new cError(cErrorType.bad_reference);
	};
	cStrucTable.prototype.clone = function (opt_ws) {
		var ws = opt_ws ? opt_ws : this.ws;
		var wb = ws.workbook;
		var oRes = new cStrucTable(this.value, wb, ws);
		oRes.tableName = this.tableName;
		oRes.oneColumnIndex = this._cloneIndex(this.oneColumnIndex);
		oRes.colStartIndex = this._cloneIndex(this.colStartIndex);
		oRes.colEndIndex = this._cloneIndex(this.colEndIndex);
		oRes.reservedColumnIndex = this.reservedColumnIndex;
		if (this.hdtIndexes) {
			oRes.hdtIndexes = this.hdtIndexes.slice(0);
		}
		oRes.hdtcstartIndex = this._cloneIndex(this.hdtcstartIndex);
		oRes.hdtcendIndex = this._cloneIndex(this.hdtcendIndex);

		oRes.isDynamic = this.isDynamic;
		if (this.area) {
			if (this.area.clone) {
				oRes.area = this.area.clone(opt_ws);
			} else {
				oRes.area = this.area;
			}
		}
		this.cloneTo(oRes);
		return oRes;
	};
	cStrucTable.prototype._cloneIndex = function (val) {
		if (val) {
			return {wsID: val.wsID, index: val.index, name: val.name};
		} else {
			return val;
		}
	};
	cStrucTable.prototype.toRef = function (opt_bbox, opt_bConvertTableFormulaToRef) {
		//opt_bbox usefull only for #This row
		//case null == opt_bbox works like FormulaTablePartInfo.data
		var table = this.wb.getDefinesNames(this.tableName, this.ws ? this.ws.getId() : null);
		if (!table || !table.ref) {
			return new cError(cErrorType.wrong_name);
		}
		if (!this.area || this.isDynamic) {
			this._updateArea(opt_bbox, true, opt_bConvertTableFormulaToRef);
		}
		return this.area;
	};
	cStrucTable.prototype.toString = function () {
		return this._toString(false);
	};
	cStrucTable.prototype.toLocaleString = function () {
		return this._toString(true);
	};
	cStrucTable.prototype._toString = function (isLocal) {
		// file works with "#This Row" - user with "@"
		// isLocal - change "#This Row", to "@"
		const table = this.wb.getDefinesNames(this.tableName, null);
		let tblStr, columns_1, columns_2;
		if (!table) {
			tblStr = this.tableName;
		} else {
			tblStr = table.name;
		}

		/* escapeTableCharacters - add special character escaping for string inside the table (escaping with single quote) */
		if (this.oneColumnIndex) {
			// TODO add this.isCrossSign to use?
			columns_1 = parserHelp.escapeTableCharacters(this.oneColumnIndex.name, true/*doEscape*/);

			if (this.isDynamic && isLocal) {
				columns_1 = "@" + columns_1;
			} else if (this.isDynamic) {
				columns_1 = "[" + this._buildLocalTableString(AscCommon.FormulaTablePartInfo.thisRow, isLocal) + "]" + 
					FormulaSeparators.functionArgumentSeparatorDef + "[" + columns_1 + "]";
			}

			tblStr += "[" + columns_1 + "]";
		} else if (this.colStartIndex && this.colEndIndex) {
			columns_1 = parserHelp.escapeTableCharacters(this.colStartIndex.name, true/*doEscape*/);
			columns_2 = parserHelp.escapeTableCharacters(this.colEndIndex.name, true/*doEscape*/);

			tblStr += "[[" + columns_1 + "]:[" + columns_2 + "]]";
		} else if (null != this.reservedColumnIndex) {
			if (this.isDynamic && isLocal && this.reservedColumnIndex === AscCommon.FormulaTablePartInfo.thisRow) {
				tblStr += "[" + "@" + "]";
			} else /*if (this.isDynamic)*/ {
				tblStr += "[" + this._buildLocalTableString(this.reservedColumnIndex, isLocal) + "]";
			}
		} else if (this.hdtIndexes || this.hdtcstartIndex || this.hdtcendIndex) {
			tblStr += '[';
			let i;

			if (this.hdtIndexes.length > 0 && this.isDynamic && isLocal && this.hdtIndexes[0] === AscCommon.FormulaTablePartInfo.thisRow) {
				let hdtcstart = this.hdtcstartIndex ? parserHelp.escapeTableCharacters(this.hdtcstartIndex.name, true) : null;
				let hdtcend = this.hdtcendIndex ? parserHelp.escapeTableCharacters(this.hdtcendIndex.name, true) : null;
				
				tblStr += "@";
				if (hdtcstart && !hdtcend) {
					// if one column is selected
					tblStr += hdtcstart;
				} else if (hdtcstart && hdtcend) {
					// if multiple columns are selected
					tblStr += '[' + hdtcstart + ']';
					tblStr += ':[' + hdtcend + ']';
				}

			} else {
				for (i = 0; i < this.hdtIndexes.length; ++i) {
					if (0 != i) {
						if (isLocal) {
							tblStr += FormulaSeparators.functionArgumentSeparator;
						} else {
							tblStr += FormulaSeparators.functionArgumentSeparatorDef;
						}
					}

					if (this.hdtcstartIndex === null && this.hdtIndexes.length === 1) {
						// If the formula contains a single hdt index, remove the inner brackets =Table[[#Headers|#All|#Data|#Totals]]
						tblStr += this._buildLocalTableString(this.hdtIndexes[i], isLocal);
					} else {
						tblStr += "[" + this._buildLocalTableString(this.hdtIndexes[i], isLocal) + "]";
					}
				}

				if (this.hdtcstartIndex) {
					if (this.hdtIndexes.length > 0) {
						if (isLocal) {
							tblStr += FormulaSeparators.functionArgumentSeparator;
						} else {
							tblStr += FormulaSeparators.functionArgumentSeparatorDef;
						}
					}
					let hdtcstart = parserHelp.escapeTableCharacters(this.hdtcstartIndex.name, true);

					tblStr += "[" + hdtcstart + "]";
					if (this.hdtcendIndex) {
						let hdtcend = parserHelp.escapeTableCharacters(this.hdtcendIndex.name, true);

						tblStr += ":[" + hdtcend + "]";
					}
				}
			}

			tblStr += ']';
		} else if (!isLocal) {
			tblStr += '[]';
		}
		return tblStr;
	};
	cStrucTable.prototype._parseVal = function (val) {
		let bRes = true, startCol, endCol;
		this.tableName = val['tableName'];

		// inside .getTableIndexColumnByName() we perform .replace for the column name we are looking for
		if (val['oneColumn']) {
			startCol = val['oneColumn']
			if (startCol[0] === "@") {
				this.isDynamic = true;
			}

			let openBracketIndex = startCol.indexOf("[");
			if (openBracketIndex !== -1) {
				let closeBracketIndex = startCol.lastIndexOf("]");
				if (closeBracketIndex !== -1) {
					startCol = startCol.slice(openBracketIndex + 1, closeBracketIndex);
				} 
			}

			this.oneColumnIndex = this.wb.getTableIndexColumnByName(this.tableName, this.isDynamic ? startCol.slice(1) : startCol);
			bRes = !!this.oneColumnIndex;
		} else if (val['columnRange']) {
			startCol = val['colStart'];
			endCol = val['colEnd'];

			if (!endCol) {
				endCol = startCol;
			}
			this.colStartIndex = this.wb.getTableIndexColumnByName(this.tableName, startCol);
			this.colEndIndex = this.wb.getTableIndexColumnByName(this.tableName, endCol);
			bRes = !!this.colStartIndex && !!this.colEndIndex;
		} else if (val['reservedColumn']) {
			this.reservedColumnIndex = parserHelp.getColumnTypeByName(val['reservedColumn']);
			if (AscCommon.FormulaTablePartInfo.thisRow == this.reservedColumnIndex ||
				AscCommon.FormulaTablePartInfo.headers == this.reservedColumnIndex ||
				AscCommon.FormulaTablePartInfo.totals == this.reservedColumnIndex) {
				this.isDynamic = true;
			}
		} else if (val['hdtcc']) {
			this.hdtIndexes = [];
			let hdtcstart = val['hdtcstart'];
			let hdtcend = val['hdtcend'];
			let re = /\[(.*?)\]|\@/ig, m;

			let isCross;
			if (val['hdt'] === "@") {
				isCross = true;
			}

			while (null !== (m = re.exec(val['hdt']))) {
				let param = parserHelp.getColumnTypeByName(isCross ? m[0] : m[1]);
				if (AscCommon.FormulaTablePartInfo.thisRow == param ||
					AscCommon.FormulaTablePartInfo.headers == param || AscCommon.FormulaTablePartInfo.totals == param) {
					this.isDynamic = true;
				}
				this.hdtIndexes.push(param);
			}

			if (hdtcstart) {
				startCol = hdtcstart;
				this.hdtcstartIndex = this.wb.getTableIndexColumnByName(this.tableName, startCol);
				bRes = !!this.hdtcstartIndex;
				if (bRes && hdtcend) {
					endCol = hdtcend;
					this.hdtcendIndex = this.wb.getTableIndexColumnByName(this.tableName, endCol);
					bRes = !!this.hdtcendIndex;
				}
			}
		}
		return bRes;
	};
	cStrucTable.prototype._updateArea = function (bbox, toRef, bConvertTableFormulaToRef) {
		var paramObj = {param: null, startCol: null, endCol: null, cell: bbox, toRef: toRef, bConvertTableFormulaToRef: bConvertTableFormulaToRef};
		var isThisRow = false;
		var tableData, refName;
		if (this.oneColumnIndex) {
			if (this.isDynamic) {
				/* this row */
				isThisRow = true;
				paramObj.param = AscCommon.FormulaTablePartInfo.thisRow;
				let thisRow = this.wb.getTableRangeForFormula(this.tableName, paramObj);
				
				let thisCol;
				if (thisRow) {
					paramObj.param = AscCommon.FormulaTablePartInfo.columns;
					paramObj.startCol = this.oneColumnIndex.name;
					paramObj.endCol = null;
					thisCol = this.wb.getTableRangeForFormula(this.tableName, paramObj);
				}

				if (!thisRow || !thisCol) {
					return this._createAreaError(isThisRow);
				}

				range = new Asc.Range(thisCol.range.c1, thisRow.range.r1, thisCol.range.c2, thisRow.range.r2);

				tableData = thisCol;
				tableData.range = range;
			} else {
				paramObj.param = AscCommon.FormulaTablePartInfo.columns;
				paramObj.startCol = this.oneColumnIndex.name;
			}
		} else if (this.colStartIndex && this.colEndIndex) {
			paramObj.param = AscCommon.FormulaTablePartInfo.columns;
			paramObj.startCol = this.colStartIndex.name;
			paramObj.endCol = this.colEndIndex.name;
		} else if (null != this.reservedColumnIndex) {
			paramObj.param = this.reservedColumnIndex;
			isThisRow = AscCommon.FormulaTablePartInfo.thisRow == paramObj.param;
		} else if (this.hdtIndexes || this.hdtcstartIndex) {
			var data, range;
			if (this.hdtIndexes) {
				for (var i = 0; i < this.hdtIndexes.length; ++i) {
					paramObj.param = this.hdtIndexes[i];
					isThisRow = AscCommon.FormulaTablePartInfo.thisRow == paramObj.param;
					data = this.wb.getTableRangeForFormula(this.tableName, paramObj);
					if (!data) {
						return this._createAreaError(isThisRow);
					}

					if (range) {
						range.union2(data.range);
					} else {
						range = data.range;
					}
				}
			}

			if (this.hdtcstartIndex) {
				paramObj.param = AscCommon.FormulaTablePartInfo.columns;
				paramObj.startCol = this.hdtcstartIndex.name;
				paramObj.endCol = null;

				if (this.hdtcendIndex) {
					paramObj.endCol = this.hdtcendIndex.name;
				}
				data = this.wb.getTableRangeForFormula(this.tableName, paramObj);
				if (!data) {
					return this._createAreaError(isThisRow);
				}
				if (range) {
					var r1Abs = range.isAbsR1();
					var c1Abs = data.range.isAbsC1();
					var r2Abs = range.isAbsR2();
					var c2Abs = data.range.isAbsC2();
					range = new Asc.Range(data.range.c1, range.r1, data.range.c2, range.r2);
					range.setAbs(r1Abs, c1Abs, r2Abs, c2Abs);
				} else {
					range = data.range;
				}
			}

			tableData = data;
			tableData.range = range;
		} else {
			paramObj.param = AscCommon.FormulaTablePartInfo.data;
		}
		if (!tableData) {
			tableData = this.wb.getTableRangeForFormula(this.tableName, paramObj);
			if (!tableData) {
				return this._createAreaError(isThisRow);
			}
		}
		if (tableData.range) {
			//всегда получаем диапазон в виде A1B1
			AscCommonExcel.executeInR1C1Mode(false, function () {
				refName = tableData.range.getName();
			});

			var wsFrom = this.wb.getWorksheetById(tableData.wsID);
			if (tableData.range.isOneCell()) {
				this.area = new cRef3D(refName, wsFrom);
			} else {
				this.area = new cArea3D(refName, wsFrom, wsFrom);
			}
		} else {
			this.area = new cError(cErrorType.bad_reference);
		}
		return this.area;
	};
	cStrucTable.prototype._createAreaError = function (isThisRow) {
		if (isThisRow) {
			return this.area = new cError(cErrorType.wrong_value_type);
		} else {
			return this.area = new cError(cErrorType.bad_reference);
		}
	};
	cStrucTable.prototype._buildLocalTableString = function (reservedColumn, local) {
		return parserHelp.getColumnNameByType(reservedColumn, local);
	};
	cStrucTable.prototype.changeDefName = function (from, to) {
		if (this.tableName == from.name) {
			this.tableName = to.name;
		}
	};
	cStrucTable.prototype.removeTableColumn = function (deleted) {
		if (this.oneColumnIndex) {
			if (deleted[this.oneColumnIndex.name]) {
				return true;
			} else {
				this.oneColumnIndex = this.wb.getTableIndexColumnByName(this.tableName, this.oneColumnIndex.name);
				if (!this.oneColumnIndex) {
					return true;
				}
			}
		}
		if (this.colStartIndex && this.colEndIndex) {
			if (deleted[this.colStartIndex.name]) {
				return true;
			} else {
				this.colStartIndex = this.wb.getTableIndexColumnByName(this.tableName, this.colStartIndex.name);
				if (!this.colStartIndex) {
					return true;
				}
			}
			if (deleted[this.colEndIndex.name]) {
				return true;
			} else {
				this.colEndIndex = this.wb.getTableIndexColumnByName(this.tableName, this.colEndIndex.name);
				if (!this.colEndIndex) {
					return true;
				}
			}
		}
		if (this.hdtcstartIndex) {
			if (deleted[this.hdtcstartIndex.name]) {
				return true;
			} else {
				this.hdtcstartIndex = this.wb.getTableIndexColumnByName(this.tableName, this.hdtcstartIndex.name);
				if (!this.hdtcstartIndex) {
					return true;
				}
			}
		}
		if (this.hdtcendIndex) {
			if (deleted[this.hdtcendIndex.name]) {
				return true;
			} else {
				this.hdtcendIndex = this.wb.getTableIndexColumnByName(this.tableName, this.hdtcendIndex.name);
				if (!this.hdtcendIndex) {
					return true;
				}
			}
		}
		return false;
	};
	cStrucTable.prototype.changeTableRef = function () {
		if (!this.isDynamic) {
			this._updateArea(null, false);
		}
	};
	cStrucTable.prototype.renameTableColumn = function () {
		var bRes = true;
		var columns1, columns2;
		if (this.oneColumnIndex) {
			columns1 = this.wb.getTableNameColumnByIndex(this.tableName, this.oneColumnIndex.index);
			if (columns1) {
				this.oneColumnIndex.name = columns1.columnName;
			} else {
				bRes = false;
			}
		} else if (this.colStartIndex && this.colEndIndex) {
			columns1 = this.wb.getTableNameColumnByIndex(this.tableName, this.colStartIndex.index);
			columns2 = this.wb.getTableNameColumnByIndex(this.tableName, this.colEndIndex.index);
			if (columns1 && columns2) {
				this.colStartIndex.name = columns1.columnName;
				this.colEndIndex.name = columns2.columnName;
			} else {
				bRes = false;
			}
		}
		if (this.hdtcstartIndex) {
			columns1 = this.wb.getTableNameColumnByIndex(this.tableName, this.hdtcstartIndex.index);
			if (columns1) {
				this.hdtcstartIndex.name = columns1.columnName;
			} else {
				bRes = false;
			}
		}
		if (this.hdtcendIndex) {
			columns1 = this.wb.getTableNameColumnByIndex(this.tableName, this.hdtcendIndex.index);
			if (columns1) {
				this.hdtcendIndex.name = columns1.columnName;
			} else {
				bRes = false;
			}
		}
		return bRes;
	};
	cStrucTable.prototype.geColumnHeadings = function() {
		var res = [];
		var table = this.wb.getTableByName(this.tableName);
		if (!table) {
			return res;
		}
		var from = 0;
		var to = table.TableColumns.length - 1;
		if (this.oneColumnIndex) {
			from = to = this.oneColumnIndex.index;
		} else if (this.colStartIndex && this.colEndIndex) {
			from = this.colStartIndex.index;
			to = this.colEndIndex.index;
		}
		if (this.hdtcstartIndex && this.hdtcendIndex) {
			from = this.hdtcstartIndex.index;
			to = this.hdtcendIndex.index;
		}
		for (var i = from; i <= to; ++i) {
			res.push(table.TableColumns[i].getTableColumnName());
		}
		return res;
	};
	cStrucTable.prototype.getTable = function() {
		return this.wb.getTableByName(this.tableName);
	};
	cStrucTable.prototype.getWS = function () {
		return this.ws;
	};
	cStrucTable.prototype.changeSheet = function(wsLast, wsNew) {
		if (this.ws === wsLast) {
			this.ws = wsNew;
			if (this.area && this.area.changeSheet) {
				this.area.changeSheet(wsLast, wsNew);
			}
		}
	};
	cStrucTable.prototype.setOffset = function(offset) {
		var t = this;

		var tryDiffHdtcIndex = function(oIndex) {
			var table = t.wb.getTableByNameAndSheet(t.tableName, oIndex.wsID);
			if(table) {
				var tableColumnsCount = table.TableColumns.length;
				var index = oIndex.index + offset.col;
				index = index - Math.floor(index / tableColumnsCount) * tableColumnsCount;
				var columnName = t.wb.getTableNameColumnByIndex(t.tableName, index);
				if(columnName) {
					oIndex.index = index;
					oIndex.name = columnName.columnName;
				}
			}
		};

		//TODO
		if(this.oneColumnIndex) {
			if(offset && offset.col) {
				tryDiffHdtcIndex(this.oneColumnIndex);
			}
		} else if(this.colStartIndex && this.colEndIndex) {

		} else if(this.hdtIndexes || this.hdtcstartIndex || this.hdtcendIndex) {
			if(offset && offset.col) {
				if(this.hdtcstartIndex) {
					tryDiffHdtcIndex(this.hdtcstartIndex);
				}
				if(this.hdtcendIndex) {
					tryDiffHdtcIndex(this.hdtcendIndex);
				}
			}
		}
	};
	cStrucTable.prototype.getRange = function () {
		return this.area && this.area.getRange && this.area.getRange();
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cStrucPivotTable(val) {
		cBaseType.call(this, val);
		if (val) {
			this.isIndex = false;
			this.fieldString = val[0];
			this.itemString = val[1];
			if (!isNaN(val[1])) {
				this.isIndex = true;
			}
		}
	}

	cStrucPivotTable.prototype = Object.create(cBaseType.prototype);
	cStrucPivotTable.prototype.constructor = cStrucPivotTable;

	cStrucPivotTable.prototype.type = cElementType.pivotTable;
	cStrucPivotTable.prototype.createFromVal = function (val) {
		//TODO check on error
		let res = new cStrucPivotTable(val);
		return res;
	};
	cStrucPivotTable.prototype.clone = function () {

	};
	cStrucPivotTable.prototype.Calculate = function (callback) {
		return callback(this.fieldString, this.itemString, this.isIndex);
	};
	cStrucPivotTable.prototype.toString = function () {
		return this._toString(false);
	};
	cStrucPivotTable.prototype.toLocaleString = function () {
		return this._toString(true);
	};
	cStrucPivotTable.prototype._toString = function (isLocal) {
		if (this.fieldString) {
			return this.fieldString + '[' + this.itemString + ']';
		}
		return this.itemString;
	};

	/**
	 * @constructor
	 * @extends {cName}
	 */
	function cName3D(val, ws, externalLink, shortLink) {
		cName.call(this, val, ws);
		this.externalLink = externalLink;
		this.shortLink = shortLink;
	}

	cName3D.prototype = Object.create(cName.prototype);
	cName3D.prototype.constructor = cName3D;
	cName3D.prototype.type = cElementType.name3D;
	cName3D.prototype.clone = function (opt_ws) {
		var ws;
		if (opt_ws && opt_ws.getName() === this.ws.getName()) {
			ws = opt_ws;
		} else {
			ws = this.ws;
		}
		var oRes = new cName3D(this.value, ws, this.externalLink);
		this.cloneTo(oRes);
		return oRes;
	};

	cName3D.prototype.toString = function () {
		let exPath = this.getExternalLinkStr(this.externalLink);
		let wsName = this.ws && this.ws.getName();
		/* short links returns without wsName */
		return parserHelp.getEscapeSheetName(this.shortLink ? exPath : (exPath +  (wsName ? wsName : "")), this.shortLink) + "!" + cName.prototype.toString.call(this);
	};
	cName3D.prototype.toLocaleString = function () {
		let exPath = this.getExternalLinkStr(this.externalLink, true, this.shortLink);
		let wsName = this.ws && this.ws.getName();
		/* short links returns without wsName */
		return parserHelp.getEscapeSheetName(this.shortLink ? exPath : (exPath +  (wsName ? wsName : ""))) + "!" + cName.prototype.toLocaleString.call(this);
	};
	cName3D.prototype.getWsId = function () {
		return this.ws && this.ws.Id;
	};

	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cArray() {
		cBaseType.call(this, undefined);
		this.array = [];
		this.rowCount = 0;
		this.countElementInRow = [];
		this.countElement = 0;

		this.realSize = null;
		this.missedValue = null;
	}

	cArray.prototype = Object.create(cBaseType.prototype);
	cArray.prototype.constructor = cArray;
	cArray.prototype.type = cElementType.array;
	cArray.prototype.addRow = function () {
		this.array[this.array.length] = [];
		this.countElementInRow[this.rowCount++] = 0;
	};
	cArray.prototype.addElement = function (element) {
		if (this.array.length === 0) {
			this.addRow();
		}
		var arr = this.array, subArr = arr[this.rowCount - 1];
		subArr[subArr.length] = element;
		this.countElementInRow[this.rowCount - 1]++;
		this.countElement++;
	};
	cArray.prototype.addElementInRow = function (element, rowIndex) {
		if (typeof rowIndex !== "number" || rowIndex < 0 || rowIndex > this.rowCount) {
			return null;
		}

		if (this.array.length === 0) {
			this.addRow();
		}
		let arr = this.array, subArr = arr[rowIndex]; 
		subArr[subArr.length] = element;
		this.countElementInRow[rowIndex]++;
		this.countElement++;
	};
	cArray.prototype.getRow = function (rowIndex) {
		if (rowIndex < 0 || rowIndex > this.array.length - 1) {
			return null;
		}
		return this.array[rowIndex];
	};
	cArray.prototype._getRow = function (rowIndex) {
		if (rowIndex < 0 || rowIndex > this.array.length - 1) {
			return null;
		}
		return [this.array[rowIndex]];
	};
	cArray.prototype.getCol = function (colIndex) {
		var col = [];
		for (var i = 0; i < this.rowCount; i++) {
			col.push(this.array[i][colIndex]);
		}
		return col;
	};
	cArray.prototype._getCol = function (colIndex) {
		let col = [];
		for (let i = 0; i < this.rowCount; i++) {
			col[i] = [];
			col[i].push(this.array[i][colIndex]);
		}
		return col;
	};
	cArray.prototype.getElementRowCol = function (row, col, checkRealSize) {
		if (row > this.rowCount || col > this.getCountElementInRow()) {
			if (checkRealSize && this.realSize && row <= this.realSize.row && col <= this.realSize.col) {
				if (this.missedValue) {
					return this.missedValue
				}
				return new cEmpty();
			}
			return new cError(cErrorType.not_available);
		}
		return this.array[row] && this.array[row][col] ? this.array[row][col] : new cEmpty();
	};
	cArray.prototype.getElement = function (index) {
		for (var i = 0; i < this.rowCount; i++) {
			//TODO length
			if (index > this.countElementInRow[i].length) {
				index -= this.countElementInRow[i].length;
			} else {
				return this.array[i][index];
			}
		}
		return null;
	};
	cArray.prototype.foreach = function (action) {
		if (typeof (action) !== 'function') {
			return true;
		}
		for (var ir = 0; ir < this.rowCount; ir++) {
			for (var ic = 0; ic < this.countElementInRow[ir]; ic++) {
				if (action.call(this, this.array[ir][ic], ir, ic)) {
					return true;
				}
			}
		}
		return undefined;
	};
	cArray.prototype.foreach2 = function (action, byCol) {
		if (typeof (action) !== 'function') {
			return true;
		}

		let ir, ic;
		if (byCol) {
			for (ic = 0; ic < this.geMaxElementInRow(); ic++) {
				for (ir = 0; ir < this.rowCount; ir++) {
					action.call(this, this.array[ir][ic], ir, ic)
				}
			}
		} else {
			for (ir = 0; ir < this.rowCount; ir++) {
				for (ic = 0; ic < this.countElementInRow[ir]; ic++) {
					action.call(this, this.array[ir][ic], ir, ic)
				}
			}
		}
	};
	cArray.prototype.getCountElement = function () {
		return this.countElement;
	};
	cArray.prototype.getCountElementInRow = function (getRealSize) {
		return getRealSize && this.realSize ? this.realSize.col : this.countElementInRow[0];
	};
	cArray.prototype.getRowCount = function (getRealSize) {
		return getRealSize && this.realSize ? this.realSize.row : this.rowCount;
	};
	cArray.prototype.geMaxElementInRow = function () {
		return Math.max.apply(null, this.countElementInRow);
	};
	cArray.prototype.getRealArraySize = function () {
		if (!this.realSize) {
			return;
		}

		return this.realSize;
	};
	cArray.prototype.getMissedValue = function () {
		if (!this.missedValue) {
			return;
		}

		return this.missedValue;
	};
	cArray.prototype.setRealArraySize = function (row, col) {
		if (row > 0 && col > 0) {
			this.realSize = {row: row, col: col}
		}
	};
	cArray.prototype.tocNumber = function () {
		let retArr = new cArray();
		retArr.realSize = this.getRealArraySize();
		retArr.missedValue = this.getMissedValue();
		for (let ir = 0; ir < this.rowCount; ir++, retArr.addRow()) {
			for (let ic = 0; ic < this.countElementInRow[ir]; ic++) {
				retArr.addElement(this.array[ir][ic].tocNumber());
			}
			if (ir === this.rowCount - 1) {
				break;
			}
		}
		return retArr;
	};
	cArray.prototype.tocString = function () {
		var retArr = new cArray();
		for (var ir = 0; ir < this.rowCount; ir++, retArr.addRow()) {
			for (var ic = 0; ic < this.countElementInRow[ir]; ic++) {
				retArr.addElement(this.array[ir][ic].tocString());
			}
			if (ir === this.rowCount - 1) {
				break;
			}
		}
		return retArr;
	};
	cArray.prototype.tocBool = function () {
		var retArr = new cArray();
		for (var ir = 0; ir < this.rowCount; ir++, retArr.addRow()) {
			for (var ic = 0; ic < this.countElementInRow[ir]; ic++) {
				retArr.addElement(this.array[ir][ic].tocBool());
			}
			if (ir === this.rowCount - 1) {
				break;
			}
		}
		return retArr;
	};
	cArray.prototype.toString = function () {
		var ret = "";
		for (var ir = 0; ir < this.rowCount; ir++, ret += FormulaSeparators.arrayRowSeparatorDef) {
			for (var ic = 0; ic < this.countElementInRow[ir]; ic++, ret += FormulaSeparators.arrayColSeparatorDef) {
				if (this.array[ir][ic] instanceof cString) {
					ret += '"' + this.array[ir][ic].toString() + '"';
				} else {
					ret += this.array[ir][ic].toString() + "";
				}
			}
			if (ret[ret.length - 1] === FormulaSeparators.arrayColSeparatorDef) {
				ret = ret.substring(0, ret.length - 1);
			}
		}
		if (ret[ret.length - 1] === FormulaSeparators.arrayRowSeparatorDef) {
			ret = ret.substring(0, ret.length - 1);
		}
		return "{" + ret + "}";
	};
	cArray.prototype.toLocaleString = function (digitDelim) {
		var ret = "";
		for (var ir = 0; ir < this.rowCount;
			 ir++, ret += digitDelim ? FormulaSeparators.arrayRowSeparator : FormulaSeparators.arrayRowSeparatorDef) {
			for (var ic = 0; ic < this.countElementInRow[ir]; ic++, ret +=
				digitDelim ? FormulaSeparators.arrayColSeparator : FormulaSeparators.arrayColSeparatorDef) {
			if (this.array[ir] && this.array[ir][ic]) {
					if (this.array[ir][ic] instanceof cString) {
						ret += '"' + this.array[ir][ic].toLocaleString(digitDelim) + '"';
					} else {
						ret += this.array[ir][ic].toLocaleString(digitDelim) + "";
					}
				}
			}
			if (ret[ret.length - 1] === digitDelim ? FormulaSeparators.arrayColSeparator :
					FormulaSeparators.arrayColSeparatorDef) {
				ret = ret.substring(0, ret.length - 1);
			}
		}
		if (ret[ret.length - 1] === digitDelim ? FormulaSeparators.arrayRowSeparator :
				FormulaSeparators.arrayRowSeparatorDef) {
			ret = ret.substring(0, ret.length - 1);
		}
		return "{" + ret + "}";
	};
	cArray.prototype.isValidArray = function () {
		if (this.countElement < 1) {
			return false;
		}
		for (var i = 0; i < this.rowCount - 1; i++) {
			if (this.countElementInRow[i] - this.countElementInRow[i + 1] !== 0) {
				return false;
			}
		}
		return true;
	};
	cArray.prototype.getValue2 = function (i, j) {
		var result = this.array[i];
		return result ? result[j] : result;
	};
	cArray.prototype.getMatrix = function () {

		//excludeErrorsVal - arguments[1]
		if(arguments[1]) {
			var retArr = new cArray();
			for (var ir = 0; ir < this.rowCount; ir++, retArr.addRow()) {
				for (var ic = 0; ic < this.countElementInRow[ir]; ic++) {
					var elem = this.array[ir][ic];
					if(AscCommonExcel.cElementType.error === elem.type) {
						elem = new cEmpty();
					}
					retArr.addElement(elem);
				}
				if (ir === this.rowCount - 1) {
					break;
				}
			}
			return retArr.array;
		}

		return this.array;
	};
	cArray.prototype.getMatrixCopy = function () {
		let retArrCopy = [];
		for (let ir = 0; ir < this.array.length; ir++) {
			retArrCopy[ir] = [];
			for (let ic = 0; ic < this.array[ir].length; ic++) {
				// let elem = this.array[ir][ic];
				retArrCopy[ir].push(this.array[ir][ic]);
			}
			if (ir === this.rowCount - 1) {
				break;
			}
		}

		return retArrCopy;

	};
	cArray.prototype.fillFromArray = function (arr, fChangeElems) {
		if (arr && arr.length !== undefined) {
			this.array = arr;
			this.rowCount = arr.length;
			for (var i = 0; i < arr.length; i++) {
				this.countElementInRow[i] = arr[i].length;
				this.countElement += arr[i].length;
				if (fChangeElems){
					for (let j = 0; j < arr[i].length; j++) {
						let changeRes = fChangeElems(arr[i][j]);
						if (changeRes !== null) {
							arr[i][j] = changeRes;
						} else {
							return null;
						}
					}
				}
			}
		}
		return true;
	};
	cArray.prototype.fillEmptyFromRange = function (range) {
		if(!range) {
			return;
		}

		for(var i = range.r1; i <= range.r2; i++) {
			this.addRow();
			for(var j = range.c1; j <= range.c2; j++) {
				this.addElement(null);
			}
		}
	};
	cArray.prototype.getDimensions = function (getRealSize) {
		let realSize = getRealSize ? this.getRealArraySize() : false;
		let col, row;
		if (!realSize) {
			col = this.getCountElementInRow();
			if (!col) {
				col = 1;
			}

			row = this.getRowCount();
			if (!row) {
				row = 1;
			}
		}

		return {col: realSize ? realSize.col : col, row: realSize ? realSize.row : row};
	};
	cArray.prototype.fillMatrix = function (replace_empty) {
		let maxColCount = Math.max.apply(null, this.countElementInRow);
		this.countElementInRow = [];
		this.countElement = 0;
		for (let i = 0; i < this.rowCount; i++) {
			let currentCount = this.array[i].length;
			if (currentCount < maxColCount) {
				for (let j = 0; j < maxColCount - currentCount; j++) {
					this.array[i].push(replace_empty);
				}
			}
			this.countElementInRow[i] = this.array[i].length;
			this.countElement += this.array[i].length;
		}
	};
	cArray.prototype.recalculateOld = function () {
		this.rowCount = this.array.length;
		this.countElementInRow = [];
		this.countElement = 0;
		for (var i = 0; i < this.array.length; i++) {
			this.countElementInRow[i] = this.array[i].length;
			this.countElement += this.array[i].length;
		}
	};
	cArray.prototype.recalculate = function (row) {
		this.rowCount = this.array.length;
		if (row === undefined) {
			// full recalculation of the number of elements in the entire array, long execution
			this.countElementInRow = [];
			this.countElement = 0;
			for (let i = 0; i < this.array.length; i++) {
				this.countElementInRow[i] = this.array[i].length;
				this.countElement += this.array[i].length;
			}
		} else {
			// changing only the affected values ​​(by row)
			let lookingRow = this.array[row];
			this.countElementInRow[row] = lookingRow.length;
			this.countElement += lookingRow.length;
		}
	};
	cArray.prototype.pushCol = function (matrix, colNum) {
		for (let i = 0; i < matrix.length; i++) {
			if (matrix[i] && matrix[i][colNum]) {
				if (!this.array[i]) {
					this.array[i] = [];
				}
				this.array[i].push(matrix[i][colNum]);
			}
		}
		this.recalculate();
	};
	cArray.prototype.pushRow = function (matrix, rowNum) {
		if (matrix && matrix[rowNum]) {
			this.array.push(matrix[rowNum]);
			this.recalculate(this.array.length - 1);
		}
	};
	cArray.prototype.crop = function (row, col) {
		let newArray = this.array;
		let dimensions = this.getDimensions();
		if (row && Math.abs(row) < dimensions.row) {
			if (row < 0) {
				newArray = newArray.splice(this.array.length - Math.abs(row));
			} else {
				newArray = newArray.splice(0, row);
			}
		}
		if (col && Math.abs(col) < dimensions.col) {
			for (let i = 0; i < newArray.length; i++) {
				if (col < 0) {
					newArray[i] = newArray[i].splice(newArray[i].length - Math.abs(col));
				} else {
					newArray[i] = newArray[i].splice(0, col);
				}
			}
		}
		let res = new cArray();
		res.fillFromArray(newArray);
		return res;
	};
	cArray.prototype.getFirstElement = function () {
		return this.getElementRowCol(0,0);	
	};
	//check two-dimensional array
	cArray.prototype.checkValidArray = function (array, bConvertToValid) {
		if (!array || !array.length) {
			return false;
		}
		let isOneDimensional = null;
		for (let i = 0; i < array.length; i++) {
			if (Array.isArray(array[i])) {
				if (isOneDimensional) {
					return false;
				}
				for (let j = 0; j < array[i].length; j++) {
					if (Array.isArray(array[i][j])) {
						return false;
					}
				}
				isOneDimensional = false;
			} else if (isOneDimensional === null) {
				isOneDimensional = true;
			}
		}
		if (isOneDimensional && bConvertToValid) {
			let temp = [];
			temp.push(array);
			array = temp;
		}
		return array;
	};



	/**
	 * @constructor
	 * @extends {cBaseType}
	 */
	function cUndefined() {
		this.value = undefined;
	}

	cUndefined.prototype = Object.create(cBaseType.prototype);
	cUndefined.prototype.constructor = cUndefined;

	function checkTypeCell(cell, opt_toLowerCase) {
		if (cell && !cell.isNullText()) {
			var type = cell.getType();
			if (CellValueType.Number === type) {
				return new cNumber(cell.getNumberValue());
			} else {
				var val = cell.getValueWithoutFormat();
				if (CellValueType.Bool === type) {
					return new cBool(val);
				} else if (CellValueType.Error === type) {
					return new cError(val);
				} else {
					return new cString(opt_toLowerCase ? val.toLowerCase() : val);
				}
			}
		} else {
			return new cEmpty();
		}
	}

  /*--------------------------------------------------------------------------*/
	/*Base classes for operators & functions */
	/** @constructor */
	function cBaseOperator(name, priority, argumentCount) {
		this.name = name ? name : '';
		this.priority = (priority !== undefined) ? priority : 10;
		this.argumentsCurrent = (argumentCount !== undefined) ? argumentCount : 2;
		this.value = null;
	}

	cBaseOperator.prototype.type = cElementType.operator;
	cBaseOperator.prototype.numFormat = cNumFormatFirstCell;
	cBaseOperator.prototype.rightAssociative = false;
	cBaseOperator.prototype.toString = function () {
		return this.name;
	};
	cBaseOperator.prototype.Calculate = function () {
		return null;
	};
	cBaseOperator.prototype.Assemble2 = function (arg, start, count) {
		var str = "";
		if (this.argumentsCurrent === 2) {
			str += arg[start + count - 2] + this.name + arg[start + count - 1];
		} else {
			str += this.name + arg[start];
		}
		return new cString(str);
	};
	cBaseOperator.prototype.Assemble2Locale = function (arg, start, count, locale, digitDelim) {
		var str = "";
		if (this.argumentsCurrent === 2 && arg[start + count - 2] && arg[start + count - 1]) {
			str += arg[start + count - 2].toLocaleString(digitDelim) + this.name +
				arg[start + count - 1].toLocaleString(digitDelim);
		} else {
			str += this.name + arg[start];
		}
		return new cString(str);
	};
	cBaseOperator.prototype._convertAreaToArray = function (areaArr) {
		var res = [];
		for(var i = 0; i < areaArr.length; i++){
			var elem = areaArr[i];
			if(elem instanceof cArea || elem instanceof cArea3D){
				elem = convertAreaToArray(elem);
			}
			res.push(elem);
		}

		if(!res.length){
			res = areaArr;
		}

		return res;
	};
	cBaseOperator.prototype.tryDoArraysOperation = function (operand1, operand2, func) {
		//применяем в случае, если один или оба операнда area/array
		//возвращаем либо null, либо array
		var res = null;

		var dimension1 = operand1 && operand1.getDimensions();
		var dimension2 = operand2 && operand2.getDimensions();

		if (dimension1 && dimension2) {
			//берём наименьший размер, исключение - когда одна строка/столбец
			var colCount = dimension1.col === 1 ? dimension2.col : (dimension2.col === 1 ? dimension1.col : Math.min(dimension1.col, dimension2.col));
			var rowCount = dimension1.row === 1 ? dimension2.row : (dimension2.row === 1 ? dimension1.row : Math.min(dimension1.row, dimension2.row));

			var matrix1, matrix2;
			if (operand1.type === cElementType.array) {
				matrix1 = operand1;
			}
			if (operand2.type === cElementType.array) {
				matrix2 = operand2;
			}
			if (operand1.type === cElementType.cellsRange || operand1.type === cElementType.cellsRange3D) {
				matrix1 = convertAreaToArray(operand1);
			}
			if (operand2.type === cElementType.cellsRange || operand2.type === cElementType.cellsRange3D) {
				matrix2 = convertAreaToArray(operand2);
			}

			if (matrix1 || matrix2) {
				res = new cArray();
				for (var iRow = 0; iRow < rowCount; iRow++, iRow < rowCount ? res.addRow() : true) {
					for (var iCol = 0; iCol < colCount; iCol++) {
						var elem1 = matrix1 ? matrix1.getElementRowCol(dimension1.row === 1 ? 0 : iRow, dimension1.col === 1 ? 0 : iCol, true) : operand1;
						var elem2 = matrix2 ? matrix2.getElementRowCol(dimension2.row === 1 ? 0 : iRow, dimension2.col === 1 ? 0 : iCol, true) : operand2;
						res.addElement(func(elem1, elem2));
					}
				}
			}
		}

		return res;
	};

	/** @constructor */
	function cBaseFunction() {
	}

	cBaseFunction.prototype.type = cElementType.func;
	cBaseFunction.prototype.argumentsMin = 0;
	cBaseFunction.prototype.argumentsMax = 255;
	cBaseFunction.prototype.numFormat = cNumFormatFirstCell;
	cBaseFunction.prototype.ca = false;
	cBaseFunction.prototype.excludeHiddenRows = false;
	cBaseFunction.prototype.excludeErrorsVal = false;
	cBaseFunction.prototype.excludeNestedStAg = false;
	cBaseFunction.prototype.bArrayFormula = null;
	//необходимо для формул массива
	//arrayIndexes - мап, где ключ - аргумент, который в функцию передаётся в виде array,area,area3d (те неизменном виде)
	//а значение - либо булево, либо объект
	//объект пока содержит только информацию в том, что если внутри лежит индекс аргумента массива, то данный аргумент не воспринимается как массив
	//те подобный вид {1: 1, 2:{0: 1}} - означает, что 1 аргумент передаётся всегда как массив, а второй агумент зависит от того, является ли 0 аргумент массивом
	//returnValueType - ипользуется константа cReturnFormulaType
	cBaseFunction.prototype.arrayIndexes = null;
	cBaseFunction.prototype.returnValueType = null;
	cBaseFunction.prototype.inheritFormat = null;
	cBaseFunction.prototype.name = null;
	cBaseFunction.prototype.argumentsType = null;
	cBaseFunction.prototype.Calculate = function () {
		return new cError(cErrorType.wrong_name);
	};
	cBaseFunction.prototype.getArrayIndex = function (index) {
		let res = false;
		if (this.arrayIndexes) {
			res = this.arrayIndexes[index];
		}
		return res;
	};
	cBaseFunction.prototype.Assemble2 = function (arg, start, count) {

		var str = "", c = start + count - 1;
		for (var i = start; i <= c; i++) {
			if(!arg[i]) {
				continue;
			}
			str += arg[i].toString();
			if (i !== c) {
				str += ",";
			}
		}
		if (this.isXLFN || this.isXLWS) {
			return new cString((this.isXLFN ? "_xlfn." : "") + (this.isXLWS ? "_xlws." : "") + this.name + "(" + str + ")");
		} else if (this.isXLUDF) {
			//return new cString("__xludf.DUMMYFUNCTION." + this.name + "(" + str + ")");
		}
		return new cString(this.toString() + "(" + str + ")");
	};
	cBaseFunction.prototype.Assemble2Locale = function (arg, start, count, locale, digitDelim) {

		var name = this.toString(), str = "", c = start + count - 1, localeName = locale ? locale[name] : name;

		localeName = localeName || this.toString();
		for (var i = start; i <= c; i++) {
			if(!arg[i]) {
				continue;
			}
			str += arg[i].toLocaleString(digitDelim);
			if (i !== c) {
				str += FormulaSeparators.functionArgumentSeparator;
			}
		}
		return new cString(localeName + "(" + str + ")");
	};
	cBaseFunction.prototype.toString = function (/*locale*/) {
		/*var name = this.toString();
		var localeName = locale ? locale[name] : name;*/
		return this.name.replace(rx_sFuncPref, "_xlfn.").replace(rx_sFuncPrefXlWS, "_xlws.").replace(rx_sFuncPrefXLUFD, "__xludf.DUMMYFUNCTION.");
	};
	cBaseFunction.prototype.toLocaleString = function (/*locale*/) {
		var name = this.toString();
		//для cUnknownFunction делаем проверку
		if(AscCommonExcel.cFormulaFunctionToLocale && undefined !== AscCommonExcel.cFormulaFunctionToLocale[name]) {
			return AscCommonExcel.cFormulaFunctionToLocale[name];
		} else {
			return name;
		}
	};
	cBaseFunction.prototype.setCalcValue = function (arg, numFormat) {
		if (numFormat !== null && numFormat !== undefined) {
			arg.numFormat = numFormat;
		}
		return arg;
	};
	cBaseFunction.prototype.checkArguments = function (countArguments) {
		return this.argumentsMin <= countArguments && countArguments <= this.argumentsMax;
	};
	cBaseFunction.prototype._findArrayInNumberArguments = function (oArguments, calculateFunc, dNotCheckNumberType){
		var argsArray = [];
		var inputArguments = oArguments.args;
		var findArgArrayIndex = oArguments.indexArr;

		var parseArray = function(array){
			array.foreach(function (elem, r, c) {

				var arg;
				argsArray = [];
				for(var j = 0; j < inputArguments.length; j++){
					if(i === j){
						arg = elem;
					}else if(cElementType.array === inputArguments[j].type){
						arg = inputArguments[j].getElementRowCol(r, c);
					}else{
						arg = inputArguments[j];
					}

					if(arg && ((dNotCheckNumberType) || (cElementType.number === arg.type && !dNotCheckNumberType))){
						argsArray[j] = arg.getValue();
					}else{
						argsArray = null;
						break;
					}
				}

				this.array[r][c] = null === argsArray ? new cError(cErrorType.wrong_value_type) : calculateFunc(argsArray);
			});
			return array;
		};

		if(null !== findArgArrayIndex){
			return parseArray(inputArguments[findArgArrayIndex]);
		}else{
			for(var i = 0; i < inputArguments.length; i++){
				if(cElementType.string === inputArguments[i].type && !dNotCheckNumberType){
					return new cError(cErrorType.wrong_value_type);
				}else{
					if(inputArguments[i].getValue){
						argsArray[i] = inputArguments[i].getValue();
					}else{
						argsArray[i] = inputArguments[i];
					}
				}
			}
		}

		return calculateFunc(argsArray);
	};
	cBaseFunction.prototype._prepareArguments = function (args, arg1, bAddFirstArrElem, typeArray, bFirstRangeElem, notArrayError) {
		var newArgs = [];
		var indexArr = null;

		var excludeHiddenRows = this && this.excludeHiddenRows;
		var excludeErrorsVal = this && this.excludeErrorsVal;
		var excludeNestedStAg = this && this.excludeNestedStAg;
		for (var i = 0; i < args.length; i++) {
			var arg = args[i];

			//для массивов отдельная ветка
			if (typeArray && cElementType.array === typeArray[i]) {
				if (cElementType.cellsRange === arg.type || cElementType.array === arg.type) {
					newArgs[i] = arg.getMatrix(excludeHiddenRows, excludeErrorsVal, excludeNestedStAg);
				} else if (cElementType.cellsRange3D === arg.type) {
					newArgs[i] = arg.getMatrix(excludeHiddenRows, excludeErrorsVal, excludeNestedStAg)[0];
				} else if (cElementType.error === arg.type) {
					newArgs[i] = arg;
				} else {
					newArgs[i] = new cError(notArrayError ? notArrayError : cErrorType.division_by_zero);
				}
			} else if (cElementType.cellsRange === arg.type || cElementType.cellsRange3D === arg.type) {
				newArgs[i] = bFirstRangeElem ? arg.getValueByRowCol(0,0) : arg.cross(arg1);
				if (newArgs[i] == null) {
					newArgs[i] = arg.cross(arg1);
				}
			} else if (cElementType.array === arg.type) {
				if (bAddFirstArrElem) {
					newArgs[i] = arg.getElementRowCol(0, 0);
				} else {
					indexArr = i;
					newArgs[i] = arg;
				}
			} else {
				newArgs[i] = arg;
			}
		}

		return {args: newArgs, indexArr: indexArr};
	};
	cBaseFunction.prototype._checkErrorArg = function (argArray) {
		for (var i = 0; i < argArray.length; i++) {
			if (argArray[i] && cElementType.error === argArray[i].type) {
				return argArray[i];
			}
		}
		return null;
	};
	cBaseFunction.prototype._checkArrayArguments = function (arg0, func) {
		var matrix, res;
		if (arg0 instanceof cArea || arg0 instanceof cArray) {
			matrix = arg0.getMatrix();
		} else if (arg0 instanceof cArea3D) {
			matrix = arg0.getMatrix()[0];
		}

		if(matrix) {
			res = new cArray();
			for (var i = 0; i < matrix.length; ++i) {
				for (var j = 0; j < matrix[i].length; ++j) {
					matrix[i][j] = func(matrix[i][j]);
				}
			}
			res.fillFromArray(matrix);
		} else {
			res = func(arg0);
		}
		return res;
	};
	cBaseFunction.prototype._getOneDimensionalArray = function (arg, type) {
		var res = [];

		var getValue = function(curArg){
			if (undefined === type || cElementType.string === type){
				return curArg.tocString().getValue();
			} else if( cElementType.number === type){
				return curArg.tocNumber().getValue();
			} else if( cElementType.bool === type){
				return curArg.toLocaleString();
			}
		};

		if (cElementType.cellsRange === arg.type || cElementType.cellsRange3D === arg.type || cElementType.array === arg.type) {

			if (cElementType.cellsRange === arg.type || cElementType.array === arg.type) {
				arg = arg.getMatrix();
			} else if (cElementType.cellsRange3D === arg.type) {
				arg = arg.getMatrix()[0];
			}

			for (var i = 0; i < arg.length; i++) {
				for (var j = 0; j < arg[i].length; j++) {
					if(cElementType.error === arg[i][j].type){
						return arg[i][j];
					} else{
						res.push(getValue(arg[i][j]));
					}
				}
			}
		}else{
			if (cElementType.error === arg.type){
				return arg;
			} else{
				res.push(getValue(arg));
			}
		}

		return res;
	};
	cBaseFunction.prototype.checkRef = function (arg) {
		var res = false;
		if (cElementType.cell3D === arg.type || cElementType.cell === arg.type || cElementType.cellsRange === arg.type ||
			cElementType.cellsRange3D === arg.type) {
			res = true;
		}
		return res;
	};
	cBaseFunction.prototype.prepareAreaArg = function (arg, arguments1) {
		var res;

		if(this.bArrayFormula) {
			res = window['AscCommonExcel'].convertAreaToArray(arg);
		} else {
			res = arg.cross(arguments1);
		}

		return res;
	};
	cBaseFunction.prototype.calculateOneArgument = function(arg0, arguments1, func, convertAreaToArray) {
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			if(convertAreaToArray) {
				//***array-formula***
				arg0 = this.prepareAreaArg(arg0, arguments1);
			} else {
				arg0 = arg0.cross(arguments1);
			}
		}
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			var array = new cArray();
			arg0.foreach(function (elem, r, c) {
				if ( !array.array[r] ) {
					array.addRow();
				}
				array.addElement(func(elem));
			});
			return array;
		} else {
			return func(arg0);
		}
	};

	cBaseFunction.prototype.calculateTwoArguments = function(arg0, arg1, arguments1, func, convertAreaToArray) {

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			if(convertAreaToArray) {
				//***array-formula***
				arg0 = this.prepareAreaArg(arg0, arguments1);
			} else {
				arg0 = arg0.cross(arguments1);
			}
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			if(convertAreaToArray) {
				//***array-formula***
				arg1 = this.prepareAreaArg(arg1, arguments1);
			} else {
				arg1 = arg1.cross(arguments1);
			}
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
			if (arg0 instanceof cError) {
				return arg0;
			} else if (arg0 instanceof cString) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				arg0 = arg0.tocNumber();
			}
		} else {
			arg0 = arg0.tocNumber();
		}

		if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
			if (arg1 instanceof cError) {
				return arg1;
			} else if (arg1 instanceof cString) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				arg1 = arg1.tocNumber();
			}
		} else {
			arg1 = arg1.tocNumber();
		}

		var array;
		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			//TODO пересмотреть и упростить обработку
			array = new cArray();
			//в случае, если первый аргумент состоит из одно строки/столбца - тогда цикл по второму аргменту
			if(1 === arg0.getRowCount() || 1 === arg0.getCountElementInRow()) {
				arg1.foreach(function (elem, r, c) {
					var b = elem, res;
					//если аргумент - строка/столбец
					var rowArg1 = r, colArg1 = c;
					if(1 === arg0.getRowCount()) {
						rowArg1 = 0;
					}
					if(1 === arg0.getCountElementInRow()) {
						colArg1 = 0;
					}
					if ( !array.array[r] ) {
						array.addRow();
					}
					var a = arg0.array[rowArg1] ? arg0.getElementRowCol(rowArg1, colArg1) : null;
					if(!a) {
						res = new cError(cErrorType.not_available);
					} else if (a instanceof cNumber && b instanceof cNumber) {
						res = func(a.getValue(), b.getValue());
					} else {
						res = new cError(cErrorType.wrong_value_type);
					}
					array.addElement(res);
				});
				return array;
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem, res;
					var rowArg1 = r, colArg1 = c;
					if(1 === arg1.getRowCount()) {
						rowArg1 = 0;
					}
					if(1 === arg1.getCountElementInRow()) {
						colArg1 = 0;
					}
					if ( !array.array[r] ) {
						array.addRow();
					}
					var b = arg1.array[rowArg1] ? arg1.getElementRowCol(rowArg1, colArg1) : null;
					if(!b) {
						res = new cError(cErrorType.not_available);
					} else if (a instanceof cNumber && b instanceof cNumber) {
						res = func(a.getValue(), b.getValue());
					} else {
						res = new cError(cErrorType.wrong_value_type);
					}
					array.addElement(res);
				});
				return array;
			}
		} else if (arg0 instanceof cArray) {
			array = new cArray();
			arg0.foreach(function (elem, r, c) {
				var a = elem, res;
				var b = arg1;
				if ( !array.array[r] ) {
					array.addRow();
				}
				if (a instanceof cNumber && b instanceof cNumber) {
					res = func(a.getValue(), b.getValue())
				} else {
					res = new cError(cErrorType.wrong_value_type);
				}
				array.addElement(res);
			});
			return array;
		} else if (arg1 instanceof cArray) {
			array = new cArray();
			arg1.foreach(function (elem, r, c) {
				var a = arg0, res;
				var b = elem;
				if ( !array.array[r] ) {
					array.addRow();
				}
				if (a instanceof cNumber && b instanceof cNumber) {
					res = func(a.getValue(), b.getValue())
				} else {
					res = new cError(cErrorType.wrong_value_type);
				}
				array.addElement(res);
			});
			return array;
		} else {
			return func(arg0.getValue(), arg1.getValue());
		}

	};
	cBaseFunction.prototype.checkFormulaArray = function (arg, opt_bbox, opt_defName, parserFormula, bIsSpecialFunction, argumentsCount) {
		var res = null;
		var t = this;

		let dynamicRange = null, dynamicArraySize = null;
		if (AscCommonExcel.bIsSupportDynamicArrays) {
			if (!parserFormula.dynamicRange && !parserFormula.ref) {
				dynamicArraySize = this.getDynamicArraySize(arg);
				if (dynamicArraySize && parserFormula.parent && parserFormula.parent.nCol != null && parserFormula.parent.nRow != null) {
					dynamicRange = Asc.Range(parserFormula.parent.nCol, parserFormula.parent.nRow, dynamicArraySize.width + parserFormula.parent.nCol - 1, dynamicArraySize.height + parserFormula.parent.nRow - 1);
					// parserFormula.ref = dynamicRange;
					parserFormula.dynamicRange = dynamicRange;
					this.bArrayFormula = true;
				}
			}
		}

		var functionsCanReturnArray = ["index"];

		var returnFormulaType = this.returnValueType;
		if (cReturnFormulaType.setArrayRefAsArg === returnFormulaType) {
			if (arg.length === 0 && parserFormula.ref) {
				res = this.Calculate([new cArea(parserFormula.ref.getName(), parserFormula.ws)], opt_bbox, opt_defName, parserFormula.ws);
			} else {
				return null;
			}
		}

		var replaceAreaByValue = cReturnFormulaType.value_replace_area === returnFormulaType;
		var replaceAreaByRefs = cReturnFormulaType.area_to_ref === returnFormulaType;
		//добавлен специальный тип для функции сT, она использует из области всегда первый аргумент
		var replaceOnlyArray = cReturnFormulaType.replace_only_array === returnFormulaType;

		// Проверка должен ли элемент поступать в формулу без изменени?
		const checkArrayIndex = function(index) {
			let res = false;
			let arrayIndex = t.getArrayIndex(index);
			if(arrayIndex) {
				if(arrayIndex === arrayIndexesType.any) {
					res = true;
				} else if(typeof arrayIndex === "object") {
					//для данной проверки запрашиваем у объекта 0 индекс, там хранится значение индекса аргумента
					//от которого зависит стоит ли вопринимать данный аргумент как массив или нет
					let tempsArgIndex = arrayIndex[0];
					if(undefined !== tempsArgIndex && arg[tempsArgIndex]) {
						if(cElementType.cellsRange === arg[tempsArgIndex].type || cElementType.cellsRange3D === arg[tempsArgIndex].type || cElementType.array === arg[tempsArgIndex].type) {
							res = true;
						}
					}
				}
			}
			return res;
		};

		const checkArayIndexType = function(index, argType) {
			// check for type of argument - whether array and range can be processed or just one of them
			let res = false;
			let arrayIndex = t.getArrayIndex(index);
			if(argType === arrayIndex) {
				res = true;
			}
			return res;
		};

		const checkOneRowCol = function() {
			let res = false;
			for (let j = 0; j < argumentsCount; j++) {
				if(cElementType.array === arg[j].type) {
					if(1 === arg[j].getRowCount() || 1 === arg[j].getCountElementInRow()) {
						res = true;
					}
				} else {
					res = false;
					break;
				}
			}
			return res;
		};

		//bIsSpecialFunction - сделано только для для функции sumproduct
		//необходимо, чтобы все внутренние функции возвращали массив, те обрабатывались как формулы массива

		if((true === this.bArrayFormula || bIsSpecialFunction) && (!returnFormulaType || replaceAreaByValue || replaceAreaByRefs || this.arrayIndexes || replaceOnlyArray)) {

			if (functionsCanReturnArray.indexOf(this.name.toLowerCase()) !== -1) {
				var _tmp = this.Calculate(arg, opt_bbox, opt_defName, this.ws, bIsSpecialFunction);
				if (_tmp && _tmp.type === cElementType.array) {
					return _tmp;
				}
			}

			//вначале перебираем все аргументы и преобразовываем из cellsRange в массив или значение в зависимости от того, как должна работать функция
			var tempArgs = [], tempArg, firstArray, _checkArrayIndex;
			for (var j = 0; j < argumentsCount; j++) {
				tempArg = arg[j];

				_checkArrayIndex = checkArrayIndex(j);
				if (!_checkArrayIndex) {
					if (cElementType.cellsRange === tempArg.type || cElementType.cellsRange3D === tempArg.type) {
						if (checkArayIndexType(j, arrayIndexesType.range)) {
							// transfer range to argument without changing 
							tempArg = tempArg;
						} else if (replaceAreaByValue) {
							tempArg = tempArg.cross(opt_bbox);
						} else if (replaceAreaByRefs) {
							//добавляю специальные заглушки для функций row/column
							//они работают с аргументами иначе, чем все остальные
							//row - игнорируем в area колонки и проходимся только по строчкам и берём 1 колонку
							//к примеру, area A1:B2 разбиваем на [a1,a1;a2,a2] вместо нормального [a1,b1;a2,b2]
							var useOnlyFirstRow = "column" === this.name.toLowerCase() ? parserFormula.ref : null;
							var useOnlyFirstColumn = "row" === this.name.toLowerCase() ? parserFormula.ref : null;
							var _bbox = tempArg.getBBox0();
							if (useOnlyFirstRow) {
								firstArray = new Asc.Range(_bbox.c1, _bbox.r1, _bbox.c2, _bbox.r1);
							} else if (useOnlyFirstColumn) {
								firstArray = new Asc.Range(_bbox.c1, _bbox.r1, _bbox.c1, _bbox.r2);
							} else {
								tempArg = window['AscCommonExcel'].convertAreaToArrayRefs(tempArg, useOnlyFirstRow, useOnlyFirstColumn);
							}
						} else if(!replaceOnlyArray){
							tempArg = window['AscCommonExcel'].convertAreaToArray(tempArg);
						}
					}

					if (cElementType.array === tempArg.type) {
						if (checkArayIndexType(j, arrayIndexesType.array)) {
							// transfer array to argument without changing
							tempArg = tempArg;
						} else if (!firstArray) {	//пытаемся найти массив, которые имеет более 1 столбца и более 1 строки
							firstArray = tempArg;
						} else if((1 === firstArray.getRowCount() || 1 === firstArray.getCountElementInRow()) && 1 !== tempArg.getRowCount() && 1 !== tempArg.getCountElementInRow()) {
							firstArray = tempArg;
						} else if((1 === firstArray.getRowCount() && 1 === firstArray.getCountElementInRow()) && (1 !== tempArg.getRowCount() || 1 !== tempArg.getCountElementInRow())){
							firstArray = tempArg;
						}
					}
				}

				tempArgs.push(tempArg);
			}


			//для функций row/column с нулевым количеством аргументов необходимо рассчитывать
			//значение для каждой ячейки массива, изменяя при этом opt_bbox
			//TODO добавляю ещё одну проверку. в будущем стоит рассмотреть использование всегда parserFormula.ref
			//TODO персмотреть проверку isOneCell/checkOneRowCol - возможно стоит смотреть по количеству данных и расширять диапазон в случае, если parserFormula.ref превышает диапазон аргументов
			if ((replaceAreaByRefs && 0 === argumentsCount) || (!bIsSpecialFunction && firstArray && parserFormula.ref && !parserFormula.ref.isOneCell() && checkOneRowCol())) {
				firstArray = new cArray();
				firstArray.fillEmptyFromRange(parserFormula.ref);
			}

			if (firstArray) {
				var array = new cArray();
				//bbox_elem -
				var doCalc = function (elem, r, c, _row, _col) {
					if (!array.array[r]) {
						array.addRow();
					}

					//формируем новые аргументы(берем r/c элмент массива у каждого аргумента)
					var newArgs = [], newArg;
					for (var j = 0; j < argumentsCount; j++) {
						newArg = tempArgs[j];
						if (cElementType.array === newArg.type && !checkArrayIndex(j)) {
							if (1 === newArg.getRowCount() && 1 === newArg.getCountElementInRow()) {
								newArg = newArg.array[0] ? newArg.array[0][0] : null;
							} else if (1 === newArg.getRowCount()) {
								newArg = newArg.array[0] ? newArg.array[0][c] : null;
							} else if (1 === newArg.getCountElementInRow()) {
								newArg = newArg.array[r] ? newArg.array[r][0] : null;
							} else {
								newArg = newArg.array[r] ? newArg.array[r][c] : null;
							}
							if (!newArg) {
								//TODO проверить что ставить, если данный эламент массива недоступен
								//пока делаю так - если не последний аргумент, то пустой элемент, если последний - undefined
								newArg = /*j === argumentsCount - 1 ? undefined : */new cError(cErrorType.not_available);
							}
						}

						newArgs.push(newArg);
					}

					//для случая с 0 аргументов
					//возможно стоит убрать проверку на количество аргументови всегда заменять bbox
					var temp_opt_bbox = opt_bbox;
					if (0 === argumentsCount && parserFormula.ref) {
						temp_opt_bbox = new Asc.Range(c + parserFormula.ref.c1, r + parserFormula.ref.r1, c + parserFormula.ref.c1, r + parserFormula.ref.r1);
					}
					array.addElement(t.Calculate(newArgs, temp_opt_bbox, opt_defName, parserFormula.ws, null, _row ? _row : r, _col ? _col : c));
				};

				if (firstArray.foreach) {
					firstArray.foreach(doCalc);
				} else {
					//сделал заглушку для рассчета row()/col() функций. если по общей схему данные функции на вход
					//принимают только ref. перед тем как рассчитать формулу массива необходимо было сформировать
					//набор этих ref. поскольку этим функциям необходимы только номер строки/столбца -
					//передаём в функцию дополнительные параметры с этими данными
					for (var i = firstArray.r1; i <= firstArray.r2; i++) {
						for (var n = firstArray.c1; n <= firstArray.c2; n++) {
							doCalc(null, i - firstArray.r1, n - firstArray.c1, i, n);
						}
					}
				}


				res = array;

			} else if(replaceOnlyArray && tempArgs && tempArgs.length) {
				res = this.Calculate(tempArgs, opt_bbox, opt_defName, parserFormula.ws/*, bIsSpecialFunction*/);
			} else {
				res = this.Calculate(arg, opt_bbox, opt_defName, parserFormula.ws/*, bIsSpecialFunction*/);
			}
		}

		if (AscCommonExcel.bIsSupportDynamicArrays && (dynamicRange || dynamicArraySize)) {
			// parserFormula.ref = null;
			this.bArrayFormula = null;
		}

		return res;
	};

	cBaseFunction.prototype.checkFormulaArray2 = function (arg, opt_bbox, opt_defName, parserFormula, bIsSpecialFunction, argumentsCount) {
		// if (AscCommonExcel.bIsSupportDynamicArrays) {
			const t = this;
			let res = null;
			let functionsCanReturnArray = ["index"];

			let returnFormulaType = this.returnValueType;
			if (cReturnFormulaType.setArrayRefAsArg === returnFormulaType) {
				// todo check if this situation occurs
				if (arg.length === 0 && parserFormula.ref) {
					res = this.Calculate([new cArea(parserFormula.ref.getName(), parserFormula.ws)], opt_bbox, opt_defName, parserFormula.ws);
				} else {
					return null;
				}
			}

			let arrayIndexes = this.arrayIndexes;
			let replaceAreaByValue = cReturnFormulaType.value_replace_area === returnFormulaType;
			let replaceAreaByRefs = cReturnFormulaType.area_to_ref === returnFormulaType;
			let replaceOnlyArray = cReturnFormulaType.replace_only_array === returnFormulaType;

			const checkArrayIndex = function(index) {
				let res = false;
				if(arrayIndexes) {
					let arrayIndex = t.getArrayIndex(index);
					if(1 === arrayIndex) {
						res = true;
					} else if(typeof arrayIndex === "object") {
						// for this situation check object 0 for an index, the value of the argument index is stored there
						// which determines whether a given argument should be treated as an array or not
						let tempsArgIndex = arrayIndex[0];
						if(undefined !== tempsArgIndex && arg[tempsArgIndex]) {
							if(cElementType.cellsRange === arg[tempsArgIndex].type || cElementType.cellsRange3D === arg[tempsArgIndex].type || cElementType.array === arg[tempsArgIndex].type) {
								res = true;
							}
						}
					}
				}
				return res;
			};
			if((!returnFormulaType || replaceAreaByValue || replaceAreaByRefs || arrayIndexes || replaceOnlyArray)) {
				if (functionsCanReturnArray.indexOf(this.name.toLowerCase()) !== -1) {
					let _tmp = this.Calculate(arg, opt_bbox, opt_defName, this.ws, bIsSpecialFunction);
					if (_tmp && _tmp.type === cElementType.array) {
						return _tmp;
					}
				}

				let tempArgs = [], tempArg, firstArray, _checkArrayIndex;
				for (let j = 0; j < argumentsCount; j++) {
					tempArg = arg[j];

					_checkArrayIndex = checkArrayIndex(j);
					if (!_checkArrayIndex) {
						if (/*cElementType.cellsRange === tempArg.type || cElementType.cellsRange3D === tempArg.type ||*/ cElementType.array === tempArg.type) {
							res = true
						}
					}
					if (res) {
						return res;
					}

					tempArgs.push(tempArg);
				}
			}

			return res;
		// }
	};

	cBaseFunction.prototype.getDynamicArraySize = function (arg) {

		if (!AscCommonExcel.bIsSupportDynamicArrays || this.returnValueType === AscCommonExcel.cReturnFormulaType.array) {
			return null;
		}

		let width = 1, height = 1;
		for (let i = 0; i < arg.length; i++) {
			if (!this.arrayIndexes || !this.getArrayIndex(i)) {
				let objSize = arg[i].getDimensions();
				if (objSize) {
					height = Math.max(objSize.row, height);
					width = Math.max(objSize.col, width);
				}
			}
		}

		if (width !== 1 || height !== 1) {
			return {width: width, height: height};
		}
		return null;
	};
	cBaseFunction.prototype.checkArgumentsTypes = function (args) {
		if (args) {
			let length = args.length;
			for (let i = 0; i < length; i++) {
				let arg = args[i];
				if (arg && this.exactTypes[i] && this.argumentsType && this.argumentsType[i] !== undefined) {
					// check types
					if (this.argumentsType[i] === Asc.c_oAscFormulaArgumentType.reference && (arg.type !== cElementType.cellsRange && arg.type !== cElementType.cellsRange3D 
						&& arg.type !== cElementType.cell && arg.type !== cElementType.cell3D)) {
							return false;
					}
					// todo add other data types for arguments to the check, if the function requires it
				}
			}
		}
		return true;
	};

	/** @constructor */
	function cUnknownFunction(name) {
		this.name = name;
		this.isXLFN = null;
		this.isXLWS = null;
		this.isXLUDF = null;
	}
	cUnknownFunction.prototype = Object.create(cBaseFunction.prototype);
	cUnknownFunction.prototype.constructor = cUnknownFunction;


	/** @constructor */
	function parentLeft() {
	}

	parentLeft.prototype.type = cElementType.operator;
	parentLeft.prototype.name = "(";
	parentLeft.prototype.argumentsCurrent = 1;
	parentLeft.prototype.toString = function () {
		return this.name;
	};
	parentLeft.prototype.Assemble2 = function (arg, start, count) {
		return new cString("(" + arg[start + count - 1] + ")");
	};
	parentLeft.prototype.Assemble2Locale = function (arg, start, count, locale, digitDelim) {
		return new cString("(" + arg[start + count - 1].toLocaleString(digitDelim) + ")");
	};

	/** @constructor */
	function parentRight() {
	}

	parentRight.prototype.type = cElementType.operator;
	parentRight.prototype.name =  ")";
	parentRight.prototype.toString = function () {
		return this.name;
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cRangeUnionOperator() {
	}

	cRangeUnionOperator.prototype = Object.create(cBaseOperator.prototype);
	cRangeUnionOperator.prototype.constructor = cRangeUnionOperator;
	cRangeUnionOperator.prototype.name = ':';
	cRangeUnionOperator.prototype.priority = 50;
	cRangeUnionOperator.prototype.argumentsCurrent = 2;
	cRangeUnionOperator.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1], ws0, ws1, ws, res;
		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}
		if (( cElementType.cell === arg0.type || cElementType.cellsRange === arg0.type ||
			cElementType.cell3D === arg0.type ||
			cElementType.cellsRange3D === arg0.type && (ws0 = arg0.wsFrom) === arg0.wsTo ) &&
			( cElementType.cell === arg1.type || cElementType.cellsRange === arg1.type ||
			cElementType.cell3D === arg1.type ||
			cElementType.cellsRange3D === arg1.type && (ws1 = arg1.wsFrom) === arg1.wsTo )) {

			if (cElementType.cellsRange3D === arg0.type) {
				ws0 = ws = arg0.wsFrom;
			} else {
				ws0 = ws = arg0.getWS();
			}

			if (cElementType.cellsRange3D === arg1.type) {
				ws1 = ws = arg1.wsFrom;
			} else {
				ws1 = ws = arg1.getWS();
			}

			if (ws0 !== ws1) {
				return new cError(cErrorType.wrong_value_type);
			}

			arg0 = arg0.getBBox0();
			arg1 = arg1.getBBox0();
			if (!arg0 || !arg1) {
				return new cError(cErrorType.wrong_value_type);
			}
			arg0 = arg0.union(arg1);
			arg0.normalize(true);
			res = arg0.isOneCell() ? new cRef(arg0.getName(), ws) : new cArea(arg0.getName(), ws);
		} else {
			res = new cError(cErrorType.wrong_value_type);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cRangeIntersectionOperator() {
	}

	cRangeIntersectionOperator.prototype = Object.create(cBaseOperator.prototype);
	cRangeIntersectionOperator.prototype.constructor = cRangeIntersectionOperator;
	cRangeIntersectionOperator.prototype.name = ' ';
	cRangeIntersectionOperator.prototype.priority = 50;
	cRangeIntersectionOperator.prototype.argumentsCurrent = 2;
	cRangeIntersectionOperator.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], ws0, ws1, ws, res;
		if (( cElementType.cell === arg0.type || cElementType.cellsRange === arg0.type ||
			cElementType.cell3D === arg0.type ||
			cElementType.cellsRange3D === arg0.type && (ws0 = arg0.wsFrom) == arg0.wsTo ) &&
			( cElementType.cell === arg1.type || cElementType.cellsRange === arg1.type ||
			cElementType.cell3D === arg1.type ||
			cElementType.cellsRange3D === arg1.type && (ws1 = arg1.wsFrom) == arg1.wsTo )) {

			if (cElementType.cellsRange3D === arg0.type) {
				ws0 = ws = arg0.wsFrom;
			} else {
				ws0 = ws = arg0.getWS();
			}

			if (cElementType.cellsRange3D === arg1.type) {
				ws1 = ws = arg1.wsFrom;
			} else {
				ws1 = ws = arg1.getWS();
			}

			if (ws0 !== ws1) {
				return new cError(cErrorType.wrong_value_type);
			}

			arg0 = arg0.getBBox0();
			arg1 = arg1.getBBox0();
			if (!arg0 || !arg1) {
				return new cError(cErrorType.wrong_value_type);
			}
			arg0 = arg0.intersection(arg1);
			if (arg0) {
				arg0.normalize(true);
				res = arg0.isOneCell() ? new cRef(arg0.getName(), ws) : new cArea(arg0.getName(), ws);
			} else {
				res = new cError(cErrorType.null_value);
			}
		} else {
			res = new cError(cErrorType.wrong_value_type);
		}

		return res;
	};


	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cUnarMinusOperator() {
	}

	cUnarMinusOperator.prototype = Object.create(cBaseOperator.prototype);
	cUnarMinusOperator.prototype.constructor = cUnarMinusOperator;
	cUnarMinusOperator.prototype.name = 'un_minus';
	cUnarMinusOperator.prototype.priority = 49;
	cUnarMinusOperator.prototype.argumentsCurrent = 1;
	cUnarMinusOperator.prototype.rightAssociative = true;
	cUnarMinusOperator.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (arrElem, r, c) {
				arrElem = arrElem.tocNumber();
				arg0.array[r][c] = arrElem instanceof cError ? arrElem : new cNumber(-arrElem.getValue());
			});
			return arg0;
		}
		arg0 = arg0.tocNumber();
		return arg0 instanceof cError ? arg0 : new cNumber(-arg0.getValue());
	};
	cUnarMinusOperator.prototype.toString = function () {        // toString function
		return '-';
	};
	cUnarMinusOperator.prototype.Assemble2 = function (arg, start, count) {
		return new cString("-" + arg[start + count - 1]);
	};
	cUnarMinusOperator.prototype.Assemble2Locale = function (arg, start, count, locale, digitDelim) {
		return arg[start + count - 1].toLocaleString ?
			new cString("-" + arg[start + count - 1].toLocaleString(digitDelim)) :
			new cString("-" + arg[start + count - 1]);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cUnarPlusOperator() {
	}

	cUnarPlusOperator.prototype = Object.create(cBaseOperator.prototype);
	cUnarPlusOperator.prototype.constructor = cUnarPlusOperator;
	cUnarPlusOperator.prototype.name = 'un_plus';
	cUnarPlusOperator.prototype.priority = 49;
	cUnarPlusOperator.prototype.argumentsCurrent = 1;
	cUnarPlusOperator.prototype.rightAssociative = true;
	cUnarPlusOperator.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		return arg0;
	};
	cUnarPlusOperator.prototype.toString = function () {
		return '+';
	};
	cUnarPlusOperator.prototype.Assemble2 = function (arg, start, count) {
		return new cString("+" + arg[start + count - 1]);
	};
	cUnarPlusOperator.prototype.Assemble2Locale = function (arg, start, count, locale, digitDelim) {
		return arg[start + count - 1].toLocaleString ?
			new cString("+" + arg[start + count - 1].toLocaleString(digitDelim)) :
			new cString("+" + arg[start + count - 1]);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cAddOperator() {
	}

	cAddOperator.prototype = Object.create(cBaseOperator.prototype);
	cAddOperator.prototype.constructor = cAddOperator;
	cAddOperator.prototype.name = '+';
	cAddOperator.prototype.priority = 20;
	cAddOperator.prototype.argumentsCurrent = 2;
	cAddOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		}
		if (arg1 instanceof cArea) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		return _func[arg0.type][arg1.type](arg0, arg1, "+", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cMinusOperator() {
	}

	cMinusOperator.prototype = Object.create(cBaseOperator.prototype);
	cMinusOperator.prototype.constructor = cMinusOperator;
	cMinusOperator.prototype.name = '-';
	cMinusOperator.prototype.priority = 20;
	cMinusOperator.prototype.argumentsCurrent = 2;
	cMinusOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		}
		if (arg1 instanceof cArea) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		return _func[arg0.type][arg1.type](arg0, arg1, "-", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cPercentOperator() {
	}

	cPercentOperator.prototype = Object.create(cBaseOperator.prototype);
	cPercentOperator.prototype.constructor = cPercentOperator;
	cPercentOperator.prototype.name = '%';
	cPercentOperator.prototype.priority = 45;
	cPercentOperator.prototype.argumentsCurrent = 1;
	cPercentOperator.prototype.rightAssociative = true;
	cPercentOperator.prototype.Calculate = function (arg) {
		var res, arg0 = arg[0];
		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (arrElem, r, c) {
				arrElem = arrElem.tocNumber();
				arg0.array[r][c] = arrElem instanceof cError ? arrElem : new cNumber(arrElem.getValue() / 100);
			});
			return arg0;
		}
		arg0 = arg0.tocNumber();
		res = arg0 instanceof cError ? arg0 : new cNumber(arg0.getValue() / 100);
		res.numFormat = 9;
		return res;
	};
	cPercentOperator.prototype.Assemble2 = function (arg, start, count) {
		return new cString(arg[start + count - 1] + this.name);
	};
	cPercentOperator.prototype.Assemble2Locale = function (arg, start, count, locale, digitDelim) {
		return new cString(arg[start + count - 1].toLocaleString(digitDelim) + this.name);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cPowOperator() {
	}

	cPowOperator.prototype = Object.create(cBaseOperator.prototype);
	cPowOperator.prototype.numFormat = cNumFormatNone;
	cPowOperator.prototype.constructor = cPowOperator;
	cPowOperator.prototype.name = '^';
	cPowOperator.prototype.priority = 40;
	cPowOperator.prototype.argumentsCurrent = 2;
	cPowOperator.prototype.Calculate = function (arg) {
		let res = AscCommonExcel.cFormulaFunction["POWER"].prototype.Calculate(arg, arguments[1]);

		if (res) {
			return res;
		}

		return new cError(cErrorType.wrong_value_type);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cMultOperator() {
	}

	cMultOperator.prototype = Object.create(cBaseOperator.prototype);
	cMultOperator.prototype.numFormat = cNumFormatNone;
	cMultOperator.prototype.constructor = cMultOperator;
	cMultOperator.prototype.name = '*';
	cMultOperator.prototype.priority = 30;
	cMultOperator.prototype.argumentsCurrent = 2;
	cMultOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		}
		if (arg1 instanceof cArea) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		return _func[arg0.type][arg1.type](arg0, arg1, "*", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cDivOperator() {
	}

	cDivOperator.prototype = Object.create(cBaseOperator.prototype);
	cDivOperator.prototype.numFormat = cNumFormatNone;
	cDivOperator.prototype.constructor = cDivOperator;
	cDivOperator.prototype.name = '/';
	cDivOperator.prototype.priority = 30;
	cDivOperator.prototype.argumentsCurrent = 2;
	cDivOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		}
		if (arg1 instanceof cArea) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		return _func[arg0.type][arg1.type](arg0, arg1, "/", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cConcatSTROperator() {
	}

	cConcatSTROperator.prototype = Object.create(cBaseOperator.prototype);
	cConcatSTROperator.prototype.constructor = cConcatSTROperator;
	cConcatSTROperator.prototype.name = '&';
	cConcatSTROperator.prototype.priority = 15;
	cConcatSTROperator.prototype.argumentsCurrent = 2;
	cConcatSTROperator.prototype.numFormat = cNumFormatNone;
	cConcatSTROperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		var doOperation = function (_arg0, _arg1) {
			_arg0 = _arg0.tocString();
			_arg1 = _arg1.tocString();
			return _arg0 instanceof cError ? _arg0 :
				_arg1 instanceof cError ? _arg1 : new cString(_arg0.toString().concat(_arg1.toString()))
		};

		if(bIsSpecialFunction){
			var array = this.tryDoArraysOperation(arg0, arg1, doOperation);
			if (array) {
				return array;
			}
		}

		if (arg0 instanceof cArea) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		}
		if (arg1 instanceof cArea) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		}

		return doOperation(arg0, arg1);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cEqualsOperator() {
	}

	cEqualsOperator.prototype = Object.create(cBaseOperator.prototype);
	cEqualsOperator.prototype.constructor = cEqualsOperator;
	cEqualsOperator.prototype.name = '=';
	cEqualsOperator.prototype.priority = 10;
	cEqualsOperator.prototype.argumentsCurrent = 2;
	cEqualsOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		if (cElementType.cellsRange === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}
		return _func[arg0.type][arg1.type](arg0, arg1, "=", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cNotEqualsOperator() {
	}

	cNotEqualsOperator.prototype = Object.create(cBaseOperator.prototype);
	cNotEqualsOperator.prototype.constructor = cNotEqualsOperator;
	cNotEqualsOperator.prototype.name = '<>';
	cNotEqualsOperator.prototype.priority = 10;
	cNotEqualsOperator.prototype.argumentsCurrent = 2;
	cNotEqualsOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}

		if (cElementType.cellsRange === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}
		return _func[arg0.type][arg1.type](arg0, arg1, "<>", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cLessOperator() {
	}

	cLessOperator.prototype = Object.create(cBaseOperator.prototype);
	cLessOperator.prototype.constructor = cLessOperator;
	cLessOperator.prototype.name = '<';
	cLessOperator.prototype.priority = 10;
	cLessOperator.prototype.argumentsCurrent = 2;
	cLessOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}

		if (cElementType.cellsRange === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}
		return _func[arg0.type][arg1.type](arg0, arg1, "<", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cLessOrEqualOperator() {
	}

	cLessOrEqualOperator.prototype = Object.create(cBaseOperator.prototype);
	cLessOrEqualOperator.prototype.constructor = cLessOrEqualOperator;
	cLessOrEqualOperator.prototype.name = '<=';
	cLessOrEqualOperator.prototype.priority = 10;
	cLessOrEqualOperator.prototype.argumentsCurrent = 2;
	cLessOrEqualOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		if (cElementType.cellsRange === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}
		return _func[arg0.type][arg1.type](arg0, arg1, "<=", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cGreaterOperator() {
	}

	cGreaterOperator.prototype = Object.create(cBaseOperator.prototype);
	cGreaterOperator.prototype.constructor = cGreaterOperator;
	cGreaterOperator.prototype.name = '>';
	cGreaterOperator.prototype.priority = 10;
	cGreaterOperator.prototype.argumentsCurrent = 2;
	cGreaterOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		if (cElementType.cellsRange === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}
		return _func[arg0.type][arg1.type](arg0, arg1, ">", arguments[1], bIsSpecialFunction);
	};

	/**
	 * @constructor
	 * @extends {cBaseOperator}
	 */
	function cGreaterOrEqualOperator() {
	}

	cGreaterOrEqualOperator.prototype = Object.create(cBaseOperator.prototype);
	cGreaterOrEqualOperator.prototype.constructor = cGreaterOrEqualOperator;
	cGreaterOrEqualOperator.prototype.name = '>=';
	cGreaterOrEqualOperator.prototype.priority = 10;
	cGreaterOrEqualOperator.prototype.argumentsCurrent = 2;
	cGreaterOrEqualOperator.prototype.Calculate = function (arg, opt_bbox, opt_defName, ws, bIsSpecialFunction) {
		var arg0 = arg[0], arg1 = arg[1];

		if(bIsSpecialFunction){
			var convertArgs = this._convertAreaToArray([arg0, arg1]);
			arg0 = convertArgs[0];
			arg1 = convertArgs[1];
		}

		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		if (cElementType.cellsRange === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1], arguments[3]);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}
		return _func[arg0.type][arg1.type](arg0, arg1, ">=", arguments[1], bIsSpecialFunction);
	};

	/** @constructor */
	function cSpecialOperandStart() {
	}

	cSpecialOperandStart.prototype.constructor = cSpecialOperandStart;
	cSpecialOperandStart.prototype.type = cElementType.specialFunctionStart;

	/** @constructor */
	function cSpecialOperandEnd() {
	}

	cSpecialOperandEnd.prototype.constructor = cSpecialOperandEnd;
	cSpecialOperandEnd.prototype.type = cElementType.specialFunctionEnd;


	/* cFormulaOperators is container for holding all ECMA-376 operators, see chapter $18.17.2.2 in "ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference" */
	var cFormulaOperators = {
		'(': parentLeft,
		')': parentRight,
		'{': function () {
			var r = {};
			r.name = '{';
			r.toString = function () {
				return this.name;
			};
			return r;
		},
		'}': function () {
			var r = {};
			r.name = '}';
			r.toString = function () {
				return this.name;
			};
			return r;
		}, /* 50 is highest priority */
		':': cRangeUnionOperator,
		' ': cRangeIntersectionOperator,
		'un_minus': cUnarMinusOperator,
		'un_plus': cUnarPlusOperator,
		'%': cPercentOperator,
		'^': cPowOperator,
		'*': cMultOperator,
		'/': cDivOperator,
		'+': cAddOperator,
		'-': cMinusOperator,
		'&': cConcatSTROperator /*concat str*/,
		'=': cEqualsOperator/*equals*/,
		'<>': cNotEqualsOperator,
		'<': cLessOperator,
		'<=': cLessOrEqualOperator,
		'>': cGreaterOperator,
		'>=': cGreaterOrEqualOperator
		/* 10 is lowest priopity */
	};

	/* cFormulaFunctionGroup is container for holding all ECMA-376 function, see chapter $18.17.7 in "ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference" */
	/*
	 Каждая формула представляет собой копию функции cBaseFunction.
	 Для реализации очередной функции необходимо указать количество (минимальное и максимальное) принимаемых аргументов. Берем в спецификации.
	 Также необходино написать реализацию методов Calculate и getInfo(возвращает название функции и вид/количетво аргументов).
	 В методе Calculate необходимо отслеживать тип принимаемых аргументов. Для примера, если мы обращаемся к ячейке A1, в которой лежит 123, то этот аргумент будет числом. Если же там лежит "123", то это уже строка. Для более подробной информации смотреть спецификацию.
	 Метод getInfo является обязательным, ибо через этот метод в интерфейс передается информация о реализованных функциях.
	 */
	var cFormulaFunctionGroup = {};
	var cFormulaFunction = {};
	var cAllFormulaFunction = {};

	function getFormulasInfo() {

		var list = [], a, b, f;
		for (var type in cFormulaFunctionGroup) {
			b = new AscCommon.asc_CFormulaGroup(type);
			for (var i = 0; i < cFormulaFunctionGroup[type].length; ++i) {
				a = new cFormulaFunctionGroup[type][i]();
				//cFormulaFunctionGroup['NotRealised'] - массив ещё не реализованных формул
				if (-1 === cFormulaFunctionGroup['NotRealised'].indexOf(cFormulaFunctionGroup[type][i])) {
					f = new AscCommon.asc_CFormula(a);
					b.asc_addFormulaElement(f);
					cFormulaFunction[f.asc_getName()] = cFormulaFunctionGroup[type][i];
				}
				cAllFormulaFunction[a.name] = cFormulaFunctionGroup[type][i];
			}
			list.push(b);
		}
		return list;
	}
	function addNewFunction(func) {
		if (!func) {
			return;
		}
		let a = new func();
		let f = new AscCommon.asc_CFormula(a);
		cFormulaFunction[f.asc_getName()] = func;
		cAllFormulaFunction[a.name] = func;
	}
	function removeCustomFunction(sName) {
		if (!sName) {
			return;
		}
		delete cFormulaFunction[sName];
		delete cAllFormulaFunction[sName];
	}
	function getRangeByRef(ref, ws, onlyRanges, checkMultiSelection, checkFormula) {
		var activeCell = ws.getSelection().activeCell;
		var bbox = new Asc.Range(activeCell.col, activeCell.row, activeCell.col, activeCell.row);
		// ToDo in parser formula
		var ranges = [];

		var pushRanges = function(item) {
			var ref;
			switch (item.oper.type) {
				case cElementType.table:
				case cElementType.name:
				case cElementType.name3D:
					ref = item.oper.toRef(bbox, (checkMultiSelection && (item.oper.type === cElementType.name || item.oper.type === cElementType.name3D)));
					break;
				case cElementType.cell:
				case cElementType.cell3D:
				case cElementType.cellsRange:
				case cElementType.cellsRange3D:
					ref = item.oper;
					break;
			}
			if (ref) {
				var pushRange = function(curRef) {
					switch(curRef.type) {
						case cElementType.cell:
						case cElementType.cell3D:
						case cElementType.cellsRange:
						case cElementType.cellsRange3D:
							ranges.push(curRef.getRange());
							break;
						case cElementType.array:
							if (!onlyRanges) {
								ranges = curRef.getMatrix();
							}
							break;
					}
				};

				if(ref.length) {
					for(var i = 0; i < ref.length; i++) {
						pushRange(ref[i]);
					}
				} else {
					pushRange(ref);
				}
			}
		};

		//TODO вызываю проверку на то, что это может быть формула только для печати. необходимо проверить везде - для этого необходимо просмотреть весь смежный функционал
		var isFormula;
		if(checkFormula && ref) {
			var parseResult = new AscCommonExcel.ParseResult([]);
			var parsed = new AscCommonExcel.parserFormula(ref, null, ws);
			parsed.parse(undefined, undefined, parseResult);
			isFormula = parsed.calculate();
		}

		if (isFormula && isFormula.type !== cElementType.error) {
			pushRanges({oper: isFormula});
		} else {
			// ToDo in parser formula
			if (ref[0] === '(') {
				ref = ref.slice(1);
			}
			if (ref[ref.length - 1] === ')') {
				ref = ref.slice(0, -1);
			}

			var arrRefs = ref.split(',');
			arrRefs.forEach(function (refItem) {
				// ToDo in parser formula
				var currentWorkbook = '[0]!';
				if (0 === refItem.indexOf(currentWorkbook)) {
					refItem = refItem.slice(currentWorkbook.length);
				}

				var _f = new AscCommonExcel.parserFormula(refItem, null, ws);
				var parseResult = new AscCommonExcel.ParseResult([]);
				if (_f.parse(null, null, parseResult)) {
					parseResult.refPos.forEach(pushRanges);
				}
			});
		}

		return ranges;
	}


	function getRangeByName(sName, ws) {
		// Early validation
		if (!sName || !ws) {
			return [];
		}

		// Initialize arrays
		const ranges = [];
		const activeCell = ws.getSelection().activeCell;
		const bbox = new Asc.Range(activeCell.col, activeCell.row, activeCell.col, activeCell.row);

		// Helper function to process names
		const processName = function(item) {
			if (item.oper.type === cElementType.name) {
				const nameRef = item.oper.toRef(bbox);

				// Skip if reference is error
				if (nameRef instanceof AscCommonExcel.cError) {
					return;
				}

				// Get range from valid reference
				if (nameRef.getRange && nameRef.getWS && nameRef.getWS() === ws) {
					ranges.push(nameRef.getRange());
				}
			}
		};

		// Parse formula and get references
		const parseResult = new AscCommonExcel.ParseResult([]);
		const parsed = new AscCommonExcel.parserFormula(sName, null, ws);

		if (parsed.parse(undefined, undefined, parseResult)) {
			// Process only name type references
			parseResult.refPos.forEach(processName);
		}

		return ranges;
	}

/*--------------------------------------------------------------------------*/


var _func = [];//для велосипеда а-ля перегрузка функций.
_func[cElementType.number] = [];
_func[cElementType.string] = [];
_func[cElementType.bool] = [];
_func[cElementType.error] = [];
_func[cElementType.cellsRange] = [];
_func[cElementType.empty] = [];
_func[cElementType.array] = [];
_func[cElementType.cell] = [];


_func[cElementType.number][cElementType.number] = function ( arg0, arg1, what ) {
	var compareNumbers = function(){
		return AscCommon.compareNumbers(arg0.getValue(), arg1.getValue());
	};
	let opt_return_bool = arguments[5];
	let res = null;
	if (what === ">") {
		res = compareNumbers() > 0;
	} else if (what === ">=") {
		res = !(compareNumbers() < 0);
	} else if (what === "<") {
		res = compareNumbers() < 0;
	} else if (what === "<=") {
		res = !(compareNumbers() > 0);
	} else if (what === "=") {
		res = (compareNumbers() === 0);
	} else if (what === "<>") {
		res = (compareNumbers() !== 0);
	} else if (what === "-") {
		return new cNumber(arg0.getValue() - arg1.getValue());
	} else if (what === "+") {
		return new cNumber(arg0.getValue() + arg1.getValue());
	} else if (what === "/") {
		if (arg1.getValue() !== 0) {
			return new cNumber(arg0.getValue() / arg1.getValue());
		} else {
			return new cError(cErrorType.division_by_zero);
		}
	} else if (what === "*") {
		return new cNumber(arg0.getValue() * arg1.getValue());
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
    return new cError( cErrorType.wrong_value_type );
};

_func[cElementType.number][cElementType.string] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;
	if (what === ">" || what === ">=") {
		res = false;
	} else if (what === "<" || what === "<=") {
		res = true;
	} else if (what === "=") {
		res = false;
	} else if (what === "<>") {
		res = true;
	} else if (what === "-" || what === "+" || what === "/" || what === "*") {
		return new cError(cErrorType.wrong_value_type);
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.number][cElementType.bool] = function ( arg0, arg1, what ) {
	let _arg;
	let res = null;
	let opt_return_bool = arguments[5];
	if (what === ">" || what === ">=") {
		res = false;
	} else if (what === "<" || what === "<=") {
		res = true;
	} else if (what === "=") {
		res = false;
	} else if (what === "<>") {
		res = true;
	} else if (what === "-") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		return new cNumber(arg0.getValue() - _arg.getValue());
	} else if (what === "+") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		return new cNumber(arg0.getValue() + _arg.getValue());
	} else if (what === "/") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		if (_arg.getValue() !== 0) {
			return new cNumber(arg0.getValue() / _arg.getValue());
		} else {
			return new cError(cErrorType.division_by_zero);
		}
	} else if (what === "*") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		return new cNumber(arg0.getValue() * _arg.getValue());
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.number][cElementType.error] = function ( arg0, arg1 ) {
    return arg1;
};

_func[cElementType.number][cElementType.empty] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;
	if (what === ">") {
		res = arg0.getValue() > 0;
	} else if (what === ">=") {
		res = arg0.getValue() >= 0;
	} else if (what === "<") {
		res = arg0.getValue() < 0;
	} else if (what === "<=") {
		res = arg0.getValue() <= 0;
	} else if (what === "=") {
		res = arg0.getValue() === 0;
	} else if (what === "<>") {
		res = arg0.getValue() !== 0;
	} else if (what === "-") {
		return new cNumber(arg0.getValue() - 0);
	} else if (what === "+") {
		return new cNumber(arg0.getValue() + 0);
	} else if (what === "/") {
		return new cError(cErrorType.division_by_zero);
	} else if (what === "*") {
		return new cNumber(0);
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};


_func[cElementType.string][cElementType.number] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;
	if (what === ">" || what === ">=") {
		res = true;
	} else if (what === "<" || what === "<=" || what === "=") {
		res = false;
	} else if (what === "<>") {
		res = true;
	} else if (what === "-" || what === "+" || what === "/" || what === "*") {
		return new cError(cErrorType.wrong_value_type);
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.string][cElementType.string] = function ( arg0, arg1, what ) {
	//TODO need change opt_return_bool. for example >  change on props -> .returnBool & .useLowerCase & ...
	let opt_return_bool = arguments[5];
	let res = null;

	let isEqualStrings = function (str1, str2) {
		return opt_return_bool ? str1 === str2 : str1.toLowerCase() === str2.toLowerCase();
	};

	let _arg0, _arg1;
	if (what === ">") {
		res = arg0.getValue(true) > arg1.getValue(true);
	} else if (what === ">=") {
		res = arg0.getValue(true) >= arg1.getValue(true);
	} else if (what === "<") {
		res = arg0.getValue(true) < arg1.getValue(true);
	} else if (what === "<=") {
		res = arg0.getValue(true) <= arg1.getValue(true);
	} else if (what === "=") {
		res = isEqualStrings(arg0.getValue(true), arg1.getValue(true));
	} else if (what === "<>") {
		res = !isEqualStrings(arg0.getValue(true), arg1.getValue(true));
	} else if (what === "-") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue(true) - _arg1.getValue(true));
	} else if (what === "+") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue(true) + _arg1.getValue(true));
	} else if (what === "/") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		if (_arg1.getValue(true) !== 0) {
			return new cNumber(_arg0.getValue(true) / _arg1.getValue(true));
		}
		return new cError(cErrorType.division_by_zero);
	} else if (what === "*") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue(true) * _arg1.getValue(true));
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.string][cElementType.bool] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;
	let _arg0, _arg1;
	if (what === ">" || what === ">=") {
		res = false;
	} else if (what === "<" || what === "<=") {
		res = true;
	} else if (what === "=") {
		res = false;
	} else if (what === "<>") {
		res = true;
	} else if (what === "-") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue() - _arg1.getValue());
	} else if (what === "+") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue() + _arg1.getValue());
	} else if (what === "/") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		if (_arg1.getValue() !== 0) {
			return new cNumber(_arg0.getValue() / _arg1.getValue());
		}
		return new cError(cErrorType.division_by_zero);
	} else if (what === "*") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg0 instanceof cError) {
			return _arg0;
		}
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue() * _arg1.getValue());
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.string][cElementType.error] = function ( arg0, arg1 ) {
    return arg1;
};

_func[cElementType.string][cElementType.empty] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;

	if (what === ">") {
		res = arg0.getValue(true).length !== 0;
	} else if (what === ">=") {
		res = arg0.getValue(true).length >= 0;
	} else if (what === "<") {
		res = false;
	} else if (what === "<=") {
		res = arg0.getValue(true).length <= 0;
	} else if (what === "=") {
		res = arg0.getValue(true).length === 0;
	} else if (what === "<>") {
		res = arg0.getValue(true).length !== 0;
	} else if (what === "-" || what === "+" || what === "/" || what === "*") {
		return new cError(cErrorType.wrong_value_type);
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};


_func[cElementType.bool][cElementType.number] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;

	var _arg;
	if (what === ">" || what === ">=") {
		res = true;
	} else if (what === "<" || what === "<=") {
		res = false;
	} else if (what === "=") {
		res = false;
	} else if (what === "<>") {
		res = true;
	} else if (what === "-") {
		_arg = arg0.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		return new cNumber(_arg.getValue() - arg1.getValue());
	} else if (what === "+") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		return new cNumber(_arg.getValue() + arg1.getValue());
	} else if (what === "/") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		if (arg1.getValue() !== 0) {
			return new cNumber(_arg.getValue() / arg1.getValue());
		} else {
			return new cError(cErrorType.division_by_zero);
		}
	} else if (what === "*") {
		_arg = arg1.tocNumber();
		if (_arg instanceof cError) {
			return _arg;
		}
		return new cNumber(_arg.getValue() * arg1.getValue());
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.bool][cElementType.string] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;

	var _arg0, _arg1;
	if (what === ">" || what === ">=") {
		res= true;
	} else if (what === "<" || what === "<=") {
		res= false;
	} else if (what === "=") {
		res = false;
	} else if (what === "<>") {
		res= true;
	} else if (what === "-") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue() - _arg1.getValue());
	} else if (what === "+") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue() + _arg1.getValue());
	} else if (what === "/") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		if (_arg1.getValue() !== 0) {
			return new cNumber(_arg0.getValue() / _arg1.getValue());
		}
		return new cError(cErrorType.division_by_zero);
	} else if (what === "*") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		if (_arg1 instanceof cError) {
			return _arg1;
		}
		return new cNumber(_arg0.getValue() * _arg1.getValue());
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};

_func[cElementType.bool][cElementType.bool] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;

	var _arg0, _arg1;
	if (what === ">") {
		res =arg0.value > arg1.value;
	} else if (what === ">=") {
		res =arg0.value >= arg1.value;
	} else if (what === "<") {
		res =arg0.value < arg1.value;
	} else if (what === "<=") {
		res =arg0.value <= arg1.value;
	} else if (what === "=") {
		res =arg0.value === arg1.value;
	} else if (what === "<>") {
		res =arg0.value !== arg1.value;
	} else if (what === "-") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		return new cNumber(_arg0.getValue() - _arg1.getValue());
	} else if (what === "+") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		return new cNumber(_arg0.getValue() + _arg1.getValue());
	} else if (what === "/") {
		if (!arg1.value) {
			return new cError(cErrorType.division_by_zero);
		}
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		return new cNumber(_arg0.getValue() / _arg1.getValue());
	} else if (what === "*") {
		_arg0 = arg0.tocNumber();
		_arg1 = arg1.tocNumber();
		return new cNumber(_arg0.getValue() * _arg1.getValue());
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
    return new cError( cErrorType.wrong_value_type );
};

_func[cElementType.bool][cElementType.error] = function ( arg0, arg1 ) {
    return arg1;
};

_func[cElementType.bool][cElementType.empty] = function ( arg0, arg1, what ) {
	let opt_return_bool = arguments[5];
	let res = null;

	if (what === ">") {
		res= arg0.value > false;
	} else if (what === ">=") {
		res= arg0.value >= false;
	} else if (what === "<") {
		res= arg0.value < false;
	} else if (what === "<=") {
		res= arg0.value <= false;
	} else if (what === "=") {
		res= arg0.value === false;
	} else if (what === "<>") {
		res= arg0.value !== false;
	} else if (what === "-") {
		res = arg0.value ? 1 : 0;
	} else if (what === "+") {
		res = arg0.value ? 1 : 0;
	} else if (what === "/") {
		return new cError(cErrorType.division_by_zero);
	} else if (what === "*") {
		return new cNumber(0);
	}
	if (res !== null) {
		return opt_return_bool ? res : new cBool(res);
	}
	return new cError(cErrorType.wrong_value_type);
};


_func[cElementType.error][cElementType.number] = _func[cElementType.error][cElementType.string] =
  _func[cElementType.error][cElementType.bool] =
    _func[cElementType.error][cElementType.error] = _func[cElementType.error][cElementType.empty] = function(arg0) {
            return arg0;
        };


_func[cElementType.empty][cElementType.number] = function ( arg0, arg1, what ) {
    if ( what === ">" ) {
        return new cBool( 0 > arg1.getValue() );
  } else if (what === ">=") {
        return new cBool( 0 >= arg1.getValue() );
  } else if (what === "<") {
        return new cBool( 0 < arg1.getValue() );
  } else if (what === "<=") {
        return new cBool( 0 <= arg1.getValue() );
  } else if (what === "=") {
        return new cBool( 0 === arg1.getValue() );
  } else if (what === "<>") {
        return new cBool( 0 !== arg1.getValue() );
  } else if (what === "-") {
        return new cNumber( 0 - arg1.getValue() );
  } else if (what === "+") {
        return new cNumber( 0 + arg1.getValue() );
  } else if (what === "/") {
        if ( arg1.getValue() === 0 ) {
            return new cError( cErrorType.not_numeric );
        }
        return new cNumber( 0 );
  } else if (what === "*") {
        return new cNumber( 0 );
    }
    return new cError( cErrorType.wrong_value_type );
};

_func[cElementType.empty][cElementType.string] = function ( arg0, arg1, what ) {
    if ( what === ">" ) {
        return new cBool( 0 > arg1.getValue(true).length );
  } else if (what === ">=") {
        return new cBool( 0 >= arg1.getValue(true).length );
  } else if (what === "<") {
        return new cBool( 0 < arg1.getValue(true).length );
  } else if (what === "<=") {
        return new cBool( 0 <= arg1.getValue(true).length );
  } else if (what === "=") {
        return new cBool( 0 === arg1.getValue(true).length );
  } else if (what === "<>") {
        return new cBool( 0 !== arg1.getValue(true).length );
  } else if (what === "-" || what === "+" || what === "/" || what === "*") {
        return new cError( cErrorType.wrong_value_type );
    }
    return new cError( cErrorType.wrong_value_type );
};

_func[cElementType.empty][cElementType.bool] = function ( arg0, arg1, what ) {
    if ( what === ">" ) {
        return new cBool( false > arg1.value );
  } else if (what === ">=") {
        return new cBool( false >= arg1.value );
  } else if (what === "<") {
        return new cBool( false < arg1.value );
  } else if (what === "<=") {
        return new cBool( false <= arg1.value );
  } else if (what === "=") {
        return new cBool( arg1.value === false );
  } else if (what === "<>") {
        return new cBool( arg1.value !== false );
  } else if (what === "-") {
        return new cNumber( 0 - arg1.value ? 1.0 : 0.0 );
  } else if (what === "+") {
        return new cNumber( arg1.value ? 1.0 : 0.0 );
  } else if (what === "/") {
        if ( arg1.value ) {
            return new cNumber( 0 );
        }
        return new cError( cErrorType.not_numeric );
  } else if (what === "*") {
        return new cNumber( 0 );
    }
    return new cError( cErrorType.wrong_value_type );
};

_func[cElementType.empty][cElementType.error] = function ( arg0, arg1 ) {
    return arg1;
};

_func[cElementType.empty][cElementType.empty] = function ( arg0, arg1, what ) {
    if ( what === ">" || what === "<" || what === "<>" ) {
        return new cBool( false );
  } else if (what === ">=" || what === "<=" || what === "=") {
        return new cBool( true );
  } else if (what === "-" || what === "+") {
        return new cNumber( 0 );
  } else if (what === "/") {
        return new cError( cErrorType.not_numeric );
  } else if (what === "*") {
        return new cNumber( 0 );
    }
    return new cError( cErrorType.wrong_value_type );
};


_func[cElementType.cellsRange][cElementType.number] = _func[cElementType.cellsRange][cElementType.string] =
    _func[cElementType.cellsRange][cElementType.bool] = _func[cElementType.cellsRange][cElementType.error] =
    _func[cElementType.cellsRange][cElementType.array] =
      _func[cElementType.cellsRange][cElementType.empty] = function(arg0, arg1, what, bbox) {
            var cross = arg0.cross( bbox );
            return _func[cross.type][arg1.type]( cross, arg1, what );
        };


_func[cElementType.number][cElementType.cellsRange] = _func[cElementType.string][cElementType.cellsRange] =
    _func[cElementType.bool][cElementType.cellsRange] = _func[cElementType.error][cElementType.cellsRange] =
    _func[cElementType.array][cElementType.cellsRange] =
      _func[cElementType.empty][cElementType.cellsRange] = function(arg0, arg1, what, bbox) {
            var cross = arg1.cross( bbox );
            return _func[arg0.type][cross.type]( arg0, cross, what );
        };


_func[cElementType.cellsRange][cElementType.cellsRange] = function ( arg0, arg1, what, bbox ) {
  var cross1 = arg0.cross(bbox), cross2 = arg1.cross(bbox);
    return _func[cross1.type][cross2.type]( cross1, cross2, what );
};

_func[cElementType.array][cElementType.array] = function ( arg0, arg1, what, bbox, bIsSpecialFunction ) {
	if (bIsSpecialFunction) {
		let specialArray = specialFuncArrayToArray(arg0, arg1, what);
		if(null !== specialArray){
			return specialArray;
		}
	}
	if ( arg0.getRowCount() !== arg1.getRowCount() || arg0.getCountElementInRow() !== arg1.getCountElementInRow() ) {
        return new cError( cErrorType.wrong_value_type );
    }
    var retArr = new cArray(), _arg0, _arg1;
    for ( var iRow = 0; iRow < arg0.getRowCount(); iRow++, iRow < arg0.getRowCount() ? retArr.addRow() : true ) {
        for ( var iCol = 0; iCol < arg0.getCountElementInRow(); iCol++ ) {
            _arg0 = arg0.getElementRowCol( iRow, iCol );
            _arg1 = arg1.getElementRowCol( iRow, iCol );
            retArr.addElement( _func[_arg0.type][_arg1.type]( _arg0, _arg1, what ) );
        }
    }
    return retArr;
};

_func[cElementType.array][cElementType.number] = _func[cElementType.array][cElementType.string] =
    _func[cElementType.array][cElementType.bool] = _func[cElementType.array][cElementType.error] =
        _func[cElementType.array][cElementType.empty] = function ( arg0, arg1, what ) {
            let res = new cArray(), realArraySize, rowDiff, colDiff, funcResult, arrayDimensions = arg0.getDimensions();

			if (arg0.realSize && arg0.missedValue) {
				realArraySize = arg0.getRealArraySize();
				rowDiff = realArraySize.row - arg0.getRowCount();
				colDiff = realArraySize.col - arg0.getCountElementInRow();
				funcResult = _func[arg0.missedValue.type][arg1.type](arg0.missedValue, arg1, what);

				// set realSize to res
				res.setRealArraySize(realArraySize.row, realArraySize.col);
				res.missedValue = funcResult;
			}

			arg0.foreach( function ( elem, r ) {
                if ( !res.array[r] ) {
                    res.addRow();
                }
                res.addElement( _func[elem.type][arg1.type]( elem, arg1, what ) );
            } );

            return res;
        };

_func[cElementType.number][cElementType.array] = _func[cElementType.string][cElementType.array] =
    _func[cElementType.bool][cElementType.array] = _func[cElementType.error][cElementType.array] =
        _func[cElementType.empty][cElementType.array] = function ( arg0, arg1, what ) {
			let res = new cArray(), realArraySize, rowDiff, colDiff, funcResult, arrayDimensions = arg1.getDimensions();

			if (arg1.realSize && arg1.missedValue) {
				realArraySize = arg1.getRealArraySize();
				rowDiff = realArraySize.row - arg1.getRowCount();
				colDiff = realArraySize.col - arg1.getCountElementInRow();
				funcResult = _func[arg0.type][arg1.missedValue.type](arg0, arg1.missedValue, what);

				// set realSize to res
				res.setRealArraySize(realArraySize.row, realArraySize.col);
				res.missedValue = funcResult;
			}

			arg1.foreach( function ( elem, r ) {
                if ( !res.array[r] ) {
                    res.addRow();
                }
                res.addElement( _func[arg0.type][elem.type]( arg0, elem, what ) );
            } );

            return res;
        };


_func.binarySearch = function ( sElem, arrTagert, regExp ) {
	var first = 0, /* The number of the first element in the array */
		last = arrTagert.length - 1, /* The number of the element in the array that comes AFTER the last one */
		/* If the viewed segment is not empty, first<last */
		mid;

	var arrTagertOneType = [], isString = false;

	for (var i = 0; i < arrTagert.length; i++) {
		if ((arrTagert[i] instanceof cString || sElem instanceof cString) && !isString) {
			i = 0;
			isString = true;
			sElem = new cString(sElem.toString().toLowerCase());
		}
		if (isString) {
			arrTagertOneType[i] = new cString(arrTagert[i].toString().toLowerCase());
		} else {
			arrTagertOneType[i] = arrTagert[i].tocNumber();
		}
	}

	// comparing the lengths of arrays and the first and last element
	if (arrTagert.length === 0) {
		return -1;
		/* array empty */
	} else if (arrTagert[0].value > sElem.value) {
		return -2;
	} else if (arrTagert[arrTagert.length - 1].value < sElem.value) {
		return arrTagert.length - 1;
	}

	// according to the sorting in MS, the comparison will be like this: cError > cBool > cText > (cNumber == cEmpty)
	while (first < last) {
		mid = Math.floor(first + (last - first) / 2);
		if (sElem.type !== arrTagert[mid].type) {
			if (sElem.type === cElementType.empty || arrTagert[mid].type === cElementType.empty) {
				if (sElem.value <= arrTagert[mid].value) {
					// cEmpty.tocNumber() ?
					last = mid;
				} else {
					first = mid + 1;
				}
			} else {
				if (cElementTypeWeight.get(sElem.type) < cElementTypeWeight.get(arrTagert[mid].type)) {
					last = mid;
				} else {
					first = mid + 1;
				}
			}
		} else {
			// if cError && cError ?
			if (sElem.value < arrTagert[mid].value || ( regExp && regExp.test(arrTagert[mid].value) )) {
				last = mid;
			} else {
				first = mid + 1;
			}
		}
	}

	/* If the conditional operator if(n==0) and so on is omitted at the beginning - then uncomment it here!    */
	if (/* last<n &&*/ arrTagert[last].value === sElem.value) {
		return last;
		/* The desired element is found. last is the desired index */
	} else {
		return last - 1;
		/* The desired element is not found. But if you suddenly need to insert it with a shift, its place is at last.    */
	}

};

_func.lookupBinarySearch = function ( sElem, arrayNoEmpty, isByRangeCall, regExp ) {
	let first = 0, last, mid;
	let typedArr;

	typedArr = prepareTypedArrayUniversal(arrayNoEmpty, sElem, isByRangeCall);
	
	if (typedArr.length === 0) {
		/* array empty */
		return -1;
	}
	// 2 elements next to each other
	if (typedArr.length === 2) {
		// todo check two element behaviour
	}
	// With 0-9 < A-Z, if query is numeric and data found is string, or
	// vice versa, the (yet another undocumented) Excel behavior is to
	// return #N/A instead.

	if (sElem.type === cElementType.string) {
		sElem = new cString(sElem.toString().toLowerCase());
	}

	let cacheIndex, isFound;
	first = 0, last = typedArr.length - 1;
	while (first < last) {
		mid = Math.floor(first + (last - first) / 2);

		let midValue = typedArr[mid].v;
		// let cmp = compareValues(sElem, midValue)
		if (sElem.value === midValue.value) {
			/* cmp === 0 */
			last = _func.getLastMatch(mid, sElem, typedArr);
			break;
		}

		if (sElem.value < midValue.value || ( regExp && regExp.test(midValue.value) )) {
			/* cmp > 0 */
			last = mid;
		} else {
			/* cmp < 0 */	
			cacheIndex = mid;														
			first = mid + 1;
		}
	}

	if (typedArr[last].v.value <= sElem.value) {
		return typedArr[last].i;
	} else if (cacheIndex !== undefined && typedArr[cacheIndex].v.value <= sElem.value) {
		return typedArr[cacheIndex].i;
	} else {
		return -2;
	}
};

_func.getLastMatch = function (startIndex, lookingElem, array) {
	// todo add compare to all types?
	let resIndex = startIndex, exactMatchIndex;
	for (let i = startIndex; i < array.length; i++) {
		if (array[i].v.type !== lookingElem.type) {
			continue;
		}
		if (lookingElem.type === cElementType.bool && array[i].v.value !== lookingElem.value) {
			break;
		}

		if (array[i].v.value === lookingElem.value) {
			exactMatchIndex = i;
		} else if (array[i].v.value <= lookingElem.value) {
			resIndex = i;
		} else if (array[i].v.value > lookingElem.value) {
			break;
		}
	}
	return exactMatchIndex ? exactMatchIndex : resIndex;

};

_func[cElementType.number][cElementType.cell] = function ( arg0, arg1, what, bbox ) {
    var ar1 = arg1.tocNumber();
    switch ( what ) {
        case ">":
        {
            return new cBool( arg0.getValue() > ar1.getValue() );
        }
        case ">=":
        {
            return new cBool( arg0.getValue() >= ar1.getValue() );
        }
        case "<":
        {
            return new cBool( arg0.getValue() < ar1.getValue() );
        }
        case "<=":
        {
            return new cBool( arg0.getValue() <= ar1.getValue() );
        }
        case "=":
        {
            return new cBool( arg0.getValue() === ar1.getValue() );
        }
        case "<>":
        {
            return new cBool( arg0.getValue() !== ar1.getValue() );
        }
        case "-":
        {
            return new cNumber( arg0.getValue() - ar1.getValue() );
        }
        case "+":
        {
            return new cNumber( arg0.getValue() + ar1.getValue() );
        }
        case "/":
        {
            if ( arg1.getValue() !== 0 ) {
                return new cNumber( arg0.getValue() / ar1.getValue() );
      } else {
                return new cError( cErrorType.division_by_zero );
            }
        }
        case "*":
        {
            return new cNumber( arg0.getValue() * ar1.getValue() );
        }
        default:
        {
            return new cError( cErrorType.wrong_value_type );
        }
    }

};
_func[cElementType.cell][cElementType.number] = function ( arg0, arg1, what, bbox ) {
    var ar0 = arg0.tocNumber();
    switch ( what ) {
        case ">":
        {
            return new cBool( ar0.getValue() > arg1.getValue() );
        }
        case ">=":
        {
            return new cBool( ar0.getValue() >= arg1.getValue() );
        }
        case "<":
        {
            return new cBool( ar0.getValue() < arg1.getValue() );
        }
        case "<=":
        {
            return new cBool( ar0.getValue() <= arg1.getValue() );
        }
        case "=":
        {
            return new cBool( ar0.getValue() === arg1.getValue() );
        }
        case "<>":
        {
            return new cBool( ar0.getValue() !== arg1.getValue() );
        }
        case "-":
        {
            return new cNumber( ar0.getValue() - arg1.getValue() );
        }
        case "+":
        {
            return new cNumber( ar0.getValue() + arg1.getValue() );
        }
        case "/":
        {
            if ( arg1.getValue() !== 0 ) {
                return new cNumber( ar0.getValue() / arg1.getValue() );
      } else {
                return new cError( cErrorType.division_by_zero );
            }
        }
        case "*":
        {
            return new cNumber( ar0.getValue() * arg1.getValue() );
        }
        default:
        {
            return new cError( cErrorType.wrong_value_type );
        }
    }
};
_func[cElementType.cell][cElementType.cell] = function ( arg0, arg1, what, bbox ) {
    var ar0 = arg0.tocNumber();
    switch ( what ) {
        case ">":
        {
            return new cBool( ar0.getValue() > arg1.getValue() );
        }
        case ">=":
        {
            return new cBool( ar0.getValue() >= arg1.getValue() );
        }
        case "<":
        {
            return new cBool( ar0.getValue() < arg1.getValue() );
        }
        case "<=":
        {
            return new cBool( ar0.getValue() <= arg1.getValue() );
        }
        case "=":
        {
            return new cBool( ar0.getValue() === arg1.getValue() );
        }
        case "<>":
        {
            return new cBool( ar0.getValue() !== arg1.getValue() );
        }
        case "-":
        {
            return new cNumber( ar0.getValue() - arg1.getValue() );
        }
        case "+":
        {
            return new cNumber( ar0.getValue() + arg1.getValue() );
        }
        case "/":
        {
            if ( arg1.getValue() !== 0 ) {
                return new cNumber( ar0.getValue() / arg1.getValue() );
      } else {
                return new cError( cErrorType.division_by_zero );
            }
        }
        case "*":
        {
            return new cNumber( ar0.getValue() * arg1.getValue() );
        }
        default:
        {
            return new cError( cErrorType.wrong_value_type );
        }
    }
};

_func[cElementType.cellsRange3D] = _func[cElementType.cellsRange];
_func[cElementType.cell3D] = _func[cElementType.cell];

	function SharedProps(ref, base) {
		this.ref = ref;
		this.base = base;
	}

	SharedProps.prototype.isOneDimension = function() {
		return this.ref && (this.ref.r1 === this.ref.r2 || this.ref.c1 === this.ref.c2);
	};
	SharedProps.prototype.isHor = function() {
		return this.ref && this.ref.r1 === this.ref.r2;
	};

	function ParseResult(refPos, elems) {
		this.refPos = refPos;
		this.elems = elems;
		this.error = undefined;
		this.operand_expected = undefined;
		this.argPos = undefined;

		//for formula wizard
		this.argPosArr = [];
		this.activeFunction = null;
		this.cursorPos = undefined;

		//в процессе добавления формулы может найтись ссылка на внешний источник, который ещё не добавлен
		//сюда добавляем индексы и после парсинга формулы, добавляем новую структуру
		this.externalReferenesNeedAdd = null;
	}

	ParseResult.prototype.addRefPos = function(start, end, index, oper, isName) {
		if (this.refPos) {
			this.refPos.push({start: start, end: end, index: index, oper: oper, isName: isName});
		}
	};
	ParseResult.prototype.addElem = function(elem) {
		if (this.elems) {
			this.elems.push(elem);
		}
	};
	ParseResult.prototype.setError = function(error) {
		this.error = error;
	};
	ParseResult.prototype.getElementByPos = function(pos) {
		var curPos = 0;
		var argCount = [], level = 0;
		for (var i = 0; i < this.elems.length; ++i) {
			curPos += this.elems[i].toLocaleString(/*AscCommonExcel.cFormulaFunctionToLocale*/).length;

			//учитываем разделители аргументов
			if("(" === this.elems[i].name) {
				level++;
			} else if(")" === this.elems[i].name) {
				level--;
			} else if (level){
				if(!argCount[level]) {
					argCount[level] = 1;
				} else {
					argCount[level]++;
				}
				if(argCount[level] > 1) {
					curPos++;
				}
			}

			if (curPos >= pos) {
				return this.elems[i];
			}
		}
		return null;
	};
	ParseResult.prototype.getElementByPos2 = function(pos, start, end) {
		var curPos = 1;
		var curFunc = [], level = -1;
		for (var i = 0; i < this.elems.length; ++i) {
			var curElem = this.elems[i];
			var curElemStr = this.elems[i].toLocaleString();
			var curElemLength = curElemStr.length;
			var isFunc = curElem.type === cElementType.func;
			var nextElem = this.elems[i + 1];
			var needAddSeparator = true;

			if (isFunc && nextElem && nextElem.name === "(") {
				level++
				curFunc[level] = {func: this.elems[i], start: curPos};
				curElemLength++;
				needAddSeparator = false;
				i++;
			}

			if (pos !== undefined && pos >= curPos && pos <= curPos + curElemLength) {
				return curFunc[level] ? curFunc[level].func : null;
			} else if (start !== undefined && curFunc[level] && start >= curFunc[level].start && end <= curPos + curElemLength) {
				return curFunc[level].func;
			}

			if (curElem.name === ")") {
				level--;
			}
			curPos += curElemLength;
			/*if (needAddSeparator && cElementType.operator !== curElem.type) {
				curPos++;
			}*/
		}
		return null;
	};

	ParseResult.prototype.getArgumentsValue = function(sFormula) {
		let res = null;
		if (sFormula && this.argPosArr) {
			for (let i = 0; i < this.argPosArr.length; i++) {
				if (!res) {
					res = [];
				}

				if (i === this.argPosArr.length - 1 && this.error === c_oAscError.ID.FrmlParenthesesCorrectCount) {
					// We don't cut off the line at the last element, but only if the formula is parsed with an error (the formula is not closed or not entered completely)
					res.push(sFormula.substring(this.argPosArr[i].start - 1, this.argPosArr[i].end));
					continue
				}
				res.push(sFormula.substring(this.argPosArr[i].start - 1, this.argPosArr[i].end - 1));
			}
		}
		return res;
	};

	ParseResult.prototype.getActiveFunction = function(start, end) {
		var res = null;
		if (this.allFunctionsPos) {
			var startFuncs, endFuncs, i, j;
			for (i = 0; i < this.allFunctionsPos.length; i++) {
				if (this.allFunctionsPos[i].start + 1 <= start && this.allFunctionsPos[i].end + 1 >= start) {
					if (!startFuncs) {
						startFuncs = [];
					}
					startFuncs.push(this.allFunctionsPos[i]);
				}
				if (start !== end && this.allFunctionsPos[i].start + 1 <= end && this.allFunctionsPos[i].end + 1 >= end) {
					if (!endFuncs) {
						endFuncs = [];
					}
					endFuncs.push(this.allFunctionsPos[i]);
				}
			}

			if (startFuncs) {
				var commonFuncs;
				if (start === end) {
					commonFuncs = startFuncs;
				} else if (endFuncs) {
					//ищем самую внутреннюю функцию, где находится и начало и конец диапазона
					for (i = 0; i < startFuncs.length; i++) {
						for (j = 0; j < endFuncs.length; j++) {
							if (startFuncs[i] === endFuncs[j]) {
								if (!commonFuncs) {
									commonFuncs = [];
								}
								commonFuncs.push(startFuncs[i]);
								break;
							}
						}
					}
				}

				//ищем самую внутреннюю функцию
				if (commonFuncs) {
					res = commonFuncs[0];
					for (i = 1; i < commonFuncs.length; i++) {
						if (commonFuncs[i].start >= res.start && commonFuncs[i].end <= res.end) {
							res = commonFuncs[i];
						}
					}
				}
			}

		}
		return res;
	};

	ParseResult.prototype.checkNumberOperator = function(elemArr) {
		//проверка оператора перед числом
		//TODO ещё необходимо сделать проверку после числа + проверку с другими типами
		var res = true;
		let lastElem;
		if (this.elems && this.elems.length) {
			lastElem = this.elems[this.elems.length - 1];
			if (lastElem && lastElem.name === " ") {
				res = false;
			}
		} else if (elemArr) {
			lastElem = elemArr[elemArr.length - 1];
			if (lastElem && lastElem.name === " ") {
				res = false;
			}
		}
		return res;
	};

	function CalculateResult(checkOnError) {
		this.checkOnError = checkOnError;
		this.error = null;
	}
	CalculateResult.prototype.setError = function(error) {
		this.error = error;
	};

	var g_defParseResult = new ParseResult(undefined, undefined);

	var lastListenerId = 0;
/** класс отвечающий за парсинг строки с формулой, подсчета формулы, перестройки формулы при манипуляции с ячейкой*/
/** @constructor */
function parserFormula( formula, parent, _ws ) {
    this.is3D = false;
    this.ws = _ws;
    this.wb = this.ws.workbook;
    this.value = null;
    this.outStack = [];
    this.Formula = formula;
    this.isParsed = false;
    this.shared = null;

	this.listenerId = lastListenerId++;
	this.ca = false;
	this.isTable = false;
	this.isInDependencies = false;
	this.parent = parent;
	this._index = undefined;

	this.ref = null;

	this.promiseResult = null;
	this.replaceFormulaAfterCalc = null;

	//mark function, when need reparse and recalculate on custom function change
	this.unknownOrCustomFunction = null;

	if (AscFonts.IsCheckSymbols) {
		AscFonts.FontPickerByCharacter.getFontsByString(this.Formula);
	}
}
  parserFormula.prototype.getWs = function() {
    return this.ws;
  };
  parserFormula.prototype.getListenerId = function() {
    return this.listenerId;
  };
	parserFormula.prototype.setIsTable = function(isTable){
		this.isTable = isTable;
	};
	parserFormula.prototype.getShared = function() {
		return this.shared;
	};
	parserFormula.prototype.setShared = function(ref, cellWithFormula) {
		this.shared = new SharedProps(ref, cellWithFormula);
	};
	parserFormula.prototype.setSharedRef = function(newRef, opt_updateBase) {
		var old = this.shared.ref;
		if (!(newRef && newRef.r1 === old.r1 && newRef.c1 === old.c1 && newRef.r2 === old.r2 && newRef.c2 === old.c2)) {
			this.removeDependencies();
			if (newRef) {
				this.shared.ref = newRef;
				//todo is any issue if base is outside ref?
				if (opt_updateBase) {
					this.shared.base.nRow += newRef.r1 - old.r1;
					this.shared.base.nCol += newRef.c1 - old.c1;
				}
				this.buildDependencies();
			}
			var index = this.ws.workbook.workbookFormulas.add(this).getIndexNumber();
			History.Add(AscCommonExcel.g_oUndoRedoSharedFormula, AscCH.historyitem_SharedFormula_ChangeShared, null,
				null, new AscCommonExcel.UndoRedoData_IndexSimpleProp(index, opt_updateBase, old, newRef), true);
		}
	};
	parserFormula.prototype.removeShared = function() {
		this.shared = null;
	};
	/**
	 * @memberof parserFormula
	 * @returns {string}
	 */
	parserFormula.prototype.getFunctionName = function () {
		const aOutStack = this.outStack;

		for (let i = aOutStack.length - 1; i >= 0; i--) {
			if (aOutStack[i].type === cElementType.func) {
				return aOutStack[i].name;
			}
		}

		return "";
	};
	/**
	 * Checks a formula is conditional.
	 * @memberof parserFormula
	 * @param {string} sFunctionName
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._isConditionalFormula = function (sFunctionName) {
		const aExcludeCondFormulas = ["IFERROR", "IFNA", "COUNTIF", "BITLSHIFT", "BITRSHIFT", "DATEDIF"];
		const aCondFormulas = ["SWITCH"];

		return !!sFunctionName && (sFunctionName.includes("IF") || aCondFormulas.includes(sFunctionName)) &&
			!aExcludeCondFormulas.includes(sFunctionName);
	};
	parserFormula.prototype.notify = function(data) {
		var eventData = {notifyData: data, assemble: null, formula: this};
		let sFunctionName = this.getFunctionName();

		if (this._isConditionalFormula(sFunctionName) && data.areaData && g_cCalcRecursion.getIsCellEdited()) {
			let oCell = null;
			if (this.parent && null != this.parent.nRow && null != this.parent.nCol) {
				this.ws._getCell(this.parent.nRow, this.parent.nCol, function (oElem) {
					oCell = oElem;
				});
			}
			if (oCell && oCell.containInFormula()) {
				this.ca = this.isRecursiveCondFormula(sFunctionName);
			}
		}
		if (AscCommon.c_oNotifyType.Dirty === data.type) {
				if (this.parent && this.parent.onFormulaEvent) {
					this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.Change, eventData);
				}
		} else if (this.shared && this.parent && this.parent.onFormulaEvent &&
			this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.Shared, eventData)) {
			;
		} else if (AscCommon.c_oNotifyType.Prepare === data.type) {
			this.removeDependencies();
			this.processNotifyPrepare(data);
		} else if (AscCommon.c_oNotifyType.ChangeExternalLink === data.type) {
			this._changeExternalLink(data);
			this.Formula = this.assemble(true);
			this.buildDependencies();
		} else {
			this.removeDependencies();
			var needAssemble = true;
			if (this.parent && this.parent.onFormulaEvent && !this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.ProcessNotify, eventData)) {
				needAssemble = this.processNotify(data);
			}
			if (needAssemble) {
				eventData.assemble = this.assemble(true);
			} else {
				eventData.assemble = this.getFormula();
			}
			if (this.parent && this.parent.onFormulaEvent) {
				this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.ChangeFormula, eventData);
			}
			this.Formula = eventData.assemble;
			this.buildDependencies();
		}
	};
	parserFormula.prototype._changeExternalLink = function(data) {
		let existedWs = data.existedWs;
		for (let i = 0; i < this.outStack.length; i++) {
			let elem = this.outStack[i];
			if (elem.type === cElementType.cell3D) {
				this.outStack[i] = new AscCommonExcel.cRef3D(elem.value, existedWs ? existedWs : elem.ws, data.data.to);
			} else if (elem.type === cElementType.cellsRange3D) {
				this.outStack[i] = new AscCommonExcel.cArea3D(elem.value, existedWs ? existedWs : elem.wsFrom, existedWs ? existedWs : elem.wsTo, data.data.to);
			} else if (elem.type === cElementType.name3D) {
				this.outStack[i] = new AscCommonExcel.cName3D(elem.value, existedWs ? existedWs : elem.ws, data.data.to);
			}
		}
	};
	parserFormula.prototype._changeExternalLinkOld = function(data) {
		for (var i = 0; i < this.outStack.length; i++) {
			if (this.outStack[i].type === cElementType.cell3D || this.outStack[i].type === cElementType.cellsRange3D || this.outStack[i].type === cElementType.name3D) {
				if (this.outStack[i].externalLink == data.data.from) {
					this.outStack[i].externalLink = data.data.to;
				}
			}
		}
	};
	parserFormula.prototype.processNotifyPrepare = function(data) {
		var needAssemble = false;
		if (AscCommon.c_oNotifyType.ChangeSheet === data.actionType) {
			var changeData = data.data;
			if (this.is3D || changeData.remove){
				if (changeData.replace || changeData.remove) {
					if (changeData.remove) {
						needAssemble = this.removeSheet(changeData.remove, changeData.tableNamesMap);
		} else {
						needAssemble = this.moveSheet(changeData.replace);
					}
					data.preparedData[this.getListenerId()] = needAssemble;
				}
			}
		}
		return needAssemble;
	};
	parserFormula.prototype.processNotify = function(data) {
			var needAssemble = true;
			if (AscCommon.c_oNotifyType.Shift === data.type || AscCommon.c_oNotifyType.Move === data.type ||
				AscCommon.c_oNotifyType.Delete === data.type) {
				this.shiftCells(data.type, data.sheetId, data.bbox, data.offset, data.sheetIdTo, data.opt_isPivot, data.isTableCreated);
			} else if (AscCommon.c_oNotifyType.ChangeDefName === data.type) {
				if (!data.to) {
					this.removeTableName(data.from, data.bConvertTableFormulaToRef);
				} else if (data.from.name !== data.to.name) {
					this.changeDefName(data.from, data.to);
				} else if (data.from.type === Asc.c_oAscDefNameType.table) {
					needAssemble = false;
					this.changeTableRef(data.from.name);
				}
			} else if (AscCommon.c_oNotifyType.DelColumnTable === data.type) {
				this.removeTableColumn(data.tableName, data.deleted);
			} else if (AscCommon.c_oNotifyType.RenameTableColumn === data.type) {
				this.renameTableColumn(data.tableName);
			} else if (AscCommon.c_oNotifyType.ChangeSheet === data.type) {
				needAssemble = false;
				var changeData = data.data;
				if (this.is3D || changeData.remove) {
					if (changeData.replace || changeData.remove) {
					needAssemble = data.preparedData[this.getListenerId()];
					} else if (changeData.rename) {
						needAssemble = true;
					}
				}
			}
		return needAssemble;
	};
	parserFormula.prototype.clone = function (formula, parent, ws) {
		var opt_ws = null;
		if (Asc["editor"] && Asc["editor"].wb && Asc["editor"].wb.model && Asc["editor"].wb.model.addingWorksheet) {
			opt_ws = Asc["editor"].wb.model.addingWorksheet;
			ws = opt_ws;
		}
		if (null == formula) {
			formula = this.Formula;
		}
		if (null == parent) {
			parent = this.parent;
		}
		if (null == ws) {
			ws = this.ws;
		}
		var oRes = new parserFormula(formula, parent, ws);
		oRes.is3D = this.is3D;
		oRes.value = this.value;
		for (var i = 0, length = this.outStack.length; i < length; i++) {
			var oCurElem = this.outStack[i];
			if (oCurElem.clone) {
				oRes.outStack.push(oCurElem.clone(opt_ws));
			} else {
				oRes.outStack.push(oCurElem);
			}
		}
		oRes.isParsed = this.isParsed;
		oRes.ref = this.ref;
		oRes.ca = this.ca;
		return oRes;
	};
	parserFormula.prototype.getParent = function() {
		return this.parent;
	};
	parserFormula.prototype.getFormula = function() {
		if (AscCommonExcel.g_ProcessShared) {
			return this.assemble(true);
		} else {
			return this.Formula;
		}
	};
	parserFormula.prototype.getFormulaRaw = function() {
		return this.Formula;
	};
	parserFormula.prototype.setFormulaString = function(formula) {
		this.Formula = formula;
	};
	parserFormula.prototype.setFormula = function (formula) {
		this.Formula = formula;
		this.is3D = false;
		this.value = null;
		this.outStack = [];
		this.isParsed = false;
		this.ca = false;
		//this.isTable = false;
		this.isInDependencies = false;
	};

	/**
	 * Returns index the first element that satisfies the provided testing function from outStack array.
	 * If no elements satisfy the testing function, -1 is returned.
	 * Searching in reverse order.
	 * @param {[]} aOutStack
	 * @param {Function} fAction - Callback function must return truthy method as indicate that operand was found or falsy otherwise
	 * @returns {number}
	 * @private
	 */
	function _findLastOperandId(aOutStack, fAction) {
		const nStartIndex = aOutStack.length - 1;
		const nEndIndex = 0;

		for (let nIndex = nStartIndex; nIndex >= nEndIndex; nIndex--) {
			if (fAction(aOutStack[nIndex])) {
				return nIndex;
			}
		}

		return -1;
	}

	/**
	 * Gets a new stack with the concatenated count of argument and function.
	 * @param {[]}aOutStack
	 * @returns {[]}
	 * @private
	 */
	function _getNewOutStack(aOutStack) {
		const aNewOutStack = [];
		const nMainFuncIndex = _findLastOperandId(aOutStack, function (oElement) {
			return oElement.type && (oElement.type === cElementType.func || oElement.type === cElementType.operator);
		});
		if (!~nMainFuncIndex) {
			return aNewOutStack;
		}
		for (let i = 0; i < aOutStack.length; i++) {
			if (!(aOutStack[i] instanceof cBaseOperator) && aOutStack[i].type === cElementType.operator) {
				continue;
			}
			if (aOutStack[i].type === cElementType.specialFunctionStart || aOutStack[i].type === cElementType.specialFunctionEnd) {
				continue;
			}
			if (typeof aOutStack[i] === 'number') {
				let nNextIndex = i + 1;
				let oNextElement = nNextIndex < aOutStack.length ? aOutStack[nNextIndex] : null;
				if (oNextElement && oNextElement.type === cElementType.func) {
					const aArgsOfFunc = [];
					const aFuncData = [oNextElement, aOutStack[i]];
					if (nMainFuncIndex !== nNextIndex && aOutStack[i] > 0) {
						let nLastIndexArg = aOutStack[i] - 1;
						for (let j = nLastIndexArg; j >= 0; j--) {
							aArgsOfFunc[j] = aNewOutStack.pop();
						}
					}
					aFuncData.push(aArgsOfFunc);
					aNewOutStack.push(aFuncData);
					i++;
					continue;
				}
			} else if (aOutStack[i].type === cElementType.operator) {
				const aArgsOfOperator = [];
				const aOperatorData = [aOutStack[i], aOutStack[i].argumentsCurrent];
				let nPrevIndex = i - 1;
				let nEndIndexArg = i - aOutStack[i].argumentsCurrent;
				for (let j = nPrevIndex; j >= nEndIndexArg; j--) {
					if (j < 0) { // over reaching minimum edge
						break;
					}
					aArgsOfOperator.unshift(aNewOutStack.pop());
				}
				if (aArgsOfOperator.length < aOperatorData[1]) {
					break;
				}
				aOperatorData.push(aArgsOfOperator);
				aNewOutStack.push(aOperatorData);
				continue;
			}
			aNewOutStack.push(aOutStack[i]);
		}

		return aNewOutStack;
	}

	/**
	 * Gets range from condition functions, who need to calculate.
	 * @param {[]} aOutStack
	 * @param {number} nCountArgs
	 * @param {string} sFunctionName
	 * @returns {Object}
	 * @private
	 */
	function _getCalcRange(aOutStack, nCountArgs, sFunctionName) {
		if (sFunctionName.includes('IFS') && sFunctionName !== 'COUNTIFS') {
			return aOutStack.shift();
		}

		return  aOutStack[nCountArgs - 1];
	}

	/**
	 * Gets args of condition formula working with range.
	 * @param {string} sFunctionName
	 * @param {[]} aOutStack
	 * @param {number} nCountArgs
	 * @returns {[]}
	 * @private
	 */
	function _getArgsRangeCondFormula(sFunctionName, aOutStack, nCountArgs) {
		const aConditions = [];
		const aCriteriaRanges = [];
		let oCalcRange = null;
		let aArgs = null;

		if (sFunctionName !== "COUNTIFS") { // COUNTIFS doesn't need to oCalcRange.
			oCalcRange = _getCalcRange(aOutStack, nCountArgs, sFunctionName);
		}
		for (let i = 0; i < nCountArgs; i++) {
			let bEvenIndex = i % 2 === 0;

			if (!aOutStack[i]) {
				continue;
			}
			if (oCalcRange && oCalcRange.value === aOutStack[i].value) {
				continue;
			}
			if (bEvenIndex && !Array.isArray(aOutStack[i])) {
				aCriteriaRanges.push(aOutStack[i]);
				continue;
			}
			if (Array.isArray(aOutStack[i])) {
				if (aOutStack[i][0].name === sFunctionName) {
					continue;
				}
				aConditions.push(aOutStack[i][0]);
				aArgs = aOutStack[i][2];
				continue;
			}
			aConditions.push(aOutStack[i]);
		}

		return [oCalcRange, aCriteriaRanges, aConditions, aArgs];
	}

	/**
	 * Gets args of condition formula
	 * @param {[]} aOutStack
	 * @param {string} sFunctionName
	 * @returns {[]}
	 * @private
	 */
	function _getArgsCondFormula(aOutStack, sFunctionName) {
		const aLogicalTests = [];
		const aTrueResults = [];
		let oFalseResult = null;
		// Uses for SWITCH formula.
		let oExpressionValue = null;
		const oEqualOperator = cFormulaOperators["="].prototype;
		let oDefaultResult = null;

		if (sFunctionName === "IF") {
			aLogicalTests.push(aOutStack.shift());
			aTrueResults.push(aOutStack.shift());
			oFalseResult = aOutStack.shift();

			return [aLogicalTests, aTrueResults, oFalseResult];
		}
		if (sFunctionName === "SWITCH") {
			oExpressionValue = aOutStack.shift();
		}

		let nMainFunctionIndex = _findLastOperandId(aOutStack, function (oElement) {
			if (Array.isArray(oElement)) {
				return oElement[0].type === cElementType.func || oElement[0].type === cElementType.operator;
			}
			return false;
		});
		aOutStack = aOutStack.slice(0, nMainFunctionIndex);
		let bEvenLength = aOutStack.length % 2 === 0;
		if (!bEvenLength) {
			oDefaultResult = aOutStack.pop();
		}
		for (let i = 0, length = aOutStack.length; i < length; i++) {
			let operand = aOutStack[i];
			let bEvenIndex = i % 2 === 0;
			if (bEvenIndex) {
				// For SWITCH formula converts data to (logical_test, true_res, ...) format like in IFS formula
				if (sFunctionName === "SWITCH") {
					const aEqualOpInfo = [oEqualOperator, oEqualOperator.argumentsCurrent];
					const aArgs = [oExpressionValue];
					aArgs.push(operand);
					aEqualOpInfo.push(aArgs);
					operand = aEqualOpInfo;
				}
				aLogicalTests.push(operand);
				continue;
			}
			aTrueResults.push(operand);
		}

		return [aLogicalTests, aTrueResults, oFalseResult, oDefaultResult];
	}

	/**
	 * Checks cell with formula is in area.
	 * @memberof parserFormula
	 * @param found_operand operand of formula
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._isAreaContainCell = function (found_operand) {
		const oParentCell = this.getParent();
		let nOperandType = found_operand.type;
		let oRange = null;

		if (!oParentCell) {
			return false;
		}
		if (!(oParentCell instanceof AscCommonExcel.CCellWithFormula)) {
			return false;
		}
		if (oParentCell.nRow == null && oParentCell.nCol == null) {
			return false;
		}
		if (nOperandType === cElementType.name || nOperandType === cElementType.name3D) {
			found_operand = found_operand.getValue();
			nOperandType = found_operand.type;
		}
		if (nOperandType === cElementType.cellsRange) {
			oRange = found_operand.getRange();
			return oRange.containCell2(oParentCell);
		}
		if (nOperandType === cElementType.cellsRange3D) {
			const aRanges = found_operand.getRanges().filter(function (oRange) {
				return oParentCell.ws.getId() === oRange.worksheet.getId();
			});

			for (let i = 0, length = aRanges.length; i < length; i++) {
				if (aRanges[i].containCell2(oParentCell)) {
					return true;
				}
			}
			return false;
		}

		return false;
	};
	/**
	 * Checks if the criteria cell has same formula as parserFormula.
	 * @memberof parserFormula
	 * @param {object} oCriteriaRange
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._criteriaCellHasFormula = function (oCriteriaRange ) {
		if (oCriteriaRange.type === cElementType.name || oCriteriaRange.type === cElementType.name3D) {
			oCriteriaRange = oCriteriaRange.toRef();
		}
		const oThis = this;
		const oRange = oCriteriaRange.getRange();
		const oBbox = oRange.bbox;
		const oParentCell = this.getParent();
		if (!oParentCell || (oParentCell && oParentCell.nRow == null && oParentCell.nCol == null)) {
			return false;
		}
		let bHasFormula = false;
		let bVertical = oBbox.c1 === oBbox.c2;
		let nRow = bVertical ? oBbox.r1 + (oParentCell.nRow - oBbox.r1) : oBbox.r1;
		let nCol = bVertical ? oBbox.c1 :  oBbox.c1 + (oParentCell.nCol - oBbox.c1);

		oRange.worksheet._getCellNoEmpty(nRow, nCol, function (oCell) {
			if (oCell && oCell.isFormula()) {
				let oParsedFormula = oCell.getFormulaParsed();
				if (oParsedFormula.Formula.replace(/\s+/g, '') === oThis.Formula.replace(/\s+/g, '')) {
					bHasFormula = true;
				}
			}
		});

		return bHasFormula;
	};
	/**
	 * Checks the cell with formula matches the criteria.
	 * @memberof parserFormula
	 * @param {{condition: object, calcRange: object, criteriaRange: object, argsFuncCondition: []}} oFormulaArgs
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._calculateMatch = function (oFormulaArgs) {
		const oParentCell = this.getParent();
		if (!oParentCell || (oParentCell && oParentCell.nRow == null && oParentCell.nCol == null)) {
			return false;
		}
		const oCalcRange = oFormulaArgs.calcRange;
		let oCriteriaRange = oFormulaArgs.criteriaRange;
		let oCondition = oFormulaArgs.condition;
		let bVertical;

		if (oCondition.type === cElementType.name || oCondition.type === cElementType.name3D) {
			oCondition = oCondition.getValue();
		}
		if (oCondition.type === cElementType.func) {
			const aArgsFuncCondition = oFormulaArgs.argsFuncCondition;
			let oBbox = oParentCell.onFormulaEvent && oParentCell.onFormulaEvent(AscCommon.c_oNotifyParentType.GetRangeCell);
			oCondition = oCondition.Calculate(aArgsFuncCondition, oBbox, undefined, this.ws);
		}
		if (oCriteriaRange.type === cElementType.name || oCriteriaRange.type === cElementType.name3D) {
			oCriteriaRange = oCriteriaRange.toRef();
		}
		let oBBoxCriteria = oCriteriaRange.getBBox0();
		if (oCalcRange) {
			let nCalcRangeCol = oCalcRange.c2 - oCalcRange.c1 + 1;
			let nCalcRangeRow = oCalcRange.r2 - oCalcRange.r1 + 1;
			let nCriteriaRangeCol = oBBoxCriteria.c2 - oBBoxCriteria.c1 + 1;
			let nCriteriaRangeRow = oBBoxCriteria.r2 - oBBoxCriteria.r1 + 1;
			if (nCalcRangeCol !== nCriteriaRangeCol || nCalcRangeRow !== nCriteriaRangeRow) {
				return false;
			}
			bVertical = oBBoxCriteria.r1 === oCalcRange.r1 && oBBoxCriteria.r2 === oCalcRange.r2;
		} else {
			bVertical = oBBoxCriteria.c1 === oBBoxCriteria.c2;
		}

		let oCriteriaRangeVal = bVertical ? oCriteriaRange.getValueByRowCol(oParentCell.nRow - oBBoxCriteria.r1, 0) : oCriteriaRange.getValueByRowCol(0, oParentCell.nCol - oBBoxCriteria.c1);
		let oMatchInfo = AscCommonExcel.matchingValue(oCondition.tocString());

		return !!oCriteriaRangeVal && AscCommonExcel.matching(oCriteriaRangeVal, oMatchInfo);
	};
	/**
	 * Checks criteria range by condition.
	 * @memberof parserFormula
	 * @param {[]} aRangeArgs
	 * @param {number} nCountArgs
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._checkRangeByCriteria = function (aRangeArgs, nCountArgs) {
		let oCalcRange = aRangeArgs[0];
		const aCriteriaRanges = aRangeArgs[1];
		const aConditions = aRangeArgs[2];
		const aArgs = aRangeArgs[3];

		if (oCalcRange && (oCalcRange.type === cElementType.name || oCalcRange.type === cElementType.name3D)) {
			oCalcRange = oCalcRange.toRef();
		}
		let nLen = Math.floor(nCountArgs / 2);
		let oRangeSum = oCalcRange && oCalcRange.getBBox0();
		let bMatch = false;

		for (let i = 0; i < nLen; i++) {
			if (this._criteriaCellHasFormula(aCriteriaRanges[i])) {
				return bMatch;
			}
			let oFormulaArgs = {
				criteriaRange: aCriteriaRanges[i],
				condition: aConditions[i],
				calcRange: oRangeSum,
			};
			if (aArgs != null) {
				oFormulaArgs.argsFuncCondition = aArgs;
			}
			bMatch = this._calculateMatch(oFormulaArgs);
			if (!bMatch) {
				return false;
			}
		}
		return true;
	};
	/**
	 * Calculates logical test of the conditional formula like IF.
	 * Recursive function
	 * @memberof parserFormula
	 * @param {[]} aLogicalTest
	 * @returns {cBool|null}
	 * @private
	 */
	parserFormula.prototype._calculateLogicalTest = function (aLogicalTest) {
		if (g_cCalcRecursion.checkRecursionCounter()) {
			g_cCalcRecursion.resetRecursionCounter();
			return null;
		}
		const aTypesWithRange = [cElementType.cell, cElementType.cell3D, cElementType.cellsRange, cElementType.cellsRange3D];
		const aNameType = [cElementType.name, cElementType.name3D];
		const oFormula = aLogicalTest[0];
		const aArgs =  aLogicalTest[2];
		const oParentCell = this.getParent();
		const oBbox = oParentCell.onFormulaEvent && oParentCell.onFormulaEvent(AscCommon.c_oNotifyParentType.GetRangeCell);
		if (!oBbox) {
			return new cError(cErrorType.not_numeric);
		}

		for (let i = 0, len = aArgs.length; i < len; i++) {
			if (aArgs[i] && aNameType.includes(aArgs[i].type)) {
				aArgs[i] = aArgs[i].toRef();
			}
			if (aArgs[i] && aArgs[i].type === cElementType.table) {
				aArgs[i] = aArgs[i].toRef(oBbox);
			}
			if (aArgs[i] && Array.isArray(aArgs[i])) {
				g_cCalcRecursion.incRecursionCounter();
				aArgs[i] = this._calculateLogicalTest(aArgs[i]);
				g_cCalcRecursion.resetRecursionCounter();
				if (aArgs[i] == null) {
					return null;
				}
			}
			// Check on recursion ref.  Recursion ref means that cell is recursion need to set ca flag to true.
			if (aArgs[i] && aTypesWithRange.includes(aArgs[i].type) && !this._isConditionalFormula(oFormula.name)) {
				if (aArgs[i].getBBox0().contains(oParentCell.nCol, oParentCell.nRow)) {
					return null;
				}
			}
		}

		return oFormula.Calculate(aArgs, oBbox, undefined, this.ws);
	};
	/**
	 * Finds recursion cell in equation with refs.
	 * Recursive function
	 * @memberof parserFormula
	 * @param {[]} aRef
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._findRecursionRef = function (aRef) {
		if (g_cCalcRecursion.checkRecursionCounter()) {
			g_cCalcRecursion.resetRecursionCounter();
			return false;
		}

		const oThis = this;
		const aArg = aRef[2];
		let bRecursiveCell = false;

		for (let i = 0, len = aArg.length; i < len; i++) {
			if (aArg[i] && (aArg[i].type === cElementType.name || aArg[i].type === cElementType.name3D)) {
				aArg[i] = aArg[i].toRef();
			}
			if (aArg[i] && (aArg[i].type === cElementType.cell || aArg[i].type === cElementType.cell3D)) {
				let oRange = aArg[i].getRange();
				oRange._foreachNoEmpty(function (oCell) {
					if (!bRecursiveCell) {
						bRecursiveCell = oCell.checkRecursiveFormula(oThis.getParent());
					}
				});

			}
			if (aArg[i] && (aArg[i].type === cElementType.cellsRange || aArg[i].type === cElementType.cellsRange3D)) {
				bRecursiveCell = this._isAreaContainCell(aArg[i]);

			}
			if (aArg[i] && Array.isArray(aArg[i])) {
				g_cCalcRecursion.incRecursionCounter();
				bRecursiveCell = this._findRecursionRef(aArg[i]);
				g_cCalcRecursion.resetRecursionCounter();
			}
			if (bRecursiveCell) {
				break;
			}
		}

		return bRecursiveCell;
	};
	/**
	 * Checks operand has a recursion.
	 * @memberof parserFormula
	 * @param {cRef|cRef3D|cName|cName3D|cString|cNumber|cBool|cArea|cArea3D} oOperand
	 * @returns {boolean}
	 */
	parserFormula.prototype._isOperandRecursive = function (oOperand) {
		const oThis = this;
		const aTypesWithRange = [cElementType.cell, cElementType.cell3D, cElementType.cellsRange, cElementType.cellsRange3D];
		const aNameType = [cElementType.name, cElementType.name3D];
		let bRecursiveCell = false;

		if (oOperand && Array.isArray(oOperand)) {
			return this._findRecursionRef(oOperand);
		}
		if (oOperand && aNameType.includes(oOperand.type)) {
			oOperand = oOperand.toRef();
		}
		if (oOperand && aTypesWithRange.includes(oOperand.type)) {
			if (oOperand.type === cElementType.cell || oOperand.type === cElementType.cell3D) {
				let oRange = oOperand.getRange();
				oRange._foreachNoEmpty(function (oCell) {
					bRecursiveCell = oCell.checkRecursiveFormula(oThis.getParent());
				});
				return bRecursiveCell;
			}
			return this._isAreaContainCell(oOperand);
		}

		return false;
	};
	/**
	 * Calculates conditional formulas like IF, IFS, and SWITCH and checks if it has recursion.
	 * @memberof parserFormula
	 * @param {[]} aOutStack
	 * @param {string} sFunctionName
	 * @param {number} nCountArgs
	 * @returns {boolean}
	 * @private
	 */
	parserFormula.prototype._evalAndCheckRecursion = function (aOutStack, sFunctionName, nCountArgs) {
		const aLogicalTestTypes = [cElementType.bool, cElementType.cell, cElementType.cell3D, cElementType.name, cElementType.name3D];
		const aArgs = _getArgsCondFormula(aOutStack, sFunctionName);
		let aLogicalTest = aArgs[0];
		let aTrueValue = aArgs[1];
		let falseValue = aArgs[2];
		let defaultValue = aArgs[3]; // For SWITCH formula
		let bRecursiveCell = false;
		let bOperandFound = false;

		// For SWITCH exclude expression argument from nCountArgs.
		let nLen = sFunctionName === "SWITCH" ? Math.floor((nCountArgs - 1) / 2) : Math.floor(nCountArgs / 2);
		for (let i = 0; i < nLen; i++) {
			let logicalTest = aLogicalTest[i];
			if (Array.isArray(logicalTest)) {
				logicalTest = this._calculateLogicalTest(logicalTest);
				if (logicalTest == null) {
					return true;
				}
			}
			if (!aLogicalTestTypes.includes(logicalTest.type)) {
				return false;
			}
			if (logicalTest.type === cElementType.name || logicalTest.type === cElementType.name3D) {
				logicalTest = logicalTest.toRef();
			}
			if (logicalTest.type === cElementType.cell || logicalTest.type === cElementType.cell3D) {
				logicalTest = logicalTest.getValue();
			}
			let value = logicalTest.value ? aTrueValue[i] : falseValue;
			bOperandFound = !!value;
			bRecursiveCell = this._isOperandRecursive(value);
			if (bRecursiveCell) {
				return bRecursiveCell;
			}
		}
		if (!bOperandFound && defaultValue) {
			return this._isOperandRecursive(defaultValue);
		}

		return bRecursiveCell;
	};
	/**
	 * Checks a condition function is recursive or not.
	 * @param {string} sFunctionName
	 * @param {[]} [aArgs]
	 * @returns {boolean}
	 */
	parserFormula.prototype.isRecursiveCondFormula = function (sFunctionName, aArgs) {
		const aCellFormulas = ['IF', 'IFS', 'SWITCH'];
		const aOutStack = aArgs && aArgs.length ? aArgs : _getNewOutStack(this.outStack);
		if (!aOutStack.length) {
			return false;
		}
		if (aOutStack.length === 1 && aOutStack[0][0].type === cElementType.operator) {
			const aArgs = aOutStack[0][2];
			let bHasRecursion = false;
			for (let i = 0, length = aArgs.length; i < length; i++) {
				if (aArgs[i] && Array.isArray(aArgs[i])) {
					let oFormula = aArgs[i][0];
					if (oFormula.type === cElementType.operator) {
						bHasRecursion = this.isRecursiveCondFormula(sFunctionName, [aArgs[i]]);
					}
					if (oFormula.type === cElementType.func) {
						if (!this._isConditionalFormula(oFormula.name)) {
							continue;
						}
						if (sFunctionName !== oFormula.name && this._isConditionalFormula(oFormula.name)) {
							sFunctionName = oFormula.name;
						}
						bHasRecursion = this.isRecursiveCondFormula(sFunctionName, aArgs[i][2]);
					}
					if (bHasRecursion) {
						return true;
					}
				}
			}
			return false;
		}
		const nCountArgs = aArgs && aArgs.length ? aOutStack.length : Number(aOutStack[aOutStack.length - 1][1]);
		const aNameType = [cElementType.name, cElementType.name3D];
		let bRecursiveCell = false;
		let bRange = !aCellFormulas.includes(sFunctionName);

		if (bRange) {
			const aAreaType = [cElementType.cellsRange, cElementType.cellsRange3D];
			// For formulas like SUMIF, COUNTIF, etc. with 2 arguments, check the range has the cycle link without criteria.
			if (nCountArgs === 2) {
				bRecursiveCell = this._isAreaContainCell(aOutStack[0]);
				if (!bRecursiveCell && (typeof aOutStack[nCountArgs - 1] !== 'number' && (aAreaType.includes(aOutStack[nCountArgs - 1].type) ||
					aNameType.includes(aOutStack[nCountArgs - 1].type)))) {
					bRecursiveCell = this._isAreaContainCell(aOutStack[nCountArgs - 1]);
				}
				return bRecursiveCell;
			}

			const aRangeArgs = _getArgsRangeCondFormula(sFunctionName, aOutStack, nCountArgs);
			let oCalcRange = aRangeArgs[0];
			const aCriteriaRanges = aRangeArgs[1];
			const aConditions = aRangeArgs[2];
			let bHasRecursiveCriteria = false;

			if (aConditions.length) {
				for (let i = 0, length = aConditions.length; i < length; i++) {
					if (this._isAreaContainCell(aConditions[i])) {
						return true;
					}
				}
			}
			if (aCriteriaRanges.length && this._isAreaContainCell(aCriteriaRanges[0])) {
				return true;
			}
			if (aCriteriaRanges.length) {
				for (let i = 1, length = aCriteriaRanges.length; i < length;  i++) {
					if (this._isAreaContainCell(aCriteriaRanges[i])) {
						bHasRecursiveCriteria = true;
						break;
					}
				}
			}
			let bRecursiveCalcRange = !!oCalcRange && this._isAreaContainCell(oCalcRange);
			// Checking criteria for the range.
			if ((bRecursiveCalcRange || bHasRecursiveCriteria) && aCriteriaRanges.length && aConditions.length) {
				return this._checkRangeByCriteria(aRangeArgs, nCountArgs);
			}
		} else {
			const MIN_COUNT_ARGS = 2;
			if (isNaN(nCountArgs) && nCountArgs < MIN_COUNT_ARGS) {
				return false;
			}
			return this._evalAndCheckRecursion(aOutStack, sFunctionName, nCountArgs);
		}
		return false;
	};
	parserFormula.prototype.parse = function (local, digitDelim, parseResult, ignoreErrors, renameSheetMap, tablesMap, opt_pivotNamesList) {
		var elemArr = [];
		var ph = {operand_str: null, pCurrPos: 0};
		var needAssemble = false;
		var cFormulaList;

		var startArrayFunc = false, counterArrayFunc = 0, isFoundImportFunctions;

		if (this.isParsed) {
			return this.isParsed;
		}

		if(!parseResult){
			parseResult = g_defParseResult;
		}
		/*
		 Парсер формулы реализует алгоритм перевода инфиксной формы записи выражения в постфиксную или Обратную Польскую Нотацию.
		 Что упрощает вычисление результата формулы.
		 При разборе формулы важен порядок проверки очередной части выражения на принадлежность тому или иному типу.
		 */

		if (this.Formula.length >= AscCommon.c_oAscMaxFormulaLength) {
			parseResult.setError(c_oAscError.ID.FrmlMaxLength);
			this.outStack = [];
			return false;
		}

		if (false) {

			var getPrevElem = function(aTokens, pos){
				for(var n = pos - 1; n >=0; n--){
					if("" !== aTokens[n].value){
						return aTokens[n];
					}
				}
				return aTokens[pos - 1];
			};

			//console.log(this.Formula);
			cFormulaList =
				(local && AscCommonExcel.cFormulaFunctionLocalized) ? AscCommonExcel.cFormulaFunctionLocalized :
					cFormulaFunction;
			var aTokens = getTokens(this.Formula);
			if (null === aTokens) {
				this.outStack = [];
				parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
				return false;
			}

			var notEndedFuncCount = 0;
			var stack = [], val, valUp, tmp, elem, len, indentCount = -1, args = [], prev, next, arr = null,
				bArrElemSign = false, wsF, wsT, arg_count;
			for (var i = 0, nLength = aTokens.length; i < nLength; ++i) {
				if(TOK_SUBTYPE_START === aTokens[i].subtype) {
					notEndedFuncCount++;
				} else if(TOK_SUBTYPE_STOP === aTokens[i].subtype) {
					notEndedFuncCount--;
				}

				found_operand = null;
				val = aTokens[i].value;
				switch (aTokens[i].type) {
					case TOK_TYPE_OPERAND: {
						if (TOK_SUBTYPE_TEXT === aTokens[i].subtype) {
							elem = new cString(val);
						} else {
							tmp = parseFloat(val);
							if (isNaN(tmp)) {
								valUp = val.toUpperCase();
								if ('TRUE' === valUp || 'FALSE' === valUp) {
									elem = new cBool(valUp);
								} else {
									if (-1 !== val.indexOf('!')) {
										tmp = AscCommonExcel.g_oRangeCache.getRange3D(val);
										if (tmp) {
											this.is3D = true;
											wsF = this.wb.getWorksheetByName(tmp.sheet);
											wsT = (null !== tmp.sheet2 && tmp.sheet !== tmp.sheet2) ?
												this.wb.getWorksheetByName(tmp.sheet2) : wsF;
											var name = tmp.getName().split("!")[1];
											elem = (tmp.isOneCell()) ? new cRef3D(name, wsF) :
												new cArea3D(name, wsF, wsT);
											parseResult.addRefPos(aTokens[i].pos - aTokens[i].length,
												aTokens[i].pos, this.outStack.length, elem);
										} else if(TOK_SUBTYPE_ERROR === aTokens[i].subtype) {
											elem = new cError(val);
										} else {
											parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
											this.outStack = [];
											return false;
										}
									} else {
										tmp = AscCommonExcel.g_oRangeCache.getAscRange(valUp);
										if (tmp) {
											//если использовать isOneCell - тогда A1:A1 -> A1
											var isOneCell = /*tmp.isOneCell()*/!valUp.split(":")[1];
											elem = isOneCell ? new cRef(valUp, this.ws) : new cArea(valUp, this.ws);
											parseResult.addRefPos(aTokens[i].pos - aTokens[i].length, aTokens[i].pos, this.outStack.length, elem);
										} else if(TOK_SUBTYPE_ERROR === aTokens[i].subtype) {
											elem = new cError(val);
										} else {
											elem = new cName(aTokens[i].value, this.ws);
											parseResult.addRefPos(aTokens[i].pos - aTokens[i].length,
												aTokens[i].pos,	this.outStack.length, elem);
										}
									}
								}
							} else {
								elem = new cNumber(tmp);
							}
						}
						if (arr) {
							if (cElementType.number !== elem.type && cElementType.bool !== elem.type &&
								cElementType.string !== elem.type) {
								this.outStack = [];
								parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
								return false;
							} else {
								if (bArrElemSign) {
									if (cElementType.number !== elem.type) {
										this.outStack = [];
										parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
										return false;
									}
									elem.value *= -1;
									bArrElemSign = false;
								}
								arr.addElement(elem);
							}
						} else {
							this.outStack.push(elem);
							parseResult.addElem(elem);
						}
						break;
					}
					case TOK_TYPE_OP_POST:
					case TOK_TYPE_OP_IN: {
						if (TOK_SUBTYPE_UNION === aTokens[i].subtype) {
							this.outStack = [];
							parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
							return false;
						}

						prev = getPrevElem(aTokens, i);
						if ('-' === val && (0 === i ||
							(TOK_TYPE_OPERAND !== prev.type && TOK_TYPE_OP_POST !== prev.type &&
							(TOK_SUBTYPE_STOP !== prev.subtype ||
							(TOK_TYPE_FUNCTION !== prev.type && TOK_TYPE_SUBEXPR !== prev.type))))) {
							elem = cFormulaOperators['un_minus'].prototype;
						} else {
							elem = cFormulaOperators[val].prototype;
						}
						if (arr) {
							if (bArrElemSign || 'un_minus' !== elem.name) {
								this.outStack = [];
								parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
								return false;
							} else {
								bArrElemSign = true;
								break;
							}
						}

						parseResult.addElem(elem);

						len = stack.length;
						while (0 !== len) {
							tmp = stack[len - 1];
							if (elem.rightAssociative ? (elem.priority < tmp.priority) :
									((elem.priority <= tmp.priority))) {
								this.outStack.push(tmp);
								--len;
							} else {
								break;
							}
						}
						stack.length = len;

						stack.push(elem);
						break;
					}
					case TOK_TYPE_FUNCTION: {
						if (TOK_SUBTYPE_START === aTokens[i].subtype) {
							val = val.toUpperCase();
							if ('ARRAY' === val) {
								if (arr) {
									this.outStack = [];
									parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
									return false;
								}
								arr = new cArray();
								break;
							} else if ('ARRAYROW' === val) {
								if (!arr) {
									this.outStack = [];
									parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
									return false;
								}
								arr.addRow();
								break;
							} else if (val in cFormulaList) {
								elem = cFormulaList[val].prototype;
							} else if (val in cAllFormulaFunction) {
								elem = cAllFormulaFunction[val].prototype;
							} else {
								elem = new cUnknownFunction(val);
								let xlfnFrefix = "_xlfn.";
								let xlwsFrefix = "_xlws.";
								//let xludfFrefix = "__xludf.DUMMYFUNCTION.";
								//_xlws only together with _xlfn
								elem.isXLFN = (val.indexOf(xlfnFrefix) === 0);
								elem.isXLWS = elem.isXLFN && xlfnFrefix.length === val.indexOf(xlwsFrefix);
							}
							if(arrayFunctionsMap[val]){
								startArrayFunc = true;

								counterArrayFunc++;
								if(1 === counterArrayFunc){
									this.outStack.push(cSpecialOperandStart.prototype);
								}
							}
							if (elem && elem.ca) {
								this.ca = elem.ca;
							}
							stack.push(elem);
							args[++indentCount] = 1;
						} else {
							if (arr) {
								if ('ARRAY' === val) {
									if (!arr.isValidArray()) {
										this.outStack = [];
										// размер массива не согласован
										parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
										return false;
									}
									this.outStack.push(arr);
									arr = null;
								} else if ('ARRAYROW' !== val) {
									this.outStack = [];
									parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
									return false;
								}
								break;
							}
							len = stack.length;
							while (0 !== len) {
								tmp = stack[len - 1];
								--len;
								this.outStack.push(tmp);
								if (cElementType.func === tmp.type) {
									prev = aTokens[i - 1];
									arg_count = args[indentCount] -
										((prev && TOK_TYPE_FUNCTION === prev.type && TOK_SUBTYPE_START ===
											prev.subtype) ? 1 : 0);
									//this.outStack.push(arg_count);
									this.outStack.splice(this.outStack.length - 1, 0, arg_count);

									if(startArrayFunc && arrayFunctionsMap[tmp.name]){
										counterArrayFunc--;
										if(counterArrayFunc < 1){
											startArrayFunc = false;
											this.outStack.push(cSpecialOperandEnd.prototype);
										}
									}

									if (!tmp.checkArguments(arg_count)) {
										this.outStack = [];
										parseResult.setError(c_oAscError.ID.FrmlWrongMaxArgument);
										return false;
									}
									break;
								}
							}
							stack.length = len;
							--indentCount;
						}
						break;
					}
					case TOK_TYPE_ARGUMENT: {
						if (arr) {
							break;
						}
						if (-1 === indentCount) {
							throw 'error!!!!!!!!!!!';
						}
						args[indentCount] += 1;
						len = stack.length;
						while (0 !== len) {
							tmp = stack[len - 1];
							if (cElementType.func === tmp.type) {
								break;
							}
							this.outStack.push(tmp);
							--len;
						}
						stack.length = len;

						next = aTokens[i + 1];
						if (next && (TOK_TYPE_ARGUMENT === next.type ||
							(TOK_TYPE_FUNCTION === next.type && TOK_SUBTYPE_START !== next.subtype))) {
							this.outStack.push(new cEmpty());
							break;
						}
						break;
					}
					case TOK_TYPE_SUBEXPR: {
						if (TOK_SUBTYPE_START === aTokens[i].subtype) {
							elem = new parentLeft();
							stack.push(elem);
						} else {
							elem = new parentRight();
							len = stack.length;
							while (0 !== len) {
								tmp = stack[len - 1];
								--len;
								this.outStack.push(tmp);
								if (tmp instanceof parentLeft) {
									break;
								}
							}
							stack.length = len;
						}
						parseResult.addElem(elem);
						break;
					}
					case TOK_TYPE_WSPACE: {
						if (0 !== i && i !== nLength - 1) {
							prev = aTokens[i - 1];
							next = aTokens[i + 1];
							if ((TOK_TYPE_OPERAND === prev.type ||
								((TOK_TYPE_FUNCTION === prev.type || TOK_TYPE_SUBEXPR === prev.type) &&
								TOK_SUBTYPE_STOP === prev.subtype)) && ((TOK_TYPE_OPERAND === next.type) ||
								((TOK_TYPE_FUNCTION === next.type || TOK_TYPE_SUBEXPR === next.type) &&
								TOK_SUBTYPE_START === next.subtype))) {
								aTokens[i].type = TOK_TYPE_OP_IN;
								aTokens[i].value = ' ';
								--i;
							}
						}
						break;
					}
				}
			}
			while (stack.length !== 0) {
				this.outStack.push(stack.pop());
			}

			if(notEndedFuncCount) {
				this.outStack = [];
				parseResult.setError(c_oAscError.ID.FrmlOperandExpected);
				return false;
			}

			if (this.outStack.length !== 0) {
				return this.isParsed = true;
			} else {
				return this.isParsed = false;
			}
		}

		parseResult.operand_expected = true;
		var wasLeftParentheses = false, wasRigthParentheses = false, found_operand = null, _3DRefTmp = null, _tableTMP = null;
		cFormulaList = (local && AscCommonExcel.cFormulaFunctionLocalized) ? AscCommonExcel.cFormulaFunctionLocalized : cFormulaFunction;
		var leftParentArgumentsCurrentArr = [];
		var referenceCount = 0;

		//позиция курсора при открытой ячейке на редактирование
		//если activePos - undefined - ищем первую функцию
		var needCalcArgPos = ignoreErrors;
		var activePos = parseResult.cursorPos;
		var needAddCursorPos = activePos === undefined;
		var needFuncLevel = 0;
		var currentFuncLevel = -1;
		var levelFuncMap = [];
		var argFuncMap = [];
		var argPosArrMap = [];
		var startArrayArg = null;
		let bConditionalFormula = false;

		var t = this;
		var _checkReferenceCount = function (weight) {
			//ввожу ограничение на максимальное количество операндов в формуле
			//для этого добавляю вес каждого операнда
			//func - 0.75, array - 2, bool - 0.5, number - _number >= 65536 || Number.isInteger(_number)) ? 1.25 : 0.5
			//string - 0.5+length*0/25
			//error - 1
			//area - 2 или 3(в зависимости от количества листов)
			//ref - 1, table - 2, defName - 0.75
			//array - 2

			referenceCount += weight;
			if (referenceCount > AscCommon.c_oAscMaxFormulaReferenceLength) {
				parseResult.setError(c_oAscError.ID.FrmlMaxReference);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			}
			return true;
		};

		var parseOperators = function () {
			wasLeftParentheses = false;
			wasRigthParentheses = false;
			var found_operator = null;

			if (parseResult.operand_expected) {
				if ('-' === ph.operand_str) {
					parseResult.operand_expected = true;
					found_operator = cFormulaOperators['un_minus'].prototype;
				} else if ('+' === ph.operand_str) {
					parseResult.operand_expected = true;
					found_operator = cFormulaOperators['un_plus'].prototype;
				} else if (' ' === ph.operand_str) {
					return true;
				} else {
					parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
					t.outStack = [];
					return false;
				}
			} else if (!parseResult.operand_expected) {
				if ('-' === ph.operand_str) {
					parseResult.operand_expected = true;
					found_operator = cFormulaOperators['-'].prototype;
				} else if ('+' === ph.operand_str) {
					parseResult.operand_expected = true;
					found_operator = cFormulaOperators['+'].prototype;
				} else if (':' === ph.operand_str) {
					parseResult.operand_expected = true;
					found_operator = cFormulaOperators[':'].prototype;
				} else if ('%' === ph.operand_str) {
					parseResult.operand_expected = false;
					found_operator = cFormulaOperators['%'].prototype;
				} else if (' ' === ph.operand_str && ph.pCurrPos === t.Formula.length) {
					return true;
				} else {
					if (ph.operand_str in cFormulaOperators) {
						found_operator = cFormulaOperators[ph.operand_str].prototype;
						parseResult.operand_expected = true;
					} else {
						if (ignoreErrors) {
							return true;
						} else {
							parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
							t.outStack = [];
							return false;
						}
					}
				}
			}

			while (0 !== elemArr.length && (found_operator.rightAssociative ? (found_operator.priority < elemArr[elemArr.length - 1].priority) :
					(found_operator.priority <= elemArr[elemArr.length - 1].priority))) {
				t.outStack.push(elemArr.pop());
			}
			elemArr.push(found_operator);
			parseResult.addElem(found_operator);
			found_operand = null;
			return true;
		};

		var parseLeftParentheses = function () {
			if (wasRigthParentheses || found_operand) {
				elemArr.push(new cMultOperator());
			}
			parseResult.operand_expected = true;
			wasLeftParentheses = true;
			wasRigthParentheses = false;
			found_operand = null;
			elemArr.push(cFormulaOperators[ph.operand_str].prototype);
			parseResult.addElem(cFormulaOperators[ph.operand_str].prototype);
			leftParentArgumentsCurrentArr[elemArr.length - 1] = 1;
			parseResult.argPos = 1;

			if (startArrayFunc) {
				counterArrayFunc++;
				if (1 === counterArrayFunc) {
					t.outStack.push(cSpecialOperandStart.prototype);
				}
			}

			argFuncMap[currentFuncLevel] = {count: 0, startPos: ph.pCurrPos + 1};
			argPosArrMap[currentFuncLevel] = [{start: ph.pCurrPos + 1}];
		};

		var parseRightParentheses = function () {

			parseResult.addElem(cFormulaOperators[ph.operand_str].prototype);
			wasRigthParentheses = true;
			var top_elem = null;
			var top_elem_arg_count = 0;
			if (0 !== elemArr.length && ((top_elem = elemArr[elemArr.length - 1]).name === '(') && parseResult.operand_expected) {
				top_elem_arg_count = leftParentArgumentsCurrentArr[elemArr.length - 1];
				if (top_elem_arg_count > 1) {
					t.outStack.push(new cEmpty());
				} else {
					leftParentArgumentsCurrentArr[elemArr.length - 1]--;
					top_elem_arg_count = leftParentArgumentsCurrentArr[elemArr.length - 1];
				}
			} else {
				while (0 !== elemArr.length && !((top_elem = elemArr[elemArr.length - 1]).name === '(')) {
					if (top_elem.name in cFormulaOperators && parseResult.operand_expected) {
						parseResult.setError(c_oAscError.ID.FrmlOperandExpected);
						t.outStack = [];
						return false;
					}
					t.outStack.push(elemArr.pop());
				}
				top_elem_arg_count = leftParentArgumentsCurrentArr[elemArr.length - 1];
			}

			if ((0 === elemArr.length || null === top_elem)) {
				parseResult.setError(c_oAscError.ID.FrmlWrongCountParentheses);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			}

			var p = top_elem, func, bError = false;
			elemArr.pop();
			if (0 !== elemArr.length && (func = elemArr[elemArr.length - 1]).type === cElementType.func) {
				p = elemArr.pop();
				if (top_elem_arg_count > func.argumentsMax) {
					parseResult.setError(c_oAscError.ID.FrmlWrongMaxArgument);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				} else {
					if (top_elem_arg_count >= func.argumentsMin) {
						t.outStack.push(null !== startArrayArg && startArrayArg < currentFuncLevel ? -top_elem_arg_count : top_elem_arg_count);
						if (!func.checkArguments(top_elem_arg_count)) {
							bError = true;
						}
					} else {
						bError = true;
					}

					if (bError) {
						parseResult.setError(c_oAscError.ID.FrmlWrongCountArgument);
						if (!ignoreErrors) {
							t.outStack = [];
							return false;
						}
					}
				}
				parseResult.argPos = leftParentArgumentsCurrentArr[elemArr.length - 1];
			} else if (wasLeftParentheses && 0 === top_elem_arg_count && elemArr[elemArr.length - 1] /*&& " " === elemArr[elemArr.length - 1].name*/) {
				//intersection with empty range
				parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			} else {
				if (wasLeftParentheses && (!elemArr[elemArr.length - 1] || '(' === elemArr[elemArr.length - 1].name)) {
					parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}
				// for (int i = 0; i < left_p.ParametersNum - 1; ++i)
				// {
				// ptgs_list.AddFirst(new PtgUnion()); // чета нужно добавить для Union.....
				// }
			}
			t.outStack.push(p);
			parseResult.operand_expected = false;
			wasLeftParentheses = false;

			if (startArrayFunc) {
				counterArrayFunc--;
				if (counterArrayFunc < 1) {
					startArrayFunc = false;
					t.outStack.push(cSpecialOperandEnd.prototype);
				}
			}

			if (func && func.type === cElementType.func) {
				if (needCalcArgPos) {
					if (needFuncLevel > 0) {
						needFuncLevel--;
					}
					if (!parseResult.activeFunction && levelFuncMap[currentFuncLevel] && levelFuncMap[currentFuncLevel].startPos <= activePos && activePos <= ph.pCurrPos) {
						parseResult.activeFunction = {func: levelFuncMap[currentFuncLevel].func, start: levelFuncMap[currentFuncLevel].startPos, end: ph.pCurrPos};
						parseResult.argPosArr = argPosArrMap[currentFuncLevel];
					}
				}
				var _argPos = argPosArrMap[currentFuncLevel];
				var lastArgPos = _argPos && _argPos[_argPos.length - 1];
				if (lastArgPos && undefined === lastArgPos.end) {
					lastArgPos.end = lastArgPos.start > ph.pCurrPos ? lastArgPos.start : ph.pCurrPos;
				}

				if (!parseResult.allFunctionsPos) {
					parseResult.allFunctionsPos = [];
				}
				parseResult.allFunctionsPos.push({func: levelFuncMap[currentFuncLevel].func, start: levelFuncMap[currentFuncLevel].startPos, end: ph.pCurrPos, args: _argPos});

				currentFuncLevel--;
			}

			return true;
		};

		const parseCommaAndArgumentsUnion = function () {
			wasLeftParentheses = false;
			wasRigthParentheses = false;
			let stackLength = elemArr.length, top_elem = null, top_elem_arg_pos;

			if (elemArr.length !== 0 && elemArr[stackLength - 1].name === "(" &&
				((!elemArr[stackLength - 2]) || (elemArr[stackLength - 2] && elemArr[stackLength - 2].type !== cElementType.func))) {
				parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			} else if (elemArr.length !== 0 && elemArr[stackLength - 1].name === "(" && parseResult.operand_expected) {
				t.outStack.push(new cEmpty());
				top_elem = elemArr[stackLength - 1];
				top_elem_arg_pos = stackLength - 1;
				wasLeftParentheses = true;
				parseResult.operand_expected = false;
			} else {
				while (stackLength !== 0) {
					top_elem = elemArr[stackLength - 1];
					top_elem_arg_pos = stackLength - 1;
					if (top_elem.name === "(") {
						wasLeftParentheses = true;
						break;
					} else {
						t.outStack.push(elemArr.pop());
						stackLength = elemArr.length;
					}
				}
			}

			if (parseResult.operand_expected) {
				parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			}

			//TODO заглушка для парсинга множественного диапазона в _xlnm.Print_Area. необходимо сделать общий парсинг подобного содержимого
			if (!wasLeftParentheses && !(t.parent && t.parent instanceof window['AscCommonExcel'].DefName /*&& t.parent.name === "_xlnm.Print_Area"*/)) {
				parseResult.setError(c_oAscError.ID.FrmlWrongCountParentheses);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			}
			leftParentArgumentsCurrentArr[top_elem_arg_pos]++;
			parseResult.argPos = leftParentArgumentsCurrentArr[top_elem_arg_pos];
			parseResult.operand_expected = true;

			if (needCalcArgPos) {
				if (needFuncLevel === 1) {
					//parseResult.argPosArr.push(ph.pCurrPos);
				}
				if (argFuncMap[currentFuncLevel] && argFuncMap[currentFuncLevel].startPos <= activePos && activePos <= ph.pCurrPos) {
					parseResult.activeArgumentPos = argFuncMap[currentFuncLevel].count;
				}
			}
			if (argPosArrMap[currentFuncLevel] && levelFuncMap[currentFuncLevel]) {
				//проверяем, вдруг данная функция может принимать в качестве данного аргумента массив
				var _curFunc = levelFuncMap[currentFuncLevel].func;
				var _curArg = argPosArrMap[currentFuncLevel].length;
				if (_curFunc.argumentsType && Asc.c_oAscFormulaArgumentType.reference === _curFunc.argumentsType[_curArg]) {
					if (null === startArrayArg || startArrayArg > currentFuncLevel) {
						startArrayArg = currentFuncLevel;
					}
				} else if (currentFuncLevel <= startArrayArg) {
					startArrayArg = null;
				}


				argPosArrMap[currentFuncLevel][argPosArrMap[currentFuncLevel].length - 1].end = ph.pCurrPos;
				argPosArrMap[currentFuncLevel][argPosArrMap[currentFuncLevel].length] = {start: ph.pCurrPos + 1};

				argFuncMap[currentFuncLevel].count++;
				argFuncMap[currentFuncLevel].startPos = ph.pCurrPos + 1;
			}

			return true;
		};

		var parseArray = function () {
			if (!_checkReferenceCount(2)) {
				return false;
			}
			wasLeftParentheses = false;
			wasRigthParentheses = false;
			var arr = new cArray(), operator = {isOperator: false, operatorName: ""};
			while (ph.pCurrPos < t.Formula.length && !parserHelp.isRightBrace.call(ph, t.Formula, ph.pCurrPos)) {
				if (parserHelp.isArraySeparator.call(ph, t.Formula, ph.pCurrPos, digitDelim)) {
					if (ph.operand_str === (digitDelim ? FormulaSeparators.arrayRowSeparator : FormulaSeparators.arrayRowSeparatorDef)) {
						arr.addRow();
					}
				} else if (parserHelp.isBoolean.call(ph, t.Formula, ph.pCurrPos, local)) {
					arr.addElement(new cBool(ph.operand_str));
				} else if (parserHelp.isString.call(ph, t.Formula, ph.pCurrPos)) {
					arr.addElement(new cString(ph.operand_str));
				} else if (parserHelp.isError.call(ph, t.Formula, ph.pCurrPos)) {
					arr.addElement(new cError(ph.operand_str));
				} else if (parserHelp.isNumber.call(ph, t.Formula, ph.pCurrPos, digitDelim)) {
					if (operator.isOperator) {
						if (operator.operatorName === "+" || operator.operatorName === "-") {
							ph.operand_str = operator.operatorName + "" + ph.operand_str
						} else {
							t.outStack = [];
							parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
							return false;
						}
					}
					arr.addElement(new cNumber(parseFloat(ph.operand_str)));
					operator = {isOperator: false, operatorName: ""};
				} else if (parserHelp.isOperator.call(ph, t.Formula, ph.pCurrPos)) {
					operator.isOperator = true;
					operator.operatorName = ph.operand_str;
				} /*else if(ignoreErrors && parserHelp.isFunc.call(ph, t.Formula, ph.pCurrPos)) {
					//TODO при нахождении функции внутри массива ms выдаёт подсказки к аргументам данной функции(lookup(,{,3,sum()
					//если расскоментировать данный код, то проверка на функцию должна осуществляться, необходимо проверить!

					if (wasRigthParentheses && parseResult.operand_expected) {
						elemArr.push(new cMultOperator());
					}

					var found_operator = null, operandStr = ph.operand_str.replace(rx_sFuncPref, "").toUpperCase();
					if (operandStr in cFormulaList) {
						found_operator = cFormulaList[operandStr].prototype;
					} else if (operandStr in cAllFormulaFunction) {
						found_operator = cAllFormulaFunction[operandStr].prototype;
					} else {
						found_operator = new cUnknownFunction(operandStr);
						found_operator.isXLFN = ( ph.operand_str.indexOf("_xlfn.") === 0 );
					}

					if (found_operator !== null) {
						if (found_operator.ca) {
							t.ca = found_operator.ca;
						}
						elemArr.push(found_operator);
						parseResult.addElem(found_operator);
						if("SUMPRODUCT" === found_operator.name){
							startSumproduct = true;
						}
					} else if(!ignoreErrors) {
						parseResult.setError(c_oAscError.ID.FrmlWrongFunctionName);
						t.outStack = [];
						return false;
					}
					parseResult.operand_expected = false;
					wasRigthParentheses = false;
					return true;
				}*/ else {
					//убираю проверку на ignoreErrors из-за зацикливания в формулах типа lookup(,{,3,sum(
					t.outStack = [];
					/*в массиве используется недопустимый параметр*/
					parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
					return false;
				}
			}
			if (!arr.isValidArray()) {
				/*размер массива не согласован*/
				parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			}
			t.outStack.push(arr);
			parseResult.operand_expected = false;
			return true;
		};
		const isRecursiveFormula = function (found_operand, parserFormula) {
			const nOperandType = found_operand.type;
			let oRange = null;
			let bRecursiveCell = parserFormula.ca;
			let sFunctionName = "";


			if (levelFuncMap.length && levelFuncMap[currentFuncLevel]) {
				sFunctionName = levelFuncMap[currentFuncLevel].func.name;
			}
			if (!bConditionalFormula) {
				bConditionalFormula = parserFormula._isConditionalFormula(sFunctionName);
			}
			if (parserFormula.getParent() == null) {
				return bRecursiveCell;
			}
			if (parserFormula.ca) {
				return bRecursiveCell;
			}
			if (sFunctionName && aExcludeRecursiveFormulas.includes(sFunctionName)) {
				return bRecursiveCell;
			}
			if (bConditionalFormula) {
				return bRecursiveCell;
			}
			if (nOperandType === cElementType.cellsRange || nOperandType === cElementType.cellsRange3D) {
				return parserFormula._isAreaContainCell(found_operand);
			}
			if (nOperandType === cElementType.name || nOperandType === cElementType.name3D) {
				const oElemValue = found_operand.getValue();
				const oElemType = oElemValue.type;
				let aRef = [cElementType.cell, cElementType.cell3D, cElementType.cellsRange, cElementType.cellsRange3D];
				if (!aRef.includes(oElemType)) {
					return bRecursiveCell;
				}
				oRange = oElemValue.getRange();
				if (oElemType === cElementType.cellsRange || oElemType === cElementType.cellsRange3D) {
					return oRange.containCell2(parserFormula.getParent());
				}
			} else if (nOperandType === cElementType.table) {
				let oRefElem = found_operand.toRef();
				oRange = oRefElem.getRange();
			} else {
				oRange = found_operand && found_operand.getRange && found_operand.getRange();
			}

			oRange && oRange._foreachNoEmpty(function (oCell) {
				if (!bRecursiveCell) {
					bRecursiveCell = oCell.checkRecursiveFormula(parserFormula.getParent());
				}
			});

			return bRecursiveCell;
		};

		var parseOperands = function () {
			found_operand = null;

			let needSplitString = true;
			let removeUnarOperator = false;
			let _doSplitString = function(str) {
				// Cache concatenation operator to avoid repeated lookups
				const concatOperator = cFormulaOperators["&"].prototype;

				// Process first part of string
				let currentPos = 0;
				const strLength = str.length;

				while (currentPos < strLength) {
					// If this is not the first part of string, add concatenation operator
					if (currentPos > 0) {
						// Add concatenation operator
						parseResult.operand_expected = true;

						// Handle operator precedence
						while (elemArr.length > 0 &&
						(concatOperator.rightAssociative ?
							(concatOperator.priority < elemArr[elemArr.length - 1].priority) :
							(concatOperator.priority <= elemArr[elemArr.length - 1].priority))) {
							t.outStack.push(elemArr.pop());
						}

						elemArr.push(concatOperator);
						parseResult.addElem(concatOperator);
					}

					// Calculate length of current part
					const partLength = Math.min(g_nFormulaStringMaxLength, strLength - currentPos);
					const part = str.slice(currentPos, currentPos + partLength);

					// Create string operand
					const stringOperand = new cString(part);

					// If this is the last part of string
					if (currentPos + partLength === strLength) {
						found_operand = stringOperand;
						break;
					}

					// Add operand to stack
					t.outStack.push(stringOperand);
					parseResult.addElem(stringOperand);
					parseResult.operand_expected = false;

					// Move to next part
					currentPos += partLength;
				}
			}

			if (wasRigthParentheses) {
				parseResult.operand_expected = true;
			}

			if (!parseResult.operand_expected) {
				parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
				if (!ignoreErrors) {
					t.outStack = [];
					return false;
				}
			}
			var prevCurrPos = ph.pCurrPos;

			/* Booleans */
			if (opt_pivotNamesList && (_tableTMP = opt_pivotNamesList.length === 0 ? parserHelp.isPivotRaw.call(ph, t.Formula, ph.pCurrPos, local) : parserHelp.isPivot.call(ph, t.Formula, ph.pCurrPos, local, opt_pivotNamesList))) {

				found_operand = cStrucPivotTable.prototype.createFromVal(_tableTMP);

				//todo undo delete column
				if (found_operand.type === cElementType.error) {
					/*используется неверный именованный диапазон или таблица*/
					parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}

				if (!_checkReferenceCount(2)) {
					return false;
				}
			} else if (parserHelp.isBoolean.call(ph, t.Formula, ph.pCurrPos, local)) {
				if (!_checkReferenceCount(0.5)) {
					return false;
				}
				found_operand = new cBool(ph.operand_str);
			} else if (parserHelp.isString.call(ph, t.Formula, ph.pCurrPos)) { /* Strings */
				if (ph.operand_str.length > g_nFormulaStringMaxLength) {
					if (needSplitString) {
						if (!_checkReferenceCount(ph.operand_str.length * 0.25 + 0.5)) {
							return false;
						}
						_doSplitString(ph.operand_str);
					} else {
						parseResult.setError(c_oAscError.ID.FrmlMaxTextLength);
						if (!ignoreErrors) {
							t.outStack = [];
							return false;
						}
					}
				}
				if (!found_operand) {
					if (!_checkReferenceCount(ph.operand_str.length * 0.25 + 0.5)) {
						return false;
					}
					found_operand = new cString(ph.operand_str);
				}
			}

			/* Errors */ else if (parserHelp.isError.call(ph, t.Formula, ph.pCurrPos, local)) {
				if (!_checkReferenceCount(1)) {
					return false;
				}
				found_operand = new cError(ph.operand_str);
			}

			/* Referens to 3D area: Sheet1:Sheet3!A1:B3, Sheet1:Sheet3!B3, Sheet1!B3*/ else if ((_3DRefTmp = parserHelp.is3DRef.call(ph, t.Formula, ph.pCurrPos, null, local))[0]) {

				t.is3D = true;

				//renameSheetMap
				if (renameSheetMap) {
					if (renameSheetMap[_3DRefTmp[1]]) {
						_3DRefTmp[1] = renameSheetMap[_3DRefTmp[1]];
						needAssemble = true;
					}
					if (_3DRefTmp[2] && renameSheetMap[_3DRefTmp[2]]) {
						_3DRefTmp[2] = renameSheetMap[_3DRefTmp[2]];
						needAssemble = true;
					}
				}

				let wsF, wsT;
				let sheetName = _3DRefTmp[1];

				let isExternalRefExist, externalLink, receivedLink, externalName, 
					createShortLink, externalProps, isCurrentFile, currentFileDefname;

				/* these flags are needed to further check the link to the current file */
				let isShortLink = _3DRefTmp[4] ? true : false;	// _3DRefTmp[4] - shortlink info
				let isFullLink = _3DRefTmp[5] ? true : false;	// _3DRefTmp[5] - current file defname from full link

				if (isShortLink) {
					externalProps = t.wb && t.wb.externalReferenceHelper && t.wb.externalReferenceHelper.check3dRef(_3DRefTmp, local);
				} else {
					externalLink = _3DRefTmp[3];
					externalName = _3DRefTmp[3];
				}

				if (!externalProps && !sheetName) {
					parseResult.setError(c_oAscError.ID.FrmlWrongReferences);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				} else if (externalProps) {
					externalLink = externalProps.externalLink;
					externalName = externalProps.externalName;
					receivedLink = externalProps.receivedLink;
					createShortLink = externalProps.isShortLink;
					isCurrentFile = externalProps.isCurrentFile;
					currentFileDefname = externalProps.currentFileDefname;
				}
				
				if (externalProps && !sheetName) {
					sheetName = externalProps.sheetName ? externalProps.sheetName : externalName;
				}

				/* if the link is not short, then we check whether we received the currentFileDefname argument, which indicates whether there is a link to the current file */
				if (!isShortLink && isFullLink) {
					if (_3DRefTmp[1] && !t.wb.getWorksheetByName(_3DRefTmp[1])) {
						// if there is sheetname in the arguments and this sheet is not exist in wb, return an error
						parseResult.setError(c_oAscError.ID.FrmlWrongReferences);
						if (!ignoreErrors) {
							t.outStack = [];
							return false;
						}
					}

					// create shortlink flag
					createShortLink = true;
					isCurrentFile = true;
					externalLink = null;
					currentFileDefname = _3DRefTmp[5];
				}

				if (externalLink && !isCurrentFile) {
					if (local) {
						externalLink = t.wb.getExternalLinkIndexByName(externalLink);
						if (externalLink === null) {
							externalLink = receivedLink ? receivedLink : _3DRefTmp[3];
							if (!parseResult.externalReferenesNeedAdd) {
								parseResult.externalReferenesNeedAdd = [];
							}
							if (!parseResult.externalReferenesNeedAdd[externalLink]) {
								parseResult.externalReferenesNeedAdd[externalLink] = [];
							}
							parseResult.externalReferenesNeedAdd[externalLink].push({sheet: sheetName /*_3DRefTmp[1]*/});
						} else {
							isExternalRefExist = true;
							if (!parseResult.externalReferenesNeedAdd) {
								parseResult.externalReferenesNeedAdd = [];
							}
							if (!parseResult.externalReferenesNeedAdd[externalName]) {
								parseResult.externalReferenesNeedAdd[externalName] = [];
							}
							parseResult.externalReferenesNeedAdd[externalName].push({sheet: sheetName /*_3DRefTmp[1]*/});
						}
					}

					wsF = sheetName ? t.wb.getExternalWorksheet(externalLink, sheetName /*_3DRefTmp[1]*/) : null;

					if (externalLink && !local && !wsF) {
						// special case when opening a file:
						// if we refer to defname that doesn't exist, but the ER itself exists, then we refer to the first existing worksheet
						// since we don't know the name of the sheet in the short link and defname doesn't exist
						wsF =  t.wb.getExternalWorksheet(externalLink, sheetName, true /* getFirtsSheet */);
						if (!wsF) {
							parseResult.setError(c_oAscError.ID.FrmlWrongReferences);
							if (!ignoreErrors) {
								t.outStack = [];
								return false;
							}
						}
					}

					wsT = wsF;
				} else {
					// isCurrentFileCheck
					let currentDefname, sheet;
					if (isCurrentFile && currentFileDefname /*&& !local*/) {
						// looking for defname from this sheet
						currentDefname = t.wb.getDefinesNames(currentFileDefname);
						if (!currentDefname && sheetName && isFullLink) {
							wsF = t.wb.getWorksheetByName(sheetName);
						} else if (!currentDefname && !isFullLink) {
							sheet = t.wb.getActiveWs();
							wsF = t.wb.getWorksheetByName(sheet.getName());
						} else {
							let exclamationMarkIndex = currentDefname.ref && currentDefname.ref.lastIndexOf("!");
							sheet = currentDefname.ref.slice(0, exclamationMarkIndex);
							wsF = t.wb.getWorksheetByName(sheet);
						}
					}

					wsF = wsF ? wsF : t.wb.getWorksheetByName(sheetName/*_3DRefTmp[1]*/);
					wsT = (null !== _3DRefTmp[2]) ? t.wb.getWorksheetByName(_3DRefTmp[2]) : wsF;
				}

				// if it's impossible to get a sheet from an external file, but the file itself is exist, then we return an error about incorrectly entering the formula
				let wsNotExist = externalLink && isExternalRefExist && !wsF;

				if ((!(wsF && wsT) && !externalLink) /*|| wsNotExist*/) {
					parseResult.setError(c_oAscError.ID.FrmlWrongReferences);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}

				if (!_checkReferenceCount(null !== _3DRefTmp[2] ? 3 : 2)) {
					return false;
				}

				if (parserHelp.isArea.call(ph, t.Formula, ph.pCurrPos)) {
					if (!(wsF && wsT)) {
						//for edit formula mode
						//found_operand = new cUnknownFunction(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase());
						found_operand = new cName(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), t.ws);
					} else {
						found_operand = new cArea3D(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), wsF, wsT, externalLink);
					}
					parseResult.addRefPos(prevCurrPos, ph.pCurrPos, t.outStack.length, found_operand);
					if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
						t.ca = isRecursiveFormula(found_operand, t);
					}
				} else if (parserHelp.isRef.call(ph, t.Formula, ph.pCurrPos)) {
					if (!(wsF && wsT)) {
						//for edit formula mode
						//found_operand = new cUnknownFunction(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase());
						found_operand = new cName(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), t.ws);
					} else if (wsT !== wsF) {
						found_operand = new cArea3D(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), wsF, wsT, externalLink);
					} else {
						found_operand = new cRef3D(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), wsF, externalLink);
					}
					parseResult.addRefPos(prevCurrPos, ph.pCurrPos, t.outStack.length, found_operand);
					if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
						t.ca = isRecursiveFormula(found_operand, t);
					}
				} else {
					parserHelp.isName.call(ph, t.Formula, ph.pCurrPos);
					// if link to the same file - set external link to zero just like in MS
					found_operand = new cName3D(ph.operand_str, wsF, isCurrentFile ? "0" : externalLink, createShortLink);
					parseResult.addRefPos(prevCurrPos, ph.pCurrPos, t.outStack.length, found_operand);
					if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
						t.ca = isRecursiveFormula(found_operand, t);
					}
				}
			}

			/* Referens to cells area A1:A10 */ else if (parserHelp.isArea.call(ph, t.Formula, ph.pCurrPos)) {
				if (!_checkReferenceCount(2)) {
					return false;
				}
				found_operand = new cArea(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), t.ws);
				parseResult.addRefPos(ph.pCurrPos - ph.operand_str.length, ph.pCurrPos, t.outStack.length, found_operand);
				if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
					t.ca = isRecursiveFormula(found_operand, t);
				}
			}
			/* Referens to cell A4 */ else if (parserHelp.isRef.call(ph, t.Formula, ph.pCurrPos)) {
				if (!_checkReferenceCount(1)) {
					return false;
				}
				found_operand = new cRef(ph.real_str ? ph.real_str.toUpperCase() : ph.operand_str.toUpperCase(), t.ws);
				parseResult.addRefPos(ph.pCurrPos - ph.operand_str.length, ph.pCurrPos, t.outStack.length, found_operand);

				if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
					t.ca = isRecursiveFormula(found_operand, t);
				}
			} else if (_tableTMP = parserHelp.isTable.call(ph, t.Formula, ph.pCurrPos, local, t)) {
				found_operand = cStrucTable.prototype.createFromVal(_tableTMP, t.wb, t.ws, tablesMap);

				//todo undo delete column
				if (found_operand.type === cElementType.error) {
					/*используется неверный именованный диапазон или таблица*/
					parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}

				if (!_checkReferenceCount(2)) {
					return false;
				}

				if (found_operand.type !== cElementType.error) {
					parseResult.addRefPos(ph.pCurrPos - ph.operand_str.length, ph.pCurrPos, t.outStack.length, found_operand, true);
				}
				if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
					t.ca = isRecursiveFormula(found_operand, t);
				}
			}

			/* Referens to DefinedNames */ else if (parserHelp.isName.call(ph, t.Formula, ph.pCurrPos)) {

				if (ph.operand_str.length > g_nFormulaStringMaxLength || !AscCommon.rx_r1c1DefError.test(ph.operand_str)) {
					//TODO стоит добавить новую ошибку
					parseResult.setError(c_oAscError.ID.FrmlWrongOperator);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}

				if (!_checkReferenceCount(0.75)) {
					return false;
				}

				//проверяем вдруг это область печати
				var defName;
				var sDefNameOperand = ph.operand_str.replace(rx_sDefNamePref, "");
				var tryTranslate = AscCommonExcel.tryTranslateToPrintArea(sDefNameOperand);
				if (tryTranslate) {
					found_operand = new cName(tryTranslate, t.ws);
					defName = found_operand.getDefName();
				}
				//TODO возможно здесь нужно else ставить
				if (!defName) {
					found_operand = new cName(sDefNameOperand, t.ws);
					defName = found_operand.getDefName();
				}

				if (defName && defName.type === Asc.c_oAscDefNameType.table && (_tableTMP = parserHelp.isTable(sDefNameOperand + "[]", 0))) {
					found_operand = cStrucTable.prototype.createFromVal(_tableTMP, t.wb, t.ws);
					//need assemble becase source formula wrong
					needAssemble = true;
				}
				parseResult.addRefPos(ph.pCurrPos - ph.operand_str.length, ph.pCurrPos, t.outStack.length, found_operand, true);
				if (local || (local === false && digitDelim === false)) { // local and digitDelim with value false using only for copypaste mode.
					t.ca = isRecursiveFormula(found_operand, t);
				}
				if (t.ca && defName && defName.parsedRef) {
					defName.parsedRef.ca = t.ca;
				}
			}

			/* Numbers*/ else if (parserHelp.isNumber.call(ph, t.Formula, ph.pCurrPos, digitDelim)) {
				if (ph.operand_str !== "." && parseResult.checkNumberOperator(elemArr)) {
					var _number = parseFloat(ph.operand_str);
					//TODO для отрицательныз числе необходимо сделать проверку
					if (!_checkReferenceCount((_number >= 65536 || !Number.isInteger(_number)) ? 1.25 : 0.5)) {
						return false;
					}
					found_operand = new cNumber(_number);
					if (local) {
						let lastElem = elemArr[elemArr.length-1];
						if (lastElem && lastElem.name && lastElem.name === "un_plus") {
							removeUnarOperator = true;
						}
					}
				} else {
					parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}
			}

			/* Function*/ else if (parserHelp.isFunc.call(ph, t.Formula, ph.pCurrPos)) {

				if (wasRigthParentheses && parseResult.operand_expected) {
					elemArr.push(new cMultOperator());
				}

				var found_operator = null, operandStr = ph.operand_str.replace(rx_sFuncPref, "").replace(rx_sFuncPrefXlWS, "").replace(rx_sFuncPrefXLUFD, "").toUpperCase();
				if (operandStr in cFormulaList) {
					found_operator = cFormulaList[operandStr].prototype;
				} else if (operandStr in cAllFormulaFunction) {
					found_operator = cAllFormulaFunction[operandStr].prototype;
				} else {
					found_operator = new cUnknownFunction(operandStr);
					let xlfnFrefix = "_xlfn.";
					let xlwsFrefix = "_xlws.";
					//let xludfFrefix = "__xludf.DUMMYFUNCTION.";

					//_xlws only together with _xlfn
					found_operator.isXLFN = (ph.operand_str.indexOf(xlfnFrefix) === 0);
					found_operator.isXLWS = found_operator.isXLFN && xlfnFrefix.length === ph.operand_str.indexOf(xlwsFrefix);

					t.unknownOrCustomFunction = operandStr;
				}

				//mark function, when need reparse and recalculate on custom function change
				let wb = Asc["editor"] && Asc["editor"].wb;
				if (wb && wb.customFunctionEngine && wb.customFunctionEngine.getFunc(operandStr)) {
					t.unknownOrCustomFunction = operandStr;
				}

				if (found_operator !== null) {
					var _nullArgCount = t.Formula[ph.pCurrPos + 1] === ")";
					if (!_checkReferenceCount(_nullArgCount ? 0.6 : 0.5)) {
						return false;
					}
					if (found_operator.ca) {
						t.ca = found_operator.ca;
					}
					elemArr.push(found_operator);
					parseResult.addElem(found_operator);
					if (arrayFunctionsMap[found_operator.name]) {
						startArrayFunc = true;
					}

					if (found_operator.name === "IMPORTRANGE") {
						isFoundImportFunctions = true;
					}

					if (needCalcArgPos) {
						if (activePos === undefined) {
							needFuncLevel++;
							activePos = ph.pCurrPos + 1;
						} else if (needFuncLevel > 0) {
							needFuncLevel++;
						}
					}
					currentFuncLevel++;
					levelFuncMap[currentFuncLevel] = {func: found_operator, startPos: ph.pCurrPos - ph.operand_str.length};

					if (found_operator.argumentsType && Asc.c_oAscFormulaArgumentType.reference === found_operator.argumentsType[0]) {
						if (null === startArrayArg || startArrayArg > currentFuncLevel) {
							startArrayArg = currentFuncLevel;
						}
					} else if (currentFuncLevel <= startArrayArg) {
						startArrayArg = null;
					}

				} else {
					parseResult.setError(c_oAscError.ID.FrmlWrongFunctionName);
					if (!ignoreErrors) {
						t.outStack = [];
						return false;
					}
				}
				parseResult.operand_expected = false;
				wasRigthParentheses = false;
				return true;
			}

			if (null !== found_operand) {
				if (removeUnarOperator) {
					elemArr.pop();
				}

				t.outStack.push(found_operand);
				parseResult.addElem(found_operand);
				parseResult.operand_expected = false;
				found_operand = null;
			} else {
				t.outStack.push(new cError(cErrorType.wrong_name));
				parseResult.setError(c_oAscError.ID.FrmlAnotherParsingError);
				return t.isParsed = false;
			}

			if (wasRigthParentheses) {
				elemArr.push(new cMultOperator());
			}
			wasLeftParentheses = false;
			wasRigthParentheses = false;
			return true;
		};

		var setArgInfo = function () {
			if (needCalcArgPos) {
				if (needAddCursorPos) {
					parseResult.cursorPos = activePos;
				}

				if (!parseResult.activeFunction && levelFuncMap[currentFuncLevel] && levelFuncMap[currentFuncLevel].startPos <= activePos && activePos <= ph.pCurrPos + 1) {
					parseResult.activeFunction = {func: levelFuncMap[currentFuncLevel].func, start: levelFuncMap[currentFuncLevel].startPos, end: ph.pCurrPos + 1};
					parseResult.argPosArr = argPosArrMap[currentFuncLevel];
				}
				if (undefined === parseResult.activeArgumentPos && argFuncMap[currentFuncLevel] && argFuncMap[currentFuncLevel].startPos <= activePos && activePos <= ph.pCurrPos +
					1) {
					parseResult.activeArgumentPos = argFuncMap[currentFuncLevel].count;
				}
				var _argPos = argPosArrMap[currentFuncLevel];
				var lastArgPos = _argPos && _argPos[_argPos.length - 1];
				if (lastArgPos && undefined === lastArgPos.end) {
					lastArgPos.end = lastArgPos.start > ph.pCurrPos ? lastArgPos.start : ph.pCurrPos;
				}
				if (levelFuncMap[currentFuncLevel]) {
					if (!parseResult.allFunctionsPos) {
						parseResult.allFunctionsPos = [];
					}
					parseResult.allFunctionsPos.push({func: levelFuncMap[currentFuncLevel].func, start: levelFuncMap[currentFuncLevel].startPos, end: ph.pCurrPos, args: _argPos});
				}
			}
		};

		while (ph.pCurrPos < this.Formula.length) {
			ph.operand_str = this.Formula[ph.pCurrPos];

			//TODO сделать так, чтобы добавлялся особый элемент - перенос строки и учитывался при сборке!!!!
			if (ph.operand_str == "\n") {
				ph.pCurrPos++;
				continue;
			}

			/* Operators*/
			if (parserHelp.isOperator.call(ph, this.Formula, ph.pCurrPos) || parserHelp.isNextPtg.call(ph, this.Formula, ph.pCurrPos)) {
				if (!parseOperators()) {
					if (ignoreErrors) {
						setArgInfo();
					}
					return false;
				}
			} /* Left Parentheses*/ else if (parserHelp.isLeftParentheses.call(ph, this.Formula, ph.pCurrPos)) {
				parseLeftParentheses();

				//TODO протестировать
				//если осталось только закрыть скобки за функции с нулевым количеством аргументов
				if (ph.pCurrPos === this.Formula.length) {
					if (elemArr[elemArr.length - 2] && 0 === elemArr[elemArr.length - 2].argumentsMax) {
						parseResult.operand_expected = false;
					}
				}

			}/* Right Parentheses */ else if (parserHelp.isRightParentheses.call(ph, this.Formula, ph.pCurrPos)) {
				if (!parseRightParentheses()) {
					if (ignoreErrors) {
						setArgInfo();
					}
					return false;
				}
			}/*Comma & arguments union*/ else if (parserHelp.isComma.call(ph, this.Formula, ph.pCurrPos)) {
				if (!parseCommaAndArgumentsUnion()) {
					if (ignoreErrors) {
						setArgInfo();
					}
					return false;
				}
			}/* Array */ else if (parserHelp.isLeftBrace.call(ph, this.Formula, ph.pCurrPos)) {
				if (!parseArray()) {
					if (ignoreErrors) {
						setArgInfo();
					}
					return false;
				}
			}/* Operands*/ else {
				if (!parseOperands()) {
					if (ignoreErrors) {
						setArgInfo();
					}
					return false;
				}
			}
		}

		setArgInfo();
		if (parseResult.operand_expected) {
			this.outStack = [];
			parseResult.setError(c_oAscError.ID.FrmlOperandExpected);
			return false;
		}
		var operand, parenthesesNotEnough = false;
		while (0 !== elemArr.length) {
			operand = elemArr.pop();
			if ('(' === operand.name) {
				this.Formula += ")";
				parenthesesNotEnough = true;
			} else if ('(' === operand.name || ')' === operand.name) {
				this.outStack = [];
				parseResult.setError(c_oAscError.ID.FrmlWrongCountParentheses);
				return false;
			} else {
				this.outStack.push(operand);
			}
		}
		if (bConditionalFormula && t.getParent() && t.getParent() instanceof AscCommonExcel.CCellWithFormula && !t.ca && !ignoreErrors) {
			t.ca = t.isRecursiveCondFormula(levelFuncMap[0].func.name);
			t.outStack.forEach(function (oOperand) {
				if (oOperand.type === cElementType.name || oOperand.type === cElementType.name3D) {
					let oDefName = oOperand.getDefName();
					if (t.ca && oDefName && oDefName.parsedRef) {
						oDefName.parsedRef.ca = t.ca;
					}
				}
			});
		}
		if (parenthesesNotEnough) {
			parseResult.setError(c_oAscError.ID.FrmlParenthesesCorrectCount);
			return this.isParsed = false;
		}

		if (0 !== this.outStack.length) {
			if (needAssemble) {
				this.Formula = this.assemble();
			}
			if (isFoundImportFunctions && !parseResult.error) {
				//share external links
				AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = true;
				this.calculate();
				AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = null;

				this.importFunctionsRangeLinks = AscCommonExcel.importRangeLinksState.importRangeLinks;

				if (this.importFunctionsRangeLinks) {
					for (let i in this.importFunctionsRangeLinks) {
						let externalLink = this.wb.getExternalLinkIndexByName(i);
						if (externalLink === null) {
							externalLink = i;
							if (!parseResult.externalReferenesNeedAdd) {
								parseResult.externalReferenesNeedAdd = [];
							}
							if (!parseResult.externalReferenesNeedAdd[externalLink]) {
								parseResult.externalReferenesNeedAdd[externalLink] = [];
							}

							for (var j = 0; j < this.importFunctionsRangeLinks[i].length; j++) {
								parseResult.externalReferenesNeedAdd[externalLink].push({sheet: this.importFunctionsRangeLinks[i][j].sheet, notUpdateId: true});
							}
						}
					}

					if (AscCommonExcel.importRangeLinksState.importRangeLinks) {
						if (!AscCommonExcel.importRangeLinksState.notUpdateIdMap) {
							AscCommonExcel.importRangeLinksState.notUpdateIdMap = {};
						}
						for (let i in AscCommonExcel.importRangeLinksState.importRangeLinks) {
							AscCommonExcel.importRangeLinksState.notUpdateIdMap[i] = true;
						}
					}

					AscCommonExcel.importRangeLinksState.importRangeLinks = null;
				}
			}
			return this.isParsed = true;
		} else {
			return this.isParsed = false;
		}
	};

	parserFormula.prototype.findRefByOutStack = function (forceCheck) {
		if (AscCommonExcel.bIsSupportDynamicArrays || forceCheck) {
			// using outStack, look at all the arguments in the formulas and compare them with the arrayIndex positions for this formula
			// go through the stack in the same order as .calculate method
			if (this.ref) {
				return true;
			}

			if (this.outStack && this.outStack.length > 0) {
				let elemArr = [], _tmp, currentElement = null, bIsSpecialFunction, argumentsCount, defNameCalcArr, defNameArgCount = 0;
				let length = this.outStack.length;
				let isRef, opt_bbox;

				if (!opt_bbox && this.parent && this.parent.onFormulaEvent) {
					opt_bbox = this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.GetRangeCell);
				}
				if (!opt_bbox) {
					opt_bbox = new Asc.Range(0, 0, 0, 0);
				}

				if (length === 1) {
					let singleElem = this.outStack[0];
					if (singleElem.type === cElementType.cellsRange || singleElem.type === cElementType.cellsRange3D || singleElem.type === cElementType.array) {
						isRef = true;
					} else if (singleElem.type === cElementType.name || singleElem.type === cElementType.name3D) {
						// let defName = singleElem.getDefName();
						let defNameResult = singleElem.Calculate(null, opt_bbox, true);
						if (defNameResult && defNameResult.type === cElementType.array || defNameResult.type === cElementType.cellsRange || defNameResult.type === cElementType.cellsRange3D) {
							isRef = true;
						}
					} else if (singleElem.type === cElementType.table) {
						let tableArea = singleElem.toRef(opt_bbox);
						if (tableArea && tableArea.type === cElementType.cellsRange || tableArea.type === cElementType.cellsRange3D || tableArea.type === cElementType.array) {
							isRef = true;
						}
					}

					if (isRef) {
						return isRef;
					}
				}

				for (let i = 0; i < this.outStack.length; i++) {
					currentElement = this.outStack[i];
					if (!currentElement) {
						continue;
					}

					if(currentElement.name === "(" || currentElement.type === cElementType.specialFunctionStart || currentElement.type === cElementType.specialFunctionEnd || "number" === typeof(currentElement)) {
						continue;
					}

					if (currentElement.type === cElementType.operator || currentElement.type === cElementType.func) {
						argumentsCount = "number" === typeof(this.outStack[i - 1]) ? this.outStack[i - 1] : currentElement.argumentsCurrent;
						if (argumentsCount < 0) {
							argumentsCount = -argumentsCount;
							currentElement.bArrayFormula = true;
						}
						if (elemArr.length < argumentsCount) {
							// elemArr = [];
							// todo test these cases
							return false;
						} else if (argumentsCount + defNameArgCount > currentElement.argumentsMax) {
							// elemArr = [];
							// todo test these cases
							return false;
						} else {
							// if operator - check whether each of the arguments is a range or an array

							let isOperator = currentElement.type === cElementType.operator;
							let arg = [];
							let _isPromise = false;
							for (let i = 0; i < argumentsCount + defNameArgCount; i++) {
								if ("number" === typeof(elemArr[elemArr.length - 1])) {
									elemArr.pop();
								}
								let tempElem = elemArr.pop();
								if (isOperator && (tempElem.type === cElementType.cellsRange || tempElem.type === cElementType.cellsRange3D || tempElem.type === cElementType.array)) {
									isRef = true;
								}

								if (fIsPromise(tempElem)) {
									_isPromise = true;
									break;
								}
								// arg.unshift(elemArr.pop());
								arg.unshift(tempElem);
							}

							if (_isPromise) {
								continue;
							}

							let isCanExpand = null;
							if (currentElement.type === cElementType.func) {
								isCanExpand = cBaseFunction.prototype.checkFormulaArray2.call(currentElement, arg, opt_bbox, null, this, bIsSpecialFunction, argumentsCount);
							} else if (currentElement.type === cElementType.operator && currentElement.bArrayFormula) {
								bIsSpecialFunction = true;
							}

							if(isCanExpand) {
								isRef = true;
							} else {
								/* results of SEQUENCE, RANDARRAY etc... can return an array when using regular values ​​in arguments */
								if (this.unknownOrCustomFunction && currentElement.returnValueType !== AscCommonExcel.cReturnFormulaType.array) {
									return false;
								}
								_tmp = currentElement.Calculate(arg, opt_bbox, null, this.ws, bIsSpecialFunction);
							}

							if (isRef || (_tmp && (_tmp.type === cElementType.array /*|| _tmp.type === cElementType.cellsRange || _tmp.type === cElementType.cellsRange3D*/))) {
								return true;
							}

							defNameArgCount = 0;
							elemArr.push(_tmp);
						}
					} else if (currentElement.type === cElementType.name || currentElement.type === cElementType.name3D) {
						// let defName = currentElement.getDefName();
						defNameCalcArr = currentElement.Calculate(null, opt_bbox, true);
						defNameArgCount = [];
						if(defNameCalcArr && defNameCalcArr.length) {
							defNameArgCount = defNameCalcArr.length - 1;
							for(let j = 0; j < defNameCalcArr.length; j++) {
								elemArr.push(defNameCalcArr[j]);
							}
						} else {
							elemArr.push(defNameCalcArr);
						}
					} else if (currentElement.type === cElementType.table) {
						elemArr.push(currentElement.toRef(opt_bbox));
					} else if (currentElement.type === cElementType.pivotTable) {
						elemArr.push(currentElement.Calculate());
					} else {
						elemArr.push(currentElement);
					}
				}

				return isRef;
			}
			return false;
		}
	};
	parserFormula.prototype.calculate = function (opt_defName, opt_bbox, opt_offset, checkMultiSelect, opt_oCalculateResult, opt_pivotCallback) {
		if (AscCommonExcel.g_LockCustomFunctionRecalculate && this.unknownOrCustomFunction) {
			return;
		}
		if (this.outStack.length < 1) {
			this.value = new cError(cErrorType.wrong_name);
			this._endCalculate();
			return this.value;
		}
		if (!opt_bbox && this.parent && this.parent.onFormulaEvent) {
			opt_bbox = this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.GetRangeCell);
		}
		if (!opt_bbox) {
			opt_bbox = new Asc.Range(0, 0, 0, 0);
		}

		let promiseCounter = 0;
		var elemArr = [], _tmp, numFormat = cNumFormatFirstCell, currentElement = null, bIsSpecialFunction, argumentsCount, defNameCalcArr, defNameArgCount = 0;
		for (var i = 0; i < this.outStack.length; i++) {
			currentElement = this.outStack[i];
			if (currentElement.name === "(") {
				continue;
			}
			if(currentElement.type === cElementType.specialFunctionStart){
				bIsSpecialFunction = true;
				continue;
			}
			if(currentElement.type === cElementType.specialFunctionEnd){
				bIsSpecialFunction = false;
				continue;
			}
			if("number" === typeof(currentElement)){
				continue;
			}

			//TODO пока проставляю у каждого элемента флаг для рассчетов. пересмотреть
			//***array-formula***
			currentElement.bArrayFormula = null;
			if(this.ref) {
				currentElement.bArrayFormula = true;
			}

			/* concatenation should be done as an array formula - via ref */
			if (currentElement.name && currentElement.name === "&") {
				currentElement.bArrayFormula = true;
			}

			if (currentElement.type === cElementType.operator || currentElement.type === cElementType.func) {
				argumentsCount = "number" === typeof(this.outStack[i - 1]) ? this.outStack[i - 1] : currentElement.argumentsCurrent;
				if (argumentsCount < 0) {
					argumentsCount = -argumentsCount;
					currentElement.bArrayFormula = true;
				}
				if (elemArr.length < argumentsCount) {
					elemArr = [];
					this.value = new cError(cErrorType.unsupported_function);
					this._endCalculate();
					return this.value;
				} else if(argumentsCount + defNameArgCount > currentElement.argumentsMax) {
					//возвращаю ошибку в случае если количество аргументов(с учетом тех аргументов, которые получили из именованного диапазона)
					//превышает максимальное допустимое количество аргументов данной функции
					elemArr = [];
					this.value = new cError(cErrorType.wrong_value_type);
					this._endCalculate();
					return this.value;
				} else {
					var arg = [];
					let _isPromise = false;
					for (var ind = 0; ind < argumentsCount + defNameArgCount; ind++) {
						if("number" === typeof(elemArr[elemArr.length - 1])){
							elemArr.pop();
						}
						let _elem = elemArr.pop();
						arg.unshift(_elem);
						if (fIsPromise(_elem)) {
							_isPromise = true;
							break;
						}
					}

					if (_isPromise) {
						break;
					}

					//***array-formula***
					//если данная функция не может возвращать массив, проходимся по всем элементам аргументов и формируем массив
					var formulaArray = null;
					if (currentElement.type === cElementType.func) {
						// checkArgumentsTypes before calculate
						if (opt_oCalculateResult && opt_oCalculateResult.checkOnError && currentElement.exactTypes && !currentElement.checkArgumentsTypes(arg)) {
							this.value = new cError(cErrorType.null_value);
							this._endCalculate();
							opt_oCalculateResult.setError(c_oAscError.ID.FrmlOperandExpected);
							return this.value;
						} else if (!(this.promiseResult && this.promiseResult[i])) {
							formulaArray = cBaseFunction.prototype.checkFormulaArray.call(currentElement, arg, opt_bbox, opt_defName, this, bIsSpecialFunction, argumentsCount);
						}

					} else if (currentElement.type === cElementType.operator && currentElement.bArrayFormula) {
						bIsSpecialFunction = true;
					}

					if (this.promiseResult && this.promiseResult[i]) {
						_tmp = this.promiseResult[i];
					} else if(formulaArray) {
						_tmp = formulaArray;
					} else {
						//if recursion - we must rewrite promise, because arguments can change
						_tmp = !g_cCalcRecursion.getIsEnabledRecursion() && this.wb.asyncFormulasManager.getPromiseByIndex(this._index, i);
						if (!_tmp) {
							_tmp = currentElement.Calculate(arg, opt_bbox, opt_defName, this.ws, bIsSpecialFunction);
						}
					}

					//check promise
					if (fIsPromise(_tmp)) {
						if (this.promiseResult && this.promiseResult[promiseCounter]) {
							_tmp = this.promiseResult[promiseCounter];
						} else {
							_tmp.parserFormula = this;
							_tmp.index = i;
							this.wb.asyncFormulasManager.addPromise(_tmp, true);
						}
						promiseCounter++;
					}

					if (this.unknownOrCustomFunction) {
						/*
						  \@\@\ - perceive it as text, remove the \. do not delete the first formula. The result of the calculation will be ("\@\@\text" -> "@@...text").

						  @@ - special case. delete the first formula and add the result to the cell without @@. ("@@text" -> "text")

						  @@\= - special case. remove the \. delete the first formula and add the result to the cell without @@. ("@@\=text" -> "=text")

						  @@= - special case. delete the formula and add a new formula(result cell) to the cell without @@. ("@@=formula" -> "=formula")

						  For other combinations \@@=, @\@,\@\@\=,@\@\= do not react. Just text result, do not change formula
						*/
						if (_tmp && _tmp.type === cElementType.string) {
							if (0 === _tmp.value.indexOf("@@\\=")) {
								_tmp.value = _tmp.value.slice(3);
								this.replaceFormulaAfterCalc = cReplaceFormulaType.val;
							} else if (0 === _tmp.value.indexOf("\\@\\@")) {
								_tmp.value = "@@" + _tmp.value.slice(4);
							} else if (0 === _tmp.value.indexOf("@@=")) {
								_tmp.value = _tmp.value.slice(2);
								this.replaceFormulaAfterCalc = cReplaceFormulaType.formula;
							} else if (0 === _tmp.value.indexOf("@@")) {
								_tmp.value = _tmp.value.slice(2);
								this.replaceFormulaAfterCalc = cReplaceFormulaType.val;
							}
						}
					}

					//_tmp = currentElement.Calculate(arg, opt_bbox, opt_defName, this.ws, bIsSpecialFunction);
					if (cNumFormatNull !== _tmp.numFormat) {
						numFormat = _tmp.numFormat;
					} else if (0 > numFormat || cNumFormatNone === currentElement.numFormat) {
						numFormat = currentElement.numFormat;
					}

					defNameArgCount = 0;
					elemArr.push(_tmp);
				}
			} else if (currentElement.type === cElementType.name || currentElement.type === cElementType.name3D) {
				var defName = currentElement.getDefName();
				if(defName && defName.parsedRef && this.ref) {
					currentElement.getDefName().parsedRef.ref = this.ref;
				}
				defNameCalcArr = currentElement.Calculate(null, opt_bbox, true);
				defNameArgCount = [];
				if(defNameCalcArr && defNameCalcArr.length) {
					defNameArgCount = defNameCalcArr.length - 1;
					for(var j = 0; j < defNameCalcArr.length; j++) {
						elemArr.push(defNameCalcArr[j]);
					}
				} else {
					elemArr.push(defNameCalcArr);
				}
			} else if (currentElement.type === cElementType.table) {
				elemArr.push(currentElement.toRef(opt_bbox));
			} else if (currentElement.type === cElementType.pivotTable) {
				elemArr.push(currentElement.Calculate(opt_pivotCallback));
			} else if (opt_offset) {
				elemArr.push(this.applyOffset(currentElement, opt_offset));
			} else {
				elemArr.push(currentElement);
			}
		}

		if (promiseCounter && /*!this.promiseResult*/ this.wb.asyncFormulasManager.isPromises()) {
			this.value = new cError(cErrorType.busy);
			this._endCalculate();
			return this.value;
		}

		// ref(CSE) - legacy array-formula
		// dynamic range(DAF) - newest dynamic array-formula
		// Differences:
		// The DAF formula is entered into one cell and is completed by simply pressing enter 'spill' occurs automatically
		// In the case of CSE, you must select the range in advance and press the cse combination after entering the formula
		// The DAF size automatically changes when the data in the original range changes. Dynamic range reference is written in D2# format
		// DAF is edited in the first cell (parent) of the range; to edit cse we need to select the entire previously created range
		// In CSE we cannot delete previously created rows, but DAF can be edited
		// CSE can expand to all cells except the same CSE arrays, tables, pivot tables
		// DAF can only expand to completely empty cells(empty values)

		let isRangeCanFitIntoCells;
		//TODO заглушка для парсинга множественного диапазона в _xlnm.Print_Area. Сюда попадаем только в одном случае - из функции findCell для отображения диапазона области печати
		if(checkMultiSelect && elemArr.length > 1 && this.parent && this.parent instanceof window['AscCommonExcel'].DefName /*&& this.parent.name === "_xlnm.Print_Area"*/) {
			this.value = elemArr;

			if (AscCommonExcel.bIsSupportDynamicArrays) {
				// check further dynamic range
				isRangeCanFitIntoCells =  this.checkDynamicRangeByElement(this.value, opt_bbox);
				if (!isRangeCanFitIntoCells) {
					this.aca = true;
					this.ca = true;
					this.value = new cError(cErrorType.cannot_be_spilled);
				} else {
					this.ca = false;
					this.aca = false;
				}
			}

			this._endCalculate();
		} else {
			let res = elemArr.pop();

			if (this.replaceFormulaAfterCalc === cReplaceFormulaType.formula) {
				if (res && res.type === cElementType.string) {
					if (0 === res.value.indexOf("=")) {
						this.Formula = _tmp.value.slice(1);
						this.isParsed = false;
						this.outStack = [];
						this.unknownOrCustomFunction = null;
						this.parse();
						this.isInDependencies = false;
						this.buildDependencies();
						this.wb.asyncFormulasManager.addReplacedFormula(this);
					} else {
						this.replaceFormulaAfterCalc = null;
					}
				} else {
					this.replaceFormulaAfterCalc = null;
				}
			}

			if (cElementType.error === res.type && res.errorType === cErrorType.busy) {
				this._endCalculate();
				return;
			}
			this.value = res;

			if (AscCommonExcel.bIsSupportDynamicArrays) {
				// check further dynamic range
				isRangeCanFitIntoCells = this.checkDynamicRangeByElement(this.value, opt_bbox);
				if (!isRangeCanFitIntoCells) {
					this.aca = true;
					this.ca = true;
					this.value = new cError(cErrorType.cannot_be_spilled);
				} else {
					this.ca = false;
					this.aca = false;
				}
			}

			this.value.numFormat = numFormat;
			//***array-formula***
			//для обработки формулы массива
			//передаётся последним параметром cell и временно подменяется parent у parserFormula для того, чтобы поменялось значение в элементе массива
			var cell = arguments[3];
			if(this.ref && cell && undefined !== cell.nRow && !(this.ref.r1 === cell.nRow && this.ref.c1 === cell.nCol)) {
				var oldParent = this.parent;
				this.parent = new AscCommonExcel.CCellWithFormula(cell.ws, cell.nRow, cell.nCol);
				this._endCalculate();
				this.parent = oldParent;
			} else {
				//TODO пересмотреть для формул массива, таких как: "=Sheet1'!$S$2:$S$1217"
				/*if(true) {
					var array = this.value.getMatrix()[0];
					var nArray = new cArray();
					nArray.fillFromArray(array);
					this.value = nArray;
				}*/

				this._endCalculate();
			}
			//***array-formula***
		}

		return this.value;
	};
	parserFormula.prototype._endCalculate = function() {
		if (this.parent && this.parent.onFormulaEvent) {
			this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.EndCalculate);
		}
	};

	/* Для обратной сборки функции иногда необходимо поменять ссылки на ячейки */
	parserFormula.prototype.changeOffset = function (offset, canResize, nChangeTable, notOffset3d) {//offset = AscCommon.CellBase
		var t = this;
		//временно комментирую из-за проблемы: при сборке формулы после обработки данной функцией в режиме R1c1
		///мы получаем вид A1. необходимо пересмотреть все функции toString/toLocaleString где возвращается value
		//+ парсинг на endTransaction запускается в режиме r1c1
		//AscCommonExcel.executeInR1C1Mode(false, function () {
			for (var i = 0; i < t.outStack.length; i++) {
				var doOffset = true;
				if (notOffset3d && t.outStack[i] && (t.outStack[i].type === cElementType.cell3D || t.outStack[i].type === cElementType.cellsRange3D)) {
					doOffset = false;
				}
				if (doOffset) {
					t._changeOffsetElem(t.outStack[i], t.outStack, i, offset, canResize, nChangeTable);
				}
			}
		//});
		return this;
	};
	parserFormula.prototype._changeOffsetElem = function(elem, container, index, offset, canResize, nChangeTable) {//offset =
		// AscCommon.CellBase
		var range, bbox = null, ws, isErr = false;
		if (cElementType.cell === elem.type || cElementType.cell3D === elem.type ||
			cElementType.cellsRange === elem.type) {
			isErr = true;
			range = elem.getRange();
			if (range) {
				bbox = range.getBBox0();
				ws = range.getWorksheet();
			}
		} else if (cElementType.cellsRange3D === elem.type) {
			isErr = true;
			bbox = elem.getBBox0NoCheck();
		} else if(cElementType.table === elem.type && !nChangeTable) {
			//когда клонируем диапазон, диапазон таблиц не изменяется
			elem.setOffset(offset);
			elem._updateArea(null, false);
		}

		if (bbox) {
			bbox = bbox.clone();
			if (bbox.setOffsetWithAbs(offset, canResize)) {
				isErr = false;
				this.changeOffsetBBox(elem, bbox, ws);
			}
		}
		if (isErr) {
			container[index] = new cError(cErrorType.bad_reference);
		}
		return elem;
	};
	parserFormula.prototype.applyOffset = function(currentElement, offset) {
		var res = currentElement;
		var cloneElem = null;
		var bbox = null;
		var ws;
		if (cElementType.cell === currentElement.type || cElementType.cell3D === currentElement.type ||
			cElementType.cellsRange === currentElement.type) {
			var range = currentElement.getRange();
			if (range) {
				bbox = range.getBBox0();
				ws = range.getWorksheet();
				if (!bbox.isAbsAll()) {
					cloneElem = currentElement.clone();
					bbox = cloneElem.getRange().getBBox0();
				}
			}
		} else if (cElementType.cellsRange3D === currentElement.type) {
			bbox = currentElement.getBBox0NoCheck();
			if (bbox && !bbox.isAbsAll()) {
				cloneElem = currentElement.clone();
				bbox = cloneElem.getBBox0NoCheck();
			}
		}
		if (cloneElem) {
			bbox.setOffsetWithAbs(offset, false, true);
			this.changeOffsetBBox(cloneElem, bbox, ws);
			res = cloneElem;
		}
		return res;
	};
	parserFormula.prototype.changeOffsetBBox = function(elem, bbox, ws) {
		if (cElementType.cellsRange3D === elem.type) {
			elem.bbox = bbox;
		} else {
			elem.range = AscCommonExcel.Range.prototype.createFromBBox(ws, bbox);
		}
		//todo remove value at all
		elem.value = bbox.getName();
	};
	parserFormula.prototype.changeDefName = function(from, to) {
		var i, elem;
		for (i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if (elem.type === cElementType.name || elem.type === cElementType.name3D || elem.type === cElementType.table) {
				elem.changeDefName(from, to);
			}
		}
	};
	parserFormula.prototype.removeTableName = function(defName, bConvertTableFormulaToRef) {
		var i, elem;
		var bbox;
		if (this.parent && this.parent.onFormulaEvent) {
			bbox= this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.GetRangeCell);
		}

		for (i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if (elem.type === cElementType.table && elem.tableName.toLowerCase() === defName.name.toLowerCase()) {
				if(bConvertTableFormulaToRef)
				{
					this.outStack[i] = this.outStack[i].toRef(bbox, bConvertTableFormulaToRef);
				}
				else
				{
					this.outStack[i] = new cError(cErrorType.bad_reference);
				}
			}
		}
	};
	parserFormula.prototype.removeTableColumn = function(tableName, deleted) {
		var i, elem;
		for (i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if (elem.type === cElementType.table && tableName && elem.tableName.toLowerCase() === tableName.toLowerCase()) {
				if (elem.removeTableColumn(deleted)) {
					this.outStack[i] = new cError(cErrorType.bad_reference);
				}
			}
		}
	};
	parserFormula.prototype.renameTableColumn = function(tableName) {
		var i, elem;
		for (i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if (elem.type === cElementType.table && tableName && elem.tableName.toLowerCase() === tableName.toLowerCase()) {
				if (!elem.renameTableColumn()) {
					this.outStack[i] = new cError(cErrorType.bad_reference);
				}
			}
		}
	};
	parserFormula.prototype.changeTableRef = function(tableName) {
		var i, elem;
		for (i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if (elem.type === cElementType.table && tableName && elem.tableName.toLowerCase() === tableName.toLowerCase()) {
				elem.changeTableRef();
			}
		}
	};
	/**
	 * Shifts the cells on the sheet in accordance with the specified parameters.
	 * 
	 * @param {number} notifyType - The type of notification or action triggering the cell shift.
	 * @param {string} sheetId - The ID of the sheet where the cells are being shifted.
	 * @param {Object} bbox - An object describing the range of the area to be shifted from.
	 * @param {Object} offset - An object describing the offset for shifting the cells.
	 * @param {string} [opt_sheetIdTo] - Optional ID of the sheet to which the cells will be shifted.
	 * @param {boolean} [opt_isPivot] - Optional flag indicating whether the shift is part of working with a pivot table.
	 * @param {boolean} [isTableCreated] - Optional flag indicating whether the table has been created and the shift must be made according to special conditions.
	 * @returns {boolean} Returns true if the cell shift is successful, and false otherwise.
	 */
	parserFormula.prototype.shiftCells = function(notifyType, sheetId, bbox, offset, opt_sheetIdTo, opt_isPivot, isTableCreated) {
		var res = false;
		var elem, bboxCell;
		var wb = this.ws.workbook;
		if (!opt_sheetIdTo) {
			opt_sheetIdTo = sheetId;
		}
		var ws = wb.getWorksheetById(sheetId);
		var wsTo = wb.getWorksheetById(opt_sheetIdTo);
		for (var i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			var _cellsRange = null;
			var _cellsBbox = null;
			if (elem.type === cElementType.cell || elem.type === cElementType.cellsRange) {
				if (sheetId === elem.getWsId() && elem.isValid()) {
					_cellsRange = elem.getRange();
					if (_cellsRange) {
						_cellsBbox = _cellsRange.getBBox0();
					}
				}
			} else if (elem.type === cElementType.cell3D) {
				if (sheetId === elem.getWsId() && elem.isValid()) {
					_cellsRange = elem.getRange();
					if (_cellsRange) {
						_cellsBbox = _cellsRange.getBBox0();
					}
				}
			} else if (elem.type === cElementType.cellsRange3D) {
				if (elem.isSingleSheet() && sheetId === elem.wsFrom.getId() && elem.isValid()) {
					_cellsBbox = elem.getBBox0();
				}
			}
			let tableOffset = null;
			if (_cellsRange || _cellsBbox) {
				var isIntersect;
				if (AscCommon.c_oNotifyType.Shift === notifyType) {
					isIntersect = bbox.isIntersectForShift(_cellsBbox, offset);
				} else if (AscCommon.c_oNotifyType.Move === notifyType) {
					isIntersect = bbox.containsRange(_cellsBbox);
					if (isTableCreated && !isIntersect && bbox.isIntersect(_cellsBbox)) {
						isIntersect = true;
						tableOffset = true;
					}
				} else if (AscCommon.c_oNotifyType.Delete === notifyType) {
					isIntersect = bbox.isIntersect(_cellsBbox);
				}
				if (isIntersect) {
					var isNoDelete;
					if (AscCommon.c_oNotifyType.Shift === notifyType) {
						isNoDelete = _cellsBbox.forShift(bbox, offset, this.wb.bUndoChanges);
					} else if (AscCommon.c_oNotifyType.Move === notifyType) {
						if (tableOffset) {
							// If we select only the first or last cell, then we make a shift by +-1
							if (bbox.r1 === _cellsBbox.r1) {
								_cellsBbox.setOffsetFirst(offset);
							} else if (bbox.r2 === _cellsBbox.r2) {
								_cellsBbox.setOffsetLast(offset);
							} else if (bbox.r1 > _cellsBbox.r1 && bbox.r2 < _cellsBbox.r2) {
							} else {
								// otherwise do shift with forshift method
								_cellsBbox.forShift(bbox, offset, this.wb.bUndoChanges);
							}
						} else {
							_cellsBbox.setOffset(offset);
						}
						isNoDelete = true;
					} else if (AscCommon.c_oNotifyType.Delete === notifyType) {
						if (bbox.containsRange(_cellsBbox)) {
							isNoDelete = false;
						} else {
							isNoDelete = true;
							if (!this.wb.bUndoChanges) {
								var ltIn = bbox.contains(_cellsBbox.c1, _cellsBbox.r1);
								var rtIn = bbox.contains(_cellsBbox.c2, _cellsBbox.r1);
								var lbIn = bbox.contains(_cellsBbox.c1, _cellsBbox.r2);
								var rbIn = bbox.contains(_cellsBbox.c2, _cellsBbox.r2);
								if (ltIn && rtIn && bbox.r1 !== _cellsBbox.r1) {
									_cellsBbox.setOffsetFirst(new AscCommon.CellBase(bbox.r2 - _cellsBbox.r1 + 1, 0));
								} else if (rtIn && rbIn && bbox.c2 !== _cellsBbox.c2) {
									_cellsBbox.setOffsetLast(new AscCommon.CellBase(0, bbox.c1 - _cellsBbox.c2 - 1));
								} else if (rbIn && lbIn && bbox.r2 !== _cellsBbox.r2) {
									_cellsBbox.setOffsetLast(new AscCommon.CellBase(bbox.r1 - _cellsBbox.r2 - 1, 0));
								} else if (lbIn && ltIn && bbox.c1 !== _cellsBbox.c1) {
									_cellsBbox.setOffsetFirst(new AscCommon.CellBase(0, bbox.c2 - _cellsBbox.c1 + 1));
								}
							}
						}
					}
					if (isNoDelete) {
						if (sheetId !== opt_sheetIdTo && (elem.type === cElementType.cell || elem.type === cElementType.cellsRange)) {
							bboxCell = null;
							if (this.parent && this.parent.onFormulaEvent) {
								bboxCell = this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.GetRangeCell);
							}
							if (!bboxCell || !bbox.containsRange(bboxCell)) {
								if (this.wb.bUndoChanges) {
									elem.changeSheet(ws, wsTo);
								} else {
									elem = elem.to3D(wsTo);
									this.outStack[i] = elem;
								}
							}
						}
						if (elem.type === cElementType.cellsRange3D) {
							elem.bbox = _cellsBbox;
							var isDefName;
							if (this.parent && this.parent.onFormulaEvent) {
								isDefName = this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.IsDefName);
							}
							//только если это defName
							if(null === isDefName) {
								elem.changeSheet(ws, wsTo);
							}
						} else {
							elem.range = _cellsRange.createFromBBox(wsTo, _cellsBbox);
						}
						elem.value = _cellsBbox.getName();
					} else if(!opt_isPivot){
						this.outStack[i] = new cError(cErrorType.bad_reference);
					}
					res = true;
				}
			}
		}
		return res;
	};
	parserFormula.prototype.getSharedIntersect = function(sheetId, bbox) {
		var ref;
		var elem;
		var bboxElem;
		for (var i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			bboxElem = undefined;
			if (elem.type === cElementType.cell || elem.type === cElementType.cellsRange ||
				elem.type === cElementType.cell3D) {
				if (sheetId === elem.getWsId() && elem.isValid()) {
					bboxElem = elem.getRange().getBBox0();
				}
			} else if (elem.type === cElementType.cellsRange3D) {
				if (elem.isSingleSheet() && sheetId === elem.wsFrom.getId() && elem.isValid()) {
					bboxElem = elem.getBBox0();
				}
			}
			if (bboxElem) {
				var sharedBBox = bboxElem.getSharedRangeBbox(this.shared.ref, this.shared.base);
				var intersection = bbox.intersection(sharedBBox);
				if (intersection) {
					var bboxSharedRef = sharedBBox.getSharedIntersect(this.shared.ref, intersection);
					ref = ref ? bboxSharedRef.union(ref) : bboxSharedRef;
				}
			}
		}
		return ref;
	};
	parserFormula.prototype.canShiftShared = function(bHor) {
		if (this.shared && this.shared.isOneDimension() && !(bHor ^ this.shared.isHor())) {
			//cut off formulas with absolute reference. it is shifted unexpectedly
			//todo transform base formula
			var elem;
			var bboxElem;
			for (var i = 0; i < this.outStack.length; i++) {
				elem = this.outStack[i];
				bboxElem = undefined;
				if (elem.type === cElementType.cell || elem.type === cElementType.cellsRange ||
					elem.type === cElementType.cell3D) {
					if (elem.isValid()) {
						if (elem.getWS() !== this.getWs()) {
							return false;
						}
						bboxElem = elem.getRange().getBBox0();
					}
				} else if (elem.type === cElementType.cellsRange3D) {
					if (elem.isValid()) {
						if (!(elem.isSingleSheet() && this.getWs() === elem.wsFrom)) {
							return false;
						}
						bboxElem = elem.getBBox0();
					}
				}
				if (bboxElem) {
					if (bHor) {
						if (bboxElem.isAbsC1() || bboxElem.isAbsC2()) {
							return false;
						}
					} else {
						if (bboxElem.isAbsR1() || bboxElem.isAbsR2()) {
							return false;
						}
					}
				}
			}
			return true;
		}
		return false;
	};
	parserFormula.prototype.renameSheetCopy = function (params) {
		var wsLast = params.lastName ? this.wb.getWorksheetByName(params.lastName) : null;
		var wsNew = params.newName ? this.wb.getWorksheetByName(params.newName) : null;
		var isInDependencies = this.isInDependencies;
		if (isInDependencies) {
			//before change outStack necessary to removeDependencies
			this.removeDependencies();
		}

		for (var i = 0; i < this.outStack.length; i++) {
			var elem = this.outStack[i];
			if (params.offset && (cElementType.cell === elem.type || cElementType.cellsRange === elem.type ||
				cElementType.cell3D === elem.type || cElementType.cellsRange3D === elem.type)) {
				elem = this._changeOffsetElem(elem, this.outStack, i, params.offset);
			}
			if (params.tableNameMap && cElementType.table === elem.type) {
				var newTableName = params.tableNameMap[elem.tableName];
				if (newTableName) {
					elem.tableName = newTableName;
				}
			}
			if (wsLast && wsNew) {
				if (cElementType.cell === elem.type || cElementType.cell3D === elem.type ||
					cElementType.cellsRange === elem.type || cElementType.table === elem.type ||
					cElementType.name === elem.type || cElementType.name3D === elem.type) {
					elem.changeSheet(wsLast, wsNew);
				} else if (cElementType.cellsRange3D === elem.type) {
					if (elem.isSingleSheet()) {
						elem.changeSheet(wsLast, wsNew);
					} else {
						if (elem.wsFrom === wsLast || elem.wsTo === wsLast) {
							this.outStack[i] = new cError(cErrorType.bad_reference);
						}
					}
				}
			}
		}
		if (isInDependencies) {
			this.buildDependencies();
		}
		return this;
	};
	parserFormula.prototype.moveToSheet = function (wsLast, wsNew, tableNameMap) {
		var isInDependencies = this.isInDependencies;
		if (isInDependencies) {
			//before change outStack necessary to removeDependencies
			this.removeDependencies();
		}
		if (this.ws === wsLast) {
			this.ws = wsNew;
		}
		for (var i = 0; i < this.outStack.length; i++) {
			var elem = this.outStack[i];
			if (tableNameMap && cElementType.table === elem.type) {
				var newTableName = tableNameMap[elem.tableName];
				if (newTableName) {
					elem.tableName = newTableName;
				}
			}
			if (wsLast && wsNew) {
				if (cElementType.cell === elem.type || cElementType.cellsRange === elem.type ||
					cElementType.table === elem.type ||	cElementType.name === elem.type) {
					elem.changeSheet(wsLast, wsNew);
				}
			}
		}
		if (isInDependencies) {
			this.buildDependencies();
		}
		return this;
	};
	parserFormula.prototype.removeSheet = function (sheetId, tableNamesMap) {
		var bRes = false;
		var ws = this.wb.getWorksheetById(sheetId);
		if (ws) {
			var wsIndex = ws.getIndex();
			var wsPrev = this.wb.getWorksheet(wsIndex - 1);
			var wsNext = this.wb.getWorksheet(wsIndex + 1);
			for (var i = 0; i < this.outStack.length; i++) {
				var elem = this.outStack[i];
				if (cElementType.cellsRange3D === elem.type) {
					if (elem.wsFrom === ws) {
						if (!elem.isSingleSheet() && null !== wsNext) {
							elem.changeSheet(ws, wsNext);
						} else {
							this.outStack[i] = new cError(cErrorType.bad_reference);
						}
						bRes = true;
					} else if (elem.wsTo === ws) {
						if (null !== wsPrev) {
							elem.changeSheet(ws, wsPrev);
						} else {
							this.outStack[i] = new cError(cErrorType.bad_reference);
						}
						bRes = true;
					}
				} else if (cElementType.cell3D === elem.type || cElementType.name3D === elem.type) {
					if (elem.getWS() === ws) {
						this.outStack[i] = new cError(cErrorType.bad_reference);
						bRes = true;
					}
				} else if (cElementType.table === elem.type) {
					if (tableNamesMap[elem.tableName]) {
						this.outStack[i] = new cError(cErrorType.bad_reference);
						bRes = true;
					}
				}
			}
		}
		return bRes;
	};
	parserFormula.prototype.moveSheet = function(tempW) {
		var bRes = false;
		for (var i = 0; i < this.outStack.length; i++) {
			var elem = this.outStack[i];
			if (cElementType.cellsRange3D === elem.type) {
				var wsToIndex = elem.wsTo.getIndex();
				var wsFromIndex = elem.wsFrom.getIndex();
				if (!elem.isSingleSheet()) {
					if (elem.wsFrom === tempW.wF) {
						if (tempW.wTI > wsToIndex) {
							bRes = true;
							var wsNext = this.wb.getWorksheet(wsFromIndex + 1);
							if (wsNext) {
								elem.changeSheet(tempW.wF, wsNext);
							} else {
								this.outStack[i] = new cError(cErrorType.bad_reference);
							}
						}
					} else if (elem.wsTo === tempW.wF) {
						if (tempW.wTI <= wsFromIndex) {
							bRes = true;
							var wsPrev = this.wb.getWorksheet(wsToIndex - 1);
							if (wsPrev) {
								elem.changeSheet(tempW.wF, wsPrev);
							} else {
								this.outStack[i] = new cError(cErrorType.bad_reference);
							}
						}
					}
				}
			}
		}
		return bRes;
	};
	/* Сборка функции в инфиксную форму */
	parserFormula.prototype.assemble = function (rFormula) {
		if (!rFormula && this.outStack.length === 1 && this.outStack[this.outStack.length - 1] instanceof cError) {
			return this.Formula;
		}

		return this._assembleExec();
	};

	/* Сборка функции в инфиксную форму */
	parserFormula.prototype.assembleLocale = function (locale, digitDelim, rFormula) {
		if (!rFormula && this.outStack.length === 1 && this.outStack[this.outStack.length - 1] instanceof cError) {
			return this.Formula;
		}

		return this._assembleExec(locale, digitDelim, true);
	};

	parserFormula.prototype._assembleExec = function (locale, digitDelim, bLocale) {
		//_numberPrevArg - количество аргументов функции в стеке
		var currentElement = null, _count = this.outStack.length, elemArr = new Array(_count), res = undefined,
			_count_arg, _numberPrevArg, _argDiff, onlyRangesElements = true, rangesStr;

		//для получаения грамотного дипапазона, устанавливаем для формул массива g_activeCell главную ячейку
		var formulaArray = this.getArrayFormulaRef();
		var oldActiveCell;
		if(AscCommonExcel.g_R1C1Mode && bLocale && formulaArray){
			AscCommonExcel.g_ActiveCell = new Asc.Range(formulaArray.c1, formulaArray.r1, formulaArray.c1, formulaArray.r1);
			oldActiveCell = AscCommonExcel.g_ActiveCell;
		}

		for (var i = 0, j = 0; i < _count; i++) {
			currentElement = this.outStack[i];

			if(currentElement.type !== cElementType.cellsRange3D && currentElement.type !== cElementType.cell3D && currentElement.type !== cElementType.name && currentElement.type !== cElementType.name3D) {
				onlyRangesElements = false;
				rangesStr = null;
			}

			if (currentElement.type === cElementType.specialFunctionStart || currentElement.type === cElementType.specialFunctionEnd) {
				continue;
			} else if("number" === typeof(currentElement)) {
				j++;
				continue;
			}
			j++;

			if (currentElement.type === cElementType.operator || currentElement.type === cElementType.func) {
				_numberPrevArg = "number" === typeof(this.outStack[i - 1]) ? Math.abs(this.outStack[i - 1]) : null;
				_count_arg = null !== _numberPrevArg ? _numberPrevArg : currentElement.argumentsCurrent;
				_argDiff = 0;
				if(null !== _numberPrevArg) {
					_argDiff++;
					/*if(this.outStack[i - 2] && cElementType.specialFunctionEnd === this.outStack[i - 2].type) {
						_argDiff++;
					}*/
				}

				if(j - _count_arg - _argDiff < 0) {
					continue;
				}

				if (bLocale) {
					res = currentElement.Assemble2Locale(elemArr, j - _count_arg - _argDiff, _count_arg, locale, digitDelim);
				} else {
					res = currentElement.Assemble2(elemArr, j - _count_arg - _argDiff, _count_arg);
				}
				j -= _count_arg + _argDiff;
				elemArr[j] = res;
			} else {
				if (cElementType.string === currentElement.type) {
					if (bLocale) {
						currentElement = new cString("\"" + currentElement.toLocaleString(digitDelim) + "\"");
					} else {
						currentElement = new cString("\"" + currentElement.toString() + "\"");
					}

				}
				res = currentElement;
				elemArr[j] = res;
				if(onlyRangesElements) {
					rangesStr = !rangesStr ? "" : rangesStr + ",";
					rangesStr += bLocale ? res.toLocaleString(digitDelim) : res.toString();
				}
			}
		}

		if (res != undefined && res != null) {
			if(rangesStr) {
				//сделана заглушка для того, чтобы диапазоны разделенные "," собирались грамотно
				//необходимо для того, чтобы мультиселект в именованных диапазонах правильно сохранялся
				//используется в областях печати
				//формулы вида "Sheet1!$B$3:$C$4,Sheet1!$D$3:$E$5,Sheet1!$G$3:$G$6,Sheet1!$J$2"
				//TODO рассмотреть вписание в общую схему
				res = rangesStr;
			} else {
				res = bLocale ? res.toLocaleString(digitDelim) : res.toString();
			}
		} else {
			res = this.Formula;
		}

		if(oldActiveCell) {
			AscCommonExcel.g_ActiveCell = oldActiveCell;
		}
		return res;
	};

	parserFormula.prototype.buildDependencies = function() {
		if (this.isInDependencies) {
			return;
		}
		this.isInDependencies = true;
		var ref, wsR;
		if (this.ca) {
			this.wb.dependencyFormulas.startListeningVolatile(this);
		}

		var isDefName;
		if (this.parent && this.parent.onFormulaEvent) {
			isDefName = this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.IsDefName);
		}

		for (var i = 0; i < this.outStack.length; i++) {
			ref = this.outStack[i];

			if (ref.type === cElementType.table) {
				this.wb.dependencyFormulas.startListeningDefName(ref.tableName, this);
			} else if (ref.type === cElementType.name) {
				this.wb.dependencyFormulas.startListeningDefName(ref.value, this);
			} else if (ref.type === cElementType.name3D) {
				this.wb.dependencyFormulas.startListeningDefName(ref.value, this, ref.ws.getId());
			} else if ((cElementType.cell === ref.type || cElementType.cell3D === ref.type ||
				cElementType.cellsRange === ref.type) && ref.isValid()) {
				this._buildDependenciesRef(ref.getWsId(), ref.getRange() && ref.getRange().getBBox0(), isDefName, true);
			} else if (cElementType.cellsRange3D === ref.type && ref.isValid()) {
				wsR = ref.range(ref.wsRange());
				for (var j = 0; j < wsR.length; j++) {
					var range = wsR[j];
					if (range) {
						this._buildDependenciesRef(range.getWorksheet().getId(), range.getBBox0(), isDefName, true);
					}
				}
			} else if (cElementType.operator === ref.type && ref.name === ":" && this.outStack[i - 1] &&
				this.outStack[i - 2] &&
				((cElementType.cell === this.outStack[i - 1].type && cElementType.cell === this.outStack[i - 2].type) ||
				(cElementType.cell3D === this.outStack[i - 1].type &&
				cElementType.cell3D === this.outStack[i - 2].type)) && this.outStack[i - 1].isValid() &&
				this.outStack[i - 2].isValid()) {
				var _wsId = this.outStack[i - 1].getWsId();
				if (_wsId === this.outStack[i - 2].getWsId()) {
					var _ref1 = this.outStack[i - 2].getRange().getBBox0();
					var _ref2 = this.outStack[i - 1].getRange().getBBox0();
					if (_ref1 && _ref2 && _ref1.c1 <= _ref2.c1 && _ref1.r1 <= _ref2.r1) {
						var _range = new Asc.Range(_ref1.c1, _ref1.r1, _ref2.c1, _ref2.r1);
						this._buildDependenciesRef(_wsId, _range, isDefName, true);
					}
				}
			}
		}

		if (this.importFunctionsRangeLinks) {
			for (let i in this.importFunctionsRangeLinks) {
				let externalLink = this.wb.getExternalLinkByName(i);
				if (externalLink) {
					for (let j in this.importFunctionsRangeLinks[i]) {
						let firstSheet;
						let _rangeInfo = this.importFunctionsRangeLinks[i][j];
						if (!_rangeInfo.sheet) {
							// get first sheet if we haven't sheet name in rangeInfo object
							firstSheet = externalLink.SheetNames && externalLink.SheetNames[0];
						}

						let _ws = externalLink.worksheets[firstSheet ? firstSheet : _rangeInfo.sheet];
						if (_ws) {
							this._buildDependenciesRef(_ws.getId(), AscCommonExcel.g_oRangeCache.getRangesFromSqRef(_rangeInfo.range)[0], null, true);
						}
					}
				}
			}

		}
	};
	parserFormula.prototype.removeDependencies = function() {
		if (!this.isInDependencies) {
			return;
		}
		this.isInDependencies = false;
		var ref;
		var wsR;
		if (this.ca) {
			this.wb.dependencyFormulas.endListeningVolatile(this);
		}

		var isDefName;
		if (this.parent && this.parent.onFormulaEvent) {
			isDefName = this.parent.onFormulaEvent(AscCommon.c_oNotifyParentType.IsDefName);
		}

		for (var i = 0; i < this.outStack.length; i++) {
			ref = this.outStack[i];

			if (ref.type === cElementType.table) {
				this.wb.dependencyFormulas.endListeningDefName(ref.tableName, this);
			} else if (ref.type === cElementType.name) {
				this.wb.dependencyFormulas.endListeningDefName(ref.value, this);
			} else if (ref.type === cElementType.name3D) {
				this.wb.dependencyFormulas.endListeningDefName(ref.value, this, ref.ws.getId());
			} else if ((cElementType.cell === ref.type || cElementType.cell3D === ref.type ||
				cElementType.cellsRange === ref.type) && ref.isValid()) {
				this._buildDependenciesRef(ref.getWsId(), ref.getRange().getBBox0(), isDefName, false);
			} else if (cElementType.cellsRange3D === ref.type && ref.isValid()) {
				wsR = ref.range(ref.wsRange());
				for (var j = 0; j < wsR.length; j++) {
					var range = wsR[j];
					if (range) {
						this._buildDependenciesRef(range.getWorksheet().getId(), range.getBBox0(), isDefName, false);
					}
				}
			} else if (cElementType.operator === ref.type && ref.name === ":" && this.outStack[i - 1] &&
				this.outStack[i - 2] &&
				((cElementType.cell === this.outStack[i - 1].type && cElementType.cell === this.outStack[i - 2].type) ||
				(cElementType.cell3D === this.outStack[i - 1].type &&
				cElementType.cell3D === this.outStack[i - 2].type)) && this.outStack[i - 1].isValid() &&
				this.outStack[i - 2].isValid()) {
				var _wsId = this.outStack[i - 1].getWsId();
				if (_wsId === this.outStack[i - 2].getWsId()) {
					var _ref1 = this.outStack[i - 2].getRange().getBBox0();
					var _ref2 = this.outStack[i - 1].getRange().getBBox0();
					if (_ref1 && _ref2 && _ref1.c1 <= _ref2.c1 && _ref1.r1 <= _ref2.r1) {
						var _range = new Asc.Range(_ref1.c1, _ref1.r1, _ref2.c1, _ref2.r1);
						this._buildDependenciesRef(_wsId, _range, isDefName, false);
					}
				}
			}
		}

		if (this.importFunctionsRangeLinks) {
			for (let i in this.importFunctionsRangeLinks) {
				let externalLink = this.wb.getExternalLinkByName(i);
				if (externalLink) {
					for (let j in this.importFunctionsRangeLinks[i]) {
						let firstSheet;
						let _rangeInfo = this.importFunctionsRangeLinks[i][j];
						if (!_rangeInfo.sheet) {
							// get first sheet if we haven't sheet name in rangeInfo object
							firstSheet = externalLink.SheetNames && externalLink.SheetNames[0];
						}

						let _ws = externalLink.worksheets[firstSheet ? firstSheet : _rangeInfo.sheet];
						if (_ws) {
							this._buildDependenciesRef(_ws.getId(), AscCommonExcel.g_oRangeCache.getRangesFromSqRef(_rangeInfo.range)[0], null, false);
						}
					}
				}
			}
		}
	};
	parserFormula.prototype._buildDependenciesRef = function(wsId, bbox, isDefName, isStart) {
		if (this.isTable) {
			//extend table formula with header/total. This allows us not to follow their change,
			//but sometimes leads to recalculate of the table although changed cells near table (it's not a problem)
			bbox = bbox.clone();
			bbox.setOffsetFirst(new AscCommon.CellBase(-1, 0));
			bbox.setOffsetLast(new AscCommon.CellBase(1, 0));
		}
		if (isDefName) {
			var bboxes = this.extendBBoxCF(isDefName, bbox);
			for (var k = 0; k < bboxes.length; ++k) {
				if (isStart) {
					this.wb.dependencyFormulas.startListeningRange(wsId, bboxes[k], this);
				} else {
					this.wb.dependencyFormulas.endListeningRange(wsId, bboxes[k], this);
				}
			}
		} else {
			bbox = this.extendBBoxDefName(isDefName, bbox);
			if (this.shared) {
				bbox = bbox.getSharedRangeBbox(this.shared.ref, this.shared.base);
			}
			if (isStart) {
				this.wb.dependencyFormulas.startListeningRange(wsId, bbox, this);
			} else {
				this.wb.dependencyFormulas.endListeningRange(wsId, bbox, this);
			}
		}
	};
	parserFormula.prototype.extendBBoxDefName = function(isDefName, bbox) {
		if (null === isDefName && !bbox.isAbsAll()) {
			bbox = bbox.clone();
			if (!bbox.isAbsR1() || !bbox.isAbsR2()) {
				bbox.r1 = 0;
				bbox.r2 = AscCommon.gc_nMaxRow0;
			}
			if (!bbox.isAbsC1() || !bbox.isAbsC2()) {
				bbox.c1 = 0;
				bbox.c2 = AscCommon.gc_nMaxCol0;
			}
		}
		return bbox;
	};
	parserFormula.prototype.extendBBoxCF = function(isDefName, bbox) {
		var res = [];
		if (!bbox.isAbsAll()) {
			var bboxCf = isDefName.bbox;
			var ranges = isDefName.ranges;
			var rowLT = bboxCf ? bboxCf.r1 : 0;
			var colLT = bboxCf ? bboxCf.c1 : 0;
			for (var i = 0; i < ranges.length; ++i) {
				var range = ranges[i];
				var newBBoxLT = bbox.clone();
				newBBoxLT.setOffsetWithAbs(new AscCommon.CellBase(range.r1 - rowLT, range.c1 - colLT), false, true);
				var newBBoxRB = newBBoxLT.clone();
				newBBoxRB.setOffsetWithAbs(new AscCommon.CellBase(range.r2 - range.r1, range.c2 - range.c1), false, true);
				var newBBox = new Asc.Range(newBBoxLT.c1, newBBoxLT.r1, newBBoxRB.c2, newBBoxRB.r2);
				//todo more accurately threshold maxRow/maxCol
				if (!(bbox.r1 <= newBBoxLT.r1 && newBBoxLT.r1 <= newBBoxLT.r2 &&
					newBBoxLT.r1 <= newBBoxRB.r1 && newBBoxRB.r1 <= newBBoxRB.r2)) {
					newBBox.r1 = 0;
					newBBox.r2 = AscCommon.gc_nMaxRow0;
				}
				if (!(bbox.c1 <= newBBoxLT.c1 && newBBoxLT.c1 <= newBBoxLT.c2 &&
					newBBoxLT.c1 <= newBBoxRB.c1 && newBBoxRB.c1 <= newBBoxRB.c2)) {
					newBBox.c1 = 0;
					newBBox.c2 = AscCommon.gc_nMaxCol0;
				}
				res.push(newBBox);
			}
		} else {
			res.push(bbox);
		}
		return res;
	};

	parserFormula.prototype.getFirstRange = function() {
		var res;
		for (var i = 0; i < this.outStack.length; i++) {
			var elem = this.outStack[i];
			if (cElementType.cell === elem.type || cElementType.cell3D === elem.type ||
				cElementType.cellsRange === elem.type || cElementType.cellsRange3D === elem.type
				|| cElementType.table === elem.type) {
				res = elem.getRange();
				break;
			}
		}
		return res;
	};
	parserFormula.prototype.getOutStackSize = function() {
		return this.outStack.length;
	};
	parserFormula.prototype.getOutStackElem = function(index) {
		return this.outStack[index];
	};
	parserFormula.prototype.getIndexNumber = function() {
		return this._index;
	};
	parserFormula.prototype.setIndexNumber = function(val) {
		this._index = val;
	};
	parserFormula.prototype.canSaveShared = function() {
		for (var i = 0; i < this.outStack.length; i++) {
			var elem = this.outStack[i];
			if (cElementType.cell3D === elem.type || cElementType.cellsRange3D === elem.type ||
				cElementType.table === elem.type || cElementType.name3D === elem.type ||
				cElementType.error === elem.type || cElementType.array === elem.type) {
				return false;
			}
		}
		return true;
	};
	parserFormula.prototype.getArrayFormulaRef = function() {
		return this.ref;
	};
	parserFormula.prototype.getDynamicRef = function() {
		if (AscCommonExcel.bIsSupportDynamicArrays) {
			return this.dynamicRange;
		}
	};
	parserFormula.prototype.setArrayFormulaRef = function(ref) {
		this.ref = ref;
	};
	parserFormula.prototype.checkFirstCellArray = function(cell) {
		//возвращаем ТОЛЬКО главную ячейку
		var res = null;
		if(this.ref) {
			if(this.parent && cell.nCol === this.ref.c1 && cell.nRow === this.ref.r1) {
				res = true;
			}
		}
		return res;
	};
	parserFormula.prototype.transpose = function(bounds) {
		for (var i = 0; i < this.outStack.length; i++) {
			//TODO пересмотреть случаи, когда возвращается ошибка
			var elem = this.outStack[i];
			var range;
			if (cElementType.cellsRange === elem.type || cElementType.cell === elem.type || cElementType.cell3D === elem.type) {
				range = elem.range && elem.range.bbox ? elem.range.bbox : null;
			} else if (cElementType.cellsRange3D === elem.type) {
				range = elem.bbox ? elem.bbox : null;
			}
			if (range) {
				var diffCol1 = range.c1 - bounds.c1;
				var diffRow1 = range.r1 - bounds.r1;
				var diffCol2 = range.c2 - bounds.c1;
				var diffRow2 = range.r2 - bounds.r1;

				range.c1 = bounds.c1 + diffRow1;
				range.r1 = bounds.r1 + diffCol1;
				range.c2 = bounds.c1 + diffRow2;
				range.r2 = bounds.r1 + diffCol2;
			}
		}
	};
	parserFormula.prototype.isFoundNestedStAg = function() {
		for (var i = 0; i < this.outStack.length; i++) {
			if (this.outStack[i] && (this.outStack[i].name === "AGGREGATE" || this.outStack[i].name === "SUBTOTAL")) {
				return true;
			}
		}
		return false;
	};
	parserFormula.prototype.simplifyRefType = function (val, opt_ws, opt_row, opt_col) {
		let ref = this.getArrayFormulaRef(), dynamicRef = this.getDynamicRef(), row, col;

		if (val == null) {
			return;
		}

		if (cElementType.cell === val.type || cElementType.cell3D === val.type) {
			val = val.getValue();
			if (cElementType.empty === val.type && opt_ws) {
				// Bug http://bugzilla.onlyoffice.com/show_bug.cgi?id=33941
				val = new cNumber(0);
			}
		} else if (cElementType.array === val.type) {
			if (ref && opt_ws) {
				// TODO check behaviour when row === 1
				row = 1 === val.array.length ? 0 : opt_row - ref.r1;
				col = (val.array[0] && 1 === val.array[0].length) ? 0 : opt_col - ref.c1;
				if (val.array[row] && val.array[row][col]) {
					val = val.getElementRowCol(row, col);
				} else {
					val = new window['AscCommonExcel'].cError(window['AscCommonExcel'].cErrorType.not_available);
				}
			} else {
				val = val.getElement(0);
			}

			//сделано для формул массива
			//внутри массива может лежать ссылка на диапазон(например, функция index возвращает area/ref)
			if (val && (cElementType.cellsRange === val.type || cElementType.cellsRange3D === val.type || cElementType.array === val.type || cElementType.cell === val.type ||
				cElementType.cell3D === val.type)) {
				val = this.simplifyRefType(val, opt_ws, opt_row, opt_col);
			}
		} else if (cElementType.cellsRange === val.type || cElementType.cellsRange3D === val.type) {
			if (opt_ws) {
				let range;
				if (ref) {
					range = val.getRange();
					if (range) {
						let bbox = range.bbox;
						let rowCount = bbox.r2 - bbox.r1 + 1,
							colCount = bbox.c2 - bbox.c1 + 1;

						row = 1 === rowCount ? 0 : opt_row - ref.r1;
						col = 1 === colCount ? 0 : opt_col - ref.c1;
						if (row > rowCount - 1 || col > colCount - 1) {
							val = null;
						} else {
							val = val.getValueByRowCol(row, col);
							if (!val) {
								val = new cEmpty();
							}
						}

						if (!val) {
							val = new window['AscCommonExcel'].cError(window['AscCommonExcel'].cErrorType.not_available);
						}
					} else {
						val = new window['AscCommonExcel'].cError(window['AscCommonExcel'].cErrorType.not_available);
					}
				} else {
					range = new Asc.Range(opt_col, opt_row, opt_col, opt_row);
					val = val.cross(range, opt_ws.getId());
				}
			} else if (cElementType.cellsRange === val.type) {
				val = val.getValue2(0, 0);
			} else {
				val = val.getValue2(new CellAddress(val.getBBox0().r1, val.getBBox0().c1, 0));
			}
		}
		return val;
	};
	parserFormula.prototype.convertTo3DRefs = function (bboxFrom) {
		var elem, bbox;
		for (var i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if (elem.type === cElementType.cell || elem.type === cElementType.cellsRange) {
				bbox = elem.getBBox0();
				if (!bboxFrom.containsRange(bbox)) {
					this.outStack[i] = elem.to3D();
				}
			}
		}
	};
	parserFormula.prototype.hasRelativeRefs = function () {
		var elem;
		for (var i = 0; i < this.outStack.length; i++) {
			elem = this.outStack[i];
			if ((elem.type === cElementType.cell || elem.type === cElementType.cellsRange || elem.type === cElementType.cell3D || elem.type === cElementType.cellsRange3D) &&
				!elem.getBBox0().isAbsAll()) {
				return true;
			}
		}
		return false;
	};

	parserFormula.prototype.getFormulaHyperlink = function () {
		for (var i = 0; i < this.outStack.length; i++) {
			if (this.outStack[i] && (this.outStack[i].name === "HYPERLINK" || this.outStack[i].name === "IMPORTRANGE")) {
				return true;
			}
		}
		return false;
	};

	parserFormula.prototype.setDynamicRef = function (range) {
		if (!range) {
			return
		}
		this.ref = range;
		this.dynamicRange = range;
	};

	parserFormula.prototype.checkDynamicRange = function () {
		/* this function checks if the current value in formula can fit in the cells */
		if (!this.dynamicRange) {
			return true
		}

		if (this.value && (this.value.type !== cElementType.array && this.value.type !== cElementType.cellsRange)) {
			return true
		} else if (this.value && (this.value.type === cElementType.array || this.value.type === cElementType.cellsRange)) {
			// go through the range and see if the array can fit into it
			let dimensions = this.value.getDimensions(),
				mainCell = this.parent, isHaveNonEmptyCell;

			if (this.value.isOneElement()) {
				return true
			}

			if (mainCell) {
				const t = this;
				let rangeRow = mainCell.nRow,
					rangeCol = mainCell.nCol;

				for (let i = rangeRow; i < (rangeRow + dimensions.row); i++) {
					for (let j = rangeCol; j < (rangeCol + dimensions.col); j++) {
						if (i === rangeRow && j === rangeCol) {
							continue
						}
						this.ws._getCellNoEmpty(i, j, function(cell) {
							if (cell) {
								let formula = cell.getFormulaParsed();
								let dynamicRangeFromCell = formula && formula.getDynamicRef();
								if (formula && dynamicRangeFromCell) {
									// check if cell belong to current dynamicRange
									// this is necessary so that spill errors do not occur during the second check of the range (since the values ​​in it have already been entered earlier)
									if (!t.dynamicRange.isEqual(dynamicRangeFromCell)) {
										// if the cell is part of another dynamic range, then the range that is in the area of ​​the previous range is displayed (except for the first cell, but we do not check it)
										// that is, if one of the ranges is “lower” or “to the right” in the editor, then it will be displayed, and the other will receive a SPILL error
										isHaveNonEmptyCell = true
									}
								} else if (cell.formulaParsed || !cell.isEmptyTextString()) {
									isHaveNonEmptyCell = true
								}
							}
						});
						if (isHaveNonEmptyCell) {
							return false
						}
					}
				}
				return true
			}
		}

		return false
	};

	parserFormula.prototype.checkDynamicRangeByElement = function (element, parentCell) {
		/* this function checks if element can fit in the cells */
		if (!element || !parentCell) {
			return true;
		}

		if (element.type !== cElementType.array && element.type !== cElementType.cellsRange && element.type !== cElementType.cellsRange3D) {
			return true;
		} else if (element.type === cElementType.array || element.type === cElementType.cellsRange || element.type === cElementType.cellsRange3D) {
			// go through the range and see if the array can fit into it
			let dimensions = element.getDimensions(), isHaveNonEmptyCell;

			if (element.isOneElement()) {
				return true
			}

			// todo if an element is defname, it has no parent element?
			const t = this;
			let rangeRow = parentCell.r1,
				rangeCol = parentCell.c1;

			let supposedDynamicRange = this.ws.getRange3(rangeRow, rangeCol, (rangeRow + dimensions.row) - 1, (rangeCol + dimensions.col) - 1);
			for (let i = rangeRow; i < (rangeRow + dimensions.row); i++) {
				for (let j = rangeCol; j < (rangeCol + dimensions.col); j++) {
					if (i === rangeRow && j === rangeCol) {
						continue
					}
					this.ws._getCellNoEmpty(i, j, function(cell) {
						if (cell) {
							let formula = cell.getFormulaParsed();
							let dynamicRangeFromCell = formula && formula.getDynamicRef();
							if (formula && dynamicRangeFromCell) {
								// check if cell belong to current dynamicRange
								// this is necessary so that spill errors do not occur during the second check of the range (since the values ​​in it have already been entered earlier)
								if (!supposedDynamicRange.bbox.isEqual(dynamicRangeFromCell)) {
									// if the cell is part of another dynamic range, then the range that is in the area of ​​the previous range is displayed (except for the first cell, but we do not check it)
									// that is, if one of the ranges is “lower” or “to the right” in the editor, then it will be displayed, and the other will receive a SPILL error
									isHaveNonEmptyCell = true
								}
							} else if (cell.formulaParsed || !cell.isEmptyTextString()) {
								isHaveNonEmptyCell = true
							}
						}
					});
					if (isHaveNonEmptyCell) {
						return false
					}
				}
			}
			return true
		}

		return false
	};

	/**
	 * Class representative an iterative calculations logic
	 * @constructor
	 */
	function CalcRecursion() {
		this.nLevel = 0;
		this.bIsForceBacktracking = false;
		this.bIsProcessRecursion = false;
		this.aElems = [];
		this.aElemsPart = [];

		this.nIterStep = 1;
		this.oStartCellIndex = null;
		this.nRecursionCounter = 0;
		this.oGroupChangedCells = null;
		this.oPrevIterResult = null;
		this.oDiffBetweenIter = null;
		this.bShowCycleWarn = true;
		this.oRecursionCells = null;
		this.nCellPasteValue = null; // for paste recursive cell
		this.bIsCellEdited = false;
		this.bIsSheetCreating = false;
		this.oIndirectFuncResult = null;
		this.oOffsetFuncResult = null;
		this.oCellContentFuncRes = null;
		this.aCycleCell = [];

		this.bIsEnabledRecursion = null;
		this.nMaxIterations = null; // Max iterations of recursion calculations. Default value: 100.
		this.nRelativeError = null; // Relative error between current and previous cell value. Default value: 1e-3.
		this.nCalcMode = Asc.c_oAscCalcMode.auto; // Calculation mode. Default value: Asc.c_oAscCalcMode.auto
		/*for chrome63(real maximum call stack size is 12575) nMaxRecursion that cause exception is 783
		by measurement: stack size in doctrenderer is one fourth smaller than chrome*/
		this.nMaxRecursion = 300; // Default value: 300
	}

	/**
	 * Method returns maximum recursion level.
	 * @memberof CalcRecursion
	 * @returns {number}
	 */
	CalcRecursion.prototype.getMaxRecursion = function () {
		return this.nMaxRecursion;
	};
	/**
	 * Method sets a flag who recognizes recursion needs force backtracking.
	 * Uses if level of recursion exceeds max level.
	 * @memberof CalcRecursion
	 * @param {boolean} bIsForceBacktracking
	 */
	CalcRecursion.prototype.setIsForceBacktracking = function (bIsForceBacktracking) {
		if (!this.getIsForceBacktracking()) {
			this.aElemsPart = [];
			this.aElems.push(this.aElemsPart);
		}
		this.bIsForceBacktracking = bIsForceBacktracking;
	};
	/**
	 * Method returns a flag who recognizes recursion needs force backtracking.
	 * Uses if level of recursion exceeds max level.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.getIsForceBacktracking = function () {
		return this.bIsForceBacktracking;
	};
	/**
	 * Method sets a flag who recognizes work with aElems in _checkDirty method is already in process.
	 * @memberof CalcRecursion
	 * @param {boolean} bIsProcessRecursion
	 */
	CalcRecursion.prototype.setIsProcessRecursion = function (bIsProcessRecursion) {
		this.bIsProcessRecursion = bIsProcessRecursion;
	};
	/**
	 * Method returns a flag who recognizes work with aElems in _checkDirty method is already in process.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.getIsProcessRecursion = function () {
		return this.bIsProcessRecursion;
	}
	/**
	 * Method increases recursion level. Uses for tracking a level of recursion in _checkDirty method.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.incLevel = function () {
		this.nLevel++;
	};
	/**
	 * Method decreases recursion level. Uses for actualizes a level of recursion
	 * in case when one of recursion is finished. Uses in _checkDirty method.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.decLevel = function () {
		this.nLevel--;
	};
	/**
	 * Method returns level of recursion in _checkDirty method.
	 * @memberof CalcRecursion
	 * @returns {number}
	 */
	CalcRecursion.prototype.getLevel = function () {
		return this.nLevel;
	};
	/**
	 * Method checks the level of recursion exceeds max level or not.
	 * Uses in _checkDirty method.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.checkLevel = function () {
		if (this.getIsForceBacktracking()) {
			return false;
		}

		let res = this.getLevel() <= this.getMaxRecursion();
		if (!res) {
			this.setIsForceBacktracking(true);
		}

		return res;
	};
	/**
	 * Method inserts cells which need to be processed in _checkDirty method again.
	 * Uses for formula chains that reached max recursion level.
	 * @memberof CalcRecursion
	 * @param {{ws:Worksheet, nRow:number, nCol:number}} oCellCoordinate
	 */
	CalcRecursion.prototype.insert = function (oCellCoordinate) {
		this.aElemsPart.push(oCellCoordinate);
	};
	/**
	 * Method executes callback for each cell from aElems in reverse order.
	 * aElems stores cell coordinates which need to be processed in _checkDirty method again.
	 * @memberof CalcRecursion
	 * @param {Function} fCallback
	 */
	CalcRecursion.prototype.foreachInReverse = function (fCallback) {
		for (let i = this.aElems.length - 1; i >= 0; i--) {
			let aElemsPart = this.aElems[i];
			for (let j = 0, length = aElemsPart.length; j < length; j++) {
				fCallback(aElemsPart[j]);
				if (this.getIsForceBacktracking()) {
					return;
				}
			}
		}
	};
	/**
	 * Method increases iteration step.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.incIterStep = function () {
		this.nIterStep++;
	};
	/**
	 * Method resets iteration step.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.resetIterStep = function () {
		this.nIterStep = 1;
	};
	/**
	 * Method returns iteration step.
	 * @memberof CalcRecursion
	 * @returns {number}
	 */
	CalcRecursion.prototype.getIterStep = function () {
		return this.nIterStep;
	};
	/**
	 * Method increments recursion counter.
	 * Uses for control recursion level of initStartCellForIterCalc and enableCalcFormulas method of Cell class.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.incRecursionCounter = function () {
		this.nRecursionCounter++;
	};
	/**
	 * Method decrements recursion counter.
	 * Uses for control recursion level of initStartCellForIterCalc and enableCalcFormulas method of Cell class.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.decRecursionCounter = function () {
		this.nRecursionCounter--;
	};
	/**
	 * Method resets recursion counter.
	 * Uses for control recursion level of initStartCellForIterCalc method.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.resetRecursionCounter = function () {
		if (this.getRecursionCounter() > 0) {
			this.decRecursionCounter();
		} else if (this.getIsForceBacktracking()) {
			this.setIsForceBacktracking(false);
		}
	};
	/**
	 * Method checks the recursion counter exceeds max level of recursion or not.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.checkRecursionCounter = function () {
		if (this.getIsForceBacktracking()) {
			return true;
		}

		let bRecursionExceeded = g_cCalcRecursion.getRecursionCounter() >= g_cCalcRecursion.getMaxRecursion();

		if (bRecursionExceeded) {
			this.setIsForceBacktracking(true);
		}

		return bRecursionExceeded;
	}
	/**
	 * Method returns recursion counter.
	 * Uses for control recursion level of initStartCellForIterCalc method.
	 * @memberof CalcRecursion
	 * @returns {number}
	 */
	CalcRecursion.prototype.getRecursionCounter = function () {
		return this.nRecursionCounter;
	};
	/**
	 * Method sets a flag who recognizes an iteration calculations setting is enabled or not.
	 * @memberof CalcRecursion
	 * @param {boolean} bIsEnabledRecursion
	 */
	CalcRecursion.prototype.setIsEnabledRecursion = function (bIsEnabledRecursion) {
		this.bIsEnabledRecursion = bIsEnabledRecursion;
	};
	/**
	 * Method returns a flag who recognizes an iteration calculations setting is enabled or not.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.getIsEnabledRecursion = function () {
		return this.bIsEnabledRecursion;
	};
	/**
	 * Method sets index of start cell. This cell is a start and finish point of iteration for a recursion formula.
	 * Uses for only with enabled iterative calculations setting.
	 * @memberof CalcRecursion
	 * @param {{cellId: number, wsName: string}|null} oStartCellIndex
	 */
	CalcRecursion.prototype.setStartCellIndex = function (oStartCellIndex) {
		this.oStartCellIndex = oStartCellIndex;
	};
	/**
	 * Method returns index of start cell. This cell is a start and finish point of iteration for a recursion formula.
	 * Uses for only with enabled iterative calculations setting.
	 * @memberof CalcRecursion
	 * @returns {{cellId: number, wsName: string}}
	 */
	CalcRecursion.prototype.getStartCellIndex = function () {
		return this.oStartCellIndex;
	};
	/**
	 * Method sets a maximum iterations.
	 * @memberof CalcRecursion
	 * @param {number} nMaxIterations
	 */
	CalcRecursion.prototype.setMaxIterations = function (nMaxIterations) {
		this.nMaxIterations = nMaxIterations;
	};
	/**
	 * Method returns a maximum iterations.
	 * @memberof CalcRecursion
	 * @returns {number}
	 */
	CalcRecursion.prototype.getMaxIterations = function () {
		return this.nMaxIterations;
	};
	/**
	 * Method sets a relative error.
	 * @memberof CalcRecursion
	 * @param {number} nRelativeError
	 */
	CalcRecursion.prototype.setRelativeError = function (nRelativeError) {
		this.nRelativeError = nRelativeError;
	};
	/**
	 * Method returns a relative error.
	 * @memberof CalcRecursion
	 * @returns {number}
	 */
	CalcRecursion.prototype.getRelativeError = function () {
		return this.nRelativeError;
	};
	/**
	 * Method sets a calculation mode.
	 * @memberof CalcRecursion
	 * @param {Asc.c_oAscCalcMode} nCalcMode
	 */
	CalcRecursion.prototype.setCalcMode = function (nCalcMode) {
		this.nCalcMode = nCalcMode;
	};
	/**
	 * Method returns a calculation mode.
	 * @memberof CalcRecursion
	 * @returns {Asc.c_oAscCalcMode}
	 */
	CalcRecursion.prototype.getCalcMode = function () {
		return this.nCalcMode;
	};
	/**
	 * Method sets a grouped changed cells.
	 * @memberof CalcRecursion
	 * @param {{wsName:{cellId: {cellId: number, wsName: string}[]}}|null} oGroupChangedCells
	 */
	CalcRecursion.prototype.setGroupChangedCells = function (oGroupChangedCells) {
		this.oGroupChangedCells = oGroupChangedCells;
	};
	/**
	 * Method returns a grouped changed cells.
	 * @memberof CalcRecursion
	 * @returns {{wsName:{cellId: {cellId: number, wsName: string}[]}}|null}
	 */
	CalcRecursion.prototype.getGroupChangedCells = function () {
		return this.oGroupChangedCells;
	};
	/**
	 * Method initializes an object for grouped changed cells.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 */
	CalcRecursion.prototype.initGroupChangedCells = function (oCell) {
		const sCellWsName = oCell.ws.getName().toLowerCase();
		let oGroupChangedCell = {};
		oGroupChangedCell[sCellWsName] = {};
		this.setGroupChangedCells(oGroupChangedCell);
	};
	/**
	 * Method returns an array of cells with recursive formula.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 * @returns {{cellId: number, wsName: string}[]}
	 */
	CalcRecursion.prototype.getRecursiveCells = function (oCell) {
		const oGroupChangedCell = this.getGroupChangedCells();
		const sCellWsName = oCell.ws.getName().toLowerCase();
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);

		if (oGroupChangedCell == null) {
			return [];
		}
		for (let sSheetName in oGroupChangedCell) {
			let oGroupChangedSheet = oGroupChangedCell[sSheetName];
			for (let sLinkedCellIndex in oGroupChangedSheet) {
				const aLinkedCells = oGroupChangedSheet[sLinkedCellIndex];
				let bHasCell = aLinkedCells.some(function (oCellIndex) {
					return oCellIndex.cellId === nCellIndex && oCellIndex.wsName === sCellWsName;
				})
				if (bHasCell) {
					return aLinkedCells;
				}
			}
		}

		return [];
	};
	/**
	 * Method checks a cell has in array with recursive cells.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.hasInRecursiveCells = function (oCell) {
		const oThis = this;
		const aRecursiveCell = this.getRecursiveCells(oCell)
		let bHasListeners = !!oCell.getListeners();
		let bHasInRecursiveCells = !!aRecursiveCell.length;

		if (!bHasListeners) {
			const aCellTypes = [cElementType.cell, cElementType.cell3D];
			let aOutStack = oCell.getFormulaParsed().outStack;
			for (let i = 0; i < aOutStack.length; i++) {
				if (aOutStack.type === cElementType.name && aOutStack.type === cElementType.name3D) {
					aOutStack[i] = aOutStack[i].toRef();
				}
				if (aCellTypes.includes(aOutStack[i].type)) {
					let oRange = aOutStack[i].getRange();
					oRange._foreachNoEmpty(function (oElem) {
						if (oElem.isFormula()) {
							bHasInRecursiveCells = !!oThis.getRecursiveCells(oElem).length;
							bHasListeners = !!oElem.getListeners();
						}
					});
				}
			}
		}

		return bHasInRecursiveCells || bHasListeners;
	};
	/**
	 * Method updates start cell index.
	 * @memberof CalcRecursion
	 * @param {{cellId: number, wsName: string}[]} aRecursiveCells
	 */
	CalcRecursion.prototype.updateStartCellIndex = function (aRecursiveCells) {
		if (!aRecursiveCells.length) {
			this.setStartCellIndex(null);
			return;
		}

		const START_CELL_INDEX = 0;
		const oStartCellIdFromArr = aRecursiveCells[START_CELL_INDEX];
		const oStartCellIndex = this.getStartCellIndex();
		if (oStartCellIndex && oStartCellIdFromArr.cellId === oStartCellIndex.cellId && oStartCellIdFromArr.wsName === oStartCellIndex.wsName) {
			return;
		}
		this.setStartCellIndex(oStartCellIdFromArr);
	};
	/**
	 * Method adds array with recursive cells in the group changed cells object.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 * @param {{cellId: number, wsName: string}[]} aRecursiveCells
	 */
	CalcRecursion.prototype.addRecursiveCells = function (oCell, aRecursiveCells) {
		const sCellWsName = oCell.ws.getName().toLowerCase();
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);
		let oGroupChangedCell = this.getGroupChangedCells();

		if (oGroupChangedCell == null) {
			this.initGroupChangedCells(oCell);
			oGroupChangedCell = this.getGroupChangedCells();
		}
		if (!oGroupChangedCell.hasOwnProperty(sCellWsName)) {
			oGroupChangedCell[sCellWsName] = {};
		}
		oGroupChangedCell[sCellWsName][nCellIndex] = aRecursiveCells;
	};
	/**
	 * Method updates array with recursive cells in the group changed cells object.
	 * @memberof CalcRecursion
	 * @param {{cellId: number, wsName: string}} oCellIndex
	 * @param {{cellId: number, wsName: string}[]} aRecursiveCells
	 */
	CalcRecursion.prototype.updateRecursiveCells = function (oCellIndex, aRecursiveCells) {
		const oGroupChangedCell = this.getGroupChangedCells();
		const sCellWsName = oCellIndex.wsName;
		const nCellIndex = oCellIndex.cellId;

		oGroupChangedCell[sCellWsName][nCellIndex] = aRecursiveCells;
	};
	/**
	 * Method removes array with recursive cells in the group changed cells object.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 */
	CalcRecursion.prototype.removeRecursionCell = function (oCell) {
		const oGroupChangedCell = this.getGroupChangedCells();
		const sCellWsName = oCell.ws.getName().toLowerCase();
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);
		if (!oGroupChangedCell[sCellWsName] || !oGroupChangedCell[sCellWsName][nCellIndex]) {
			return;
		}
		delete oGroupChangedCell[sCellWsName][nCellIndex];
	};
	/**
	 * Method sets a previous iteration result.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 */
	CalcRecursion.prototype.setPrevIterResult = function (oCell) {
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);
		const sWsName = oCell.ws.getName().toLowerCase();

		if (this.oPrevIterResult == null) {
			this.oPrevIterResult = {};
		}
		if (!this.oPrevIterResult.hasOwnProperty(sWsName)) {
			this.oPrevIterResult[sWsName] = {};
		}
		this.oPrevIterResult[sWsName][nCellIndex] = oCell.getNumberValue();
	};
	/**
	 * Method returns a previous iteration result.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 * @returns {number}
	 */
	CalcRecursion.prototype.getPrevIterResult = function (oCell) {
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);
		const sWsName = oCell.ws.getName().toLowerCase();
		const oPrevIterResult = this.oPrevIterResult;

		if (oPrevIterResult == null) {
			return NaN;
		}
		if (!oPrevIterResult.hasOwnProperty(sWsName) || !oPrevIterResult[sWsName].hasOwnProperty(nCellIndex)) {
			return NaN;
		}

		return oPrevIterResult[sWsName][nCellIndex];
	};
	/**
	 * Method clears a previous iteration results.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.clearPrevIterResult = function () {
		this.oPrevIterResult = null;
	};
	/**
	 * Method sets result of a difference between iterations.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 * @param {number} nResult
	 */
	CalcRecursion.prototype.setDiffBetweenIter = function (oCell, nResult) {
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);
		const sWsName = oCell.ws.getName().toLowerCase();

		if (this.oDiffBetweenIter == null) {
			this.oDiffBetweenIter = {};
		}
		if (!this.oDiffBetweenIter.hasOwnProperty(sWsName)) {
			this.oDiffBetweenIter[sWsName] = {};
		}
		this.oDiffBetweenIter[sWsName][nCellIndex] = nResult;
	}
	/**
	 * Method calculates a result of a difference between iterations.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 */
	CalcRecursion.prototype.calcDiffBetweenIter = function (oCell) {
		const nPrevIterResult = this.getPrevIterResult(oCell);
		const nCurrentIterResult = oCell.getNumberValue();
		const nChainLength = this.getRecursiveCells(oCell).length;

		if (this.getIterStep() <= nChainLength && nCurrentIterResult === nPrevIterResult) {
			return;
		}
		this.setDiffBetweenIter(oCell, Math.abs(nCurrentIterResult - nPrevIterResult));
	};
	/**
	 * Method returns a result of a difference between iterations.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 * @returns {number}
	 */
	CalcRecursion.prototype.getDiffBetweenIter = function (oCell) {
		const nCellIndex = AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol);
		const sWsName = oCell.ws.getName().toLowerCase();
		const oDiffBetweenIter = this.oDiffBetweenIter;

		if (oDiffBetweenIter == null) {
			return NaN;
		}
		if (!oDiffBetweenIter.hasOwnProperty(sWsName) || !oDiffBetweenIter[sWsName].hasOwnProperty(nCellIndex)) {
			return NaN;
		}

		return oDiffBetweenIter[sWsName][nCellIndex];
	};
	/**
	 * Method clears a result of a difference between iterations.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.clearDiffBetweenIter = function () {
		this.oDiffBetweenIter = null;
	};
	/**
	 * Method returns a flag that checks a recursive call is needed.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.needRecursiveCall = function () {
		if (!this.getIsEnabledRecursion()) {
			return false;
		}
		const oGroupChangedCells = this.getGroupChangedCells();
		const bMaxStepNotExceeded = this.getIterStep() <= this.getMaxIterations();
		let bHasRecursiveCell = false;

		for (let sSheetName in oGroupChangedCells) {
			let oGroupChangedSheet = oGroupChangedCells[sSheetName];
			for (let sCellIndex in oGroupChangedSheet) {
				let aRecursiveCells = oGroupChangedSheet[sCellIndex];
				if (aRecursiveCells.length) {
					bHasRecursiveCell = true;
					break;
				}
			}
			if (bHasRecursiveCell) {
				break;
			}
		}

		return bHasRecursiveCell && bMaxStepNotExceeded;
	};
	/**
	 * Method initializes calculation properties.
	 * @memberof CalcRecursion
	 * @param {CCalcPr} oCalcPr
	 */
	CalcRecursion.prototype.initCalcProperties = function (oCalcPr) {
		const oCalcSettings = Asc.editor.asc_GetCalcSettings(); // Object with default values

		if (!oCalcSettings) {
			return;
		}
		this.setIsEnabledRecursion(oCalcPr.getIterate() ? oCalcPr.getIterate() : oCalcSettings.asc_getIterativeCalc());
		this.setRelativeError(oCalcPr.getIterateDelta() != null ? oCalcPr.getIterateDelta() : oCalcSettings.asc_getMaxChange());
		this.setMaxIterations(oCalcPr.getIterateCount() ? oCalcPr.getIterateCount() : oCalcSettings.asc_getMaxIterations());
		if (oCalcPr.getCalcMode()) {
			this.setCalcMode(oCalcPr.getCalcMode());
		}
	};
	/**
	 * Method returns a flag who recognizes show warn about cycle reference error or not.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.getShowCycleWarn = function () {
		return this.bShowCycleWarn;
	};
	/**
	 * Method sets a flag who recognizes show warn about cycle reference error or not.
	 * @memberof CalcRecursion
	 * @param {boolean} bShowCycleWarn
	 */
	CalcRecursion.prototype.setShowCycleWarn = function (bShowCycleWarn) {
		this.bShowCycleWarn = bShowCycleWarn;
	};
	/**
	 * Method adds an index of a recursive cell to the list of recursive cells.
	 * @memberof CalcRecursion
	 * @param {number} nCellIndex
	 */
	CalcRecursion.prototype.addRecursiveCell = function (nCellIndex) {
		if (!this.oRecursionCells) {
			this.oRecursionCells = {};
		}

		this.oRecursionCells[nCellIndex] = true;
	};
	/**
	 * Method checks a cell is recursive or not by index of cell.
	 * @memberof CalcRecursion
	 * @param {number} nCellId
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.isRecursiveCell = function (nCellId) {
		return !!(this.oRecursionCells && this.oRecursionCells[nCellId]);
	};
	/**
	 * Method clears a list of recursive cells.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.clearRecursionCells = function () {
		this.oRecursionCells = null;
	};
	/**
	 * Method finds recursive cell by parserFormula.
	 * @memberof CalcRecursion
	 * @param {parserFormula} oParserFormula
	 */
	CalcRecursion.prototype.findRecursionCell = function (oParserFormula) {
		const oThis = this;
		const oParentCell = oParserFormula.getParent();

		if (!(oParentCell instanceof AscCommonExcel.CCellWithFormula)) {
			return;
		}
		oParserFormula.ws._getCell(oParentCell.nRow, oParentCell.nCol, function (oCell) {
			if (oCell.isFormula()) {
				oCell.initStartCellForIterCalc(); // check cell has recursion formula
				if (oThis.getStartCellIndex()) {
					oThis.addRecursiveCell(AscCommonExcel.getCellIndex(oCell.nRow, oCell.nCol));
					oThis.setStartCellIndex(null);
				}
			}
		});
	};
	/**
	 * Method sets a value from a copying cell for a paste cell.
	 * @memberof CalcRecursion
	 * @param {number|null} nCellPasteValue
	 */
	CalcRecursion.prototype.setCellPasteValue = function (nCellPasteValue) {
		this.nCellPasteValue = nCellPasteValue;
	};
	/**
	 * Method gets a value from a copying cell for a paste cell.
	 * @memberof CalcRecursion
	 * @returns {number|null}
	 */
	CalcRecursion.prototype.getCellPasteValue = function () {
		return this.nCellPasteValue;
	};
	/**
	 * Method sets flag that checks cell is in edited mode
	 * * true - cell is editing. File in the editor already opened.
	 * * false - cell isn't editing. File in the editor is opening.
	 * @param {boolean} bIsCellEdited
	 */
	CalcRecursion.prototype.setIsCellEdited = function (bIsCellEdited) {
		this.bIsCellEdited = bIsCellEdited;
	};
	/**
	 * Method gets flag that checks cell is in edited mode
	 * * true - cell is editing. File in the editor already opened.
	 * * false - cell isn't editing. File in the editor is opening.
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.getIsCellEdited = function () {
		return this.bIsCellEdited;
	};
	/**
	 * Method gets the flag that checks whether the editor is making an "Add sheet" operation.
	 * * true - Editor is making an "Add sheet" operation.
	 * * false - Editor isn't making an "Add sheet" operation.
	 * @memberof CalcRecursion
	 * @returns {boolean}
	 */
	CalcRecursion.prototype.getIsSheetCreating = function () {
		return this.bIsSheetCreating;
	};
	/**
	 * Method sets the flag that checks whether the editor is making an "Add sheet" operation.
	 * * true - Editor is making an "Add sheet" operation.
	 * * false - Editor isn't making an "Add sheet" operation.
	 * @memberof CalcRecursion
	 * @param {boolean} bIsSheetCreating
	 */
	CalcRecursion.prototype.setIsSheetCreating = function (bIsSheetCreating) {
		this.bIsSheetCreating = bIsSheetCreating;
	};
	/**
	 * Method saves the result of the formula, which needs to be checked for cycle after being calculated.
	 * @memberof CalcRecursion
	 * @param {string}sFuncName
	 * @param {cRef|cRef3D|cArea|cArea3D|cName|cName3D}oResult
	 */
	CalcRecursion.prototype.saveFunctionResult = function (sFuncName, oResult) {
		if (oResult.type === cElementType.error) {
			return;
		}
		switch (sFuncName) {
			case 'INDIRECT':
				this.oIndirectFuncResult = oResult;
				break;
			case 'OFFSET':
				this.oOffsetFuncResult = oResult;
				break;
			case 'CELL':
				this.oCellContentFuncRes = oResult;
				break;
		}
	};
	/**
	 * Method returns the array of result of functions, which need to be checked for cycle after being calculated.
	 * @memberof CalcRecursion
	 * @returns {[]}
	 */
	CalcRecursion.prototype.getFunctionsResult = function () {
		const aFunctionResults = [];

		if (this.oIndirectFuncResult) {
			aFunctionResults.push(this.oIndirectFuncResult);
		}
		if (this.oOffsetFuncResult) {
			aFunctionResults.push(this.oOffsetFuncResult);
		}
		if (this.oCellContentFuncRes) {
			aFunctionResults.push(this.oCellContentFuncRes);
		}

		return aFunctionResults;
	};
	/**
	 * Method clears result of formulas, which need to be checked for cycle after being calculated.
	 *  @memberof CalcRecursion
	 */
	CalcRecursion.prototype.clearFunctionsResult = function () {
		this.oIndirectFuncResult = null;
		this.oOffsetFuncResult = null;
		this.oCellContentFuncRes = null;
	};
	/**
	 * Method returns array of cycle cells.
	 * Uses when "Iteration calculation" setting is disabled.
	 * @memberof CalcRecursion
	 * @returns {Cell[]}
	 */
	CalcRecursion.prototype.getCycleCells = function () {
		return this.aCycleCell;
	};
	/**
	 * Method adds cycle cell to array.
	 * Uses when "Iteration calculation" setting is disabled.
	 * @memberof CalcRecursion
	 * @param {Cell} oCell
	 */
	CalcRecursion.prototype.addCycleCell = function (oCell) {
		let bDuplicateElem = this.aCycleCell.some(function (oElem) {
			return oElem.nRow === oCell.nRow && oElem.nCol === oCell.nCol && oElem.ws.getName() === oCell.ws.getName();
		});
		if (bDuplicateElem) {
			return;
		}
		this.aCycleCell.push(oCell);
	};
	/**
	 * Method clears array of cycle cells.
	 * @memberof CalcRecursion
	 */
	CalcRecursion.prototype.clearCycleCells = function () {
		this.aCycleCell = [];
	};

	const g_cCalcRecursion = new CalcRecursion();

	function parseNum(str) {
		if (str.indexOf("x") > -1 || str == "" || str.match(/^\s+$/))//исключаем запись числа в 16-ричной форме из числа.
		{
			return false;
		}
		return !isNaN(str);
	}

	var matchingOperators = new RegExp("^(=|<>|<=|>=|<|>).*");

	function matchingValue(oVal) {
		var res;
		if (cElementType.string === oVal.type) {
			var search, op;
			var val = oVal.getValue();
			var match = val.match(matchingOperators);
			if (match) {
				search = val.substr(match[1].length);
				op = match[1].replace(/\s/g, "");
			} else {
				search = val;
				op = null;
			}

			var parseRes = AscCommon.g_oFormatParser.parse(search);
			res = {val: parseRes ? new cNumber(parseRes.value) : new cString(search), op: op};
		} else {
			res = {val: oVal, op: null};
		}

		return res;
	}

	function matching(x, matchingInfo, doNotParseNum, doNotParseFormat) {
		var y = matchingInfo.val;
		var operator = matchingInfo.op;
		var res = false, rS;
		if (cElementType.string === y.type) {
			if ('<' === operator || '>' === operator || '<=' === operator || '>=' === operator) {
				var _funcVal = _func[x.type][y.type](x, y, operator);
				if (cElementType.error === _funcVal.type) {
					return false;
				}
				return _funcVal.toBool();
			}

			y = y.toString();

			// Equal only string values
			if(cElementType.empty === x.type && '' === y){
				rS = true;
			} else if(cElementType.bool === x.type){
				x = x.tocString();
				rS = x.value === y;
			}else if(cElementType.error === x.type){
				rS = x.value === y;
			}else{
				rS = (cElementType.string === x.type) ? searchRegExp2(x.value, y) : false;
			}

			switch (operator) {
				case "<>":
					res = !rS;
					break;
				case "=":
				default:
					res = rS;
					break;
			}
		} else if (cElementType.number === y.type) {
			rS = (x.type === y.type);
			switch (operator) {
				case "<>":
					res = !rS || (x.value != y.value);
					break;
				case ">":
					res = rS && (x.value > y.value);
					break;
				case "<":
					res = rS && (x.value < y.value);
					break;
				case ">=":
					res = rS && (x.value >= y.value);
					break;
				case "<=":
					res = rS && (x.value <= y.value);
					break;
				case "=":
				default:
					if (cElementType.string === x.type) {
						var toNumberX = x.tocNumber(doNotParseNum);
						if (toNumberX.value === y.value) {
							res = true;
							break;
						}
						var parseRes = !doNotParseFormat && AscCommon.g_oFormatParser.parse(x.value);
						if (parseRes && parseRes.value === y.value) {
							res = true;
							break;
						}
					}
					res = (x.value === y.value);
					break;
			}
		} else if (cElementType.bool === y.type || cElementType.error === y.type) {
			if (y.type === x.type && x.value === y.value) {
				res = true;
			}
		}
		return res;
	}

	function GetDiffDate360(nDay1, nMonth1, nYear1, nDay2, nMonth2, nYear2, bUSAMethod) {
		var nDayDiff;
		var startTime = new Date(nYear1, nMonth1 - 1, nDay1), endTime = new Date(nYear2, nMonth2 -
			1, nDay2), nY, nM, nD;

		if (startTime > endTime) {
			nY = nYear1;
			nYear1 = nYear2;
			nYear2 = nY;
			nM = nMonth1;
			nMonth1 = nMonth2;
			nMonth2 = nM;
			nD = nDay1;
			nDay1 = nDay2;
			nDay2 = nD;
		}

		if (bUSAMethod) {
			if (nDay1 == 31) {
				nDay1--;
			}
			if (nDay1 == 30 && nDay2 == 31) {
				nDay2--;
			} else {
				if (nMonth1 == 2 && nDay1 == ( new cDate(nYear1, 0, 1).isLeapYear() ? 29 : 28 )) {
					nDay1 = 30;
					if (nMonth2 == 2 && nDay2 == ( new cDate(nYear2, 0, 1).isLeapYear() ? 29 : 28 )) {
						nDay2 = 30;
					}
				}
			}
		//nDayDiff = ( nYear2 - nYear1 ) * 360 + ( nMonth2 - nMonth1 ) * 30 + ( nDay2 - nDay1 );
		} else {
			if (nDay1 == 31) {
				nDay1--;
			}
			if (nDay2 == 31) {
				nDay2--;
			}
		}
		nDayDiff = ( nYear2 - nYear1 ) * 360 + ( nMonth2 - nMonth1 ) * 30 + ( nDay2 - nDay1 );
		return nDayDiff;
	}

	function searchRegExp2(s, mask) {
		//todo протестировать
		var bRes = true;
		s = s.toString().toLowerCase();
		mask = mask.toString().toLowerCase();
		var cCurMask;
		var nSIndex = 0;
		var nMaskIndex = 0;
		var nSLastIndex = 0;
		var nMaskLastIndex = 0;
		var nSLength = s.length;
		var nMaskLength = mask.length;
		var t = false;
		for (; nSIndex < nSLength; nMaskIndex++, nSIndex++, t = false) {
			cCurMask = mask[nMaskIndex];
			if ('~' === cCurMask) {
				nMaskIndex++;
				cCurMask = mask[nMaskIndex];
				t = true;
			} else if ('*' === cCurMask) {
				break;
			}
			if (( cCurMask !== s[nSIndex] && '?' !== cCurMask ) || ( cCurMask !== s[nSIndex] && t)) {
				bRes = false;
				break;
			}
		}
		if (bRes) {
			while (1) {
				cCurMask = mask[nMaskIndex];
				if (nSIndex >= nSLength) {
					while ('*' === cCurMask && nMaskIndex < nMaskLength) {
						nMaskIndex++;
						cCurMask = mask[nMaskIndex];
					}
					bRes = nMaskIndex >= nMaskLength;
					break;
				} else if ('*' === cCurMask) {
					nMaskIndex++;
					if (nMaskIndex >= nMaskLength) {
						bRes = true;
						break;
					}
					nSLastIndex = nSIndex + 1;
					nMaskLastIndex = nMaskIndex;
				} else if (cCurMask !== s[nSIndex] && '?' !== cCurMask) {
					nMaskIndex = nMaskLastIndex;
					nSIndex = nSLastIndex++;
				} else {
					nSIndex++;
					nMaskIndex++;
				}
			}
		}
		return bRes;
	}

	function getArrayHelper(args, func, exceptions) {
		// check for arrays and find max length
		let isContainsArray = false,
			maxRows = 1,
			maxColumns = 1;

		for (let i = 0; i < args.length; i++) {
			if ((cElementType.cellsRange === args[i].type || cElementType.cellsRange3D === args[i].type || cElementType.array === args[i].type) && (!exceptions || (exceptions && !exceptions.get(i)))) {
				let argDimensions = args[i].getDimensions();
				maxRows = argDimensions.row > maxRows ? argDimensions.row : maxRows;
				maxColumns = argDimensions.col > maxColumns ? argDimensions.col : maxColumns;
				isContainsArray = true;
			}
		}

		if (!isContainsArray) {
			return false;
		}

		let resultArr = new cArray();

		for (let i = 0; i < maxRows; i++) {
			resultArr.addRow();
			for (let j = 0; j < maxColumns; j++) {
				let values = [];

				for (let k = 0; k < args.length; k++) {
					let value = args[k];

					if ((cElementType.cellsRange === value.type || cElementType.cellsRange3D === value.type || cElementType.array === value.type) && (!exceptions || (exceptions && !exceptions.get(k)))) {
						let valueDimensions = value.getDimensions();
						if (value.isOneElement()) {
							// single row with single element
							value = value.getFirstElement();
						} else if (valueDimensions.col !== 1 && valueDimensions.row === 1) {
							// single row with many elements
							value = _getValueInRange(value, 0, j);
						} else if (valueDimensions.col === 1 && valueDimensions.row !== 1) {
							// many rows with single element
							value = _getValueInRange(value, i, 0);
						} else {
							value = _getValueInRange(value, i, j);
						}
					}

					values.push(value);
				}

				resultArr.addElement(func(values, true));
			}
		}

		return resultArr;
	}

	// if went beyond the cellsRange
	const _getValueInRange = function (array, _row, _col) {
		let sizes = array.getDimensions();
		if (_row > sizes.row - 1 || _col > sizes.col - 1) {
			return new cError(cErrorType.not_available);
		}
		let res = array.getValueByRowCol ? array.getValueByRowCol(_row, _col, true) : array.getElementRowCol(_row, _col);
		return res;
	}

	/*
	 * Code below has been taken from OpenOffice Source.
	 */

	function lcl_Erf0065(x) {
		var pn = [1.12837916709551256, 1.35894887627277916E-1, 4.03259488531795274E-2, 1.20339380863079457E-3,
			6.49254556481904354E-5], qn = [1.00000000000000000, 4.53767041780002545E-1, 8.69936222615385890E-2,
			8.49717371168693357E-3, 3.64915280629351082E-4];
		var pSum = 0.0, qSum = 0.0, xPow = 1.0;
		for (var i = 0; i <= 4; ++i) {
			pSum += pn[i] * xPow;
			qSum += qn[i] * xPow;
			xPow *= x * x;
		}
		return x * pSum / qSum;
	}

	/** Approximation algorithm for erfc for 0.65 < x < 6.0. */
	function lcl_Erfc0600(x) {
		var pSum = 0, qSum = 0, xPow = 1, pn, qn;

		if (x < 2.2) {
			pn = [9.99999992049799098E-1, 1.33154163936765307, 8.78115804155881782E-1, 3.31899559578213215E-1,
				7.14193832506776067E-2, 7.06940843763253131E-3];
			qn = [1.00000000000000000, 2.45992070144245533, 2.65383972869775752, 1.61876655543871376,
				5.94651311286481502E-1, 1.26579413030177940E-1, 1.25304936549413393E-2];
		} else {
			pn = [9.99921140009714409E-1, 1.62356584489366647, 1.26739901455873222, 5.81528574177741135E-1,
				1.57289620742838702E-1, 2.25716982919217555E-2];
			qn = [1.00000000000000000, 2.75143870676376208, 3.37367334657284535, 2.38574194785344389,
				1.05074004614827206, 2.78788439273628983E-1, 4.00072964526861362E-2];
		}

		for (var i = 0; i < 6; ++i) {
			pSum += pn[i] * xPow;
			qSum += qn[i] * xPow;
			xPow *= x;
		}
		qSum += qn[6] * xPow;
		return Math.exp(-1 * x * x) * pSum / qSum;
	}

	/** Approximation algorithm for erfc for 6.0 < x < 26.54 (but used for all x > 6.0). */
	function lcl_Erfc2654(x) {
		var pn = [5.64189583547756078E-1, 8.80253746105525775, 3.84683103716117320E1, 4.77209965874436377E1,
			8.08040729052301677], qn = [1.00000000000000000, 1.61020914205869003E1, 7.54843505665954743E1,
			1.12123870801026015E2, 3.73997570145040850E1];

		var pSum = 0, qSum = 0, xPow = 1;

		for (var i = 0; i <= 4; ++i) {
			pSum += pn[i] * xPow;
			qSum += qn[i] * xPow;
			xPow /= x * x;
		}
		return Math.exp(-1 * x * x) * pSum / (x * qSum);
	}

	function rtl_math_erf(x) {
		if (x == 0) {
			return 0;
		}

		var bNegative = false;
		if (x < 0) {
			x = Math.abs(x);
			bNegative = true;
		}

		var res = 1;
		if (x < 1.0e-10) {
			res = parseFloat(x * 1.1283791670955125738961589031215452);
		} else if (x < 0.65) {
			res = lcl_Erf0065(x);
		} else {
			res = 1 - rtl_math_erfc(x);
		}

		if (bNegative) {
			res *= -1;
		}

		return res;
	}

	function rtl_math_erfc(x) {
		if (x == 0) {
			return 1;
		}

		var bNegative = false;
		if (x < 0) {
			x = Math.abs(x);
			bNegative = true;
		}

		var fErfc = 0;
		if (x >= 0.65) {
			if (x < 6) {
				fErfc = lcl_Erfc0600(x);
			} else {
				fErfc = lcl_Erfc2654(x);
			}
		} else {
			fErfc = 1 - rtl_math_erf(x);
		}

		if (bNegative) {
			fErfc = 2 - fErfc;
		}

		return fErfc;
	}

	// ToDo use Array.prototype.max, but some like to use for..in without hasOwnProperty
	function getArrayMax (array) {
		//Math.min and Math.max crash on large arrays
		let maxValue = array[0], i, length = array.length;
		for (i = 1; i < length; i++) {
			if (array[i] > maxValue) {
				maxValue = array[i];
			}
		}

		return maxValue;
	}
	// ToDo use Array.prototype.min, but some like to use for..in without hasOwnProperty
	function getArrayMin (array) {
		//Math.min and Math.max crash on large arrays
		let minValue = array[0], i, length = array.length;
		for (i = 1; i < length; i++) {
			if (array[i] < minValue) {
				minValue = array[i];
			}
		}

		return minValue;
	}

	function compareFormula(formula1, refPos1, formula2, offsetRow) {
		if (formula1.length === formula2.length) {
			var index = 0;
			var i, j, bbox, bboxRef, bboxPrev, _3DRefTmp, wsF, wsT;
			for (i = 0; i < refPos1.length; ++i) {
				var refPos = refPos1[i];
				if (!refPos.isName && refPos.oper) {
					for (j = index; j < refPos.start; ++j) {
						if (formula1[j] !== formula2[j]) {
							return false;
						}
					}
					switch (refPos.oper.type) {
						case cElementType.cell:
						case cElementType.cellsRange:
							bboxRef = formula2.substring(refPos.start, refPos.end);
							bbox = AscCommonExcel.g_oRangeCache.getAscRange(bboxRef);
							bboxPrev = refPos.oper.getBBox0();
							break;
						case cElementType.cell3D:
						case cElementType.cellsRange3D:
							_3DRefTmp = parserHelp.is3DRef.call(parserHelp, formula2, refPos.start);
							if (_3DRefTmp[0]) {
								if ((_3DRefTmp[3] || refPos.oper.externalLink) && _3DRefTmp[3] !== refPos.oper.externalLink) {
									return false;
								}
								if (cElementType.cell3D === refPos.oper.type) {
									if (_3DRefTmp[1] !== refPos.oper.getWS().getName()) {
										return false;
									}
									bboxPrev = refPos.oper.getBBox0();
								} else {
									wsF = _3DRefTmp[1];
									wsT = (null !== _3DRefTmp[2]) ? _3DRefTmp[2] : wsF;
									if (!(wsF === refPos.oper.wsFrom.getName() && wsT === refPos.oper.wsTo.getName())) {
										return false;
									}
									bboxPrev = refPos.oper.getBBox0NoCheck();
								}
								bboxRef = formula2.substring(parserHelp.pCurrPos + refPos.start, refPos.end);
								bbox = AscCommonExcel.g_oRangeCache.getAscRange(bboxRef);
							} else {
								return false;
							}
							break;
					}
					if (bboxPrev) {
						if (!(bbox && bboxPrev.isEqualWithOffsetRow(bbox, offsetRow))) {
							return false;
						}
						index = refPos.end;
					}
				}
			}
			for (j = index; j < formula2.length; ++j) {
				if (formula1[j] !== formula2[j]) {
					return false;
				}
			}
			return true;
		}
		return false;
	}

	function convertAreaToArray(area){
		let retArr = new cArray(), _arg0;
		let dimension = area.getDimensions();

		retArr.realSize = {row: dimension.row, col: dimension.col}

		let ws;
		if(cElementType.cellsRange3D === area.type) {
			ws = area.wsFrom;
			area = area.getMatrixNoEmpty()[0];
		} else {
			ws = area.ws;
			area = area.getMatrixNoEmpty();
		}

		if (dimension) {
			let oBBox = dimension.bbox,
				minC = Math.min( ws.getColDataLength(), oBBox.c2 ),
				minR = Math.min( ws.cellsByColRowsCount - 1, oBBox.r2 ),
				rowCount = (minR - oBBox.r1) >= 0 ? minR - oBBox.r1 + 1 : 0,
				colCount = (minC - oBBox.c1) >= 0 ? minC - oBBox.c1 + 1 : 0;

			if (rowCount < dimension.row || colCount < dimension.col) {
				retArr.missedValue = new cEmpty();
			} else {
				retArr.realSize = null;
			}

			if (area && area.length < 1) {
				// let emptyElem = new cEmpty();
				// if array is empty - add info about range size and set missedValue
				retArr.setRealArraySize(dimension.row, dimension.col);
				retArr.missedValue = new cEmpty();

				/* we add one element to the array so as not to return a completely empty array to the formula */
				if (!retArr.countElement) {
					retArr.addRow();
					retArr.addElement(retArr.missedValue);
				}

			} else {
				for ( let iRow = 0; iRow < rowCount; iRow++, iRow < rowCount ? retArr.addRow() : true ) {
					for ( let iCol = 0; iCol < colCount; iCol++ ) {
						_arg0 = area[iRow] && area[iRow][iCol] ? area[iRow][iCol] : new cEmpty();
						retArr.addElement(_arg0);
					}
				}
			}
		}

		return retArr;
	}

	function convertRefToRowCol (ref, curRef) {
		var cellAddress = new AscCommon.CellAddress(ref);

		var res = "R";
		res += !cellAddress.bRowAbs && curRef ? "[" + (cellAddress.row - curRef.nRow - 1) + "]" : cellAddress.row;
		res += "C";
		res += !cellAddress.bColAbs && curRef ? "[" + (cellAddress.col - curRef.nCol - 1) + "]" : cellAddress.col;

		return res;
	}
	function convertAreaToArrayRefs(area, useOnlyFirstRow, useOnlyFirstColumn){
		var retArr = new cArray(), ref, is3d;
		var range, ws;
		if(cElementType.cellsRange === area.type) {
			range = area.range;
			ws = area.ws;
		} else if (cElementType.cellsRange3D === area.type && area.isSingleSheet()) {
			range = area.getRanges()[0];
			ws = area.wsFrom;
			is3d = true;
		}

		if(range) {
			var bbox = range.bbox;

			var countRow = useOnlyFirstRow ? 0 : bbox.r2 - bbox.r1;
			var countCol = useOnlyFirstColumn ? 0 : bbox.c2 - bbox.c1;

			for ( var iRow = bbox.r1; iRow <= countRow + bbox.r1; iRow++, iRow <= countRow + bbox.r1 ? retArr.addRow() : true ) {
				for ( var iCol = bbox.c1; iCol <= countCol + bbox.c1; iCol++ ) {
					var curCol = useOnlyFirstColumn ? bbox.c1 : iCol;
					var curRow = useOnlyFirstRow ? bbox.r1 : iRow;
					ref = new Asc.Range(curCol, curRow, curCol, curRow);
					ref = is3d ? new cRef3D(ref.getName(), ws) : new cRef(ref.getName(), ws);
					retArr.addElement(ref);
				}
			}
		}

		return retArr;
	}

	function specialFuncArrayToArray(arg0, arg1, what) {
		let retArr = null, _arg0, _arg1;
		let iRow, iCol;

		let arg0RowCount = arg0.getRowCount(true),
			arg0ColCount = arg0.getCountElementInRow(true),
			arg1RowCount = arg1.getRowCount(true),
			arg1ColCount = arg1.getCountElementInRow(true);

		if (arg0RowCount === arg1RowCount && 1 === arg0ColCount) {
			retArr = new cArray();
			for (iRow = 0; iRow < arg1RowCount; iRow++, iRow < arg1RowCount ? retArr.addRow() : true) {
				for (iCol = 0; iCol < arg1ColCount; iCol++) {
					_arg0 = arg0.getElementRowCol(iRow, 0, true);
					_arg1 = arg1.getElementRowCol(iRow, iCol, true);
					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		} else if (arg0RowCount === arg1RowCount && 1 === arg1ColCount) {
			retArr = new cArray();
			for (iRow = 0; iRow < arg0RowCount; iRow++, iRow < arg0RowCount ? retArr.addRow() : true) {
				for (iCol = 0; iCol < arg0ColCount; iCol++) {
					_arg0 = arg0.getElementRowCol(iRow, iCol, true);
					_arg1 = arg1.getElementRowCol(iRow, 0, true);
					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		} else if (arg0ColCount === arg1ColCount && 1 === arg0RowCount) {
			retArr = new cArray();
			for (iRow = 0; iRow < arg1RowCount; iRow++, iRow < arg1RowCount ? retArr.addRow() : true) {
				for (iCol = 0; iCol < arg1ColCount; iCol++) {
					_arg0 = arg0.getElementRowCol(0, iCol, true);
					_arg1 = arg1.getElementRowCol(iRow, iCol, true);
					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		} else if (arg0ColCount === arg1ColCount && 1 === arg1RowCount) {
			retArr = new cArray();
			for (iRow = 0; iRow < arg0RowCount; iRow++, iRow < arg0RowCount ? retArr.addRow() : true) {
				for (iCol = 0; iCol < arg0ColCount; iCol++) {
					_arg0 = arg0.getElementRowCol(iRow, iCol, true);
					_arg1 = arg1.getElementRowCol(0, iCol, true);
					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		} else if (1 === arg0ColCount && 1 === arg1RowCount) {
			retArr = new cArray();
			for (iRow = 0; iRow < arg0RowCount; iRow++, iRow < arg0RowCount ? retArr.addRow() : true) {
				for (iCol = 0; iCol < arg1ColCount; iCol++) {
					_arg0 = arg0.getElementRowCol(iRow, 0, true);
					_arg1 = arg1.getElementRowCol(0, iCol, true);
					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		} else if (1 === arg1ColCount && 1 === arg0RowCount) {
			retArr = new cArray();
			for (iRow = 0; iRow < arg1RowCount; iRow++, iRow < arg1RowCount ? retArr.addRow() : true) {
				for (iCol = 0; iCol < arg0ColCount; iCol++) {
					_arg0 = arg0.getElementRowCol(0, iCol, true);
					_arg1 = arg1.getElementRowCol(iRow, 0, true);
					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		} else if (arg0.getCountElement() !== arg1.getCountElement() || arg0RowCount !== arg1RowCount || arg0ColCount !== arg1ColCount) {
			let errNA = new cError(cErrorType.not_available);

			// if there is only one element in the range, get this element and call the function again
			if (arg0.isOneElement()) {
				let arg0FirstElem = arg0.getFirstElement();
				return _func[arg0FirstElem.type][arg1.type](arg0FirstElem, arg1, what);
			} else if (arg1.isOneElement()) {
				let arg1FirstElem = arg1.getFirstElement();
				return _func[arg0.type][arg1FirstElem.type](arg0, arg1FirstElem, what);
			}

			// Logic:
			// find the effective range (the one that is involved in the calculations)
			// calculate it
			// then fill the remaining rows and columns with N/A errors

			let arrayMaxRows = Math.max(arg0RowCount, arg1RowCount), arrayMaxCols = Math.max(arg0ColCount, arg1ColCount);
			retArr = new cArray();
			retArr.setRealArraySize(arrayMaxRows, arrayMaxCols);

			// arg0RowCount, arg0ColCount, arg1RowCount, arg1ColCount
			let usefulCol = Math.min(arg0ColCount, arg1ColCount),
				usefulRow = Math.min(arg0RowCount, arg1RowCount),
				arg0Dimensions = arg0.getDimensions(),
				arg1Dimensions = arg1.getDimensions();

			// if we have one of the element with single row|col, set the value to be obtained from this particular row or column
			let fromArg0Row, fromArg1Row, fromArg0Col, fromArg1Col;
			if (arg0Dimensions.row === 1 && arrayMaxRows > 1) {
				// arg1.row more than arg0.row
				usefulRow = arrayMaxRows;
				fromArg0Row = 0;
			}
			if (arg1Dimensions.row === 1 && arrayMaxRows > 1) {
				// arg0.row more than arg1.row
				usefulRow = arrayMaxRows;
				fromArg1Row = 0;
			}
			if (arg0Dimensions.col === 1 && arrayMaxCols > 1) {
				// arg1.col more than arg0.col
				usefulCol = arrayMaxCols;
				fromArg0Col = 0;
			}
			if (arg1Dimensions.col === 1 && arrayMaxCols > 1) {
				// arg0.col more than arg1.col
				usefulCol = arrayMaxCols;
				fromArg1Col = 0;
			}

			// fill the array
			for (let iRow = 0; iRow < arrayMaxRows; iRow++, iRow < usefulRow ? retArr.addRow() : true) {
				if (iRow >= usefulRow) {
					// fill row with N/A and continue
					let errRow = new Array(arrayMaxCols).fill(errNA);
					retArr.pushRow([errRow], 0);
					continue
				}

				for (let iCol = 0; iCol < arrayMaxCols; iCol++) {
					if (iCol >= usefulCol) {
						// add N/A error and continue
						retArr.addElement(errNA);
						continue
					}

					_arg0 = arg0.getElementRowCol(fromArg0Row !== undefined ? fromArg0Row : iRow, fromArg0Col !== undefined ? fromArg0Col : iCol, true);
					_arg1 = arg1.getElementRowCol(fromArg1Row !== undefined ? fromArg1Row : iRow, fromArg1Col !== undefined ? fromArg1Col : iCol, true);

					retArr.addElement(_func[_arg0.type][_arg1.type](_arg0, _arg1, what));
				}
			}
		}
		return retArr;
	}

	//----------------------------------------------------------export----------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].cElementType = cElementType;
	window['AscCommonExcel'].cErrorType = cErrorType;
	window['AscCommonExcel'].cElementTypeWeight = cElementTypeWeight;
	window['AscCommonExcel'].cExcelSignificantDigits = cExcelSignificantDigits;
	window['AscCommonExcel'].cExcelMaxExponent = cExcelMaxExponent;
	window['AscCommonExcel'].cExcelMinExponent = cExcelMinExponent;
	window['AscCommonExcel'].c_Date1904Const = c_Date1904Const;
	window['AscCommonExcel'].c_Date1900Const = c_Date1900Const;
	window['AscCommonExcel'].c_DateCorrectConst = c_Date1900Const;
	window['AscCommonExcel'].cNumFormatFirstCell = cNumFormatFirstCell;
	window['AscCommonExcel'].cNumFormatNone = cNumFormatNone;
	window['AscCommonExcel'].g_cCalcRecursion = g_cCalcRecursion;
	window['AscCommonExcel'].g_ProcessShared = false;
	window['AscCommonExcel'].cReturnFormulaType = cReturnFormulaType;
	window['AscCommonExcel'].arrayIndexesType = arrayIndexesType;

	window['AscCommonExcel'].bIsSupportArrayFormula = bIsSupportArrayFormula;
	window['AscCommonExcel'].bIsSupportDynamicArrays = bIsSupportDynamicArrays;

	window['AscCommonExcel'].aExcludeRecursiveFormulas = aExcludeRecursiveFormulas;

	window['AscCommonExcel'].cReplaceFormulaType = cReplaceFormulaType;


	window['AscCommonExcel'].cNumber = cNumber;
	window['AscCommonExcel'].cString = cString;
	window['AscCommonExcel'].cBool = cBool;
	window['AscCommonExcel'].cError = cError;
	window['AscCommonExcel'].cArea = cArea;
	window['AscCommonExcel'].cArea3D = cArea3D;
	window['AscCommonExcel'].cRef = cRef;
	window['AscCommonExcel'].cRef3D = cRef3D;
	window['AscCommonExcel'].cEmpty = cEmpty;
	window['AscCommonExcel'].cName = cName;
	window['AscCommonExcel'].cName3D = cName3D;
	window['AscCommonExcel'].cArray = cArray;
	window['AscCommonExcel'].cUndefined = cUndefined;
	window['AscCommonExcel'].cBaseFunction = cBaseFunction;
	window['AscCommonExcel'].cUnknownFunction = cUnknownFunction;
	window['AscCommonExcel'].cStrucTable = cStrucTable;
	window['AscCommonExcel'].cBaseOperator = cBaseOperator;
	window['AscCommonExcel'].cStrucPivotTable = cStrucPivotTable;

	window['AscCommonExcel'].checkTypeCell = checkTypeCell;
	window['AscCommonExcel'].cFormulaFunctionGroup = cFormulaFunctionGroup;
	window['AscCommonExcel'].cFormulaFunction = cFormulaFunction;

	window['AscCommonExcel'].cFormulaFunctionLocalized = null;
	window['AscCommonExcel'].cFormulaFunctionToLocale = null;

	window['AscCommonExcel'].getFormulasInfo = getFormulasInfo;
	window['AscCommonExcel'].getRangeByRef = getRangeByRef;
	window['AscCommonExcel'].addNewFunction = addNewFunction;
	window['AscCommonExcel'].removeCustomFunction = removeCustomFunction;
	window['AscCommonExcel'].getRangeByName = getRangeByName;

	window['AscCommonExcel']._func = _func;

	window['AscCommonExcel'].parserFormula = parserFormula;
	window['AscCommonExcel'].ParseResult = ParseResult;
	window['AscCommonExcel'].CalculateResult = CalculateResult;

	window['AscCommonExcel'].parseNum = parseNum;
	window['AscCommonExcel'].matching = matching;
	window['AscCommonExcel'].matchingValue = matchingValue;
	window['AscCommonExcel'].GetDiffDate360 = GetDiffDate360;
	window['AscCommonExcel'].searchRegExp2 = searchRegExp2;
	window['AscCommonExcel'].rtl_math_erf = rtl_math_erf;
	window['AscCommonExcel'].rtl_math_erfc = rtl_math_erfc;
	window['AscCommonExcel'].getArrayMax = getArrayMax;
	window['AscCommonExcel'].getArrayMin = getArrayMin;
	window['AscCommonExcel'].compareFormula = compareFormula;
	window['AscCommonExcel'].convertRefToRowCol = convertRefToRowCol;
	window['AscCommonExcel'].convertAreaToArray = convertAreaToArray;
	window['AscCommonExcel'].convertAreaToArrayRefs = convertAreaToArrayRefs;
	window['AscCommonExcel'].getArrayHelper = getArrayHelper;
	window['AscCommonExcel'].getMaxDate = getMaxDate;

	window['AscCommonExcel'].importRangeLinksState = importRangeLinksState;

})(window);
