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

(/**
* @param {Window} window
* @param {undefined} undefined
*/
function(window, undefined) {
// Import
var CellValueType = AscCommon.CellValueType;

var c_oAscNumFormatType = Asc.c_oAscNumFormatType;

var gc_sFormatDecimalPoint = ".";
var gc_sFormatThousandSeparator = ",";
var LocaleFormatSymbol ={};
var numFormat_Text = 0;
var numFormat_TextPlaceholder = 1;
var numFormat_Bracket = 2;
var numFormat_Digit = 3;
var numFormat_DigitNoDisp = 4;
var numFormat_DigitSpace = 5;
var numFormat_DecimalPoint = 6;
var numFormat_DecimalFrac = 7;
var numFormat_Thousand = 8;
var numFormat_Scientific = 9;
var numFormat_Repeat = 10;
var numFormat_Skip = 11;
var numFormat_Year = 12;
var numFormat_Month = 13;
var numFormat_Minute = 14;
var numFormat_Hour = 15;
var numFormat_Day = 16;
var numFormat_Second = 17;
var numFormat_Milliseconds = 18;
var numFormat_AmPm = 19;
var numFormat_DateSeparator = 20;
var numFormat_TimeSeparator = 21;
var numFormat_DecimalPointText = 22;
//Вспомогательные типы, которые заменятюся в _prepareFormat
var numFormat_MonthMinute = 101;
var numFormat_Percent = 102;
var numFormat_General = 103;
var numFormat_DigitDrop = 104;
var numFormat_Plus = 105;
var numFormat_Minus = 106;
var numFormat_ThousandText = 107;
var numFormat_DayOfWeek = 110;

var FormatStates = {Decimal: 1, Frac: 2, Scientific: 3, Slash: 4, SlashFrac: 5};
var SignType = {Negative: 1, Null:2, Positive: 3};

var gc_nMaxDigCount = 15;//Максимальное число знаков точности
var gc_nMaxDigCountView = 11;//Максимальное число знаков в ячейке
var gc_nMaxMantissa = Math.pow(10, gc_nMaxDigCount);
var gc_aTimeFormats = ['[$-F400]h:mm:ss AM/PM', 'h:mm;@', 'h:mm AM/PM;@', 'h:mm:ss;@', 'h:mm:ss AM/PM;@', 'mm:ss.0;@',
	'[h]:mm:ss;@'];
var gc_aFractionFormats = ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'];
const dBNum1Numbers = ['\u3007','\u4E00','\u4E8C','\u4E09','\u56DB','\u4E94','\u516D','\u4E03','\u516B','\u4E5D'];

var NumComporationOperators =
{
	equal: 1,
	greater: 2,
	less: 3,
	greaterorequal: 4,
	lessorequal: 5,
	notequal: 6
};
var NumFormatType =
{
	Excel: 1,
	WordFieldDate: 2,
	WordFieldNumeric: 3,
	PDFFormDate: 4
};

function getNumberParts(x)
{
    var sig = SignType.Null;
    if (!isFinite(x))
        x = 0;
	if(x > 0)
		sig = SignType.Positive;
	else if(x < 0)
	{
		sig = SignType.Negative;
		x = Math.abs(x);
	}
    var exp = - gc_nMaxDigCount;
	var man = 0;
	if(SignType.Null != sig)
	{
		exp = Math.floor( Math.log(x) * Math.LOG10E ) - gc_nMaxDigCount + 1;
		//хотелось бы поставить здесь floor, чтобы не округлялось число 0.9999999999999999, но обнаружились проблемы с числом 0.999999999999999
		//после умножения оно превращается в 999999999999998.9
		man = Math.round(x / Math.pow(10, exp));
		if(man >= gc_nMaxMantissa)
		{
			exp++;
			man/=10;
		}
	}
    return {mantissa: man, exponent: exp, sign: sig};//для 0,123 exponent == - gc_nMaxDigCount
}

	function compareNumbers(val1, val2) {
		var res = 0;
		var parts1 = getNumberParts(val1);
		var parts2 = getNumberParts(val2);
		if (parts1.sign === parts2.sign) {
			if (parts1.exponent === parts2.exponent) {
				res = parts1.mantissa - parts2.mantissa;
			} else {
				res = parts1.exponent - parts2.exponent;
			}
			if (SignType.Negative === parts1.sign) {
				res = -res;
			}
		} else {
			res = parts1.sign - parts2.sign;
		}
		return res;
	}

    function isNumber(n) {
        return !isNaN(parseFloat(n)) && isFinite(n);
    }
	function round10(value, exp1, exp2) {
		//todo use Math.round10
		// Shift
		value = value.toString().split('e');
		value = Math.round(+(value[0] + 'e' + (value[1] ? (+value[1] + exp1) : exp1)));
		// Shift back
		value = value.toString().split('e');
		return +(value[0] + 'e' + (value[1] ? (+value[1] - exp2) : -exp2));
	}

function FormatObj(type, val)
{
    this.type = type;
    this.val = val;//что здесь лежит определяется типом
}
function FormatObjScientific(val, format, sign)
{
    this.type = numFormat_Scientific;
    this.val = val;//E или e
    this.format = format;//array формата
    this.sign = sign;
}
function FormatObjDecimalFrac(aLeft, aRight)
{
    this.type = numFormat_DecimalFrac;
    this.aLeft = aLeft;//array формата левой части
    this.aRight = aRight;//array формата правой части
    this.bNumRight = false;
	this.numerator = 0;
	this.denominator = 0;
}
function FormatObjDateVal(type, nCount, bElapsed)
{
    this.type = type;
    this.val = nCount;//Количество знаков подряд
    this.bElapsed = bElapsed;//true == [hhh]; в квадратных скобках
}
function FormatObjBracket(sData)
{
    this.type = numFormat_Bracket;
    this.val = sData;
    this.parse = function(data)
    {
        var length = data.length;
        if(length > 0)
        {
            var first = data[0];
            if("$" == first)
            {
                var aParams = data.substring(1).split('-');
				if (aParams[0].length > 0) {
					this.CurrencyString = aParams[0];
				} if(aParams.length > 1 && aParams[1].length > 0) {
					this.Lid = aParams[1];
				}
            }
			else if("=" == first || ">" == first || "<" == first)
			{
				var nIndex = 1;
				var sOperator = first;
				if(length > 1 && (">" == first || "<" == first))
				{
					var second = data[1];
					if("=" == second || (">" == second && "<" == first))
					{
						sOperator += second;
						nIndex = 2;
					}
				}
				switch(sOperator)
				{
					case "=": this.operator = NumComporationOperators.equal;break;
					case ">": this.operator = NumComporationOperators.greater;break;
					case "<": this.operator = NumComporationOperators.less;break;
					case ">=": this.operator = NumComporationOperators.greaterorequal;break;
					case "<=": this.operator = NumComporationOperators.lessorequal;break;
					case "<>": this.operator = NumComporationOperators.notequal;break;
				}
				this.operatorValue = parseFloat(data.substring(nIndex));
			}
            else
            {
				var sLowerColor = data.toLowerCase();
                //todo Color1-56
                if("black" == sLowerColor)
                    this.color = 0x000000;
                else if("blue" == sLowerColor)
                    this.color = 0x0000ff;
                else if("cyan" == sLowerColor)
                    this.color = 0x00ffff;
                else if("green" == sLowerColor)
                    this.color = 0x00ff00;
                else if("magenta" == sLowerColor)
                    this.color = 0xff00ff;
                else if("red" == sLowerColor)
                    this.color = 0xff0000;
                else if("white" == sLowerColor)
                    this.color = 0xffffff;
                else if("yellow" == sLowerColor)
                    this.color = 0xffff00;
                else if("y" == first || "m" == first || "d" == first || "h" == first || "s" == first ||
                    "Y" == first || "M" == first || "D" == first || "H" == first || "S" == first ||
					"a" == first)
                {
                    var bSame = true;
                    var nCount = 1;
                    for(var i = 1; i < length; ++i)
                    {
                        if(first != data[i])
                        {
                            bSame = false;
                            break;
                        }
                        nCount++;
                    }
                    if(true == bSame)
                    {
                        switch(first)
                        {
                            case "Y":
                            case "y": this.dataObj = new FormatObjDateVal(numFormat_Year, nCount, true);break;
                            case "M":
                            case "m": this.dataObj = new FormatObjDateVal(numFormat_MonthMinute, nCount, true);break;
                            case "D":
                            case "d": this.dataObj = new FormatObjDateVal(numFormat_Day, nCount, true);break;
                            case "H":
                            case "h": this.dataObj = new FormatObjDateVal(numFormat_Hour, nCount, true);break;
                            case "S":
                            case "s": this.dataObj = new FormatObjDateVal(numFormat_Second, nCount, true);break;
                            case "a": this.dataObj = new FormatObjDateVal(numFormat_DayOfWeek, nCount, true);break;
                        }
                    }
                }
            }
        }
    };
    this.parse(sData);
}
function ParseLocalFormatSymbol(Name)
{
	LocaleFormatSymbol['Y'] = 'Y';
	LocaleFormatSymbol['y'] = 'y';
	LocaleFormatSymbol['M'] = 'M';
	LocaleFormatSymbol['m'] = 'm';
	LocaleFormatSymbol['D'] = 'D';
	LocaleFormatSymbol['d'] = 'd';
	LocaleFormatSymbol['H'] = 'H';
	LocaleFormatSymbol['h'] = 'h';
	LocaleFormatSymbol['Minute'] = 'M';
	LocaleFormatSymbol['minute'] = 'm';
	LocaleFormatSymbol['S'] = 'S';
	LocaleFormatSymbol['s'] = 's';
	LocaleFormatSymbol['a'] = 'a';
	LocaleFormatSymbol['general'] = 'General';
	switch (Name) {
//___________________________________________________fi________________________________________________________________
		case("fi"):
		case("smn"):
		case("sms"):
		case("fi-FI"):
		case("se-FI"):
		case("smn-FI"):
		case("sms-FI"):
		case("sv-AX"):
		case("sv-FI"):
		case("en-FI"): {
			LocaleFormatSymbol['Y'] = 'V';
			LocaleFormatSymbol['y'] = 'v';
			LocaleFormatSymbol['M'] = 'K';
			LocaleFormatSymbol['m'] = 'k';
			LocaleFormatSymbol['D'] = 'P';
			LocaleFormatSymbol['d'] = 'p';
			LocaleFormatSymbol['H'] = 'T';
			LocaleFormatSymbol['h'] = 't';
			LocaleFormatSymbol['general'] = 'Yleinen';
			break;
		}
//___________________________________________________fy________________________________________________________________
		case("fy"):
		case("nds"):
		case("nl"):
		case("en-NL"):
		case("fy-NL"):
		case("nds-NL"):
		case("nl-BE"):
		case("nl-NL"): {
			LocaleFormatSymbol['Y'] = 'J';
			LocaleFormatSymbol['y'] = 'j';
			LocaleFormatSymbol['H'] = 'U';
			LocaleFormatSymbol['h'] = 'u';
			LocaleFormatSymbol['general'] = 'Standaard';
			break;
		}
//___________________________________________________ES________________________________________________________________
		case("ast"):
		case("eu"):
		case("gl"):
		case("ast-ES"):
		case("ca-ES"):
		case("es-ES"):
		case("es-MX"):
		case("eu-ES"):
		case("gl-ES"):
		case("ca-ES-valencia"): {
			LocaleFormatSymbol['Y'] = 'A';
			LocaleFormatSymbol['y'] = 'a';
			LocaleFormatSymbol['a'] = 'o';
			LocaleFormatSymbol['general'] = 'Estándar';
			break;
		}
		case("pt-BR"):
		case("es-BR"): {
			LocaleFormatSymbol['Y'] = 'A';
			LocaleFormatSymbol['y'] = 'a';
			LocaleFormatSymbol['a'] = 'o';
			LocaleFormatSymbol['general'] = 'Geral';
			break;
		}
		case("pt"):
		case("pt-PT"): {
			LocaleFormatSymbol['Y'] = 'A';
			LocaleFormatSymbol['y'] = 'a';
			LocaleFormatSymbol['a'] = 'o';
			LocaleFormatSymbol['general'] = 'Éstandar';
			break;
		}
//____________________________________________________ru_______________________________________________________________
		case("ba"):
		case("ce"):
		case("cu"):
		case("kk"):
		case("os"):
		case("rm"):
		case("ru"):
		case("sah"):
		case("tt"):
		case("wae"):
		case("ba-RU"):
		case("ce-RU"):
		case("cu-RU"):
		case("de-BE"):
		case("en-BE"):
		case("en-CH"):
		case("kk-KZ"):
		case("os-RU"):
		case("pt-CH"):
		case("rm-CH"):
		case("ru-KZ"):
		case("ru-RU"):
		case("sah-RU"):
		case("tt-RU"):
		case("wae-CH"): {
			LocaleFormatSymbol['Y'] = 'Г';
			LocaleFormatSymbol['y'] = 'г';
			LocaleFormatSymbol['M'] = 'М';
			LocaleFormatSymbol['m'] = 'М';
			LocaleFormatSymbol['D'] = 'Д';
			LocaleFormatSymbol['d'] = 'д';
			LocaleFormatSymbol['H'] = 'Ч';
			LocaleFormatSymbol['h'] = 'ч';
			LocaleFormatSymbol['Minute'] = 'М';
			LocaleFormatSymbol['minute'] = 'м'
			LocaleFormatSymbol['S'] = 'C';
			LocaleFormatSymbol['s'] = 'с';
			LocaleFormatSymbol['general'] = 'Основной';
			break;
		}
//____________________________________________________fr_______________________________________________________________
		case("oc"):
		case("br"):
		case("co"):
		case("fr"):
		case("br-FR"):
		case("ca-FR"):
		case("co-FR"):
		case("fr-BE"):
		case("fr-CA"):
		case("fr-CH"):
		case("fr-FR"):
		case("gsw-FR"): {
			LocaleFormatSymbol['Y'] = 'A';
			LocaleFormatSymbol['y'] = 'a';
			LocaleFormatSymbol['D'] = 'J';
			LocaleFormatSymbol['d'] = 'j';
			LocaleFormatSymbol['a'] = 'o';
			LocaleFormatSymbol['general'] = 'Standard';
			break;
		}
//____________________________________________________de_______________________________________________________________
		case("de"):
		case("ksh"):
		case("dsb"):
		case("hsb"):
		case("de-AT"):
		case("de-CH"):
		case("de-DE"):
		case("dsb-DE"):
		case("en-AT"):
		case("en-DE"):
		case("hsb-DE"):
		case("ksh-DE"):
		case("nds-DE"): {
			LocaleFormatSymbol['Y'] = 'J';
			LocaleFormatSymbol['y'] = 'j';
			LocaleFormatSymbol['M'] = 'M';
			LocaleFormatSymbol['m'] = 'M';
			LocaleFormatSymbol['Minute'] = 'M';
			LocaleFormatSymbol['minute'] = 'm';
			LocaleFormatSymbol['D'] = 'T';
			LocaleFormatSymbol['d'] = 't';
			LocaleFormatSymbol['general'] = 'Standard';
			break;
		}
//____________________________________________________it_______________________________________________________________
		case("ca"):
		case("it"):
		case("fur"):
		case("ca-IT"):
		case("de-IT"):
		case("fur-IT"):
		case("it-CH"):
		case("it-IT"):
		case("it-VA"): {
			LocaleFormatSymbol['Y'] = 'A';
			LocaleFormatSymbol['y'] = 'a';
			LocaleFormatSymbol['D'] = 'G';
			LocaleFormatSymbol['d'] = 'g';
			LocaleFormatSymbol['a'] = 'o';
			LocaleFormatSymbol['general'] = 'Standard';
			break;
		}
//____________________________________________________da_______________________________________________________________
		case("sv"):
		case("en-SE"):
		case("se-SE"):
		case("sma-SE"):
		case("smj-SE"):
		case("sv-SE"): {
			LocaleFormatSymbol['Y'] = 'Å';
			LocaleFormatSymbol['y'] = 'å';
			LocaleFormatSymbol['m'] = 'M'
			LocaleFormatSymbol['M'] = 'M';
			LocaleFormatSymbol['Minute'] = 'M';
			LocaleFormatSymbol['minute'] = 'm';
			LocaleFormatSymbol['H'] = 'T';
			LocaleFormatSymbol['h'] = 't';
			LocaleFormatSymbol['general'] = 'Standard';
			break;
		}
		case("nb"):
		case("nn"):
		case("se"):
		case("smj"):
		case("sma"):
		case("fo"):
		case("da"):
		case("smj-NO"):
		case("sma-NO"):
		case("se-NO"):
		case("nn-NO"):
		case("nb-SJ"):
		case("nb-NO"):
		case("fo-DK"):
		case("da-DK"): {
			LocaleFormatSymbol['Y'] = 'Å';
			LocaleFormatSymbol['y'] = 'å';
			LocaleFormatSymbol['H'] = 'T';
			LocaleFormatSymbol['h'] = 't';
			LocaleFormatSymbol['general'] = 'Standard';
			break;
		}
//_____________________________________________________ch______________________________________________________________
		case("bo"):
		case("ii"):
		case("ug"):
		case("zh"):
		case("bo-CN"):
		case("ii-CN"):
		case("mn-Mong-CN"):
		case("ug-CN"):
		case("zh-CN"):
		case("zh-Hans"):
		case("zh-TW"): {
			LocaleFormatSymbol['general'] = 'G/通用格式';
			break;
		}
//_________________________________________________special_____________________________________________________________
		case("el"):
		case("el-GR"): {
			LocaleFormatSymbol['Y'] = 'Ε';
			LocaleFormatSymbol['y'] = 'ε';
			LocaleFormatSymbol['M'] = 'Μ';
			LocaleFormatSymbol['m'] = 'μ';
			LocaleFormatSymbol['D'] = 'Η';
			LocaleFormatSymbol['d'] = 'η';
			LocaleFormatSymbol['H'] = 'Ω';
			LocaleFormatSymbol['h'] = 'ω';
			LocaleFormatSymbol['Minute'] = 'Λ';
			LocaleFormatSymbol['minute'] = 'λ';
			LocaleFormatSymbol['S'] = 'Δ';
			LocaleFormatSymbol['s'] = 'δ';
			LocaleFormatSymbol['general'] = 'Γενικός τύπος';
			break;
		}
		case("hu"):
		case("hu-HU"): {
			LocaleFormatSymbol['Y'] = 'É';
			LocaleFormatSymbol['y'] = 'é';
			LocaleFormatSymbol['M'] = 'H';
			LocaleFormatSymbol['m'] = 'h';
			LocaleFormatSymbol['D'] = 'N';
			LocaleFormatSymbol['d'] = 'n';
			LocaleFormatSymbol['H'] = 'Ó';
			LocaleFormatSymbol['h'] = 'ó';
			LocaleFormatSymbol['Minute'] = 'P';
			LocaleFormatSymbol['minute'] = 'p';
			LocaleFormatSymbol['S'] = 'M';
			LocaleFormatSymbol['s'] = 'm';
			LocaleFormatSymbol['general'] = 'Normál';
			break;
		}
		case("tr"):
		case("tr-TR"): {
			LocaleFormatSymbol['M'] = 'A';
			LocaleFormatSymbol['m'] = 'a';
			LocaleFormatSymbol['D'] = 'G';
			LocaleFormatSymbol['d'] = 'g';
			LocaleFormatSymbol['H'] = 'S';
			LocaleFormatSymbol['h'] = 's';
			LocaleFormatSymbol['Minute'] = 'D';
			LocaleFormatSymbol['minute'] = 'd';
			LocaleFormatSymbol['S'] = 'N';
			LocaleFormatSymbol['s'] = 'n';
			LocaleFormatSymbol['a'] = 'o';
			LocaleFormatSymbol['general'] = 'Genel';
			break;
		}
		case("pl"):
		case("pl-PL"): {
			LocaleFormatSymbol['Y'] = 'R';
			LocaleFormatSymbol['y'] = 'r';
			LocaleFormatSymbol['H'] = 'G';
			LocaleFormatSymbol['h'] = 'g';
			LocaleFormatSymbol['general'] = 'Standardowy';
			break;
		}
		case("cs"):
		case("cs-CZ"): {
			LocaleFormatSymbol['Y'] = 'R';
			LocaleFormatSymbol['y'] = 'r';
			LocaleFormatSymbol['general'] = 'Vęeobecný';
			break;
		}
		case("ja"):
		case("ja-JP"): {
			LocaleFormatSymbol['general'] = 'G/標準';
			break;
		}
		case("ko"):
		case("ko-KR"): {
			LocaleFormatSymbol['general'] = 'G/표준';
			break;
		}
	}
	return true;
}
function NumFormat(bAddMinusIfNes)
{
    //Stream чтения формата
    this.formatString = "";
    this.length = this.formatString.length;
    this.index = 0;
    this.EOF = -1;
    
    //Формат
    this.aRawFormat = [];
    this.aDecFormat = [];
    this.aFracFormat = [];
    this.bDateTime = false;
	this.bDate = false;
	this.bTime = false;//флаг, чтобы отличить формат даты с временем, от простой даты
	this.bDay = false;//чтобы отличать когда надо использовать MonthGenitiveNames
    this.nPercent = 0;
    this.bScientific = false;
    this.bThousandSep = false;
    this.nThousandScale = 0;
    this.bTextFormat = false;
    this.bTimePeriod = false;
    this.bMillisec = false;
    this.bSlash = false;
    this.bWhole = false;
	this.bCurrency = false;
	this.bRepeat = false;
    this.Color = -1;
	this.ComporationOperator = null;
	this.LCID = null;
	this.CurrencyString = null;
	this.DBNum = 0;

	this.bGeneralChart = false;//если в формате только один текст(например в chart "Основной")
    this.bAddMinusIfNes = bAddMinusIfNes;//когда не задано форматирование для отрицательных чисел иногда надо вставлять минус
}
NumFormat.prototype =
{
    _getChar : function()
    {
        if(this.index < this.length)
        {
            return this.formatString[this.index];
        }
        return this.EOF;
    },
    _readChar : function()
    {
        var curChar = this._getChar();
        if(this.index < this.length)
            this.index++;
        return curChar;
    },
    _skip : function(val)
    {
        var nNewIndex = this.index + val;
        if(nNewIndex >= 0)
            this.index = nNewIndex;
    },
    _addToFormat : function(type, val)
    {
        var oFormatObj = new FormatObj(type, val);
        this.aRawFormat.push(oFormatObj);
    },
    _addToFormat2 : function(oFormatObj)
    {
        this.aRawFormat.push(oFormatObj);
    },
    _ReadText : function(endChar)
    {
        var sText = "";
        while(true)
        {
            var next = this._readChar();
            if(this.EOF == next || endChar == next)
                break;
            else
            {
                sText += next;
            }
        }
        this._addToFormat(numFormat_Text, sText);
    },
    _GetText : function(len)
    {
        return this.formatString.substr(this.index, len);
    },
    _ReadChar : function()
    {
        var next = this._readChar();
        if(this.EOF != next)
            this._addToFormat(numFormat_Text, next);
    },
    _ReadBracket : function()
    {
        var sBracket = "";
        while(true)
        {
            var next = this._readChar();
            if(this.EOF == next || "]" == next)
                break;
            else
            {
                sBracket += next;
            }
        }
		var oFormatObjBracket = new FormatObjBracket(sBracket);
		if(null != oFormatObjBracket.operator)
			this.ComporationOperator = oFormatObjBracket;
        this._addToFormat2(oFormatObjBracket);
    },
    _ReadAmPm : function(next)
    {
		if ("A" === next || "a" === next) {
			let ampm = "AM/PM";
			if (ampm.substring(1) === this._GetText(ampm.length - 1).toUpperCase()) {
				this._addToFormat2(new FormatObj(numFormat_AmPm));
				this.bTimePeriod = true;
				this.bDateTime = true;
				this._skip(ampm.length - 1);
				return true;
			}
		}
		if ("上" === next) {
			let ampm = "上午/下午";
			if (ampm.substring(1) === this._GetText(ampm.length - 1).toUpperCase()) {
				this._addToFormat2(new FormatObj(numFormat_AmPm));
				this.bTimePeriod = true;
				this.bDateTime = true;
				this._skip(ampm.length - 1);
				return true;
			}
		}
		return false;
    },
	_ReadAmPmPDF : function(next)
    {
		let bAmPm = true;
		let nttCount = 1;
        while(true)
        {
            next = this._readChar();
            if(this.EOF == next)
                break;
            else if ("t" == next)
            {
				nttCount++;
            }
            else
            {
				// если больше двух tt не добавляем am/pm
				if (nttCount > 2) {
					bAmPm = false;
				}

				this._skip(-1);
				break;
            }
        }
        if(bAmPm == true)
        {
            this._addToFormat2(new FormatObj(numFormat_AmPm));
            this.bTimePeriod = true;
            this.bDateTime = true;
        }
    },
    _parseFormat : function(digitSpaceSymbol, useLocaleFormat)
    {
        var sGeneral;
        var DecimalSeparator;
        var GroupSeparator;
        var TimeSeparator;
        var Year;
        var Month;
        var Day;
        var Hour;
        var year;
        var month;
        var day;
        var hour;
        var Minute;
        var minute;
        var Second;
        var second;
		var dayOfWeek;
		if (useLocaleFormat) {
			sGeneral = LocaleFormatSymbol['general'].toLowerCase();
			DecimalSeparator = g_oDefaultCultureInfo.NumberDecimalSeparator;
			TimeSeparator = g_oDefaultCultureInfo.TimeSeparator;
			GroupSeparator = g_oDefaultCultureInfo.NumberGroupSeparator;
			Year = LocaleFormatSymbol['Y'];
			year = LocaleFormatSymbol['y'];
			Month = LocaleFormatSymbol['M'];
			month = LocaleFormatSymbol['m'];
			Day = LocaleFormatSymbol['D'];
			day = LocaleFormatSymbol['d'];
			Hour = LocaleFormatSymbol['H'];
			hour = LocaleFormatSymbol['h'];
			Minute = LocaleFormatSymbol['Minute'];
			minute = LocaleFormatSymbol['minute'];
			Second = LocaleFormatSymbol['S'];
			second = LocaleFormatSymbol['s'];
			dayOfWeek = LocaleFormatSymbol['a'];
		} else {
			sGeneral = AscCommon.g_cGeneralFormat.toLowerCase();
			DecimalSeparator = gc_sFormatDecimalPoint;
			TimeSeparator = ':';
			GroupSeparator = gc_sFormatThousandSeparator;
			Year = 'Y';
			year = 'y';
			Month = 'M';
			month = 'm';
			Day = 'D';
			day = 'd';
			Hour = 'H';
			hour = 'h';
			Minute = 'M';
			minute = 'm';
			Second = 'S';
			second = 's';
			dayOfWeek = 'a';
		}
        var sGeneralFirst = sGeneral[0];
        this.bGeneralChart = true;
        while(true)
        {
            var next = this._readChar();
            var bNoFormat = false;
            if(this.EOF == next)
                break;
            else if("[" == next)
                this._ReadBracket();
            else if("\"" == next)
                this._ReadText("\"");
            else if("\\" == next)
                this._ReadChar();
            else if("%" == next)
            {
                this._addToFormat(numFormat_Percent);
            }
            else if(TimeSeparator == next)
            {
                this._addToFormat(numFormat_TimeSeparator);
            }
            else if('0' === next)
            {
                this._addToFormat(numFormat_Digit, 0);
            }
            else if("#" == next)
            {
                this._addToFormat(numFormat_DigitNoDisp);
            }
            else if(digitSpaceSymbol == next)
            {
                this._addToFormat(numFormat_DigitSpace);
            }
            else if(DecimalSeparator == next)
            {
                this._addToFormat(numFormat_DecimalPoint);
            }
            else if("/" == next)
            {
                this._addToFormat2(new FormatObjDecimalFrac([], []));
            }
            else if(GroupSeparator == next)
            {
                this._addToFormat(numFormat_Thousand, 1);
            }
            else if("$" == next || "+" == next || "-" == next || "(" == next || ")" == next || " " == next)
            {
                this._addToFormat(numFormat_Text, next);
            }
            else if (sGeneralFirst === next.toLowerCase() &&
                sGeneral === (next + this._GetText(sGeneral.length - 1)).toLowerCase()) {
                this._addToFormat(numFormat_General);
                this._skip(sGeneral.length - 1);
            }
			else if (this._ReadAmPm(next))
			{

			}
            else if("E" == next || "e" == next)
            {
                var nextnext = this._readChar();
                if(this.EOF != nextnext && "+" == nextnext || "-" == nextnext)
                {
                    var sign = ("+" == nextnext) ? SignType.Positive : SignType.Negative;
                    this._addToFormat2(new FormatObjScientific(next, "", sign));
                }
            }
            else if("*" == next)
            {
                var nextnext = this._readChar();
                if(this.EOF != nextnext)
                    this._addToFormat(numFormat_Repeat, nextnext);
            }
            else if("_" == next)
            {
                var nextnext = this._readChar();
                if(this.EOF != nextnext)
                    this._addToFormat(numFormat_Skip, nextnext);
            }
            else if("@" == next)
            {
                this._addToFormat(numFormat_TextPlaceholder);
            }
            else if(Year == next || year == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Year, 1, false));
            }
            else if(Month == next || month == next)
            {
                if (Month === Minute) {
                    this._addToFormat2(new FormatObjDateVal(numFormat_MonthMinute, 1, false));
                } else {
                    this._addToFormat2(new FormatObjDateVal(numFormat_Month, 1, false));
                }
            }
            else if(Day == next || day == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Day, 1, false));
            }
            else if(Hour == next || hour == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Hour, 1, false));
            }
            else if(Minute == next || minute == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Minute, 1, false));
            }
            else if(Second == next || second == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Second, 1, false));
            }
			else if (dayOfWeek == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_DayOfWeek, 1, false));
			}
            else {
                bNoFormat = true;
                this._addToFormat(numFormat_Text, next);
            }
            if (!bNoFormat)
                this.bGeneralChart = false;
        }
        return true;
    },
    _parseFormatWordDateTime : function()
    {
        while(true)
        {
            var next = this._readChar();
			if(this.EOF == next)
				break;
			else if("\'" == next)
				this._ReadText("\'");
			else if (this._ReadAmPm(next))
			{

			}
			else if("Y" == next || "y" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Year, 1, false));
			}
			else if("M" == next || "m" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_MonthMinute, 1, false));
			}
			else if("D" == next || "d" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Day, 1, false));
			}
			else if("H" == next || "h" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Hour, 1, false));
			}
			else if("S" == next || "s" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Second, 1, false));
			}
			else if ("a" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_DayOfWeek, 1, false));
			}
			else {
					this._addToFormat(numFormat_Text, next);
			}
        }
        return true;
    },
	_parseFormatPDFDateTime : function()
    {
        while(true)
        {
            var next = this._readChar();
			if(this.EOF == next)
				break;
			else if("\'" == next)
				this._ReadText("\'");
			else if ("y" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Year, 1, false));
			}
			else if ("m" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Month, 1, false));
			}
			else if ("M" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Minute, 1, false));
			}
			else if ("d" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Day, 1, false));
			}
			else if ("h" == next || "H" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Hour, 1, false));
			}
			else if ("s" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Second, 1, false));
			}
			else if ("t" == next) {
				this._ReadAmPmPDF(next);
			}
			else {
				this._addToFormat(numFormat_Text, next);
			}
        }
        return true;
    },
	_parseFormatWordNumeric : function(digitSpaceSymbol)
	{
		while(true)
		{
			var next = this._readChar();
			if (this.EOF == next) {
				break;
			} else if ("\'" === next) {
				this._ReadText("\'");
			} else if ('0' === next) {
				this._addToFormat(numFormat_Digit, 0);
			} else if (digitSpaceSymbol === next) {
				this._addToFormat(numFormat_DigitSpace);
			} else if ('x' === next || 'X' === next) {
				this._addToFormat(numFormat_DigitDrop);
			} else if (gc_sFormatDecimalPoint === next) {
				this._addToFormat(numFormat_DecimalPoint);
			} else if (gc_sFormatThousandSeparator === next) {
				this._addToFormat(numFormat_Thousand, 1);
			} else if ('+' === next) {
				this._addToFormat(numFormat_Plus);
			} else if ('-' === next) {
				this._addToFormat(numFormat_Minus);
			} else {
				this._addToFormat(numFormat_Text, next);
			}
		}
		return true;
	},
	_isDigitType: function(type) {
		return numFormat_Digit === type || numFormat_DigitNoDisp === type || numFormat_DigitSpace === type ||
			numFormat_DigitDrop === type;
	},
    _prepareFormat : function()
    {
        //Color
		for(var i = 0, length = this.aRawFormat.length; i < length; ++i)
        {
            var oCurItem = this.aRawFormat[i];
            if(numFormat_Bracket == oCurItem.type && null != oCurItem.color)
                this.Color = oCurItem.color;
        }
        this.bRepeat = false;
        var nFormatLength = this.aRawFormat.length;

        //Группируем несколько элемнтов подряд в один спецсимвол
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_Repeat == item.type)
            {
                //Оставляем только последний numFormat_Repeat
                if(false == this.bRepeat)
                    this.bRepeat = true;
                else
                {
                    this.aRawFormat.splice(i, 1);
                    nFormatLength--;
                }
            }
            else if(numFormat_Bracket == item.type)
            {
                //Разруливаем [hhh]
                var oNewObj = item.dataObj;
                if(null != oNewObj)
                {
                    this.aRawFormat.splice(i, 1, oNewObj);
                    this.bDateTime = true;
                    if(numFormat_Hour == oNewObj.type || numFormat_Minute == oNewObj.type || numFormat_Second == oNewObj.type || numFormat_Milliseconds == oNewObj.type)
                        this.bTime = true;
                    else if (numFormat_Year == oNewObj.type || numFormat_Month == oNewObj.type || numFormat_Day == oNewObj.type) {
                        this.bDate = true;
                        if (numFormat_Day == oNewObj.type)
                            this.bDay = true;
                    }
                }
            }
            else if(numFormat_Year == item.type || numFormat_MonthMinute == item.type || numFormat_Month == item.type || numFormat_Day == item.type || numFormat_Hour == item.type || numFormat_Minute == item.type || numFormat_Second == item.type || numFormat_Thousand == item.type ||
				numFormat_DayOfWeek == item.type)
            {
                //Собираем в одно целое последовательности hhh
                var nStartType = item.type;
                var nEndIndex = i;
                for(var j = i + 1; j < nFormatLength; ++j)
                {
                    if(nStartType == this.aRawFormat[j].type)
                        nEndIndex = j;
                    else
                        break;
                }
                if(i != nEndIndex)
                {
                    item.val = nEndIndex - i + 1;
                    var nDelCount = item.val - 1;
                    this.aRawFormat.splice(i + 1, nDelCount);
                    nFormatLength -= nDelCount;
                }
                if(numFormat_Thousand != item.type)
                {
                    this.bDateTime = true;
                    if(numFormat_Hour == item.type || numFormat_Minute == item.type || numFormat_Second == item.type || numFormat_Milliseconds == item.type)
                        this.bTime = true;
                    else if (numFormat_Year == item.type || numFormat_Month == item.type || numFormat_Day == item.type) {
                        this.bDate = true;
                        if (numFormat_Day == item.type)
                            this.bDay = true;
                    }
                }
            }
            else if(numFormat_Scientific == item.type)
            {
                var bAsText = false;
                if(true == this.bScientific)
                {
                    bAsText = true;
                }
                else
                {
                    var aDigitArray = [];
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var nextItem = this.aRawFormat[j];
                        if(this._isDigitType(nextItem.type))
                            aDigitArray.push(nextItem);
                    }
                    if(aDigitArray.length > 0)
                    {
                        item.format = aDigitArray;
                        this.bScientific = true;
                    }
                    else
                        bAsText = true;
                }
                if(false != bAsText)
                {
                    //заменяем на текст
                    item.type = numFormat_Text;
                    item.val = item.val + "+";
                }
            }
            else if(numFormat_DecimalFrac == item.type)
            {
                var bValid = false;
                //собираем правую и левую часть дроби
                var nLeft = i;
                for(var j = i - 1; j >= 0; --j)
                {
                    var subitem = this.aRawFormat[j];
                    if(this._isDigitType(subitem.type))
                        nLeft = j;
                    else
                        break;
                }
                var nRight = i;
                if(nLeft < i)
                {
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var subitem = this.aRawFormat[j];
                        if(this._isDigitType(subitem.type) || (numFormat_Text === subitem.type && '0' <= subitem.val && subitem.val <= '9'))
                            nRight = j;
                        else
                            break;
                    }
                    if(nRight > i)
                    {
                        bValid = true;
                        item.aRight = this.aRawFormat.splice(i + 1, nRight - i);
                        item.aLeft = this.aRawFormat.splice(nLeft, i - nLeft);
                        nFormatLength -= nRight - nLeft;
                        i -= i - nLeft;
                        this.bSlash = true;

                        var flag = (item.aRight.length > 0) && (item.aRight[0].type == numFormat_Digit || item.aRight[0].type == numFormat_Text) && (parseInt(item.aRight[0].val) > 0);
                        if(flag)
                        {
                            var rPart = 0;
                            for(var j = 0; j< item.aRight.length; j++)
                            {
                                if(item.aRight[j].type == numFormat_Digit || item.aRight[j].type == numFormat_Text)
                                    rPart = rPart*10 + parseInt(item.aRight[j].val);
                                else
                                {
                                    bValid = false;
                                    this.bSlash = false;
                                    break;
                                }
                            }
                            if(bValid == true)
                            {
                                item.aRight = [];
                                item.aRight.push(new FormatObj(numFormat_Digit, rPart));
                                item.bNumRight = true;
                            }
                        }
                    }

                }

                if(false == bValid)
                {
                    item.type = numFormat_DateSeparator;
                }
            }
        }
        
        var nReadState = FormatStates.Decimal;
        var bDecimal = true;
        nFormatLength = this.aRawFormat.length;
        //Разруливаем конфликтные ситуации, выставляем значения свойств
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_DecimalPoint == item.type)
            {
                //миллисекунды
                //Если после DecimalPoint идут numFormat_Digit, и есть формат для даты времени, то это миллисекунды
                if(this.bDateTime)
                {
                    var nStartIndex = i;
                    var nEndIndex = nStartIndex;
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var subItem = this.aRawFormat[j];
                        if(numFormat_Digit == subItem.type)
                            nEndIndex = j;
                        else
                            break;
                    }
                    if(nStartIndex < nEndIndex)
                    {
                        var nDigCount = nEndIndex - nStartIndex;
                        var oNewItem = new FormatObjDateVal(numFormat_Milliseconds, nDigCount, false);
                        var nDelCount = nDigCount;
                        oNewItem.format = this.aRawFormat.splice(i + 1, nDelCount, oNewItem);
                        nFormatLength -= (nDigCount - 1);
                        i++;
                        this.bMillisec = true;

                    }
                    //преобразуем в текст все последующие
                    item.type = numFormat_DecimalPointText;
                    item.val = null;
                }
                else if(FormatStates.Decimal == nReadState)
                    nReadState = FormatStates.Frac;
            }
            else if(numFormat_MonthMinute == item.type)
            {
                //Разрешаем конфликты numFormat_MonthMinute
                var bRightCond = false;
                //ищем вперед первый элемент с типом datetime 
                for(var j = i + 1; j < nFormatLength; ++j)
                {
                    var subItem = this.aRawFormat[j];
                    if(numFormat_Year == subItem.type || numFormat_Month == subItem.type || numFormat_Day == subItem.type || numFormat_MonthMinute == subItem.type ||
                    numFormat_Hour == subItem.type || numFormat_Minute == subItem.type || numFormat_Second == subItem.type || numFormat_Milliseconds == subItem.type)
                    {
                        if(numFormat_Second == subItem.type)
                            bRightCond = true;
                        break;
                    }
                }
                var bLeftCond = false;
                if(false == bRightCond)
                {
                    //ищем назад первый элемент с типом hh или ss
                    var bFindSec = false;//чтобы разрулить случай mm:ss:mm должно быть Минуты:Секунды:Месяцы
                    for(var j = i - 1; j >= 0; --j)
                    {
                        var subItem = this.aRawFormat[j];
                        
                        if(numFormat_Hour == subItem.type)
                        {
                            bLeftCond = true;
                            break;
                        }
                        else if(numFormat_Second == subItem.type)
                        {
                            //продолжаем смотреть дальше, пока не встретиться следующий date time обьект
                            bFindSec = true;
                        }
                        else if(numFormat_Minute == subItem.type || numFormat_Month == subItem.type || numFormat_MonthMinute == subItem.type)
                        {
                            if(true == bFindSec && numFormat_Minute == subItem.type)
                                bFindSec = false;
                            break;
                        }
                        else if(numFormat_Year == subItem.type || numFormat_Day == subItem.type || numFormat_Hour == subItem.type || numFormat_Second == subItem.type || numFormat_Milliseconds == subItem.type)
                        {
                            if(true == bFindSec)
                                break;
                        }
                    }
                    if(true == bFindSec)
                        bLeftCond = true;
                }
                
                if((true == bLeftCond || true == bRightCond) && item.val <= 2)
				{
                    item.type = numFormat_Minute;
					this.bTime = true;
				}
                else
				{
                    item.type = numFormat_Month;
					this.bDate = true;
				}
            }
            else if(numFormat_Percent == item.type)
            {
                this.nPercent++;
                //заменяем на текст
                item.type = numFormat_Text;
                item.val = "%";
            }
            else if(numFormat_Thousand == item.type)
            {
                var isPrevDigit = i > 0 && this._isDigitType(this.aRawFormat[i - 1].type);
                var isPrevDecimalPoint = i > 0 && numFormat_DecimalPoint === this.aRawFormat[i - 1].type;
                var isNextDigit = i + 1 < nFormatLength && this._isDigitType(this.aRawFormat[i + 1].type);
                if (isPrevDigit && isNextDigit) {
                    if(FormatStates.Decimal == nReadState) {
                        this.bThousandSep = true;
                    }
                } else if (isPrevDigit || isPrevDecimalPoint) {
                    this.nThousandScale = item.val;
                } else {
                    item.type = numFormat_ThousandText;
                }
            }
            else if(this._isDigitType(item.type))
            {
                this.nThousandScale = 0;
                if(FormatStates.Decimal == nReadState)
                {
                    this.aDecFormat.push(item);

                    if(this.bSlash === true)
                        this.bWhole = true;
                }
                else if(FormatStates.Frac == nReadState)
                    this.aFracFormat.push(item);

            }
            else if(numFormat_Scientific == item.type)
                nReadState = FormatStates.Scientific;
            else if(numFormat_TextPlaceholder == item.type)
            {
                this.bTextFormat = true;
            }
        }
        return true;
    },
	_prepareFormatDatePDF : function()
    {
		var nFormatLength = this.aRawFormat.length;
        //Группируем несколько элемнтов подряд в один спецсимвол
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_Year == item.type || numFormat_Month == item.type || numFormat_Day == item.type)
            {
                //Удаляем итемы у которых val > 4 (для года удаляем если "yyy")
				if(item.val === 3 && numFormat_Year == item.type)
                {
                    this.aRawFormat.splice(i, 1);
					nFormatLength -= 1;
                }
                if(item.val > 4)
                {
                    this.aRawFormat.splice(i, 1);
					nFormatLength -= 1;
                }
            }
			else if(numFormat_Hour == item.type || numFormat_Minute == item.type || numFormat_Second == item.type)
            {
				//Удаляем итемы у которых val > 2
                if(item.val > 2)
                {
                    this.aRawFormat.splice(i, 1);
					nFormatLength -= 1;
                }
            }
        }
    },
	_calsScientific : function(nDecLen, nRealExp)
	{
		var nKoef = 0;
		if(true == this.bThousandSep)
			nKoef = 4;
		if(nDecLen > nKoef)
			nKoef = nDecLen;
		if(nRealExp > 0 && nKoef > 0)
		{
			var nTemp = nRealExp % nKoef;
			if(0 == nTemp)
				nTemp = nKoef;
			nKoef = nTemp;
		}
		return nKoef;
	},
	_parseNumber : function(number, aDecFormat, nFracLen, nValType)
    {
        var res = {bDigit: false, dec: 0, frac: 0, fraction: 0, exponent: 0, exponentFrac: 0, scientific: 0, sign: SignType.Positive, date: {}};
        if(CellValueType.String != nValType)
            res.bDigit = (number == number - 0);
        if(res.bDigit)
        {
			var numberAbs = Math.abs(number);
			res.fraction = numberAbs - Math.floor(numberAbs);
			//Округляем
			var parts = getNumberParts(number);
			res.sign = parts.sign;
			var nRealExp = gc_nMaxDigCount + parts.exponent;//nRealExp == 0, при 0,123
			if(SignType.Null != parts.sign)
			{
				if(true == this.bScientific)
				{
					var nKoef = this._calsScientific(aDecFormat.length, nRealExp);
					res.scientific = nRealExp - nKoef;
					nRealExp = nKoef;
				}
				else
				{
					//Percent
					for(var i = 0; i < this.nPercent; ++i)
						nRealExp += 2;
					//Thousands separator
					for(var i = 0; i < this.nThousandScale; ++i)
						nRealExp -= 3;		
				}
				//округляем после операций которые могут изменить nRealExp
				if(false == this.bSlash)
				{
					var nOldRealExp = nRealExp;
					parts = getNumberParts(round10(parts.mantissa, nFracLen + nRealExp - gc_nMaxDigCount, nFracLen));
					if(SignType.Null != parts.sign)
					{
						nRealExp = gc_nMaxDigCount + parts.exponent;
						if(nOldRealExp != nRealExp && true == this.bScientific)
						{
							var nKoef = this._calsScientific(aDecFormat.length, nRealExp);
							res.scientific += nRealExp - nOldRealExp;
							nRealExp = nKoef;
						}
					}
				}
				res.exponent = nRealExp;
				res.exponentFrac = nRealExp;
				if(nRealExp > 0 && nRealExp < gc_nMaxDigCount)
				{
					var sNumber = parts.mantissa.toString();
					var nExponentFrac = 0;
					for(var i = nRealExp, length = sNumber.length; i < length; ++i)
					{
						if("0" == sNumber[i])
							nExponentFrac++;
						else
							break;
					}
					if(nRealExp + nExponentFrac < sNumber.length)
						res.exponentFrac = - nExponentFrac;
				}
				if(SignType.Null != parts.sign)
				{
					if(nRealExp <= 0)
					{
						if(this.bSlash == true)
						{
							res.dec = 0;
							res.frac = parts.mantissa;
						}
						else
						{
							if(nFracLen > 0)
							{
								res.dec = 0;
								res.frac = 0;
								if(nFracLen + nRealExp > 0)
								{
									var sTemp = parts.mantissa.toString();
									res.frac = sTemp.substring(0, nFracLen + nRealExp) - 0;
								}
							}
							else
							{
								res.dec = 0;
								res.frac = 0;
							}
						}
					}
					else if(nRealExp >= gc_nMaxDigCount)
					{
						res.dec = parts.mantissa;
						res.frac = 0;
					}
					else
					{
						var sTemp = parts.mantissa.toString();
						if(this.bSlash == true)
						{
							res.dec = sTemp.substring(0, nRealExp) - 0;
							if(nRealExp < sTemp.length)
								res.frac = sTemp.substring(nRealExp) - 0;
							else
								res.frac = 0;
						}
						else
						{
							if(nFracLen > 0 )
							{
								res.dec = sTemp.substring(0, nRealExp) - 0;
								res.frac = 0;
								var nStart = nRealExp;
								var nEnd = nRealExp + nFracLen;
								if(nStart < sTemp.length)
									res.frac = sTemp.substring(nStart, nEnd) - 0;
							}
							else
							{
								res.dec = sTemp.substring(0, nRealExp) - 0;
								res.frac = 0;
							}
						}
					}
				}
				if(0 == res.frac && 0 == res.dec && false === this.bDateTime)
					res.sign = SignType.Null;
			}
            //После округления может получиться ноль,
            //но не стала перестаскивать проверку на знак сюда, т.к. для округления нужно неотриц число

            if(this.bDateTime === true)
				res.date = this.parseDate(number);
        }
        return res;
    },
	_parseNumberForPDFDate : function(number) {
		let oDateTmp = new Date();
		oDateTmp.setTime(number * (86400 * 1000));
	 
		return {
			date: {
				d:			oDateTmp.getDate(),
				dayWeek:	oDateTmp.getDay(),
				hour:		oDateTmp.getHours(),
				min:		oDateTmp.getMinutes(),
				month:		oDateTmp.getMonth(),
				ms:			0,
				//ms:			oDateTmp.getMilliseconds(),
				sec:		oDateTmp.getSeconds(),
				year:		oDateTmp.getFullYear()
			}
		}
	},
	parseDate : function(number)
	{
        var d = {val: 0, coeff: 1}, h = {val: 0, coeff: 24},
            min = {val: 0, coeff: 60}, s = {val: 0, coeff: 60}, ms = {val: 0, coeff: 1000};
        //number is negative in case of bDate1904
        var numberAbs = this.formatType == AscCommon.NumFormatType.PDFFormDate ? number : Math.abs(number);
        var tmp = numberAbs;
        var ttimes = [d, h, min, s, ms];
        for(var i = 0; i < 4; i++)
        {
            var v = tmp*ttimes[i].coeff;
            ttimes[i].val = Math.floor(v);
            tmp = v - ttimes[i].val;
        }
        ms.val = Math.round(tmp*1000);
        for(i = 4; i > 0 && (ttimes[i].val === ttimes[i].coeff); i--)
        {
            ttimes[i].val = 0;
            ttimes[i-1].val++;
        }
        var stDate, day, month, year, dayWeek;
		if(AscCommon.bDate1904)
		{
			stDate = new Date(Date.UTC(1904,0,1,0,0,0));
			if(d.val)
				stDate.setUTCDate( stDate.getUTCDate() + d.val );
			day = stDate.getUTCDate();
			dayWeek = stDate.getUTCDay();
			month = stDate.getUTCMonth();
			year = stDate.getUTCFullYear();
		}
		else
		{
			if (60 <= numberAbs && numberAbs < 61)
			{
				day = 29;
				month = 1;
				year = 1900;
				dayWeek = 3;
			}
			else if (0 <= numberAbs && numberAbs < 1)
			{
				//TODO необходимо использовать cDate везде
				stDate = new Asc.cDate(Date.UTC(1899,11,31,0,0,0));
				day = stDate.getUTCDate();
				dayWeek = ( stDate.getUTCDay() > 0) ? stDate.getUTCDay() - 1 : 6;
				month = stDate.getUTCMonth();
				year = stDate.getUTCFullYear();
			}
			else if(numberAbs < 60 && number > 0)
			{
				stDate = new Date(Date.UTC(1899,11,31,0,0,0));
				if(d.val)
				// setUTCDate doesn't consider the transition from 1899 to 1900 when adding d.val
					stDate.setUTCDate( stDate.getUTCDate() + d.val );
				day = stDate.getUTCDate();
				dayWeek = ( stDate.getUTCDay() > 0) ? stDate.getUTCDay() - 1 : 6;
				month = stDate.getUTCMonth();
				year = stDate.getUTCFullYear();
			}
			else
			{
				stDate = new Date(Date.UTC(1899,11,30,0,0,0));
				if(d.val)
					stDate.setUTCDate( stDate.getUTCDate() + d.val );
				day = stDate.getUTCDate();
				dayWeek = stDate.getUTCDay();
				month = stDate.getUTCMonth();
				year = stDate.getUTCFullYear();
			}
		}
        return {d: day, month: month, year: year, dayWeek: dayWeek, hour: h.val, min: min.val, sec: s.val, ms: ms.val, countDay: d.val };
	},
	_FormatNumber: function (number, exponent, format, nReadState, cultureInfo, opt_forceNull)
	{
        var aRes = [];
        var nFormatLen = format.length;
        if(nFormatLen > 0)
        {
            if(FormatStates.Frac != nReadState && FormatStates.SlashFrac != nReadState)
            {
				var sNumber = number + "";
				var nNumberLen = sNumber.length;
				//для бага Bug 14325 - В загруженной таблице число с 30 знаками после разделителя отображается неправильно.
				//например число "1.23456789123456e+23" и формат "0.000000000000000000000000000000"
				if(exponent > nNumberLen)
				{
					for(var i = 0; i < exponent - nNumberLen; ++i)
						sNumber += "0";
					nNumberLen = sNumber.length;
				}
                var bIsNUll = false;
                if("0" == sNumber && !opt_forceNull)
                    bIsNUll = true;
                //выравниваем длину
                if(nNumberLen > nFormatLen)
                {
                    if(false === bIsNUll)
                    {
						var item = format.shift();
						if (numFormat_DigitDrop !== item.type) {
							var nSplitIndex = nNumberLen - nFormatLen + 1;
							aRes.push(new FormatObj(numFormat_Text, sNumber.slice(0, nSplitIndex)));
							sNumber = sNumber.substring(nSplitIndex);
						} else {
							sNumber = sNumber.substring(nNumberLen - nFormatLen);
						}
                    }
                }
                else if(nNumberLen < nFormatLen)
                {
                    //просто копируем, здесь будут только нули и пропуски
                    for(var i = 0, length = nFormatLen - nNumberLen; i < length; ++i)
                    {
                        var item = format.shift();
                        aRes.push(new FormatObj(item.type));
                    }
                }
                //просто заполняем текстом
                for(var i = 0, length = sNumber.length; i < length; ++i)
                {
                    var sCurNumber = sNumber[i];
					var numFormat = numFormat_Text;
                    var item = format.shift();
                    if(true == bIsNUll && null != item && FormatStates.Scientific != nReadState)
					{
						if(numFormat_DigitNoDisp == item.type)
							sCurNumber = "";
						else if(numFormat_DigitSpace == item.type)
						{
							numFormat = numFormat_DigitSpace;
							sCurNumber = null;
						}
					}
                    aRes.push(new FormatObj(numFormat, sCurNumber));
                }
                
                //Вставляем разделители 
                if(true == this.bThousandSep && FormatStates.Slash != nReadState)
                {
					var sThousandSep = cultureInfo.NumberGroupSeparator;
					var aGroupSize = cultureInfo.NumberGroupSizes;
					var nCurGroupIndex = 0;
					var nCurGroupSize = 0;
					if (nCurGroupIndex < aGroupSize.length)
					    nCurGroupSize = aGroupSize[nCurGroupIndex++];
					else
					    nCurGroupSize = 0;
                    var nIndex = 0;
                    for(var i = aRes.length - 1; i >= 0; --i)
                    {
                        var item = aRes[i];
                        if(numFormat_Text == item.type)
                        {
                            var aNewText = [];
                            var nTextLength = item.val.length;
                            for(var j = nTextLength - 1; j >= 0; --j)
                            {
                                if (nCurGroupSize == nIndex)
                                {
                                    aNewText.push(sThousandSep);
                                    nTextLength++;
                                }
                                aNewText.push(item.val[j]);
                                if(0 != j)
                                {
                                    nIndex++;
                                    if (nCurGroupSize + 1 == nIndex) {
                                        nIndex = 1;
                                        if (nCurGroupIndex < aGroupSize.length)
                                            nCurGroupSize = aGroupSize[nCurGroupIndex++];
                                    }
                                }
                            }
                            if(nTextLength > 1)
                                aNewText.reverse();
                            item.val = aNewText.join("");
                        }
                        else if(numFormat_DigitNoDisp != item.type)
                        {
                            //не добавляем пробел только перед numFormat_DigitNoDisp
                            if (nCurGroupSize == nIndex)
                            {
                                item.val = sThousandSep;
                                aRes[i] = item;
                            }
                        }
                        nIndex++;
                        if (nCurGroupSize + 1 == nIndex) {
                            nIndex = 1;
                            if (nCurGroupIndex < aGroupSize.length)
                                nCurGroupSize = aGroupSize[nCurGroupIndex++];
                        }
                    }
                }
            }
            else
            {
				var val = number;
				var exp = exponent;
                //Считаем количество нулей в начале
                var nStartNulls = 0;
				if(exp < 0)
					nStartNulls = Math.abs(exp);
                var sNumber = val.toString();
                var nNumberLen = sNumber.length;
				//удаляем 0 на конце
				var nLastNoNull = nNumberLen;
                for(var i = nNumberLen - 1; i >= 0; --i)
                {
					if ("0" != sNumber[i])
						break;
					nLastNoNull = i;
				}
				if (nLastNoNull < nNumberLen && (FormatStates.SlashFrac != nReadState || 0 == nLastNoNull)) {
					sNumber = sNumber.substring(0, nLastNoNull);
					nNumberLen = sNumber.length;
				}
                //заполняем первые нули
                for(var i = 0; i < nStartNulls; ++i)
                    aRes.push(new FormatObj(numFormat_Text, "0"));
                //просто заполняем текстом
                for(var i = 0, length = nNumberLen; i < length; ++i)
                    aRes.push(new FormatObj(numFormat_Text, sNumber[i]));
                //просто копируем, здесь будут только нули и пропуски
                for(var i = nNumberLen + nStartNulls; i < nFormatLen; ++i)
                {
                    var item = format[i];
                    aRes.push(new FormatObj(item.type));
                }
            }
        }
        return aRes;
    },
	_replaceDBNumDigit: function (val) {
		//todo DBNum 1-4
		if (1 !== this.DBNum) {
			return val;
		}
		let locale = Asc.g_oLcidIdToNameMap[this.LCID];
		if (!locale) {
			return val;
		}
		locale = locale.substring(0, 2);
		if ('zh' === locale || 'ja' === locale || 'ko' === locale) {
			let dBNumVal = '';
			for (let j = 0; j < val.length; ++j) {
				if ('0' <= val[j] && val[j] <= '9') {
					dBNumVal += dBNum1Numbers[val[j] - '0'];
				} else {
					dBNumVal += val[j];
				}
			}
			val = dBNumVal;
		}
		return val;
	},
    _AddDigItem : function(res, oCurText, item)
    {
        if(numFormat_Text == item.type)
            oCurText.text += item.val;
        else if(numFormat_Digit == item.type)
        {
            //text.val может заполниться в Thousand
            oCurText.text += "0";
            if(null != item.val)
                oCurText.text += item.val;
        }
        else if(numFormat_DigitNoDisp == item.type)
        {
            oCurText.text += "";
            if(null != item.val)
                oCurText.text += item.val;
        }
        else if(numFormat_DigitSpace == item.type || numFormat_DigitDrop == item.type)
        {
            var oNewFont = new AscCommonExcel.Font();
			oNewFont.skip = true;
            this._CommitText(res, oCurText, "0", oNewFont);
            if(null != item.val)
                oCurText.text += item.val;
        }
    },
    _ZeroPad: function(n)
    {
        return (n < 10) ? "0" + n : n;
    },
    _CommitText: function(res, oCurText, textVal, format)
    {
        if(null != oCurText && oCurText.text.length > 0)
        {
            this._CommitText(res, null, oCurText.text, null);
            oCurText.text = "";
        }
        if(null != textVal && textVal.length > 0)
        {
            var length = res.length;
            var prev = null;
            if(length > 0)
                prev = res[length - 1];
            if(-1 != this.Color)
            {
                if(null == format)
                    format = new AscCommonExcel.Font();
                format.c = new AscCommonExcel.RgbColor(this.Color);
            }
            if(null != prev && ((null == prev.format && null == format) || (null != prev.format && null != format && format.isEqual(prev.format))))
            {
                prev.text += textVal;
            }
            else
            {
                if(null == format)
                    prev = {text: textVal};
                else
                    prev = {text: textVal, format: format};
                res.push(prev);
            }
        }
    },
    setFormat: function(format, cultureInfo, formatType, useLocaleFormat) {
		if (null == cultureInfo) {
            cultureInfo = g_oDefaultCultureInfo;
        }
        this.formatString = format;
        this.length = this.formatString.length;
        //string -> tokens
		if (NumFormatType.WordFieldDate === formatType) {
			this.valid = this._parseFormatWordDateTime();
		} else if (NumFormatType.PDFFormDate === formatType) {
			this.valid = this._parseFormatPDFDateTime();
		} else if (NumFormatType.WordFieldNumeric === formatType) {
			this.valid = this._parseFormatWordNumeric("#");
		} else {
			this.valid = this._parseFormat("?", useLocaleFormat);
		}
        if (true == this.valid) {
            //prepare tokens
            // this.valid = formatType != NumFormatType.PDFFormDate ? this._prepareFormat() : this._prepareFormatPDF();
            this.valid = this._prepareFormat();
            if (this.valid) {
                //additional prepare
                var aCurrencySymbols = ["$", "€", "£", "¥", "р.", cultureInfo.CurrencySymbol];
                var sText = "";
                for (var i = 0, length = this.aRawFormat.length; i < length; ++i) {
                    var item = this.aRawFormat[i];
                    if (numFormat_Text == item.type) {
                        sText += item.val;
                    } else if (numFormat_Bracket == item.type) {
						let dbnum = item.val.match(/DBNum(\d)/);
						if (dbnum) {
							this.DBNum = parseInt(dbnum[1]);
						} else {
							if (null != item.CurrencyString) {
								this.bCurrency = true;
								this.CurrencyString = item.CurrencyString;
								sText += item.CurrencyString;
							}
							if (null != item.Lid) {
								//Excel sometimes add 0x10000(0x442 and 0x10442)
								this.LCID = parseInt(item.Lid, 16) & 0xFFFF;
							}
						}
                    }
                    else if (numFormat_DecimalPoint == item.type) {
                        sText += gc_sFormatDecimalPoint;
                    } else if (numFormat_DecimalPointText == item.type) {
                        sText += gc_sFormatDecimalPoint;
                    }
                }
                if ("" != sText) {
                    for (var i = 0, length = aCurrencySymbols.length; i < length; ++i) {
                        if (-1 != sText.indexOf(aCurrencySymbols[i])) {
                            this.bCurrency = true;
                            break;
                        }
                    }
                }
                    }
                }
        return this.valid;
    },
    isInvalidDateValue : function(number)
    {
        return (number == number - 0) && ((number < 0 && !AscCommon.bDate1904) || number > 2958465.9999884);
    },
    _applyGeneralFormat: function(number, nValType, dDigitsCount, bChart, cultureInfo){
        var res = null;
        //todo fIsFitMeasurer and decrease dDigitsCount by other format tokens
        var sGeneral = DecodeGeneralFormat(number, nValType, dDigitsCount);
        if (null != sGeneral) {
            var numFormat = oNumFormatCache.get(sGeneral);
            if (null != numFormat) {
                res = numFormat.format(number, nValType, dDigitsCount, bChart, cultureInfo, true);
            }
        }
        if(!res){
            res = [{text: number.toString()}];
        }
        if (-1 != this.Color) {
            for (var i = 0; i < res.length; ++i) {
                var elem = res[i];
                if (null == elem.format) {
                    elem.format = new AscCommonExcel.Font();
                }
                elem.format.c = new AscCommonExcel.RgbColor(this.Color);
            }
        }
        return res;
    },
	_formatDecimalFrac: function(oParsedNumber) {
		var forceNull = false;
		for (var i = 0; i < this.aRawFormat.length; ++i) {
			var item = this.aRawFormat[i];
			if (numFormat_DecimalFrac == item.type) {
				var frac = oParsedNumber.fraction;
				var numerator = 0;
				var denominator = 0;
				if (item.bNumRight === true) {
					//todo max denominator - 99999
					denominator = item.aRight[0].val;
					numerator = Math.round(denominator * frac);
				} else if (frac > 0) {
					//Continued fraction
					//7 - excel max denominator length
					var denominatorLen = Math.min(7, item.aRight.length);
					var denominatorBound = Math.pow(10, denominatorLen);
					var an = Math.floor(frac);
					var xn1 = frac - an;
					var pn1 = an;
					var qn1 = 1;
					var pn2 = 1;
					var qn2 = 0;
					do {
						an = Math.floor(1 / xn1);
						xn1 = 1 / xn1 - an;
						var pn = an * pn1 + pn2;
						var qn = an * qn1 + qn2;
						pn2 = pn1;
						pn1 = pn;
						qn2 = qn1;
						qn1 = qn;
					} while (qn < denominatorBound);
					numerator = pn2;
					denominator = qn2;
				}
				if (numerator <= 0) {
					numerator = 0;
					if (this.bWhole === false) {
						if (denominator <= 0) {
							denominator = 1;
						}
					} else {
						denominator = 0;
					}
				}
				if (this.bWhole === false) {
					numerator += denominator * oParsedNumber.dec;
				} else if (numerator === denominator && 0 !== numerator) {
					oParsedNumber.dec++;
					numerator = 0;
					denominator = 0;
				}
				if (0 === numerator && 0 === denominator) {
					forceNull = true;
				}
				item.numerator = numerator;
				item.denominator = denominator;
			}
		}
		return forceNull;
	},
    format: function (number, nValType, dDigitsCount, cultureInfo, bChart, opt_forceNull)
    {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        var cultureInfoLCID = cultureInfo;
        if (null != this.LCID) {
            cultureInfoLCID = g_aCultureInfos[this.LCID] || cultureInfo;
        }
        if(null == nValType)
            nValType = CellValueType.Number;
        var res = [];
        var oCurText = {text: ""};
        if(true == this.valid)
        {
            if(true === this.bDateTime)
            {
                if(this.isInvalidDateValue(number) && this.formatType != AscCommon.NumFormatType.PDFFormDate)
                {
                    var oNewFont = new AscCommonExcel.Font();
					oNewFont.repeat = true;
                    this._CommitText(res, null, "#", oNewFont);
                    return res;
                }
            }
            var oParsedNumber;
			if (this.formatType == AscCommon.NumFormatType.PDFFormDate)
				oParsedNumber = this._parseNumberForPDFDate(number);
			else
				oParsedNumber = this._parseNumber(number, this.aDecFormat, this.aFracFormat.length, nValType);

            if (true == this.isGeneral() || (true == oParsedNumber.bDigit && true == this.bTextFormat) || (false == oParsedNumber.bDigit && false == this.bTextFormat) || (bChart && this.bGeneralChart))
            {
                return this._applyGeneralFormat(number, nValType, dDigitsCount, bChart, cultureInfo);
            }
			var forceNull = !!opt_forceNull;
			if (this.bSlash) {
				forceNull = this._formatDecimalFrac(oParsedNumber);
			}
            var aDec = [];
            var aFrac = [];
            var aScientific = [];
            if(true == oParsedNumber.bDigit)
            {
                aDec = this._FormatNumber(oParsedNumber.dec, oParsedNumber.exponent, this.aDecFormat.concat(), FormatStates.Decimal, cultureInfo, forceNull);
                aFrac = this._FormatNumber(oParsedNumber.frac, oParsedNumber.exponentFrac, this.aFracFormat.concat(), FormatStates.Frac, cultureInfo);
            }

            var bNoDecFormat = false;
            if((null == aDec || 0 == aDec.length) && 0 != oParsedNumber.dec)
            {
                //случай ".00"
                bNoDecFormat = true;
            }
            var hasSign = false;
            var nReadState = FormatStates.Decimal;
            var nFormatLength = this.aRawFormat.length;
			let isArabic = (lcid_ar === cultureInfoLCID.LCID
				|| lcid_arSY === cultureInfoLCID.LCID
				|| lcid_arSA === cultureInfoLCID.LCID
				|| lcid_arAE === cultureInfoLCID.LCID
				|| lcid_arBH === cultureInfoLCID.LCID
				|| lcid_arDZ === cultureInfoLCID.LCID
				|| lcid_arEG === cultureInfoLCID.LCID
				|| lcid_arIQ === cultureInfoLCID.LCID
				|| lcid_arJO === cultureInfoLCID.LCID
				|| lcid_arKW === cultureInfoLCID.LCID
				|| lcid_arQA === cultureInfoLCID.LCID
			);
			
			let _t = this;
			function checkRLM(prev)
			{
				if (!isArabic)
					return;
				
				if (undefined === prev
					|| prev < 0
					|| (numFormat_TimeSeparator !== _t.aRawFormat[prev].type
						&& (numFormat_Text !== _t.aRawFormat[prev].type || ":" !== _t.aRawFormat[prev].val)))
					oCurText.text += "‏";
			}
			
            for(var i = 0; i < nFormatLength; ++i)
            {
                var item = this.aRawFormat[i];
                if(numFormat_Bracket == item.type)
                {
                    if(null != item.CurrencyString)
                        oCurText.text += item.CurrencyString;
                }
                else if(numFormat_DecimalPoint == item.type)
                {
                    if(bNoDecFormat && null != oParsedNumber.dec && FormatStates.Decimal == nReadState)
                    {
                        oCurText.text += oParsedNumber.dec;
                    }
					oCurText.text += cultureInfo.NumberDecimalSeparator;
                    nReadState = FormatStates.Frac;
                }
                else if (numFormat_DecimalPointText == item.type) {
                    oCurText.text += cultureInfo.NumberDecimalSeparator;
                }
                else if (numFormat_ThousandText == item.type) {
                    oCurText.text += cultureInfo.NumberGroupSeparator;
                }
                else if(this._isDigitType(item.type))
                {
                    var text = null;
                    if(nReadState == FormatStates.Decimal)
                        text = aDec.shift();
                    else if(nReadState == FormatStates.Frac)
                        text = aFrac.shift();
                    else if(nReadState == FormatStates.Scientific)
                        text = aScientific.shift();
                    if(null != text)
                    {
                        this._AddDigItem(res, oCurText, text);
                    }
                }
                else if(numFormat_Text == item.type)
                {
					if(',' === item.val && isArabic) {
						oCurText.text += "،";
					} else {
						oCurText.text += item.val;
					}
                }
                else if(numFormat_TextPlaceholder == item.type)
                {
                    oCurText.text += number;
                }
                else if(numFormat_Scientific == item.type)
                {
                    if(null != item.format)
                    {
                        oCurText.text += item.val;

                        if(oParsedNumber.scientific < 0)
                            oCurText.text += "-";
                        else if(item.sign == SignType.Positive)
                            oCurText.text += "+";

                        
                        aScientific = this._FormatNumber(Math.abs(oParsedNumber.scientific), 0, item.format.concat(), FormatStates.Scientific, cultureInfo);
                        nReadState = FormatStates.Scientific;
                    }
                }
                else if(numFormat_DecimalFrac == item.type)
                {
                    var curForceNull = this.bWhole === false;
					var aLeft = this._FormatNumber(item.numerator, 0, item.aLeft.concat(), FormatStates.Slash, cultureInfo, curForceNull);
					for (var j = 0, length = aLeft.length; j < length; ++j) {
						var subitem = aLeft[j];
						if (subitem) {
							this._AddDigItem(res, oCurText, subitem);
						}
					}
					if ((item.numerator > 0 && item.denominator > 0) || curForceNull) {
						oCurText.text += "/";
					} else {
						var oNewFont = new AscCommonExcel.Font();
						oNewFont.skip = true;
						this._CommitText(res, oCurText, "/", oNewFont);
					}
					if (item.bNumRight === true) {
						var rightVal = item.aRight[0].val;
						if (rightVal) {
							if (item.denominator > 0) {
								oCurText.text += rightVal;
							} else {
								for (var rightIdx = 0; rightIdx < rightVal.toString().length; ++rightIdx) {
									var oNewFont = new AscCommonExcel.Font();
									oNewFont.skip = true;
									this._CommitText(res, oCurText, "0", oNewFont);
								}
							}
						}
					} else {
						var aRight = this._FormatNumber(item.denominator, 0, item.aRight.concat(), FormatStates.SlashFrac, cultureInfo);
						for (var j = 0, length = aRight.length; j < length; ++j) {
							var subitem = aRight[j];
							if (subitem) {
								this._AddDigItem(res, oCurText, subitem);
							}
						}
					}
                }
                else if(numFormat_Repeat == item.type)
                {
                    var oNewFont = new AscCommonExcel.Font();
					oNewFont.repeat = true;
                    this._CommitText(res, oCurText, item.val, oNewFont);
                }
                else if(numFormat_Skip == item.type)
                {
                    var oNewFont = new AscCommonExcel.Font();
					oNewFont.skip = true;
                    this._CommitText(res, oCurText, item.val, oNewFont);
                }
				else if(numFormat_DateSeparator == item.type)
                {
                    oCurText.text += cultureInfo.DateSeparator;
				}
				else if(numFormat_TimeSeparator == item.type)
                {
                    oCurText.text += cultureInfo.TimeSeparator;
				}
				else if(numFormat_DayOfWeek == item.type)
				{
					if (item.val === 3)
					{
						oCurText.text += cultureInfoLCID.AbbreviatedDayNames[oParsedNumber.date.dayWeek];
					}
					else if (item.val > 3)
					{
						oCurText.text += cultureInfoLCID.DayNames[oParsedNumber.date.dayWeek];
					}
					else
					{
						checkRLM();
						oCurText.text += 'a'.repeat(item.val);
					}
				}
                else if(numFormat_Year == item.type)
                {
                  if (item.val > 0) {
					  checkRLM();
                    if (item.val <= 2) {
						oCurText.text += (oParsedNumber.date.year.toString().slice(-2));
                    } else {
						if (oParsedNumber.date.year.toString().length < 4)
                    		oCurText.text += '0' + oParsedNumber.date.year;
						else
							oCurText.text += oParsedNumber.date.year;
                    }
                  }
                }
                else if(numFormat_Month == item.type)
                {
                    var m = oParsedNumber.date.month;
					if (item.val === 1) {
						checkRLM();
						oCurText.text += m + 1;
					} else if (item.val === 2) {
						checkRLM();
						oCurText.text += this._ZeroPad(m + 1);
					}
                    else if (item.val == 3) {
                        if (this.bDay && cultureInfoLCID.AbbreviatedMonthGenitiveNames.length > 0)
                            oCurText.text += cultureInfoLCID.AbbreviatedMonthGenitiveNames[m];
                        else
                            oCurText.text += cultureInfoLCID.AbbreviatedMonthNames[m];
                    }
                    else if (item.val == 5) {
                        var sMonthName = cultureInfoLCID.MonthNames[m];
                        if (sMonthName.length > 0)
                            oCurText.text += sMonthName[0];
                    }
                    else if (item.val > 0){
                        if (this.bDay && cultureInfoLCID.MonthGenitiveNames.length > 0)
                            oCurText.text += cultureInfoLCID.MonthGenitiveNames[m];
                        else
                            oCurText.text += cultureInfoLCID.MonthNames[m];
                    }
                }
                else if(numFormat_Day == item.type)
                {
                    if(item.val == 1) {
						checkRLM();
						oCurText.text += oParsedNumber.date.d;
					} else if(item.val === 2) {
						checkRLM();
						oCurText.text += this._ZeroPad(oParsedNumber.date.d);
					}
                    else if(item.val == 3)
                        oCurText.text += cultureInfoLCID.AbbreviatedDayNames[oParsedNumber.date.dayWeek];
                    else if(item.val > 0)
                        oCurText.text += cultureInfoLCID.DayNames[oParsedNumber.date.dayWeek];
                    
                }
                else if(numFormat_Hour == item.type)
                {
                    var h = oParsedNumber.date.hour;
                    if(item.bElapsed === true)
                        h = oParsedNumber.date.countDay*24 + oParsedNumber.date.hour;
                    if(this.bTimePeriod === true)
                        h = h%12||12;
					
					if (item.val > 0) {
						checkRLM(i - 1);
						if (item.val === 1)
							oCurText.text += h;
						else
							oCurText.text += this._ZeroPad(h);
					}
                }
                else if(numFormat_Minute == item.type)
                {
                    var min = oParsedNumber.date.min;
                    if(item.bElapsed === true)
                        min = oParsedNumber.date.countDay*24*60 + oParsedNumber.date.hour*60 + oParsedNumber.date.min;
					if (item.val > 0) {
						checkRLM(i - 1);
						if (item.val === 1)
							oCurText.text += min;
						else
							oCurText.text += this._ZeroPad(min);
					}
                }
                else if(numFormat_Second == item.type)
                {
                    var s = oParsedNumber.date.sec;
                    if(this.bMillisec === false)
                        s = oParsedNumber.date.sec + Math.round(oParsedNumber.date.ms/1000);
                    if(item.bElapsed === true)
                        s = oParsedNumber.date.countDay*24*60*60 + oParsedNumber.date.hour*60*60 + oParsedNumber.date.min*60 + s;
	
					if (item.val > 0) {
						checkRLM(i - 1);
						if (item.val === 1)
							oCurText.text += s;
						else
							oCurText.text += this._ZeroPad(s);
					}
                }
                else if (numFormat_AmPm == item.type) {
                    if (cultureInfoLCID.AMDesignator.length > 0 && cultureInfoLCID.PMDesignator.length > 0)
                        oCurText.text += (oParsedNumber.date.hour < 12) ? cultureInfoLCID.AMDesignator : cultureInfoLCID.PMDesignator;
                    else
                        oCurText.text += (oParsedNumber.date.hour < 12) ? "AM" : "PM";
                }
                else if (numFormat_Milliseconds == item.type) {
                    var nMsFormatLength = item.format.length;
                    var dMs = oParsedNumber.date.ms;
                    //Округляем
                    if (nMsFormatLength < 3) {
                        var dTemp = dMs / Math.pow(10, 3 - nMsFormatLength);
                        dTemp = Math.round(dTemp);
                        dMs = dTemp * Math.pow(10, 3 - nMsFormatLength);
                    }
                    var nExponent = 0;
                    if(0 == dMs)
                        nExponent = -1;
                    else if (dMs < 10)
                        nExponent = -2;
                    else if (dMs < 100)
                        nExponent = -1;
                    var aMilSec = this._FormatNumber(dMs, nExponent, item.format.concat(), FormatStates.Frac, cultureInfo);
					checkRLM(i - 1);
                    for (var k = 0; k < aMilSec.length; k++)
                        this._AddDigItem(res, oCurText, aMilSec[k]);
                }
                else if (numFormat_General == item.type) {
                    this._CommitText(res, oCurText, null, null);
                    //todo minus sign
                    res = res.concat(this._applyGeneralFormat(Math.abs(number), nValType, dDigitsCount, bChart, cultureInfo));
                } else if (numFormat_Plus == item.type) {
					hasSign = true;
					if (number > 0) {
						oCurText.text += '+';
					} else if (number < 0) {
						oCurText.text += '-';
					} else {
						oCurText.text += ' ';
					}
				} else if (numFormat_Minus == item.type) {
					hasSign = true;
					if (number < 0) {
						oCurText.text += '-';
					} else {
						oCurText.text += ' ';
					}
				}
            }

			if (true == this.bAddMinusIfNes && SignType.Negative == oParsedNumber.sign && !hasSign) {
				//todo разобраться с минусами
				//Добавляем в самое начало знак минус
				res.unshift({text: "-"});
			}
            this._CommitText(res, oCurText, null, null);
			if(0 == res.length)
                res = [{text: ""}];
        }
        else
        {
            if(0 == res.length)
                res = [{text: number.toString()}];
        }
		//длина результирующей строки не должна быть длиннее c_oAscMaxColumnWidth
		var nLen = 0;
		for(var i = 0; i < res.length; ++i){
			var elem = res[i];
			if (elem.text) {
				elem.text = this._replaceDBNumDigit(elem.text);
				nLen += elem.text.length;
			}
		}
		if(nLen > Asc.c_oAscMaxColumnWidth){
			var oNewFont = new AscCommonExcel.Font();
			oNewFont.repeat = true;
			res = [{text: "#", format: oNewFont}];
		}
        return res;
    },
	shiftFormat : function(output, nShift, useLocaleFormat) {
		if (this.bDateTime || this.bSlash || this.bTextFormat)
			return false;
		output.format = this.toString(nShift, useLocaleFormat);
		return true;
	},
	toString : function(nShift, useLocaleFormat)
	{
		var sGeneral;
		var DecimalSeparator;
		var GroupSeparator;
		var TimeSeparator;
		var year;
		var month;
		var day;
		var hour;
		var minute;
		var second;
		var dayOfWeek;
		if (useLocaleFormat) {
			sGeneral = LocaleFormatSymbol['general'];
			DecimalSeparator = g_oDefaultCultureInfo.NumberDecimalSeparator;
			TimeSeparator = g_oDefaultCultureInfo.TimeSeparator;
			GroupSeparator = g_oDefaultCultureInfo.NumberGroupSeparator;
			if (LocaleFormatSymbol['M'] === LocaleFormatSymbol['m']) {
				year = LocaleFormatSymbol['Y'];
				month = LocaleFormatSymbol['M'];
				day = LocaleFormatSymbol['D'];
			} else {
				year = LocaleFormatSymbol['y'];
				month = LocaleFormatSymbol['m'];
				day = LocaleFormatSymbol['d'];
			}
			hour = LocaleFormatSymbol['h'];
			minute = LocaleFormatSymbol['minute'];
			second = LocaleFormatSymbol['s'];
			dayOfWeek = LocaleFormatSymbol['a'];
		} else {
			sGeneral = AscCommon.g_cGeneralFormat;
			DecimalSeparator = gc_sFormatDecimalPoint;
			TimeSeparator = ':';
			GroupSeparator = gc_sFormatThousandSeparator;
			year = 'y';
			month = 'm';
			day = 'd';
			hour = 'h';
			minute = 'm';
			second = 's';
			dayOfWeek = 'a';
		}
        var nDecLength = this.aDecFormat.length;
        var nDecIndex = 0;
        var nFracLength = this.aFracFormat.length;
        var nFracIndex = 0;
        var nNewFracLength = nFracLength + nShift;
        if(nNewFracLength < 0)
            nNewFracLength = 0;
        var nReadState = FormatStates.Decimal;
        var res = "";
        var fFormatToString = function(aFormat)
        {
            var res = "";
            for(var i = 0, length = aFormat.length; i < length; ++i)
            {
                var item = aFormat[i];
                if(numFormat_Digit == item.type)
                {
                    if(null != item.val)
                        res += item.val;
                    else
                        res += "0";
                }
                else if(numFormat_DigitNoDisp == item.type)
                    res += "#";
                else if(numFormat_DigitSpace == item.type)
                    res += "?";
				else if(numFormat_DigitDrop == item.type)
					res += "x";
            }
            return res;
        };
        //Color
        if(null != this.Color)
        {
            switch(this.Color)
            {
            case 0x000000: res += "[Black]";break;
            case 0x0000ff: res += "[Blue]";break;
            case 0x00ffff: res += "[Cyan]";break;
            case 0x00ff00: res += "[Green]";break;
            case 0xff00ff: res += "[Magenta]";break;
            case 0xff0000: res += "[Red]";break;
            case 0xffffff: res += "[White]";break;
            case 0xffff00: res += "[Yellow]";break;
            }
        }
		//Comporation operator
        if(null != this.ComporationOperator)
        {
			switch(this.ComporationOperator.operator)
			{
				case NumComporationOperators.equal: res += "[=" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.greater: res += "[>" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.less: res += "[<" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.greaterorequal: res += "[>=" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.lessorequal: res += "[<=" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.notequal: res += "[<>" + this.ComporationOperator.operatorValue +"]";break;
			}
		}
		if (this.DBNum > 0)
		{
			res += '[DBNum' + this.DBNum + ']';
		}

        var nFormatLength = this.aRawFormat.length;    
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_Bracket == item.type)
            {
                if(null != item.CurrencyString || null != item.Lid)
                {
                    res += "[$";
                    if(null != item.CurrencyString)
                        res += item.CurrencyString;
					if (null != item.Lid) {
						res += "-";
						res += item.Lid;
					}
                    res += "]";
                }
            }
            else if(numFormat_DecimalPoint == item.type)
            {
                nReadState = FormatStates.Frac;
                if(0 != nNewFracLength)
                    res += DecimalSeparator;
            }
            else if (numFormat_DecimalPointText == item.type) {
                res += DecimalSeparator;
            }
            else if(numFormat_Thousand == item.type || numFormat_ThousandText == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += GroupSeparator;
            }
            else if(this._isDigitType(item.type))
            {
                if(FormatStates.Decimal == nReadState)
                    nDecIndex++;
                else
                    nFracIndex++;
                if(nReadState == FormatStates.Frac && nFracIndex > nNewFracLength)
                    ;
                else
                {
                    var sCurSimbol;
                    if(numFormat_Digit == item.type)
                        sCurSimbol = "0";
                    else if(numFormat_DigitNoDisp == item.type)
                        sCurSimbol = "#";
                    else if(numFormat_DigitSpace == item.type)
                        sCurSimbol = "?";
					else if(numFormat_DigitDrop == item.type)
						sCurSimbol = "x";
                    res += sCurSimbol;
                    if(nReadState == FormatStates.Frac && nFracIndex == nFracLength)
                    {
                        for(var j = 0; j < nShift; ++j)
                            res += sCurSimbol;
                    }
                }
                if(0 == nFracLength && nShift > 0 && FormatStates.Decimal == nReadState && nDecIndex == nDecLength)
                {
                    res += gc_sFormatDecimalPoint;
                    for(var j = 0; j < nShift; ++j)
                        res += "0";
                }
            }
            else if(numFormat_Text == item.type)
            {
                if("%" == item.val)
                    res += item.val;
                else
                    res += "\"" + item.val + "\"";
            }
            else if(numFormat_TextPlaceholder == item.type)
                res += "@";
            else if(numFormat_Scientific == item.type)
            {
                nReadState = FormatStates.Scientific;
                res += item.val;
                if(item.sign == SignType.Positive)
                    res += "+";
                else
                    res += "-";
            }
            else if(numFormat_DecimalFrac == item.type)
            {
                res += fFormatToString(item.aLeft);
                res += "/";
                res += fFormatToString(item.aRight);
            }
            else if(numFormat_Repeat == item.type)
                res += "*" + item.val;
            else if(numFormat_Skip == item.type)
                res += "_" + item.val;
			else if(numFormat_DateSeparator == item.type)
                res += "/";
			else if(numFormat_TimeSeparator == item.type)
                res += TimeSeparator;
            else if(numFormat_Year == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += year;
            }
            else if(numFormat_Month == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += month;
            }
            else if(numFormat_Day == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += day;
            }
            else if(numFormat_Hour == item.type)
            {
				if (item.bElapsed) {
					res += "[";
				}
				for(var j = 0; j < item.val; ++j)
					res += hour;
				if (item.bElapsed) {
					res += "]";
				}
            }
            else if(numFormat_Minute == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += minute;
            }
            else if(numFormat_Second == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += second;
            }
			else if(numFormat_DayOfWeek == item.type)
			{
				var nIndex = (item.val > 3) ? 3 : item.val;
				for(var j = 0; j < nIndex; ++j)
					res += dayOfWeek;
			}
            else if(numFormat_AmPm == item.type)
                res += "AM/PM";
            else if(numFormat_Milliseconds == item.type)
                res += fFormatToString(item.format);
			else if(numFormat_Plus == item.type)
				res += "+";
			else if(numFormat_Minus == item.type)
				res += "-";
			else if(numFormat_General == item.type)
				res += sGeneral;
        }
        return res;
    },
	getFormatCellsInfo: function() {
		var info = new Asc.asc_CFormatCellsInfo();
		info.asc_setDecimalPlaces(this.aFracFormat.length);
		info.asc_setSeparator(this.bThousandSep);
		info.asc_setSymbol(this.LCID);
		info.asc_setCurrencySymbol(this.CurrencyString);
		return info;
	},
	isGeneral: function() {
		return 1 == this.aRawFormat.length && numFormat_General == this.aRawFormat[0].type;
	}
};
function NumFormatCache()
{
    this.oNumFormats = {};
}
NumFormatCache.prototype =
{
	cleanCache : function(){
		this.oNumFormats = {};
	},
    get : function(format, formatType)
    {
		var key = format + String.fromCharCode(5) + formatType;
        var res = this.oNumFormats[key];
        if(null == res)
        {
            res = new CellFormat(format, formatType, false);
            this.oNumFormats[key] = res;
        }
        return res;
    }
};
//кеш структур по строке формата
var oNumFormatCache = new NumFormatCache();

function CellFormat(format, formatType, useLocaleFormat)
{
    this.sFormat = format;
    this.oPositiveFormat = null;
    this.oNegativeFormat = null;
    this.oNullFormat = null;
    this.oTextFormat = null;
	this.aComporationFormats = null;
    var aFormats = format.split(";");
	var aParsedFormats = [];
	for(var i = 0; i < aFormats.length; ++i)
	{
    var sNewFormat = aFormats[i];
    //если sNewFormat заканчивается на нечетное число '\', значит ';' был экранирован и это текст
    while(true){
      var formatTail = sNewFormat.match(/\\+$/g);
      if (formatTail && formatTail.length > 0 && 1 === formatTail[0].length % 2 && i + 1 < aFormats.length) {
        sNewFormat += ';';
        sNewFormat += aFormats[++i];
      } else {
        break;
      }
    }
		var oNewFormat = new NumFormat(false);
		oNewFormat.setFormat(sNewFormat, undefined, formatType, useLocaleFormat);
		if (oNewFormat.LCID === 0xF800) {
			sNewFormat = '[$-F800]' + g_oDefaultCultureInfo.LongDatePattern;
			oNewFormat = new NumFormat(false);
			oNewFormat.setFormat(sNewFormat, undefined, formatType, useLocaleFormat);
		}
		aParsedFormats.push(oNewFormat);
	}
  var nFormatsLength = aParsedFormats.length;
	var noComparisonn = aParsedFormats.every(function(format) {return !format.ComporationOperator});
	if(noComparisonn)
	{
		if(4 <= nFormatsLength)
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oNegativeFormat = aParsedFormats[1];
			this.oNullFormat = aParsedFormats[2];
			this.oTextFormat = aParsedFormats[3];
			//for ';;;' format, if 4 formats exist fourth always used for text
			this.oTextFormat.bTextFormat = true;
		}
		else if(3 == nFormatsLength)
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oNegativeFormat = aParsedFormats[1];
			this.oNullFormat = aParsedFormats[2];
			this.oTextFormat = this.oPositiveFormat;
			if (this.oNullFormat.bTextFormat) {
				this.oTextFormat = this.oNullFormat;
				this.oNullFormat = this.oPositiveFormat;
			}
		}
		else if(2 == nFormatsLength)
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oNegativeFormat = aParsedFormats[1];
			this.oNullFormat = this.oPositiveFormat;
			this.oTextFormat = this.oPositiveFormat;
			if (this.oNegativeFormat.bTextFormat) {
				this.oTextFormat = this.oNegativeFormat;
				this.oNegativeFormat = this.oPositiveFormat;
				this.oPositiveFormat.bAddMinusIfNes = true;
			}
		}
		else
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oPositiveFormat.bAddMinusIfNes = true;
			this.oNegativeFormat = this.oPositiveFormat;
			this.oNullFormat = this.oPositiveFormat;
			this.oTextFormat = this.oPositiveFormat;
		}
	}
	else
	{
		this.oTextFormat = new NumFormat(false);
		this.oTextFormat.setFormat("@", undefined, undefined, useLocaleFormat);
		//по результатам опытов, если оператор сравнения проходит через 0, то надо добавлять знак минус в зависимости от значения
		//пример [<100] надо добавлять знак, [<-100] знак добавлять не надо
		for (let i = 0; i < aParsedFormats.length && i < 2; ++i) {
			let oCurFormat = aParsedFormats[i];
			if (oCurFormat.ComporationOperator) {
				let operator = oCurFormat.ComporationOperator.operator;
				let operatorValue = oCurFormat.ComporationOperator.operatorValue;
				if (0 < operatorValue && (operator === NumComporationOperators.less || operator === NumComporationOperators.lessorequal))
					oCurFormat.bAddMinusIfNes = true;
				else if (0 > operatorValue && (operator === NumComporationOperators.greater || operator === NumComporationOperators.greaterorequal))
					oCurFormat.bAddMinusIfNes = true;
			}
		}
		if (aParsedFormats.length > 2) {
			aParsedFormats[2].bAddMinusIfNes = true;
		}
		this.aComporationFormats = aParsedFormats.slice(0, 3);
	}
    this.formatCache = {};
}
CellFormat.prototype =
{
	isTextFormat : function()
	{
		if (this.oPositiveFormat  != null) {
			return this.oPositiveFormat.bTextFormat;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bTextFormat;
		}
		return false;
	},
	isGeneralFormat : function()
	{
		if (this.oPositiveFormat != null) {
			return this.oPositiveFormat.isGeneral();
		} else if (this.aComporationFormats != null  && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].isGeneral();
		}
		return false;
	},
	isDateTimeFormat : function()
	{
		if (this.oPositiveFormat != null) {
			return this.oPositiveFormat.bDateTime;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bDateTime;
		}
		return false;
	},
	isTimeFormat : function() {
		if (this.oPositiveFormat != null) {
			return this.oPositiveFormat.bTime;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bTime;
		}
		return false;
	},
	isDateFormat : function() {
		if ( this.oPositiveFormat != null) {
			return this.oPositiveFormat.bDate;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bDate;
		}
		return false;
	},
	getTextFormat: function () {
	    var oRes = null;
	    if (null == this.aComporationFormats) {
	        if (null != this.oTextFormat && this.oTextFormat.bTextFormat)
	            oRes = this.oTextFormat;
	    } else {
	        for (var i = 0, length = this.aComporationFormats.length; i < length ; ++i) {
	            var oCurFormat = this.aComporationFormats[i];
	            if (null == oCurFormat.ComporationOperator && oCurFormat.bTextFormat) {
	                oRes = oCurFormat;
	                break;
	            }
	        }
	    }
	    return oRes;
	},
	getFormatByValue : function(dNumber)
	{
		var oRes = null;
		if(null == this.aComporationFormats)
		{
			if(dNumber > 0 && null != this.oPositiveFormat)
				oRes = this.oPositiveFormat;
			else if(dNumber < 0 && null != this.oNegativeFormat)
				oRes = this.oNegativeFormat;
			else if(null != this.oNullFormat)
				oRes = this.oNullFormat;
		}
		else
		{
			//ищем совпадение
			for (let i = 0; i < this.aComporationFormats.length && i < 2; ++i)
			{
				let oCurFormat = this.aComporationFormats[i];
				let oOperationValue, operator;
				if (null != oCurFormat.ComporationOperator) {
					operator = oCurFormat.ComporationOperator.operator;
					oOperationValue = oCurFormat.ComporationOperator.operatorValue;
				} else {
					oOperationValue = 0;
					operator = 0 === i ? NumComporationOperators.greater : NumComporationOperators.less;
				}
				let isMatch = (operator === NumComporationOperators.equal && dNumber === oOperationValue) ||
					(operator === NumComporationOperators.greater && dNumber > oOperationValue) ||
					(operator === NumComporationOperators.less && dNumber < oOperationValue) ||
					(operator === NumComporationOperators.greaterorequal && dNumber >= oOperationValue) ||
					(operator === NumComporationOperators.lessorequal && dNumber <= oOperationValue) ||
					(operator === NumComporationOperators.notequal && dNumber !== oOperationValue);
				if (isMatch) {
					oRes = oCurFormat;
					break;
				}
			}
			if (null == oRes && null != this.aComporationFormats.length > 2)
				oRes = this.aComporationFormats[2];
		}
		return oRes;
	},
    format : function(number, nValType, dDigitsCount, bChart, cultureInfo, opt_withoutCache, opt_forceNull)
    {
        var res = null;
        if (null == bChart)
            bChart = false;
        var lcid = cultureInfo ? cultureInfo.LCID : 0;
        var cacheKey, cacheVal;
        if (!opt_withoutCache) {
            cacheKey = number + '-' + nValType + '-' + dDigitsCount + '-' + lcid;
            cacheVal = this.formatCache[cacheKey];
            if(null != cacheVal)
            {
                if (bChart)
                    res = cacheVal.chart;
                else
                    res = cacheVal.nochart;
                if (null != res)
                    return res;
            }
        }
        res = [{text: number.toString()}];
        var dNumber = number - 0;
        var oFormat = null;
		if(CellValueType.String != nValType && number == dNumber)
		{
			oFormat = this.getFormatByValue(dNumber);
			if(null != oFormat)
			    res = oFormat.format(number, nValType, dDigitsCount, cultureInfo, bChart, opt_forceNull);
			else if(null != this.aComporationFormats)
			{
			    var oNewFont = new AscCommonExcel.Font();
				oNewFont.repeat = true;
				res = [{text: "#", format: oNewFont}];
			}
		}
		else
		{
			//text
		    if (null != this.oTextFormat) {
		        oFormat = this.oTextFormat;
		        res = oFormat.format(number, nValType, dDigitsCount, cultureInfo, bChart, opt_forceNull);
		    }
		}
        if (!opt_withoutCache) {
            if (null == cacheVal) {
                cacheVal = {chart: null, nochart: null};
                this.formatCache[cacheKey] = cacheVal;
            }
            if (null != oFormat && oFormat.bGeneralChart) {
                if (bChart)
                    cacheVal.chart = res;
                else
                    cacheVal.nochart = res;
            }
            else {
                cacheVal.chart = res;
                cacheVal.nochart = res;
            }
        }
        return res;
    },
    shiftFormat : function(output, nShift, useLocaleFormat)
    {
        var bRes = false;
        var bCurRes = true;
		if(null == this.aComporationFormats)
		{
			bCurRes = this.oPositiveFormat.shiftFormat(output, nShift, useLocaleFormat);
			if(false == bCurRes)
				output.format = this.oPositiveFormat.formatString;
			bRes |= bCurRes;
			if(null != this.oNegativeFormat && this.oPositiveFormat != this.oNegativeFormat)
			{
				var oTempOutput = {};
				bCurRes = this.oNegativeFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(false == bCurRes)
					output.format += ";" + this.oNegativeFormat.formatString;
				else
					output.format += ";" + oTempOutput.format;
				bRes |= bCurRes;
			}
			if(null != this.oNullFormat && this.oPositiveFormat != this.oNullFormat)
			{
				var oTempOutput = {};
				bCurRes = this.oNullFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(false == bCurRes)
					output.format += ";" + this.oNullFormat.formatString;
				else
					output.format += ";" + oTempOutput.format;
				bRes |= bCurRes;
			}
			if(null != this.oTextFormat && this.oPositiveFormat != this.oTextFormat)
			{
				var oTempOutput = {};
				bCurRes = this.oTextFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(false == bCurRes)
					output.format += ";" + this.oTextFormat.formatString;
				else
					output.format += ";" + oTempOutput.format;
				bRes |= bCurRes;
			}
		}
		else
		{
			var length = this.aComporationFormats.length;
			output.format = "";
			for(var i = 0; i < length; ++i)
			{
				var oTempOutput = {};
				var oCurFormat = this.aComporationFormats[i];
				var bCurRes = oCurFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(0 != i)
					output.format += ";";
				if(false == bCurRes)
					output.format += oCurFormat.formatString;
				else
					output.format += oTempOutput.format;
				bRes |= bCurRes;
			}
		}
        return bRes;
    },
	toString: function(nShift, useLocaleFormat) {
		var res = '';
		if (null == this.aComporationFormats) {
			res += this.oPositiveFormat.toString(nShift, useLocaleFormat);
			if (null != this.oNegativeFormat && this.oPositiveFormat != this.oNegativeFormat) {
				res += ";" + this.oNegativeFormat.toString(nShift, useLocaleFormat);
			}
			if (null != this.oNullFormat && this.oPositiveFormat != this.oNullFormat) {
				res += ";" + this.oNullFormat.toString(nShift, useLocaleFormat);
			}
			if (null != this.oTextFormat && this.oPositiveFormat != this.oTextFormat) {
				res += ";" + this.oTextFormat.toString(nShift, useLocaleFormat);
			}
		}
		else {
			var length = this.aComporationFormats.length;
			for (var i = 0; i < length; ++i) {
				var oCurFormat = this.aComporationFormats[i];
				if (0 != i) {
					res += ";";
				}
				res += oCurFormat.toString(nShift, useLocaleFormat);
			}
		}
		return res;
	},
	formatToMathInfo : function(number, nValType, dDigitsCount)
	{
		return this._formatToText(number, nValType, dDigitsCount, false);
	},
	formatToChart : function(number, dDigitsCount, cultureInfo)
	{
		return this._formatToText(number, CellValueType.Number, dDigitsCount || gc_nMaxDigCount, true, cultureInfo);
	},
	formatToWord : function(number, dDigitsCount, cultureInfo)
	{
		return this._formatToText(number, CellValueType.Number, dDigitsCount || gc_nMaxDigCount, false, cultureInfo, true);
	},
	_formatToText : function(number, nValType, dDigitsCount, bChart, cultureInfo, opt_forceNull)
	{
		var result = "";
		var arrFormat = this.format(number, nValType, dDigitsCount, bChart, cultureInfo, undefined, opt_forceNull);
		for (var i = 0, item; i < arrFormat.length; ++i) {
			item = arrFormat[i];
			if (item.format) {
				if (item.format.repeat)
					continue;
				if (item.format.skip) {
					result += " ";
					continue;
				}
			}
			if (item.text)
				result += item.text;
		}
		return result;
	},
	getType: function() {
		return this.getTypeInfo().type;
	},
	getTypeInfo: function() {
		var info;
		if (null != this.oPositiveFormat) {
			info = this.oPositiveFormat.getFormatCellsInfo();
			info.asc_setType(this._getType(this.oPositiveFormat));
		} else if (null != this.aComporationFormats && this.aComporationFormats.length > 0) {
			info = this.aComporationFormats[0].getFormatCellsInfo();
			info.asc_setType(this._getType(this.aComporationFormats[0]));
		} else {
			info = new Asc.asc_CFormatCellsInfo();
			info.asc_setType(c_oAscNumFormatType.General);
			info.asc_setDecimalPlaces(0);
			info.asc_setSeparator(false);
			info.asc_setSymbol(null);
		}
		return info;
	},
	_getType: function(format) {
		var nType = c_oAscNumFormatType.Custom;
		if (format.isGeneral()) {
			nType = c_oAscNumFormatType.General;
		}
		else if (format.bDateTime) {
			if (format.bDate) {
				nType = c_oAscNumFormatType.Date;
			} else {
				nType = c_oAscNumFormatType.Time;
			}
		}
		else if (format.bCurrency) {
			if (format.bRepeat) {
				nType = c_oAscNumFormatType.Accounting;
			} else {
				nType = c_oAscNumFormatType.Currency;
			}
		} else {
			var info = format.getFormatCellsInfo();
			var types = [c_oAscNumFormatType.Text, c_oAscNumFormatType.Percent, c_oAscNumFormatType.Scientific,
				c_oAscNumFormatType.Number, c_oAscNumFormatType.Fraction, c_oAscNumFormatType.Currency,
				c_oAscNumFormatType.Accounting
			];
			for (var i = 0; i < types.length; ++i) {
				var type = types[i];
				info.asc_setType(type);
				var formats = getFormatCells(info);
				if (-1 != formats.indexOf(this.sFormat)) {
					nType = type;
					break;
				}
			}
		}
		return nType;
	},
	checkCultureInfoFontPicker: function() {
		if (null !== this.sFormat) {
			AscFonts.FontPickerByCharacter.getFontsByString(this.sFormat);
		}
		if (null !== this.oPositiveFormat && null !== this.oPositiveFormat.LCID) {
			checkCultureInfoFontPicker(this.oPositiveFormat.LCID);
		}
		if (null !== this.oNegativeFormat && null !== this.oNegativeFormat.LCID) {
			checkCultureInfoFontPicker(this.oNegativeFormat.LCID);
		}
		if (null !== this.oNullFormat && null !== this.oNullFormat.LCID) {
			checkCultureInfoFontPicker(this.oNullFormat.LCID);
		}
		if (null !== this.oTextFormat && null !== this.oTextFormat.LCID) {
			checkCultureInfoFontPicker(this.oTextFormat.LCID);
		}
		if (this.aComporationFormats) {
			for (var i = 0, length = this.aComporationFormats.length; i < length; ++i) {
				var oCurFormat = this.aComporationFormats[i];
				if (null !== oCurFormat.LCID) {
					checkCultureInfoFontPicker(oCurFormat.LCID);
				}

			}
		}
	}
};
var oDecodeGeneralFormatCache = {};
function DecodeGeneralFormat(val, nValType, dDigitsCount)
{
    var cacheVal = oDecodeGeneralFormatCache[val];
    if(null != cacheVal)
    {
        cacheVal = cacheVal[nValType];
        if(null != cacheVal)
        {
            cacheVal = cacheVal[dDigitsCount];
            if(null != cacheVal)
                return cacheVal;
        }
    }
    var res = DecodeGeneralFormat_Raw(val, nValType, dDigitsCount);
    var cacheVal = oDecodeGeneralFormatCache[val];
    if(null == cacheVal)
    {
        cacheVal = {};
        oDecodeGeneralFormatCache[val] = cacheVal;
    }
    var cacheType = cacheVal[nValType];
    if(null == cacheType)
    {
        cacheType = {};
        cacheVal[nValType] = cacheType;
    }
    cacheType[dDigitsCount] = res;
    return res;
}
function DecodeGeneralFormat_Raw(val, nValType, dDigitsCount)
{
    if(CellValueType.String == nValType)
        return "@";
    var number = val - 0;
    if(number != val)
        return "@";
    if(0 == number)
        return "0";
    var nDigitsCount;
    if(null == dDigitsCount || dDigitsCount > gc_nMaxDigCountView)
        nDigitsCount = gc_nMaxDigCountView;
    else
        nDigitsCount = parseInt(dDigitsCount);//пока не подключена измерялся не используем нецелые метрики
    if(number < 0)
    {
        //todo возможно нужно nDigitsCount--
        //nDigitsCount--;//на знак '-'
        number = -number;
    }
    if(nDigitsCount < 1)
        return "0";//можно возвращать любой числовой формат, все равно при nDigitsCount < 1 он учитываться не будет
	var bContinue = true;
	var parts = getNumberParts(number);
	while(bContinue)
	{
		bContinue = false;
		var nRealExp = gc_nMaxDigCount + parts.exponent;//nRealExp == 0, при 0,123
		var nRealExpAbs = Math.abs(nRealExp);
		var nExpMinDigitsCount;//число знаков в формате 'E+00'
		if(nRealExpAbs < 100)
			nExpMinDigitsCount = 4;
		else
			nExpMinDigitsCount = 2 + nRealExpAbs.toString().length;
		
		var suffix = "";
		if (nRealExp > 0)
		{
			if(nRealExp > nDigitsCount)
			{
				if(nDigitsCount >= nExpMinDigitsCount + 1)//1 на еще один символ перед E (*E+00)
				{
					suffix = "E+";
					for(var i = 2; i < nExpMinDigitsCount; ++i)
						suffix += "0";
					nDigitsCount -= nExpMinDigitsCount;
				}
				else
					return "0";//можно возвращать любой числовой формат, все равно будут решетки
			}
		}
		else
		{
			var nVarian1 = nDigitsCount - 2 + nRealExp;//без E+00, 2 на знаки "0."
			var nVarian2 = nDigitsCount - nExpMinDigitsCount;// с E+00
			if(nVarian2 > 2)
				nVarian2--;//на знак '.'
			else if(nVarian2 > 0)
				nVarian2 = 1;
			if(nVarian1 <= 0 && nVarian2 <= 0)
				return "0";
			if(nVarian1 < nVarian2)
			{
				//если в nVarian1 число помещается полностью, то применяем nVarian1
				var bUseVarian1 = false;
				if(nVarian1 > 0 && 0 == (parts.mantissa % Math.pow(10, gc_nMaxDigCount - nVarian1)))
					bUseVarian1 = true;
				if(false == bUseVarian1)
				{
					if(nDigitsCount >= nExpMinDigitsCount + 1)
					{
						suffix = "E+";
						for(var i = 2; i < nExpMinDigitsCount; ++i)
							suffix += "0";
						nDigitsCount -= nExpMinDigitsCount;
					}
					else
						return "0";//можно возвращать любой числовой формат, все равно будут решетки
				}
			}
		}
		var dec_num_digits = nRealExp;
		if(suffix)
			dec_num_digits = 1;
		//округляем мантиссу, чтобы правильно обрабатывать ситуацию 0,999, когда nDigitsCount = 4
		var nRoundDigCount = 0;
		if(dec_num_digits <= 0)
		{
			//2 на знаки '0.'
			var nTemp = nDigitsCount + dec_num_digits - 2;
			if(nTemp > 0)
				nRoundDigCount = nTemp;
		}
		else if(dec_num_digits < gc_nMaxDigCount)
		{
			if(dec_num_digits <= nDigitsCount)
			{
				//1 на знаки '.'
				if(dec_num_digits + 1 < nDigitsCount)
					nRoundDigCount = nDigitsCount - 1;
				else
					nRoundDigCount = dec_num_digits;
			}
		}
		if(nRoundDigCount > 0)
		{
			var nTemp = Math.pow(10, gc_nMaxDigCount - nRoundDigCount);
			number = Math.round(parts.mantissa / nTemp) * nTemp * Math.pow(10, parts.exponent);
			
			var oNewParts = getNumberParts(number);
			//если в результате округления изменилось число разрядов, надо начинать заново
			if(oNewParts.exponent != parts.exponent)
				bContinue = true;
			else
				bContinue = false;
			parts = oNewParts;
		}
	}
	
    var frac_num_digits;
    if(dec_num_digits > 0)
        frac_num_digits = nDigitsCount - 1 - dec_num_digits;//1 на знак '.'
    else
        frac_num_digits = nDigitsCount - 2 + dec_num_digits;//2 на знаки '0.' 
        
    //считаем необходимое число знаков после запятой
    if(frac_num_digits > 0)
    {
		var sTempNumber = parts.mantissa.toString();
		var sTempNumber;
		if(dec_num_digits > 0)
			sTempNumber = sTempNumber.substring(dec_num_digits, dec_num_digits + frac_num_digits);
		else
			sTempNumber = sTempNumber.substring(0, frac_num_digits);
        var nTempNumberLength = sTempNumber.length;
        var nreal_frac_num_digits = frac_num_digits;
        for(var i = frac_num_digits - 1; i >= 0; --i)
        {
            if("0" == sTempNumber[i])
                nreal_frac_num_digits--;
            else
                break;
        }
        frac_num_digits = nreal_frac_num_digits;
		if(dec_num_digits < 0)
			frac_num_digits += (-dec_num_digits);
    }
    if(frac_num_digits <= 0)
        return "0" + suffix;

    //собираем формат
    var number_format_string = "0" + gc_sFormatDecimalPoint;
    for(var i = 0; i < frac_num_digits; ++i)
        number_format_string += "0";
    number_format_string += suffix;
    return number_format_string;
}
function GeneralEditFormatCache()
{
    this.oCache = {};
}
GeneralEditFormatCache.prototype =
{
	cleanCache : function(){
		this.oCache = {};
	},
    format: function (number, cultureInfo)
    {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        //преобразуем число так чтобы в строке было только 15 значящих цифр.
        var value = this.oCache[number];
        if(null == value)
        {
			if(0 == number)
				value = "0";
			else
			{
				var sRes = "";
				var parts = getNumberParts(number);
				var nRealExp = gc_nMaxDigCount + parts.exponent;//nRealExp == 0, при 0,123
				if(parts.exponent >= 0)//nRealExp >= -gc_nMaxDigCount
				{
					if(nRealExp <= 21)
					{
						sRes = parts.mantissa.toString();
						for(var i = 0; i < parts.exponent; ++i)
							sRes += "0";
					}
					else
					{
					    sRes = this._removeTileZeros(parts.mantissa.toString(), cultureInfo);
						if(sRes.length > 1)
						{
							var temp = sRes.substring(0, 1);
							temp += cultureInfo.NumberDecimalSeparator;
							temp += sRes.substring(1);
							sRes = temp;
						}
						sRes += "E+" + (nRealExp - 1);
					}
				}
				else
				{
					if(nRealExp > 0)
					{
						sRes = parts.mantissa.toString();
						if(sRes.length > nRealExp)
						{
							var temp = sRes.substring(0, nRealExp);
							temp += cultureInfo.NumberDecimalSeparator;
							temp += sRes.substring(nRealExp);
							sRes = temp;
						}
						sRes = this._removeTileZeros(sRes, cultureInfo);
					}
					else
					{
						if(nRealExp >= -18)
						{
							sRes = "0";
							sRes += cultureInfo.NumberDecimalSeparator;
							for(var i = 0; i < -nRealExp; ++i)
								sRes += "0";
							var sTemp = parts.mantissa.toString();
							sTemp = sTemp.substring(0, 19 + nRealExp);
							sRes += this._removeTileZeros(sTemp, cultureInfo);
						}
						else
						{
							sRes = parts.mantissa.toString();
							if(sRes.length > 1)
							{
								var temp = sRes.substring(0, 1);
								temp += cultureInfo.NumberDecimalSeparator;
								temp += sRes.substring(1);
								temp = this._removeTileZeros(temp, cultureInfo);
								sRes = temp;
							}
							sRes += "E-" + (1 - nRealExp);
						}
					}
				}
				if( SignType.Negative == parts.sign)
					value = "-" + sRes;
				else
					value = sRes;
			}
            this.oCache[number] = value;
        }
        return value;
    },
    _removeTileZeros: function (val, cultureInfo)
    {
		var res = val;
		var nLength = val.length;
		var nLastNoZero = nLength - 1;
		for(var i = val.length - 1; i >= 0; --i)
		{
			nLastNoZero = i;
			if("0" != val[i])
				break;
		}
		if(nLastNoZero != nLength - 1)
		{
		    if (cultureInfo.NumberDecimalSeparator == res[nLastNoZero])
				res = res.substring(0, nLastNoZero);
			else
				res = res.substring(0, nLastNoZero + 1);
		}
		return res;
	}
};
var oGeneralEditFormatCache = new GeneralEditFormatCache();

function FormatParser()
{
	this.days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
	this.daysLeap = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
}
FormatParser.prototype =
{
    isLocaleNumber: function (val, cultureInfo) {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        //javascript decimal separator is '.'
        if ("." != cultureInfo.NumberDecimalSeparator) {
            val = val.replace(".", "q");//заменяем на символ с которым не распознается, как в Excel
            val = val.replace(cultureInfo.NumberDecimalSeparator, ".");
        }
        //parseNum исключаем запись числа в 16-ричной форме из числа.
        return AscCommonExcel.parseNum(val) && Asc.isNumberInfinity(val);
    },
    parseLocaleNumber: function (val, cultureInfo) {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        //javascript decimal separator is '.'
        if ("." != cultureInfo.NumberDecimalSeparator) {
            val = val.replace(".", "q");//заменяем на символ с которым не распознается, как в Excel
            val = val.replace(cultureInfo.NumberDecimalSeparator, ".");
        }
        return val - 0;
    },
    parse: function (value, cultureInfo)
    {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        var res = null;
        var bError = false;
        //replace Non-breaking space(0xA0) with White-space(0x20)
        if (" " == cultureInfo.NumberGroupSeparator)
            value = value.replace(new RegExp(String.fromCharCode(0xA0), "g"));
        var rx_thouthand = new RegExp("^(([ \\+\\-%\\$€£¥\\(]|" + escapeRegExp(cultureInfo.CurrencySymbol) + ")*)((\\d+" + escapeRegExp(cultureInfo.NumberGroupSeparator) + "\\d+)*\\d*" + escapeRegExp(cultureInfo.NumberDecimalSeparator) + "?\\d*)(([ %\\)]|р.|" + escapeRegExp(cultureInfo.CurrencySymbol) + ")*)$");
        var match = value.match(rx_thouthand);
        if (null != match) {
            var sBefore = match[1];
            var sVal = match[3];
            var sAfter = match[5];
			var oChartCount = {};
			if(null != sBefore)
			    this._parseStringLetters(sBefore, cultureInfo.CurrencySymbol, true, oChartCount);
			if(null != sAfter)
			    this._parseStringLetters(sAfter, cultureInfo.CurrencySymbol, false, oChartCount);
			var bMinus = false;
			var bPercent = false;
			var sCurrency = null;
			var oCurrencyElem = null;
			var nBracket = 0;
			for(var sChar in oChartCount){
				var elem = oChartCount[sChar];
				if(" " == sChar)
					continue;
				else if("+" == sChar){
					if(elem.all > 1)
						bError = true;
				}
				else if("-" == sChar){
					if(elem.all > 1)
						bError = true;
					else
						bMinus = true;
				}
				else if("-" == sChar){
					if(elem.all > 1)
						bError = true;
					else
						bMinus = true;
				}
				else if("(" == sChar){
					if(1 == elem.all && 1 == elem.before)
						nBracket++;
					else
						bError = true;
				}
				else if(")" == sChar){
					if(1 == elem.all && 1 == elem.after)
						nBracket++;
					else
						bError = true;
				}
				else if("%" == sChar){
					if(1 == elem.all)
						bPercent = true;
					else
						bError = true;
				}
				else{
					if(null == sCurrency && 1 == elem.all){
						sCurrency = sChar;
						oCurrencyElem = elem;
					}
					else
						bError = true;
				}
			}
			if (nBracket > 0) {
			    if (2 == nBracket)
			        bMinus = true;
			    else
			        bError = true;
			}
			var CurrencyNegativePattern = cultureInfo.CurrencyNegativePattern;
			if(null != sCurrency){
			    if (sCurrency == cultureInfo.CurrencySymbol) {
			        var nPattern = cultureInfo.CurrencyNegativePattern;
			        if (0 == nPattern || 1 == nPattern || 2 == nPattern || 3 == nPattern || 9 == nPattern || 11 == nPattern || 12 == nPattern || 14 == nPattern) {
			            if (1 != oCurrencyElem.before)
			                bError = true;
			        }
			        else if (1 != oCurrencyElem.after)
			            bError = true;
			    }
			    else if(-1 != "$€£¥".indexOf(sCurrency)){
			        if (1 == oCurrencyElem.before) {
			            CurrencyNegativePattern = 0;
			        }
                    else
						bError = true;
				}
				else if(-1 != "р.".indexOf(sCurrency)){
				    if (1 == oCurrencyElem.after) {
				        CurrencyNegativePattern = 5;
				    }
                    else
						bError = true;
				}
				else
				    bError = true;
			}
			if(!bError){
				var oVal = this._parseThouthand(sVal, cultureInfo);
				if (oVal) {
					res = {format: null, value: null, bDateTime: false, bDate: false, bTime: false, bPercent: false, bCurrency: false};
					var dVal = oVal.number;
					if (bMinus)
						dVal = -dVal;
					var sFracFormat = "";
					if (parseInt(dVal) != dVal)
						sFracFormat = gc_sFormatDecimalPoint + "00";
					var sFormat = null;
					if (bPercent) {
						res.bPercent = true;
						dVal /= 100;
						sFormat = "0" + sFracFormat + "%";
					}
					else if (sCurrency) {
						res.bCurrency = true;
					    var sNumberFormat = "#" + gc_sFormatThousandSeparator + "##0" + sFracFormat;
					    var sCurrencyFormat;
					    if(sCurrency.length > 1)
					        sCurrencyFormat = "\"" + sCurrency + "\"";
					    else
					        sCurrencyFormat = "\\" + sCurrency;
					    var sPositivePattern;
					    var sNegativePattern;
					    switch (CurrencyNegativePattern) {
					        case 0:
					            sPositivePattern = sCurrencyFormat + sNumberFormat + "_)";
					            sNegativePattern = "[Red](" + sCurrencyFormat + sNumberFormat + ")";
					            break;
					        case 1:
					            sPositivePattern = sCurrencyFormat + sNumberFormat;
					            sNegativePattern = "[Red]-" + sCurrencyFormat + sNumberFormat;
					            break;
					        case 2:
					            sPositivePattern = sCurrencyFormat + sNumberFormat;
					            sNegativePattern = "[Red]" + sCurrencyFormat + "-" + sNumberFormat;
					            break;
					        case 3:
					            sPositivePattern = sCurrencyFormat + sNumberFormat + "_-";
					            sNegativePattern = "[Red]" + sCurrencyFormat + sNumberFormat + "-";
					            break;
					        case 4:
					            sPositivePattern = sNumberFormat + sCurrencyFormat + "_)";
					            sNegativePattern = "[Red](" + sNumberFormat + sCurrencyFormat + ")";
					            break;
					        case 5:
					            sPositivePattern = sNumberFormat + sCurrencyFormat;
					            sNegativePattern = "[Red]-" + sNumberFormat + sCurrencyFormat;
					            break;
					        case 6:
					            sPositivePattern = sNumberFormat + "-" + sCurrencyFormat;
					            sNegativePattern = "[Red]" + sNumberFormat + "-" + sCurrencyFormat;
					            break;
					        case 7:
					            sPositivePattern = sNumberFormat + sCurrencyFormat + "_-";
					            sNegativePattern = "[Red]" + sNumberFormat + sCurrencyFormat + "-";
					            break;
					        case 8:
					            sPositivePattern = sNumberFormat + " " + sCurrencyFormat;
					            sNegativePattern = "[Red]-" + sNumberFormat + " " + sCurrencyFormat;
					            break;
					        case 9:
					            sPositivePattern = sCurrencyFormat + " " + sNumberFormat;
					            sNegativePattern = "[Red]-" + sCurrencyFormat + " " + sNumberFormat;
					            break;
					        case 10:
					            sPositivePattern = sNumberFormat + " " + sCurrencyFormat + "_-";
					            sNegativePattern = "[Red]" + sNumberFormat + " " + sCurrencyFormat + "-";
					            break;
					        case 11:
					            sPositivePattern = sCurrencyFormat + " " + sNumberFormat + "_-";
					            sNegativePattern = "[Red]" + sCurrencyFormat + " " + sNumberFormat + "-";
					            break;
					        case 12:
					            sPositivePattern = sCurrencyFormat + " " + sNumberFormat;
					            sNegativePattern = "[Red]" + sCurrencyFormat + " -" + sNumberFormat;
					            break;
					        case 13:
					            sPositivePattern = sNumberFormat + " " + sCurrencyFormat;
					            sNegativePattern = "[Red]" + sNumberFormat + "- " + sCurrencyFormat;
					            break;
					        case 14:
					            sPositivePattern = sCurrencyFormat + " " + sNumberFormat + "_)";
					            sNegativePattern = "[Red](" + sCurrencyFormat + " " + sNumberFormat + ")";
					            break;
					        case 15:
					            sPositivePattern = sNumberFormat + " " + sCurrencyFormat + "_)";
					            sNegativePattern = "[Red](" + sNumberFormat + " " + sCurrencyFormat + ")";
					            break;
					    }
					    sFormat = sPositivePattern + ";" + sNegativePattern;
					}
					else if (oVal.thouthand) {
						sFormat = "#" + gc_sFormatThousandSeparator + "##0" + sFracFormat;
					}
					else
						sFormat = AscCommon.g_cGeneralFormat;
					res.format = sFormat;
					res.value = dVal;
				}
			}
        }
        if (null == res && !bError)
            res = this.parseDate(value, cultureInfo);
        return res;
    },
    _parseStringLetters: function (sVal, currencySymbol, bBefore, oRes) {
        //отдельно обрабатываем 'р.' и currencySymbol потому что они могут быть не односимвольными
        var aTemp = ["р.", currencySymbol];
        for (var i = 0, length = aTemp.length; i < length; i++){
            var sChar = aTemp[i];
            var nIndex = -1;
            var nCount = 0;
            while(-1 != (nIndex = sVal.indexOf(sChar, nIndex + 1)))
                nCount++;
            if(nCount > 0)
            {
                sVal = sVal.replace(new RegExp(escapeRegExp(sChar), "g"), "");
                var elem = oRes[sChar];
                if(!elem){
                    elem = {before: 0, after: 0, all: 0};
                    oRes[sChar] = elem;
                }
                if(bBefore)
                    elem.before += nCount;
                else
                    elem.after += nCount;
                elem.all += nCount;
            }
        }
		for(var i = 0, length = sVal.length; i < length; i++){
			var sChar = sVal[i];
			var elem = oRes[sChar];
			if(!elem){
				elem = {before: 0, after: 0, all: 0};
				oRes[sChar] = elem;
			}
			if(bBefore)
				elem.before++;
			else
				elem.after++;
			elem.all++;
		}
	},
    _parseThouthand: function (val, cultureInfo)
    {
        var oRes = null;
        var bThouthand = false;
        //reverse
        var sReverseVal = "";
        for (var i = val.length - 1; i >= 0; --i)
            sReverseVal += val[i];
        var nGroupSizeIndex = 0;
        var nGroupSize = cultureInfo.NumberGroupSizes[nGroupSizeIndex];
        var nPrevIndex = 0;
        var nIndex = -1;
        var bError = false;
        while (-1 != (nIndex = sReverseVal.indexOf(cultureInfo.NumberGroupSeparator, nIndex + 1))) {
            var nCurLength = nIndex - nPrevIndex;
            if (nCurLength < nGroupSize) {
                bError = true;
                break;
            }
            if (nGroupSizeIndex < cultureInfo.NumberGroupSizes.length - 1) {
                nGroupSizeIndex++;
                nGroupSize = cultureInfo.NumberGroupSizes[nGroupSizeIndex];
            }
            nPrevIndex = nIndex + 1;
        }
        if (!bError) {
            if (0 != nPrevIndex) {
                //чтобы не распознавалось 0,001
                if (nPrevIndex < val.length && parseInt(val.substr(0, val.length - nPrevIndex)) > 0) {
                    val = val.replace(new RegExp(escapeRegExp(cultureInfo.NumberGroupSeparator), "g"), '');
                    bThouthand = true;
                }
            }
			if (g_oFormatParser.isLocaleNumber(val, cultureInfo)) {
				var dNumber = g_oFormatParser.parseLocaleNumber(val, cultureInfo);
				oRes = { number: dNumber, thouthand: bThouthand };
			}
        }
		return oRes;
	},
    _parseDateFromArray: function (match, oDataTypes, cultureInfo)
	{
        var res = null;
        var bError = false;
        //в первый проход разделяем date и time с помощью delimiter
        for (var i = 0, length = match.length; i < length; i++) {
            var elem = match[i];
            if (elem.type == oDataTypes.delimiter) {
                bError = true;
                if(i - 1 >= 0 && i + 1 < length){
                    var prev = match[i - 1];
                    var next = match[i + 1];
                    if(prev.type != oDataTypes.delimiter && next.type != oDataTypes.delimite){
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if(false == prev.date && false == next.date){
                                bError = false;
                                prev.time = true;
                                next.time = true;
                            }
                        }
                        else{
                            if(false == prev.time && false == next.time){
                                bError = false;
                                prev.date = true;
                                next.date = true;
                            }
                        }
                    }
                }
                else if (i - 1 >= 0 && i + 1 == length) {
                    //случай "10:"
                    var prev = match[i - 1];
                    if (prev.type != oDataTypes.delimiter) {
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if (false == prev.date) {
                                bError = false;
                                prev.time = true;
                            }
                        }
                    }
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            //разделяем date и time с помощью Am/Pm и имена месяцев
            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.type == oDataTypes.letter){
                    var valLower = elem.val.toLowerCase();
                    if (elem.am || elem.pm) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.date) {
                                prev.time = true;
                            }
                        }
                        //AmPm должна быть последней записью
                        if (i + 1 != length) {
                            bError = true;
                        }
                    }
                    else if (null != elem.month) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.time)
                                prev.date = true;
                        }
                        if (i + 1 < length) {
                            let next = match[i + 1]
                            // processing the option when the date is given as the format "October 11, 2008"
                            if (i === 0 && i + 2 < length) {
                                let afterNext = match[i + 2];
                                if (oDataTypes.digit == afterNext.type && false == afterNext.time) {
                                    afterNext.date = true;
                                }
                            }
                            if (oDataTypes.digit == next.type && false == next.time)
                                next.date = true;
                        }
                    }
                    else
                        bError = true;
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            var aDate = [];
            var nMonthIndex = null;
			var sMonthFormat = null;
            var aTime = [];
            var am = false;
            var pm = false;

            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.date) {
                    if (elem.type == oDataTypes.digit)
                        aDate.push(elem.val);
                    else if (elem.type == oDataTypes.letter && null != elem.month) {
                        nMonthIndex = aDate.length;
                        sMonthFormat = elem.month.format;
                        aDate.push(elem.month.val);
                    }
                    else
                        bError = true;
                }
                else if (elem.time) {
                    if (elem.type == oDataTypes.digit)
                        aTime.push(elem.val);
                    else if (elem.type == oDataTypes.letter && (elem.am || elem.pm)) {
                        am = elem.am;
                        pm = elem.pm;
                    }
                    else
                        bError = true;
                }
                else if (oDataTypes.digit == elem.type)
                    bError = true;//случай "1-2-3 10"
            }
            var nDateLength = aDate.length;
            if (nDateLength > 0 && !(2 <= nDateLength && nDateLength <= 3 && (null == nMonthIndex || (3 == nDateLength && 1 == nMonthIndex) || 2 == nDateLength || (3 == nDateLength && 0 == nMonthIndex))))
                bError = true;
            var nTimeLength = aTime.length;
            if (nTimeLength > 3)
                bError = true;
            if(!bError){
                res = { d: null, m: null, y: null, h: null, min: null, s: null, am: am, pm: pm, sDateFormat: null };
                if (nDateLength > 0) {
                    var nIndexD = Math.max(cultureInfo.ShortDatePattern.indexOf("0"), cultureInfo.ShortDatePattern.indexOf("1"));
                    var nIndexM = Math.max(cultureInfo.ShortDatePattern.indexOf("2"), cultureInfo.ShortDatePattern.indexOf("3"));
                    var nIndexY = Math.max(cultureInfo.ShortDatePattern.indexOf("4"), cultureInfo.ShortDatePattern.indexOf("5"));
                    if (null != nMonthIndex) {
                        if (2 == nDateLength) {
                            res.d = aDate[nDateLength - 1 - nMonthIndex];
                            res.m = aDate[nMonthIndex];
                            //приоритет у формата d-mmm, но если он не подходит пробуем сделать mmm-yy
                            if (this.isValidDate((new Date()).getFullYear(), res.m - 1, res.d))
                                res.sDateFormat = "d-mmm";
                            else {
                                //не в классическом случае(!= dd/mm/yyyy) меняем местами d и m перед тем как пробовать y
                                if (!isDMY(cultureInfo) && this.isValidDate((new Date()).getFullYear(), res.d - 1, res.m)) {
                                    res.sDateFormat = "d-mmm";
                                    var temp = res.d;
                                    res.d = res.m;
                                    res.m = temp;
                                }
                                else {
                                    //если текстовый месяц стоит вторым, то первый параметр может быть только днем
                                    if (0 == nMonthIndex) {
                                        res.sDateFormat = "mmm-yy";
                                        res.d = null;
                                        res.m = aDate[0];
                                        res.y = aDate[1];
                                    }
                                    else
                                        bError = true;
                                }
                            }
                        } else {
                            if (nMonthIndex == 0) {
                                res.sDateFormat = "dd-mmm-yy";
                                res.m = aDate[0];
                                res.d = aDate[1];
                                res.y = aDate[2];
                            } else {
                                res.sDateFormat = "d-mmm-yy";
                                res.d = aDate[0];
                                res.m = aDate[1];
                                res.y = aDate[2];
                            }
                        }
                    }
                    else {
                        //смотрим порядок в default формат
                        if (2 == nDateLength) {
                            //в приоритете d и m
                            if (nIndexD < nIndexM) {
                                res.d = aDate[0];
                                res.m = aDate[1];
                            }
                            else {
                                res.m = aDate[0];
                                res.d = aDate[1];
                            }
                            if (this.isValidDate((new Date()).getFullYear(), res.m - 1, res.d))
                                res.sDateFormat = "d-mmm";
                            else{
                                //в обратной записи(== yyyy/mm/dd) меняем местами d и m перед тем как пробовать y
                                if (isYMD(cultureInfo) && this.isValidDate((new Date()).getFullYear(), res.d - 1, res.m)) {
                                    res.sDateFormat = "d-mmm";
                                    var temp = res.d;
                                    res.d = res.m;
                                    res.m = temp;
                                }
                                else{
                                    res.sDateFormat = "mmm-yy";
                                    res.d = null;
                                    if (nIndexM < nIndexY) {
                                        res.m = aDate[0];
                                        res.y = aDate[1];
                                    }
                                    else {
                                        res.y = aDate[0];
                                        res.m = aDate[1];
                                    }
                                }
                            }
                        } else if(3 == nDateLength && aDate[0] > 1000) {
                            res.y = aDate[0];
                            res.m = aDate[1];
                            res.d = aDate[2];
                            res.sDateFormat = getShortDateFormat(cultureInfo);
                        } else {
                            for (var i = 0, length = cultureInfo.ShortDatePattern.length; i < length; i++)
                            {
                                var nIndex = cultureInfo.ShortDatePattern[i] - 0;
                                var val = aDate[i];
                                if (0 == nIndex || 1 == nIndex) {
                                    res.d = val;
                                } else if (2 == nIndex || 3 == nIndex) {
                                    res.m = val;
                                } else if (4 == nIndex || 5 == nIndex) {
                                    res.y = val;
                                }
                            }
                            res.sDateFormat = getShortDateFormat(cultureInfo);
                        }
                    }
                    if(null != res.y)
                    {
                        if(res.y < 30)
                            res.y = 2000 + res.y;
                        else if(res.y < 100)
                            res.y = 1900 + res.y;
                    }
                }
                if(nTimeLength > 0){
                    res.h = aTime[0];
                    if(nTimeLength > 1)
                        res.min = aTime[1];
                    if(nTimeLength > 2)
                        res.s = aTime[2];
                }
                if(bError)
                    res = null;
            }
        }
		return res;
    },
	_parseDateFromArrayPDF: function (match, oDataTypes, cultureInfo, oFormat)
	{
        var res = null;
        var bError = false;
        //в первый проход разделяем date и time с помощью delimiter
        for (var i = 0, length = match.length; i < length; i++) {
            var elem = match[i];
            if (elem.type == oDataTypes.delimiter) {
                bError = true;
                if(i - 1 >= 0 && i + 1 < length){
                    var prev = match[i - 1];
                    var next = match[i + 1];
                    if(prev.type != oDataTypes.delimiter && next.type != oDataTypes.delimite){
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if(false == prev.date && false == next.date){
                                bError = false;
                                prev.time = true;
                                next.time = true;
                            }
                        }
                        else{
                            if(false == prev.time && false == next.time){
                                bError = false;
                                prev.date = true;
                                next.date = true;
                            }
                        }
                    }
                }
                else if (i - 1 >= 0 && i + 1 == length) {
                    //случай "10:"
                    var prev = match[i - 1];
                    if (prev.type != oDataTypes.delimiter) {
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if (false == prev.date) {
                                bError = false;
                                prev.time = true;
                            }
                        }
                    }
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            //разделяем date и time с помощью Am/Pm и имена месяцев
            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.type == oDataTypes.letter){
                    var valLower = elem.val.toLowerCase();
                    if (elem.am || elem.pm) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.date) {
                                prev.time = true;
                            }
                        }
                        //AmPm должна быть последней записью
                        if (i + 1 != length) {
                            bError = true;
                        }
                    }
                    else if (null != elem.month) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.time)
                                prev.date = true;
                        }
                        if (i + 1 < length) {
                            var next = match[i + 1];
                            if (oDataTypes.digit == next.type && false == next.time)
                                next.date = true;
                        }
                    }
                    else
                        bError = true;
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            var aDate = [];
            var nMonthIndex = null;
			var sMonthFormat = null;
			var monthDone = false;
            var aTime = [];
            var am = false;
            var pm = false;

			var nIndexD = Math.max(cultureInfo.ShortDatePattern.indexOf("0"), cultureInfo.ShortDatePattern.indexOf("1"));
			var nIndexM = Math.max(cultureInfo.ShortDatePattern.indexOf("2"), cultureInfo.ShortDatePattern.indexOf("3"));
            var nIndexY = Math.max(cultureInfo.ShortDatePattern.indexOf("4"), cultureInfo.ShortDatePattern.indexOf("5"));

            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.date || (elem.time == false && elem.type == oDataTypes.digit)) {
                    if (elem.type == oDataTypes.digit)
                        aDate.push(elem.val);
                    else if (elem.type == oDataTypes.letter && null != elem.month) {
                        if (aDate.length >= 3)
							continue;
							
						nMonthIndex = aDate.length;
                        sMonthFormat = elem.month.format;
                        aDate.push(elem.month.val);
						monthDone = true;
                    }
                    else
                        bError = true;
                }
                else if (elem.time) {
                    if (elem.type == oDataTypes.digit)
                        aTime.push(elem.val);
                    else if (elem.type == oDataTypes.letter && (elem.am || elem.pm)) {
                        am = elem.am;
                        pm = elem.pm;
                    }
                    else
                        bError = true;
                }
            }
			if (aDate.length > 3)
				aDate.length = 3;

            var nDateLength = aDate.length;
            var nTimeLength = aTime.length;
            if (nTimeLength > 3)
                aTime.length = 3;
            if(!bError){
                res = { d: null, m: null, y: null, h: null, min: null, s: null, am: am, pm: pm, sDateFormat: null };
                if (nDateLength > 0) {
                    if (null != nMonthIndex) {
                        res.m = aDate[nMonthIndex];

						if (nIndexD != -1) {
							if (nIndexD != nMonthIndex) {
								res.d = aDate[nIndexD];
							}
							else {
								if (aDate[0] <= 31) {
									res.d = aDate[0];
									res.y = aDate[2];
								}
								else {
									res.d = aDate[2];
									res.y = aDate[0];
								}
							}
						}
						
						if (nIndexY != -1 && res.y == null) {
							if (nIndexY != nMonthIndex) {
								res.y = aDate[nIndexY];
							}
							else {
								res.d = aDate[0];
								res.y = aDate[2];
							}
						}
                    }
                    else {
                        res.m = aDate[nIndexM];
						res.d = aDate[nIndexD];
						res.y = aDate[nIndexY];
                    }
                    if(null != res.y)
                    {
                        if(res.y < 30)
                            res.y = 2000 + res.y;
                        else if(res.y < 100)
                            res.y = 1900 + res.y;
                    }
                }
                if(nTimeLength > 0){
                    res.h = aTime[0];
                    if(nTimeLength > 1)
                        res.min = aTime[1];
                    if(nTimeLength > 2)
                        res.s = aTime[2];
                }
                if(bError)
                    res = null;
            }
        }
		return res;
    },
    strcmp: function (s1, s2, index1, length, index2) {
        if (null == index2)
            index2 = 0;
        var bRes = true;
        for (var i = 0; i < length; ++i) {
            if (s1[index1 + i] != s2[index2 + i]) {
                bRes = false;
                break;
            }
        }
        return length === 0 ? false: bRes;
    },
	parseDate: function (value, cultureInfo)
	{
		//todo "11: AM" should fail
		var res = null;
		var match = [];
		var sCurValue = null;
		var oCurDataType = null;
		var oPrevType = null;
		var bAmPm = false;
		var bMonth = false;
		var bError = false;
		var oDataTypes = {letter: {id: 0, min: 2, max: 9}, digit: {id: 1, min: 1, max: 4}, delimiter: {id: 2, min: 1, max: 1}, space: {id: 3, min: null, max: null}};
		var valueLower = value.toLowerCase();
		for(var i = 0, length = value.length; i < length; i++)
		{
		    var sChar = value[i];
		    var oDataType = null;
		    if("0" <= sChar && sChar <= "9")
		        oDataType = oDataTypes.digit;
		    else if(" " == sChar || "," == sChar)
		        oDataType = oDataTypes.space;
		    else if ("/" == sChar || "-" == sChar || ":" == sChar || cultureInfo.DateSeparator == sChar || cultureInfo.TimeSeparator == sChar)
		        oDataType = oDataTypes.delimiter;
		    else
		        oDataType = oDataTypes.letter;
			    
		    if(null != oDataType)
		    {
		        if(null == oCurDataType)
		            sCurValue = sChar;
		        else
		        {
		            if(oCurDataType == oDataType)
		            {
		                if(null == oCurDataType.max || sCurValue.length < oCurDataType.max)
		                    sCurValue += sChar;
		                else
		                    bError = true;
		            }
		            else
		            {
		                if (null == oCurDataType.min || sCurValue.length >= oCurDataType.min) {
		                    if (oDataTypes.space != oCurDataType) {
		                        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		                        if (oDataTypes.digit == oCurDataType)
		                            oNewElem.val = oNewElem.val - 0;
		                        match.push(oNewElem);
		                    }
		                    sCurValue = sChar;
		                    oPrevType = oCurDataType;
		                }
		                else
		                    bError = true;
		            }
		        }
		        oCurDataType = oDataType;
		    }
		    else
		        bError = true;
		    if(oDataTypes.letter == oDataType){
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        var bAm = false;
		        var bPm = false;
		        if (!bAmPm && ((bAm = this.strcmp(valueLower, "am", i, 2)) || (bPm = this.strcmp(valueLower, "pm", i, 2)))) {
		            bAmPm = true;
		            oNewElem.am = bAm;
		            oNewElem.pm = bPm;
		            oNewElem.time = true;
		            match.push(oNewElem);
		            i += 2 - 1;
		            if (oPrevType != oDataTypes.space)
		                bError = true;
		        }
		        else if (!bMonth) {
		            bMonth = true;
					let aArraysToCheck = [{ arr: cultureInfo.MonthNames, format: "mmmm" }, { arr: cultureInfo.AbbreviatedMonthNames, format: "mmm" }];
		            var bFound = false;
		            for (var index in aArraysToCheck) {
		                var aArrayTemp = aArraysToCheck[index];
		                for (var j = 0, length2 = aArrayTemp.arr.length; j < length2; j++) {
		                    var sCmpVal = aArrayTemp.arr[j].toLowerCase();
		                    var sCmpValCrop = sCmpVal.replace(/\./g, "");
		                    var bCrop = false;
		                    if (this.strcmp(valueLower, sCmpVal, i, sCmpVal.length) || (bCrop = (sCmpVal != sCmpValCrop && this.strcmp(valueLower, sCmpValCrop, i, sCmpValCrop.length)))) {
		                        bFound = true;
		                        oNewElem.month = { val: j + 1, format: aArrayTemp.format };
		                        oNewElem.date = true;
		                        if (bCrop)
		                            i += sCmpValCrop.length - 1;
		                        else
		                            i += sCmpVal.length - 1;
		                        break;
		                    }
		                }
		                if (bFound)
		                    break;
		            }
		            //ничего кроме имени месяца больше быть не может
		            if (bFound)
		                match.push(oNewElem);
		            else
		                bError = true;
		        }
		        else
		            bError = true;
		        oCurDataType = null;
		        sCurValue = null;
		    }
			if (bError)
			{
				match = null;
				break;
			}
		}
		if (null != match && null != sCurValue) {
		    if (oDataTypes.space != oCurDataType) {
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        if (oDataTypes.digit == oCurDataType)
		            oNewElem.val = oNewElem.val - 0;
		        match.push(oNewElem);
		    }
		}
		if(null != match && match.length > 0)
		{
		    var oParsedDate = this._parseDateFromArray(match, oDataTypes, cultureInfo);
			if(null != oParsedDate)
			{
				var d = oParsedDate.d;
				var m = oParsedDate.m;
				var y = oParsedDate.y;
				var h = oParsedDate.h;
				var min = oParsedDate.min;
				var s = oParsedDate.s;
				var am = oParsedDate.am;
				var pm = oParsedDate.pm;
				var sDateFormat = oParsedDate.sDateFormat;
				
				var bDate = false;
				var bTime = false;
				var bSeconds = false;
				var nDay;
				var nMounth;
				var nYear;
				if(AscCommon.bDate1904)
				{
					nDay = 1;
					nMounth = 0;
					nYear = 1904;
				}
				else
				{
					nDay = 31;
					nMounth = 11;
					nYear = 1899;
				}
				var nHour = 0;
				var nMinute = 0;
				var nSecond = 0;
				var dValue = 0;
				var bValidDate = true;
				if(null != m && (null != d || null != y))
				{
					bDate = true;
					var oNowDate;
					if(null != d)
						nDay = d - 0;
					else
						nDay = 1;
					nMounth = m - 1;
					if(null != y)
						nYear = y - 0;
					else
                    {
                        oNowDate = new Date();
						nYear = oNowDate.getFullYear();
                    }
					
					//проверяем дату на валидность
					bValidDate = this.isValidDate(nYear, nMounth, nDay);
				}
				if(null != h)
				{
					bTime = true;
					nHour = h - 0;
					if (am || pm)
					{
						if(nHour <= 23)
						{
							//переводим 24
							nHour = nHour % 12;
							if(pm)
								nHour += 12;
						}
						else
							bValidDate = false;
					}
					if(null != min)
					{
						nMinute = min - 0;
						if(nMinute > 59)
							bValidDate = false;
					}
					if(null != s)
					{
						nSecond = s - 0;
						if (0 <= nSecond && nSecond < 60) {
							bSeconds = true;
						} else {
							bValidDate = false;
						}
					}
				}
				if(true == bValidDate && (true == bDate || true == bTime))
				{
					if(AscCommon.bDate1904)
						dValue = (Date.UTC(nYear,nMounth,nDay,nHour,nMinute,nSecond) - Date.UTC(1904,0,1,0,0,0)) / (86400 * 1000);
					else
					{
						if(1900 < nYear || (1900 == nYear && 1 < nMounth ))
							dValue = (Date.UTC(nYear,nMounth,nDay,nHour,nMinute,nSecond) - Date.UTC(1899,11,30,0,0,0)) / (86400 * 1000);
						else if(1900 == nYear && 1 == nMounth && 29 == nDay)
							dValue = 60;
						else
							dValue = (Date.UTC(nYear,nMounth,nDay,nHour,nMinute,nSecond) - Date.UTC(1899,11,31,0,0,0)) / (86400 * 1000);
					}
					if(dValue >= 0)
					{
						var sFormat = "";
						if (bDate) {
							if (bTime && nHour > 23) {
								sFormat = AscCommon.g_cGeneralFormat;
							} else {
								sFormat += sDateFormat;
								if (bTime) {
									sFormat += " h:mm";
								}
							}
						} else {
							if (dValue > 1) {
								sFormat += "[h]:mm";
							} else {
								sFormat += "h:mm";
							}
							if (bSeconds || dValue > 1) {
								sFormat += ":ss";
							}
							if (am || pm)
								sFormat += " AM/PM";
						}
						res = {format: sFormat, value: dValue, bDateTime: true, bDate: bDate, bTime: bTime, bPercent: false, bCurrency: false};
					}
				}
            }
        }
		return res;
	},
	parseDatePDF: function (value, cultureInfo, oFormat)
	{
		let res = null;
		let match = [];
		let sCurValue = null;
		let oCurDataType = null;
		let oPrevType = null;
		let bAmPm = false;
		let bMonth = false;
		let bError = false;
		let oDataTypes = {letter: {id: 0, min: 2, max: 9}, digit: {id: 1, min: 1, max: 4}, delimiter: {id: 2, min: 1, max: 1}, space: {id: 3, min: null, max: null}};
		let valueLower = value.toLowerCase();
		for(var i = 0, length = value.length; i < length; i++)
		{
		    var sChar = value[i];
		    var oDataType = null;
		    if("0" <= sChar && sChar <= "9")
		        oDataType = oDataTypes.digit;
		    else if(" " == sChar)
		        oDataType = oDataTypes.space;
		    else if ("." == sChar || "/" == sChar || "-" == sChar || ":" == sChar || "," == sChar || cultureInfo.DateSeparator == sChar || cultureInfo.TimeSeparator == sChar)
		        oDataType = oDataTypes.delimiter;
		    else
		        oDataType = oDataTypes.letter;
			    
			// после разделителя может быть опять месяц
			if (oDataType == oDataTypes.delimiter)
				bMonth = false;

		    if(null != oDataType)
		    {
		        if(null == oCurDataType)
		            sCurValue = sChar;
		        else
		        {
		            if(oCurDataType == oDataType)
		            {
		                if(null == oCurDataType.max || sCurValue.length < oCurDataType.max)
		                    sCurValue += sChar;
		                else
		                    bError = true;
		            }
		            else
		            {
		                if (null == oCurDataType.min || sCurValue.length >= oCurDataType.min) {
		                    if (oDataTypes.space != oCurDataType) {
		                        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		                        if (oDataTypes.digit == oCurDataType)
		                            oNewElem.val = oNewElem.val - 0;
								if (oNewElem.val < 100 && sCurValue.length == 4)
									bError = true; // год до ста лет, пример: 0001 год
		                        
								match.push(oNewElem);
		                    }
		                    sCurValue = sChar;
		                    oPrevType = oCurDataType;
		                }
		                else
		                    bError = true;
		            }
		        }
		        oCurDataType = oDataType;
		    }
		    else
		        bError = true;
		    if(oDataTypes.letter == oDataType){
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        var bAm = false;
		        var bPm = false;
		        if (!bAmPm && ((bAm = this.strcmp(valueLower, "am", i, 2)) || (bPm = this.strcmp(valueLower, "pm", i, 2)))) {
		            bAmPm = true;
		            oNewElem.am = bAm;
		            oNewElem.pm = bPm;
		            oNewElem.time = true;
		            match.push(oNewElem);
		            i += 2 - 1;
		            if (oPrevType != oDataTypes.space)
		                bError = true;
		        }
		        else if (!bMonth) {
		            bMonth = true;
		            var aArraysToCheck = [{ arr: cultureInfo.MonthNames, format: "mmmm" }, { arr: cultureInfo.AbbreviatedMonthNames, format: "mmm" }];
		            var bFound = false;
		            for (var index in aArraysToCheck) {
		                var aArrayTemp = aArraysToCheck[index];
		                for (var j = 0, length2 = aArrayTemp.arr.length; j < length2; j++) {
		                    var sCmpVal = aArrayTemp.arr[j].toLowerCase();
		                    var sCmpValCrop = sCmpVal.replace(/\./g, "");
		                    var bCrop = false;
		                    if (this.strcmp(valueLower, sCmpVal, i, sCmpVal.length) || (bCrop = (sCmpVal != sCmpValCrop && this.strcmp(valueLower, sCmpValCrop, i, sCmpValCrop.length)))) {
		                        bFound = true;
		                        oNewElem.month = { val: j + 1, format: aArrayTemp.format };
		                        oNewElem.date = true;
		                        if (bCrop)
		                            i += sCmpValCrop.length - 1;
		                        else
		                            i += sCmpVal.length - 1;
		                        break;
		                    }
		                }
		                if (bFound)
		                    break;
		            }
		            //ничего кроме имени месяца больше быть не может
		            if (bFound)
		                match.push(oNewElem);
		            else
		                bError = true;
		        }
		        else
		            bError = true;
		        oCurDataType = null;
		        sCurValue = null;
		    }
			if (bError)
			{
				match = null;
				break;
			}
		}
		if (null != match && null != sCurValue) {
		    if (oDataTypes.space != oCurDataType) {
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        if (oDataTypes.digit == oCurDataType)
		            oNewElem.val = oNewElem.val - 0;

		        match.push(oNewElem);
		    }
		}
		if(null != match && match.length > 0)
		{
		    var oParsedDate = this._parseDateFromArrayPDF(match, oDataTypes, cultureInfo, oFormat);
			if(null != oParsedDate)
			{
				var d = oParsedDate.d;
				var m = oParsedDate.m;
				var y = oParsedDate.y;
				var h = oParsedDate.h;
				var min = oParsedDate.min;
				var s = oParsedDate.s;
				var am = oParsedDate.am;
				var pm = oParsedDate.pm;
				var sDateFormat = oParsedDate.sDateFormat;
				
				var bDate = false;
				var bTime = false;
				var nDay;
				var nMounth;
				var nYear;
				if(AscCommon.bDate1904)
				{
					nDay = 1;
					nMounth = 0;
					nYear = 1904;
				}
				else
				{
					nDay = 31;
					nMounth = 11;
					nYear = 1899;
				}
				var nHour = 0;
				var nMinute = 0;
				var nSecond = 0;
				var dValue = 0;
				var bValidDate = true;
				if(null != m && (null != d || null != y))
				{
					bDate = true;
					var oNowDate;
					if(null != d)
						nDay = d - 0;
					else
						nDay = 1;
					nMounth = m - 1;
					if(null != y)
						nYear = y - 0;
					else
                    {
                        oNowDate = new Date();
						nYear = oNowDate.getFullYear();
                    }
					
					//проверяем дату на валидность
					bValidDate = this.isValidDatePDF(nYear, nMounth, nDay);
				}
				if(null != h)
				{
					bTime = true;
					nHour = h - 0;
					if (am || pm)
					{
						if(nHour <= 23)
						{
							//переводим 24
							nHour = nHour % 12;
							if(pm)
								nHour += 12;
						}
						else
							bValidDate = false;
					}
					if(null != min)
					{
						nMinute = min - 0;
						if(nMinute > 59)
							bValidDate = false;
					}
					if(null != s)
					{
						nSecond = s - 0;
						if(nSecond > 59)
							bValidDate = false;
					}
				}
				if(true == bValidDate && (true == bDate || true == bTime))
				{
					var oDateTmp = new Date();
					oDateTmp.setFullYear(nYear, nMounth, nDay);
					oDateTmp.setHours(nHour, nMinute, nSecond);
					dValue = oDateTmp.getTime() / (86400 * 1000);

					var sFormat;
					if(true == bDate && true == bTime)
					{
						sFormat = sDateFormat + " h:mm:ss";
						if (am || pm)
							sFormat += " AM/PM";
					}
					else if(true == bDate)
						sFormat = sDateFormat;
					else
					{
						if(dValue > 1)
							sFormat = "[h]:mm:ss";
						else if (am || pm)
							sFormat = "h:mm:ss AM/PM";
						else
							sFormat = "h:mm:ss";
					}
					res = {format: sFormat, value: dValue, bDateTime: true, bDate: bDate, bTime: bTime, bPercent: false, bCurrency: false};
				}
            }
        }
		return res;
	},
	isValidDate : function(nYear, nMounth, nDay)
	{
		if(nYear < 1900 && !(1899 === nYear && 11 == nMounth && 31 == nDay))
			return false;
		else
		{
			if(nMounth < 0 || nMounth > 11)
				return false;
			else if(this.isValidDay(nYear, nMounth, nDay))
				return true;
			else if(1900 == nYear && 1 == nMounth && 29 == nDay)
				return true;
		}
		return false;
	},
	isValidDatePDF : function(nYear, nMounth, nDay)
	{
		if(nMounth < 0 || nMounth > 11)
			return false;
		else if(this.isValidDay(nYear, nMounth, nDay))
			return true;
		else if(1900 == nYear && 1 == nMounth && 29 == nDay)
			return true;
		return false;
	},
	isValidDay : function(nYear, nMounth, nDay){
		if(this.isLeapYear(nYear))
		{
			if(nDay <= 0 || nDay > this.daysLeap[nMounth])
				return false;
		}
		else
		{
			if(nDay <= 0 || nDay > this.days[nMounth])
				return false;
		}
		return true;
	},
	isLeapYear : function(year)
	{
		return (0 == (year % 4)) && (0 != (year % 100) || 0 == (year % 400))
	}
};
var g_oFormatParser = new FormatParser();
function escapeRegExp(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}
function setCurrentCultureInfo (LCID, decimalSeparator, groupSeparator) {
	var res = false;
	var cultureInfoNew = g_aCultureInfos[LCID];
	if (cultureInfoNew) {
		if (LCID !== g_oLCID) {
			g_oLCID = LCID;
			AscCommon.g_oDefaultCultureInfo = g_oDefaultCultureInfo = JSON.parse(JSON.stringify(cultureInfoNew)); // ToDo clone
			res = true;
		}
		ParseLocalFormatSymbol(g_oDefaultCultureInfo.Name);
		decimalSeparator = (null != decimalSeparator) ? decimalSeparator : cultureInfoNew.NumberDecimalSeparator;
		if (decimalSeparator !== g_oDefaultCultureInfo.NumberDecimalSeparator) {
			g_oDefaultCultureInfo.NumberDecimalSeparator = decimalSeparator;
			res = true;
		}
		groupSeparator = (null != groupSeparator) ? groupSeparator : cultureInfoNew.NumberGroupSeparator;
		if (groupSeparator !== g_oDefaultCultureInfo.NumberGroupSeparator) {
			g_oDefaultCultureInfo.NumberGroupSeparator = groupSeparator;
			res = true;
		}
	}
	return res;
}
	function checkCultureInfoFontPicker(LCID) {
		var ci = g_aCultureInfos[LCID] || g_oDefaultCultureInfo;
		AscFonts.FontPickerByCharacter.getFontsByString(ci.CurrencySymbol);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.NumberDecimalSeparator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.NumberGroupSeparator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.AMDesignator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.PMDesignator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.DateSeparator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.TimeSeparator);
		var arrays = [ci.DayNames, ci.AbbreviatedDayNames, ci.MonthNames, ci.AbbreviatedMonthNames,
			ci.MonthGenitiveNames, ci.AbbreviatedMonthGenitiveNames
		];
		arrays.forEach(function(arr){
			arr.forEach(function(text) {
				AscFonts.FontPickerByCharacter.getFontsByString(text);
			});
		});
	}

	function isDMY(cultureInfo) {
		//day month year
		var res = true;
		for (var i = 0; i < cultureInfo.ShortDatePattern.length - 1; ++i) {
			if (cultureInfo.ShortDatePattern.charCodeAt(i) > cultureInfo.ShortDatePattern.charCodeAt(i + 1)) {
				return false;
			}
		}
		return true;
	}
	function isYMD(cultureInfo) {
		//year month day
		var res = true;
		for (var i = 0; i < cultureInfo.ShortDatePattern.length - 1; ++i) {
			if (cultureInfo.ShortDatePattern.charCodeAt(i) < cultureInfo.ShortDatePattern.charCodeAt(i + 1)) {
				return false;
			}
		}
		return true;
	}
	function getShortDateMonthFormat(bDate, bYear, opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var separator;
		if ('/' == g_oDefaultCultureInfo.DateSeparator) {
			separator = '-';
		} else {
			separator = '/';
		}
		var sRes = '';
		if (bDate) {
			if (-1 != cultureInfo.ShortDatePattern.indexOf('1')) {
				sRes += 'dd';
			} else {
				sRes += 'd';
			}
			sRes += separator;
		}
		sRes += 'mmm';
		if (bYear) {
			sRes += separator;
			sRes += 'yy';
		}
		return sRes;
	}
	function getShortDateFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var dateElems = [];
		for (var i = 0; i < cultureInfo.ShortDatePattern.length; ++i) {
			switch (cultureInfo.ShortDatePattern[i]) {
				case '0':
					dateElems.push('d');
					break;
				case '1':
					dateElems.push('dd');
					break;
				case '2':
					dateElems.push('m');
					break;
				case '3':
					dateElems.push('mm');
					break;
				case '4':
					dateElems.push('yy');
					break;
				case '5':
					dateElems.push('yyyy');
					break;
			}
		}
		return dateElems.join('/');
	}

	function getShortDateFormat2(day, month, year, opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var dateElems = [];
		for (var i = 0; i < cultureInfo.ShortDatePattern.length; ++i) {
			switch (cultureInfo.ShortDatePattern[i]) {
				case '0':
				case '1':
					if (day > 0) {
						dateElems.push('d'.repeat(day));
					}
					break;
				case '2':
				case '3':
					if (month > 0) {
						dateElems.push('m'.repeat(month));
					}
					break;
				case '4':
				case '5':
					if (year > 0) {
						dateElems.push('y'.repeat(year));
					}
					break;
			}
		}
		return dateElems.join('/');
	}

	function getShortTimeFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		if (AscCommon.is12HourTimeFormat(cultureInfo)) {
			return 'h:mm AM/PM;@';
		} else {
			return 'h:mm;@'
		}
	}
	function getLongTimeFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		if (AscCommon.is12HourTimeFormat(cultureInfo)) {
			return 'h:mm:ss AM/PM;@';
		} else {
			return 'h:mm:ss;@'
		}
	}

	function getNumberFormatSimple(opt_separate, opt_fraction) {
		var numberFormat = opt_separate ? '#,##0' : '0';
		if (opt_fraction > 0) {
			numberFormat += '.' + '0'.repeat(opt_fraction);
		}
		return numberFormat;
	}

	function getNumberFormat(opt_cultureInfo, opt_separate, opt_fraction, opt_red) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(opt_separate, opt_fraction);
		var red = opt_red ? '[Red]' : '';

		var positiveFormat;
		var negativeFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
			case 4:
			case 14:
			case 15:
				positiveFormat = numberFormat + '_)';
				negativeFormat = '\\(' + numberFormat + '\\)';
				break;
			default:
				positiveFormat = numberFormat + '_ ';
				negativeFormat = '\\-' + numberFormat + '\\ ';
				break;
		}
		return positiveFormat + ';' + red + negativeFormat;
	}

	function getLocaleFormat(opt_cultureInfo, opt_currency) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var symbol = opt_currency ? cultureInfo.CurrencySymbol : '';
		return '[$' + symbol + '-' + cultureInfo.LCID.toString(16).toUpperCase() + ']';
	}
	function getCurrencyCustomFormat(symbol) {
		return '[$' + symbol + ']';
	}

	function getCurrencyFormatSimple(opt_cultureInfo, opt_fraction, opt_currency, opt_currencyLocale, opt_currencySymbol, opt_red) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(true, opt_fraction);
		var signCurrencyFormat;
		var signCurrencyFormatEnd;
		var signCurrencyFormatSpace;
		if (opt_currency) {
			if (opt_currencySymbol) {
				signCurrencyFormat = getCurrencyCustomFormat(opt_currencySymbol);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormat = signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			} else {
				if (opt_currencyLocale) {
					signCurrencyFormat = getLocaleFormat(cultureInfo, true);
				} else {
					signCurrencyFormat = '"' + cultureInfo.CurrencySymbol + '"';
				}
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			}
		} else {
			signCurrencyFormatEnd = signCurrencyFormat = signCurrencyFormatSpace = '';
			for (var i = 0; i < cultureInfo.CurrencySymbol.length; ++i) {
				signCurrencyFormatEnd += '_' + cultureInfo.CurrencySymbol[i];
			}
		}
		var red = opt_red ? '[Red]' : '';

		var prefixs = ['_ ', '_-', '_(', '_)'];
		var postfix = '';
		var positiveFormat;
		var negativeFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
				postfix = prefixs[3];
				negativeFormat = '\\(' + signCurrencyFormat + numberFormat + '\\)';
				break;
			case 1:
				negativeFormat = '\\-' + signCurrencyFormat + numberFormat;
				break;
			case 2:
				negativeFormat = signCurrencyFormatSpace + '\\-' + numberFormat;
				break;
			case 3:
				postfix = prefixs[1];
				negativeFormat = signCurrencyFormatSpace + numberFormat + '\\-';
				break;
			case 4:
				postfix = prefixs[3];
				negativeFormat = '\\(' + numberFormat + signCurrencyFormatEnd + '\\)';
				break;
			case 5:
				negativeFormat = '\\-' + numberFormat + signCurrencyFormatEnd;
				break;
			case 6:
				negativeFormat = numberFormat + '\\-' + signCurrencyFormatEnd;
				break;
			case 7:
				postfix = prefixs[1];
				negativeFormat = numberFormat + signCurrencyFormatEnd + '\\-';
				break;
			case 8:
				negativeFormat = '\\-' + numberFormat + '\\ ' + signCurrencyFormatEnd;
				break;
			case 9:
				negativeFormat = '\\-' + signCurrencyFormatSpace + numberFormat;
				break;
			case 10:
				postfix = prefixs[1];
				negativeFormat = numberFormat + '\\ ' + signCurrencyFormatEnd + '\\-';
				break;
			case 11:
				postfix = prefixs[1];
				negativeFormat = signCurrencyFormatSpace + numberFormat + '\\-';
				break;
			case 12:
				negativeFormat = signCurrencyFormatSpace + '\\-' + numberFormat;
				break;
			case 13:
				negativeFormat = numberFormat + '\\-\\ ' + signCurrencyFormatEnd;
				break;
			case 14:
				postfix = prefixs[3];
				negativeFormat = '(' + signCurrencyFormat + numberFormat + '\\)';
				break;
			case 15:
				postfix = prefixs[3];
				negativeFormat = '\\(' + numberFormat + signCurrencyFormatEnd + '\\)';
				break;
		}
		switch (cultureInfo.CurrencyPositivePattern) {
			case 0:
				positiveFormat = signCurrencyFormat + numberFormat;
				break;
			case 1:
				positiveFormat = numberFormat + signCurrencyFormatEnd;
				break;
			case 2:
				positiveFormat = signCurrencyFormatSpace + numberFormat;
				break;
			case 3:
				positiveFormat = numberFormat + '\\ ' + signCurrencyFormatEnd;
				break;
		}
		positiveFormat = positiveFormat + postfix;
		return positiveFormat + ';' + red + negativeFormat;
	}

	function getCurrencyFormatSimple2(opt_cultureInfo, opt_fraction, opt_currency, opt_currencySymbol, opt_negative) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(true, opt_fraction);
		var signCurrencyFormat;
		var signCurrencyFormatEnd;
		var signCurrencyFormatSpace;
		if (opt_currency) {
			if (opt_currencySymbol) {
				signCurrencyFormat = getCurrencyCustomFormat(opt_currencySymbol);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormat = signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			} else {
				signCurrencyFormat = getLocaleFormat(cultureInfo, true);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			}
		} else {
			signCurrencyFormatEnd = signCurrencyFormat = signCurrencyFormatSpace = '';
			for (var i = 0; i < cultureInfo.CurrencySymbol.length; ++i) {
				signCurrencyFormatEnd += '_' + cultureInfo.CurrencySymbol[i];
			}
		}
		var positiveFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
			case 1:
			case 14:
				positiveFormat = signCurrencyFormat + numberFormat;
				break;
			case 2:
			case 3:
			case 9:
			case 10:
			case 11:
			case 12:
				positiveFormat = signCurrencyFormatSpace + numberFormat;
				break;
			case 4:
			case 5:
			case 6:
			case 7:
			case 15:
				positiveFormat = numberFormat + signCurrencyFormatEnd;
				break;
			case 8:
			case 13:
				positiveFormat = numberFormat + '\\ ' + signCurrencyFormatEnd;
				break;
		}
		return opt_negative ? positiveFormat + ';[Red]' + positiveFormat : positiveFormat;
	}

	function getCurrencyFormat(opt_cultureInfo, opt_fraction, opt_currency, opt_currencyLocale, opt_currencySymbol) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(true, opt_fraction);
		var nullSignFormat = '* "-"';
		if (opt_fraction) {
			nullSignFormat += '?'.repeat(opt_fraction);
		}
		var signCurrencyFormat;
		var signCurrencyFormatEnd;
		var signCurrencyFormatSpace;
		if (opt_currency) {
			if (opt_currencySymbol) {
				signCurrencyFormat = getCurrencyCustomFormat(opt_currencySymbol);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormat = signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			} else {
				if (opt_currencyLocale) {
					signCurrencyFormat = getLocaleFormat(cultureInfo, true);
				} else {
					signCurrencyFormat = '"' + cultureInfo.CurrencySymbol + '"';
				}
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			}
		} else {
			signCurrencyFormatEnd = signCurrencyFormat = signCurrencyFormatSpace = '';
			for (var i = 0; i < cultureInfo.CurrencySymbol.length; ++i) {
				signCurrencyFormatEnd += '_' + cultureInfo.CurrencySymbol[i];
			}
		}

		var prefixs = ['_ ', '_-', '_(', '_)'];
		var prefix = prefixs[0];
		var postfix = prefixs[0];
		var positiveNumberFormat = '* ' + numberFormat;
		var positiveFormat;
		var negativeFormat;
		var nullFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + signCurrencyFormat + '* \\(' + numberFormat + '\\)';
				break;
			case 1:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-' + signCurrencyFormat + '* ' + numberFormat + postfix;
				break;
			case 2:
				negativeFormat = prefix + signCurrencyFormatSpace + '* \\-' + numberFormat + postfix;
				break;
			case 3:
				prefix = postfix = prefixs[1];
				negativeFormat = prefix + signCurrencyFormatSpace + '* ' + numberFormat + '\\-';
				break;
			case 4:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + '* \\(' + numberFormat + '\\)' + signCurrencyFormatEnd + postfix;
				break;
			case 5:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-* ' + numberFormat + signCurrencyFormatEnd + postfix;
				break;
			case 6:
				negativeFormat = prefix + '* ' + numberFormat + '\\-' + signCurrencyFormatEnd + postfix;
				break;
			case 7:
				negativeFormat = prefix + '* ' + numberFormat + signCurrencyFormatEnd + '\\-';
				break;
			case 8:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-* ' + numberFormat + '\\ ' + signCurrencyFormatEnd + postfix;
				break;
			case 9:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-' + signCurrencyFormatSpace + '* ' + numberFormat + postfix;
				break;
			case 10:
				negativeFormat = prefix + '* ' + numberFormat + '\\ ' + signCurrencyFormatEnd + '\\-';
				break;
			case 11:
				negativeFormat = prefix + signCurrencyFormatSpace + '* ' + numberFormat + '\\-';
				break;
			case 12:
				negativeFormat = prefix + signCurrencyFormatSpace + '* \\-' + numberFormat + postfix;
				break;
			case 13:
				negativeFormat = prefix + '* ' + numberFormat + '\\-\\ ' + signCurrencyFormatEnd + postfix;
				break;
			case 14:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + signCurrencyFormatSpace + '* \\(' + numberFormat + '\\)';
				break;
			case 15:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + '* \\(' + numberFormat + '\\)\\ ' + signCurrencyFormatEnd + postfix;
				break;
		}
		switch (cultureInfo.CurrencyPositivePattern) {
			case 0:
				positiveFormat = signCurrencyFormat + positiveNumberFormat;
				nullFormat = signCurrencyFormat + nullSignFormat;
				break;
			case 1:
				positiveFormat = positiveNumberFormat + signCurrencyFormatEnd;
				nullFormat = nullSignFormat + signCurrencyFormatEnd;
				break;
			case 2:
				positiveFormat = signCurrencyFormatSpace + positiveNumberFormat;
				nullFormat = signCurrencyFormatSpace + nullSignFormat;
				break;
			case 3:
				positiveFormat = positiveNumberFormat + '\\ ' + signCurrencyFormatEnd;
				nullFormat = nullSignFormat + '\\ ' + signCurrencyFormatEnd;
				break;
		}
		positiveFormat = prefix + positiveFormat + postfix;
		nullFormat = prefix + nullFormat + postfix;
		var textFormat = prefix + '@' + postfix;
		return positiveFormat + ';' + negativeFormat + ';' + nullFormat + ';' + textFormat;
	}

	function getFormatCells(info) {
		var res = [];
		if (info) {
			var format;
			var i;
			var currencySymbol = info.currency;
			var cultureInfo = g_aCultureInfos[info.symbol];
			var hasCurrency = !!cultureInfo || !!currencySymbol;
			if (Asc.c_oAscNumFormatType.General === info.type) {
				res.push(AscCommon.g_cGeneralFormat);
			} else if (Asc.c_oAscNumFormatType.Number === info.type) {
				var numberFormat = getNumberFormatSimple(info.separator, info.decimalPlaces);
				res.push(numberFormat);
				res.push(numberFormat + ';[Red]' + numberFormat);
				res.push(getNumberFormat(cultureInfo, info.separator, info.decimalPlaces, false));
				res.push(getNumberFormat(cultureInfo, info.separator, info.decimalPlaces, true));
			} else if (Asc.c_oAscNumFormatType.Currency === info.type) {
				res.push(getCurrencyFormatSimple2(cultureInfo, info.decimalPlaces, hasCurrency, currencySymbol, false));
				res.push(getCurrencyFormatSimple2(cultureInfo, info.decimalPlaces, hasCurrency, currencySymbol, true));
				res.push(getCurrencyFormatSimple(cultureInfo, info.decimalPlaces, hasCurrency, true, currencySymbol, false));
				res.push(getCurrencyFormatSimple(cultureInfo, info.decimalPlaces, hasCurrency, true, currencySymbol, true));
			} else if (Asc.c_oAscNumFormatType.Accounting === info.type) {
				res.push(getCurrencyFormat(cultureInfo, info.decimalPlaces, hasCurrency, true, currencySymbol));
			} else if (Asc.c_oAscNumFormatType.Date === info.type) {
				//todo locale dependence
				if (info.symbol == g_oDefaultCultureInfo.LCID) {
					res.push(getShortDateFormat(cultureInfo));
					res.push('[$-F800]' + cultureInfo.LongDatePattern);
				}
				res.push(getShortDateFormat2(1, 1, 0, cultureInfo) + ';@');
				res.push(getShortDateFormat2(2, 2, 0, cultureInfo) + ';@');
				res.push(getShortDateFormat2(1, 1, 2, cultureInfo) + ';@');
				res.push(getShortDateFormat2(2, 2, 2, cultureInfo) + ';@');
				res.push(getShortDateFormat2(1, 1, 4, cultureInfo) + ';@');
				res.push(getShortDateFormat2(2, 2, 4, cultureInfo) + ';@');
				res.push(getShortDateFormat2(1, 1, 2, cultureInfo) + ' h:mm;@');
				res.push(getShortDateFormat2(2, 2, 2, cultureInfo) + ' h:mm;@');
				res.push('[$-409]' + getShortDateFormat2(1, 1, 2, cultureInfo) + ' h:mm AM/PM;@');
				var locale = getLocaleFormat(cultureInfo, false);
				res.push(locale + 'mmmmm;@');
				res.push(locale + 'mmmm d, yyyy;@');
				var separators = ['-', '/', ' '];
				for (i = 0; i < separators.length; ++i) {
					var separator = separators[i];
					res.push(locale + 'd' + separator + 'mmm;@');
					res.push(locale + 'd' + separator + 'mmm' + separator + 'yy;@');
					res.push(locale + 'dd' + separator + 'mmm' + separator + 'yy;@');
					res.push(locale + 'mmm' + separator + 'yy;@');
					res.push(locale + 'mmmm' + separator + 'yy;@');
					res.push(locale + 'mmmmm' + separator + 'yy;@');
					res.push(locale + 'yy' + separator + 'mmm;@');
					res.push(locale + 'd' + separator + 'mmm' + separator + 'yyyy;@');
					res.push(locale + 'yyyy' + separator + 'mmm' + separator + 'd;@');
					res.push(locale + 'yy' + separator + 'mmm' + separator + 'd;@');
					res.push('yy' + separator + 'm' + separator + 'd;@');
					res.push('yy' + separator + 'mm' + separator + 'dd;@');
					res.push('yyyy' + separator + 'm' + separator + 'd;@');
					res.push('yyyy' + separator + 'mm' + separator + 'dd;@');
				}
			} else if (Asc.c_oAscNumFormatType.Time === info.type) {
				if (AscCommon.is12HourTimeFormat(cultureInfo)) {
					res = ['[$-F400]h:mm:ss AM/PM', 'h:mm;@', 'h:mm AM/PM;@', 'h:mm:ss;@', 'h:mm:ss AM/PM;@', 'mm:ss.0;@', '[h]:mm:ss;@'];
				} else {
					res = ['[$-F400]h:mm:ss', 'h:mm;@', 'h:mm AM/PM;@', 'h:mm:ss;@', 'h:mm:ss AM/PM;@', 'mm:ss.0;@', '[h]:mm:ss;@'];
				}
			} else if (Asc.c_oAscNumFormatType.Percent === info.type) {
				format = '0';
				if (info.decimalPlaces > 0) {
					format += '.' + '0'.repeat(info.decimalPlaces);
				}
				format += '%';
				res.push(format);
			} else if (Asc.c_oAscNumFormatType.Fraction === info.type) {
				res = gc_aFractionFormats;
			} else if (Asc.c_oAscNumFormatType.Scientific === info.type) {
				format = '0.' + '0'.repeat(info.decimalPlaces) + 'E+00';
				res.push(format);
			} else if (Asc.c_oAscNumFormatType.Text === info.type) {
				res.push('@');
			} else if (Asc.c_oAscNumFormatType.Custom === info.type) {
				for (i = 0; i <= 4; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(getCurrencyFormatSimple(null, 0, false, false, null, false));
				res.push(getCurrencyFormatSimple(null, 0, false, false, null, true));
				res.push(getCurrencyFormatSimple(null, 2, false, false, null, false));
				res.push(getCurrencyFormatSimple(null, 2, false, false, null, true));
				res.push(getCurrencyFormatSimple(null, 0, true, false, null, false));
				res.push(getCurrencyFormatSimple(null, 0, true, false, null, true));
				res.push(getCurrencyFormatSimple(null, 2, true, false, null, false));
				res.push(getCurrencyFormatSimple(null, 2, true, false, null, true));
				for (i = 9; i <= 13; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(getShortDateFormat(null));
				res.push(getShortDateMonthFormat(true, true, null));
				res.push(getShortDateMonthFormat(true, false, null));
				res.push(getShortDateMonthFormat(false, true, null));
				for (i = 18; i <= 21; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(getShortDateFormat(null) + " h:mm");
				for (i = 45; i <= 49; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(AscCommon.getCurrencyFormat(null, 0, true, false, null));
				res.push(AscCommon.getCurrencyFormat(null, 0, false, false, null));
				res.push(AscCommon.getCurrencyFormat(null, 2, true, false, null));
				res.push(AscCommon.getCurrencyFormat(null, 2, false, false, null));
			} else {
				res.push(AscCommon.g_cGeneralFormat);
				res.push('0.00');
				res.push('0.00E+00');
				res.push(getCurrencyFormat(cultureInfo, 2, hasCurrency, true, currencySymbol));
				res.push(getCurrencyFormatSimple2(cultureInfo, 2, hasCurrency, currencySymbol, false));
				res.push(getShortDateFormat(cultureInfo));
				res.push('[$-F800]' + cultureInfo.LongDatePattern);
				//todo F400
				if (AscCommon.is12HourTimeFormat(cultureInfo)) {
					res.push('[$-F400]h:mm:ss AM/PM');
				} else {
					res.push('[$-F400]h:mm:ss');
				}
				res.push('0.00%');
				res.push('# ?/?');
				res.push('@');
			}
		}
		return res;
	}
	function getFormatByCulturalStandardId(id, opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		let formats;
		let localeStart = cultureInfo.Name.substring(0, 2);
		let LCID = cultureInfo.LCID;
		if ('zh' === localeStart) {
			if (4 === LCID || 2052 === LCID || 4100 === LCID || 30724 === LCID) {
				// zh
				// zh-Hans
				// zh-CN
				// zh-SG
				formats = {
					27: 'yyyy"年"m"月"',
					28: 'm"月"d"日"',
					29: 'm"月"d"日"',
					30: 'm-d-yy',
					31: 'yyyy"年"m"月"d"日"',
					32: 'h"时"mm"分"',
					33: 'h"时"mm"分"ss"秒"',
					34: '上午/下午h"时"mm"分"',
					35: '上午/下午h"时"mm"分"ss"秒"',
					36: 'yyyy"年"m"月"',
					50: 'yyyy"年"m"月"',
					51: 'm"月"d"日"',
					52: 'yyyy"年"m"月"',
					53: 'm"月"d"日"',
					54: 'm"月"d"日"',
					55: '上午/下午h"时"mm"分"',
					56: '上午/下午h"时"mm"分"ss"秒"',
					57: 'yyyy"年"m"月"',
					58: 'm"月"d"日"'
				}
			} else {
				// zh-Hant
				// zh-TW
				// zh-HK
				// zh-MO
				formats = {
					27: '[$-404]e/m/d',
					28: '[$-404]e"年"m"月"d"日"',
					29: '[$-404]e"年"m"月"d"日"',
					30: 'm/d/yy',
					31: 'yyyy"年"m"月"d"日"',
					32: 'hh"時"mm"分"',
					33: 'hh"時"mm"分"ss"秒"',
					34: '上午/下午hh"時"mm"分"',
					35: '上午/下午hh"時"mm"分"ss"秒"',
					36: '[$-404]e/m/d',
					50: '[$-404]e/m/d',
					51: '[$-404]e"年"m"月"d"日"',
					52: '上午/下午hh"時"mm"分"',
					53: '上午/下午hh"時"mm"分"ss"秒"',
					54: '上午/下午hh"時"mm"分"',
					55: '上午/下午hh"時"mm"分"ss"秒"',
					56: '[$-404]e/m/d',
					57: '[$-404]e"年"m"月"d"日"',
					58: '[$-404]e"年"m"月"d"日"'
				}
			}
		} else if ('ja' === localeStart) {
			//"ja-jp"
			formats = {
				27: '[$-411]ge.m.d',
				28: '[$-411]ggge"年"m"月"d"日"',
				29: '[$-411]ggge"年"m"月"d"日"',
				30: 'm/d/yy',
				31: 'yyyy"年"m"月"d"日"',
				32: 'h"時"mm"分"',
				33: 'h"時"mm"分"ss"秒"',
				34: 'yyyy"年"m"月"',
				35: 'm"月"d"日"',
				36: '[$-411]ge.m.d',
				50: '[$-411]ge.m.d',
				51: '[$-411]ggge"年"m"月"d"日"',
				52: 'yyyy"年"m"月"',
				53: 'm"月"d"日"',
				54: '[$-411]ggge"年"m"月"d"日"',
				55: 'yyyy"年"m"月"',
				56: 'm"月"d"日"',
				57: '[$-411]ge.m.d',
				58: '[$-411]ggge"年"m"月"d"日"'
			}
		} else if ('ko' === localeStart) {
			//"ko-kr"
			formats = {
				27: 'yyyy"年" mm"月" dd"日"',
				28: 'mm-dd',
				29: 'mm-dd',
				30: 'mm-dd-yy',
				31: 'yyyy"년" mm"월" dd"일"',
				32: 'h"시" mm"분"',
				33: 'h"시" mm"분" ss"초"',
				34: 'yyyy-mm-dd',
				35: 'yyyy-mm-dd',
				36: 'yyyy"年" mm"月" dd"日"',
				50: 'yyyy"年" mm"月" dd"日"',
				51: 'mm-dd',
				52: 'yyyy-mm-dd',
				53: 'yyyy-mm-dd',
				54: 'mm-dd',
				55: 'yyyy-mm-dd',
				56: 'yyyy-mm-dd',
				57: 'yyyy"年" mm"月" dd"日"',
				58: 'mm-dd'
			}
		} else if ('th' === localeStart) {
			//"th-th"
			formats = {
				59: 't0',
					60: 't0.00',
					61: 't#,##0',
					62: 't#,##0.00',
					67: 't0%',
					68: 't0.00%',
					69: 't# ?/?',
					70: 't# ??/??',
					71: 'ว/ด/ปปปป',
					72: 'ว-ดดด-ปป',
					73: 'ว-ดดด',
					74: 'ดดด-ปป',
					75: 'ช:นน',
					76: 'ช:นน:ทท',
					77: 'ว/ด/ปปปป ช:นน',
					78: 'นน:ทท',
					79: '[ช]:นน:ทท',
					80: '80 นน:ทท.0',
					81: 'd/m/bb'
			}
		}
		return formats && formats[id] || null;
	}
	function getFormatByStandardId(id, opt_cultureInfo) {
		var res = getFormatByCulturalStandardId(id, opt_cultureInfo);
		if (res) {
			return res;
		}
		if (59 <= id && id <= 78) {
			if (69 <= id && id <= 71) {
				id += 1;
			}
			id -= 58;
		} else if (79 <= id && id <= 81) {
			id -= 34;
		}
			//todo currencyLocale true/false?
			var currencyLocale = true;
			switch (id) {
				case 5:
					res = AscCommon.getCurrencyFormatSimple(null, 0, true, currencyLocale, null, false);
					break;
				case 6:
					res = AscCommon.getCurrencyFormatSimple(null, 0, true, currencyLocale, null, true);
					break;
				case 7:
					res = AscCommon.getCurrencyFormatSimple(null, 2, true, currencyLocale, null, false);
					break;
				case 8:
					res = AscCommon.getCurrencyFormatSimple(null, 2, true, currencyLocale, null, true);
					break;
				case 14:
					res = AscCommon.getShortDateFormat(null);
					break;
			case 15:
				res = AscCommon.getShortDateMonthFormat(true, true, null);
				break;
			case 16:
				res = AscCommon.getShortDateMonthFormat(true, false, null);
				break;
			case 17:
				res = AscCommon.getShortDateMonthFormat(false, true, null);
				break;
				case 22:
					res = AscCommon.getShortDateFormat(null) + " h:mm";
					break;
			case 23:
			case 24:
			case 25:
			case 26:
				//like 0
				res = "General";
				break;
				case 27:
				case 28:
				case 29:
				case 30:
				case 31:
				//like 14
				res = AscCommon.getShortDateFormat(null);
				break;
			case 32:
			case 33:
			case 34:
			case 35:
				//like 21
				res = AscCommonExcel.aStandartNumFormats[21];
				break;
				case 36:
				//like 14
					res = AscCommon.getShortDateFormat(null);
					break;
				case 37:
					res = AscCommon.getCurrencyFormatSimple(null, 0, false, currencyLocale, null, false);
					break;
				case 38:
					res = AscCommon.getCurrencyFormatSimple(null, 0, false, currencyLocale, null, true);
					break;
				case 39:
					res = AscCommon.getCurrencyFormatSimple(null, 2, false, currencyLocale, null, false);
					break;
				case 40:
					res = AscCommon.getCurrencyFormatSimple(null, 2, false, currencyLocale, null, true);
					break;
				case 41:
					res = AscCommon.getCurrencyFormat(null, 0, false, currencyLocale, null);
					break;
				case 42:
					res = AscCommon.getCurrencyFormat(null, 0, true, currencyLocale, null);
					break;
				case 43:
					res = AscCommon.getCurrencyFormat(null, 2, false, currencyLocale, null);
					break;
				case 44:
					res = AscCommon.getCurrencyFormat(null, 2, true, currencyLocale, null);
					break;
			case 50:
			case 51:
			case 52:
			case 53:
			case 54:
			case 55:
			case 56:
			case 57:
			case 58:
				//like 14
				res = AscCommon.getShortDateFormat(null);
				break;
                default:
                    res = AscCommonExcel.aStandartNumFormats[id];
                    break;
			}
		return res;
	}
	function canGetFormatByStandardId(id) {
		return (5 <= id && id <= 8) || (14 <= id && id <= 17) || 22 == id || (27 <= id && id <= 81);
	}
	function is12HourTimeFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		return cultureInfo.UseAMPM > 0;
	}

	//Excel uses DateSeparator with 2 letters only in date patterns
var g_aCultureInfos = {
	1: {LCID: 1, Name: "ar", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.س.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], AbbreviatedMonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "134", LongDatePattern: "dd/mmmm/yyyy"},
	4: {LCID: 4, Name: "zh-Hans", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	5: {LCID: 5, Name: "cs", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Kč", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["neděle", "pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota"], AbbreviatedDayNames: ["ne", "po", "út", "st", "čt", "pá", "so"], MonthNames: ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec", ""], AbbreviatedMonthNames: ["led", "úno", "bře", "dub", "kvě", "čvn", "čvc", "srp", "zář", "říj", "lis", "pro", ""], MonthGenitiveNames: ["ledna", "února", "března", "dubna", "května", "června", "července", "srpna", "září", "října", "listopadu", "prosince", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "odp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	6: {LCID: 6, Name: "da", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["søndag", "mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag"], AbbreviatedDayNames: ["sø", "ma", "ti", "on", "to", "fr", "lø"], MonthNames: ["januar", "februar", "marts", "april", "maj", "juni", "juli", "august", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy"},
	7: {LCID: 7, Name: "de", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	8: {LCID: 8, Name: "el", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Κυριακή", "Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο"], AbbreviatedDayNames: ["Κυρ", "Δευ", "Τρι", "Τετ", "Πεμ", "Παρ", "Σαβ"], MonthNames: ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος", "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος", ""], AbbreviatedMonthNames: ["Ιαν", "Φεβ", "Μαρ", "Απρ", "Μαϊ", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ", ""], MonthGenitiveNames: ["Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου", "Ιουνίου", "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου", "Νοεμβρίου", "Δεκεμβρίου", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "πμ", PMDesignator: "μμ", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	9: {LCID: 9, Name: "en", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ d\\,\\ yyyy"},
	10: {LCID: 10, Name: "es", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["do.", "lu.", "ma.", "mi.", "ju.", "vi.", "sá."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	11: {LCID: 11, Name: "fi", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sunnuntai", "maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai"], AbbreviatedDayNames: ["su", "ma", "ti", "ke", "to", "pe", "la"], MonthNames: ["tammikuu", "helmikuu", "maaliskuu", "huhtikuu", "toukokuu", "kesäkuu", "heinäkuu", "elokuu", "syyskuu", "lokakuu", "marraskuu", "joulukuu", ""], AbbreviatedMonthNames: ["tammi", "helmi", "maalis", "huhti", "touko", "kesä", "heinä", "elo", "syys", "loka", "marras", "joulu", ""], MonthGenitiveNames: ["tammikuuta", "helmikuuta", "maaliskuuta", "huhtikuuta", "toukokuuta", "kesäkuuta", "heinäkuuta", "elokuuta", "syyskuuta", "lokakuuta", "marraskuuta", "joulukuuta", ""], AbbreviatedMonthGenitiveNames: ["tammik.", "helmik.", "maalisk.", "huhtik.", "toukok.", "kesäk.", "heinäk.", "elok.", "syysk.", "lokak.", "marrask.", "jouluk.", ""], AMDesignator: "ap.", PMDesignator: "ip.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ".", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	12: {LCID: 12, Name: "fr", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	14: {LCID: 14, Name: "hu", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Ft", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat"], AbbreviatedDayNames: ["V", "H", "K", "Sze", "Cs", "P", "Szo"], MonthNames: ["január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "febr.", "márc.", "ápr.", "máj.", "jún.", "júl.", "aug.", "szept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "de.", PMDesignator: "du.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.\\ mmmm\\ d\\.\\,\\ dddd"},
	16: {LCID: 16, Name: "it", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"], AbbreviatedDayNames: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"], MonthNames: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre", ""], AbbreviatedMonthNames: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	17: {LCID: 17, Name: "ja", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"], AbbreviatedDayNames: ["日", "月", "火", "水", "木", "金", "土"], MonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "午前", PMDesignator: "午後", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	18: {LCID: 18, Name: "ko", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₩", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"], AbbreviatedDayNames: ["일", "월", "화", "수", "목", "금", "토"], MonthNames: ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "오전", PMDesignator: "오후", UseAMPM: 1, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"년\"\\ m\"월\"\\ d\"일\"\\ dddd"},
	21: {LCID: 21, Name: "pl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "zł", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"], AbbreviatedDayNames: ["niedz.", "pon.", "wt.", "śr.", "czw.", "pt.", "sob."], MonthNames: ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień", ""], AbbreviatedMonthNames: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru", ""], MonthGenitiveNames: ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września", "października", "listopada", "grudnia", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	22: {LCID: 22, Name: "pt", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "R$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"], AbbreviatedDayNames: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"], MonthNames: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro", ""], AbbreviatedMonthNames: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	25: {LCID: 25, Name: "ru", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₽", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", ""], AbbreviatedMonthNames: ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\ \"г.\""},
	29: {LCID: 29, Name: "sv", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"], AbbreviatedDayNames: ["sön", "mån", "tis", "ons", "tor", "fre", "lör"], MonthNames: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "\"den \"d\\ mmmm\\ yyyy"},
	31: {LCID: 31, Name: "tr", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₺", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"], AbbreviatedDayNames: ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"], MonthNames: ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık", ""], AbbreviatedMonthNames: ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ÖÖ", PMDesignator: "ÖS", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "d\\ mmmm\\ yyyy\\ dddd"},
	33: {LCID: 33, Name: "id", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Rp", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"], AbbreviatedDayNames: ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"], MonthNames: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	34: {LCID: 34, Name: "uk", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₴", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["неділя", "понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота"], AbbreviatedDayNames: ["Нд", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["січень", "лютий", "березень", "квітень", "травень", "червень", "липень", "серпень", "вересень", "жовтень", "листопад", "грудень", ""], AbbreviatedMonthNames: ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру", ""], MonthGenitiveNames: ["січня", "лютого", "березня", "квітня", "травня", "червня", "липня", "серпня", "вересня", "жовтня", "листопада", "грудня", ""], AbbreviatedMonthGenitiveNames: ["січ", "лют", "бер", "кві", "тра", "чер", "лип", "сер", "вер", "жов", "лис", "гру", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\" р.\""},
	36: {LCID: 36, Name: "sl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljek", "torek", "sreda", "četrtek", "petek", "sobota"], AbbreviatedDayNames: ["ned.", "pon.", "tor.", "sre.", "čet.", "pet.", "sob."], MonthNames: ["januar", "februar", "marec", "april", "maj", "junij", "julij", "avgust", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "feb.", "mar.", "apr.", "maj", "jun.", "jul.", "avg.", "sep.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "pop.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy"},
	38: {LCID: 38, Name: "lv", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["svētdiena", "pirmdiena", "otrdiena", "trešdiena", "ceturtdiena", "piektdiena", "sestdiena"], AbbreviatedDayNames: ["svētd.", "pirmd.", "otrd.", "trešd.", "ceturtd.", "piektd.", "sestd."], MonthNames: ["janvāris", "februāris", "marts", "aprīlis", "maijs", "jūnijs", "jūlijs", "augusts", "septembris", "oktobris", "novembris", "decembris", ""], AbbreviatedMonthNames: ["janv.", "febr.", "marts", "apr.", "maijs", "jūn.", "jūl.", "aug.", "sept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priekšp.", PMDesignator: "pēcp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ yyyy\\.\\ \"gada\"\\ d\\.\\ mmmm"},
	39: {LCID: 39, Name: "lt", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sekmadienis", "pirmadienis", "antradienis", "trečiadienis", "ketvirtadienis", "penktadienis", "šeštadienis"], AbbreviatedDayNames: ["sk", "pr", "an", "tr", "kt", "pn", "št"], MonthNames: ["sausis", "vasaris", "kovas", "balandis", "gegužė", "birželis", "liepa", "rugpjūtis", "rugsėjis", "spalis", "lapkritis", "gruodis", ""], AbbreviatedMonthNames: ["saus.", "vas.", "kov.", "bal.", "geg.", "birž.", "liep.", "rugp.", "rugs.", "spal.", "lapkr.", "gruod.", ""], MonthGenitiveNames: ["sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio", "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priešpiet", PMDesignator: "popiet", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\ \"m\"\\.\\ mmmm\\ d\\ \"d\"\\.\\,\\ dddd"},
	42: {LCID: 42, Name: "vi", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₫", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Chủ Nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"], AbbreviatedDayNames: ["CN", "T2", "T3", "T4", "T5", "T6", "T7"], MonthNames: ["Tháng Giêng", "Tháng Hai", "Tháng Ba", "Tháng Tư", "Tháng Năm", "Tháng Sáu", "Tháng Bảy", "Tháng Tám", "Tháng Chín", "Tháng Mười", "Tháng Mười Một", "Tháng Mười Hai", ""], AbbreviatedMonthNames: ["Thg1", "Thg2", "Thg3", "Thg4", "Thg5", "Thg6", "Thg7", "Thg8", "Thg9", "Thg10", "Thg11", "Thg12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "SA", PMDesignator: "CH", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	44: {LCID: 44, Name: "az", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["bazar", "bazar ertəsi", "çərşənbə axşamı", "çərşənbə", "cümə axşamı", "cümə", "şənbə"], AbbreviatedDayNames: ["B.", "B.E.", "Ç.A.", "Ç.", "C.A.", "C.", "Ş."], MonthNames: ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul", "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr", ""], AbbreviatedMonthNames: ["yan", "fev", "mar", "apr", "may", "iyn", "iyl", "avq", "sen", "okt", "noy", "dek", ""], MonthGenitiveNames: ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\,\\ dddd"},
	63: {LCID: 63, Name: "kk", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₸", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["жексенбі", "дүйсенбі", "сейсенбі", "сәрсенбі", "бейсенбі", "жұма", "сенбі"], AbbreviatedDayNames: ["жс", "дс", "сс", "ср", "бс", "жм", "сб"], MonthNames: ["Қаңтар", "Ақпан", "Наурыз", "Сәуір", "Мамыр", "Маусым", "Шілде", "Тамыз", "Қыркүйек", "Қазан", "Қараша", "Желтоқсан", ""], AbbreviatedMonthNames: ["қаң.", "ақп.", "нау.", "сәу.", "мам.", "мау.", "шіл.", "там.", "қыр.", "қаз.", "қар.", "жел.", ""], MonthGenitiveNames: ["қаңтар", "ақпан", "наурыз", "сәуір", "мамыр", "маусым", "шілде", "тамыз", "қыркүйек", "қазан", "қараша", "желтоқсан", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "yyyy\\ \"ж\"\\.\\ d\\ mmmm\\,\\ dddd"},
	80: {LCID: 80, Name: "mn", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["ням", "даваа", "мягмар", "лхагва", "пүрэв", "баасан", "бямба"], AbbreviatedDayNames: ["Ня", "Да", "Мя", "Лх", "Пү", "Ба", "Бя"], MonthNames: ["Нэгдүгээр сар", "Хоёрдугаар сар", "Гуравдугаар сар", "Дөрөвдүгээр сар", "Тавдугаар сар", "Зургаадугаар сар", "Долоодугаар сар", "Наймдугаар сар", "Есдүгээр сар", "Аравдугаар сар", "Арван нэгдүгээр сар", "Арван хоёрдугаар сар", ""], AbbreviatedMonthNames: ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар", "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар", ""], MonthGenitiveNames: ["нэгдүгээр сар", "хоёрдугаар сар", "гуравдугаар сар", "дөрөвдүгээр сар", "тавдугаар сар", "зургаадугаар сар", "долоодугаар сар", "наймдугаар сар", "есдүгээр сар", "аравдугаар сар", "арван нэгдүгээр сар", "арван хоёрдугаар сар", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ү.ө.", PMDesignator: "ү.х.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.mm\\.dd\\,\\ dddd"},
	1025: {LCID: 1025, Name: "ar-SA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.س.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], AbbreviatedMonthNames: ["محرم", "صفر", "ربيع الأول", "ربيع الثاني", "جمادى الأولى", "جمادى الثانية", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "134", LongDatePattern: "dd/mmmm/yyyy"},
	1026: {LCID: 1026, Name: "bg-BG", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "лв.", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["неделя", "понеделник", "вторник", "сряда", "четвъртък", "петък", "събота"], AbbreviatedDayNames: ["нед", "пон", "вт", "ср", "четв", "пет", "съб"], MonthNames: ["януари", "февруари", "март", "април", "май", "юни", "юли", "август", "септември", "октомври", "ноември", "декември", ""], AbbreviatedMonthNames: ["яну", "фев", "мар", "апр", "май", "юни", "юли", "авг", "сеп", "окт", "ное", "дек", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dd\\ mmmm\\ yyyy\\ \"г.\""},
	1028: {LCID: 1028, Name: "zh-TW", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "NT$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	1029: {LCID: 1029, Name: "cs-CZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Kč", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["neděle", "pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota"], AbbreviatedDayNames: ["ne", "po", "út", "st", "čt", "pá", "so"], MonthNames: ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec", ""], AbbreviatedMonthNames: ["led", "úno", "bře", "dub", "kvě", "čvn", "čvc", "srp", "zář", "říj", "lis", "pro", ""], MonthGenitiveNames: ["ledna", "února", "března", "dubna", "května", "června", "července", "srpna", "září", "října", "listopadu", "prosince", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "odp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	1030: {LCID: 1030, Name: "da-DK", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["søndag", "mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag"], AbbreviatedDayNames: ["sø", "ma", "ti", "on", "to", "fr", "lø"], MonthNames: ["januar", "februar", "marts", "april", "maj", "juni", "juli", "august", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy"},
	1031: {LCID: 1031, Name: "de-DE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	1032: {LCID: 1032, Name: "el-GR", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Κυριακή", "Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο"], AbbreviatedDayNames: ["Κυρ", "Δευ", "Τρι", "Τετ", "Πεμ", "Παρ", "Σαβ"], MonthNames: ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος", "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος", ""], AbbreviatedMonthNames: ["Ιαν", "Φεβ", "Μαρ", "Απρ", "Μαϊ", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ", ""], MonthGenitiveNames: ["Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου", "Ιουνίου", "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου", "Νοεμβρίου", "Δεκεμβρίου", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "πμ", PMDesignator: "μμ", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	1033: {LCID: 1033, Name: "en-US", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ d\\,\\ yyyy"},
	1035: {LCID: 1035, Name: "fi-FI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sunnuntai", "maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai"], AbbreviatedDayNames: ["su", "ma", "ti", "ke", "to", "pe", "la"], MonthNames: ["tammikuu", "helmikuu", "maaliskuu", "huhtikuu", "toukokuu", "kesäkuu", "heinäkuu", "elokuu", "syyskuu", "lokakuu", "marraskuu", "joulukuu", ""], AbbreviatedMonthNames: ["tammi", "helmi", "maalis", "huhti", "touko", "kesä", "heinä", "elo", "syys", "loka", "marras", "joulu", ""], MonthGenitiveNames: ["tammikuuta", "helmikuuta", "maaliskuuta", "huhtikuuta", "toukokuuta", "kesäkuuta", "heinäkuuta", "elokuuta", "syyskuuta", "lokakuuta", "marraskuuta", "joulukuuta", ""], AbbreviatedMonthGenitiveNames: ["tammik.", "helmik.", "maalisk.", "huhtik.", "toukok.", "kesäk.", "heinäk.", "elok.", "syysk.", "lokak.", "marrask.", "jouluk.", ""], AMDesignator: "ap.", PMDesignator: "ip.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ".", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	1036: {LCID: 1036, Name: "fr-FR", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	1038: {LCID: 1038, Name: "hu-HU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "Ft", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat"], AbbreviatedDayNames: ["V", "H", "K", "Sze", "Cs", "P", "Szo"], MonthNames: ["január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "febr.", "márc.", "ápr.", "máj.", "jún.", "júl.", "aug.", "szept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "de.", PMDesignator: "du.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.\\ mmmm\\ d\\.\\,\\ dddd"},
	1040: {LCID: 1040, Name: "it-IT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"], AbbreviatedDayNames: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"], MonthNames: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre", ""], AbbreviatedMonthNames: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	1041: {LCID: 1041, Name: "ja-JP", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"], AbbreviatedDayNames: ["日", "月", "火", "水", "木", "金", "土"], MonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "午前", PMDesignator: "午後", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	1042: {LCID: 1042, Name: "ko-KR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₩", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"], AbbreviatedDayNames: ["일", "월", "화", "수", "목", "금", "토"], MonthNames: ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", ""], AbbreviatedMonthNames: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "오전", PMDesignator: "오후", UseAMPM: 1, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\"년\"\\ m\"월\"\\ d\"일\"\\ dddd"},
	1043: {LCID: 1043, Name: "nl-NL", CurrencyPositivePattern: 2, CurrencyNegativePattern: 12, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"], AbbreviatedDayNames: ["zo", "ma", "di", "wo", "do", "vr", "za"], MonthNames: ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	1045: {LCID: 1045, Name: "pl-PL", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "zł", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"], AbbreviatedDayNames: ["niedz.", "pon.", "wt.", "śr.", "czw.", "pt.", "sob."], MonthNames: ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień", ""], AbbreviatedMonthNames: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru", ""], MonthGenitiveNames: ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września", "października", "listopada", "grudnia", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	1046: {LCID: 1046, Name: "pt-BR", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "R$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"], AbbreviatedDayNames: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"], MonthNames: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro", ""], AbbreviatedMonthNames: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	1049: {LCID: 1049, Name: "ru-RU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₽", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", ""], AbbreviatedMonthNames: ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\ \"г.\""},
	1050: {LCID: 1050, Name: "hr-HR", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedjelja", "ponedjeljak", "utorak", "srijeda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sri", "čet", "pet", "sub"], MonthNames: ["siječanj", "veljača", "ožujak", "travanj", "svibanj", "lipanj", "srpanj", "kolovoz", "rujan", "listopad", "studeni", "prosinac", ""], AbbreviatedMonthNames: ["sij", "vlj", "ožu", "tra", "svi", "lip", "srp", "kol", "ruj", "lis", "stu", "pro", ""], MonthGenitiveNames: ["siječnja", "veljače", "ožujka", "travnja", "svibnja", "lipnja", "srpnja", "kolovoza", "rujna", "listopada", "studenog", "prosinca", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "d\\.\\ mmmm\\ yyyy\\."},
	1051: {LCID: 1051, Name: "sk-SK", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["nedeľa", "pondelok", "utorok", "streda", "štvrtok", "piatok", "sobota"], AbbreviatedDayNames: ["ne", "po", "ut", "st", "št", "pi", "so"], MonthNames: ["január", "február", "marec", "apríl", "máj", "jún", "júl", "august", "september", "október", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "máj", "jún", "júl", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: ["januára", "februára", "marca", "apríla", "mája", "júna", "júla", "augusta", "septembra", "októbra", "novembra", "decembra", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ d\\.\\ mmmm\\ yyyy"},
	1053: {LCID: 1053, Name: "sv-SE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "kr", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"], AbbreviatedDayNames: ["sön", "mån", "tis", "ons", "tor", "fre", "lör"], MonthNames: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "\"den \"d\\ mmmm\\ yyyy"},
	1055: {LCID: 1055, Name: "tr-TR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₺", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"], AbbreviatedDayNames: ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"], MonthNames: ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık", ""], AbbreviatedMonthNames: ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ÖÖ", PMDesignator: "ÖS", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "d\\ mmmm\\ yyyy\\ dddd"},
	1057: {LCID: 1057, Name: "id-ID", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Rp", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"], AbbreviatedDayNames: ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"], MonthNames: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	1058: {LCID: 1058, Name: "uk-UA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₴", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["неділя", "понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота"], AbbreviatedDayNames: ["Нд", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"], MonthNames: ["січень", "лютий", "березень", "квітень", "травень", "червень", "липень", "серпень", "вересень", "жовтень", "листопад", "грудень", ""], AbbreviatedMonthNames: ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру", ""], MonthGenitiveNames: ["січня", "лютого", "березня", "квітня", "травня", "червня", "липня", "серпня", "вересня", "жовтня", "листопада", "грудня", ""], AbbreviatedMonthGenitiveNames: ["січ", "лют", "бер", "кві", "тра", "чер", "лип", "сер", "вер", "жов", "лис", "гру", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\" р.\""},
	1060: {LCID: 1060, Name: "sl-SI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljek", "torek", "sreda", "četrtek", "petek", "sobota"], AbbreviatedDayNames: ["ned.", "pon.", "tor.", "sre.", "čet.", "pet.", "sob."], MonthNames: ["januar", "februar", "marec", "april", "maj", "junij", "julij", "avgust", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "feb.", "mar.", "apr.", "maj", "jun.", "jul.", "avg.", "sep.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "dop.", PMDesignator: "pop.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy"},
	1062: {LCID: 1062, Name: "lv-LV", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["svētdiena", "pirmdiena", "otrdiena", "trešdiena", "ceturtdiena", "piektdiena", "sestdiena"], AbbreviatedDayNames: ["svētd.", "pirmd.", "otrd.", "trešd.", "ceturtd.", "piektd.", "sestd."], MonthNames: ["janvāris", "februāris", "marts", "aprīlis", "maijs", "jūnijs", "jūlijs", "augusts", "septembris", "oktobris", "novembris", "decembris", ""], AbbreviatedMonthNames: ["janv.", "febr.", "marts", "apr.", "maijs", "jūn.", "jūl.", "aug.", "sept.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priekšp.", PMDesignator: "pēcp.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ yyyy\\.\\ \"gada\"\\ d\\.\\ mmmm"},
	1063: {LCID: 1063, Name: "lt-LT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["sekmadienis", "pirmadienis", "antradienis", "trečiadienis", "ketvirtadienis", "penktadienis", "šeštadienis"], AbbreviatedDayNames: ["sk", "pr", "an", "tr", "kt", "pn", "št"], MonthNames: ["sausis", "vasaris", "kovas", "balandis", "gegužė", "birželis", "liepa", "rugpjūtis", "rugsėjis", "spalis", "lapkritis", "gruodis", ""], AbbreviatedMonthNames: ["saus.", "vas.", "kov.", "bal.", "geg.", "birž.", "liep.", "rugp.", "rugs.", "spal.", "lapkr.", "gruod.", ""], MonthGenitiveNames: ["sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio", "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "priešpiet", PMDesignator: "popiet", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\ \"m\"\\.\\ mmmm\\ d\\ \"d\"\\.\\,\\ dddd"},
	1066: {LCID: 1066, Name: "vi-VN", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₫", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Chủ Nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"], AbbreviatedDayNames: ["CN", "T2", "T3", "T4", "T5", "T6", "T7"], MonthNames: ["Tháng Giêng", "Tháng Hai", "Tháng Ba", "Tháng Tư", "Tháng Năm", "Tháng Sáu", "Tháng Bảy", "Tháng Tám", "Tháng Chín", "Tháng Mười", "Tháng Mười Một", "Tháng Mười Hai", ""], AbbreviatedMonthNames: ["Thg1", "Thg2", "Thg3", "Thg4", "Thg5", "Thg6", "Thg7", "Thg8", "Thg9", "Thg10", "Thg11", "Thg12", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "SA", PMDesignator: "CH", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	1068: {LCID: 1068, Name: "az-Latn-AZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["bazar", "bazar ertəsi", "çərşənbə axşamı", "çərşənbə", "cümə axşamı", "cümə", "şənbə"], AbbreviatedDayNames: ["B.", "B.E.", "Ç.A.", "Ç.", "C.A.", "C.", "Ş."], MonthNames: ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul", "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr", ""], AbbreviatedMonthNames: ["yan", "fev", "mar", "apr", "may", "iyn", "iyl", "avq", "sen", "okt", "noy", "dek", ""], MonthGenitiveNames: ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\,\\ dddd"},
	1087: {LCID: 1087, Name: "kk-KZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₸", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["жексенбі", "дүйсенбі", "сейсенбі", "сәрсенбі", "бейсенбі", "жұма", "сенбі"], AbbreviatedDayNames: ["жс", "дс", "сс", "ср", "бс", "жм", "сб"], MonthNames: ["Қаңтар", "Ақпан", "Наурыз", "Сәуір", "Мамыр", "Маусым", "Шілде", "Тамыз", "Қыркүйек", "Қазан", "Қараша", "Желтоқсан", ""], AbbreviatedMonthNames: ["қаң.", "ақп.", "нау.", "сәу.", "мам.", "мау.", "шіл.", "там.", "қыр.", "қаз.", "қар.", "жел.", ""], MonthGenitiveNames: ["қаңтар", "ақпан", "наурыз", "сәуір", "мамыр", "маусым", "шілде", "тамыз", "қыркүйек", "қазан", "қараша", "желтоқсан", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "yyyy\\ \"ж\"\\.\\ d\\ mmmm\\,\\ dddd"},
	1104: {LCID: 1104, Name: "mn-MN", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["ням", "даваа", "мягмар", "лхагва", "пүрэв", "баасан", "бямба"], AbbreviatedDayNames: ["Ня", "Да", "Мя", "Лх", "Пү", "Ба", "Бя"], MonthNames: ["Нэгдүгээр сар", "Хоёрдугаар сар", "Гуравдугаар сар", "Дөрөвдүгээр сар", "Тавдугаар сар", "Зургаадугаар сар", "Долоодугаар сар", "Наймдугаар сар", "Есдүгээр сар", "Аравдугаар сар", "Арван нэгдүгээр сар", "Арван хоёрдугаар сар", ""], AbbreviatedMonthNames: ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар", "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар", ""], MonthGenitiveNames: ["нэгдүгээр сар", "хоёрдугаар сар", "гуравдугаар сар", "дөрөвдүгээр сар", "тавдугаар сар", "зургаадугаар сар", "долоодугаар сар", "наймдугаар сар", "есдүгээр сар", "аравдугаар сар", "арван нэгдүгээр сар", "арван хоёрдугаар сар", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ү.ө.", PMDesignator: "ү.х.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.mm\\.dd\\,\\ dddd"},
	2049: {LCID: 2049, Name: "ar-IQ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ع.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	2052: {LCID: 2052, Name: "zh-CN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	2055: {LCID: 2055, Name: "de-CH", CurrencyPositivePattern: 2, CurrencyNegativePattern: 2, CurrencySymbol: "CHF", NumberDecimalSeparator: ".", NumberGroupSeparator: "’", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	2057: {LCID: 2057, Name: "en-GB", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "£", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	2058: {LCID: 2058, Name: "es-MX", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	2060: {LCID: 2060, Name: "fr-BE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "134", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	2064: {LCID: 2064, Name: "it-CH", CurrencyPositivePattern: 2, CurrencyNegativePattern: 2, CurrencySymbol: "CHF", NumberDecimalSeparator: ".", NumberGroupSeparator: "’", NumberGroupSizes: [3], DayNames: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"], AbbreviatedDayNames: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"], MonthNames: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre", ""], AbbreviatedMonthNames: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	2070: {LCID: 2070, Name: "pt-PT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"], AbbreviatedDayNames: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"], MonthNames: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro", ""], AbbreviatedMonthNames: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\" de \"mmmm\" de \"yyyy"},
	2073: {LCID: 2073, Name: "ru-MD", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "L", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"], AbbreviatedDayNames: ["вс", "пн", "вт", "ср", "чт", "пт", "сб"], MonthNames: ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь", ""], AbbreviatedMonthNames: ["янв.", "февр.", "март", "апр.", "май", "июнь", "июль", "авг.", "сент.", "окт.", "нояб.", "дек.", ""], MonthGenitiveNames: ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря", ""], AbbreviatedMonthGenitiveNames: ["янв.", "февр.", "мар.", "апр.", "мая", "июн.", "июл.", "авг.", "сент.", "окт.", "нояб.", "дек.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy\\ \"г\"\\."},
	2077: {LCID: 2077, Name: "sv-FI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"], AbbreviatedDayNames: ["sön", "mån", "tis", "ons", "tors", "fre", "lör"], MonthNames: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december", ""], AbbreviatedMonthNames: ["jan.", "feb.", "mars", "apr.", "maj", "juni", "juli", "aug.", "sep.", "okt.", "nov.", "dec.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "fm", PMDesignator: "em", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	2092: {LCID: 2092, Name: "az-Cyrl-AZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["базар", "базар ертәси", "чәршәнбә ахшамы", "чәршәнбә", "ҹүмә ахшамы", "ҹүмә", "шәнбә"], AbbreviatedDayNames: ["Б", "Бе", "Ча", "Ч", "Ҹа", "Ҹ", "Ш"], MonthNames: ["jанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthNames: ["Јан", "Фев", "Мар", "Апр", "Мај", "Ијун", "Ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], MonthGenitiveNames: ["јанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthGenitiveNames: ["Јан", "Фев", "Мар", "Апр", "мая", "ијун", "ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy"},
	2128: {LCID: 2128, Name: "mn-Mong-CN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 0], DayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], AbbreviatedDayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], MonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], AbbreviatedMonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd"},
	3073: {LCID: 3073, Name: "ar-EG", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ج.م.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	3076: {LCID: 3076, Name: "zh-HK", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "HK$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	3079: {LCID: 3079, Name: "de-AT", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Jänner", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jän", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jän.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sep.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	3081: {LCID: 3081, Name: "en-AU", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	3082: {LCID: 3082, Name: "es-ES", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["do.", "lu.", "ma.", "mi.", "ju.", "vi.", "sá."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\" de \"mmmm\" de \"yyyy"},
	3084: {LCID: 3084, Name: "fr-CA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 15, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "d\\ mmmm\\ yyyy"},
	3152: {LCID: 3152, Name: "mn-Mong-MN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 0], DayNames: ["ᠨᠢᠮ᠎ᠠ", "ᠳᠠᠸᠠ", "ᠮᠢᠭᠮᠠᠷ", "ᡀᠠᠭᠪᠠ", "ᠫᠦᠷᠪᠦ", "ᠪᠠᠰᠠᠩ", "ᠪᠢᠮᠪᠠ"], AbbreviatedDayNames: ["ᠨᠢᠮ᠎ᠠ", "ᠳᠠᠸᠠ", "ᠮᠢᠭᠮᠠᠷ", "ᡀᠠᠭᠪᠠ", "ᠫᠦᠷᠪᠦ", "ᠪᠠᠰᠠᠩ", "ᠪᠢᠮᠪᠠ"], MonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], AbbreviatedMonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd"},
	4097: {LCID: 4097, Name: "ar-LY", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ل.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	4100: {LCID: 4100, Name: "zh-SG", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	4103: {LCID: 4103, Name: "de-LU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	4105: {LCID: 4105, Name: "en-CA", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "mmmm\\ d\\,\\ yyyy"},
	4106: {LCID: 4106, Name: "es-GT", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Q", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	4108: {LCID: 4108, Name: "fr-CH", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CHF", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	4122: {LCID: 4122, Name: "hr-BA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "KM", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedjelja", "ponedjeljak", "utorak", "srijeda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sri", "čet", "pet", "sub"], MonthNames: ["siječanj", "veljača", "ožujak", "travanj", "svibanj", "lipanj", "srpanj", "kolovoz", "rujan", "listopad", "studeni", "prosinac", ""], AbbreviatedMonthNames: ["sij", "velj", "ožu", "tra", "svi", "lip", "srp", "kol", "ruj", "lis", "stu", "pro", ""], MonthGenitiveNames: ["siječnja", "veljače", "ožujka", "travnja", "svibnja", "lipnja", "srpnja", "kolovoza", "rujna", "listopada", "studenoga", "prosinca", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy\\."},
	5121: {LCID: 5121, Name: "ar-DZ", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ج.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	5124: {LCID: 5124, Name: "zh-MO", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "MOP", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	5127: {LCID: 5127, Name: "de-LI", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "CHF", NumberDecimalSeparator: ".", NumberGroupSeparator: "’", NumberGroupSizes: [3], DayNames: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], AbbreviatedDayNames: ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."], MonthNames: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\.\\ mmmm\\ yyyy"},
	5129: {LCID: 5129, Name: "en-NZ", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	5130: {LCID: 5130, Name: "es-CR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₡", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	5132: {LCID: 5132, Name: "fr-LU", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	6153: {LCID: 6153, Name: "en-IE", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "€", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	6154: {LCID: 6154, Name: "es-PA", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "B/.", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "315", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	6156: {LCID: 6156, Name: "fr-MC", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	7169: {LCID: 7169, Name: "ar-TN", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ت.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["جانفييه", "فيفرييه", "مارس", "أفريل", "مي", "جوان", "جوييه", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	7177: {LCID: 7177, Name: "en-ZA", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "R", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	7178: {LCID: 7178, Name: "es-DO", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	7180: {LCID: 7180, Name: "fr-029", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "EC$", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre", ""], AbbreviatedMonthNames: ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc.", ""], MonthGenitiveNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthGenitiveNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	8193: {LCID: 8193, Name: "ar-OM", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.ع.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	8201: {LCID: 8201, Name: "en-JM", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	8202: {LCID: 8202, Name: "es-VE", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "Bs.S", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	8204: {LCID: 8204, Name: "fr-RE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "€", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	9217: {LCID: 9217, Name: "ar-YE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.ي.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	9225: {LCID: 9225, Name: "en-029", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "EC$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	9226: {LCID: 9226, Name: "es-CO", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	9228: {LCID: 9228, Name: "fr-CD", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "FC", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	9242: {LCID: 9242, Name: "sr-Latn-RS", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "RSD", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljak", "utorak", "sreda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sre", "čet", "pet", "sub"], MonthNames: ["januar", "februar", "mart", "april", "maj", "jun", "jul", "avgust", "septembar", "oktobar", "novembar", "decembar", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "avg", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "pre podne", PMDesignator: "po podne", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy\\."},
	10241: {LCID: 10241, Name: "ar-SY", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ل.س.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	10249: {LCID: 10249, Name: "en-BZ", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	10250: {LCID: 10250, Name: "es-PE", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "S/", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", ""], AbbreviatedMonthNames: ["Ene.", "Feb.", "Mar.", "Abr.", "May.", "Jun.", "Jul.", "Ago.", "Set.", "Oct.", "Nov.", "Dic.", ""], MonthGenitiveNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "set.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "035", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	10252: {LCID: 10252, Name: "fr-SN", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	10266: {LCID: 10266, Name: "sr-Cyrl-RS", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "дин.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["недеља", "понедељак", "уторак", "среда", "четвртак", "петак", "субота"], AbbreviatedDayNames: ["нед.", "пон.", "ут.", "ср.", "чет.", "пет.", "суб."], MonthNames: ["јануар", "фебруар", "март", "април", "мај", "јун", "јул", "август", "септембар", "октобар", "новембар", "децембар", ""], AbbreviatedMonthNames: ["јан.", "феб.", "март", "апр.", "мај", "јун", "јул", "авг.", "септ.", "окт.", "нов.", "дец.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy\\."},
	11265: {LCID: 11265, Name: "ar-JO", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ا.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	11273: {LCID: 11273, Name: "en-TT", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	11274: {LCID: 11274, Name: "es-AR", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	11276: {LCID: 11276, Name: "fr-CM", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "FCFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "mat.", PMDesignator: "soir", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	12289: {LCID: 12289, Name: "ar-LB", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ل.ل.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], AbbreviatedMonthNames: ["كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران", "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	12297: {LCID: 12297, Name: "en-ZW", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "US$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ dd\\ mmmm\\ yyyy"},
	12298: {LCID: 12298, Name: "es-EC", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	12300: {LCID: 12300, Name: "fr-CI", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	13313: {LCID: 13313, Name: "ar-KW", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ك.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	13321: {LCID: 13321, Name: "en-PH", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "₱", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	13322: {LCID: 13322, Name: "es-CL", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	13324: {LCID: 13324, Name: "fr-ML", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "CFA", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	14337: {LCID: 14337, Name: "ar-AE", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.إ.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	14345: {LCID: 14345, Name: "en-ID", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Rp", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	14346: {LCID: 14346, Name: "es-UY", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "$", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", ""], AbbreviatedMonthNames: ["Ene.", "Feb.", "Mar.", "Abr.", "May.", "Jun.", "Jul.", "Ago.", "Set.", "Oct.", "Nov.", "Dic.", ""], MonthGenitiveNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthGenitiveNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "set.", "oct.", "nov.", "dic.", ""], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	14348: {LCID: 14348, Name: "fr-MA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "DH", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["jan.", "fév.", "mar.", "avr.", "mai", "jui.", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	15361: {LCID: 15361, Name: "ar-BH", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "د.ب.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "ابريل", "مايو", "يونيو", "يوليو", "اغسطس", "سبتمبر", "اكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	15369: {LCID: 15369, Name: "en-HK", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	15370: {LCID: 15370, Name: "es-PY", CurrencyPositivePattern: 2, CurrencyNegativePattern: 12, CurrencySymbol: "₲", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	15372: {LCID: 15372, Name: "fr-HT", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "G", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], AbbreviatedDayNames: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."], MonthNames: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", ""], AbbreviatedMonthNames: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dddd\\ d\\ mmmm\\ yyyy"},
	16385: {LCID: 16385, Name: "ar-QA", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "ر.ق.‏", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], AbbreviatedDayNames: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"], MonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], AbbreviatedMonthNames: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ص", PMDesignator: "م", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\,\\ yyyy"},
	16393: {LCID: 16393, Name: "en-IN", CurrencyPositivePattern: 2, CurrencyNegativePattern: 12, CurrencySymbol: "₹", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 2], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: "-", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "dd\\ mmmm\\ yyyy"},
	16394: {LCID: 16394, Name: "es-BO", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "Bs", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	17417: {LCID: 17417, Name: "en-MY", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "RM", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\,\\ yyyy"},
	17418: {LCID: 17418, Name: "es-SV", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	18441: {LCID: 18441, Name: "en-SG", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], AbbreviatedDayNames: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], MonthNames: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", ""], AbbreviatedMonthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "am", PMDesignator: "pm", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ mmmm\\ yyyy"},
	18442: {LCID: 18442, Name: "es-HN", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "L", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\ dd\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	19466: {LCID: 19466, Name: "es-NI", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "C$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	20490: {LCID: 20490, Name: "es-PR", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a. m.", PMDesignator: "p. m.", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "315", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	21514: {LCID: 21514, Name: "es-US", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 1, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "205", LongDatePattern: "dddd\\,\\ mmmm\\ dd\\,\\ yyyy"},
	22538: {LCID: 22538, Name: "es-419", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "XDR", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a.m.", PMDesignator: "p.m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	23562: {LCID: 23562, Name: "es-CU", CurrencyPositivePattern: 0, CurrencyNegativePattern: 1, CurrencySymbol: "$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], AbbreviatedDayNames: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], MonthNames: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre", ""], AbbreviatedMonthNames: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "a.m.", PMDesignator: "p.m.", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ d\\ \"de\"\\ mmmm\\ \"de\"\\ yyyy"},
	27674: {LCID: 27674, Name: "sr-Cyrl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "дин.", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["недеља", "понедељак", "уторак", "среда", "четвртак", "петак", "субота"], AbbreviatedDayNames: ["нед.", "пон.", "ут.", "ср.", "чет.", "пет.", "суб."], MonthNames: ["јануар", "фебруар", "март", "април", "мај", "јун", "јул", "август", "септембар", "октобар", "новембар", "децембар", ""], AbbreviatedMonthNames: ["јан.", "феб.", "март", "апр.", "мај", "јун", "јул", "авг.", "септ.", "окт.", "нов.", "дец.", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\.\\ mmmm\\ yyyy\\."},
	28698: {LCID: 28698, Name: "sr-Latn", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "RSD", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["nedelja", "ponedeljak", "utorak", "sreda", "četvrtak", "petak", "subota"], AbbreviatedDayNames: ["ned", "pon", "uto", "sre", "čet", "pet", "sub"], MonthNames: ["januar", "februar", "mart", "april", "maj", "jun", "jul", "avgust", "septembar", "oktobar", "novembar", "decembar", ""], AbbreviatedMonthNames: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "avg", "sep", "okt", "nov", "dec", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "pre podne", PMDesignator: "po podne", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "dddd\\,\\ dd\\.\\ mmmm\\ yyyy\\."},
	29740: {LCID: 29740, Name: "az-Cyrl", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: " ", NumberGroupSizes: [3], DayNames: ["базар", "базар ертәси", "чәршәнбә ахшамы", "чәршәнбә", "ҹүмә ахшамы", "ҹүмә", "шәнбә"], AbbreviatedDayNames: ["Б", "Бе", "Ча", "Ч", "Ҹа", "Ҹ", "Ш"], MonthNames: ["jанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthNames: ["Јан", "Фев", "Мар", "Апр", "Мај", "Ијун", "Ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], MonthGenitiveNames: ["јанвар", "феврал", "март", "апрел", "мај", "ијун", "ијул", "август", "сентјабр", "октјабр", "нојабр", "декабр", ""], AbbreviatedMonthGenitiveNames: ["Јан", "Фев", "Мар", "Апр", "мая", "ијун", "ијул", "Авг", "Сен", "Окт", "Ноя", "Дек", ""], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy"},
	30724: {LCID: 30724, Name: "zh", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	30764: {LCID: 30764, Name: "az-Latn", CurrencyPositivePattern: 3, CurrencyNegativePattern: 8, CurrencySymbol: "₼", NumberDecimalSeparator: ",", NumberGroupSeparator: ".", NumberGroupSizes: [3], DayNames: ["bazar", "bazar ertəsi", "çərşənbə axşamı", "çərşənbə", "cümə axşamı", "cümə", "şənbə"], AbbreviatedDayNames: ["B.", "B.E.", "Ç.A.", "Ç.", "C.A.", "C.", "Ş."], MonthNames: ["Yanvar", "Fevral", "Mart", "Aprel", "May", "İyun", "İyul", "Avqust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr", ""], AbbreviatedMonthNames: ["yan", "fev", "mar", "apr", "may", "iyn", "iyl", "avq", "sen", "okt", "noy", "dek", ""], MonthGenitiveNames: ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "AM", PMDesignator: "PM", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "135", LongDatePattern: "d\\ mmmm\\ yyyy\\,\\ dddd"},
	30800: {LCID: 30800, Name: "mn-Cyrl", CurrencyPositivePattern: 2, CurrencyNegativePattern: 9, CurrencySymbol: "₮", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["ням", "даваа", "мягмар", "лхагва", "пүрэв", "баасан", "бямба"], AbbreviatedDayNames: ["Ня", "Да", "Мя", "Лх", "Пү", "Ба", "Бя"], MonthNames: ["Нэгдүгээр сар", "Хоёрдугаар сар", "Гуравдугаар сар", "Дөрөвдүгээр сар", "Тавдугаар сар", "Зургаадугаар сар", "Долоодугаар сар", "Наймдугаар сар", "Есдүгээр сар", "Аравдугаар сар", "Арван нэгдүгээр сар", "Арван хоёрдугаар сар", ""], AbbreviatedMonthNames: ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар", "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар", ""], MonthGenitiveNames: ["нэгдүгээр сар", "хоёрдугаар сар", "гуравдугаар сар", "дөрөвдүгээр сар", "тавдугаар сар", "зургаадугаар сар", "долоодугаар сар", "наймдугаар сар", "есдүгээр сар", "аравдугаар сар", "арван нэгдүгээр сар", "арван хоёрдугаар сар", ""], AbbreviatedMonthGenitiveNames: [], AMDesignator: "ү.ө.", PMDesignator: "ү.х.", UseAMPM: 0, DateSeparator: ".", TimeSeparator: ":", ShortDatePattern: "531", LongDatePattern: "yyyy\\.mm\\.dd\\,\\ dddd"},
	31748: {LCID: 31748, Name: "zh-Hant", CurrencyPositivePattern: 0, CurrencyNegativePattern: 0, CurrencySymbol: "HK$", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3], DayNames: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"], AbbreviatedDayNames: ["週日", "週一", "週二", "週三", "週四", "週五", "週六"], MonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], AbbreviatedMonthNames: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "上午", PMDesignator: "下午", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "025", LongDatePattern: "yyyy\"年\"m\"月\"d\"日\""},
	31824: {LCID: 31824, Name: "mn-Mong", CurrencyPositivePattern: 0, CurrencyNegativePattern: 2, CurrencySymbol: "¥", NumberDecimalSeparator: ".", NumberGroupSeparator: ",", NumberGroupSizes: [3, 0], DayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], AbbreviatedDayNames: ["ᠭᠠᠷᠠᠭ ᠤᠨ ᠡᠳᠦᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠨᠢᠭᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠬᠣᠶᠠᠷ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠭᠤᠷᠪᠠᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠳᠥᠷᠪᠡᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠲᠠᠪᠤᠨ", "ᠭᠠᠷᠠᠭ ᠤᠨ ᠵᠢᠷᠭᠤᠭᠠᠨ"], MonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], AbbreviatedMonthNames: ["ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠭᠤᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠦᠷᠪᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠠᠪᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠵᠢᠷᠭᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠲᠤᠯᠤᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠨᠠᠢᠮᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠶᠢᠰᠦᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠨᠢᠭᠡᠳᠦᠭᠡᠷ ᠰᠠᠷ᠎ᠠ", "ᠠᠷᠪᠠᠨ ᠬᠤᠶ᠋ᠠᠳᠤᠭᠠᠷ ᠰᠠᠷ᠎ᠠ", ""], MonthGenitiveNames: [], AbbreviatedMonthGenitiveNames: [], AMDesignator: "", PMDesignator: "", UseAMPM: 0, DateSeparator: "/", TimeSeparator: ":", ShortDatePattern: "520", LongDatePattern: "yyyy\\ᠣ\\ᠨ\\ mmmm\\ d\\ᠡ\\ᠳ\\ᠦ\\ᠷ\\᠂\\ dddd"},
};
var g_oDefaultCultureInfo, g_oLCID;
setCurrentCultureInfo(1033);//en-US//1033//fr-FR//1036//basq//1069//ru-Ru//1049//hindi//1081
	var g_aAdditionalCurrencySymbols = ["ADP", "AED", "AFA", "AFN", "ALL", "AMD", "ANG", "AOA", "ARS", "ATS", "AUD",
		"AWG", "AZM", "AZN", "BAM", "BBD", "BDT", "BEF", "BGL", "BGN", "BHD", "BIF", "BMD", "BND", "BOB", "BOV", "BRL",
		"BSD", "BTN", "BWP", "BYB", "BYN", "BYR", "BZD", "CAD", "CDF", "CHE", "CHF", "CHW", "CLF", "CLP", "CNY", "COP",
		"COU", "CRC", "CSD", "CUC", "CUP", "CVE", "CYP", "CZK", "DEM", "DJF", "DKK", "DOP", "DZD", "ECS", "ECV", "EEK",
		"EGP", "ERN", "ESP", "ETB", "EUR", "FIM", "FJD", "FKP", "FRF", "GBP", "GEL", "GHC", "GHS", "GIP", "GMD", "GNF",
		"GRD", "GTQ", "GYD", "HKD", "HNL", "HRK", "HTG", "HUF", "IDR", "IEP", "ILS", "INR", "IQD", "IRR", "ISK", "ITL",
		"JMD", "JOD", "JPY", "KAF", "KES", "KGS", "KHR", "KMF", "KPW", "KRW", "KWD", "KYD", "KZT", "LAK", "LBP", "LKR",
		"LRD", "LSL", "LTL", "LUF", "LVL", "LYD", "MAD", "MDL", "MGA", "MGF", "MKD", "MMK", "MNT", "MOP", "MRO", "MRU",
		"MTL", "MUR", "MVR", "MWK", "MXN", "MXV", "MYR", "MZM", "MZN", "NAD", "NGN", "NIO", "NLG", "NOK", "NPR", "NTD",
		"NZD", "OMR", "PAB", "PEN", "PGK", "PHP", "PKR", "PLN", "PTE", "PYG", "QAR", "ROL", "RON", "RSD", "RUB", "RUR",
		"RWF", "SAR", "SBD", "SCR", "SDD", "SDG", "SDP", "SEK", "SGD", "SHP", "SIT", "SKK", "SLL", "SOS", "SPL", "SRD",
		"SRG", "STD", "SVC", "SYP", "SZL", "THB", "TJR", "TJS", "TMM", "TMT", "TND", "TOP", "TRL", "TRY", "TTD", "TWD",
		"TZS", "UAH", "UGX", "USD", "USN", "USS", "UYI", "UYU", "UZS", "VEB", "VEF", "VES", "VND", "VUV", "WST", "XAF",
		"XAG", "XAU", "XB5", "XBA", "XBB", "XBC", "XBD", "XCD", "XDR", "XFO", "XFU", "XOF", "XPD", "XPF", "XPT", "XTS",
		"XXX", "YER", "YUM", "ZAR", "ZMK", "ZMW", "ZWD", "ZWL", "ZWN", "ZWR"
	];

    //---------------------------------------------------------export---------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].isNumber = isNumber;
    window["AscCommon"].NumFormat = NumFormat;
    window["AscCommon"].CellFormat = CellFormat;
    window["AscCommon"].DecodeGeneralFormat = DecodeGeneralFormat;
    window["AscCommon"].setCurrentCultureInfo = setCurrentCultureInfo;
	window["AscCommon"].checkCultureInfoFontPicker = checkCultureInfoFontPicker;
	window['AscCommon'].getShortDateFormat = getShortDateFormat;
	window['AscCommon'].getShortDateFormat2 = getShortDateFormat2;
	window['AscCommon'].getShortTimeFormat = getShortTimeFormat;
	window['AscCommon'].getLongTimeFormat = getLongTimeFormat;
	window['AscCommon'].getShortDateMonthFormat = getShortDateMonthFormat;
	window['AscCommon'].getNumberFormatSimple = getNumberFormatSimple;
	window['AscCommon'].getNumberFormat = getNumberFormat;
	window['AscCommon'].getLocaleFormat = getLocaleFormat;
	window['AscCommon'].getCurrencyFormatSimple = getCurrencyFormatSimple;
	window['AscCommon'].getCurrencyFormatSimple2 = getCurrencyFormatSimple2;
	window['AscCommon'].getCurrencyFormat = getCurrencyFormat;
	window['AscCommon'].getFormatCells = getFormatCells;
	window['AscCommon'].canGetFormatByStandardId = canGetFormatByStandardId;
	window['AscCommon'].getFormatByStandardId = getFormatByStandardId;
	window['AscCommon'].is12HourTimeFormat = is12HourTimeFormat;
	window['AscCommon'].compareNumbers = compareNumbers;

    window["AscCommon"].gc_nMaxDigCount = gc_nMaxDigCount;
    window["AscCommon"].gc_nMaxDigCountView = gc_nMaxDigCountView;
    window["AscCommon"].oNumFormatCache = oNumFormatCache;
    window["AscCommon"].oGeneralEditFormatCache = oGeneralEditFormatCache;
    window["AscCommon"].g_oFormatParser = g_oFormatParser;
    window["AscCommon"].g_aCultureInfos = g_aCultureInfos;
    window["AscCommon"].g_oDefaultCultureInfo = g_oDefaultCultureInfo;
	window["AscCommon"].g_aAdditionalCurrencySymbols = g_aAdditionalCurrencySymbols;
	window["AscCommon"].NumFormatType = NumFormatType;

	window["AscCommon"].escapeRegExp = escapeRegExp;


})(window);
