/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

	$(function () {

		var cDate = Asc.cDate;

		function toFixed(n) {
			return n;//.toFixed( AscCommonExcel.cExcelSignificantDigits ) - 0;
		}

		function difBetween(a, b) {
			return Math.abs(a - b) < dif
		}

		function _getPMT(fZins, fZzr, fBw, fZw, nF) {
			var fRmz;
			if (fZins == 0.0) {
				fRmz = (fBw + fZw) / fZzr;
			} else {
				var fTerm = Math.pow(1.0 + fZins, fZzr);
				if (nF > 0) {
					fRmz = (fZw * fZins / (fTerm - 1.0) + fBw * fZins / (1.0 - 1.0 / fTerm)) / (1.0 + fZins);
				} else {
					fRmz = fZw * fZins / (fTerm - 1.0) + fBw * fZins / (1.0 - 1.0 / fTerm);
				}
			}

			return -fRmz;
		}

		function _getFV(fZins, fZzr, fRmz, fBw, nF) {
			var fZw;
			if (fZins == 0.0) {
				fZw = fBw + fRmz * fZzr;
			} else {
				var fTerm = Math.pow(1.0 + fZins, fZzr);
				if (nF > 0) {
					fZw = fBw * fTerm + fRmz * (1.0 + fZins) * (fTerm - 1.0) / fZins;
				} else {
					fZw = fBw * fTerm + fRmz * (fTerm - 1.0) / fZins;
				}
			}

			return -fZw;
		}

		function _getDDB(cost, salvage, life, period, factor) {
			var ddb, ipmt, oldCost, newCost;
			ipmt = factor / life;
			if (ipmt >= 1) {
				ipmt = 1;
				if (period == 1) {
					oldCost = cost;
				} else {
					oldCost = 0;
				}
			} else {
				oldCost = cost * Math.pow(1 - ipmt, period - 1);
			}
			newCost = cost * Math.pow(1 - ipmt, period);

			if (newCost < salvage) {
				ddb = oldCost - salvage;
			} else {
				ddb = oldCost - newCost;
			}
			if (ddb < 0) {
				ddb = 0;
			}
			return ddb;
		}

		function _getIPMT(rate, per, pv, type, pmt) {
			var ipmt;

			if (per == 1) {
				if (type > 0) {
					ipmt = 0;
				} else {
					ipmt = -pv;
				}
			} else {
				if (type > 0) {
					ipmt = _getFV(rate, per - 2, pmt, pv, 1) - pmt;
				} else {
					ipmt = _getFV(rate, per - 1, pmt, pv, 0);
				}
			}
			return ipmt * rate
		}

		function _diffDate(d1, d2, mode) {
			var date1 = d1.getDate(), month1 = d1.getMonth(), year1 = d1.getFullYear(), date2 = d2.getDate(), month2 = d2.getMonth(), year2 = d2.getFullYear();

			switch (mode) {
				case 0:
					return Math.abs(GetDiffDate360(date1, month1, year1, date2, month2, year2, true));
				case 1:
					var yc = Math.abs(year2 - year1), sd = year1 > year2 ? d2 : d1, yearAverage = sd.isLeapYear() ? 366 : 365, dayDiff = Math.abs(d2 - d1);
					for (var i = 0; i < yc; i++) {
						sd.addYears(1);
						yearAverage += sd.isLeapYear() ? 366 : 365;
					}
					yearAverage /= (yc + 1);
					dayDiff /= c_msPerDay;
					return dayDiff;
				case 2:
					var dayDiff = Math.abs(d2 - d1);
					dayDiff /= c_msPerDay;
					return dayDiff;
				case 3:
					var dayDiff = Math.abs(d2 - d1);
					dayDiff /= c_msPerDay;
					return dayDiff;
				case 4:
					return Math.abs(GetDiffDate360(date1, month1, year1, date2, month2, year2, false));
				default:
					return "#NUM!";
			}
		}

		function _yearFrac(d1, d2, mode) {
			var date1 = d1.getDate(), month1 = d1.getMonth() + 1, year1 = d1.getFullYear(), date2 = d2.getDate(), month2 = d2.getMonth() + 1, year2 = d2.getFullYear();

			switch (mode) {
				case 0:
					return Math.abs(GetDiffDate360(date1, month1, year1, date2, month2, year2, true)) / 360;
				case 1:
					var yc = /*Math.abs*/(year2 - year1), sd = year1 > year2 ? new cDate(d2) : new cDate(d1), yearAverage = sd.isLeapYear() ? 366 : 365,
						dayDiff = /*Math.abs*/(d2 - d1);
					for (var i = 0; i < yc; i++) {
						sd.addYears(1);
						yearAverage += sd.isLeapYear() ? 366 : 365;
					}
					yearAverage /= (yc + 1);
					dayDiff /= (yearAverage * c_msPerDay);
					return dayDiff;
				case 2:
					var dayDiff = Math.abs(d2 - d1);
					dayDiff /= (360 * c_msPerDay);
					return dayDiff;
				case 3:
					var dayDiff = Math.abs(d2 - d1);
					dayDiff /= (365 * c_msPerDay);
					return dayDiff;
				case 4:
					return Math.abs(GetDiffDate360(date1, month1, year1, date2, month2, year2, false)) / 360;
				default:
					return "#NUM!";
			}
		}

		function _lcl_GetCouppcd(settl, matur, freq) {
			matur.setFullYear(settl.getFullYear());
			if (matur < settl) {
				matur.addYears(1);
			}
			while (matur > settl) {
				matur.addMonths(-12 / freq);
			}
		}

		function _lcl_GetCoupncd(settl, matur, freq) {
			matur.setFullYear(settl.getFullYear());
			if (matur > settl) {
				matur.addYears(-1);
			}
			while (matur <= settl) {
				matur.addMonths(12 / freq);
			}
		}

		function _getcoupdaybs(settl, matur, frequency, basis) {
			_lcl_GetCouppcd(settl, matur, frequency);
			return _diffDate(settl, matur, basis);
		}

		function _getcoupdays(settl, matur, frequency, basis) {
			_lcl_GetCouppcd(settl, matur, frequency);
			var n = new cDate(matur);
			n.addMonths(12 / frequency);
			return _diffDate(matur, n, basis);
		}

		function _getdiffdate(d1, d2, nMode) {
			var bNeg = d1 > d2;

			if (bNeg) {
				var n = d2;
				d2 = d1;
				d1 = n;
			}

			var nRet, pOptDaysIn1stYear;

			var nD1 = d1.getDate(), nM1 = d1.getMonth(), nY1 = d1.getFullYear(), nD2 = d2.getDate(), nM2 = d2.getMonth(), nY2 = d2.getFullYear();

			switch (nMode) {
				case 0:			// 0=USA (NASD) 30/360
				case 4:			// 4=Europe 30/360
				{
					var bLeap = d1.isLeapYear();
					var nDays, nMonths/*, nYears*/;

					nMonths = nM2 - nM1;
					nDays = nD2 - nD1;

					nMonths += (nY2 - nY1) * 12;

					nRet = nMonths * 30 + nDays;
					if (nMode == 0 && nM1 == 2 && nM2 != 2 && nY1 == nY2) {
						nRet -= bLeap ? 1 : 2;
					}

					pOptDaysIn1stYear = 360;
				}
					break;
				case 1:			// 1=exact/exact
					pOptDaysIn1stYear = d1.isLeapYear() ? 366 : 365;
					nRet = d2 - d1;
					break;
				case 2:			// 2=exact/360
					nRet = d2 - d1;
					pOptDaysIn1stYear = 360;
					break;
				case 3:			//3=exact/365
					nRet = d2 - d1;
					pOptDaysIn1stYear = 365;
					break;
			}

			return (bNeg ? -nRet : nRet) / c_msPerDay / pOptDaysIn1stYear;
		}

		function _getprice(nSettle, nMat, fRate, fYield, fRedemp, nFreq, nBase) {

			var fdays = AscCommonExcel.getcoupdays(new cDate(nSettle), new cDate(nMat), nFreq, nBase),
				fdaybs = AscCommonExcel.getcoupdaybs(new cDate(nSettle), new cDate(nMat), nFreq, nBase), fnum = AscCommonExcel.getcoupnum(new cDate(nSettle), (nMat), nFreq, nBase),
				fdaysnc = (fdays - fdaybs) / fdays, fT1 = 100 * fRate / nFreq, fT2 = 1 + fYield / nFreq, res = fRedemp / (Math.pow(1 + fYield / nFreq, fnum - 1 + fdaysnc));

			/*var fRet = fRedemp / ( Math.pow( 1.0 + fYield / nFreq, fnum - 1.0 + fdaysnc ) );
        fRet -= 100.0 * fRate / nFreq * fdaybs / fdays;

        var fT1 = 100.0 * fRate / nFreq;
        var fT2 = 1.0 + fYield / nFreq;

        for( var fK = 0.0 ; fK < fnum ; fK++ ){
            fRet += fT1 / Math.pow( fT2, fK + fdaysnc );
        }

        return fRet;*/

			if (fnum == 1) {
				return (fRedemp + fT1) / (1 + fdaysnc * fYield / nFreq) - 100 * fRate / nFreq * fdaybs / fdays;
			}

			res -= 100 * fRate / nFreq * fdaybs / fdays;

			for (var i = 0; i < fnum; i++) {
				res += fT1 / Math.pow(fT2, i + fdaysnc);
			}

			return res;
		}

		function _getYield(nSettle, nMat, fCoup, fPrice, fRedemp, nFreq, nBase) {
			var fRate = fCoup, fPriceN = 0.0, fYield1 = 0.0, fYield2 = 1.0;
			var fPrice1 = _getprice(nSettle, nMat, fRate, fYield1, fRedemp, nFreq, nBase);
			var fPrice2 = _getprice(nSettle, nMat, fRate, fYield2, fRedemp, nFreq, nBase);
			var fYieldN = (fYield2 - fYield1) * 0.5;

			for (var nIter = 0; nIter < 100 && fPriceN != fPrice; nIter++) {
				fPriceN = _getprice(nSettle, nMat, fRate, fYieldN, fRedemp, nFreq, nBase);

				if (fPrice == fPrice1) {
					return fYield1;
				} else if (fPrice == fPrice2) {
					return fYield2;
				} else if (fPrice == fPriceN) {
					return fYieldN;
				} else if (fPrice < fPrice2) {
					fYield2 *= 2.0;
					fPrice2 = _getprice(nSettle, nMat, fRate, fYield2, fRedemp, nFreq, nBase);

					fYieldN = (fYield2 - fYield1) * 0.5;
				} else {
					if (fPrice < fPriceN) {
						fYield1 = fYieldN;
						fPrice1 = fPriceN;
					} else {
						fYield2 = fYieldN;
						fPrice2 = fPriceN;
					}

					fYieldN = fYield2 - (fYield2 - fYield1) * ((fPrice - fPrice2) / (fPrice1 - fPrice2));
				}
			}

			if (Math.abs(fPrice - fPriceN) > fPrice / 100.0) {
				return "#NUM!";
			}		// result not precise enough

			return fYieldN;
		}

		function _getyieldmat(nSettle, nMat, nIssue, fRate, fPrice, nBase) {

			var fIssMat = _yearFrac(nIssue, nMat, nBase);
			var fIssSet = _yearFrac(nIssue, nSettle, nBase);
			var fSetMat = _yearFrac(nSettle, nMat, nBase);

			var y = 1.0 + fIssMat * fRate;
			y /= fPrice / 100.0 + fIssSet * fRate;
			y--;
			y /= fSetMat;

			return y;

		}

		function _coupnum(settlement, maturity, frequency, basis) {

			basis = (basis !== undefined ? basis : 0);

			var n = new cDate(maturity);
			_lcl_GetCouppcd(settlement, n, frequency);
			var nMonths = (maturity.getFullYear() - n.getFullYear()) * 12 + maturity.getMonth() - n.getMonth();
			return nMonths * frequency / 12;

		}

		function _duration(settlement, maturity, coupon, yld, frequency, basis) {
			var dbc = AscCommonExcel.getcoupdaybs(new cDate(settlement), new cDate(maturity), frequency, basis),
				coupD = AscCommonExcel.getcoupdays(new cDate(settlement), new cDate(maturity), frequency, basis),
				numCoup = AscCommonExcel.getcoupnum(new cDate(settlement), new cDate(maturity), frequency);

			if (settlement >= maturity || basis < 0 || basis > 4 || (frequency != 1 && frequency != 2 && frequency != 4) || yld < 0 || coupon < 0) {
				return "#NUM!";
			}

			var duration = 0, p = 0;

			var dsc = coupD - dbc;
			var diff = dsc / coupD - 1;
			yld = yld / frequency + 1;


			coupon *= 100 / frequency;

			for (var index = 1; index <= numCoup; index++) {
				var di = index + diff;

				var yldPOW = Math.pow(yld, di);

				duration += di * coupon / yldPOW;

				p += coupon / yldPOW;
			}

			duration += (diff + numCoup) * 100 / Math.pow(yld, diff + numCoup);
			p += 100 / Math.pow(yld, diff + numCoup);

			return duration / p / frequency;
		}

		function numDivFact(num, fact) {
			var res = num / Math.fact(fact);
			res = res.toString();
			return res;
		}

		function testArrayFormula(assert, func, dNotSupportAreaArg) {

			var getValue = function (ref) {
				oParser = new parserFormula(func + "(" + ref + ")", "A2", ws);
				assert.ok(oParser.parse());
				return oParser.calculate().getValue();
			};

			//***array-formula***
			ws.getRange2("A100").setValue("1");
			ws.getRange2("B100").setValue("3");
			ws.getRange2("C100").setValue("-4");
			ws.getRange2("A101").setValue("2");
			ws.getRange2("B101").setValue("4");
			ws.getRange2("C101").setValue("5");


			oParser = new parserFormula(func + "(A100:C101)", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
			assert.ok(oParser.parse());
			var array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue("A100"));
				assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue("B100"));
				assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue("C100"));
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), getValue("A101"));
				assert.strictEqual(array.getElementRowCol(1, 1).getValue(), getValue("B101"));
				assert.strictEqual(array.getElementRowCol(1, 2).getValue(), getValue("C101"));
			} else {
				if (!dNotSupportAreaArg) {
					assert.strictEqual(false, true);
				}
				consoleLog("func: " + func + " don't return area array");
			}

			oParser = new parserFormula(func + "({1,2,-3})", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue(1));
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue(2));
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue(-3));
		}

		//returnOnlyValue - те функции, на вход которых всегда должны подаваться массивы и которые возвращают единственное значение
		function testArrayFormula2(assert, func, minArgCount, maxArgCount, dNotSupportAreaArg, returnOnlyValue) {

			var getValue = function (ref, countArg) {
				var argStr = "(";
				for (var j = 1; j <= countArg; j++) {
					argStr += ref;
					if (i !== j) {
						argStr += ",";
					} else {
						argStr += ")";
					}
				}
				oParser = new parserFormula(func + argStr, "A2", ws);
				assert.ok(oParser.parse());
				return oParser.calculate().getValue();
			};


			//***array-formula***
			ws.getRange2("A100").setValue("1");
			ws.getRange2("B100").setValue("3");
			ws.getRange2("C100").setValue("-4");
			ws.getRange2("A101").setValue("2");
			ws.getRange2("B101").setValue("4");
			ws.getRange2("C101").setValue("5");

			//формируем массив значений
			var randomArray = [];
			var randomStrArray = "{";
			var maxArg = 4;
			for (var i = 1; i <= maxArg; i++) {
				var randVal = Math.random();
				randomArray.push(randVal);
				randomStrArray += randVal;
				if (i !== maxArg) {
					randomStrArray += ",";
				} else {
					randomStrArray += "}";
				}
			}

			for (var i = minArgCount; i <= maxArgCount; i++) {
				var argStrArr = "(";
				var randomArgStrArr = "(";
				for (var j = 1; j <= i; j++) {
					argStrArr += "A100:C101";
					randomArgStrArr += randomStrArray;
					if (i !== j) {
						argStrArr += ",";
						randomArgStrArr += ",";
					} else {
						argStrArr += ")";
						randomArgStrArr += ")";
					}
				}

				oParser = new parserFormula(func + argStrArr, "A1", ws);
				oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
				assert.ok(oParser.parse());
				var array = oParser.calculate();
				if (AscCommonExcel.cElementType.array === array.type) {
					assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue("A100", i));
					assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue("B100", i));
					assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue("C100", i));
					assert.strictEqual(array.getElementRowCol(1, 0).getValue(), getValue("A101", i));
					assert.strictEqual(array.getElementRowCol(1, 1).getValue(), getValue("B101", i));
					assert.strictEqual(array.getElementRowCol(1, 2).getValue(), getValue("C101", i));
				} else {
					if (!(dNotSupportAreaArg || returnOnlyValue)) {
						assert.strictEqual(false, true);
					}
					consoleLog("func: " + func + " don't return area array");
				}

				oParser = new parserFormula(func + randomArgStrArr, "A1", ws);
				oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
				assert.ok(oParser.parse());
				array = oParser.calculate();
				if (AscCommonExcel.cElementType.array === array.type) {
					assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue(randomArray[0], i));
					assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue(randomArray[1], i));
					assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue(randomArray[2], i));
				} else {
					if (!returnOnlyValue) {
						assert.strictEqual(false, true);
					}
					consoleLog("func: " + func + " don't return array");
				}
			}
		}

		function testArrayFormulaEqualsValues(assert, str, formula, isNotLowerCase) {
			//***array-formula***
			ws.getRange2("A1").setValue("1");
			ws.getRange2("B1").setValue("3.123");
			ws.getRange2("C1").setValue("-4");
			ws.getRange2("A2").setValue("2");
			ws.getRange2("B2").setValue("4");
			ws.getRange2("C2").setValue("5");

			oParser = new parserFormula(formula, "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E6:H8").bbox);
			assert.ok(oParser.parse());
			var array = oParser.calculate();

			var splitStr = str.split(";");

			for (var i = 0; i < splitStr.length; i++) {
				var subSplitStr = splitStr[i].split(",");
				for (var j = 0; j < subSplitStr.length; j++) {
					var valMs = subSplitStr[j];
					var element;
					if (array.getElementRowCol) {
						var row = 1 === array.array.length ? 0 : i;
						var col = 1 === array.array[0].length ? 0 : j;
						if (array.array[row] && array.array[row][col]) {
							element = array.getElementRowCol(row, col);
						} else {
							element = new window['AscCommonExcel'].cError(window['AscCommonExcel'].cErrorType.not_available);
						}
					} else {
						element = array;
					}
					var ourVal = element && undefined != element.value ? element.value.toString() : "#N/A";
					if (!isNotLowerCase) {
						valMs = valMs.toLowerCase();
						ourVal = ourVal.toLowerCase();
					}
					assert.strictEqual(valMs, ourVal, "formula: " + formula + " i: " + i + " j: " + j)
				}
			}
		}

		function _getValue(from, row, col) {
			var res;
			if (from.type === AscCommonExcel.cElementType.array) {
				res = from.getElementRowCol(row !== undefined ? row : 0, col !== undefined ? col : 0).getValue();
			} else if (from.type === AscCommonExcel.cElementType.cellsRange || from.type === AscCommonExcel.cElementType.cellsRange3D) {
				res = from.getValueByRowCol(row !== undefined ? row : 0, col !== undefined ? col : 0).getValue();
			} else if (from.type === AscCommonExcel.cElementType.cell || from.type === AscCommonExcel.cElementType.cell3D) {
				res = from.getValue().getValue();
			} else {
				res = from.getValue();
			}
			return res;
		}

		function consoleLog(val) {
			//console.log(val);
		}

		var newFormulaParser = false;

		var c_msPerDay = AscCommonExcel.c_msPerDay;
		var parserFormula = AscCommonExcel.parserFormula;
		var GetDiffDate360 = AscCommonExcel.GetDiffDate360;
		var fSortAscending = AscCommon.fSortAscending;
		var g_oIdCounter = AscCommon.g_oIdCounter;

		var oParser, wb, ws, dif = 1e-9, sData = AscCommon.getEmpty(), tmp, array;
		if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {
			let api = new Asc.spreadsheet_api({
				'id-view': 'editor_sdk'
			});

			let docInfo = new Asc.asc_CDocInfo();
			docInfo.asc_putTitle("TeSt.xlsx");
			api.DocInfo = docInfo;

			window["Asc"]["editor"] = api;

			wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api);
			AscCommon.History.init(wb);
			wb.maxDigitWidth = 7;
			wb.paddingPlusBorder = 5;

			AscCommon.g_oTableId.init();
			if (this.User) {
				g_oIdCounter.Set_UserId(this.User.asc_getId());
			}

			AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
			AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
			AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
			AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
			AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
			AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
			AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
			AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
			g_oIdCounter.Set_Load(false);

			var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
			oBinaryFileReader.Read(sData, wb);
			ws = wb.getWorksheet(wb.getActive());
			AscCommonExcel.getFormulasInfo();
		}

		wb.dependencyFormulas.lockRecal();

		QUnit.module("Formula");
		QUnit.test("Test: \"ABS\"", function (assert) {

			ws.getRange2("A22").setValue("-4");

			oParser = new parserFormula("ABS(2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);
			oParser = new parserFormula("ABS(-2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);
			oParser = new parserFormula("ABS(A22)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			testArrayFormula(assert, "ABS");
		});

		QUnit.test("Test: \"Absolute reference\"", function (assert) {

			ws.getRange2("A7").setValue("1");
			ws.getRange2("A8").setValue("2");
			ws.getRange2("A9").setValue("3");
			oParser = new parserFormula('A$7+A8', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula('A$7+A$8', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula('$A$7+$A$8', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula('SUM($A$7:$A$9)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 6);
		});

		QUnit.test("Test: \"Asc\"", function (assert) {
			oParser = new parserFormula('ASC("ｔｅＳｔ")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "teSt");

			oParser = new parserFormula('ASC("デジタル")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "デジタル");

			oParser = new parserFormula('ASC("￯")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "");
		});

		QUnit.test("Test: \"Cross\"", function (assert) {

			ws.getRange2("A7").setValue("1");
			ws.getRange2("A8").setValue("2");
			ws.getRange2("A9").setValue("3");
			oParser = new parserFormula('A7:A9', null, ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 5, 0, 5), ws.getId()).getValue(), "#VALUE!");
			assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 6, 0, 6), ws.getId()).getValue(), 1);
			assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 7, 0, 7), ws.getId()).getValue(), 2);
			assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 8, 0, 8), ws.getId()).getValue(), 3);
			assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 9, 0, 9), ws.getId()).getValue(), "#VALUE!");

		});

		QUnit.test("Test: \"Defined names cycle\"", function (assert) {

			var newNameQ = new Asc.asc_CDefName("q", "SUM('" + ws.getName() + "'!A2)");
			wb.editDefinesNames(null, newNameQ);
			ws.getRange2("Q1").setValue("=q");
			ws.getRange2("Q2").setValue("=q");
			ws.getRange2("Q3").setValue("1");
			assert.strictEqual(ws.getRange2("Q1").getValueWithFormat(), "1");
			assert.strictEqual(ws.getRange2("Q2").getValueWithFormat(), "1");

			var newNameW = new Asc.asc_CDefName("w", "'" + ws.getName() + "'!A1");
			wb.editDefinesNames(null, newNameW);
			ws.getRange2("Q4").setValue("=w");
			assert.strictEqual(ws.getRange2("Q4").getValueWithFormat(), "#REF!");
			//clean up
			ws.getRange2("Q1:Q4").cleanAll();
			wb.delDefinesNames(newNameW);
			wb.delDefinesNames(newNameQ);
		});

		QUnit.test("Test: \"Parse intersection\"", function (assert) {

			ws.getRange2("A7").setValue("1");
			ws.getRange2("A8").setValue("2");
			ws.getRange2("A9").setValue("3");
			oParser = new parserFormula('1     +    (    A7   +A8   )   *   2', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.assemble(), "1+(A7+A8)*2");
			assert.strictEqual(oParser.calculate().getValue(), 7);

			oParser = new parserFormula('sum                    A1:A5', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.assemble(), "sum A1:A5");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('sum(   A1:A5    ,        B1:B5     )     ', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.assemble(), "SUM(A1:A5,B1:B5)");
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula('sum(   A1:A5    ,        B1:B5  , "    3 , 14 15 92 6 "   )     ', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.assemble(), 'SUM(A1:A5,B1:B5,"    3 , 14 15 92 6 ")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		});


		QUnit.test("Test: \"Arithmetical operations\"", function (assert) {
			oParser = new parserFormula('1+3', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula('(1+2)*4+3', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), (1 + 2) * 4 + 3);

			oParser = new parserFormula('2^52', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), Math.pow(2, 52));

			oParser = new parserFormula('-10', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -10);

			oParser = new parserFormula('-10*2', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -20);

			oParser = new parserFormula('-10+10', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula('12%', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0.12);

			oParser = new parserFormula("2<>\"3\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", "2<>\"3\"");

			oParser = new parserFormula("2=\"3\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", "2=\"3\"");

			oParser = new parserFormula("2>\"3\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", "2>\"3\"");

			oParser = new parserFormula("\"f\">\"3\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula("\"f\"<\"3\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual("FALSE", oParser.calculate().getValue(), "FALSE");

			oParser = new parserFormula("FALSE>=FALSE", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula("\"TRUE\"&\"TRUE\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUETRUE");

			oParser = new parserFormula("10*\"\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("-TRUE", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -1);


			ws.getRange2("M106").setValue("1");
			ws.getRange2("M107").setValue("2");
			ws.getRange2("M108").setValue("2");
			ws.getRange2("M109").setValue("4");
			ws.getRange2("M110").setValue("5");
			ws.getRange2("M111").setValue("-23");
			ws.getRange2("M112").setValue("6");
			ws.getRange2("M113").setValue("5");

			ws.getRange2("N106").setValue("1");
			ws.getRange2("N107").setValue("");
			ws.getRange2("N108").setValue("");
			ws.getRange2("N109").setValue("3");
			ws.getRange2("N110").setValue("");
			ws.getRange2("N111").setValue("2");
			ws.getRange2("N112").setValue("");
			ws.getRange2("N113").setValue("3");

			ws.getRange2("O106").setValue("1");
			ws.getRange2("O107").setValue("3");
			ws.getRange2("O108").setValue("2");
			ws.getRange2("O109").setValue("12");
			ws.getRange2("O110").setValue("3");
			ws.getRange2("O111").setValue("4");
			ws.getRange2("O112").setValue("3");
			ws.getRange2("O113").setValue("2");

			ws.getRange2("P106").setValue("3");
			ws.getRange2("P107").setValue("4");
			ws.getRange2("P108").setValue("5");
			ws.getRange2("P109").setValue("1");
			ws.getRange2("P110").setValue("23");
			ws.getRange2("P111").setValue("4");
			ws.getRange2("P112").setValue("3");
			ws.getRange2("P113").setValue("1");

			oParser = new parserFormula("M106:P113+M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			var array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2);
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 8);
				assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 10);
				assert.strictEqual(array.getElementRowCol(5, 0).getValue(), -46);
				assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 12);
				assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 10);

				assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 2);
				assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 6);
				assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(7, 1).getValue(), 6);

				assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 2);
				assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 6);
				assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 24);
				assert.strictEqual(array.getElementRowCol(4, 2).getValue(), 6);
				assert.strictEqual(array.getElementRowCol(5, 2).getValue(), 8);
				assert.strictEqual(array.getElementRowCol(6, 2).getValue(), 6);
				assert.strictEqual(array.getElementRowCol(7, 2).getValue(), 4);

				assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 6);
				assert.strictEqual(array.getElementRowCol(1, 3).getValue(), 8);
				assert.strictEqual(array.getElementRowCol(2, 3).getValue(), 10);
				assert.strictEqual(array.getElementRowCol(3, 3).getValue(), 2);
				assert.strictEqual(array.getElementRowCol(4, 3).getValue(), 46);
				assert.strictEqual(array.getElementRowCol(5, 3).getValue(), 8);
				assert.strictEqual(array.getElementRowCol(6, 3).getValue(), 6);
				assert.strictEqual(array.getElementRowCol(7, 3).getValue(), 2);

			}

			oParser = new parserFormula("M106:P113*M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 16);
				assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 25);
				assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 529);
				assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 36);
				assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 25);

				assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1);
				assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 9);
				assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(7, 1).getValue(), 9);

				assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1);
				assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 9);
				assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 4);
				assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 144);
				assert.strictEqual(array.getElementRowCol(4, 2).getValue(), 9);
				assert.strictEqual(array.getElementRowCol(5, 2).getValue(), 16);
				assert.strictEqual(array.getElementRowCol(6, 2).getValue(), 9);
				assert.strictEqual(array.getElementRowCol(7, 2).getValue(), 4);

				assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 9);
				assert.strictEqual(array.getElementRowCol(1, 3).getValue(), 16);
				assert.strictEqual(array.getElementRowCol(2, 3).getValue(), 25);
				assert.strictEqual(array.getElementRowCol(3, 3).getValue(), 1);
				assert.strictEqual(array.getElementRowCol(4, 3).getValue(), 529);
				assert.strictEqual(array.getElementRowCol(5, 3).getValue(), 16);
				assert.strictEqual(array.getElementRowCol(6, 3).getValue(), 9);
				assert.strictEqual(array.getElementRowCol(7, 3).getValue(), 1);
			}

			oParser = new parserFormula("M106:P113-M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 0);
				assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 0);
			}

			oParser = new parserFormula("M106:P113=M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE");
			}

			oParser = new parserFormula("M106:P113/M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1);
				assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!");
			}

			oParser = new parserFormula("M106:P113<>M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE");
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE");
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE");
			}

			oParser = new parserFormula("M106:P113>M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE");
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE");
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE");
			}

			oParser = new parserFormula("M106:P113<M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE");
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE");
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE");
			}

			oParser = new parserFormula("M106:P113>=M106:P113", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE");
				assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE");
			}

			oParser = new parserFormula("SUM(M:P*M:P)", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			assert.strictEqual(array.getValue(), 1465);

			oParser = new parserFormula("SUM(M:P+M:P)", "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
			assert.ok(oParser.parse());
			array = oParser.calculate();
			assert.strictEqual(array.getValue(), 170);

		});

		QUnit.test("Test: \"ACOS\"", function (assert) {
			oParser = new parserFormula('ACOS(-0.5)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 2.094395102);

			testArrayFormula(assert, "ACOS");
		});

		QUnit.test("Test: \"ACOSH\"", function (assert) {
			oParser = new parserFormula('ACOSH(1)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula('ACOSH(10)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 2.9932228);

			testArrayFormula(assert, "ACOSH");
		});

		QUnit.test("Test: \"ASIN\"", function (assert) {
			oParser = new parserFormula('ASIN(-0.5)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, -0.523598776);

			testArrayFormula(assert, "ASIN");
		});

		QUnit.test("Test: \"ASINH\"", function (assert) {
			oParser = new parserFormula('ASINH(-2.5)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, -1.647231146);

			oParser = new parserFormula('ASINH(10)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 2.99822295);

			testArrayFormula(assert, "ASINH");
		});

		QUnit.test("Test: \"SIN have wrong arguments count\"", function (assert) {
			oParser = new parserFormula('SIN(3.1415926,3.1415926*2)', "A1", ws);
			assert.ok(!oParser.parse());
		});

		QUnit.test("Test: \"SIN(3.1415926)\"", function (assert) {
			oParser = new parserFormula('SIN(3.1415926)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), Math.sin(3.1415926));

			testArrayFormula(assert, "SIN");
		});

		QUnit.test("Test: \"SQRT\"", function (assert) {
			ws.getRange2("A202").setValue("-16");

			oParser = new parserFormula('SQRT(16)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula('SQRT(A202)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula('SQRT(ABS(A202))', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			testArrayFormula(assert, "SQRT");
		});

		QUnit.test("Test: \"SQRTPI\"", function (assert) {
			oParser = new parserFormula('SQRTPI(1)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 1.772454);

			oParser = new parserFormula('SQRTPI(2)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 2.506628);

			testArrayFormula(assert, "SQRTPI", true);
		});

		QUnit.test("Test: \"COS(PI()/2)\"", function (assert) {
			oParser = new parserFormula('COS(PI()/2)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), Math.cos(Math.PI / 2));
		});

		QUnit.test("Test: \"ACOT(2)\"", function (assert) {
			oParser = new parserFormula('ACOT(2)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), Math.PI / 2 - Math.atan(2));
		});

		QUnit.test("Test: \"ACOTH(6)\"", function (assert) {
			oParser = new parserFormula('ACOTH(6)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), Math.atanh(1 / 6));

			testArrayFormula(assert, "ACOTH");
		});


		QUnit.test("Test: \"COT\"", function (assert) {
			oParser = new parserFormula('COT(30)', "A1", ws);
			assert.ok(oParser.parse(), 'COT(30)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(3) - 0, -0.156, 'COT(30)');

			oParser = new parserFormula('COT(0)', "A1", ws);
			assert.ok(oParser.parse(), 'COT(0)');
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'COT(0)');

			oParser = new parserFormula('COT(1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'COT(1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'COT(1000000000)');

			oParser = new parserFormula('COT(-1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'COT(-1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'COT(-1000000000)');

			oParser = new parserFormula('COT(test)', "A1", ws);
			assert.ok(oParser.parse(), 'COT(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'COT(test)');

			oParser = new parserFormula('COT("test")', "A1", ws);
			assert.ok(oParser.parse(), 'COT("test")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'COT("test")');

			testArrayFormula(assert, "COT");
		});

		QUnit.test("Test: \"COTH\"", function (assert) {
			oParser = new parserFormula('COTH(2)', "A1", ws);
			assert.ok(oParser.parse(), 'COTH(2)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(3) - 0, 1.037, 'COTH(2)');

			oParser = new parserFormula('COTH(0)', "A1", ws);
			assert.ok(oParser.parse(), 'COTH(0)');
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'COTH(0)');

			oParser = new parserFormula('COTH(1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'COTH(1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), 1, 'COTH(1000000000)');

			oParser = new parserFormula('COTH(-1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'COTH(-1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), -1, 'COTH(-1000000000)');

			oParser = new parserFormula('COTH(test)', "A1", ws);
			assert.ok(oParser.parse(), 'COTH(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'COTH(test)');

			oParser = new parserFormula('COTH("test")', "A1", ws);
			assert.ok(oParser.parse(), 'COTH("test")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'COTH("test")');

			testArrayFormula(assert, "COTH");
		});

		QUnit.test("Test: \"CSC\"", function (assert) {
			oParser = new parserFormula('CSC(15)', "A1", ws);
			assert.ok(oParser.parse(), 'CSC(15)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(3) - 0, 1.538, 'CSC(15)');

			oParser = new parserFormula('CSC(0)', "A1", ws);
			assert.ok(oParser.parse(), 'CSC(0)');
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'CSC(0)');

			oParser = new parserFormula('CSC(1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'CSC(1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'CSC(1000000000)');

			oParser = new parserFormula('CSC(-1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'CSC(-1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'CSC(-1000000000)');

			oParser = new parserFormula('CSC(test)', "A1", ws);
			assert.ok(oParser.parse(), 'CSC(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'CSC(test)');

			oParser = new parserFormula('CSC("test")', "A1", ws);
			assert.ok(oParser.parse(), 'CSC("test")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'CSC("test")');

			testArrayFormula(assert, "CSC");
		});

		QUnit.test("Test: \"CSCH\"", function (assert) {
			oParser = new parserFormula('CSCH(1.5)', "A1", ws);
			assert.ok(oParser.parse(), 'CSCH(1.5)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(4) - 0, 0.4696, 'CSCH(1.5)');

			oParser = new parserFormula('CSCH(0)', "A1", ws);
			assert.ok(oParser.parse(), 'CSCH(0)');
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'CSCH(0)');

			oParser = new parserFormula('CSCH(1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'CSCH(1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), 0, 'CSCH(1000000000)');

			oParser = new parserFormula('CSCH(-1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'CSCH(-1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), 0, 'CSCH(-1000000000)');

			oParser = new parserFormula('CSCH(test)', "A1", ws);
			assert.ok(oParser.parse(), 'CSCH(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'CSCH(test)');

			oParser = new parserFormula('CSCH("test")', "A1", ws);
			assert.ok(oParser.parse(), 'CSCH("test")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'CSCH("test")');

			testArrayFormula(assert, "CSCH");
		});

		QUnit.test("Test: \"CLEAN\"", function (assert) {
			ws.getRange2("A202").setValue('=CHAR(9)&"Monthly report"&CHAR(10)');

			oParser = new parserFormula('CLEAN(A202)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "Monthly report");

			testArrayFormula(assert, "CLEAN");
		});

		QUnit.test("Test: \"DEGREES\"", function (assert) {
			oParser = new parserFormula('DEGREES(PI())', "A1", ws);
			assert.ok(oParser.parse(), 'DEGREES(PI())');
			assert.strictEqual(oParser.calculate().getValue(), 180, 'DEGREES(PI())');

			testArrayFormula(assert, "DEGREES");
		});

		QUnit.test("Test: \"SEC\"", function (assert) {
			oParser = new parserFormula('SEC(45)', "A1", ws);
			assert.ok(oParser.parse(), 'SEC(45)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(5) - 0, 1.90359, 'SEC(45)');

			oParser = new parserFormula('SEC(30)', "A1", ws);
			assert.ok(oParser.parse(), 'SEC(30)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(5) - 0, 6.48292, 'SEC(30)');

			oParser = new parserFormula('SEC(0)', "A1", ws);
			assert.ok(oParser.parse(), 'SEC(0)');
			assert.strictEqual(oParser.calculate().getValue(), 1, 'SEC(0)');

			oParser = new parserFormula('SEC(1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'SEC(1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'SEC(1000000000)');

			oParser = new parserFormula('SEC(test)', "A1", ws);
			assert.ok(oParser.parse(), 'SEC(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SEC(test)');

			oParser = new parserFormula('SEC("test")', "A1", ws);
			assert.ok(oParser.parse(), 'SEC("test")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'SEC("test")');

			testArrayFormula(assert, "SEC");
		});

		QUnit.test("Test: \"SECH\"", function (assert) {
			oParser = new parserFormula('SECH(5)', "A1", ws);
			assert.ok(oParser.parse(), 'SECH(5)');
			assert.strictEqual(oParser.calculate().getValue().toFixed(3) - 0, 0.013, 'SECH(5)');

			oParser = new parserFormula('SECH(0)', "A1", ws);
			assert.ok(oParser.parse(), 'SECH(0)');
			assert.strictEqual(oParser.calculate().getValue(), 1, 'SECH(0)');

			oParser = new parserFormula('SECH(1000000000)', "A1", ws);
			assert.ok(oParser.parse(), 'SECH(1000000000)');
			assert.strictEqual(oParser.calculate().getValue(), 0, 'SECH(1000000000)');

			oParser = new parserFormula('SECH(test)', "A1", ws);
			assert.ok(oParser.parse(), 'SECH(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SECH(test)');

			oParser = new parserFormula('SECH("test")', "A1", ws);
			assert.ok(oParser.parse(), 'SECH("test")');
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'SECH("test")');

			testArrayFormula(assert, "SECH");
		});

		QUnit.test("Test: \"SECOND\"", function (assert) {

			ws.getRange2("A202").setValue("12:45:03 PM");
			ws.getRange2("A203").setValue("4:48:18 PM");
			ws.getRange2("A204").setValue("4:48 PM");

			oParser = new parserFormula("SECOND(A202)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula("SECOND(A203)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 18);

			oParser = new parserFormula("SECOND(A204)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			testArrayFormula2(assert, "SECOND", 1, 1);
		});

		QUnit.test("Test: \"FLOOR\"", function (assert) {
			oParser = new parserFormula('FLOOR(3.7,2)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR(3.7,2)');
			assert.strictEqual(oParser.calculate().getValue(), 2, 'FLOOR(3.7,2)');

			oParser = new parserFormula('FLOOR(-2.5,-2)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR(-2.5,-2)');
			assert.strictEqual(oParser.calculate().getValue(), -2, 'FLOOR(-2.5,-2)');

			oParser = new parserFormula('FLOOR(2.5,-2)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR(2.5,-2)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'FLOOR(2.5,-2)');

			oParser = new parserFormula('FLOOR(1.58,0.1)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR(1.58,0.1)');
			assert.strictEqual(oParser.calculate().getValue(), 1.5, 'FLOOR(1.58,0.1)');

			oParser = new parserFormula('FLOOR(0.234,0.01)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR(0.234,0.01)');
			assert.strictEqual(oParser.calculate().getValue(), 0.23, 'FLOOR(0.234,0.01)');

			testArrayFormula2(assert, "FLOOR", 2, 2);
		});

		QUnit.test("Test: \"FLOOR.PRECISE\"", function (assert) {
			oParser = new parserFormula('FLOOR.PRECISE(-3.2, -1)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.PRECISE(-3.2, -1)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'FLOOR.PRECISE(-3.2, -1)');

			oParser = new parserFormula('FLOOR.PRECISE(3.2, 1)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.PRECISE(3.2, 1)');
			assert.strictEqual(oParser.calculate().getValue(), 3, 'FLOOR.PRECISE(3.2, 1)');

			oParser = new parserFormula('FLOOR.PRECISE(-3.2, 1)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.PRECISE(-3.2, 1)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'FLOOR.PRECISE(-3.2, 1)');

			oParser = new parserFormula('FLOOR.PRECISE(3.2, -1)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.PRECISE(3.2, -1)');
			assert.strictEqual(oParser.calculate().getValue(), 3, 'FLOOR.PRECISE(3.2, -1)');

			oParser = new parserFormula('FLOOR.PRECISE(3.2)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.PRECISE(3.2)');
			assert.strictEqual(oParser.calculate().getValue(), 3, 'FLOOR.PRECISE(3.2)');

			oParser = new parserFormula('FLOOR.PRECISE(test)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.PRECISE(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'FLOOR.PRECISE(test)');

			testArrayFormula2(assert, "FLOOR.PRECISE", 1, 2);
		});

		QUnit.test("Test: \"FLOOR.MATH\"", function (assert) {
			oParser = new parserFormula('FLOOR.MATH(24.3, 5)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.MATH(24.3, 5)');
			assert.strictEqual(oParser.calculate().getValue(), 20, 'FLOOR.MATH(24.3, 5)');

			oParser = new parserFormula('FLOOR.MATH(6.7)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.MATH(6.7)');
			assert.strictEqual(oParser.calculate().getValue(), 6, 'FLOOR.MATH(6.7)');

			oParser = new parserFormula('FLOOR.MATH(-8.1, 5)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.MATH(-8.1, 5)');
			assert.strictEqual(oParser.calculate().getValue(), -10, 'FLOOR.MATH(-8.1, 5)');

			oParser = new parserFormula('FLOOR.MATH(-5.5, 2, -1)', "A1", ws);
			assert.ok(oParser.parse(), 'FLOOR.MATH(-5.5, 2, -1)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'FLOOR.MATH(-5.5, 2, -1)');

			testArrayFormula2(assert, "FLOOR.MATH", 1, 3);
		});

		QUnit.test("Test: \"CEILING.MATH\"", function (assert) {
			oParser = new parserFormula('CEILING.MATH(24.3, 5)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.MATH(24.3, 5)');
			assert.strictEqual(oParser.calculate().getValue(), 25, 'CEILING.MATH(24.3, 5)');

			oParser = new parserFormula('CEILING.MATH(6.7)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.MATH(6.7)');
			assert.strictEqual(oParser.calculate().getValue(), 7, 'CEILING.MATH(6.7)');

			oParser = new parserFormula('CEILING.MATH(-8.1, 2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.MATH(-8.1, 2)');
			assert.strictEqual(oParser.calculate().getValue(), -8, 'CEILING.MATH(-8.1, 2)');

			oParser = new parserFormula('CEILING.MATH(-5.5, 2, -1)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.MATH(-5.5, 2, -1)');
			assert.strictEqual(oParser.calculate().getValue(), -6, 'CEILING.MATH(-5.5, 2, -1)');

			testArrayFormula2(assert, "CEILING.MATH", 1, 3);
		});

		QUnit.test("Test: \"CEILING.PRECISE\"", function (assert) {
			oParser = new parserFormula('CEILING.PRECISE(4.3)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(4.3)');
			assert.strictEqual(oParser.calculate().getValue(), 5, 'CEILING.PRECISE(4.3)');

			oParser = new parserFormula('CEILING.PRECISE(-4.3)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(-4.3)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'CEILING.PRECISE(-4.3)');

			oParser = new parserFormula('CEILING.PRECISE(4.3, 2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(4.3, 2)');
			assert.strictEqual(oParser.calculate().getValue(), 6, 'CEILING.PRECISE(4.3, 2)');

			oParser = new parserFormula('CEILING.PRECISE(4.3,-2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(4.3,-2)');
			assert.strictEqual(oParser.calculate().getValue(), 6, 'CEILING.PRECISE(4.3,-2)');

			oParser = new parserFormula('CEILING.PRECISE(-4.3,2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(-4.3,2)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'CEILING.PRECISE(-4.3,2)');

			oParser = new parserFormula('CEILING.PRECISE(-4.3,-2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(-4.3,-2)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'CEILING.PRECISE(-4.3,-2)');

			oParser = new parserFormula('CEILING.PRECISE(test)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING.PRECISE(test)');
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'CEILING.PRECISE(test)');

			testArrayFormula2(assert, "CEILING.PRECISE", 1, 2);
		});

		QUnit.test("Test: \"ISO.CEILING\"", function (assert) {
			oParser = new parserFormula('ISO.CEILING(4.3)', "A1", ws);
			assert.ok(oParser.parse(), 'ISO.CEILING(4.3)');
			assert.strictEqual(oParser.calculate().getValue(), 5, 'ISO.CEILING(4.3)');

			oParser = new parserFormula('ISO.CEILING(-4.3)', "A1", ws);
			assert.ok(oParser.parse(), 'ISO.CEILING(-4.3)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'ISO.CEILING(-4.3)');

			oParser = new parserFormula('ISO.CEILING(4.3, 2)', "A1", ws);
			assert.ok(oParser.parse(), 'ISO.CEILING(4.3, 2)');
			assert.strictEqual(oParser.calculate().getValue(), 6, 'ISO.CEILING(4.3, 2)');

			oParser = new parserFormula('ISO.CEILING(4.3,-2)', "A1", ws);
			assert.ok(oParser.parse(), 'ISO.CEILING(4.3,-2)');
			assert.strictEqual(oParser.calculate().getValue(), 6, 'ISO.CEILING(4.3,-2)');

			oParser = new parserFormula('ISO.CEILING(-4.3,2)', "A1", ws);
			assert.ok(oParser.parse(), 'ISO.CEILING(-4.3,2)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'ISO.CEILING(-4.3,2)');

			oParser = new parserFormula('ISO.CEILING(-4.3,-2)', "A1", ws);
			assert.ok(oParser.parse(), 'ISO.CEILING(-4.3,-2)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'ISO.CEILING(-4.3,-2)');

			testArrayFormula2(assert, "ISO.CEILING", 1, 2);
		});

		QUnit.test("Test: \"ISBLANK\"", function (assert) {

			ws.getRange2("A202").setValue("");
			ws.getRange2("A203").setValue("test");

			oParser = new parserFormula('ISBLANK(A202)', "A1", ws);
			assert.ok(oParser.parse(), 'ISBLANK(A202)');
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISBLANK(A202)');

			oParser = new parserFormula('ISBLANK(A203)', "A1", ws);
			assert.ok(oParser.parse(), 'ISBLANK(A203)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISBLANK(A203)');

			testArrayFormula2(assert, "ISBLANK", 1, 1);
		});

		QUnit.test("Test: \"ISERROR\"", function (assert) {

			ws.getRange2("A202").setValue("");
			ws.getRange2("A203").setValue("#N/A");

			oParser = new parserFormula('ISERROR(A202)', "A1", ws);
			assert.ok(oParser.parse(), 'ISERROR(A202)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERROR(A202)');

			oParser = new parserFormula('ISERROR(A203)', "A1", ws);
			assert.ok(oParser.parse(), 'ISERROR(A203)');
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISERROR(A203)');

			testArrayFormula2(assert, "ISERROR", 1, 1);
		});

		QUnit.test("Test: \"ISERR\"", function (assert) {

			ws.getRange2("A202").setValue("");
			ws.getRange2("A203").setValue("#N/A");
			ws.getRange2("A204").setValue("#VALUE!");

			oParser = new parserFormula('ISERR(A202)', "A1", ws);
			assert.ok(oParser.parse(), 'ISERR(A202)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERR(A202)');

			oParser = new parserFormula('ISERR(A203)', "A1", ws);
			assert.ok(oParser.parse(), 'ISERR(A203)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERR(A203)');

			oParser = new parserFormula('ISERR(A203)', "A1", ws);
			assert.ok(oParser.parse(), 'ISERR(A203)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERR(A203)');

			testArrayFormula2(assert, "ISERR", 1, 1);
		});

		QUnit.test("Test: \"ISEVEN\"", function (assert) {

			oParser = new parserFormula('ISEVEN(-1)', "A1", ws);
			assert.ok(oParser.parse(), 'ISEVEN(-1)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISEVEN(-1)');

			oParser = new parserFormula('ISEVEN(2.5)', "A1", ws);
			assert.ok(oParser.parse(), 'ISEVEN(2.5)');
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISEVEN(2.5)');

			oParser = new parserFormula('ISEVEN(5)', "A1", ws);
			assert.ok(oParser.parse(), 'ISEVEN(5)');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISEVEN(5)');

			oParser = new parserFormula('ISEVEN(0)', "A1", ws);
			assert.ok(oParser.parse(), 'ISEVEN(0)');
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISEVEN(0)');

			oParser = new parserFormula('ISEVEN(12/23/2011)', "A1", ws);
			assert.ok(oParser.parse(), 'ISEVEN(12/23/2011)');
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISEVEN(12/23/2011)');

			testArrayFormula2(assert, "ISEVEN", 1, 1, true, null);
		});

		QUnit.test("Test: \"ISLOGICAL\"", function (assert) {

			oParser = new parserFormula('ISLOGICAL(TRUE)', "A1", ws);
			assert.ok(oParser.parse(), 'ISLOGICAL(TRUE)');
			assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISLOGICAL(TRUE)');

			oParser = new parserFormula('ISLOGICAL("TRUE")', "A1", ws);
			assert.ok(oParser.parse(), 'ISLOGICAL("TRUE")');
			assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISLOGICAL("TRUE")');

			testArrayFormula2(assert, "ISLOGICAL", 1, 1);
		});

		QUnit.test("Test: \"CEILING\"", function (assert) {

			oParser = new parserFormula('CEILING(2.5, 1)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING(2.5, 1)');
			assert.strictEqual(oParser.calculate().getValue(), 3, 'CEILING(2.5, 1)');

			oParser = new parserFormula('CEILING(-2.5, -2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING(-2.5, -2)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'CEILING(-2.5, -2)');

			oParser = new parserFormula('CEILING(-2.5, 2)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING(-2.5, 2)');
			assert.strictEqual(oParser.calculate().getValue(), -2, 'CEILING(-2.5, 2)');

			oParser = new parserFormula('CEILING(1.5, 0.1)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING(1.5, 0.1)');
			assert.strictEqual(oParser.calculate().getValue(), 1.5, 'CEILING(1.5, 0.1)');

			oParser = new parserFormula('CEILING(0.234, 0.01)', "A1", ws);
			assert.ok(oParser.parse(), 'CEILING(0.234, 0.01)');
			assert.strictEqual(oParser.calculate().getValue(), 0.24, 'CEILING(0.234, 0.01)');

			testArrayFormula2(assert, "CEILING", 2, 2);
		});

		QUnit.test("Test: \"ECMA.CEILING\"", function (assert) {

			oParser = new parserFormula('ECMA.CEILING(2.5, 1)', "A1", ws);
			assert.ok(oParser.parse(), 'ECMA.CEILING(2.5, 1)');
			assert.strictEqual(oParser.calculate().getValue(), 3, 'ECMA.CEILING(2.5, 1)');

			oParser = new parserFormula('ECMA.CEILING(-2.5, -2)', "A1", ws);
			assert.ok(oParser.parse(), 'ECMA.CEILING(-2.5, -2)');
			assert.strictEqual(oParser.calculate().getValue(), -4, 'ECMA.CEILING(-2.5, -2)');

			oParser = new parserFormula('ECMA.CEILING(-2.5, 2)', "A1", ws);
			assert.ok(oParser.parse(), 'ECMA.CEILING(-2.5, 2)');
			assert.strictEqual(oParser.calculate().getValue(), -2, 'ECMA.CEILING(-2.5, 2)');

			oParser = new parserFormula('ECMA.CEILING(1.5, 0.1)', "A1", ws);
			assert.ok(oParser.parse(), 'ECMA.CEILING(1.5, 0.1)');
			assert.strictEqual(oParser.calculate().getValue(), 1.5, 'ECMA.CEILING(1.5, 0.1)');

			oParser = new parserFormula('ECMA.CEILING(0.234, 0.01)', "A1", ws);
			assert.ok(oParser.parse(), 'ECMA.CEILING(0.234, 0.01)');
			assert.strictEqual(oParser.calculate().getValue(), 0.24, 'ECMA.CEILING(0.234, 0.01)');

		});

		QUnit.test("Test: \"COMBINA\"", function (assert) {
			oParser = new parserFormula('COMBINA(4,3)', "A1", ws);
			assert.ok(oParser.parse(), 'COMBINA(4,3)');
			assert.strictEqual(oParser.calculate().getValue(), 20, 'COMBINA(4,3)');

			oParser = new parserFormula('COMBINA(10,3)', "A1", ws);
			assert.ok(oParser.parse(), 'COMBINA(10,3)');
			assert.strictEqual(oParser.calculate().getValue(), 220, 'COMBINA(10,3)');

			oParser = new parserFormula('COMBINA(3,10)', "A1", ws);
			assert.ok(oParser.parse(), 'COMBINA(3,10)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'COMBINA(10,3)');

			oParser = new parserFormula('COMBINA(10,-3)', "A1", ws);
			assert.ok(oParser.parse(), 'COMBINA(10,-3)');
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'COMBINA(10,-3)');

			testArrayFormula2(assert, "COMBINA", 2, 2);
		});

		QUnit.test("Test: \"DECIMAL\"", function (assert) {
			oParser = new parserFormula('DECIMAL("FF",16)', "A1", ws);
			assert.ok(oParser.parse(), 'DECIMAL("FF",16)');
			assert.strictEqual(oParser.calculate().getValue(), 255, 'DECIMAL("FF",16)');

			oParser = new parserFormula('DECIMAL(111,2)', "A1", ws);
			assert.ok(oParser.parse(), 'DECIMAL(111,2)');
			assert.strictEqual(oParser.calculate().getValue(), 7, 'DECIMAL(111,2)');

			oParser = new parserFormula('DECIMAL("zap",36)', "A1", ws);
			assert.ok(oParser.parse(), 'DECIMAL("zap",36)');
			assert.strictEqual(oParser.calculate().getValue(), 45745, 'DECIMAL("zap",36)');

			oParser = new parserFormula('DECIMAL("00FF",16)', "A1", ws);
			assert.ok(oParser.parse(), 'DECIMAL("00FF",16)');
			assert.strictEqual(oParser.calculate().getValue(), 255, 'DECIMAL("00FF",16)');

			oParser = new parserFormula('DECIMAL("101b",2)', "A1", ws);
			assert.ok(oParser.parse(), 'DECIMAL("101b",2)');
			assert.strictEqual(oParser.calculate().getValue(), 5, 'DECIMAL("101b",2)');

			testArrayFormula2(assert, "DECIMAL", 2, 2);
		});

		QUnit.test("Test: \"BASE\"", function (assert) {
			oParser = new parserFormula('BASE(7,2)', "A1", ws);
			assert.ok(oParser.parse(), 'BASE(7,2)');
			assert.strictEqual(oParser.calculate().getValue(), "111", 'BASE(7,2)');

			oParser = new parserFormula('BASE(100,16)', "A1", ws);
			assert.ok(oParser.parse(), 'BASE(100,16)');
			assert.strictEqual(oParser.calculate().getValue(), "64", 'BASE(100,16)');

			oParser = new parserFormula('BASE(15,2,10)', "A1", ws);
			assert.ok(oParser.parse(), 'BASE(15,2,10)');
			assert.strictEqual(oParser.calculate().getValue(), "0000001111", 'BASE(15,2,10)');

			testArrayFormula2(assert, "BASE", 2, 3);
		});

		QUnit.test("Test: \"ARABIC('LVII')\"", function (assert) {
			oParser = new parserFormula('ARABIC("LVII")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 57);
		});

		QUnit.test("Test: \"TDIST\"", function (assert) {
			oParser = new parserFormula("TDIST(60,1,2)", "A1", ws);
			assert.ok(oParser.parse(), "TDIST(60,1,2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.010609347, "TDIST(60,1,2)");

			oParser = new parserFormula("TDIST(8,3,1)", "A1", ws);
			assert.ok(oParser.parse(), "TDIST(8,3,1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.002038289, "TDIST(8,3,1)");

			ws.getRange2("A2").setValue("1.959999998");
			ws.getRange2("A3").setValue("60");

			oParser = new parserFormula("TDIST(A2,A3,2)", "A1", ws);
			assert.ok(oParser.parse(), "TDIST(A2,A3,2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.054644930, "TDIST(A2,A3,2)");

			oParser = new parserFormula("TDIST(A2,A3,1)", "A1", ws);
			assert.ok(oParser.parse(), "TDIST(A2,A3,1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.027322465, "TDIST(A2,A3,1)");

			testArrayFormula2(assert, "TDIST", 3, 3);
		});

		QUnit.test("Test: \"T.DIST\"", function (assert) {
			oParser = new parserFormula("T.DIST(60,1,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "T.DIST(60,1,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.99469533, "T.DIST(60,1,TRUE)");

			oParser = new parserFormula("T.DIST(8,3,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "T.DIST(8,3,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.00073691, "T.DIST(8,3,FALSE)");

			testArrayFormula2(assert, "T.DIST", 3, 3);
		});

		QUnit.test("Test: \"T.DIST.2T\"", function (assert) {
			ws.getRange2("A2").setValue("1.959999998");
			ws.getRange2("A3").setValue("60");

			oParser = new parserFormula("T.DIST.2T(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "T.DIST.2T(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.054644930, "T.DIST.2T(A2,A3)");

			testArrayFormula2(assert, "T.DIST.2T", 2, 2);
		});

		QUnit.test("Test: \"T.DIST.RT\"", function (assert) {
			ws.getRange2("A2").setValue("1.959999998");
			ws.getRange2("A3").setValue("60");

			oParser = new parserFormula("T.DIST.RT(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "T.DIST.RT(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.027322, "T.DIST.RT(A2,A3)");

			testArrayFormula2(assert, "T.DIST.RT", 2, 2);
		});

		QUnit.test("Test: \"TTEST\"", function (assert) {
			ws.getRange2("A2").setValue("3");
			ws.getRange2("A3").setValue("4");
			ws.getRange2("A4").setValue("5");
			ws.getRange2("A5").setValue("8");
			ws.getRange2("A6").setValue("9");
			ws.getRange2("A7").setValue("1");
			ws.getRange2("A8").setValue("2");
			ws.getRange2("A9").setValue("4");
			ws.getRange2("A10").setValue("5");

			ws.getRange2("B2").setValue("6");
			ws.getRange2("B3").setValue("19");
			ws.getRange2("B4").setValue("3");
			ws.getRange2("B5").setValue("2");
			ws.getRange2("B6").setValue("14");
			ws.getRange2("B7").setValue("4");
			ws.getRange2("B8").setValue("5");
			ws.getRange2("B9").setValue("17");
			ws.getRange2("B10").setValue("1");

			oParser = new parserFormula("TTEST(A2:A10,B2:B10,2,1)", "A1", ws);
			assert.ok(oParser.parse(), "TTEST(A2:A10,B2:B10,2,1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.196016, "TTEST(A2:A10,B2:B10,2,1)");

			//TODO нужна другая функция для тестирования
			//testArrayFormula2(assert, "TTEST", 4, 4, null, true);
		});

		QUnit.test("Test: \"T.TEST\"", function (assert) {
			ws.getRange2("A2").setValue("3");
			ws.getRange2("A3").setValue("4");
			ws.getRange2("A4").setValue("5");
			ws.getRange2("A5").setValue("8");
			ws.getRange2("A6").setValue("9");
			ws.getRange2("A7").setValue("1");
			ws.getRange2("A8").setValue("2");
			ws.getRange2("A9").setValue("4");
			ws.getRange2("A10").setValue("5");

			ws.getRange2("B2").setValue("6");
			ws.getRange2("B3").setValue("19");
			ws.getRange2("B4").setValue("3");
			ws.getRange2("B5").setValue("2");
			ws.getRange2("B6").setValue("14");
			ws.getRange2("B7").setValue("4");
			ws.getRange2("B8").setValue("5");
			ws.getRange2("B9").setValue("17");
			ws.getRange2("B10").setValue("1");

			oParser = new parserFormula("T.TEST(A2:A10,B2:B10,2,1)", "A1", ws);
			assert.ok(oParser.parse(), "T.TEST(A2:A10,B2:B10,2,1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(5) - 0, 0.19602, "T.TEST(A2:A10,B2:B10,2,1)");
		});

		QUnit.test("Test: \"ZTEST\"", function (assert) {
			ws.getRange2("A2").setValue("3");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("7");
			ws.getRange2("A5").setValue("8");
			ws.getRange2("A6").setValue("6");
			ws.getRange2("A7").setValue("5");
			ws.getRange2("A8").setValue("4");
			ws.getRange2("A9").setValue("2");
			ws.getRange2("A10").setValue("1");
			ws.getRange2("A11").setValue("9");

			oParser = new parserFormula("ZTEST(A2:A11,4)", "A1", ws);
			assert.ok(oParser.parse(), "ZTEST(A2:A11,4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.090574, "ZTEST(A2:A11,4)");

			oParser = new parserFormula("2 * MIN(ZTEST(A2:A11,4), 1 - ZTEST(A2:A11,4))", "A1", ws);
			assert.ok(oParser.parse(), "2 * MIN(ZTEST(A2:A11,4), 1 - ZTEST(A2:A11,4))");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.181148, "2 * MIN(ZTEST(A2:A11,4), 1 - ZTEST(A2:A11,4))");

			oParser = new parserFormula("ZTEST(A2:A11,6)", "A1", ws);
			assert.ok(oParser.parse(), "ZTEST(A2:A11,6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.863043, "ZTEST(A2:A11,6)");

			oParser = new parserFormula("2 * MIN(ZTEST(A2:A11,6), 1 - ZTEST(A2:A11,6))", "A1", ws);
			assert.ok(oParser.parse(), "2 * MIN(ZTEST(A2:A11,6), 1 - ZTEST(A2:A11,6))");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.273913, "2 * MIN(ZTEST(A2:A11,6), 1 - ZTEST(A2:A11,6))");

			//TODO нужна другая функция для тестирования
			//testArrayFormula2(assert, "Z.TEST", 2, 3, null, true);
		});

		QUnit.test("Test: \"Z.TEST\"", function (assert) {
			ws.getRange2("A2").setValue("3");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("7");
			ws.getRange2("A5").setValue("8");
			ws.getRange2("A6").setValue("6");
			ws.getRange2("A7").setValue("5");
			ws.getRange2("A8").setValue("4");
			ws.getRange2("A9").setValue("2");
			ws.getRange2("A10").setValue("1");
			ws.getRange2("A11").setValue("9");

			oParser = new parserFormula("Z.TEST(A2:A11,4)", "A1", ws);
			assert.ok(oParser.parse(), "Z.TEST(A2:A11,4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.090574, "Z.TEST(A2:A11,4)");

			oParser = new parserFormula("2 * MIN(Z.TEST(A2:A11,4), 1 - Z.TEST(A2:A11,4))", "A1", ws);
			assert.ok(oParser.parse(), "2 * MIN(Z.TEST(A2:A11,4), 1 - Z.TEST(A2:A11,4))");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.181148, "2 * MIN(Z.TEST(A2:A11,4), 1 - Z.TEST(A2:A11,4))");

			oParser = new parserFormula("Z.TEST(A2:A11,6)", "A1", ws);
			assert.ok(oParser.parse(), "Z.TEST(A2:A11,6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.863043, "Z.TEST(A2:A11,6)");

			oParser = new parserFormula("2 * MIN(Z.TEST(A2:A11,6), 1 - Z.TEST(A2:A11,6))", "A1", ws);
			assert.ok(oParser.parse(), "2 * MIN(Z.TEST(A2:A11,6), 1 - Z.TEST(A2:A11,6))");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.273913, "2 * MIN(Z.TEST(A2:A11,6), 1 - Z.TEST(A2:A11,6))");

			//TODO нужна другая функция для тестирования
			//testArrayFormula2(assert, "Z.TEST", 2, 3, null, true);
		});


		QUnit.test("Test: \"F.DIST\"", function (assert) {
			ws.getRange2("A2").setValue("15.2069");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("4");

			oParser = new parserFormula("F.DIST(A2,A3,A4,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "F.DIST(A2,A3,A4,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.99, "F.DIST(A2,A3,A4,TRUE)");

			oParser = new parserFormula("F.DIST(A2,A3,A4,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "F.DIST(A2,A3,A4,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0012238, "F.DIST(A2,A3,A4,FALSE)");

			testArrayFormula2(assert, "F.DIST", 4, 4);
		});

		QUnit.test("Test: \"F.DIST.RT\"", function (assert) {
			ws.getRange2("A2").setValue("15.2069");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("4");

			oParser = new parserFormula("F.DIST.RT(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "F.DIST.RT(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.01, "F.DIST.RT(A2,A3,A4)");

			testArrayFormula2(assert, "F.DIST.RT", 3, 3);
		});

		QUnit.test("Test: \"FDIST\"", function (assert) {
			ws.getRange2("A2").setValue("15.2069");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("4");

			oParser = new parserFormula("FDIST(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "FDIST(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.01, "FDIST(A2,A3,A4)");
		});

		QUnit.test("Test: \"FINV\"", function (assert) {
			ws.getRange2("A2").setValue("0.01");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("4");

			oParser = new parserFormula("FINV(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "FINV(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 15.206865, "FINV(A2,A3,A4)");

			testArrayFormula2(assert, "FINV", 3, 3);
		});

		QUnit.test("Test: \"F.INV\"", function (assert) {
			ws.getRange2("A2").setValue("0.01");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("4");

			oParser = new parserFormula("F.INV(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "F.INV(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.10930991, "F.INV(A2,A3,A4)");

			testArrayFormula2(assert, "F.INV", 3, 3);
		});

		QUnit.test("Test: \"F.INV.RT\"", function (assert) {
			ws.getRange2("A2").setValue("0.01");
			ws.getRange2("A3").setValue("6");
			ws.getRange2("A4").setValue("4");

			oParser = new parserFormula("F.INV.RT(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "F.INV.RT(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(5) - 0, 15.20686, "F.INV.RT(A2,A3,A4)");
		});

		function fTestFormulaTest(assert) {
			ws.getRange2("A2").setValue("6");
			ws.getRange2("A3").setValue("7");
			ws.getRange2("A4").setValue("9");
			ws.getRange2("A5").setValue("15");
			ws.getRange2("A6").setValue("21");

			ws.getRange2("B2").setValue("20");
			ws.getRange2("B3").setValue("28");
			ws.getRange2("B4").setValue("31");
			ws.getRange2("B5").setValue("38");
			ws.getRange2("B6").setValue("40");

			oParser = new parserFormula("FTEST(A2:A6,B2:B6)", "A1", ws);
			assert.ok(oParser.parse(), "FTEST(A2:A6,B2:B6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.64831785, "FTEST(A2:A6,B2:B6)");

			oParser = new parserFormula("FTEST(A2,B2:B6)", "A1", ws);
			assert.ok(oParser.parse(), "FTEST(A2,B2:B6)");
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "FTEST(A2,B2:B6)");

			oParser = new parserFormula("FTEST(1,B2:B6)", "A1", ws);
			assert.ok(oParser.parse(), "FTEST(1,B2:B6)");
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "FTEST(1,B2:B6)");

			oParser = new parserFormula("FTEST({1,2,3},{2,3,4,5})", "A1", ws);
			assert.ok(oParser.parse(), "FTEST({1,2,3},{2,3,4,5})");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.792636779, "FTEST({1,2,3},{2,3,4,5})");

			oParser = new parserFormula("FTEST({1,\"test\",\"test\"},{2,3,4,5})", "A1", ws);
			assert.ok(oParser.parse(), "FTEST({1,\"test\",\"test\"},{2,3,4,5})");
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "FTEST({1,\"test\",\"test\"},{2,3,4,5})");
		}

		QUnit.test("Test: \"FTEST\"", function (assert) {
			fTestFormulaTest(assert);
			testArrayFormula2(assert, "FTEST", 2, 2, null, true);
		});

		QUnit.test("Test: \"F.TEST\"", function (assert) {
			fTestFormulaTest(assert);
			testArrayFormula2(assert, "F.TEST", 2, 2, null, true);
		});

		QUnit.test("Test: \"T.INV\"", function (assert) {
			oParser = new parserFormula("T.INV(0.75,2)", "A1", ws);
			assert.ok(oParser.parse(), "T.INV(0.75,2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.8164966, "T.INV(0.75,2)");

			testArrayFormula2(assert, "T.INV", 2, 2);
		});

		QUnit.test("Test: \"T.INV.2T\"", function (assert) {
			ws.getRange2("A2").setValue("0.546449");
			ws.getRange2("A3").setValue("60");

			oParser = new parserFormula("T.INV.2T(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "T.INV.2T(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.606533, "T.INV.2T(A2,A3)");

			testArrayFormula2(assert, "T.INV.2T", 2, 2);
		});

		QUnit.test("Test: \"RANK\"", function (assert) {
			ws.getRange2("A2").setValue("7");
			ws.getRange2("A3").setValue("3.5");
			ws.getRange2("A4").setValue("3.5");
			ws.getRange2("A5").setValue("1");
			ws.getRange2("A6").setValue("2");

			oParser = new parserFormula("RANK(A3,A2:A6,1)", "A1", ws);
			assert.ok(oParser.parse(), "RANK(A3,A2:A6,1)");
			assert.strictEqual(oParser.calculate().getValue(), 3, "RANK(A3,A2:A6,1)");

			oParser = new parserFormula("RANK(A2,A2:A6,1)", "A1", ws);
			assert.ok(oParser.parse(), "RANK(A2,A2:A6,1)");
			assert.strictEqual(oParser.calculate().getValue(), 5, "RANK(A2,A2:A6,1)");
		});

		QUnit.test("Test: \"RANK.EQ\"", function (assert) {
			ws.getRange2("A2").setValue("7");
			ws.getRange2("A3").setValue("3.5");
			ws.getRange2("A4").setValue("3.5");
			ws.getRange2("A5").setValue("1");
			ws.getRange2("A6").setValue("2");

			oParser = new parserFormula("RANK.EQ(A2,A2:A6,1)", "A1", ws);
			assert.ok(oParser.parse(), "RANK.EQ(A2,A2:A6,1)");
			assert.strictEqual(oParser.calculate().getValue(), 5, "RANK.EQ(A2,A2:A6,1)");

			oParser = new parserFormula("RANK.EQ(A6,A2:A6)", "A1", ws);
			assert.ok(oParser.parse(), "RANK.EQ(A6,A2:A6)");
			assert.strictEqual(oParser.calculate().getValue(), 4, "RANK.EQ(A6,A2:A6)");

			oParser = new parserFormula("RANK.EQ(A3,A2:A6,1)", "A1", ws);
			assert.ok(oParser.parse(), "RANK.EQ(A3,A2:A6,1)");
			assert.strictEqual(oParser.calculate().getValue(), 3, "RANK.EQ(A3,A2:A6,1)");
		});

		QUnit.test("Test: \"RANK.AVG\"", function (assert) {
			ws.getRange2("A2").setValue("89");
			ws.getRange2("A3").setValue("88");
			ws.getRange2("A4").setValue("92");
			ws.getRange2("A5").setValue("101");
			ws.getRange2("A6").setValue("94");
			ws.getRange2("A7").setValue("97");
			ws.getRange2("A8").setValue("95");

			oParser = new parserFormula("RANK.AVG(94,A2:A8)", "A1", ws);
			assert.ok(oParser.parse(), "RANK.AVG(94,A2:A8)");
			assert.strictEqual(oParser.calculate().getValue(), 4, "RANK.AVG(94,A2:A8)");
		});

		QUnit.test("Test: \"RADIANS\"", function (assert) {
			oParser = new parserFormula("RADIANS(270)", "A1", ws);
			assert.ok(oParser.parse(), "RADIANS(270)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 4.712389);

			testArrayFormula(assert, "RADIANS");
		});

		QUnit.test("Test: \"LOG\"", function (assert) {
			oParser = new parserFormula("LOG(10)", "A1", ws);
			assert.ok(oParser.parse(), "LOG(10)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "LOG(10)");

			oParser = new parserFormula("LOG(8,2)", "A1", ws);
			assert.ok(oParser.parse(), "LOG(8,2)");
			assert.strictEqual(oParser.calculate().getValue(), 3, "LOG(8,2)");

			oParser = new parserFormula("LOG(86, 2.7182818)", "A1", ws);
			assert.ok(oParser.parse(), "LOG(86, 2.7182818)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 4.4543473, "LOG(86, 2.7182818)");

			oParser = new parserFormula("LOG(8,1)", "A1", ws);
			assert.ok(oParser.parse(), "LOG(8,1)");
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "LOG(8,1)");

			testArrayFormula(assert, "LOG", 1, 2);
		});

		QUnit.test("Test: \"LOGNORM.DIST\"", function (assert) {
			ws.getRange2("A2").setValue("4");
			ws.getRange2("A3").setValue("3.5");
			ws.getRange2("A4").setValue("1.2");

			oParser = new parserFormula("LOGNORM.DIST(A2,A3,A4,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "LOGNORM.DIST(A2,A3,A4,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0390836, "LOGNORM.DIST(A2,A3,A4,TRUE)");

			oParser = new parserFormula("LOGNORM.DIST(A2,A3,A4,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "LOGNORM.DIST(A2,A3,A4,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0176176, "LOGNORM.DIST(A2,A3,A4,FALSE)");

			testArrayFormula2(assert, "LOGNORM.DIST", 4, 4);
		});

		QUnit.test("Test: \"LOGNORM.INV\"", function (assert) {
			ws.getRange2("A2").setValue("0.039084");
			ws.getRange2("A3").setValue("3.5");
			ws.getRange2("A4").setValue("1.2");

			oParser = new parserFormula("LOGNORM.INV(A2, A3, A4)", "A1", ws);
			assert.ok(oParser.parse(), "LOGNORM.INV(A2, A3, A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 4.0000252, "LOGNORM.INV(A2, A3, A4)");

			testArrayFormula2(assert, "LOGNORM.INV", 3, 3);
		});

		QUnit.test("Test: \"LOGNORMDIST\"", function (assert) {
			ws.getRange2("A2").setValue("4");
			ws.getRange2("A3").setValue("3.5");
			ws.getRange2("A4").setValue("1.2");

			oParser = new parserFormula("LOGNORMDIST(A2, A3, A4)", "A1", ws);
			assert.ok(oParser.parse(), "LOGNORMDIST(A2, A3, A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0390836, "LOGNORMDIST(A2, A3, A4)");

			testArrayFormula2(assert, "LOGNORMDIST", 3, 3);
		});

		QUnit.test("Test: \"LOWER\"", function (assert) {
			ws.getRange2("A2").setValue("E. E. Cummings");
			ws.getRange2("A3").setValue("Apt. 2B");

			oParser = new parserFormula("LOWER(A2)", "A1", ws);
			assert.ok(oParser.parse(), "LOWER(A2)");
			assert.strictEqual(oParser.calculate().getValue(), "e. e. cummings", "LOWER(A2)");

			oParser = new parserFormula("LOWER(A3)", "A1", ws);
			assert.ok(oParser.parse(), "LOWER(A3)");
			assert.strictEqual(oParser.calculate().getValue(), "apt. 2b", "LOWER(A3)");

			testArrayFormula2(assert, "LOWER", 1, 1);
		});

		QUnit.test("Test: \"EXPON.DIST\"", function (assert) {
			ws.getRange2("A2").setValue("0.2");
			ws.getRange2("A3").setValue("10");

			oParser = new parserFormula("EXPON.DIST(A2,A3,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "EXPON.DIST(A2,A3,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.86466472, "EXPON.DIST(A2,A3,TRUE)");

			oParser = new parserFormula("EXPON.DIST(0.2,10,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "EXPON.DIST(0.2,10,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 1.35335283, "EXPON.DIST(0.2,10,FALSE)");

			testArrayFormula2(assert, "EXPON.DIST", 3, 3);
		});

		QUnit.test("Test: \"GAMMA.DIST\"", function (assert) {
			ws.getRange2("A2").setValue("10.00001131");
			ws.getRange2("A3").setValue("9");
			ws.getRange2("A4").setValue("2");

			oParser = new parserFormula("GAMMA.DIST(A2,A3,A4,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMA.DIST(A2,A3,A4,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.032639, "GAMMA.DIST(A2,A3,A4,FALSE)");

			oParser = new parserFormula("GAMMA.DIST(A2,A3,A4,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMA.DIST(A2,A3,A4,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.068094, "GAMMA.DIST(A2,A3,A4,TRUE)");

			testArrayFormula2(assert, "GAMMA.DIST", 4, 4);
		});

		QUnit.test("Test: \"GAMMADIST\"", function (assert) {
			ws.getRange2("A2").setValue("10.00001131");
			ws.getRange2("A3").setValue("9");
			ws.getRange2("A4").setValue("2");

			oParser = new parserFormula("GAMMADIST(A2,A3,A4,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMADIST(A2,A3,A4,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.032639, "GAMMADIST(A2,A3,A4,FALSE)");

			oParser = new parserFormula("GAMMADIST(A2,A3,A4,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMADIST(A2,A3,A4,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.068094, "GAMMADIST(A2,A3,A4,TRUE)");
		});

		QUnit.test("Test: \"GAMMADIST\"", function (assert) {

			oParser = new parserFormula("GAMMADIST(A2,A3,A4,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMADIST(A2,A3,A4,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.032639, "GAMMADIST(A2,A3,A4,FALSE)");

			oParser = new parserFormula("GAMMADIST(A2,A3,A4,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMADIST(A2,A3,A4,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.068094, "GAMMADIST(A2,A3,A4,TRUE)");
		});


		QUnit.test("Test: \"GAMMA\"", function (assert) {

			oParser = new parserFormula("GAMMA(2.5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(3), "1.329");

			oParser = new parserFormula("GAMMA(-3.75)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(3), "0.268");

			oParser = new parserFormula("GAMMA(0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula("GAMMA(-2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");


			testArrayFormula2(assert, "GAMMA", 1, 1);
		});

		QUnit.test("Test: \"CHITEST\"", function (assert) {
			ws.getRange2("A2").setValue("58");
			ws.getRange2("A3").setValue("11");
			ws.getRange2("A4").setValue("10");
			ws.getRange2("A5").setValue("x");
			ws.getRange2("A6").setValue("45.35");
			ws.getRange2("A7").setValue("17.56");
			ws.getRange2("A8").setValue("16.09");

			ws.getRange2("B2").setValue("35");
			ws.getRange2("B3").setValue("25");
			ws.getRange2("B4").setValue("23");
			ws.getRange2("B5").setValue("x");
			ws.getRange2("B6").setValue("47.65");
			ws.getRange2("B7").setValue("18.44");
			ws.getRange2("B8").setValue("16.91");

			oParser = new parserFormula("CHITEST(A2:B4,A6:B8)", "A1", ws);
			assert.ok(oParser.parse(), "CHITEST(A2:B4,A6:B8)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0003082, "CHITEST(A2:B4,A6:B8)");

			testArrayFormula2(assert, "CHITEST", 2, 2, null, true);
		});

		QUnit.test("Test: \"CHISQ.TEST\"", function (assert) {
			ws.getRange2("A2").setValue("58");
			ws.getRange2("A3").setValue("11");
			ws.getRange2("A4").setValue("10");
			ws.getRange2("A5").setValue("x");
			ws.getRange2("A6").setValue("45.35");
			ws.getRange2("A7").setValue("17.56");
			ws.getRange2("A8").setValue("16.09");

			ws.getRange2("B2").setValue("35");
			ws.getRange2("B3").setValue("25");
			ws.getRange2("B4").setValue("23");
			ws.getRange2("B5").setValue("x");
			ws.getRange2("B6").setValue("47.65");
			ws.getRange2("B7").setValue("18.44");
			ws.getRange2("B8").setValue("16.91");

			oParser = new parserFormula("CHISQ.TEST(A2:B4,A6:B8)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.TEST(A2:B4,A6:B8)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0003082, "CHISQ.TEST(A2:B4,A6:B8)");
		});

		QUnit.test("Test: \"CHITEST\"", function (assert) {
			ws.getRange2("A2").setValue("18.307");
			ws.getRange2("A3").setValue("10");

			oParser = new parserFormula("CHIDIST(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "CHIDIST(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0500006, "CHIDIST(A2,A3)");

			testArrayFormula2(assert, "CHIDIST", 2, 2);
		});

		QUnit.test("Test: \"GAUSS\"", function (assert) {
			oParser = new parserFormula("GAUSS(2)", "A1", ws);
			assert.ok(oParser.parse(), "GAUSS(2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(5) - 0, 0.47725, "GAUSS(2)");

			testArrayFormula2(assert, "GAUSS", 1, 1);
		});

		QUnit.test("Test: \"CHISQ.DIST.RT\"", function (assert) {
			ws.getRange2("A2").setValue("18.307");
			ws.getRange2("A3").setValue("10");

			oParser = new parserFormula("CHISQ.DIST.RT(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.DIST.RT(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0500006, "CHISQ.DIST.RT(A2,A3)");

			testArrayFormula2(assert, "CHISQ.INV.RT", 2, 2);
		});

		QUnit.test("Test: \"CHISQ.INV\"", function (assert) {
			oParser = new parserFormula("CHISQ.INV(0.93,1)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.INV(0.93,1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 3.283020287, "CHISQ.INV(0.93,1)");

			oParser = new parserFormula("CHISQ.INV(0.6,2)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.INV(0.6,2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 1.832581464, "CHISQ.INV(0.6,2)");

			testArrayFormula2(assert, "CHISQ.INV", 2, 2);
		});

		QUnit.test("Test: \"CHISQ.DIST\"", function (assert) {
			oParser = new parserFormula("CHISQ.DIST(0.5,1,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.DIST(0.5,1,TRUE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.52049988, "CHISQ.DIST(0.5,1,TRUE)");

			oParser = new parserFormula("CHISQ.DIST(2,3,FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.DIST(2,3,FALSE)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 0.20755375, "CHISQ.DIST(2,3,FALSE)");

			testArrayFormula2(assert, "CHISQ.DIST", 3, 3);
		});

		QUnit.test("Test: \"CHIINV\"", function (assert) {
			ws.getRange2("A2").setValue("0.050001");
			ws.getRange2("A3").setValue("10");

			oParser = new parserFormula("CHIINV(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "CHIINV(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 18.306973, "CHIINV(A2,A3)");

			testArrayFormula2(assert, "CHIINV", 2, 2);
		});

		QUnit.test("Test: \"CHISQ.INV.RT\"", function (assert) {
			ws.getRange2("A2").setValue("0.050001");
			ws.getRange2("A3").setValue("10");

			oParser = new parserFormula("CHISQ.INV.RT(A2,A3)", "A1", ws);
			assert.ok(oParser.parse(), "CHISQ.INV.RT(A2,A3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 18.306973, "CHISQ.INV.RT(A2,A3)");

			testArrayFormula2(assert, "CHISQ.INV.RT", 2, 2);
		});

		QUnit.test("Test: \"CHOOSE\"", function (assert) {
			ws.getRange2("A2").setValue("st");
			ws.getRange2("A3").setValue("2nd");
			ws.getRange2("A4").setValue("3rd");
			ws.getRange2("A5").setValue("Finished");

			ws.getRange2("B2").setValue("Nails");
			ws.getRange2("B3").setValue("Screws");
			ws.getRange2("B4").setValue("Nuts");
			ws.getRange2("B5").setValue("Bolts");

			oParser = new parserFormula("CHOOSE(2,A2,A3,A4,A5)", "A1", ws);
			assert.ok(oParser.parse(), "CHOOSE(2,A2,A3,A4,A5)");
			assert.strictEqual(oParser.calculate().getValue().getValue(), "2nd", "CHOOSE(2,A2,A3,A4,A5)");

			oParser = new parserFormula("CHOOSE(4,B2,B3,B4,B5)", "A1", ws);
			assert.ok(oParser.parse(), "CHOOSE(4,B2,B3,B4,B5)");
			assert.strictEqual(oParser.calculate().getValue().getValue(), "Bolts", "CHOOSE(4,B2,B3,B4,B5))");

			oParser = new parserFormula('CHOOSE(3,"Wide",115,"world",8)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "world");

			//функция возвращает ref
			//testArrayFormula2(assert, "CHOOSE", 2, 9);
		});

		QUnit.test("Test: \"CHOOSECOLS\"", function (assert) {
			//1. добавляем общие тесты

			ws.getRange2("A1").setValue("1");
			ws.getRange2("A2").setValue("2");
			ws.getRange2("A3").setValue("");
			ws.getRange2("A4").setValue("4");
			ws.getRange2("A5").setValue("#N/A");
			ws.getRange2("A6").setValue("f");

			ws.getRange2("B1").setValue("q");
			ws.getRange2("B2").setValue("w");
			ws.getRange2("B3").setValue("e");
			ws.getRange2("B4").setValue("test");
			ws.getRange2("B5").setValue("hhh");
			ws.getRange2("B6").setValue("g");

			ws.getRange2("C1").setValue("r");
			ws.getRange2("C2").setValue("3");
			ws.getRange2("C3").setValue("5");
			ws.getRange2("C4").setValue("");
			ws.getRange2("C5").setValue("6");
			ws.getRange2("C6").setValue("h");


			oParser = new parserFormula("CHOOSECOLS(A1:C6;-1;1)", "A1", ws);
			assert.ok(oParser.parse());
			let array = oParser.calculate();

			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'r');
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 5);
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), '');
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 'h');

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), '');
			assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(4, 1).getValue(), '#N/A');
			assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 'f');


			oParser = new parserFormula("CHOOSECOLS(A1:C6;-2;3)", "A1", ws);
			assert.ok(oParser.parse());
			array = oParser.calculate();

			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'q');
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'w');
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 'e');
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 'test');
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 'hhh');
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 'g');

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'r');
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 3);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 5);
			assert.strictEqual(array.getElementRowCol(3, 1).getValue(), '');
			assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 'h');


			oParser = new parserFormula("CHOOSECOLS(A1:C6;-4;3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("CHOOSECOLS(A1:C6;-2;4)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("CHOOSECOLS(A1:C6;-2;0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");


			//2. аргументы - разные типы. нужно пербрать все аргументы
			//2.1 аргумент - number
			oParser = new parserFormula("CHOOSECOLS(1,1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			//2.2 аргумент - string
			oParser = new parserFormula("CHOOSECOLS(\"test\",1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test");
			//2.3 аргумент - bool
			oParser = new parserFormula("CHOOSECOLS(true,1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "TRUE");
			//2.4 аргумент - error
			oParser = new parserFormula("CHOOSECOLS(#VALUE!,3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.5 аргумент - empty
			oParser = new parserFormula("CHOOSECOLS(,2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.6 аргумент - cellsRange
			//2.7 аргумент - cell
			oParser = new parserFormula("CHOOSECOLS(B1, 1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "q");

			//2.8 аргумент - array
			oParser = new parserFormula("CHOOSECOLS({2;\"\";\"test\"},3)", "A1", ws);
			assert.ok(oParser.parse());
			array = oParser.calculate();

			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			//2.8 аргумент - array
			oParser = new parserFormula("CHOOSECOLS({2,\"\",\"test\"},3)", "A1", ws);
			assert.ok(oParser.parse());
			array = oParser.calculate();

			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test");


			//2.2 аргумент - string
			oParser = new parserFormula("CHOOSECOLS(1,\"test\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.3 аргумент - bool
			oParser = new parserFormula("CHOOSECOLS(1,true)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			//2.4 аргумент - error
			oParser = new parserFormula("CHOOSECOLS(1, #VALUE!)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.5 аргумент - empty
			oParser = new parserFormula("CHOOSECOLS(1,)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");


			//2.6 аргумент - cellsRange
			//2.7 аргумент - cell
			oParser = new parserFormula("CHOOSECOLS(1,A1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);

			oParser = new parserFormula("CHOOSECOLS(1,A1:B5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			//2.8 аргумент - array
			oParser = new parserFormula("CHOOSECOLS(1,{2;\"\";\"test\"})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			ws.getRange2("F1").setValue("1");
			ws.getRange2("G1").setValue("3");

			oParser = new parserFormula("CHOOSECOLS(A1:C2,F1:G1,F1:G1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "r");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), 3);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 2).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 2).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 3).getValue(), "r");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 3).getValue(), 3);

			oParser = new parserFormula("CHOOSECOLS(A1:C2,F1:G1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "r");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), 3);

			oParser = new parserFormula("CHOOSECOLS(A1:C2,{1,2},{1,2},{1,2,3})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 2).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 2).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 3).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 3).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 4).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 4).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 5).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 5).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 6).getValue(), "r");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 6).getValue(), 3);

			oParser = new parserFormula("CHOOSECOLS(A1:C2,{1;2},{1;2},{1;2;3})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 2).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 2).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 3).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 3).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 4).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 4).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 5).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 5).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 6).getValue(), "r");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 6).getValue(), 3);

			oParser = new parserFormula("CHOOSECOLS(A1:C2,{1;2},{1,1;2,1},{1;2;3})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("CHOOSECOLS(A1:C2,{1;2},F1:G2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
		});

		QUnit.test("Test: \"CHOOSEROWS\"", function (assert) {
			//1. добавляем общие тесты

			ws.getRange2("A1").setValue("1");
			ws.getRange2("A2").setValue("2");
			ws.getRange2("A3").setValue("");
			ws.getRange2("A4").setValue("4");
			ws.getRange2("A5").setValue("#N/A");
			ws.getRange2("A6").setValue("f");

			ws.getRange2("B1").setValue("q");
			ws.getRange2("B2").setValue("w");
			ws.getRange2("B3").setValue("e");
			ws.getRange2("B4").setValue("test");
			ws.getRange2("B5").setValue("hhh");
			ws.getRange2("B6").setValue("g");

			ws.getRange2("C1").setValue("r");
			ws.getRange2("C2").setValue("3");
			ws.getRange2("C3").setValue("5");
			ws.getRange2("C4").setValue("");
			ws.getRange2("C5").setValue("6");
			ws.getRange2("C6").setValue("h");


			oParser = new parserFormula("CHOOSEROWS(A1:C6;-1;1)", "A1", ws);
			assert.ok(oParser.parse());
			let array = oParser.calculate();

			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'f');
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1);

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'g');
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 'q');

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 'h');
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 'r');



			oParser = new parserFormula("CHOOSEROWS(A1:C6;-2;3)", "A1", ws);
			assert.ok(oParser.parse());
			array = oParser.calculate();

			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), '#N/A');
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '');

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'hhh');
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 'e');

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 5);


			oParser = new parserFormula("CHOOSEROWS(A1:C6;-4;20)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("CHOOSEROWS(A1:C6;-10;4)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("CHOOSEROWS(A1:C6;-2;0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");


			//2. аргументы - разные типы. нужно пербрать все аргументы
			//2.1 аргумент - number
			oParser = new parserFormula("CHOOSEROWS(1,1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			//2.2 аргумент - string
			oParser = new parserFormula("CHOOSEROWS(\"test\",1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test");
			//2.3 аргумент - bool
			oParser = new parserFormula("CHOOSEROWS(true,1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "TRUE");
			//2.4 аргумент - error
			oParser = new parserFormula("CHOOSEROWS(#VALUE!,3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.5 аргумент - empty
			oParser = new parserFormula("CHOOSEROWS(,2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.6 аргумент - cellsRange
			//2.7 аргумент - cell
			oParser = new parserFormula("CHOOSEROWS(B1, 1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "q");

			//2.8 аргумент - array
			oParser = new parserFormula("CHOOSEROWS({2;\"\";\"test\"},3)", "A1", ws);
			assert.ok(oParser.parse());
			array = oParser.calculate();

			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test");

			//2.8 аргумент - array
			oParser = new parserFormula("CHOOSEROWS({2,\"\",\"test\"},3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");


			//2.2 аргумент - string
			oParser = new parserFormula("CHOOSEROWS(1,\"test\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.3 аргумент - bool
			oParser = new parserFormula("CHOOSEROWS(1,true)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			//2.4 аргумент - error
			oParser = new parserFormula("CHOOSEROWS(1, #VALUE!)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
			//2.5 аргумент - empty
			oParser = new parserFormula("CHOOSEROWS(1,)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");


			//2.6 аргумент - cellsRange
			//2.7 аргумент - cell
			oParser = new parserFormula("CHOOSEROWS(1,A1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);

			oParser = new parserFormula("CHOOSEROWS(1,A1:B5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			//2.8 аргумент - array
			oParser = new parserFormula("CHOOSEROWS(1,{2;\"\";\"test\"})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			ws.getRange2("F1").setValue("1");
			ws.getRange2("G1").setValue("3");

			oParser = new parserFormula("CHOOSEROWS(A1:C2,F1:G1,F1:G1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			ws.getRange2("G1").setValue("2");

			oParser = new parserFormula("CHOOSEROWS(A1:C2,F1:G1,F1:G1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 2).getValue(), "r");

			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 2).getValue(), 3);

			assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 2).getValue(), "r");

			assert.strictEqual(oParser.calculate().getElementRowCol(3, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(3, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(3, 2).getValue(), 3);

			oParser = new parserFormula("CHOOSEROWS(A1:C2,{1;2},{1;2})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 2).getValue(), "r");

			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 2).getValue(), 3);

			assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 2).getValue(), "r");

			assert.strictEqual(oParser.calculate().getElementRowCol(3, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(3, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(3, 2).getValue(), 3);

			oParser = new parserFormula("CHOOSEROWS(A1:C2,{1;2},{1;3})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("CHOOSEROWS(A1:C2,{1,2},{1,2})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 2).getValue(), "r");

			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 2).getValue(), 3);

			assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), 1);
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 1).getValue(), "q");
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 2).getValue(), "r");

			assert.strictEqual(oParser.calculate().getElementRowCol(3, 0).getValue(), 2);
			assert.strictEqual(oParser.calculate().getElementRowCol(3, 1).getValue(), "w");
			assert.strictEqual(oParser.calculate().getElementRowCol(3, 2).getValue(), 3);

		});

		QUnit.test("Test: \"BETA.INV\"", function (assert) {
			ws.getRange2("A2").setValue("0.685470581");
			ws.getRange2("A3").setValue("8");
			ws.getRange2("A4").setValue("10");
			ws.getRange2("A5").setValue("1");
			ws.getRange2("A6").setValue("3");

			oParser = new parserFormula("BETA.INV(A2,A3,A4,A5,A6)", "A1", ws);
			assert.ok(oParser.parse(), "BETA.INV(A2,A3,A4,A5,A6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(1) - 0, 2, "BETA.INV(A2,A3,A4,A5,A6)");

			testArrayFormula2(assert, "BETA.INV", 3, 5);
		});

		QUnit.test("Test: \"BETAINV\"", function (assert) {
			ws.getRange2("A2").setValue("0.685470581");
			ws.getRange2("A3").setValue("8");
			ws.getRange2("A4").setValue("10");
			ws.getRange2("A5").setValue("1");
			ws.getRange2("A6").setValue("3");

			oParser = new parserFormula("BETAINV(A2,A3,A4,A5,A6)", "A1", ws);
			assert.ok(oParser.parse(), "BETAINV(A2,A3,A4,A5,A6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(1) - 0, 2, "BETAINV(A2,A3,A4,A5,A6)");

			testArrayFormula2(assert, "BETAINV", 3, 5);
		});

		QUnit.test("Test: \"BETA.DIST\"", function (assert) {
			ws.getRange2("A2").setValue("2");
			ws.getRange2("A3").setValue("8");
			ws.getRange2("A4").setValue("10");
			ws.getRange2("A5").setValue("1");
			ws.getRange2("A6").setValue("3");

			oParser = new parserFormula("BETA.DIST(A2,A3,A4,TRUE,A5,A6)", "A1", ws);
			assert.ok(oParser.parse(), "BETA.DIST(A2,A3,A4,TRUE,A5,A6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.6854706, "BETA.DIST(A2,A3,A4,TRUE,A5,A6)");

			oParser = new parserFormula("BETA.DIST(A2,A3,A4,FALSE,A5,A6)", "A1", ws);
			assert.ok(oParser.parse(), "BETA.DIST(A2,A3,A4,FALSE,A5,A6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 1.4837646, "BETA.DIST(A2,A3,A4,FALSE,A5,A6)");

			testArrayFormula2(assert, "BETA.DIST", 4, 6);
		});

		QUnit.test("Test: \"BETADIST\"", function (assert) {
			ws.getRange2("A2").setValue("2");
			ws.getRange2("A3").setValue("8");
			ws.getRange2("A4").setValue("10");
			ws.getRange2("A5").setValue("1");
			ws.getRange2("A6").setValue("3");

			oParser = new parserFormula("BETADIST(A2,A3,A4,A5,A6)", "A1", ws);
			assert.ok(oParser.parse(), "BETADIST(A2,A3,A4,A5,A6)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.6854706, "BETADIST(A2,A3,A4,A5,A6)");

			oParser = new parserFormula("BETADIST(1,2,3,1,6)", "A1", ws);
			assert.ok(oParser.parse(), "BETADIST(1,2,3,1,6)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "BETADIST(1,2,3,1,6)");

			oParser = new parserFormula("BETADIST(6,2,3,1,6)", "A1", ws);
			assert.ok(oParser.parse(), "BETADIST(6,2,3,1,6)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "BETADIST(6,2,3,1,6)");

			testArrayFormula2(assert, "BETADIST", 3, 5);
		});

		QUnit.test("Test: \"BESSELJ\"", function (assert) {

			oParser = new parserFormula("BESSELJ(1.9, 2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELJ(1.9, 2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.329925728, "BESSELJ(1.9, 2)");

			oParser = new parserFormula("BESSELJ(1.9, 2.4)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELJ(1.9, 2.4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.329925728, "BESSELJ(1.9, 2.4)");

			oParser = new parserFormula("BESSELJ(-1.9, 2.4)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELJ(-1.9, 2.4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.329925728, "BESSELJ(-1.9, 2.4)");

			oParser = new parserFormula("BESSELJ(-1.9, -2.4)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELJ(-1.9, -2.4)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			testArrayFormula2(assert, "BESSELJ", 2, 2, true, null);
		});

		QUnit.test("Test: \"BESSELK\"", function (assert) {

			oParser = new parserFormula("BESSELK(1.5, 1)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELK(1.5, 1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.277387804, "BESSELK(1.5, 1)");

			oParser = new parserFormula("BESSELK(1, 3)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELK(1, 3)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(8) - 0, 7.10126281, "BESSELK(1, 3)");

			oParser = new parserFormula("BESSELK(-1.123,2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELK(-1.123,2)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula("BESSELK(1,-2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELK(1,-2)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			testArrayFormula2(assert, "BESSELK", 2, 2, true, null);

		});

		QUnit.test("Test: \"BESSELY\"", function (assert) {

			oParser = new parserFormula("BESSELY(2.5, 1)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELY(2.5, 1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.1459181, "BESSELY(2.5, 1)");

			oParser = new parserFormula("BESSELY(1,-2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELY(1,-2)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "BESSELY(1,-2)");

			oParser = new parserFormula("BESSELY(-1,2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELY(-1,2)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "BESSELY(-1,2)");

			testArrayFormula2(assert, "BESSELY", 2, 2, true, null);

		});

		QUnit.test("Test: \"BESSELI\"", function (assert) {
			//есть различия excel в некоторых формулах(неточности в 7 цифре после точки)
			oParser = new parserFormula("BESSELI(1.5, 1)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELI(1.5, 1)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.981666, "BESSELI(1.5, 1)");

			oParser = new parserFormula("BESSELI(1,2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELI(1,2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.135748, "BESSELI(1,2)");

			oParser = new parserFormula("BESSELI(1,-2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELI(1,-2)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "BESSELI(1,-2)");

			oParser = new parserFormula("BESSELI(-1,2)", "A1", ws);
			assert.ok(oParser.parse(), "BESSELI(-1,2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(6) - 0, 0.135748, "BESSELI(-1,2)");

			testArrayFormula2(assert, "BESSELI", 2, 2, true, null);
		});

		QUnit.test("Test: \"GAMMA.INV\"", function (assert) {
			ws.getRange2("A2").setValue("0.068094");
			ws.getRange2("A3").setValue("9");
			ws.getRange2("A4").setValue("2");

			oParser = new parserFormula("GAMMA.INV(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMA.INV(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 10.0000112, "GAMMA.INV(A2,A3,A4)");

			testArrayFormula2(assert, "GAMMA.INV", 3, 3);
		});

		QUnit.test("Test: \"GAMMAINV\"", function (assert) {
			ws.getRange2("A2").setValue("0.068094");
			ws.getRange2("A3").setValue("9");
			ws.getRange2("A4").setValue("2");

			oParser = new parserFormula("GAMMAINV(A2,A3,A4)", "A1", ws);
			assert.ok(oParser.parse(), "GAMMAINV(A2,A3,A4)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 10.0000112, "GAMMAINV(A2,A3,A4)");
		});

		QUnit.test("Test: \"SUM(1,2,3)\"", function (assert) {
			oParser = new parserFormula('SUM(1,2,3)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1 + 2 + 3);

			testArrayFormula2(assert, "SUM", 1, 8, null, true);
		});

		QUnit.test("Test: \"\"s\"&5\"", function (assert) {
			oParser = new parserFormula("\"s\"&5", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "s5");
		});

		QUnit.test("Test: \"String+Number\"", function (assert) {
			oParser = new parserFormula("1+\"099\"", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 100);

			ws.getRange2("A1469").setValue("'099");
			ws.getRange2("A1470").setValue("\"099\"");

			oParser = new parserFormula("1+A1469", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 100);


			oParser = new parserFormula("1+A1470", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		});

		QUnit.test("Test: \"POWER(2,8)\"", function (assert) {
			oParser = new parserFormula("POWER(2,8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), Math.pow(2, 8));
		});

		QUnit.test("Test: \"POWER(0,-3)\"", function (assert) {
			oParser = new parserFormula("POWER(0,-3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");

			testArrayFormula2(assert, "POWER", 2, 2);
		});

		QUnit.test("Test: \"ISNA(A1)\"", function (assert) {
			ws.getRange2("A1").setValue("#N/A");

			oParser = new parserFormula("ISNA(A1)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			testArrayFormula2(assert, "ISNA", 1, 1);
		});

		QUnit.test("Test: \"ISNONTEXT\"", function (assert) {
			oParser = new parserFormula('ISNONTEXT("123")', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE");

			testArrayFormula2(assert, "ISNONTEXT", 1, 1);
		});

		QUnit.test("Test: \"ISNUMBER\"", function (assert) {
			ws.getRange2("A1").setValue("123");

			oParser = new parserFormula('ISNUMBER(4)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula('ISNUMBER(A1)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			testArrayFormula2(assert, "ISNUMBER", 1, 1);
		});

		QUnit.test("Test: \"ISODD\"", function (assert) {
			oParser = new parserFormula('ISODD(-1)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula('ISODD(2.5)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE");

			oParser = new parserFormula('ISODD(5)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			testArrayFormula2(assert, "ISODD", 1, 1, true, null);
		});

		QUnit.test("Test: \"ROUND\"", function (assert) {
			oParser = new parserFormula("ROUND(2.15, 1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2.2);

			oParser = new parserFormula("ROUND(2.149, 1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2.1);

			oParser = new parserFormula("ROUND(-1.475, 2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -1.48);

			oParser = new parserFormula("ROUND(21.5, -1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 20);

			oParser = new parserFormula("ROUND(626.3,-3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1000);

			oParser = new parserFormula("ROUND(1.98,-1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("ROUND(-50.55,-2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -100);

			oParser = new parserFormula('ROUND("test",-2.1)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('ROUND(123.431,"test")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('ROUND(123.431,#NUM!)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula('ROUND(#NUM!,123.431)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula("ROUND(-50.55,-2.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -100);

			oParser = new parserFormula("ROUND(-50.55,-2.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -100);

			oParser = new parserFormula("ROUND(-50.55,0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -51);

			oParser = new parserFormula("ROUND(-50.55,0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -51);


			testArrayFormula2(assert, "ROUND", 2, 2);
		});

		QUnit.test("Test: \"ROUNDUP(31415.92654,-2)\"", function (assert) {
			oParser = new parserFormula("ROUNDUP(31415.92654,-2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 31500);
		});

		QUnit.test("Test: \"ROUNDUP(3.2,0)\"", function (assert) {
			oParser = new parserFormula("ROUNDUP(3.2,0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);
		});

		QUnit.test("Test: \"ROUNDUP(-3.14159,1)\"", function (assert) {
			oParser = new parserFormula("ROUNDUP(-3.14159,1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -3.2);
		});

		QUnit.test("Test: \"ROUNDUP(3.14159,3)\"", function (assert) {
			oParser = new parserFormula("ROUNDUP(3.14159,3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3.142);

			testArrayFormula2(assert, "ROUNDUP", 2, 2);
		});

		QUnit.test("Test: \"ROUNDUP\"", function (assert) {
			oParser = new parserFormula("ROUNDUP(2.1123,4)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue().toFixed(4) - 0, 2.1123);

			//TODO в хроме при расчёте разница, временно убираю
			oParser = new parserFormula("ROUNDUP(2,4)", "A1", ws);
			assert.ok(oParser.parse());
			//assert.strictEqual( oParser.calculate().getValue(), 2 );

			oParser = new parserFormula("ROUNDUP(2,0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);

			oParser = new parserFormula("ROUNDUP(2.1123,-1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 10);

			oParser = new parserFormula("ROUNDUP(2.1123,0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula("ROUNDUP(123.431,0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 124);

			oParser = new parserFormula("ROUNDUP(123.431,0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 124);

			oParser = new parserFormula("ROUNDUP(123.431,-0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 124);

			oParser = new parserFormula("ROUNDUP(123.431,-0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 124);

			oParser = new parserFormula("ROUNDUP(123.431,-2.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 200);

			oParser = new parserFormula('ROUNDUP("test",-2.1)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('ROUNDUP(123.431,"test")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('ROUNDUP(123.431,#NUM!)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula('ROUNDUP(#NUM!,123.431)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula('ROUNDUP(123.431,-1.9)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 130);

			oParser = new parserFormula("ROUNDUP(-50.55,0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -51);

			oParser = new parserFormula("ROUNDUP(-50.55,0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -51);

		});


		QUnit.test("Test: \"ROUNDDOWN(31415.92654,-2)\"", function (assert) {
			oParser = new parserFormula("ROUNDDOWN(31415.92654,-2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 31400);
		});

		QUnit.test("Test: \"ROUNDDOWN(-3.14159,1)\"", function (assert) {
			oParser = new parserFormula("ROUNDDOWN(-3.14159,1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -3.1);
		});

		QUnit.test("Test: \"ROUNDDOWN(3.14159,3)\"", function (assert) {
			oParser = new parserFormula("ROUNDDOWN(3.14159,3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3.141);
		});

		QUnit.test("Test: \"ROUNDDOWN(3.2,0)\"", function (assert) {
			oParser = new parserFormula("ROUNDDOWN(3.2,0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			testArrayFormula2(assert, "ROUNDDOWN", 2, 2);
		});

		QUnit.test("Test: \"ROUNDDOWN\"", function (assert) {
			oParser = new parserFormula("ROUNDDOWN(123.431,0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 123);

			oParser = new parserFormula("ROUNDDOWN(123.431,0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 123);

			oParser = new parserFormula("ROUNDDOWN(123.431,-0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 123);

			oParser = new parserFormula("ROUNDDOWN(123.431,-0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 123);

			oParser = new parserFormula("ROUNDDOWN(123.431,-2.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue() - 0, 100);

			oParser = new parserFormula('ROUNDDOWN("test",-2.1)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('ROUNDDOWN(123.431,"test")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula('ROUNDDOWN(123.431,#NUM!)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula('ROUNDDOWN(#NUM!,123.431)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			oParser = new parserFormula("ROUNDDOWN(-50.55,0.9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -50);

			oParser = new parserFormula("ROUNDDOWN(-50.55,0.1)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -50);
		});


		QUnit.test("Test: \"MROUND\"", function (assert) {
			var multiple;//должен равняться значению второго аргумента
			function mroundHelper(num) {
				var multiplier = Math.pow(10, Math.floor(Math.log(Math.abs(num)) / Math.log(10)) - AscCommonExcel.cExcelSignificantDigits + 1);
				var nolpiat = 0.5 * (num > 0 ? 1 : num < 0 ? -1 : 0) * multiplier;
				var y = (num + nolpiat) / multiplier;
				y = y / Math.abs(y) * Math.floor(Math.abs(y))
				var x = y * multiplier / multiple

				// var x = number / multiple;
				var nolpiat = 5 * (x / Math.abs(x)) * Math.pow(10, Math.floor(Math.log(Math.abs(x)) / Math.log(10)) - AscCommonExcel.cExcelSignificantDigits);
				x = x + nolpiat;
				x = x | x;

				return x * multiple;
			}


			oParser = new parserFormula("MROUND(10,3)", "A1", ws);
			assert.ok(oParser.parse());
			multiple = 3;
			assert.strictEqual(oParser.calculate().getValue(), mroundHelper(10 + 3 / 2));

			oParser = new parserFormula("MROUND(-10,-3)", "A1", ws);
			assert.ok(oParser.parse());
			multiple = -3;
			assert.strictEqual(oParser.calculate().getValue(), mroundHelper(-10 + -3 / 2));

			oParser = new parserFormula("MROUND(1.3,0.2)", "A1", ws);
			assert.ok(oParser.parse());
			multiple = 0.2;
			assert.strictEqual(oParser.calculate().getValue(), mroundHelper(1.3 + 0.2 / 2));

			testArrayFormula2(assert, "MROUND", 2, 2, true, null);
		});

		QUnit.test("Test: \"T(\"HELLO\")\"", function (assert) {
			oParser = new parserFormula("T(\"HELLO\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "HELLO");
		});

		QUnit.test("Test: \"MMULT\"", function (assert) {
			ws.getRange2("AAA102").setValue("4");
			ws.getRange2("AAA103").setValue("5");
			ws.getRange2("AAA104").setValue("6");
			ws.getRange2("AAA105").setValue("7");
			ws.getRange2("AAB102").setValue("1");
			ws.getRange2("AAB103").setValue("2");
			ws.getRange2("AAB104").setValue("3");
			ws.getRange2("AAB105").setValue("2");
			ws.getRange2("AAC102").setValue("4");
			ws.getRange2("AAC103").setValue("5");
			ws.getRange2("AAC104").setValue("6");
			ws.getRange2("AAC105").setValue("3");
			ws.getRange2("AAD102").setValue("7");
			ws.getRange2("AAD103").setValue("8");
			ws.getRange2("AAD104").setValue("9");
			ws.getRange2("AAD105").setValue("4");

			ws.getRange2("AAF102").setValue("1");
			ws.getRange2("AAF103").setValue("2");
			ws.getRange2("AAF104").setValue("3");
			ws.getRange2("AAF105").setValue("6");

			ws.getRange2("AAG102").setValue("2");
			ws.getRange2("AAG103").setValue("3");
			ws.getRange2("AAG104").setValue("4");
			ws.getRange2("AAG105").setValue("5");

			oParser = new parserFormula("MMULT(AAC102,AAF104)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 12);

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAF104)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("MMULT(AAC102,AAF104)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 12);

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAF102:AAG105)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 60);
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), 62);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 72);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), 76);

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAF102:AAF105)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 60);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 72);

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAF102:AAF105)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 60);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 72);

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAF102:AAF104)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAK110:AAN110)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("MMULT(AAA102:AAD105,AAA102:AAD105)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 94);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 116);
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), 138);

			assert.strictEqual(oParser.calculate().getElementRowCol(0, 1).getValue(), 32);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 1).getValue(), 40);
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 1).getValue(), 48);

			oParser = new parserFormula("MMULT(AAF102:AAF105,AAG102:AAG105)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("MMULT(AAF102:AAF105,AAA102:AAD102)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 4);
			assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), 8);
			assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), 12);

		});

		QUnit.test("Test: \"T(123)\"", function (assert) {
			oParser = new parserFormula("T(123)", "A1", ws);
			assert.ok(oParser.parse());
			assert.ok(!oParser.calculate().getValue(), "123");
		});

		QUnit.test("Test: YEAR", function (assert) {
			oParser = new parserFormula("YEAR(2013)", "A1", ws);
			assert.ok(oParser.parse());
			if (AscCommon.bDate1904) {
				assert.strictEqual(oParser.calculate().getValue(), 1909);
			} else {
				assert.strictEqual(oParser.calculate().getValue(), 1905);
			}

			testArrayFormula2(assert, "YEAR", 1, 1);
		});

		QUnit.test("Test: DAY", function (assert) {
			oParser = new parserFormula("DAY(2013)", "A1", ws);
			assert.ok(oParser.parse());
			if (AscCommon.bDate1904) {
				assert.strictEqual(oParser.calculate().getValue(), 6);
			} else {
				assert.strictEqual(oParser.calculate().getValue(), 5);
			}

			testArrayFormula2(assert, "DAY", 1, 1);
		});

		QUnit.test("Test: DAYS", function (assert) {
			ws.getRange2("A2").setValue("12/31/2011");
			ws.getRange2("A3").setValue("1/1/2011");

			oParser = new parserFormula('DAYS("3/15/11","2/1/11")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 42);

			oParser = new parserFormula("DAYS(A2,A3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 364);

			oParser = new parserFormula("DAYS(A2,A3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 364);

			oParser = new parserFormula('DAYS("2008-03-03","2008-03-01")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);

			oParser = new parserFormula('DAYS("2008-03-01","2008-03-03")', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -2);


			testArrayFormula2(assert, "DAYS", 2, 2);
		});

		QUnit.test("Test: DAY 2", function (assert) {
			oParser = new parserFormula("DAY(\"20 may 2045\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 20);
		});

		QUnit.test("Test: MONTH #1", function (assert) {
			oParser = new parserFormula("MONTH(2013)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 7);
		});

		QUnit.test("Test: MONTH #2", function (assert) {
			oParser = new parserFormula("MONTH(DATE(2013,2,2))", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);
		});

		QUnit.test("Test: MONTH #3", function (assert) {
			oParser = new parserFormula("MONTH(NOW())", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), new cDate().getUTCMonth() + 1);

			testArrayFormula2(assert, "MONTH", 1, 1);
		});

		QUnit.test("Test: \"10-3\"", function (assert) {
			oParser = new parserFormula("10-3", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 7);
		});

		QUnit.test("Test: \"SUM\"", function (assert) {

			ws.getRange2("S5").setValue("1");
			ws.getRange2("S6").setValue(numDivFact(-1, 2));
			ws.getRange2("S7").setValue(numDivFact(1, 4));
			ws.getRange2("S8").setValue(numDivFact(-1, 6));

			oParser = new parserFormula("SUM(S5:S8)", "A1", ws);
			assert.ok(oParser.parse());
//        assert.strictEqual( oParser.calculate().getValue(), 1-1/Math.fact(2)+1/Math.fact(4)-1/Math.fact(6) );
			assert.ok(Math.abs(oParser.calculate().getValue() - (1 - 1 / Math.fact(2) + 1 / Math.fact(4) - 1 / Math.fact(6))) < dif);
		});

		QUnit.test("Test: \"MAX\"", function (assert) {

			oParser = new parserFormula("MAX(-1, TRUE)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			oParser = new parserFormula("MAX(0, FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(0, FALSE)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of MAX(0, FALSE)");

			oParser = new parserFormula("MAX(25, 25.1, 25.01, 25.02, 25.2, 25.222, 25.333, 25.3334)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(25, 25.1, 25.01, 25.02, 25.2, 25.222, 25.333, 25.3334)");
			assert.strictEqual(oParser.calculate().getValue(), 25.3334, "Result of MAX(25, 25.1, 25.01, 25.02, 25.2, 25.222, 25.333, 25.3334)");

			oParser = new parserFormula("MAX(TRUE, FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(TRUE, FALSE)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MAX(TRUE, FALSE)");

			oParser = new parserFormula("MAX(FALSE, FALSE)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(FALSE, FALSE)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of MAX(FALSE, FALSE)");

			oParser = new parserFormula("MAX(str)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(str)");
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Result of MAX(str)");

			oParser = new parserFormula("MAX(49.08 - 432.81, 0)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(49.08 - 432.81, 0)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of MAX(49.08 - 432.81, 0)");

			oParser = new parserFormula("MAX(FALSE,-1-2,3-8,FALSE,TRUE)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(FALSE,-1-2,3-8,FALSE,TRUE)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MAX(FALSE,-1-2,3-8,FALSE,TRUE)");

			oParser = new parserFormula("MAX(49.08 - 432.81, 9,99999999999999E+43)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(49.08 - 432.81, 9,99999999999999E+43)");
			assert.strictEqual(oParser.calculate().getValue(), 9.9999999999999e+56, "Result of MAX(49.08 - 432.81, 9,99999999999999E+43)");

			oParser = new parserFormula("MAX(49.08 - 432.81, {9,99999999999999E+43})", "A1", ws);
			assert.ok(oParser.parse(), "MAX(49.08 - 432.81, {9,99999999999999E+43})");
			assert.strictEqual(oParser.calculate().getValue(), 9.9999999999999e+56, "Result of MAX(49.08 - 432.81, {9,99999999999999E+43})");

			oParser = new parserFormula("MAX(49.08 - 432.81, {12,13;14,15})", "A1", ws);
			assert.ok(oParser.parse(), "MAX(49.08 - 432.81, {12,13;14,15})");
			assert.strictEqual(oParser.calculate().getValue(), 15, "Result of MAX(49.08 - 432.81, {12,13;14,15})");

			oParser = new parserFormula("MAX({1,1,TRUE,2})", "A1", ws);
			assert.ok(oParser.parse(), "MAX({1,1,TRUE,2})");
			assert.strictEqual(oParser.calculate().getValue(), 2, "Result of MAX({1,1,TRUE,2})");

			oParser = new parserFormula("MAX({1,1,TRUE,2},{1,2,3,4,5,6,7,8,9,11,1})", "A1", ws);
			assert.ok(oParser.parse(), "MAX({1,1,TRUE,2},{1,2,3,4,5,6,7,8,9,11,1})");
			assert.strictEqual(oParser.calculate().getValue(), 11, "Result of MAX({1,1,TRUE,2},{1,2,3,4,5,6,7,8,9,11,1})");

			oParser = new parserFormula("MAX({1,1,TRUE,2},{12;12;13;11},{1,2,3,4,5,6,7,8,9,11,1})", "A1", ws);
			assert.ok(oParser.parse(), "MAX({1,1,TRUE,2},{12;12;13;11},{1,2,3,4,5,6,7,8,9,11,1})");
			assert.strictEqual(oParser.calculate().getValue(), 13, "Result of MAX({1,1,TRUE,2},{12;12;13;11},{1,2,3,4,5,6,7,8,9,11,1})");

			ws.getRange2("S5").setValue("1");
			ws.getRange2("S6").setValue(numDivFact(-1, 2));
			ws.getRange2("S7").setValue(numDivFact(1, 4));
			ws.getRange2("S8").setValue(numDivFact(-1, 6));

			oParser = new parserFormula("MAX(S5:S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			ws.getRange2("S5").setValue("#DIV/0!");
			ws.getRange2("S6").setValue("TRUE");
			ws.getRange2("S7").setValue("qwe");
			ws.getRange2("S8").setValue("");
			ws.getRange2("S9").setValue("-1");

			ws.getRange2("J10").setValue();
			ws.getRange2("J11").setValue("");
			ws.getRange2("J12").setValue("10");
			ws.getRange2("J13").setValue("7");
			ws.getRange2("J14").setValue("2");
			ws.getRange2("J15").setValue("27");
			ws.getRange2("J16").setValue("TRUE");
			ws.getRange2("J17").setValue("FALSE");
			ws.getRange2("J18").setValue("#N/A");
			ws.getRange2("J19").setValue("{2;3;4;5}");
			ws.getRange2("J20").setValue("{999;2;3;4;5}");
			ws.getRange2("J21").setValue("9.99999999999999E+43");
			ws.getRange2("J22").setValue("-9.99999999999999E+43");
			ws.getRange2("J23").setValue("0.000009");
			ws.getRange2("J24").setValue("-0.000009");
			ws.getRange2("J25").setValue("255");
			// string
			ws.getRange2("J25").setNumFormat("@");

			oParser = new parserFormula("MAX(J10)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J10)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of MAX(J10)");

			oParser = new parserFormula("MAX(J11)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J11)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of MAX(J11)");

			oParser = new parserFormula("MAX(J12)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J12)");
			assert.strictEqual(oParser.calculate().getValue(), 10, "Result of MAX(J12)");

			oParser = new parserFormula("MAX(J10:J17,J19:J24)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J10:J17,J19:J24)");
			assert.strictEqual(oParser.calculate().getValue(), 9.99999999999999E+43, "Result of MAX(J10:J17,J19:J24)");

			oParser = new parserFormula("MAX(J12:J19)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J12:J19)");
			assert.strictEqual(oParser.calculate().getValue(), "#N/A", "Result of MAX(J12:J19)");

			oParser = new parserFormula("MAX(J10:J25)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J10:J25)");
			assert.strictEqual(oParser.calculate().getValue(), "#N/A", "Result of MAX(J10:J25)");

			oParser = new parserFormula("MAX(J25, J10:J17)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J25, J10:J17)");
			assert.strictEqual(oParser.calculate().getValue(), 255, "Result of MAX(J25, J10:J17)");

			oParser = new parserFormula("MAX(J25, J10:J17, J18)", "A1", ws);
			assert.ok(oParser.parse(), "MAX(J25, J10:J17, J18)");
			assert.strictEqual(oParser.calculate().getValue(), "#N/A", "Result of MAX(J25, J10:J17, J18)");

			oParser = new parserFormula("MAX(S5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");

			oParser = new parserFormula("MAX(S6)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MAX(S7)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MAX(S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MAX(S5:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");

			oParser = new parserFormula("MAX(S6:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -1);

			testArrayFormula2(assert, "MAX", 1, 8, null, true);
		});

		QUnit.test("Test: \"MAXA\"", function (assert) {

			ws.getRange2("S5").setValue("1");
			ws.getRange2("S6").setValue(numDivFact(-1, 2));
			ws.getRange2("S7").setValue(numDivFact(1, 4));
			ws.getRange2("S8").setValue(numDivFact(-1, 6));

			oParser = new parserFormula("MAXA(S5:S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			ws.getRange2("S5").setValue("#DIV/0!");
			ws.getRange2("S6").setValue("TRUE");
			ws.getRange2("S7").setValue("qwe");
			ws.getRange2("S8").setValue("");
			ws.getRange2("S9").setValue("-1");
			oParser = new parserFormula("MAXA(S5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");
			oParser = new parserFormula("MAXA(S6)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);
			oParser = new parserFormula("MAXA(S7)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MAXA(S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MAXA(S5:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");
			oParser = new parserFormula("MAXA(S6:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);
			oParser = new parserFormula("MAXA(-1, TRUE)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			testArrayFormula2(assert, "MAXA", 1, 8, null, true);
		});

		QUnit.test("Test: \"MIN\"", function (assert) {

			ws.getRange2("S5").setValue("1");
			ws.getRange2("S6").setValue(numDivFact(-1, 2));
			ws.getRange2("S7").setValue(numDivFact(1, 4));
			ws.getRange2("S8").setValue(numDivFact(-1, 6));

			oParser = new parserFormula("MIN(S5:S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -1 / Math.fact(2));

			ws.getRange2("S5").setValue("#DIV/0!");
			ws.getRange2("S6").setValue("TRUE");
			ws.getRange2("S7").setValue("qwe");
			ws.getRange2("S8").setValue("");
			ws.getRange2("S9").setValue("2");
			oParser = new parserFormula("MIN(S5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");
			oParser = new parserFormula("MIN(S6)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MIN(S7)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MIN(S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MIN(S5:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");
			oParser = new parserFormula("MIN(S6:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);
			oParser = new parserFormula("MIN(2, TRUE)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			testArrayFormula2(assert, "min", 1, 8, null, true);
		});

		QUnit.test("Test: \"MINA\"", function (assert) {

			ws.getRange2("S5").setValue("1");
			ws.getRange2("S6").setValue(numDivFact(-1, 2));
			ws.getRange2("S7").setValue(numDivFact(1, 4));
			ws.getRange2("S8").setValue(numDivFact(-1, 6));

			oParser = new parserFormula("MINA(S5:S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -1 / Math.fact(2));

			ws.getRange2("S5").setValue("#DIV/0!");
			ws.getRange2("S6").setValue("TRUE");
			ws.getRange2("S7").setValue("qwe");
			ws.getRange2("S8").setValue("");
			ws.getRange2("S9").setValue("2");
			oParser = new parserFormula("MINA(S5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");
			oParser = new parserFormula("MINA(S6)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);
			oParser = new parserFormula("MINA(S7)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MINA(S8)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MINA(S5:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");
			oParser = new parserFormula("MINA(S6:S9)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);
			oParser = new parserFormula("MINA(2, TRUE)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			testArrayFormula2(assert, "mina", 1, 8, null, true);
		});

		QUnit.test("Test: SUM(S7:S9,{1,2,3})", function (assert) {
			ws.getRange2("S7").setValue("1");
			ws.getRange2("S8").setValue("2");
			ws.getRange2("S9").setValue("3");

			oParser = new parserFormula("SUM(S7:S9,{1,2,3})", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 12);
		});

		QUnit.test("Test: ISREF", function (assert) {
			oParser = new parserFormula("ISREF(G0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE");

			testArrayFormula2(assert, "ISREF", 1, 1, null, true);
		});

		QUnit.test("Test: ISTEXT", function (assert) {
			ws.getRange2("S7").setValue("test");

			oParser = new parserFormula("ISTEXT(S7)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			testArrayFormula2(assert, "ISTEXT", 1, 1);
		});

		QUnit.test("Test: MOD", function (assert) {
			oParser = new parserFormula("MOD(7,3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			oParser = new parserFormula("MOD(-10,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MOD(-9,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			oParser = new parserFormula("MOD(-8,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);

			oParser = new parserFormula("MOD(-7,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula("MOD(-6,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula("MOD(-5,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MOD(10,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MOD(9,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula("MOD(8,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 3);

			oParser = new parserFormula("MOD(15,5)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("MOD(15,0)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!");

			testArrayFormula2(assert, "MOD", 2, 2);
		});

		QUnit.test("Test: rename sheet #1", function (assert) {
			wb.dependencyFormulas.unlockRecal();
			ws.getRange2("S95").setValue("2");
			ws.getRange2("S100").setValue("=" + wb.getWorksheet(0).getName() + "!S95");
			ws.setName("SheetTmp");
			assert.strictEqual(ws.getCell2("S100").getFormula(), ws.getName() + "!S95");
			wb.dependencyFormulas.lockRecal();
		});

		QUnit.test("Test: wrong ref", function (assert) {
			oParser = new parserFormula("1+XXX1", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?");
		});

		QUnit.test("Test: \"CODE\"", function (assert) {
			oParser = new parserFormula("CODE(\"abc\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 97);

			oParser = new parserFormula("CODE(TRUE)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 84);

			testArrayFormula2(assert, "CODE", 1, 1);
		});

		QUnit.test("Test: \"CHAR\"", function (assert) {
			oParser = new parserFormula("CHAR(97)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "a");

			testArrayFormula2(assert, "CHAR", 1, 1);
		});

		QUnit.test("Test: \"CHAR(CODE())\"", function (assert) {
			oParser = new parserFormula("CHAR(CODE(\"A\"))", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "A");
		});

		QUnit.test("Test: \"PROPER\"", function (assert) {

			oParser = new parserFormula("PROPER(\"2-cent's worth\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "2-Cent'S Worth");

			oParser = new parserFormula("PROPER(\"76BudGet\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "76Budget");

			oParser = new parserFormula("PROPER(\"this is a TITLE\")", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "This Is A Title");

			oParser = new parserFormula('PROPER(TRUE)', "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "True");

			testArrayFormula2(assert, "PROPER", 1, 1);
		});

		QUnit.test("Test: \"GCD\"", function (assert) {
			oParser = new parserFormula("GCD(10,100,50)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 10);
			oParser = new parserFormula("GCD(24.6,36.2)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 12);
			oParser = new parserFormula("GCD(-1,39,52)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

			testArrayFormula2(assert, "GCD", 1, 8, null, true);
		});

		QUnit.test("Test: \"FIXED\"", function (assert) {
			oParser = new parserFormula("FIXED(1234567,-3)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "1,235,000");
			oParser = new parserFormula("FIXED(.555555,10)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "0.5555550000");
			oParser = new parserFormula("FIXED(1234567.555555,4,TRUE)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "1234567.5556");
			oParser = new parserFormula("FIXED(1234567)", "A1", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "1,234,567.00");

			testArrayFormula2(assert, "FIXED", 2, 3);
		});

		QUnit.test("Test: \"REPLACE\"", function (assert) {

			oParser = new parserFormula("REPLACE(\"abcdefghijk\",3,4,\"XY\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "abXYghijk");

			oParser = new parserFormula("REPLACE(\"abcdefghijk\",3,1,\"12345\")", "B2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "ab12345defghijk");

			oParser = new parserFormula("REPLACE(\"abcdefghijk\",15,4,\"XY\")", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "abcdefghijkXY");

			testArrayFormula2(assert, "REPLACE", 4, 4);
		});

		QUnit.test("Test: \"SEARCH\"", function (assert) {

			oParser = new parserFormula("SEARCH(\"~*\",\"abc*dEF\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula("SEARCH(\"~\",\"abc~dEF\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula("SEARCH(\"de\",\"abcdEF\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula("SEARCH(\"?c*e\",\"abcdEF\")", "B2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);

			oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",3)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 7);

			oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",30)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("SEARCH(\"pe\",\"dEFabcdEF\",2)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",2)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 7);

			oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",0)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",-2)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			testArrayFormula2(assert, "SEARCH", 2, 3);
		});

		QUnit.test("Test: \"SUBSTITUTE\"", function (assert) {

			oParser = new parserFormula("SUBSTITUTE(\"abcaAabca\",\"a\",\"xx\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "xxbcxxAxxbcxx");

			oParser = new parserFormula("SUBSTITUTE(\"abcaaabca\",\"a\",\"xx\")", "B2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "xxbcxxxxxxbcxx");

			oParser = new parserFormula("SUBSTITUTE(\"abcaaabca\",\"a\",\"\",10)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "bcbc");

			oParser = new parserFormula("SUBSTITUTE(\"abcaaabca\",\"a\",\"xx\",3)", "C2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "abcaxxabca");

			testArrayFormula2(assert, "SUBSTITUTE", 3, 4);
		});

		QUnit.test("Test: \"SHEET\"", function (assert) {

			oParser = new parserFormula("SHEET(Hi_Temps)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?");

			testArrayFormula2(assert, "SHEET", 1, 1, null, true);
		});

		QUnit.test("Test: \"SHEETS\"", function (assert) {

			oParser = new parserFormula("SHEETS(Hi_Temps)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#NAME?");

			oParser = new parserFormula("SHEETS()", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			testArrayFormula2(assert, "SHEETS", 1, 1, null, true);
		});

		QUnit.test("Test: \"TRIM\"", function (assert) {

			oParser = new parserFormula("TRIM(\"     abc         def      \")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "abc def");

			oParser = new parserFormula("TRIM(\" First Quarter Earnings \")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "First Quarter Earnings");

			testArrayFormula2(assert, "TRIM", 1, 1);
		});

		QUnit.test("Test: \"TRIMMEAN\"", function (assert) {
			ws.getRange2("A2").setValue("4");
			ws.getRange2("A3").setValue("5");
			ws.getRange2("A4").setValue("6");
			ws.getRange2("A5").setValue("7");
			ws.getRange2("A6").setValue("2");
			ws.getRange2("A7").setValue("3");
			ws.getRange2("A8").setValue("4");
			ws.getRange2("A9").setValue("5");
			ws.getRange2("A10").setValue("1");
			ws.getRange2("A11").setValue("2");
			ws.getRange2("A12").setValue("3");

			oParser = new parserFormula("TRIMMEAN(A2:A12,0.2)", "A1", ws);
			assert.ok(oParser.parse(), "TRIMMEAN(A2:A12,0.2)");
			assert.strictEqual(oParser.calculate().getValue().toFixed(3) - 0, 3.778, "TRIMMEAN(A2:A12,0.2)");

			//TODO нужна другая функция для тестирования
			//testArrayFormula2(assert, "TRIMMEAN", 2, 2);
		});

		QUnit.test("Test: \"DOLLAR\"", function (assert) {

			oParser = new parserFormula("DOLLAR(1234.567)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "$1,234.57");

			oParser = new parserFormula("DOLLAR(1234.567,-2)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "$1,200");

			oParser = new parserFormula("DOLLAR(-1234.567,4)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "($1,234.5670)");

			testArrayFormula2(assert, "DOLLAR", 2, 2);
		});

		QUnit.test("Test: \"EXACT\"", function (assert) {

			ws.getRange2("A2").setValue("word");
			ws.getRange2("A3").setValue("Word");
			ws.getRange2("A4").setValue("w ord");
			ws.getRange2("B2").setValue("word");
			ws.getRange2("B3").setValue("word");
			ws.getRange2("B4").setValue("word");

			oParser = new parserFormula("EXACT(A2,B2)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula("EXACT(A3,B3)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE");

			oParser = new parserFormula("EXACT(A4,B4)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE");

			oParser = new parserFormula("EXACT(TRUE,TRUE)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula('EXACT("TRUE",TRUE)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula('EXACT("TRUE","TRUE")', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "TRUE");

			oParser = new parserFormula('EXACT("true",TRUE)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "FALSE");

			testArrayFormula2(assert, "EXACT", 2, 2);
		});

		QUnit.test("Test: \"LEFT\"", function (assert) {

			ws.getRange2("A2").setValue("Sale Price");
			ws.getRange2("A3").setValue("Sweden");


			oParser = new parserFormula("LEFT(A2,4)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "Sale");

			oParser = new parserFormula("LEFT(A3)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "S");

			testArrayFormula2(assert, "LEFT", 1, 2);
		});

		QUnit.test("Test: \"LEN\"", function (assert) {

			ws.getRange2("A201").setValue("Phoenix, AZ");
			ws.getRange2("A202").setValue("");
			ws.getRange2("A203").setValue("     One   ");

			oParser = new parserFormula("LEN(A201)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 11);

			oParser = new parserFormula("LEN(A202)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("LEN(A203)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 11);

			oParser = new parserFormula('LEN(TRUE)', "A2", ws);
			assert.ok(oParser.parse(), 'LEN(TRUE)');
			assert.strictEqual(oParser.calculate().getValue(), 4, 'LEN(TRUE)');

			testArrayFormula2(assert, "LEN", 1, 1);
		});

		QUnit.test("Test: \"REPT\"", function (assert) {

			oParser = new parserFormula('REPT("*-", 3)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "*-*-*-");

			oParser = new parserFormula('REPT("-",10)', "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "----------");

			testArrayFormula2(assert, "REPT", 2, 2);
		});

		QUnit.test("Test: \"RIGHT\"", function (assert) {

			ws.getRange2("A2").setValue("Sale Price");
			ws.getRange2("A3").setValue("Stock Number");

			oParser = new parserFormula("RIGHT(A2,5)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "Price");

			oParser = new parserFormula("RIGHT(A3)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "r");

			testArrayFormula2(assert, "RIGHT", 1, 2);
		});


		QUnit.test("Test: \"VALUE\"", function (assert) {

			oParser = new parserFormula("VALUE(\"123.456\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 123.456);

			oParser = new parserFormula("VALUE(\"$1,000\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1000);

			oParser = new parserFormula("VALUE(\"23-Mar-2002\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 37338);

			oParser = new parserFormula("VALUE(\"03-26-2006\")", "A2", ws);
			assert.ok(oParser.parse());

			if (AscCommon.bDate1904) {
				assert.strictEqual(oParser.calculate().getValue(), 37340);
			} else {
				assert.strictEqual(oParser.calculate().getValue(), 38802);
			}

			oParser = new parserFormula("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), AscCommon.g_oFormatParser.parse("16:48:00").value - AscCommon.g_oFormatParser.parse("12:17:12").value);

			testArrayFormula2(assert, "value", 1, 1);
		});

		QUnit.test("Test: \"DATE\"", function (assert) {

			testArrayFormula2(assert, "DATE", 3, 3);
		});

		QUnit.test("Test: \"DATEVALUE\"", function (assert) {

			oParser = new parserFormula("DATEVALUE(\"10-10-2010 10:26\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 40461);

			oParser = new parserFormula("DATEVALUE(\"10-10-2010 10:26\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 40461);

			tmp = ws.getRange2("A7");
			tmp.setNumFormat('@');
			tmp.setValue("3-Mar");
			oParser = new parserFormula("DATEVALUE(A7)", "A2", ws);
			assert.ok(oParser.parse());
			var d = new cDate();
			d.setUTCMonth(2);
			d.setUTCDate(3);
			assert.strictEqual(oParser.calculate().getValue(), d.getExcelDate());

			oParser = new parserFormula("DATEVALUE(\"$1,000\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("DATEVALUE(\"23-Mar-2002\")", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 37338);

			oParser = new parserFormula("DATEVALUE(\"03-26-2006\")", "A2", ws);
			assert.ok(oParser.parse());

			if (AscCommon.bDate1904) {
				assert.strictEqual(oParser.calculate().getValue(), 37340);
			} else {
				assert.strictEqual(oParser.calculate().getValue(), 38802);
			}

			testArrayFormula(assert, "DATEVALUE");
		});

		QUnit.test("Test: \"EDATE\"", function (assert) {

			if (!AscCommon.bDate1904) {
				oParser = new parserFormula("EDATE(DATE(2006,1,31),5)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 38898);

				oParser = new parserFormula("EDATE(DATE(2004,2,29),12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 38411);

				ws.getRange2("A7").setValue("02-28-2004");
				oParser = new parserFormula("EDATE(A7,12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 38411);

				oParser = new parserFormula("EDATE(DATE(2004,1,15),-23)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 37302);
			} else {
				oParser = new parserFormula("EDATE(DATE(2006,1,31),5)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 37436);

				oParser = new parserFormula("EDATE(DATE(2004,2,29),12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 36949);

				ws.getRange2("A7").setValue("02-28-2004");
				oParser = new parserFormula("EDATE(A7,12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 36949);

				oParser = new parserFormula("EDATE(DATE(2004,1,15),-23)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 35840);
			}

			testArrayFormula2(assert, "EDATE", 2, 2, true, null);
		});

		QUnit.test("Test: \"EOMONTH\"", function (assert) {

			if (!AscCommon.bDate1904) {
				oParser = new parserFormula("EOMONTH(DATE(2006,1,31),5)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 38898);

				oParser = new parserFormula("EOMONTH(DATE(2004,2,29),12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 38411);

				ws.getRange2("A7").setValue("02-28-2004");
				oParser = new parserFormula("EOMONTH(A7,12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 38411);

				oParser = new parserFormula("EOMONTH(DATE(2004,1,15),-23)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 37315);
			} else {
				oParser = new parserFormula("EOMONTH(DATE(2006,1,31),5)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 37436);

				oParser = new parserFormula("EOMONTH(DATE(2004,2,29),12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 36949);

				ws.getRange2("A7").setValue("02-28-2004");
				oParser = new parserFormula("EOMONTH(A7,12)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 36949);

				oParser = new parserFormula("EOMONTH(DATE(2004,1,15),-23)", "A2", ws);
				assert.ok(oParser.parse());
				assert.strictEqual(oParser.calculate().getValue(), 35853);
			}

			testArrayFormula2(assert, "EOMONTH", 2, 2, true, null);
		});

		QUnit.test("Test: \"EVEN\"", function (assert) {

			oParser = new parserFormula("EVEN(1.5)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);

			oParser = new parserFormula("EVEN(3)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4);

			oParser = new parserFormula("EVEN(2)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2);

			oParser = new parserFormula("EVEN(-1)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -2);

			testArrayFormula(assert, "EVEN");

		});


		QUnit.test("Test: \"NETWORKDAYS\"", function (assert) {

			oParser = new parserFormula("NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31))", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 22);

			oParser = new parserFormula("NETWORKDAYS(DATE(2006,1,31),DATE(2006,1,1))", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), -22);

			oParser = new parserFormula("NETWORKDAYS(DATE(1700,1,1),DATE(1700,2,2))", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 23);

			oParser = new parserFormula("NETWORKDAYS(DATE(2006,1,1),DATE(2006,2,1),{\"01-02-2006\",\"01-16-2006\"})", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 21);

			oParser = new parserFormula("NETWORKDAYS(0,0)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(0,0)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(0,0)");

			// in js new Date(1900,0,1) === monday, in ms 01.01.1990 === sunday
			oParser = new parserFormula("NETWORKDAYS(1,1)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1,1)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(1,1)");

			oParser = new parserFormula("NETWORKDAYS(2,2)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(2,2)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(2,2)");

			oParser = new parserFormula("NETWORKDAYS(3,3)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(3,3)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(3,3)");

			oParser = new parserFormula("NETWORKDAYS(4,4)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(4,4)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(4,4)");

			oParser = new parserFormula("NETWORKDAYS(5,5)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(5,5)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(5,5)");

			oParser = new parserFormula("NETWORKDAYS(6,6)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(6,6)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(6,6)");

			oParser = new parserFormula("NETWORKDAYS(7,7)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(7,7)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(7,7)");

			oParser = new parserFormula("NETWORKDAYS(8,8)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(8,8)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(8,8)");

			oParser = new parserFormula("NETWORKDAYS(9,9)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(9,9)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(9,9)");

			oParser = new parserFormula("NETWORKDAYS(10,10)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(10,10)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(10,10)");

			oParser = new parserFormula("NETWORKDAYS(11,11)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(11,11)");
			assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(11,11)");

			oParser = new parserFormula("NETWORKDAYS(0,11)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(0,11)");
			assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS(0,11)");

			oParser = new parserFormula("NETWORKDAYS(1,11)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1,11)");
			assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS(1,11)");

			oParser = new parserFormula("NETWORKDAYS(11,0)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(11,0)");
			assert.strictEqual(oParser.calculate().getValue(), -8, "Result of NETWORKDAYS(11,0)");

			oParser = new parserFormula("NETWORKDAYS(11,1)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(11,1)");
			assert.strictEqual(oParser.calculate().getValue(), -8, "Result of NETWORKDAYS(11,1)");

			oParser = new parserFormula("NETWORKDAYS(-1,15)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(-1,15)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of NETWORKDAYS(-1,15)");

			oParser = new parserFormula("NETWORKDAYS(15,-1)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(15,-1)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of NETWORKDAYS(15,-1)");

			oParser = new parserFormula("NETWORKDAYS(-1,-15)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(-1,-15)");
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of NETWORKDAYS(-1,-15)");

			oParser = new parserFormula("NETWORKDAYS(1,3889)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1,3889)");
			assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1,3889)");

			oParser = new parserFormula("NETWORKDAYS(1,45689)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1,45689)");
			assert.strictEqual(oParser.calculate().getValue(), 32635, "Result of NETWORKDAYS(1,45689)");

			oParser = new parserFormula("NETWORKDAYS(0.1,0.9)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(0.1,0.9)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(0.1,0.9)");

			oParser = new parserFormula("NETWORKDAYS(1.1,3889)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1.1,3889)");
			assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1.1,3889)");

			oParser = new parserFormula("NETWORKDAYS(1.9,3889)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1.9,3889)");
			assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1.9,3889)");

			oParser = new parserFormula("NETWORKDAYS(1,3889.1)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1,3889.1)");
			assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1,3889.1)");

			oParser = new parserFormula("NETWORKDAYS(1.9,3889.9)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1.9,3889.9)");
			assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1.9,3889.9)");

			// bool
			oParser = new parserFormula("NETWORKDAYS(11,TRUE)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(11,TRUE)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(11,TRUE)");

			oParser = new parserFormula("NETWORKDAYS(TRUE,TRUE)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(TRUE,TRUE)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(TRUE,TRUE)");

			oParser = new parserFormula("NETWORKDAYS(TRUE,11)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(TRUE,11)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(TRUE,11)");

			oParser = new parserFormula("NETWORKDAYS(#VALUE!,#NUM!)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(#VALUE!,#NUM!)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(#VALUE!,#NUM!)");

			// array
			oParser = new parserFormula("NETWORKDAYS({1,11,255},11)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS({1,11,255},11)");
			assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS({1,11,255},11)");

			oParser = new parserFormula("NETWORKDAYS(1,{11,85,255})", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(1,{11,85,255})");
			assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS(1,{11,85,255})");

			oParser = new parserFormula("NETWORKDAYS({1,11,255},{11,85,255})", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS({1,11,255},{11,85,255})");
			assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS({1,11,255},{11,85,255})");

			ws.getRange2("A101").setValue();
			ws.getRange2("A102").setValue("");
			ws.getRange2("A103").setValue("0");
			ws.getRange2("A104").setValue("9");
			ws.getRange2("A105").setValue("25");
			ws.getRange2("A106").setValue("TRUE");
			ws.getRange2("A107").setValue("FALSE");
			ws.getRange2("A108").setValue("{999,25,0}");
			ws.getRange2("A109").setValue("{777,25,0}");
			ws.getRange2("A110").setValue("{0,777,25,0}");
			ws.getRange2("A111").setValue("#N/A");
			ws.getRange2("A112").setValue("99999999999999999999");
			ws.getRange2("A113").setValue("-99999999999999999999");
			ws.getRange2("A114").setValue("str");
			ws.getRange2("A115").setValue("str2");

			ws.getRange2("B101").setValue("0");
			ws.getRange2("B102").setValue("1");
			ws.getRange2("B103").setValue("4");
			ws.getRange2("B104").setValue("9");
			ws.getRange2("B105").setValue("25");
			ws.getRange2("B106").setValue("255");
			ws.getRange2("B107").setValue("312778");

			// cellsrange
			oParser = new parserFormula("NETWORKDAYS(A101:A105,A105)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A101:A105,25)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(A101:A105,25)");

			oParser = new parserFormula("NETWORKDAYS(A104,A101:A105)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A104,A101:A105)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(A104,A101:A105)");

			oParser = new parserFormula("NETWORKDAYS(A101:A105,A101:A105)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A101:A105,A101:A105)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(A101:A105,A101:A105)");

			oParser = new parserFormula("NETWORKDAYS(B101:B107,B101:B107)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(B101:B107,B101:B107)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(B101:B107,B101:B107)");

			// cells
			oParser = new parserFormula("NETWORKDAYS(A102,A102)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A102,A102)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS('','')");

			oParser = new parserFormula("NETWORKDAYS(A102:A102,A102:A102)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A102:A102,A102:A102)");
			assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS('','')");

			oParser = new parserFormula("NETWORKDAYS(A103,A104)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A103,A104)");
			assert.strictEqual(oParser.calculate().getValue(), 6, "Result of NETWORKDAYS(0,9)");

			oParser = new parserFormula("NETWORKDAYS(A104:A104,A104:A104)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			oParser = new parserFormula("NETWORKDAYS(A106,A107)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("NETWORKDAYS(A109,A108)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS({777,25,0},{999,25,0})");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS({777,25,0},{999,25,0})");

			oParser = new parserFormula("NETWORKDAYS(A105,A108)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(25,{999,25,0})");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(25,{999,25,0})");

			oParser = new parserFormula("NETWORKDAYS(A108,25)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS({999,25,0},25)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS({999,25,0},25)");

			oParser = new parserFormula("NETWORKDAYS(A111,A105)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A114,A115)");
			assert.strictEqual(oParser.calculate().getValue(), "#N/A", "Result of NETWORKDAYS(str,str2)");

			oParser = new parserFormula("NETWORKDAYS(A114,A115)", "A2", ws);
			assert.ok(oParser.parse(), "NETWORKDAYS(A114,A115)");
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(str,str2)");

			// bug case
			oParser = new parserFormula("NETWORKDAYS(A101,A101)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("NETWORKDAYS(A101,A102)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("NETWORKDAYS(A101,A109)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("NETWORKDAYS(A102,A109)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

			oParser = new parserFormula("NETWORKDAYS(A101:A101,A101:A101)", "A2", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);


			testArrayFormula2(assert, "NETWORKDAYS", 2, 3, true, null);
		});

		QUnit.test("Test: \"NETWORKDAYS.INTL\"", function (assert) {

			var formulaStr = "NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))";
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 22, formulaStr);

			formulaStr = "NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))";
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), -21, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 22, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),17,{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 26, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"1111111",{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 0, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 20, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0000000",{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 30, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"19",{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),19,{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), "#NUM!", formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(1901,1,1),DATE(2006,2,1),"0000000",{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 38381, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(DATE(1901,1,1),DATE(2006,2,1),17,{"1/2/2006","1/16/2006"})';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 32898, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,11)';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 8490, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,1)';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 7075, formulaStr);

			formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,2)';
			oParser = new parserFormula(formulaStr, "A2", ws);
			assert.ok(oParser.parse(), formulaStr);
			assert.strictEqual(oParser.calculate().getValue(), 7075, formulaStr);

			//TODO посмотреть почему неверно считается
			//проблема повторяется с новым и со старым вариантом реализации NETWORKDAYS.INTL

			/*formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,5)';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7074, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,5,{123,1000})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7073, formulaStr );*/
		});

		QUnit.test("Test: \"N\"", function (assert) {

			ws.getRange2("A2").setValue("7");
			ws.getRange2("A3").setValue("Even");
			ws.getRange2("A4").setValue("TRUE");
			ws.getRange2("A5").setValue("4/17/2011");

			oParser = new parserFormula("N(A2)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 7);

			oParser = new parserFormula("N(A3)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("N(A4)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 1);

			oParser = new parserFormula("N(A5)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 40650);

			oParser = new parserFormula('N("7")', "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			//TODO нужна другая функция для тестирования
			//testArrayFormula2(assert, "N", 1, 1);
		});

		QUnit.test("Test: \"SUMIF\"", function (assert) {

			ws.getRange2("A2").setValue("100000");
			ws.getRange2("A3").setValue("200000");
			ws.getRange2("A4").setValue("300000");
			ws.getRange2("A5").setValue("400000");

			ws.getRange2("B2").setValue("7000");
			ws.getRange2("B3").setValue("14000");
			ws.getRange2("B4").setValue("21000");
			ws.getRange2("B5").setValue("28000");

			ws.getRange2("C2").setValue("250000");

			oParser = new parserFormula("SUMIF(A2:A5,\">160000\",B2:B5)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 63000);

			oParser = new parserFormula("SUMIF(A2:A5,\">160000\")", "A8", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 900000);

			oParser = new parserFormula("SUMIF(A2:A5,300000,B2:B5)", "A9", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 21000);

			oParser = new parserFormula("SUMIF(A2:A5,\">\" & C2,B2:B5)", "A10", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 49000);

			oParser = new parserFormula("SUMIF(A2,\">160000\",B2:B5)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("SUMIF(A3,\">160000\",B2:B5)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 7000);

			oParser = new parserFormula("SUMIF(A4,\">160000\",B4:B5)", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 21000);

			oParser = new parserFormula("SUMIF(A4,\">160000\")", "A7", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 300000);


			ws.getRange2("A12").setValue("Vegetables");
			ws.getRange2("A13").setValue("Vegetables");
			ws.getRange2("A14").setValue("Fruits");
			ws.getRange2("A15").setValue("");
			ws.getRange2("A16").setValue("Vegetables");
			ws.getRange2("A17").setValue("Fruits");

			ws.getRange2("B12").setValue("Tomatoes");
			ws.getRange2("B13").setValue("Celery");
			ws.getRange2("B14").setValue("Oranges");
			ws.getRange2("B15").setValue("Butter");
			ws.getRange2("B16").setValue("Carrots");
			ws.getRange2("B17").setValue("Apples");

			ws.getRange2("C12").setValue("2300");
			ws.getRange2("C13").setValue("5500");
			ws.getRange2("C14").setValue("800");
			ws.getRange2("C15").setValue("400");
			ws.getRange2("C16").setValue("4200");
			ws.getRange2("C17").setValue("1200");

			oParser = new parserFormula("SUMIF(A12:A17,\"Fruits\",C12:C17)", "A19", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 2000);

			oParser = new parserFormula("SUMIF(A12:A17,\"Vegetables\",C12:C17)", "A20", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 12000);

			oParser = new parserFormula("SUMIF(B12:B17,\"*es\",C12:C17)", "A21", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 4300);

			oParser = new parserFormula("SUMIF(A12:A17,\"\",C12:C17)", "A22", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 400);

		});

		QUnit.test("Test: \"SUMIFS\"", function (assert) {

			ws.getRange2("A2").setValue("5");
			ws.getRange2("A3").setValue("4");
			ws.getRange2("A4").setValue("15");
			ws.getRange2("A5").setValue("3");
			ws.getRange2("A6").setValue("22");
			ws.getRange2("A7").setValue("12");
			ws.getRange2("A8").setValue("10");
			ws.getRange2("A9").setValue("33");

			ws.getRange2("B2").setValue("Apples");
			ws.getRange2("B3").setValue("Apples");
			ws.getRange2("B4").setValue("Artichokes");
			ws.getRange2("B5").setValue("Artichokes");
			ws.getRange2("B6").setValue("Bananas");
			ws.getRange2("B7").setValue("Bananas");
			ws.getRange2("B8").setValue("Carrots");
			ws.getRange2("B9").setValue("Carrots");

			ws.getRange2("C2").setValue("Tom");
			ws.getRange2("C3").setValue("Sarah");
			ws.getRange2("C4").setValue("Tom");
			ws.getRange2("C5").setValue("Sarah");
			ws.getRange2("C6").setValue("Tom");
			ws.getRange2("C7").setValue("Sarah");
			ws.getRange2("C8").setValue("Tom");
			ws.getRange2("C9").setValue("Sarah");

			oParser = new parserFormula("SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")", "A10", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 20);

			oParser = new parserFormula("SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")", "A11", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 30);

			oParser = new parserFormula("SUMIFS(D:D,E:E,$H2)", "A11", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			oParser = new parserFormula("SUMIFS(C:D,E:E,$H2)", "A11", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");
		});

		QUnit.test("Test: \"MAXIFS\"", function (assert) {

			ws.getRange2("AAA2").setValue("10");
			ws.getRange2("AAA3").setValue("1");
			ws.getRange2("AAA4").setValue("100");
			ws.getRange2("AAA5").setValue("1");
			ws.getRange2("AAA6").setValue("1");
			ws.getRange2("AAA7").setValue("50");

			ws.getRange2("BBB2").setValue("b");
			ws.getRange2("BBB3").setValue("a");
			ws.getRange2("BBB4").setValue("a");
			ws.getRange2("BBB5").setValue("b");
			ws.getRange2("BBB6").setValue("a");
			ws.getRange2("BBB7").setValue("b");

			ws.getRange2("DDD2").setValue("100");
			ws.getRange2("DDD3").setValue("100");
			ws.getRange2("DDD4").setValue("200");
			ws.getRange2("DDD5").setValue("300");
			ws.getRange2("DDD6").setValue("100");
			ws.getRange2("DDD7").setValue("400");

			oParser = new parserFormula('MAXIFS(AAA2:AAA7,BBB2:BBB7,"b",DDD2:DDD7,">100")', "A22", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 50);

			oParser = new parserFormula('MAXIFS(AAA2:AAA6,BBB2:BBB6,"a",DDD2:DDD6,">200")', "A22", ws);
			assert.ok(oParser.parse());
			assert.strictEqual(oParser.calculate().getValue(), 0);

			testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A", "MAXIFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2,A1:C2)");
			testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;0,0,0,#N/A;#N/A,#N/A,#N/A,#N/A", "MAXIFS(A1:C2,A1:C2,A1:A1,A1:C2,A1:C2,A1:C2,A1:C2)");
			testArrayFormulaEqualsValues(assert, "1,0,0,#N/A;2,0,0,#N/A;#N/A,#N/A,#N/A,#N/A", "MAXIFS(A1:C2,A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");
		});





		wb.dependencyFormulas.unlockRecal();
	});
	