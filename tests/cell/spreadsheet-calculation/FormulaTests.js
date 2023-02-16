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


			let docInfo = new Asc.asc_CDocInfo();
			docInfo.asc_putTitle("TeSt.xlsx");

			let api = {
				wb: {
					getWorksheet: function () {
					}
				}, DocInfo: docInfo,
				isEyedropperStarted: function() {return false;}
			};
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




		wb.dependencyFormulas.unlockRecal();
	});
