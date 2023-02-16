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

		

		wb.dependencyFormulas.unlockRecal();
	});
