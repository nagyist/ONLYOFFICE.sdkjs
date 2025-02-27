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
function (window, undefined) {
	var cDate = Asc.cDate;
	var g_oFormatParser = AscCommon.g_oFormatParser;

	var cErrorType = AscCommonExcel.cErrorType;
	var c_sPerDay = AscCommonExcel.c_sPerDay;
	var c_msPerDay = AscCommonExcel.c_msPerDay;
	var cNumber = AscCommonExcel.cNumber;
	var cString = AscCommonExcel.cString;
	var cBool = AscCommonExcel.cBool;
	var cError = AscCommonExcel.cError;
	var cArea = AscCommonExcel.cArea;
	var cArea3D = AscCommonExcel.cArea3D;
	var cRef = AscCommonExcel.cRef;
	var cRef3D = AscCommonExcel.cRef3D;
	var cEmpty = AscCommonExcel.cEmpty;
	var cArray = AscCommonExcel.cArray;
	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var cElementType = AscCommonExcel.cElementType;
	var argType = Asc.c_oAscFormulaArgumentType;

	var GetDiffDate360 = AscCommonExcel.GetDiffDate360;

	var cExcelDateTimeDigits = 8; //количество цифр после запятой в числах отвечающих за время специализация $18.17.4.2

	var DayCountBasis = {
		// US 30/360
		UsPsa30_360: 0, // Actual/Actual
		ActualActual: 1, // Actual/360
		Actual360: 2, // Actual/365
		Actual365: 3, // European 30/360
		Europ30_360: 4
	};

	function yearFrac(d1, d2, mode) {

		d1.truncate();
		d2.truncate();

		var date1 = d1.getUTCDate(), month1 = d1.getUTCMonth() + 1, year1 = d1.getUTCFullYear(),
			date2 = d2.getUTCDate(), month2 = d2.getUTCMonth() + 1, year2 = d2.getUTCFullYear();

		switch (mode) {
			case DayCountBasis.UsPsa30_360:
				return new cNumber(Math.abs(GetDiffDate360(date1, month1, year1, date2, month2, year2, true)) / 360);
			case DayCountBasis.ActualActual:
				var yc = Math.abs(year2 - year1), sd = year1 > year2 ? new cDate(d2) : new cDate(d1),
					yearAverage = sd.isLeapYear() ? 366 : 365, dayDiff = /*Math.abs*/(d2 - d1);
				for (var i = 0; i < yc; i++) {
					sd.addYears(1);
					yearAverage += sd.isLeapYear() ? 366 : 365;
				}
				yearAverage /= (yc + 1);
				dayDiff /= (yearAverage * c_msPerDay);
				return new cNumber(Math.abs(dayDiff));
			case DayCountBasis.Actual360:
				var dayDiff = Math.abs(d2 - d1);
				dayDiff /= (360 * c_msPerDay);
				return new cNumber(dayDiff);
			case DayCountBasis.Actual365:
				var dayDiff = Math.abs(d2 - d1);
				dayDiff /= (365 * c_msPerDay);
				return new cNumber(dayDiff);
			case DayCountBasis.Europ30_360:
				return new cNumber(Math.abs(GetDiffDate360(date1, month1, year1, date2, month2, year2, false)) / 360);
			default:
				return new cError(cErrorType.not_numeric);
		}
	}

	function diffDate(d1, d2, mode) {
		var date1 = d1.getUTCDate(), month1 = d1.getUTCMonth() + 1, year1 = d1.getUTCFullYear(),
			date2 = d2.getUTCDate(), month2 = d2.getUTCMonth() + 1, year2 = d2.getUTCFullYear();

		switch (mode) {
			case DayCountBasis.UsPsa30_360:
				return new cNumber(GetDiffDate360(date1, month1, year1, date2, month2, year2, true));
			case DayCountBasis.ActualActual:
				return new cNumber((d2 - d1) / c_msPerDay);
			case DayCountBasis.Actual360:
				var dayDiff = d2 - d1;
				dayDiff /= c_msPerDay;
				return new cNumber((d2 - d1) / c_msPerDay);
			case DayCountBasis.Actual365:
				var dayDiff = d2 - d1;
				dayDiff /= c_msPerDay;
				return new cNumber((d2 - d1) / c_msPerDay);
			case DayCountBasis.Europ30_360:
				return new cNumber(GetDiffDate360(date1, month1, year1, date2, month2, year2, false));
			default:
				return new cError(cErrorType.not_numeric);
		}
	}

	function diffDate2(d1, d2, mode) {
		var date1 = d1.getUTCDate(), month1 = d1.getUTCMonth(), year1 = d1.getUTCFullYear(), date2 = d2.getUTCDate(),
			month2 = d2.getUTCMonth(), year2 = d2.getUTCFullYear();

		var nDaysInYear, nYears, nDayDiff;

		switch (mode) {
			case DayCountBasis.UsPsa30_360:
				nDaysInYear = 360;
				nYears = year1 - year2;
				nDayDiff = Math.abs(GetDiffDate360(date1, month1 + 1, year1, date2, month2 + 1, year2, true)) -
					nYears * nDaysInYear;
				return new cNumber(nYears + nDayDiff / nDaysInYear);
			case DayCountBasis.ActualActual:
				nYears = year2 - year1;
				nDaysInYear = d1.isLeapYear() ? 366 : 365;

				var dayDiff;

				if (nYears && (month1 > month2 || (month1 == month2 && date1 > date2))) {
					nYears--;
				}

				if (nYears) {
					dayDiff = parseInt((d2 - new cDate(Date.UTC(year2, month1, date1))) / c_msPerDay);
				} else {
					dayDiff = parseInt((d2 - d1) / c_msPerDay);
				}

				if (dayDiff < 0) {
					dayDiff += nDaysInYear;
				}
				return new cNumber(nYears + dayDiff / nDaysInYear);
			case DayCountBasis.Actual360:
				nDaysInYear = 360;
				nYears = parseInt((d2 - d1) / c_msPerDay / nDaysInYear);
				nDayDiff = (d2 - d1) / c_msPerDay;
				nDayDiff %= nDaysInYear;
				return new cNumber(nYears + nDayDiff / nDaysInYear);
			case DayCountBasis.Actual365:
				nDaysInYear = 365;
				nYears = parseInt((d2 - d1) / c_msPerDay / nDaysInYear);
				nDayDiff = (d2 - d1) / c_msPerDay;
				nDayDiff %= nDaysInYear;
				return new cNumber(nYears + nDayDiff / nDaysInYear);
			case DayCountBasis.Europ30_360:
				nDaysInYear = 360;
				nYears = year1 - year2;
				nDayDiff = Math.abs(GetDiffDate360(date1, month1 + 1, year1, date2, month2 + 1, year2, false)) -
					nYears * nDaysInYear;
				return new cNumber(nYears + nDayDiff / nDaysInYear);
			default:
				return new cError(cErrorType.not_numeric);
		}
	}

	function GetDiffDate(d1, d2, nMode, av) {
		var bNeg = d1 > d2, nRet, pOptDaysIn1stYear;

		if (bNeg) {
			var n = d2;
			d2 = d1;
			d1 = n;
		}

		var nD1 = d1.getUTCDate(), nM1 = d1.getUTCMonth(), nY1 = d1.getUTCFullYear(), nD2 = d2.getUTCDate(),
			nM2 = d2.getUTCMonth(), nY2 = d2.getUTCFullYear();

		switch (nMode) {
			case DayCountBasis.UsPsa30_360:            // 0=USA (NASD) 30/360
			case DayCountBasis.Europ30_360: {            // 4=Europe 30/360
				var bLeap = d1.isLeapYear(), nDays, nMonths/*, nYears*/;

				nMonths = nM2 - nM1;
				nDays = nD2 - nD1;
				nMonths += (nY2 - nY1) * 12;
				nRet = nMonths * 30 + nDays;
				if (nMode == 0 && nM1 == 2 && nM2 != 2 && nY1 == nY2) {
					nRet -= bLeap ? 1 : 2;
				}

				pOptDaysIn1stYear = 360;
				break;
			}
			case DayCountBasis.ActualActual:            // 1=exact/exact
				pOptDaysIn1stYear = d1.isLeapYear() ? 366 : 365;
				nRet = d2 - d1;
				break;
			case DayCountBasis.Actual360:            // 2=exact/360
				nRet = d2 - d1;
				pOptDaysIn1stYear = 360;
				break;
			case DayCountBasis.Actual365:            //3=exact/365
				nRet = d2 - d1;
				pOptDaysIn1stYear = 365;
				break;
		}

		return (bNeg ? -nRet : nRet) / c_msPerDay / (av ? 1 : pOptDaysIn1stYear);
	}

	function days360(date1, date2, flag, isAccrint) {
		let sign;
		let nY1 = date1.getUTCFullYear(), nM1 = date1.getUTCMonth() + 1, nD1 = date1.getUTCDate(),
			nY2 = date2.getUTCFullYear(), nM2 = date2.getUTCMonth() + 1, nD2 = date2.getUTCDate();

		if (flag && !isAccrint && (date2 < date1)) {
			sign = date1;
			date1 = date2;
			date2 = sign;
			sign = -1;
		} else {
			sign = 1;
		}
		if (nD1 == 31) {
			nD1 -= 1;
		} else if (!flag) {
			if (nM1 == 2) {
				switch (nD1) {
					case 28 :
						if (!date1.isLeapYear()) {
							nD1 = 30;
						}
						break;
					case 29 :
						nD1 = 30;
						break;
				}
			}
		}
		if (nD2 == 31) {
			if (!flag) {
				if (nD1 == 30) {
					nD2--;
				}
			} else {
				nD2 = 30;
			}
		}
		return sign * (nD2 - nD1 + (nM2 - nM1) * 30 + (nY2 - nY1) * 360);
	}

	function daysInYear(date, basis) {
		switch (basis) {
			case DayCountBasis.UsPsa30_360:         // 0=USA (NASD) 30/360
			case DayCountBasis.Actual360:         // 2=exact/360
			case DayCountBasis.Europ30_360:         // 4=Europe 30/360
				return new cNumber(360);
			case DayCountBasis.ActualActual: {         // 1=exact/exact
				var d = cDate.prototype.getDateFromExcel(date);
				return new cNumber(d.isLeapYear() ? 366 : 365);
			}
			case DayCountBasis.Actual365:         //3=exact/365
				return new cNumber(365);
			default:
				return new cError(cErrorType.not_numeric);
		}
	}

	function getCorrectDate(val) {
		if (!AscCommon.bDate1904) {
			if (val < 60) {
				val = new cDate((val - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
			} else if (val == 60) {
				val = new cDate((val - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay);
			} else {
				val = new cDate((val - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay);
			}
		} else {
			val = new cDate((val - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
		}
		return val;
	}

	function getCorrectDate2(val) {
		if (!AscCommon.bDate1904) {
			val = new cDate((val - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay);
		} else {
			val = new cDate((val - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
		}
		return val;
	}


	function getWeekends(val) {
		var res = [];
		if (val) {
			if (cElementType.number === val.type) {
				//0 - SUNDAY, 1 - MONDAY, 2 - TUESDAY, 3 - WEDNESDAY, 4 - THURSDAY, 5 - FRIDAY, 6 - SATURDAY
				var numberVal = val.getValue();
				switch (numberVal) {
					case 1 :
						res[6] = true;
						res[0] = true;
						break;
					case 2 :
						res[0] = true;
						res[1] = true;
						break;
					case 3 :
						res[1] = true;
						res[2] = true;
						break;
					case 4 :
						res[2] = true;
						res[3] = true;
						break;
					case 5 :
						res[3] = true;
						res[4] = true;
						break;
					case 6 :
						res[4] = true;
						res[5] = true;
						break;
					case 7 :
						res[5] = true;
						res[6] = true;
						break;

					case 11 :
						res[0] = true;
						break;
					case 12 :
						res[1] = true;
						break;
					case 13 :
						res[2] = true;
						break;
					case 14 :
						res[3] = true;
						break;
					case 15 :
						res[4] = true;
						break;
					case 16 :
						res[5] = true;
						break;
					case 17 :
						res[6] = true;
						break;

					default :
						return new cError(cErrorType.not_numeric);
				}
			} else if (cElementType.string === val.type) {
				var stringVal = val.getValue();
				if (stringVal.length !== 7) {
					return new cError(cErrorType.wrong_value_type);
				}
				//start with monday
				for (var i = 0; i < 7; i++) {
					var num = 6 === i ? 0 : i + 1;
					switch (stringVal[i]) {
						case '0' :
							res[num] = false;
							break;
						case '1' :
							res[num] = true;
							break;
						default  :
							return new cError(cErrorType.wrong_value_type);
					}
				}
			} else {
				return new cError(cErrorType.not_numeric);
			}
		} else {
			res[6] = true;
			res[0] = true;
		}

		return res;
	}

	function getHolidays(val) {
		var holidays = [];
		if (val) {
			if (val instanceof cRef) {
				var a = val.getValue();
				if (a instanceof cNumber && a.getValue() >= 0) {
					holidays.push(a);
				}
			} else if (val instanceof cArea || val instanceof cArea3D) {
				var arr = val.getValue();
				for (var i = 0; i < arr.length; i++) {
					if (arr[i] instanceof cNumber && arr[i].getValue() >= 0) {
						holidays.push(arr[i]);
					}
				}
			} else if (val instanceof cArray) {
				var bIsError = false;

				val.foreach(function (elem) {
					if (elem instanceof cNumber) {
						holidays.push(elem);
					} else if (elem instanceof cString) {
						var res = elem.tocNumber();
						if (res instanceof cError || res instanceof cEmpty) {
							var d = new cDate(elem.getValue());
							if (isNaN(d)) {
								d = g_oFormatParser.parseDate(elem.getValue(), AscCommon.g_oDefaultCultureInfo);
								if (d === null) {
									return new cError(cErrorType.wrong_value_type);
								}
								res = d.value;
							} else {
								res = (d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay +
									(AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1));
							}
						} else {
							res = elem.tocNumber().getValue();
						}

						if (res && res > 0) {
							holidays.push(new cNumber(parseInt(res)));
						} else {
							return bIsError = new cError(cErrorType.wrong_value_type);
						}
					}
				});

				if (bIsError) {
					return bIsError;
				}
			} else {
				val = val.tocNumber();
				if (val instanceof cError) {
					return bIsError = new cError(val);
				} else if (val instanceof cNumber) {
					holidays.push(val);
				}
			}
		}

		for (var i = 0; i < holidays.length; i++) {
			holidays[i] = cDate.prototype.getDateFromExcel(holidays[i].getValue());
		}

		return holidays;
	}

	function _includeInHolidays(date, holidays) {
		for (var i = 0; i < holidays.length; i++) {
			if (date.getTime() == holidays[i].getTime()) {
				return true;
			}
		}
		return false;
	}

	function _includeInHolidaysTime(time, holidays) {
		for (var i = 0; i < holidays.length; i++) {
			if (time == holidays[i].getTime()) {
				return true;
			}
		}
		return false;
	}

	function weekNumber(dt, iso, type) {
		if (undefined === iso) {
			iso = [0, 1, 2, 3, 4, 5, 6];
		}
		if (undefined === type) {
			type = 0;
		}

		dt.setUTCHours(0, 0, 0);
		var utcFullYear = dt.getUTCFullYear();
		if (type) {
			//если это исключение не обработать, то будет бесконечная рекурсия
			utcFullYear = Date.prototype.getUTCFullYear.call(dt);
		}
		var startOfYear = new cDate(Date.UTC(utcFullYear, 0, 1));
		var endOfYear = new cDate(dt);
		endOfYear.setUTCMonth(11);
		endOfYear.setUTCDate(31);
		var wk = parseInt(((dt - startOfYear) / c_msPerDay + iso[startOfYear.getUTCDay()]) / 7);
		if (type) {
			switch (wk) {
				case 0:
					// Возвращаем номер недели от 31 декабря предыдущего года
					startOfYear.setUTCDate(0);
					return weekNumber(startOfYear, iso, type);
				case 53:
					// Если 31 декабря выпадает до четверга 1 недели следующего года
					if (endOfYear.getUTCDay() < 4) {
						return new cNumber(1);
					} else {
						return new cNumber(wk);
					}
				default:
					return new cNumber(wk);
			}
		} else {
			wk = parseInt(((dt - startOfYear) / c_msPerDay + iso[startOfYear.getUTCDay()] + 7) / 7);
			return new cNumber(wk);
		}
	}

	/**
	 * Returns the last day in the month in Date format.
	 * @param {number} nExcelDateVal
	 * @param {number} nMonths - The number of months before or after nExcelDateVal.
	 * A positive value for months yields a future date; a negative value yields a past date.
	 * @returns {cDate}
	 */
	function getLastDayInMonth (nExcelDateVal, nMonths) {
		let dtValue;
		if (!AscCommon.bDate1904) {
			if (nExcelDateVal === 0) {
				dtValue = new cDate((nExcelDateVal - AscCommonExcel.c_DateCorrectConst + 1) * c_msPerDay);
			} else if (nExcelDateVal < 60) {
				dtValue = new cDate((nExcelDateVal - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
			} else if (nExcelDateVal === 60) {
				dtValue = new cDate((nExcelDateVal - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay);
			} else {
				dtValue = new cDate((nExcelDateVal - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay);
			}
		} else {
			dtValue = new cDate((nExcelDateVal - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
		}
		// setUTCDate doesn't work properly for -UTC time
		dtValue.setUTCDate(1);
		dtValue.setUTCMonth(dtValue.getUTCMonth() + nMonths);
		dtValue.setUTCDate(dtValue.getDaysInMonth());

		return dtValue;
	}

	cFormulaFunctionGroup['DateAndTime'] = cFormulaFunctionGroup['DateAndTime'] || [];
	cFormulaFunctionGroup['DateAndTime'].push(cDATE, cDATEDIF, cDATEVALUE, cDAY, cDAYS, cDAYS360, cEDATE, cEOMONTH,
		cHOUR, cISOWEEKNUM, cMINUTE, cMONTH, cNETWORKDAYS, cNETWORKDAYS_INTL, cNOW, cSECOND, cTIME, cTIMEVALUE, cTODAY,
		cWEEKDAY, cWEEKNUM, cWORKDAY, cWORKDAY_INTL, cYEAR, cYEARFRAC);

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDATE() {
	}

	//***array-formula***
	cDATE.prototype = Object.create(cBaseFunction.prototype);
	cDATE.prototype.constructor = cDATE;
	cDATE.prototype.name = 'DATE';
	cDATE.prototype.argumentsMin = 3;
	cDATE.prototype.argumentsMax = 3;
	cDATE.prototype.argumentsType = [argType.number, argType.number, argType.number];
	cDATE.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2], year, month, day;

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElement(0);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElement(0);
		}
		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]);
		} else if (arg2 instanceof cArray) {
			arg2 = arg2.getElement(0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		arg2 = arg2.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
		}

		year = arg0.getValue();
		month = arg1.getValue();
		day = arg2.getValue();

		if (year >= 0 && year <= 1899) {
			year += 1900;
		}

		var res;
		if (year === 1900 && month === 2 && day === 29) {
			res = new cNumber(60);
		} else {
			var _num = Math.round(new cDate(Date.UTC(year, month - 1, day)).getExcelDate());
			if (_num < 0) {
				return new cError(cErrorType.not_numeric);
			}
			res = new cNumber(_num);
		}
		res.numFormat = 14;
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDATEDIF() {
	}

	//***array-formula***
	cDATEDIF.prototype = Object.create(cBaseFunction.prototype);
	cDATEDIF.prototype.constructor = cDATEDIF;
	cDATEDIF.prototype.name = 'DATEDIF';
	cDATEDIF.prototype.argumentsMin = 3;
	cDATEDIF.prototype.argumentsMax = 3;
	cDATEDIF.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDATEDIF.prototype.argumentsType = [argType.number, argType.number, argType.text];
	cDATEDIF.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1], arg2 = arg[2];

		// --------------- check first argument[arg0] ----------------//
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.array === arg0.type) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		// --------------- check second argument[arg1] ----------------//
		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.array === arg1.type) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		// --------------- check third argument[arg2] ----------------//
		if (cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type) {
			arg2 = arg2.cross(arguments[1]);
		} else if (cElementType.array === arg2.type) {
			arg2 = arg2.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		arg2 = arg2.tocString();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}
		if (cElementType.error === arg2.type) {
			return arg2;
		}

		let val0 = arg0.getValue(), val1 = arg1.getValue();

		if (val0 < 0 || val1 < 0 || val0 > val1) {
			return new cError(cErrorType.not_numeric);
		}

		// TODO при передаче в аргументы числовых значений в единственном числе, итоговая дата различается с той что в ms на 1 день, из-за этого результаты могут различатся
		val0 = cDate.prototype.getDateFromExcel(val0);
		val1 = cDate.prototype.getDateFromExcel(val1);

		function newDateDiff(date1, date2) {
			const results = {
				years: date2.getUTCFullYear() - date1.getUTCFullYear(),
				months: date2.getUTCMonth() - date1.getUTCMonth(),
				days: date2.getUTCDate() - date1.getUTCDate(),
				hours: date2.getUTCHours() - date1.getUTCHours()
			}

			if (results.hours < 0) {
				results.days--;
				results.hours += 24;
			}
			if (results.days < 0) {
				let copy1 = new Date(date1.getTime());
				copy1.setDate(32);
				results.months--;
				results.days = 32 - date1.getUTCDate() - copy1.getUTCDate() + date2.getUTCDate();
			}
			if (results.months < 0) {
				results.years--;
				results.months += 12;
			}

			let years = results.years;
			let months = results.years * 12 + results.months;
			let days = results.days;

			return [years, months, days];
		}

		// function dateDiff(date1, date2) {
		// 	var years = date2.getUTCFullYear() - date1.getUTCFullYear();
		// 	var months = years * 12 + date2.getUTCMonth() - date1.getUTCMonth();
		// 	var days = date2.getUTCDate() - date1.getUTCDate();
		// 	var daysDiff = (date2 - date1) / c_msPerDay; // ms -> s -> m -> h -> d
		// 	years -= date2.getUTCMonth() < date1.getUTCMonth();
		// 	months -= date2.getUTCDate() < date1.getUTCDate();
		// 	days += days < 0 ? new cDate(Date.UTC(date2.getUTCFullYear(), date2.getUTCMonth() - 1, 0)).getUTCDate() + 1 : 0;
		// 	if(years === 1 && daysDiff < 365) {
		// 		years = 0;
		// 	}

		// 	return [years, months, days];
		// }

		switch (arg2.getValue().toUpperCase()) {
			case "Y":
				return new cNumber(newDateDiff(val0, val1)[0]);
				break;
			case "M":
				return new cNumber(newDateDiff(val0, val1)[1]);
				break;
			case "D":
				return new cNumber(parseInt((val1 - val0) / c_msPerDay));
				break;
			case "MD":
				if (val0.getUTCDate() > val1.getUTCDate()) {
					return new cNumber(Math.abs(
						new cDate(Date.UTC(val0.getUTCFullYear(), val0.getUTCMonth(), val0.getUTCDate())) -
						new cDate(Date.UTC(val0.getUTCFullYear(), val0.getUTCMonth() + 1, val1.getUTCDate()))) /
						c_msPerDay);
				} else {
					return new cNumber(val1.getUTCDate() - val0.getUTCDate());
				}
				break;
			case "YM":
				let d = newDateDiff(val0, val1);
				return new cNumber(d[1] - d[0] * 12);
				break;
			case "YD":
				if (val0.getUTCMonth() > val1.getUTCMonth()) {
					return new cNumber(Math.abs(
						new cDate(Date.UTC(val0.getUTCFullYear(), val0.getUTCMonth(), val0.getUTCDate())) -
						new cDate(Date.UTC(val0.getUTCFullYear() + 1, val1.getUTCMonth(), val1.getUTCDate()))) /
						c_msPerDay);
				} else {
					return new cNumber(Math.abs(
						new cDate(Date.UTC(val0.getUTCFullYear(), val0.getUTCMonth(), val0.getUTCDate())) -
						new cDate(Date.UTC(val0.getUTCFullYear(), val1.getUTCMonth(), val1.getUTCDate()))) /
						c_msPerDay);
				}
				break;
			default:
				return new cError(cErrorType.not_numeric)
		}

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDATEVALUE() {
	}

	//***array-formula***
	cDATEVALUE.prototype = Object.create(cBaseFunction.prototype);
	cDATEVALUE.prototype.constructor = cDATEVALUE;
	cDATEVALUE.prototype.name = 'DATEVALUE';
	cDATEVALUE.prototype.argumentsMin = 1;
	cDATEVALUE.prototype.argumentsMax = 1;
	cDATEVALUE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDATEVALUE.prototype.argumentsType = [argType.number];
	cDATEVALUE.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocString();

		if (arg0 instanceof cError) {
			return arg0;
		}

		var res = g_oFormatParser.parse(arg0.getValue());
		if (res && res.bDateTime) {
			return new cNumber(parseInt(res.value));
		} else {
			return new cError(cErrorType.wrong_value_type);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDAY() {
	}

	//***array-formula***
	cDAY.prototype = Object.create(cBaseFunction.prototype);
	cDAY.prototype.constructor = cDAY;
	cDAY.prototype.name = 'DAY';
	cDAY.prototype.argumentsMin = 1;
	cDAY.prototype.argumentsMax = 1;
	cDAY.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDAY.prototype.argumentsType = [argType.number];
	cDAY.prototype.Calculate = function (arg) {
		var t = this;
		var bIsSpecialFunction = arguments[4];

		var calculateFunc = function (curArg) {
			var val;
			if (curArg instanceof cArray) {
				curArg = curArg.getElement(0);
			} else if (curArg instanceof cArea || curArg instanceof cArea3D) {
				curArg = curArg.cross(arguments[1]).tocNumber();
				val = curArg.tocNumber().getValue();
			}
			if (curArg instanceof cError) {
				return curArg;
			} else if (curArg instanceof cNumber || curArg instanceof cBool || curArg instanceof cEmpty) {
				val = curArg.tocNumber().getValue();
			} else if (curArg instanceof cRef || curArg instanceof cRef3D) {
				val = curArg.getValue().tocNumber();
				if (val instanceof cNumber || val instanceof cBool || curArg instanceof cEmpty) {
					val = curArg.tocNumber().getValue();
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			} else if (curArg instanceof cString) {
				val = curArg.tocNumber();
				if (val instanceof cError || val instanceof cEmpty) {
					var d = new cDate(curArg.getValue());
					if (isNaN(d)) {
						return new cError(cErrorType.wrong_value_type);
					} else {
						val = Math.floor((d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay +
							(AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1)));
					}
				} else {
					val = curArg.tocNumber().getValue();
				}
			}
			if (val < 0) {
				return new cError(cErrorType.not_numeric);
			} else if (!AscCommon.bDate1904) {
				if (val < 60) {
					return t.setCalcValue(
						new cNumber((new cDate((val - AscCommonExcel.c_DateCorrectConst) * c_msPerDay)).getUTCDate()), 0);
				} else if (val == 60) {
					return t.setCalcValue(new cNumber(
						(new cDate((val - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay)).getUTCDate() + 1), 0);
				} else {
					return t.setCalcValue(
						new cNumber((new cDate((val - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay)).getUTCDate()),
						0);
				}
			} else {
				return t.setCalcValue(
					new cNumber((new cDate((val - AscCommonExcel.c_DateCorrectConst) * c_msPerDay)).getUTCDate()), 0);
			}
		};

		var arg0 = arg[0], res;
		if (!bIsSpecialFunction) {

			if (arg0 instanceof cArray) {
				arg0 = arg0.getElement(0);
			} else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
				arg0 = arg0.cross(arguments[1]).tocNumber();
			}
			res = calculateFunc(arg0);
		} else {
			res = this._checkArrayArguments(arg0, calculateFunc);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDAYS() {
	}

	//***array-formula***
	cDAYS.prototype = Object.create(cBaseFunction.prototype);
	cDAYS.prototype.constructor = cDAYS;
	cDAYS.prototype.name = 'DAYS';
	cDAYS.prototype.argumentsMin = 2;
	cDAYS.prototype.argumentsMax = 2;
	cDAYS.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDAYS.prototype.isXLFN = true;
	cDAYS.prototype.argumentsType = [argType.number, argType.number];
	cDAYS.prototype.Calculate = function (arg) {
		let oArguments = this._prepareArguments(arg, arguments[1], true);
		let argClone = oArguments.args;

		const calulateDay = function (curArg) {
			let val;
			if (curArg.type === cElementType.cell || curArg.type === cElementType.cell3D) {
				curArg = curArg.getValue();
			} else if (curArg.type === cElementType.array) {
				curArg = curArg.getElement(0);
			} else if (curArg.type === cElementType.cellsRange || curArg.type === cElementType.cellsRange3D) {
				curArg = curArg.cross(arguments[1]);
			}

			val = curArg ? curArg : new cNumber(0);

			if (val.type === cElementType.error) {
				return val;
			} else if (val.type === cElementType.number || val.type === cElementType.bool || val.type === cElementType.empty) {
				val = val.tocNumber().getValue();
			} else if (val.type === cElementType.string) {
				val = val.tocNumber();
				if (val.type === cElementType.error || val.type === cElementType.empty) {
					let d = new cDate(val.getValue());
					if (isNaN(d)) {
						return new cError(cErrorType.wrong_value_type);
					} else {
						val = Math.floor((d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay +
							(AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1)));
					}
				} else {
					val = val.tocNumber().getValue();
				}
			}

			return val;
		};

		const calulateDays = function (argArray) {
			let end_date = calulateDay(argArray[0]);
			let start_date = calulateDay(argArray[1]);
			if (end_date.type === cElementType.error) {
				return end_date;
			} else if (start_date.type === cElementType.error) {
				return start_date;
			}

			if (end_date < 0 || start_date < 0) {
				return new cError(cErrorType.not_numeric);
			} else {
				let endInteger = Math.floor(end_date);
				let startInteger = Math.floor(start_date);
				return new cNumber(endInteger - startInteger);

			}
		};

		return calulateDays(argClone);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDAYS360() {
	}

	//***array-formula***
	cDAYS360.prototype = Object.create(cBaseFunction.prototype);
	cDAYS360.prototype.constructor = cDAYS360;
	cDAYS360.prototype.name = 'DAYS360';
	cDAYS360.prototype.argumentsMin = 2;
	cDAYS360.prototype.argumentsMax = 3;
	cDAYS360.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDAYS360.prototype.argumentsType = [argType.number, argType.number, argType.logical];
	cDAYS360.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2] ? arg[2] : new cBool(false);

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]);
		} else if (arg2 instanceof cArray) {
			arg2 = arg2.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		arg2 = arg2.tocBool();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
		}

		if (arg0.getValue() < 0) {
			return new cError(cErrorType.not_numeric);
		}
		if (arg1.getValue() < 0) {
			return new cError(cErrorType.not_numeric);
		}

		var date1 = cDate.prototype.getDateFromExcel(arg0.getValue()),
			date2 = cDate.prototype.getDateFromExcel(arg1.getValue());

		return new cNumber(days360(date1, date2, arg2.toBool()));

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cEDATE() {
	}

	//***array-formula***
	cEDATE.prototype = Object.create(cBaseFunction.prototype);
	cEDATE.prototype.constructor = cEDATE;
	cEDATE.prototype.name = 'EDATE';
	cEDATE.prototype.argumentsMin = 2;
	cEDATE.prototype.argumentsMax = 2;
	cEDATE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cEDATE.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cEDATE.prototype.argumentsType = [argType.any, argType.any];
	cEDATE.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1];

		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			let dimensions = arg0.getDimensions();
			if (dimensions.row > 1 || dimensions.col > 1) {
				return new cError(cErrorType.wrong_value_type);
			}
			arg0 = arg0.getFirstElement();
		} else if (arg0.type === cElementType.array) {
			arg0 = arg0.getElementRowCol(0, 0);
		} else if (arg0.type === cElementType.bool) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg1.type === cElementType.cellsRange || arg1.type === cElementType.cellsRange3D) {
			let dimensions = arg1.getDimensions();
			if (dimensions.row > 1 || dimensions.col > 1) {
				return new cError(cErrorType.wrong_value_type);
			}
			arg1 = arg1.getFirstElement();
		} else if (arg1.type === cElementType.array) {
			arg1 = arg1.getElementRowCol(0, 0);
		} else if (arg1.type === cElementType.bool) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg0.type === cElementType.empty || arg1.type === cElementType.empty) {
			return new cError(cErrorType.not_available);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0.type === cElementType.error) {
			return arg0;
		}
		if (arg1.type === cElementType.error) {
			return arg1;
		}

		let val = arg0.getValue(), addedMonths = arg1.getValue(), date, _date;
		let now = new cDate();

		let offsetInMinutes = now.getTimezoneOffset();
		let offsetInHours = (offsetInMinutes / 60) * -1;
		let UTCshift = offsetInHours >= 0 ? 1 : 0;

		if (val < 0) {
			return new cError(cErrorType.not_numeric);
		} else if (addedMonths === 0) {
			return arg0;
		} else if (!AscCommon.bDate1904) {
			if (val < 60) {
				val = new cDate((Math.floor(val) - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
			} else if (val === 60) {
				// 29 february 1900 - exception - this date doesn't exist in js
				// val = new cDate(Date.UTC( 1900, 1, 29 ));
				val = new cDate((Math.floor(val) - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
			} else {
				val = new cDate((Math.floor(val) - AscCommonExcel.c_DateCorrectConst - UTCshift) * c_msPerDay);
			}
		} else {
			val = new cDate((Math.floor(val) - AscCommonExcel.c_DateCorrectConst) * c_msPerDay);
		}

		date = new cDate(val);

		let localDay = date.getDate(),
			localMonth = date.getMonth(),
			isLeap = date.isLeapYear();

		if (0 <= localDay && 28 >= localDay) {
			val = new cDate(val.setUTCMonth(val.getUTCMonth() + addedMonths))
		} else if (29 <= localDay && 31 >= localDay) {
			// Add the month separately. If the target month is February, the year is a leap year, and the day is greater than 28, return 28 or 29.
			// It should not depend on UTC since we are working with an isolated date.
			let newDay = localDay;
			if (localMonth + addedMonths == 1) {
				// set day to the last day of february
				newDay = isLeap ? 29 : 28;
			}
			date.setDate(1);
			date.setMonth(localMonth + addedMonths);

			// get and compare last day of month
			if (newDay > (_date = date.getLocalDaysInMonth())) {
				date.setDate(_date)
			} else {
				date.setDate(newDay);
			}

			val = new cDate(date);
		}

		let res = Math.floor((val.getTime() / 1000 - val.getTimezoneOffset() * 60) / c_sPerDay + (AscCommonExcel.c_DateCorrectConst + 1));
		if (res < 0) {
			return new cError(cErrorType.not_numeric);
		}
		// shift for 1904 mode
		if (AscCommon.bDate1904) {
			res -= 1;
		}

		return new cNumber(res);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cEOMONTH() {
	}

	//***array-formula***
	cEOMONTH.prototype = Object.create(cBaseFunction.prototype);
	cEOMONTH.prototype.constructor = cEOMONTH;
	cEOMONTH.prototype.name = 'EOMONTH';
	cEOMONTH.prototype.argumentsMin = 2;
	cEOMONTH.prototype.argumentsMax = 2;
	cEOMONTH.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cEOMONTH.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cEOMONTH.prototype.argumentsType = [argType.any, argType.any];
	cEOMONTH.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1];
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		} else if (cElementType.array === arg0.type) {
			arg0 = arg0.getElementRowCol(0, 0);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}

		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (cElementType.array === arg1.type) {
			arg1 = arg1.getElementRowCol(0, 0);
		} else if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}

		if (cElementType.bool === arg0.type) {
			return new cError(cErrorType.wrong_value_type);
		}
		if (cElementType.bool === arg1.type) {
			return new cError(cErrorType.wrong_value_type);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		let val = arg0.getValue();
		val = parseInt(val);
		if (val < 0) {
			return new cError(cErrorType.not_numeric);
		}
		val = getLastDayInMonth(val, arg1.getValue());

		return new cNumber(Math.round(val.getExcelDateWithTime()));
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cHOUR() {
	}

	//***array-formula***
	cHOUR.prototype = Object.create(cBaseFunction.prototype);
	cHOUR.prototype.constructor = cHOUR;
	cHOUR.prototype.name = 'HOUR';
	cHOUR.prototype.argumentsMin = 1;
	cHOUR.prototype.argumentsMax = 1;
	cHOUR.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cHOUR.prototype.argumentsType = [argType.number];
	cHOUR.prototype.Calculate = function (arg) {
		var t = this;
		var bIsSpecialFunction = arguments[4];

		var calculateFunc = function (curArg) {
			var val;
			if (curArg instanceof cArray) {
				curArg = curArg.getElement(0);
			} else if (curArg instanceof cArea || curArg instanceof cArea3D) {
				curArg = curArg.cross(arguments[1]).tocNumber();
			}

			if (curArg instanceof cError) {
				return curArg;
			} else if (curArg instanceof cNumber || curArg instanceof cBool || curArg instanceof cEmpty) {
				val = curArg.tocNumber().getValue();
			} else if (curArg instanceof cRef || curArg instanceof cRef3D) {
				val = curArg.getValue();
				if (val instanceof cError) {
					return val;
				} else if (val instanceof cNumber || val instanceof cBool || curArg instanceof cEmpty) {
					val = curArg.tocNumber().getValue();
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			} else if (curArg instanceof cString) {
				val = curArg.tocNumber();
				if (val instanceof cError || val instanceof cEmpty) {
					var d = new cDate(curArg.getValue());
					if (isNaN(d)) {
						d = g_oFormatParser.parseDate(curArg.getValue(), AscCommon.g_oDefaultCultureInfo);
						if (d == null) {
							return new cError(cErrorType.wrong_value_type);
						}
						val = d.value;
					} else {
						val = (d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay + (AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1));
					}
				} else {
					val = curArg.tocNumber().getValue();
				}
			}
			if (val < 0) {
				return new cError(cErrorType.not_numeric);
			} else                             //1		 2 3 4					   4	3		 	 					2 1
			{
				return t.setCalcValue(new cNumber(parseInt(((val - Math.floor(val)) * 24).toFixed(cExcelDateTimeDigits))), 0);
			}
		};

		var arg0 = arg[0], res;
		if (!bIsSpecialFunction) {

			if (arg0 instanceof cArray) {
				arg0 = arg0.getElement(0);
			} else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
				arg0 = arg0.cross(arguments[1]).tocNumber();
			}
			res = calculateFunc(arg0);
		} else {
			res = this._checkArrayArguments(arg0, calculateFunc);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cISOWEEKNUM() {
	}

	//***array-formula***
	cISOWEEKNUM.prototype = Object.create(cBaseFunction.prototype);
	cISOWEEKNUM.prototype.constructor = cISOWEEKNUM;
	cISOWEEKNUM.prototype.name = 'ISOWEEKNUM';
	cISOWEEKNUM.prototype.argumentsMin = 1;
	cISOWEEKNUM.prototype.argumentsMax = 1;
	cISOWEEKNUM.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cISOWEEKNUM.prototype.isXLFN = true;
	cISOWEEKNUM.prototype.argumentsType = [argType.number];
	cISOWEEKNUM.prototype.Calculate = function (arg) {
		let oArguments = this._prepareArguments(arg, arguments[1], true);
		let argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();

		let argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		let arg0 = argClone[0];
		if (arg0.getValue() < 0) {
			return new cError(cErrorType.not_numeric);
		}

		let weekdayStartDay = [6, 7, 8, 9, 10, 4, 5], type = 1;

		return new cNumber(weekNumber(cDate.prototype.getDateFromExcel(arg0.getValue()), weekdayStartDay, type));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMINUTE() {
	}

	//***array-formula***
	cMINUTE.prototype = Object.create(cBaseFunction.prototype);
	cMINUTE.prototype.constructor = cMINUTE;
	cMINUTE.prototype.name = 'MINUTE';
	cMINUTE.prototype.argumentsMin = 1;
	cMINUTE.prototype.argumentsMax = 1;
	cMINUTE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cMINUTE.prototype.argumentsType = [argType.number];
	cMINUTE.prototype.Calculate = function (arg) {
		var t = this;
		var bIsSpecialFunction = arguments[4];

		var calculateFunc = function (curArg) {
			var val;
			if (curArg instanceof cArray) {
				curArg = curArg.getElement(0);
			} else if (curArg instanceof cArea || curArg instanceof cArea3D) {
				curArg = curArg.cross(arguments[1]).tocNumber();
			}

			if (curArg instanceof cError) {
				return curArg;
			} else if (curArg instanceof cNumber || curArg instanceof cBool || curArg instanceof cEmpty) {
				val = curArg.tocNumber().getValue();
			} else if (curArg instanceof cRef || curArg instanceof cRef3D) {
				val = curArg.getValue();
				if (val instanceof cError) {
					return val;
				} else if (val instanceof cNumber || val instanceof cBool || curArg instanceof cEmpty) {
					val = curArg.tocNumber().getValue();
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			} else if (curArg instanceof cString) {
				val = curArg.tocNumber();
				if (val instanceof cError || val instanceof cEmpty) {
					var d = new cDate(curArg.getValue());
					if (isNaN(d)) {
						d = g_oFormatParser.parseDate(curArg.getValue(), AscCommon.g_oDefaultCultureInfo);
						if (d == null) {
							return new cError(cErrorType.wrong_value_type);
						}
						val = d.value;
					} else {
						val = (d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay + (AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1));
					}
				} else {
					val = curArg.tocNumber().getValue();
				}
			}
			if (val < 0) {
				return new cError(cErrorType.not_numeric);
			} else {
				//TODO исплользую функцию parseDate. по идее необходима только первая часть этой функции
				d = AscCommon.NumFormat.prototype.parseDate(val);
				val = d.min;

				return t.setCalcValue(new cNumber(val), 0);
			}
		};

		var arg0 = arg[0], res;
		if (!bIsSpecialFunction) {

			if (arg0 instanceof cArray) {
				arg0 = arg0.getElement(0);
			} else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
				arg0 = arg0.cross(arguments[1]).tocNumber();
			}
			res = calculateFunc(arg0);
		} else {
			res = this._checkArrayArguments(arg0, calculateFunc);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMONTH() {
	}

	//***array-formula***
	cMONTH.prototype = Object.create(cBaseFunction.prototype);
	cMONTH.prototype.constructor = cMONTH;
	cMONTH.prototype.name = 'MONTH';
	cMONTH.prototype.argumentsMin = 1;
	cMONTH.prototype.argumentsMax = 1;
	cMONTH.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cMONTH.prototype.argumentsType = [argType.number];
	cMONTH.prototype.Calculate = function (arg) {
		let t = this;
		let bIsSpecialFunction = arguments[4];

		let calculateFunc = function (curArg) {
			let val;

			if (curArg.type === cElementType.cell || curArg.type === cElementType.cell3D) {
				curArg = curArg.getValue();
			}

			if (curArg.type === cElementType.error) {
				return curArg;
			} else if (curArg.type === cElementType.number || curArg.type === cElementType.bool || curArg.type === cElementType.empty) {
				val = curArg.tocNumber().getValue();
			} else if (curArg.type === cElementType.string) {
				val = curArg.tocNumber();
				if (val.type === cElementType.error || val.type === cElementType.empty) {
					let d = new cDate(curArg.getValue());
					if (isNaN(d)) {
						return new cError(cErrorType.wrong_value_type);
					} else {
						val = Math.floor((d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay + (AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1)));
					}
				} else {
					val = curArg.tocNumber().getValue();
				}
			}
			if (val < 0) {
				return new cError(cErrorType.not_numeric);
			}
			if (!AscCommon.bDate1904) {
				if (val === 60) {
					return t.setCalcValue(new cNumber(2), 0);
				} else {
					return t.setCalcValue(new cNumber((new cDate(((val === 0 ? 1 : val) - AscCommonExcel.c_DateCorrectConst - 1) * c_msPerDay)).getUTCMonth() + 1), 0);
				}
			} else {
				return t.setCalcValue(new cNumber((new cDate(((val === 0 ? 1 : val) - AscCommonExcel.c_DateCorrectConst) * c_msPerDay)).getUTCMonth() + 1), 0);
			}
		};

		let arg0 = arg[0], res;
		if (!bIsSpecialFunction) {
			if (arg0.type === cElementType.array) {
				arg0 = arg0.getElement(0);
			} else if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
				arg0 = arg0.cross(arguments[1]).tocNumber();
			}
			res = calculateFunc(arg0);
		} else {
			res = this._checkArrayArguments(arg0, calculateFunc);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cNETWORKDAYS() {
	}

	//***array-formula***
	cNETWORKDAYS.prototype = Object.create(cBaseFunction.prototype);
	cNETWORKDAYS.prototype.constructor = cNETWORKDAYS;
	cNETWORKDAYS.prototype.name = 'NETWORKDAYS';
	cNETWORKDAYS.prototype.argumentsMin = 2;
	cNETWORKDAYS.prototype.argumentsMax = 3;
	cNETWORKDAYS.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cNETWORKDAYS.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cNETWORKDAYS.prototype.arrayIndexes = {2: 1};
	cNETWORKDAYS.prototype.argumentsType = [argType.any, argType.any, argType.any];
	cNETWORKDAYS.prototype.Calculate = function (arg) {
		let oArguments = this._prepareArguments([arg[0], arg[1]], arguments[1]);
		let argClone = oArguments.args;

		let arg0 = arg[0],
			arg1 = arg[1],
			arg2 = arg[2];

		// arg0 type check
		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		} else if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			if (arg0.isOneElement()) {
				arg0 = arg0.getFirstElement();
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		} else if (cElementType.array === arg0.type) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (cElementType.bool === arg0.type) {
			arg0 = new cError(cErrorType.wrong_value_type);
		} else if (cElementType.empty === arg0.type) {
			arg0 = new cNumber(0);
		}

		// arg1 type check
		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		} else if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			if (arg1.isOneElement()) {
				arg1 = arg1.getFirstElement();
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		} else if (cElementType.array === arg1.type) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (cElementType.bool === arg1.type) {
			arg1 = new cError(cErrorType.wrong_value_type);
		} else if (cElementType.empty === arg1.type) {
			arg1 = new cNumber(0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		if (cElementType.error === arg1.type) {
			return arg1;
		}

		let val0 = Math.trunc(arg0.getValue()), val1 = Math.trunc(arg1.getValue());

		if (val0 < 0 || val1 < 0) {
			return new cError(cErrorType.not_numeric);
		} else if (val0 === 0 && val1 === 0) {
			return new cNumber(0);
		}

		let argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		//function weekday also don't separate < 60
		val0 = getCorrectDate2(val0);
		val1 = getCorrectDate2(val1);

		//Holidays
		let holidays = getHolidays(arg2);
		if (holidays.type === cElementType.error) {
			return holidays;
		}

		const calcDate = function () {
			let count = 0;
			let start = val0;
			let end = val1;
			let dif = val1 - val0;
			if (dif < 0) {
				start = val1;
				end = val0;
			}

			let difAbs = (end - start);
			difAbs = (difAbs + (c_msPerDay)) / c_msPerDay;

			for (let i = 0; i < difAbs; i++) {
				let date = new cDate(start);
				date.setUTCDate(Date.prototype.getUTCDate.call(start) + i);
				if (date.getUTCDay() !== 6 && date.getUTCDay() !== 0 && !_includeInHolidays(date, holidays)) {
					count++;
				}
			}

			return new cNumber((dif < 0 ? -1 : 1) * count);
		};

		return calcDate();
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cNETWORKDAYS_INTL() {
	}

	//***array-formula***
	cNETWORKDAYS_INTL.prototype = Object.create(cBaseFunction.prototype);
	cNETWORKDAYS_INTL.prototype.constructor = cNETWORKDAYS_INTL;
	cNETWORKDAYS_INTL.prototype.name = 'NETWORKDAYS.INTL';
	cNETWORKDAYS_INTL.prototype.argumentsMin = 2;
	cNETWORKDAYS_INTL.prototype.argumentsMax = 4;
	cNETWORKDAYS_INTL.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cNETWORKDAYS_INTL.prototype.arrayIndexes = {2: 1, 3: 1};
	cNETWORKDAYS_INTL.prototype.argumentsType = [argType.any, argType.any, argType.number, argType.any];
	//TODO в данном случае есть различия с ms. при 3 и 4 аргументах - замена результата на ошибку не происходит.
	cNETWORKDAYS_INTL.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cNETWORKDAYS_INTL.prototype.Calculate = function (arg) {
		var tempArgs = arg[2] ? [arg[0], arg[1], arg[2]] : [arg[0], arg[1]];
		var oArguments = this._prepareArguments(tempArgs, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1].tocNumber();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		var arg0 = argClone[0], arg1 = argClone[1], arg2 = argClone[2], arg3 = arg[3];
		var val0 = arg0.getValue(), val1 = arg1.getValue();

		if (val0 < 0) {
			return new cError(cErrorType.not_numeric);
		}
		if (val1 < 0) {
			return new cError(cErrorType.not_numeric);
		}

		val0 = getCorrectDate(val0);
		val1 = getCorrectDate(val1);

		//Weekend
		var weekends = getWeekends(arg2);
		if (weekends instanceof cError) {
			return weekends;
		}

		//Holidays
		var holidays = getHolidays(arg3);
		if (holidays instanceof cError) {
			return holidays;
		}

		var calcDate = function () {
			var count = 0;
			var start = val0;
			var end = val1;
			var dif = val1 - val0;
			if (dif < 0) {
				start = val1;
				end = val0;
			}

			var difAbs = (end - start);
			difAbs = (difAbs + (c_msPerDay)) / c_msPerDay;

			var date = new cDate(start);
			var startTime = date.getTime();
			var startDay = date.getUTCDay();
			for (var i = 0; i < difAbs; i++) {
				if (!weekends[startDay] && !_includeInHolidaysTime(startTime, holidays)) {
					count++;
				}
				startDay = (startDay + 1) % 7;
				startTime += AscCommonExcel.c_msPerDay;

				// date.setTime(startTime + i * AscCommonExcel.c_msPerDay);
				//
				// if (!_includeInHolidays(date, holidays) && !weekends[date.getUTCDay()]) {
				// 	count++;
				// }
			}

			return new cNumber((dif < 0 ? -1 : 1) * count);
		};

		return calcDate();
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cNOW() {
	}

	//***array-formula***
	cNOW.prototype = Object.create(cBaseFunction.prototype);
	cNOW.prototype.constructor = cNOW;
	cNOW.prototype.name = 'NOW';
	cNOW.prototype.argumentsMax = 0;
	cNOW.prototype.ca = true;
	cNOW.prototype.argumentsType = null;
	cNOW.prototype.Calculate = function () {
		var d = new cDate();
		var res =
			new cNumber(d.getExcelDate(true) + (d.getHours() * 60 * 60 + d.getMinutes() * 60 + d.getSeconds()) / c_sPerDay);
		res.numFormat = 22;
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSECOND() {
	}

	//***array-formula***
	cSECOND.prototype = Object.create(cBaseFunction.prototype);
	cSECOND.prototype.constructor = cSECOND;
	cSECOND.prototype.name = 'SECOND';
	cSECOND.prototype.argumentsMin = 1;
	cSECOND.prototype.argumentsMax = 1;
	cSECOND.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cSECOND.prototype.argumentsType = [argType.number];
	cSECOND.prototype.Calculate = function (arg) {
		var t = this;
		var bIsSpecialFunction = arguments[4];

		var calculateFunc = function (curArg) {
			var val;
			if (curArg instanceof cArray) {
				curArg = curArg.getElement(0);
			} else if (curArg instanceof cArea || curArg instanceof cArea3D) {
				curArg = curArg.cross(arguments[1]).tocNumber();
			}

			if (curArg instanceof cError) {
				return curArg;
			} else if (curArg instanceof cNumber || curArg instanceof cBool || curArg instanceof cEmpty) {
				val = curArg.tocNumber().getValue();
			} else if (curArg instanceof cRef || curArg instanceof cRef3D) {
				val = curArg.getValue();
				if (val instanceof cError) {
					return val;
				} else if (val instanceof cNumber || val instanceof cBool || curArg instanceof cEmpty) {
					val = curArg.tocNumber().getValue();
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			} else if (curArg instanceof cString) {
				val = curArg.tocNumber();
				if (val instanceof cError || val instanceof cEmpty) {
					var d = new cDate(curArg.getValue());
					if (isNaN(d)) {
						d = g_oFormatParser.parseDate(curArg.getValue(), AscCommon.g_oDefaultCultureInfo);
						if (d == null) {
							return new cError(cErrorType.wrong_value_type);
						}
						val = d.value;
					} else {
						val = (d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay +
							(AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1));
					}
				} else {
					val = curArg.tocNumber().getValue();
				}
			}
			if (val < 0) {
				return new cError(cErrorType.not_numeric);
			} else {
				//val = parseInt((( val * 24 * 60 - Math.floor(val * 24 * 60) ) * 60).toFixed(cExcelDateTimeDigits)) % 60;
				val = Math.floor((val * c_sPerDay + 0.5).toFixed(cExcelDateTimeDigits)) % 60;
				return t.setCalcValue(new cNumber(val), 0);
			}
		};

		var arg0 = arg[0], res;
		if (!bIsSpecialFunction) {

			if (arg0 instanceof cArray) {
				arg0 = arg0.getElement(0);
			} else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
				arg0 = arg0.cross(arguments[1]).tocNumber();
			}
			res = calculateFunc(arg0);
		} else {
			res = this._checkArrayArguments(arg0, calculateFunc);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTIME() {
	}

	//***array-formula***
	cTIME.prototype = Object.create(cBaseFunction.prototype);
	cTIME.prototype.constructor = cTIME;
	cTIME.prototype.name = 'TIME';
	cTIME.prototype.argumentsMin = 3;
	cTIME.prototype.argumentsMax = 3;
	cTIME.prototype.argumentsType = [argType.number, argType.number, argType.number];
	cTIME.prototype.Calculate = function (arg) {
		var hour = arg[0], minute = arg[1], second = arg[2];

		if (hour instanceof cArea || hour instanceof cArea3D) {
			hour = hour.cross(arguments[1]);
		} else if (hour instanceof cArray) {
			hour = hour.getElement(0);
		}
		if (minute instanceof cArea || minute instanceof cArea3D) {
			minute = minute.cross(arguments[1]);
		} else if (minute instanceof cArray) {
			minute = minute.getElement(0);
		}
		if (second instanceof cArea || second instanceof cArea3D) {
			second = second.cross(arguments[1]);
		} else if (second instanceof cArray) {
			second = second.getElement(0);
		}

		hour = hour.tocNumber();
		minute = minute.tocNumber();
		second = second.tocNumber();

		if (hour instanceof cError) {
			return hour;
		}
		if (minute instanceof cError) {
			return minute;
		}
		if (second instanceof cError) {
			return second;
		}

		//LO - не округляет. ms - округляет.
		hour = parseInt(hour.getValue());
		minute = parseInt(minute.getValue());
		second = parseInt(second.getValue());

		var v = (hour * 60 * 60 + minute * 60 + second) / c_sPerDay;
		if (v < 0) {
			return new cError(cErrorType.not_numeric);
		}

		var res = new cNumber(v - Math.floor(v));
		res.numFormat = 18;
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTIMEVALUE() {
	}

	//***array-formula***
	cTIMEVALUE.prototype = Object.create(cBaseFunction.prototype);
	cTIMEVALUE.prototype.constructor = cTIMEVALUE;
	cTIMEVALUE.prototype.name = 'TIMEVALUE';
	cTIMEVALUE.prototype.argumentsMin = 1;
	cTIMEVALUE.prototype.argumentsMax = 1;
	cTIMEVALUE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTIMEVALUE.prototype.argumentsType = [argType.number];
	cTIMEVALUE.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocString();

		if (arg0 instanceof cError) {
			return arg0;
		}

		var res = g_oFormatParser.parse(arg0.getValue());

		if (res && res.bDateTime) {
			return new cNumber(res.value - parseInt(res.value));
		} else {
			return new cError(cErrorType.wrong_value_type);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTODAY() {
	}

	//***array-formula***
	cTODAY.prototype = Object.create(cBaseFunction.prototype);
	cTODAY.prototype.constructor = cTODAY;
	cTODAY.prototype.name = 'TODAY';
	cTODAY.prototype.argumentsMax = 0;
	cTODAY.prototype.ca = true;
	cTODAY.prototype.argumentsType = null;
	cTODAY.prototype.Calculate = function () {
		var res = new cNumber(new cDate().getExcelDate(true));
		res.numFormat = 14;
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWEEKDAY() {
	}

	//***array-formula***
	cWEEKDAY.prototype = Object.create(cBaseFunction.prototype);
	cWEEKDAY.prototype.constructor = cWEEKDAY;
	cWEEKDAY.prototype.name = 'WEEKDAY';
	cWEEKDAY.prototype.argumentsMin = 1;
	cWEEKDAY.prototype.argumentsMax = 2;
	cWEEKDAY.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cWEEKDAY.prototype.argumentsType = [argType.number, argType.number];
	cWEEKDAY.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(1);

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg1 instanceof cError) {
			return arg1;
		}

		var weekday;

		switch (arg1.getValue()) {
			case 1: /* 1 (Sunday) through 7 (Saturday)  */
			case 17:/* 1 (Sunday) through 7 (Saturday) */
				weekday = [1, 2, 3, 4, 5, 6, 7];
				break;
			case 2: /* 1 (Monday) through 7 (Sunday)    */
			case 11:/* 1 (Monday) through 7 (Sunday)    */
				weekday = [7, 1, 2, 3, 4, 5, 6];
				break;
			case 3: /* 0 (Monday) through 6 (Sunday)    */
				weekday = [6, 0, 1, 2, 3, 4, 5];
				break;
			case 12:/* 1 (Tuesday) through 7 (Monday)   */
				weekday = [6, 7, 1, 2, 3, 4, 5];
				break;
			case 13:/* 1 (Wednesday) through 7 (Tuesday) */
				weekday = [5, 6, 7, 1, 2, 3, 4];
				break;
			case 14:/* 1 (Thursday) through 7 (Wednesday) */
				weekday = [4, 5, 6, 7, 1, 2, 3];
				break;
			case 15:/* 1 (Friday) through 7 (Thursday) */
				weekday = [3, 4, 5, 6, 7, 1, 2];
				break;
			case 16:/* 1 (Saturday) through 7 (Friday) */
				weekday = [2, 3, 4, 5, 6, 7, 1];
				break;
			default:
				return new cError(cErrorType.not_numeric);
		}
		if (arg0.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}
		// ???
		return new cNumber(
			weekday[new cDate((arg0.getValue() - (AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1))) * c_msPerDay).getUTCDay()]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWEEKNUM() {
	}

	//***array-formula***
	cWEEKNUM.prototype = Object.create(cBaseFunction.prototype);
	cWEEKNUM.prototype.constructor = cWEEKNUM;
	cWEEKNUM.prototype.name = 'WEEKNUM';
	cWEEKNUM.prototype.argumentsMin = 1;
	cWEEKNUM.prototype.argumentsMax = 2;
	cWEEKNUM.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cWEEKNUM.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cWEEKNUM.prototype.argumentsType = [argType.any, argType.any];
	cWEEKNUM.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(1), type = 0;

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0.getValue() < 0) {
			return new cError(cErrorType.not_numeric);
		}

		var weekdayStartDay;

		switch (arg1.getValue()) {
			case 1: /* 1 (Sunday) through 7 (Saturday)  */
			case 17:/* 1 (Sunday) through 7 (Saturday) */
				weekdayStartDay = [0, 1, 2, 3, 4, 5, 6];
				break;
			case 2: /* 1 (Monday) through 7 (Sunday)    */
			case 11:/* 1 (Monday) through 7 (Sunday)    */
				weekdayStartDay = [6, 0, 1, 2, 3, 4, 5];
				break;
			case 12:/* 1 (Tuesday) through 7 (Monday)   */
				weekdayStartDay = [5, 6, 0, 1, 2, 3, 4];
				break;
			case 13:/* 1 (Wednesday) through 7 (Tuesday) */
				weekdayStartDay = [4, 5, 6, 0, 1, 2, 3];
				break;
			case 14:/* 1 (Thursday) through 7 (Wednesday) */
				weekdayStartDay = [3, 4, 5, 6, 0, 1, 2];
				break;
			case 15:/* 1 (Friday) through 7 (Thursday) */
				weekdayStartDay = [2, 3, 4, 5, 6, 0, 1];
				break;
			case 16:/* 1 (Saturday) through 7 (Friday) */
				weekdayStartDay = [1, 2, 3, 4, 5, 6, 0];
				break;
			case 21:
				weekdayStartDay = [6, 7, 8, 9, 10, 4, 5];
//                { 6, 7, 8, 9, 10, 4, 5 }
				type = 1;
				break;
			default:
				return new cError(cErrorType.not_numeric);
		}

		return new cNumber(weekNumber(cDate.prototype.getDateFromExcel(arg0.getValue()), weekdayStartDay, type));

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWORKDAY() {
	}

	//***array-formula***
	cWORKDAY.prototype = Object.create(cBaseFunction.prototype);
	cWORKDAY.prototype.constructor = cWORKDAY;
	cWORKDAY.prototype.name = 'WORKDAY';
	cWORKDAY.prototype.argumentsMin = 2;
	cWORKDAY.prototype.argumentsMax = 3;
	cWORKDAY.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cWORKDAY.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cWORKDAY.prototype.arrayIndexes = {2: 1};
	cWORKDAY.prototype.argumentsType = [argType.any, argType.any, argType.any];
	cWORKDAY.prototype.Calculate = function (arg) {
		var t = this;
		var oArguments = this._prepareArguments([arg[0], arg[1]], arguments[1], true);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1].tocNumber();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		var arg0 = argClone[0], arg1 = argClone[1], arg2 = arg[2];

		var val0 = arg0.getValue();
		if (val0 < 0) {
			return new cError(cErrorType.not_numeric);
		}
		val0 = getCorrectDate(val0);

		//Holidays
		var holidays = getHolidays(arg2);
		if (holidays instanceof cError) {
			return holidays;
		}

		var calcDate = function (startdate, workdaysCount, holidays) {
			workdaysCount = Math.floor(workdaysCount);
			var diff = Math.sign(workdaysCount);
			var daysCounter = 0;
			var currentDate = new cDate(startdate);

			while (daysCounter !== workdaysCount) {
				currentDate = new cDate(currentDate.getTime() + diff * c_msPerDay);

				//TODO поверить когда переходим через argVal0 = 60
				var dayOfWeek;
				var argVal0 = argClone && argClone[0] ? argClone[0].getValue() : null;
				if (argVal0 !== null && argVal0 < 60) {
					dayOfWeek = (currentDate.getUTCDay() > 0) ? currentDate.getUTCDay() - 1 : 6;
				} else {
					dayOfWeek = currentDate.getUTCDay();
				}

				if (dayOfWeek !== 0 && dayOfWeek !== 6 && !_includeInHolidays(currentDate, holidays)) {
					daysCounter += diff;
				}
			}

			return currentDate;
		};

		if (undefined === arg1.getValue()) {
			return new cError(cErrorType.not_numeric);
		}

		var date = calcDate(val0, arg1.getValue(), holidays);
		var val = date.getExcelDate();

		if (val < 0) {
			return new cError(cErrorType.not_numeric);
		}

		return t.setCalcValue(new cNumber(val), 14);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWORKDAY_INTL() {
	}

	//***array-formula***
	cWORKDAY_INTL.prototype = Object.create(cBaseFunction.prototype);
	cWORKDAY_INTL.prototype.constructor = cWORKDAY_INTL;
	cWORKDAY_INTL.prototype.name = 'WORKDAY.INTL';
	cWORKDAY_INTL.prototype.argumentsMin = 2;
	cWORKDAY_INTL.prototype.argumentsMax = 4;
	cWORKDAY_INTL.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cWORKDAY_INTL.prototype.arrayIndexes = {0: AscCommonExcel.arrayIndexesType.range, 1: AscCommonExcel.arrayIndexesType.range, 3: 1};
	cWORKDAY_INTL.prototype.argumentsType = [argType.any, argType.any, argType.number, argType.any];
	//TODO в данном случае есть различия с ms. при 3 и 4 аргументах - замена результата на ошибку не происходит.
	// cWORKDAY_INTL.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cWORKDAY_INTL.prototype.Calculate = function (arg) {
		//TODO проблема с формулами следующего типа - WORKDAY.INTL(8,60,"0000000")
		let t = this;
		let tempArgs = arg[2] ? [arg[0], arg[1], arg[2]] : [arg[0], arg[1]];
		let oArguments = this._prepareArguments(tempArgs, arguments[1]);
		let argClone = oArguments.args;

		if ((arg[0] && (arg[0].type === cElementType.cellsRange || arg[0].type === cElementType.cellsRange3D)) || 
			(arg[1] && (arg[1].type === cElementType.cellsRange || arg[1].type === cElementType.cellsRange3D))) {
			return new cError(cErrorType.wrong_value_type);
		}

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1].tocNumber();

		let argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		let arg0 = argClone[0], arg1 = argClone[1], arg2 = argClone[2], arg3 = arg[3];

		let val0 = arg0.getValue();
		if (val0 < 0) {
			return new cError(cErrorType.not_numeric);
		}
		val0 = val0 < 61 ? new cDate((val0 - (AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1))) * c_msPerDay) : getCorrectDate(val0);

		// Weekend
		if (arg2 && "1111111" === arg2.getValue()) {
			return new cError(cErrorType.wrong_value_type);
		}
		let weekends = getWeekends(arg2);
		if (weekends instanceof cError) {
			return weekends;
		}

		// Holidays
		let holidays = getHolidays(arg3);
		if (holidays instanceof cError) {
			return holidays;
		}

		const calcDate = function () {
			let dif = arg1.getValue(), count = 0, dif1 = dif > 0 ? 1 : dif < 0 ? -1 : 0, val, date = val0,
				isEndOfCycle = false;
			while (Math.abs(dif) > count) {
				date = new cDate(val0.getTime() + dif1 * c_msPerDay);
				if (!_includeInHolidays(date, holidays) && !weekends[date.getUTCDay()]) {
					count++;
				}
				// if last iteration:
				if (!(Math.abs(dif) > count)) {
					// check if it's next weekend. if found - add
					date = new cDate(val0.getTime() + dif1 * c_msPerDay);
					for (let i = 0; i < 7; i++) {
						if (weekends[date.getUTCDay()]) {
							dif >= 0 ? dif1++ : dif1--;
							date = new cDate(val0.getTime() + (dif1) * c_msPerDay);
						} else {
							isEndOfCycle = true;
							break;
						}
					}
				}
				!isEndOfCycle ? (dif >= 0 ? dif1++ : dif1--) : null;
			}
			date = new cDate(val0.getTime() + dif1 * c_msPerDay);
			val = date.getExcelDate();

			if (val < 0) {
				return new cError(cErrorType.not_numeric);
			}
			// shift for date less than 1/3/1900
			if (arg0.getValue() < 61) {
				val++;
			}

			return t.setCalcValue(new cNumber(val), 14);
		};

		return calcDate();
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cYEAR() {
	}

	//***array-formula***
	cYEAR.prototype = Object.create(cBaseFunction.prototype);
	cYEAR.prototype.constructor = cYEAR;
	cYEAR.prototype.name = 'YEAR';
	cYEAR.prototype.argumentsMin = 1;
	cYEAR.prototype.argumentsMax = 1;
	cYEAR.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cYEAR.prototype.argumentsType = [argType.number];
	cYEAR.prototype.Calculate = function (arg) {
		let t = this;
		let bIsSpecialFunction = arguments[4];

		let calculateFunc = function (curArg) {
			let val;
			if (curArg.type === cElementType.cell || curArg.type === cElementType.cell3D) {
				curArg = curArg.getValue();
			}

			if (curArg.type === cElementType.error) {
				return curArg;
			} else if (curArg.type === cElementType.number || curArg.type === cElementType.bool || curArg.type === cElementType.empty) {
				val = curArg.tocNumber().getValue();
			} else if (curArg.type === cElementType.cellsRange || curArg.type === cElementType.cellsRange3D) {
				return new cError(cErrorType.wrong_value_type);
			} else if (curArg.type === cElementType.string) {
				val = curArg.tocNumber();
				if (val.type === cElementType.error || val.type === cElementType.empty) {
					let d = new cDate(curArg.getValue());
					if (isNaN(d)) {
						return new cError(cErrorType.wrong_value_type);
					} else {
						val = Math.floor((d.getTime() / 1000 - d.getTimezoneOffset() * 60) / c_sPerDay + (AscCommonExcel.c_DateCorrectConst + (AscCommon.bDate1904 ? 0 : 1)));
					}
				} else {
					val = curArg.tocNumber().getValue();
				}
			}

			if (val < 0) {
				return t.setCalcValue(new cError(cErrorType.not_numeric), 0);
			} else {
				if (AscCommon.bDate1904 && val === 0) {
					val += 1;
				}
				return t.setCalcValue(new cNumber((new cDate((val - (AscCommonExcel.c_DateCorrectConst + 1)) * c_msPerDay)).getUTCFullYear()), 0);
			}
		};

		let arg0 = arg[0], res;
		if (!bIsSpecialFunction) {
			if (arg0.type === cElementType.array) {
				arg0 = arg0.getElement(0);
			} else if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
				arg0 = arg0.cross(arguments[1]).tocNumber();
			}
			res = calculateFunc(arg0);
		} else {
			res = this._checkArrayArguments(arg0, calculateFunc);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cYEARFRAC() {
	}

	//***array-formula***
	cYEARFRAC.prototype = Object.create(cBaseFunction.prototype);
	cYEARFRAC.prototype.constructor = cYEARFRAC;
	cYEARFRAC.prototype.name = 'YEARFRAC';
	cYEARFRAC.prototype.argumentsMin = 2;
	cYEARFRAC.prototype.argumentsMax = 3;
	cYEARFRAC.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cYEARFRAC.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cYEARFRAC.prototype.argumentsType = [argType.any, argType.any, argType.any];
	cYEARFRAC.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2] ? arg[2] : new cNumber(0);
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]);
		} else if (arg2 instanceof cArray) {
			arg2 = arg2.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		arg2 = arg2.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
		}

		var val0 = arg0.getValue(), val1 = arg1.getValue();

		if (val0 < 0 || val1 < 0) {
			return new cError(cErrorType.not_numeric);
		}

		val0 = cDate.prototype.getDateFromExcel(val0);
		val1 = cDate.prototype.getDateFromExcel(val1);

		return yearFrac(val0, val1, arg2.getValue());
//    return diffDate2( val0, val1, arg2.getValue() );

	};

//----------------------------------------------------------export----------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].DayCountBasis = DayCountBasis;
	window['AscCommonExcel'].yearFrac = yearFrac;
	window['AscCommonExcel'].diffDate = diffDate;
	window['AscCommonExcel'].days360 = days360;
	window['AscCommonExcel'].getCorrectDate = getCorrectDate;
	window['AscCommonExcel'].daysInYear = daysInYear;
	window['AscCommonExcel'].getLastDayInMonth = getLastDayInMonth;
})(window);
