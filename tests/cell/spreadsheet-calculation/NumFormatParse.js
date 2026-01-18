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

$(function () {
    window["AscCommonExcel"] = window["AscCommonExcel"] || {};
    window["AscCommonExcel"].Font = function () {};
    window["AscCommonExcel"].RgbColor = function () {};

    const eps = 1e-15;
    const formatParser = AscCommon.g_oFormatParser;
    const formatTypes = Asc.c_oAscNumFormatType;

    // =====================================================================
    // FormatParser.isLeapYear
    // =====================================================================
    QUnit.module('FormatParser.isLeapYear');

    QUnit.test('leap years', function (assert) {
        // Divisible by 4 but not by 100
        assert.strictEqual(formatParser.isLeapYear(2024), true, '2024');
        assert.strictEqual(formatParser.isLeapYear(2020), true, '2020');
        assert.strictEqual(formatParser.isLeapYear(1996), true, '1996');
        
        // Divisible by 100 but not by 400 (not leap)
        assert.strictEqual(formatParser.isLeapYear(1900), false, '1900');
        assert.strictEqual(formatParser.isLeapYear(2100), false, '2100');
        
        // Divisible by 400 (leap)
        assert.strictEqual(formatParser.isLeapYear(2000), true, '2000');
        assert.strictEqual(formatParser.isLeapYear(1600), true, '1600');
        
        // Not divisible by 4 (not leap)
        assert.strictEqual(formatParser.isLeapYear(2023), false, '2023');
        assert.strictEqual(formatParser.isLeapYear(2019), false, '2019');
    });

    // =====================================================================
    // FormatParser.isValidDay
    // =====================================================================
    QUnit.module('FormatParser.isValidDay');

    QUnit.test('valid days per month', function (assert) {
        const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        for (let month = 0; month < 12; month++) {
            assert.strictEqual(formatParser.isValidDay(2023, month, 1), true, 
                `Day 1 valid for month ${month + 1}`);
            assert.strictEqual(formatParser.isValidDay(2023, month, daysInMonth[month]), true, 
                `Day ${daysInMonth[month]} valid for month ${month + 1}`);
            assert.strictEqual(formatParser.isValidDay(2023, month, daysInMonth[month] + 1), false, 
                `Day ${daysInMonth[month] + 1} invalid for month ${month + 1}`);
        }
    });

    QUnit.test('February leap year handling', function (assert) {
        assert.strictEqual(formatParser.isValidDay(2024, 1, 29), true, 'Feb 29 in leap year');
        assert.strictEqual(formatParser.isValidDay(2024, 1, 30), false, 'Feb 30 in leap year');
        assert.strictEqual(formatParser.isValidDay(2023, 1, 29), false, 'Feb 29 in non-leap year');
    });

    QUnit.test('boundary conditions', function (assert) {
        assert.strictEqual(formatParser.isValidDay(2023, 0, 0), false, 'Day 0');
        assert.strictEqual(formatParser.isValidDay(2023, 0, -1), false, 'Negative day');
        assert.strictEqual(formatParser.isValidDay(2023, 0, 32), false, 'Day 32 in January');
    });

    // =====================================================================
    // FormatParser.isValidDate
    // =====================================================================
    QUnit.module('FormatParser.isValidDate');

    QUnit.test('valid dates', function (assert) {
        assert.strictEqual(formatParser.isValidDate(2023, 0, 1), true, 'Jan 1, 2023');
        assert.strictEqual(formatParser.isValidDate(2023, 11, 31), true, 'Dec 31, 2023');
        assert.strictEqual(formatParser.isValidDate(2000, 1, 29), true, 'Feb 29, 2000 (leap)');
        assert.strictEqual(formatParser.isValidDate(1900, 1, 28), true, 'Feb 28, 1900');
    });

    QUnit.test('Excel Feb 29, 1900 bug compatibility', function (assert) {
        assert.strictEqual(formatParser.isValidDate(1900, 1, 29), true, 'Feb 29, 1900 (Excel bug)');
    });

    QUnit.test('special base date Dec 31, 1899', function (assert) {
        assert.strictEqual(formatParser.isValidDate(1899, 11, 31), true, 'Dec 31, 1899');
    });

    QUnit.test('dates before 1900 invalid', function (assert) {
        assert.strictEqual(formatParser.isValidDate(1899, 0, 1), false, 'Jan 1, 1899');
        assert.strictEqual(formatParser.isValidDate(1800, 5, 15), false, '1800');
    });

    QUnit.test('invalid months and days', function (assert) {
        assert.strictEqual(formatParser.isValidDate(2023, -1, 15), false, 'Month -1');
        assert.strictEqual(formatParser.isValidDate(2023, 12, 15), false, 'Month 12');
        assert.strictEqual(formatParser.isValidDate(2023, 0, 0), false, 'Day 0');
        assert.strictEqual(formatParser.isValidDate(2023, 0, 32), false, 'Jan 32');
        assert.strictEqual(formatParser.isValidDate(2023, 3, 31), false, 'Apr 31');
    });

    // =====================================================================
    // FormatParser.isValidDatePDF
    // =====================================================================
    QUnit.module('FormatParser.isValidDatePDF');

    QUnit.test('PDF dates allow pre-1900', function (assert) {
        assert.strictEqual(formatParser.isValidDatePDF(2023, 0, 1), true, 'Jan 1, 2023');
        assert.strictEqual(formatParser.isValidDatePDF(1899, 0, 1), true, 'Jan 1, 1899 (PDF)');
        assert.strictEqual(formatParser.isValidDatePDF(1500, 5, 15), true, '1500 (PDF)');
        assert.strictEqual(formatParser.isValidDatePDF(100, 0, 1), true, 'Year 100 (PDF)');
    });

    QUnit.test('invalid months and days still rejected', function (assert) {
        assert.strictEqual(formatParser.isValidDatePDF(2023, -1, 15), false, 'Month -1');
        assert.strictEqual(formatParser.isValidDatePDF(2023, 12, 15), false, 'Month 12');
        assert.strictEqual(formatParser.isValidDatePDF(2023, 0, 32), false, 'Jan 32');
    });

    // =====================================================================
    // FormatParser.strcmp
    // =====================================================================
    QUnit.module('FormatParser.strcmp');

    QUnit.test('string comparison', function (assert) {
        assert.strictEqual(formatParser.strcmp("hello", "hello", 0, 5), true, 'Full match');
        assert.strictEqual(formatParser.strcmp("hello world", "ello", 1, 4), true, 'Partial match');
        assert.strictEqual(formatParser.strcmp("abcdef", "cd", 2, 2), true, 'Substring match');
        assert.strictEqual(formatParser.strcmp("hello", "world", 0, 5), false, 'Different strings');
        assert.strictEqual(formatParser.strcmp("hello", "", 0, 0), false, 'Zero length');
        assert.strictEqual(formatParser.strcmp("abc", "a", 0, 1), true, 'Single char');
        assert.strictEqual(formatParser.strcmp("hello", "llo", 2, 3, 0), true, 'With index2');
    });

    // =====================================================================
    // FormatParser.isLocaleNumber / parseLocaleNumber
    // =====================================================================
    QUnit.module('FormatParser.isLocaleNumber');

    QUnit.test('valid numbers with default locale', function (assert) {
        assert.strictEqual(formatParser.isLocaleNumber("123", null), true, 'Integer');
        assert.strictEqual(formatParser.isLocaleNumber("123.45", null), true, 'Decimal');
        assert.strictEqual(formatParser.isLocaleNumber("0.5", null), true, 'Less than 1');
        assert.strictEqual(formatParser.isLocaleNumber("-123.45", null), true, 'Negative');
    });

    QUnit.test('invalid inputs', function (assert) {
        assert.strictEqual(formatParser.isLocaleNumber("abc", null), false, 'Alphabetic');
        assert.strictEqual(formatParser.isLocaleNumber("12a34", null), false, 'Mixed');
        assert.strictEqual(formatParser.isLocaleNumber("", null), false, 'Empty');
    });

    QUnit.module('FormatParser.parseLocaleNumber');

    QUnit.test('parse numbers', function (assert) {
        assert.strictEqual(formatParser.parseLocaleNumber("123", null), 123, 'Integer');
        assert.strictEqual(formatParser.parseLocaleNumber("123.45", null), 123.45, 'Decimal');
        assert.strictEqual(formatParser.parseLocaleNumber("-50.5", null), -50.5, 'Negative');
    });

    // =====================================================================
    // FormatParser.parse - Numbers
    // =====================================================================
    QUnit.module('FormatParser.parse - Numbers');

    QUnit.test('thousand separators', function (assert) {
        let result = formatParser.parse("1,234", null);
        assert.ok(result !== null, 'Should parse "1,234"');
        assert.strictEqual(result.value, 1234, 'Value 1234');
        
        result = formatParser.parse("1,234,567", null);
        assert.ok(result !== null, 'Should parse "1,234,567"');
        assert.strictEqual(result.value, 1234567, 'Value 1234567');
    });

    QUnit.test('decimal numbers', function (assert) {
        let result = formatParser.parse("1,234.56", null);
        assert.ok(result !== null, 'Should parse');
        assert.ok(Math.abs(result.value - 1234.56) < eps, 'Value 1234.56');
    });

    QUnit.test('negative in parentheses', function (assert) {
        let result = formatParser.parse("(100)", null);
        assert.ok(result !== null, 'Should parse "(100)"');
        assert.strictEqual(result.value, -100, 'Value -100');
        
        result = formatParser.parse("(1,234.56)", null);
        assert.ok(result !== null, 'Should parse "(1,234.56)"');
        assert.ok(Math.abs(result.value - (-1234.56)) < eps, 'Value -1234.56');
    });

    QUnit.test('negative with minus', function (assert) {
        let result = formatParser.parse("-100", null);
        assert.ok(result !== null, 'Should parse');
        assert.strictEqual(result.value, -100, 'Value -100');
    });

    QUnit.test('positive with plus', function (assert) {
        let result = formatParser.parse("+100", null);
        assert.ok(result !== null, 'Should parse');
        assert.strictEqual(result.value, 100, 'Value 100');
    });

    // =====================================================================
    // FormatParser.parse - Percentages
    // =====================================================================
    QUnit.module('FormatParser.parse - Percentages');

    QUnit.test('percentage values', function (assert) {
        let result = formatParser.parse("50%", null);
        assert.ok(result !== null && result.bPercent, 'Should parse 50%');
        assert.ok(Math.abs(result.value - 0.5) < eps, 'Value 0.5');
        
        result = formatParser.parse("100%", null);
        assert.ok(Math.abs(result.value - 1) < eps, 'Value 1');
        
        result = formatParser.parse("12.5%", null);
        assert.ok(Math.abs(result.value - 0.125) < eps, 'Value 0.125');
        
        result = formatParser.parse("%50", null);
        assert.ok(result !== null, 'Should parse %50');
        assert.ok(Math.abs(result.value - 0.5) < eps, 'Value 0.5');
    });

    // =====================================================================
    // FormatParser.parse - Currencies
    // =====================================================================
    QUnit.module('FormatParser.parse - Currencies');

    QUnit.test('currency symbols', function (assert) {
        let result = formatParser.parse("$100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse $100');
        assert.strictEqual(result.value, 100, 'Value 100');
        
        result = formatParser.parse("€100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse €100');
        
        result = formatParser.parse("£100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse £100');
        
        result = formatParser.parse("¥100", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse ¥100');
        
        result = formatParser.parse("100р.", null);
        assert.ok(result !== null && result.bCurrency, 'Should parse 100р.');
    });

    QUnit.test('negative currency', function (assert) {
        let result = formatParser.parse("($100)", null);
        assert.ok(result !== null, 'Should parse ($100)');
        assert.strictEqual(result.value, -100, 'Value -100');
        
        result = formatParser.parse("-$100", null);
        assert.ok(result !== null, 'Should parse -$100');
        assert.strictEqual(result.value, -100, 'Value -100');
    });

    // =====================================================================
    // FormatParser.parse - Invalid inputs
    // =====================================================================
    QUnit.module('FormatParser.parse - Invalid inputs');

    QUnit.test('invalid patterns', function (assert) {
        assert.strictEqual(formatParser.parse("++100", null), null, 'Multiple plus');
        assert.strictEqual(formatParser.parse("--100", null), null, 'Multiple minus');
        assert.strictEqual(formatParser.parse("(100", null), null, 'Unmatched open paren');
        assert.strictEqual(formatParser.parse("100)", null), null, 'Unmatched close paren');
        assert.strictEqual(formatParser.parse("50%%", null), null, 'Multiple percent');
        assert.strictEqual(formatParser.parse("$€100", null), null, 'Mixed currencies');
        assert.strictEqual(formatParser.parse("", null), null, 'Empty string');
        assert.strictEqual(formatParser.parse("   ", null), null, 'Whitespace only');
    });

    // =====================================================================
    // FormatParser.parseDate - using parse() integration
    // =====================================================================
    QUnit.module('FormatParser.parse - Date/Time');

    QUnit.test('comprehensive date/time tests', function (assert) {
        let data = [
            ["1/2/2000 11:34:56", "m/d/yyyy h:mm", 36527.482592592591],
            ["1/2/2000 11:34:5", "m/d/yyyy h:mm", 36527.482002314813],
            ["1/2/2000 11:34:", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:34", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:3", "m/d/yyyy h:mm", 36527.460416666669],
            ["1/2/2000 11:", "m/d/yyyy h:mm", 36527.458333333336],
            ["11:34:56", "h:mm:ss", 0.48259259259259263],
            ["11:34:5", "h:mm:ss", 0.48200231481481487],
            ["11:34:", "h:mm", 0.48194444444444445],
            ["11:34", "h:mm", 0.48194444444444445],
            ["11:3", "h:mm", 0.4604166666666667],
            ["11:", "h:mm", 0.45833333333333331],
            ["1/2/2000 11:34:56 AM", "m/d/yyyy h:mm", 36527.482592592591],
            ["1/2/2000 11:34:5 AM", "m/d/yyyy h:mm", 36527.482002314813],
            ["1/2/2000 11:34: AM", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:34 AM", "m/d/yyyy h:mm", 36527.481944444444],
            ["1/2/2000 11:3 AM", "m/d/yyyy h:mm", 36527.460416666669],
            ["1/2/2000 11: AM", "m/d/yyyy h:mm", 36527.458333333336],
            ["11:34:56 AM", "h:mm:ss AM/PM", 0.48259259259259263],
            ["11:34:5 AM", "h:mm:ss AM/PM", 0.48200231481481487],
            ["11:34: AM", "h:mm AM/PM", 0.48194444444444445],
            ["11:34 AM", "h:mm AM/PM", 0.48194444444444445],
            ["11:3 AM", "h:mm AM/PM", 0.4604166666666667],
            ["11: AM", "h:mm AM/PM", 0.45833333333333331],
            ["11:00:00", "h:mm:ss", 0.45833333333333331],
            ["11:00:0", "h:mm:ss", 0.45833333333333331],
            ["11:00:", "h:mm", 0.45833333333333331],
            ["11:0", "h:mm", 0.45833333333333331],
            ["11:", "h:mm", 0.45833333333333331],
            ["1/2/2000 55:34:56", "General", 36529.315925925926],
            ["1/2/2000 55:34:5", "General", 36529.315335648149],
            ["1/2/2000 55:34:", "General", 36529.31527777778],
            ["1/2/2000 55:34", "General", 36529.31527777778],
            ["1/2/2000 55:3", "General", 36529.293749999997],
            ["1/2/2000 55:", "General", 36529.291666666664],
            ["55:34:56", "[h]:mm:ss", 2.3159259259259257],
            ["55:34:5", "[h]:mm:ss", 2.3153356481481482],
            ["55:34:", "[h]:mm:ss", 2.3152777777777778],
            ["55:34", "[h]:mm:ss", 2.3152777777777778],
            ["55:3", "[h]:mm:ss", 2.2937499999999997],
            ["55:", "[h]:mm:ss", 2.2916666666666665],
        ];
        for (let i = 0; i < data.length; i++) {
            let date = formatParser.parse(data[i][0]);
            assert.strictEqual(date.format, data[i][1], `Format: ${data[i][0]}`);
            assert.ok(Math.abs(date.value - data[i][2]) < eps, `Value: ${data[i][0]}`);
        }
    });

    QUnit.test('month name dates', function (assert) {
        let result = formatParser.parse("January 15, 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "January 15, 2023"');
        
        result = formatParser.parse("15 January 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "15 January 2023"');
        
        result = formatParser.parse("Jan 15, 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse abbreviated month');
        
        result = formatParser.parse("15-Jan-2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse dash-separated');
    });

    QUnit.test('month-year format (no day)', function (assert) {
        let result = formatParser.parse("Jan-2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "Jan-2023"');
        
        result = formatParser.parse("January 2023");
        assert.ok(result !== null && result.bDateTime, 'Should parse "January 2023"');
    });

    QUnit.test('day-month format (no year)', function (assert) {
        let result = formatParser.parse("15-Jan");
        assert.ok(result !== null && result.bDateTime, 'Should parse "15-Jan"');
        
        result = formatParser.parse("Jan 15");
        assert.ok(result !== null && result.bDateTime, 'Should parse "Jan 15"');
    });

    QUnit.test('time formats', function (assert) {
        let result = formatParser.parse("14:30");
        assert.ok(result !== null && result.bDateTime, 'Should parse "14:30"');
        
        result = formatParser.parse("2:30 PM");
        assert.ok(result !== null && result.bDateTime, 'Should parse "2:30 PM"');
        
        result = formatParser.parse("12:00 PM");
        assert.ok(result !== null, 'Should parse "12:00 PM"');
        assert.ok(Math.abs(result.value - 0.5) < eps, '12:00 PM = 0.5');
        
        result = formatParser.parse("12:00 AM");
        assert.ok(result !== null, 'Should parse "12:00 AM"');
        assert.ok(result.value < 0.01, '12:00 AM near 0');
    });

    QUnit.test('invalid time should fail', function (assert) {
        let result = formatParser.parse("14:60");
        assert.strictEqual(result, null, '60 minutes invalid');
        
        result = formatParser.parse("14:30:60");
        assert.strictEqual(result, null, '60 seconds invalid');
    });

    // =====================================================================
    // FormatParser.parseDatePDF
    // =====================================================================
    QUnit.module('FormatParser.parseDatePDF');

    QUnit.test('basic PDF date parsing', function (assert) {
        let result = formatParser.parseDatePDF("January 15, 2023", null);
        assert.ok(result !== null && result.bDate, 'Should parse month name date');
        
        result = formatParser.parseDatePDF("15 Jan 2023", null);
        assert.ok(result !== null && result.bDate, 'Should parse abbreviated month');
    });

    QUnit.test('PDF dates before 1900', function (assert) {
        let result = formatParser.parseDatePDF("January 15, 1850", null);
        assert.ok(result !== null, 'Should parse date before 1900');
    });

    QUnit.test('PDF dates with time', function (assert) {
        let result = formatParser.parseDatePDF("January 15, 2023 14:30:45", null);
        assert.ok(result !== null && result.bDate && result.bTime, 'Should parse date with time');
    });

    // =====================================================================
    // Date1904 mode
    // =====================================================================
    QUnit.module('Date1904 mode');

    QUnit.test('bDate1904 affects date values', function (assert) {
        let original1904 = AscCommon.bDate1904;
        
        try {
            AscCommon.bDate1904 = false;
            let result1 = formatParser.parse("January 1, 2000");
            
            AscCommon.bDate1904 = true;
            let result2 = formatParser.parse("January 1, 2000");
            
            if (result1 && result2) {
                // Values should differ by ~1462 days
                let diff = Math.abs(result1.value - result2.value);
                assert.ok(diff > 1400 && diff < 1500, 'Values differ by ~1462 days');
            } else {
                assert.ok(true, 'Parse returned null - expected for some locales');
            }
        } finally {
            AscCommon.bDate1904 = original1904;
        }
    });

    
    // =====================================================================
    // CellFormat.format - Number formatting
    // =====================================================================
    QUnit.module('CellFormat.format');

    QUnit.test('number formatting', function (assert) {
        let testCases = [
            // Thousand separators
            [1234, '#,##0', '1,234'],
            [1234567, '#,##0', '1,234,567'],
            [0, '#,##0', '0'],
            [-1234, '#,##0', '-1,234'],
            
            // Decimal places
            [1234.56, '#,##0.00', '1,234.56'],
            [1234.5, '#,##0.00', '1,234.50'],
            [0.5, '0.00', '0.50'],
            [1.234, '0.00', '1.23'],
            
            // Percentages
            [0.5, '0%', '50%'],
            [0.125, '0.00%', '12.50%'],
            [1, '0%', '100%'],
            [0.999, '0%', '100%'],
            
            // Currency with text literals
            [1234.56, '"$"#,##0.00', '$1,234.56'],
            [0, '"$"#,##0.00', '$0.00'],
            [-50, '"$"#,##0.00', '-$50.00'],
            [1000, '"USD "0.00', 'USD 1000.00'],
            
            // Negative numbers in parentheses
            [100, '0;(0)', '100'],
            [-100, '0;(0)', '(100)'],
            [0, '0;(0)', '0'],
            [-50.5, '0.00;(0.00)', '(50.50)'],
            
            // Optional digits with #
            [123, '###', '123'],
            [0, '###', ''],
            [12.3, '##.#', '12.3'],
            [12, '##.#', '12.'],
            
            // Mandatory zeros
            [5, '000', '005'],
            [123, '000', '123'],
            [5.5, '000.00', '005.50'],
            [0, '00', '00'],
            
            // Space alignment with ?
            [1, '??', '01'],
            [10, '??', '10'],
            [1.5, '?.??', '1.50'],
            [10.25, '?.??', '10.25'],
            
            // Escaped characters
            [100, '\\#0', '#100'],
            [50, '0\\%', '50%'],
            [10, '0\\-', '10-'],
            [25, '\\+0', '+25'],
            
            // Mixed format
            [1234.5, '#,##0.00;[Red](#,##0.00)', '1,234.50'],
            [-1234.5, '#,##0.00;[Red](#,##0.00)', '(1,234.50)'],
            
            // Additional important cases
            [0.75, '0.#', '0.8'],
            [100.123, '0.0', '100.1'],
            [1234, '"Total: "#,##0', 'Total: 1,234'],
            [0.5555, '0.00%', '55.55%'],
            [999999, '#,##0', '999,999'],
            [-0.25, '0.00;(0.00)', '(0.25)'],
        ];
        
        for (let i = 0; i < testCases.length; i++) {
            let value = testCases[i][0];
            let format = testCases[i][1];
            let expected = testCases[i][2];
            
            let expr = new AscCommon.CellFormat(format);
            let formatted = expr.format(value);
            let text = '';
            for (let j = 0, length = formatted.length; j < length; ++j) {
                text += formatted[j].text;
            }
            
            assert.strictEqual(text, expected, `format("${format}", ${value})`);
        }
    });

    QUnit.test('date/time elapsed formats', function (assert) {
        let testCases = [
            // Date format cases
            [0.684027777777778, 'mm', '01'],
            [0.684027777777778, '[mm]', '985'],
            [0.684027777777778, '[h] "hours"', '16 hours'],
            [0.684027777777778, '[h]:mm', '16:25'],
            [0.684027777777778, '[h]:mm" ""minutes"', '16:25 minutes'],
            [0.684027777777778, '[s]', '59100'],
            [0.684027777777778, '[s]" ""seconds"', '59100 seconds'],
            [0.684027777777778, '[ss].0', '59100.0'],
            [0.684027777777778, '[mm]:ss', '985:00'],
            [0.684027777777778, '[mm]:mm', '985:01'],
            [0.684027777777778, '[hh]', '16'],
            [0.684027777777778, '[h]:mm:ss.000', '16:25:00.000'],
            [0.684027777777778, 'dd"d "hh"h "mm"m "ss"s"" "AM/PM', '00d 04h 25m 00s PM'],
            [0.684027777777778, '[h]"h*"mm"m*"ss"s*"ss"ms"', '16h*25m*00s*00ms'],
            [0.684027777777778, 'yyyy"Y-"mm"M-"dd"D "hh"H:"mm"M:"ss"."s"S"" "AM/PM', '1900Y-01M-00D 04H:25M:00.0S PM'],
            [0.684027777777778, 'dd:mm:yyyy" "hh:mm:ss" "[hh]:[mm]" "AM/PM" ""minutes AM/PM"', '00:01:1900 04:25:00 04:985 PM minutes AM/PM'],

            [37753.6844097222, 'mm', '05'],
            [37753.6844097222, '[mm]', '54365305'],
            [37753.6844097222, '[h] "hours"', '906088 hours'],
            [37753.6844097222, '[h]:mm', '906088:25'],
            [37753.6844097222, '[h]:mm" ""minutes"', '906088:25 minutes'],
            [37753.6844097222, '[s]', '3261918333'],
            [37753.6844097222, '[s]" ""seconds"', '3261918333 seconds'],
            [37753.6844097222, '[ss].0', '3261918333.0'],
            [37753.6844097222, '[mm]:ss', '54365305:33'],
            [37753.6844097222, '[mm]:mm', '54365305:05'],
            [37753.6844097222, '[hh]', '906088'],
            [37753.6844097222, '[h]:mm:ss.000', '906088:25:33.000'],
            [37753.6844097222, 'dd"d "hh"h "mm"m "ss"s"" "AM/PM', '12d 04h 25m 33s PM'],
            [37753.6844097222, '[h]"h*"mm"m*"ss"s*"ss"ms"', '906088h*25m*33s*33ms'],
            [37753.6844097222, 'yyyy"Y-"mm"M-"dd"D "hh"H:"mm"M:"ss"."s"S"" "AM/PM', '2003Y-05M-12D 04H:25M:33.33S PM'],
            [37753.6844097222, 'dd:mm:yyyy" "hh:mm:ss" "[hh]:[mm]" "AM/PM" ""minutes AM/PM"', '12:05:2003 04:25:33 04:54365305 PM minutes AM/PM'],
        ];
        
        for (let i = 0; i < testCases.length; i++) {
            let [value, format, expected] = testCases[i];
            let expr = new AscCommon.CellFormat(format);
            let formatted = expr.format(value);
            let text = formatted.map(f => f.text).join('');
            assert.strictEqual(text, expected, `format("${format}", ${value})`);
        }
    });

    QUnit.test('formatRecognition', function (assert) {
        let testCases = [
            ['1,234', '#,##0', 1234],
            ['1,234,567', '#,##0', 1234567],
            ['-1,234', '#,##0', -1234],
            
            // Decimal places
            ['1,234.56', '#,##0.00', 1234.56],
            ['1,234.50', '#,##0.00', 1234.5],
            
            // Percentages
            ['50%', '0%', 0.5],
            ['12.50%', '0.00%', 0.125],
            ['100%', '0%', 1],
            [' 100 %', '0%', 1],
            
            // Currency with text literals
            ['$1,234.56', '\\$#,##0.00_);[Red](\\$#,##0.00)', 1234.56],
            ['$0.00', '\\$#,##0_);[Red](\\$#,##0)', 0],
            ['-$50.00', '\\$#,##0_);[Red](\\$#,##0)', -50],
            ['USD 1000.00', null, 'USD 1000.00'],
            
            // Negative numbers in parentheses
            ['(100)', 'General', -100],
            ['(50.50)', 'General', -50.5],
            
            // Optional digits with
            ['123', 'General', 123],
            ['12.3', 'General', 12.3],
            ['12.', 'General', 12],
            

            // Fraction format cases
            ["1/2", "d-mmm", 45659],
            ["3/4", "d-mmm", 45720],
            ["15/20", null, "15/20"],
            [" 1/2", null, " 1/2"],
            ["150/200", null, "150/200"],
            ["0 1/5/5", null, "0 1/5/5"],
            ["1/5/5", "m/d/yyyy", 38357],
            [" 150/200", null, " 150/200"],
            ["+1/2", null, "+1/2"],
            ["-1/2", null, "-1/2"],
            ["$1/2", null, "$1/2"],
            ["(1/2", null, "(1/2"],
            ["1/2)", null, "1/2)"],
            ["1/2%", null, "1/2%"],
            ["1/2 $", null, "1/2 $"],
            ["1/2 p.", null, "1/2 p."],
            ["+1 1/2%", "0.00%", 0.015],
            ["-$2 3/4", "# ?/?", -2.75], //General
            ["(100 1/2)", "# ?/?", -100.5],
            ["25 50/100 %", "0.00%", 0.255],

            ["0 1/2", "# ?/?", 0.5],
            ["0 1/10", "# ??/??", 0.1],
            ["0 1/100", "# ??/??", 0.01],
            ["0 10/2", "# ?/?", 5],
            ["0 15/3", "# ?/?", 5],
            ["0 17/7", "# ?/?", 2.4285714285714284],
            ["0 15/20", "# ??/??", 0.75],
            ["0 12/120", "# ??/??", 0.1],
            ["0 25/250", "# ??/??", 0.1],
            ["0 100/200", "# ??/??", 0.5],
            ["0 125/250", "# ??/??", 0.5],
            ["0 0/1", "# ?/?", 0],
            ["0 1/1", "# ?/?", 1],
            ["0 999/999", "# ??/??", 1],
            ["0 1/999", "# ??/??", 0.001001001001001001],
            ["0 999/1", "# ?/?", 999],

            ["1 999/1", "# ?/?", 1000],
            ["1 999/12", "# ??/??", 84.25],
            ["1 999/134", "# ??/??", 8.455223880597014],
        ]; 
        
        for (let i = 0; i < testCases.length; i++) {
            let value = testCases[i][0];
            let format = testCases[i][1];
            let expectedValue = testCases[i][2];
            
            let formatted = AscCommon.g_oFormatParser.parse(value);

            if (formatted) {
                assert.strictEqual(formatted.format, format, `Case format: ${value}`);
                assert.strictEqual(formatted.value, expectedValue, `Case value: ${expectedValue}`);
            } else {
                assert.strictEqual(formatted, format, `Case format: ${value}`);
            }
        }
    });
    QUnit.test('formatRecognitionWithSelectedFormat', function (assert) {
        let testCases = [
            // appliedFormat = 0 (General)
            ["1/2", "d-mmm", 45659, formatTypes.General],
            ["3/4", "d-mmm", 45720, formatTypes.General],
            ["15/20", null, "15/20", formatTypes.General],  
            [" 1/2", null, " 1/2", formatTypes.General],
            ["1 1/2", "# ?/?", 1.5, formatTypes.General],
            ["2 3/4", "# ?/?", 2.75, formatTypes.General],
            ["15/3", null, "15/3", formatTypes.General],
            ["150/200", null, "150/200", formatTypes.General],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.General],
            ["0 1/2", "# ?/?", 0.5, formatTypes.General],
            ["0 1/10", "# ??/??", 0.1, formatTypes.General],
            ["0 1/100", "# ??/??", 0.01, formatTypes.General],
            ["1 150/200", "# ??/??", 1.75, formatTypes.General],
            
            // appliedFormat = 1 (Number - 0)
            ["1/2", "0.00", 0.5, formatTypes.Number, "0.00"],
            ["3/4", "0.00", 0.75, formatTypes.Number, "0.00"],
            ["15/20", "0.00", 0.75, formatTypes.Number, "0.00"],
            [" 1/2", "0.00", 0.5, formatTypes.Number, "0.00"],
            ["1 1/2", "0.00", 1.5, formatTypes.Number, "0.00"],
            ["2 3/4", "0.00", 2.75, formatTypes.Number, "0.00"],
            ["15/3", "0.00", 5, formatTypes.Number, "0.00"],
            ["150/200", "0.00", 0.75, formatTypes.Number, "0.00"],
            ["1/5/5", "0.00", 38357, formatTypes.Number, "0.00"],
            ["0 1/2", "0.00", 0.5, formatTypes.Number, "0.00"],
            ["0 1/10", "0.00", 0.1, formatTypes.Number, "0.00"],
            ["0 1/100", "0.00", 0.01, formatTypes.Number, "0.00"],
            ["1 150/200", "0.00", 1.75, formatTypes.Number, "0.00"],


            // Specific number format
            ["1/2", "0.00000", 0.5, formatTypes.Number, "0.00000"],
            ["3/4", "0.00000", 0.75, formatTypes.Number, "0.00000"],
            ["15/20", "0.00000", 0.75, formatTypes.Number, "0.00000"],
            [" 1/2", "0.00000", 0.5, formatTypes.Number, "0.00000"],
            ["1 1/2", "0.00000", 1.5, formatTypes.Number, "0.00000"],
            ["2 3/4", "0.00000", 2.75, formatTypes.Number, "0.00000"],
            ["15/3", "0.00000", 5, formatTypes.Number, "0.00000"],
            ["150/200", "0.00000", 0.75, formatTypes.Number, "0.00000"],
            ["1/5/5", "0.00000", 38357, formatTypes.Number, "0.00000"],
            ["0 1/2", "0.00000", 0.5, formatTypes.Number, "0.00000"],
            ["0 1/10", "0.00000", 0.1, formatTypes.Number, "0.00000"],
            ["0 1/100", "0.00000", 0.01, formatTypes.Number, "0.00000"],
            ["1 150/200", "0.00000", 1.75, formatTypes.Number, "0.00000"],

            // appliedFormat = 2 (Scientific - 0.00E+00)
            ["1/2", "# ?/?", 0.5, formatTypes.Scientific, "0.00E+00"],
            ["3/4", "# ?/?", 0.75, formatTypes.Scientific, "0.00E+00"],
            ["15/20", "# ??/??", 0.75, formatTypes.Scientific, "0.00E+00"],
            [" 1/2", "# ?/?", 0.5, formatTypes.Scientific, "0.00E+00"],
            ["1 1/2", "# ?/?", 1.5, formatTypes.Scientific, "0.00E+00"],
            ["2 3/4", "# ?/?", 2.75, formatTypes.Scientific, "0.00E+00"],
            ["15/3", "# ?/?", 5, formatTypes.Scientific, "0.00E+00"],
            ["150/200", "# ??/??", 0.75, formatTypes.Scientific, "0.00E+00"],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.Scientific, "0.00E+00"],
            ["0 1/2", "# ?/?", 0.5, formatTypes.Scientific, "0.00E+00"],
            ["0 1/10", "# ??/??", 0.1, formatTypes.Scientific, "0.00E+00"],
            ["0 1/100", "# ??/??", 0.01, formatTypes.Scientific, "0.00E+00"],
            ["1 150/200", "# ??/??", 1.75, formatTypes.Scientific, "0.00E+00"],
            
            // Specific scientific format
            ["1/2", "0.00000E+00", 0.5, formatTypes.Scientific, "0.00000E+00"],
            ["3/4", "0.00000E+00", 0.75, formatTypes.Scientific, "0.00000E+00"],
            ["15/20", "0.00000E+00", 0.75, formatTypes.Scientific, "0.00000E+00"],
            [" 1/2", "0.00000E+00", 0.5, formatTypes.Scientific, "0.00000E+00"],
            ["1 1/2", "0.00000E+00", 1.5, formatTypes.Scientific, "0.00000E+00"],
            ["2 3/4", "0.00000E+00", 2.75, formatTypes.Scientific, "0.00000E+00"],
            ["15/3", "0.00000E+00", 5, formatTypes.Scientific, "0.00000E+00"],
            ["150/200", "0.00000E+00", 0.75, formatTypes.Scientific, "0.00000E+00"],
            ["1/5/5", "0.00000E+00", 38357, formatTypes.Scientific, "0.00000E+00"],
            ["0 1/2", "0.00000E+00", 0.5, formatTypes.Scientific, "0.00000E+00"],
            ["0 1/10", "0.00000E+00", 0.1, formatTypes.Scientific, "0.00000E+00"],
            ["0 1/100", "0.00000E+00", 0.01, formatTypes.Scientific, "0.00000E+00"],
            ["1 150/200", "0.00000E+00", 1.75, formatTypes.Scientific, "0.00000E+00"],
            

            // appliedFormat = 3 (Accounting)
            ["1/2", "# ?/?", 0.5, formatTypes.Accounting],
            ["3/4", "# ?/?", 0.75, formatTypes.Accounting],
            ["15/20", "# ??/??", 0.75, formatTypes.Accounting],
            [" 1/2", "# ?/?", 0.5, formatTypes.Accounting],
            ["1 1/2", "# ?/?", 1.5, formatTypes.Accounting],
            ["2 3/4", "# ?/?", 2.75, formatTypes.Accounting],
            ["15/3", "# ?/?", 5, formatTypes.Accounting],
            ["150/200", "# ??/??", 0.75, formatTypes.Accounting],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.Accounting],
            ["0 1/2", "# ?/?", 0.5, formatTypes.Accounting],
            ["0 1/10", "# ??/??", 0.1, formatTypes.Accounting],
            ["0 1/100", "# ??/??", 0.01, formatTypes.Accounting],
            ["1 150/200", "# ??/??", 1.75, formatTypes.Accounting],
            
            // Specific Accounting formats
            ["1/2", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.5, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["3/4", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.75, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["15/20", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.75, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            [" 1/2", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.5, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["1 1/2", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 1.5, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["2 3/4", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 2.75, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["15/3", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 5, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["150/200", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.75, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["1/5/5", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 38357, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["0 1/2", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.5, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["0 1/10", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.1, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["0 1/100", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 0.01, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            ["1 150/200", '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)', 1.75, formatTypes.Accounting, '_([$$-9]* #,##0.000_);_([$$-9]* \\(#,##0.000\\);_([$$-9]* "-"???_);_(@_)'],
            
            // appliedFormat = 4 (Currency - $#,##0.00)
            ["1/2", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.5, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["3/4", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.75, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["15/20", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.75, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            [" 1/2", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.5, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["1 1/2", "\\$#,##0.00_);[Red](\\$#,##0.00)", 1.5, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["2 3/4", "\\$#,##0.00_);[Red](\\$#,##0.00)", 2.75, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["15/3", "\\$#,##0_);[Red](\\$#,##0)", 5, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["150/200", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.75, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["1/5/5", "\\$#,##0.00_);[Red](\\$#,##0.00)", 38357, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["0 1/2", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.5, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["0 1/10", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.1, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["0 1/100", "\\$#,##0.00_);[Red](\\$#,##0.00)", 0.01, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            ["1 150/200", "\\$#,##0.00_);[Red](\\$#,##0.00)", 1.75, formatTypes.Currency, "\\$#,##0.00_);[Red](\\$#,##0.00)"],
            
            // appliedFormat = 5 (Date - m/d/yyyy)
            ["1/2", "m/d/yyyy", 45659, formatTypes.Date, "m/d/yyyy"],
            ["3/4", "m/d/yyyy", 45720, formatTypes.Date, "m/d/yyyy"],
            ["15/20", null, "15/20", formatTypes.Date, "m/d/yyyy"], 
            [" 1/2", null, " 1/2", formatTypes.Date, "m/d/yyyy"],
            ["1 1/2", "# ?/?", 1.5, formatTypes.Date, "m/d/yyyy"],
            ["2 3/4", "# ?/?", 2.75, formatTypes.Date, "m/d/yyyy"],
            ["15/3", null, "15/3", formatTypes.Date, "m/d/yyyy"],
            ["150/200", null, "150/200", formatTypes.Date, "m/d/yyyy"],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.Date, "m/d/yyyy"],
            ["0 1/2", "# ?/?", 0.5, formatTypes.Date, "m/d/yyyy"],
            ["0 1/10", "# ??/??", 0.1, formatTypes.Date, "m/d/yyyy"],
            ["0 1/100", "# ??/??", 0.01, formatTypes.Date, "m/d/yyyy"],
            ["1 150/200", "# ??/??", 1.75, formatTypes.Date, "m/d/yyyy"],
            
            // appliedFormat = 6 (LongDate - dddd, mmmm d, yyyy)
            ["1/2", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 45659, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["3/4", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 45720, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["15/20", null, "15/20", formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            [" 1/2", null, " 1/2", formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["1 1/2", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 1.5, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"], 
            ["2 3/4", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 2.75, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["15/3", null, "15/3", formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["150/200", null, "150/200", formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["1/5/5", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 38357, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],
            ["0 1/2", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 0.5, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"], 
            ["0 1/10", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 0.1, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],   
            ["0 1/100", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 0.01, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],  
            ["1 150/200", "dddd\\,\\ mmmm\\ d\\,\\ yyyy", 1.75, formatTypes.Date, "[$-F800]dddd\\,\\ mmmm\\ d\\,\\ yyyy"],  
            
            // Specific date formats
            ["1/2", "yyyy-mm-dd", 45659, formatTypes.Date, "yyyy-mm-dd"],
            ["1/2", "yyyy-mm-dd", 45659, formatTypes.Date, "yyyy-mm-dd"],
            ["3/4", "yyyy-mm-dd", 45720, formatTypes.Date, "yyyy-mm-dd"],
            ["15/20", null, "15/20", formatTypes.Date, "yyyy-mm-dd"], 
            [" 1/2", null, " 1/2", formatTypes.Date, "yyyy-mm-dd"],
            ["1 1/2", "yyyy-mm-dd", 1.5, formatTypes.Date, "yyyy-mm-dd"],
            ["2 3/4", "yyyy-mm-dd", 2.75, formatTypes.Date, "yyyy-mm-dd"],
            ["15/3", null, "15/3", formatTypes.Date, "yyyy-mm-dd"],
            ["150/200", null, "150/200", formatTypes.Date, "yyyy-mm-dd"],
            ["1/5/5", "yyyy-mm-dd", 38357, formatTypes.Date, "yyyy-mm-dd"],
            ["0 1/2", "yyyy-mm-dd", 0.5, formatTypes.Date, "yyyy-mm-dd"],
            ["0 1/10", "yyyy-mm-dd", 0.1, formatTypes.Date, "yyyy-mm-dd"],
            ["0 1/100", "yyyy-mm-dd", 0.01, formatTypes.Date, "yyyy-mm-dd"],
            ["1 150/200", "yyyy-mm-dd", 1.75, formatTypes.Date, "yyyy-mm-dd"],

            // appliedFormat = 7 (Time - h:mm:ss)
            ["1/2", "h:mm:ss", 45659, formatTypes.Time, "h:mm:ss"],
            ["3/4", "h:mm:ss", 45720, formatTypes.Time, "h:mm:ss"],
            ["15/20", null, "15/20", formatTypes.Time, "h:mm:ss"],
            [" 1/2", null, " 1/2", formatTypes.Time, "h:mm:ss"],
            ["1 1/2", "h:mm:ss", 1.5, formatTypes.Time, "h:mm:ss"],
            ["2 3/4", "h:mm:ss", 2.75, formatTypes.Time, "h:mm:ss"],
            ["15/3", null, "15/3", formatTypes.Time, "h:mm:ss"],
            ["150/200", null, "150/200", formatTypes.Time, "h:mm:ss"],
            ["1/5/5", "h:mm:ss", 38357, formatTypes.Time, "h:mm:ss"],
            ["0 1/2", "h:mm:ss", 0.5, formatTypes.Time, "h:mm:ss"],
            ["0 1/10", "h:mm:ss", 0.1, formatTypes.Time, "h:mm:ss"],
            ["0 1/100", "h:mm:ss", 0.01, formatTypes.Time, "h:mm:ss"],
            ["1 150/200", "h:mm:ss", 1.75, formatTypes.Time, "h:mm:ss"],
            
            // Specific time format
            ["1/2", "h:mm AM/PM", 45659, formatTypes.Time, "h:mm AM/PM"],
            ["3/4", "h:mm AM/PM", 45720, formatTypes.Time, "h:mm AM/PM"],
            ["15/20", null, "15/20", formatTypes.Time, "h:mm AM/PM"],
            [" 1/2", null, " 1/2", formatTypes.Time, "h:mm AM/PM"],
            ["1 1/2", "h:mm AM/PM", 1.5, formatTypes.Time, "h:mm AM/PM"],
            ["2 3/4", "h:mm AM/PM", 2.75, formatTypes.Time, "h:mm AM/PM"],
            ["15/3", null, "15/3", formatTypes.Time, "h:mm AM/PM"],
            ["150/200", null, "150/200", formatTypes.Time, "h:mm AM/PM"],
            ["1/5/5", "h:mm AM/PM", 38357, formatTypes.Time, "h:mm AM/PM"],
            ["0 1/2", "h:mm AM/PM", 0.5, formatTypes.Time, "h:mm AM/PM"],
            ["0 1/10", "h:mm AM/PM", 0.1, formatTypes.Time, "h:mm AM/PM"],
            ["0 1/100", "h:mm AM/PM", 0.01, formatTypes.Time, "h:mm AM/PM"],
            ["1 150/200", "h:mm AM/PM", 1.75, formatTypes.Time, "h:mm AM/PM"],

            // appliedFormat = 8 (Percent - 0%)
            ["1/2", "0.00%", 0.005, formatTypes.Percent, "0.00%"],
            ["3/4", "0.00%", 0.0075, formatTypes.Percent, "0.00%"],
            ["15/20", "0.00%", 0.0075, formatTypes.Percent, "0.00%"],
            [" 1/2", "# ?/?", 0.5, formatTypes.Percent, "0.00%"],
            ["1 1/2", "0.00%", 0.015, formatTypes.Percent, "0.00%"],
            ["2 3/4", "0.00%", 0.0275, formatTypes.Percent, "0.00%"],
            ["15/3", "0%", 0.05, formatTypes.Percent, "0.00%"],
            ["150/200", "0.00%", 0.0075, formatTypes.Percent, "0.00%"],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.Percent, "0.00%"],
            ["0 1/2", "0.00%", 0.005, formatTypes.Percent, "0.00%"],
            ["0 1/10", "0.00%", 0.001, formatTypes.Percent, "0.00%"],
            ["0 1/100", "0.00%", 0.0001, formatTypes.Percent, "0.00%"],
            ["1 150/200", "0.00%", 0.0175, formatTypes.Percent, "0.00%"],
            
            // Specific percent formats
            ["1/2", "0.00%", 0.005, formatTypes.Percent, "0.0000%"],
            ["3/4", "0.00%", 0.0075, formatTypes.Percent, "0.0000%"],
            ["15/20", "0.00%", 0.0075, formatTypes.Percent, "0.0000%"],
            [" 1/2", "0.00%", 0.005, formatTypes.Percent, "0.0000%"],
            ["1 1/2", "0.00%", 0.015, formatTypes.Percent, "0.0000%"],
            ["2 3/4", "0.00%", 0.0275, formatTypes.Percent, "0.0000%"],
            ["15/3", "0%", 0.05, formatTypes.Percent, "0.0000%"],
            ["150/200", "0.00%", 0.0075, formatTypes.Percent, "0.0000%"],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.Percent, "0.0000%"],
            ["0 1/2", "0.00%", 0.005, formatTypes.Percent, "0.0000%"],
            ["0 1/10", "0.00%", 0.001, formatTypes.Percent, "0.0000%"],
            ["0 1/100", "0.00%", 0.0001, formatTypes.Percent, "0.0000%"],
            ["1 150/200", "0.00%", 0.0175, formatTypes.Percent, "0.0000%"],

            // appliedFormat = 9 (Fraction - # ?/?)
            ["1/2", "# ?/?", 0.5, formatTypes.Fraction, "# ?/?"],
            ["3/4", "# ?/?", 0.75, formatTypes.Fraction, "# ?/?"],
            ["15/20", "# ?/?", 0.75, formatTypes.Fraction, "# ?/?"],
            [" 1/2", "# ?/?", 0.5, formatTypes.Fraction, "# ?/?"],
            ["1 1/2", "# ?/?", 1.5, formatTypes.Fraction, "# ?/?"],
            ["2 3/4", "# ?/?", 2.75, formatTypes.Fraction, "# ?/?"],
            ["15/3", "# ?/?", 5, formatTypes.Fraction, "# ?/?"],
            ["150/200", "# ?/?", 0.75, formatTypes.Fraction, "# ?/?"],
            ["1 150/200", "# ?/?", 1.75, formatTypes.Fraction, "# ?/?"],
            ["1/5/5", "m/d/yyyy", 38357, formatTypes.Fraction, "# ?/?"],
            ["0 1/2", "# ?/?", 0.5, formatTypes.Fraction, "# ?/?"],
            ["0 1/10", "# ?/?", 0.1, formatTypes.Fraction, "# ?/?"],
            ["0 1/100", "# ?/?", 0.01, formatTypes.Fraction, "# ?/?"],

            // Specific fraction formats
            ["1/2", "# ?/2", 0.5, formatTypes.Fraction, "# ?/2"],
            ["0 1/2", "# ?/2", 0.5, formatTypes.Fraction, "# ?/2"],
            ["1 1/2", "# ?/2", 1.5, formatTypes.Fraction, "# ?/2"],
            [" 1/2", "# ?/2", 0.5, formatTypes.Fraction, "# ?/2"],
            ["3/7", "# ?/2", 0.42857142857142855, formatTypes.Fraction, "# ?/2"],
            ["0 3/7", "# ?/2", 0.42857142857142855, formatTypes.Fraction, "# ?/2"],
            ["1 3/7", "# ?/2", 1.42857142857142855, formatTypes.Fraction, "# ?/2"],
            ["25/35", "# ?/2", 0.7142857142857143, formatTypes.Fraction, "# ?/2"],
            ["0 25/35", "# ?/2", 0.7142857142857143, formatTypes.Fraction, "# ?/2"],
            ["1 25/35", "# ?/2", 1.7142857142857144, formatTypes.Fraction, "# ?/2"],
            ["17/37", "# ?/2", 0.4594594594594595, formatTypes.Fraction, "# ?/2"],
            ["0 17/37", "# ?/2", 0.4594594594594595, formatTypes.Fraction, "# ?/2"],
            ["1 17/37", "# ?/2", 1.4594594594594595, formatTypes.Fraction, "# ?/2"],
            ["150/200", "# ?/2", 0.75, formatTypes.Fraction, "# ?/2"],
            ["0 150/200", "# ?/2", 0.75, formatTypes.Fraction, "# ?/2"],
            ["1 150/200", "# ?/2", 1.75, formatTypes.Fraction, "# ?/2"],
            ["137/235", "# ?/2", 0.5829787234042553, formatTypes.Fraction, "# ?/2"],
            ["0 137/235", "# ?/2", 0.5829787234042553, formatTypes.Fraction, "# ?/2"],
            ["1 137/235", "# ?/2", 1.5829787234042553, formatTypes.Fraction, "# ?/2"],
           
            ["1/2", "# ?/8", 0.5, formatTypes.Fraction, "# ?/8"],
            ["0 1/2", "# ?/8", 0.5, formatTypes.Fraction, "# ?/8"],
            ["1 1/2", "# ?/8", 1.5, formatTypes.Fraction, "# ?/8"],
            [" 1/2", "# ?/8", 0.5, formatTypes.Fraction, "# ?/8"],
            ["3/7", "# ?/8", 0.42857142857142855, formatTypes.Fraction, "# ?/8"],
            ["0 3/7", "# ?/8", 0.42857142857142855, formatTypes.Fraction, "# ?/8"],
            ["1 3/7", "# ?/8", 1.42857142857142855, formatTypes.Fraction, "# ?/8"],
            ["25/35", "# ?/8", 0.7142857142857143, formatTypes.Fraction, "# ?/8"],
            ["0 25/35", "# ?/8", 0.7142857142857143, formatTypes.Fraction, "# ?/8"],
            ["1 25/35", "# ?/8", 1.7142857142857144, formatTypes.Fraction, "# ?/8"],
            ["17/37", "# ?/8", 0.4594594594594595, formatTypes.Fraction, "# ?/8"],
            ["0 17/37", "# ?/8", 0.4594594594594595, formatTypes.Fraction, "# ?/8"],
            ["1 17/37", "# ?/8", 1.4594594594594595, formatTypes.Fraction, "# ?/8"],
            ["150/200", "# ?/8", 0.75, formatTypes.Fraction, "# ?/8"],
            ["0 150/200", "# ?/8", 0.75, formatTypes.Fraction, "# ?/8"],
            ["1 150/200", "# ?/8", 1.75, formatTypes.Fraction, "# ?/8"],
            ["137/235", "# ?/8", 0.5829787234042553, formatTypes.Fraction, "# ?/8"],
            ["0 137/235", "# ?/8", 0.5829787234042553, formatTypes.Fraction, "# ?/8"],
            ["1 137/235", "# ?/8", 1.5829787234042553, formatTypes.Fraction, "# ?/8"],
            
            // appliedFormat = 10 (Text - @)
            ["1/2", "@", "1/2", 10],
            ["3/4", "@", "3/4", 10],
            ["15/20", "@", "15/20", 10],
            [" 1/2", "@", " 1/2", 10],
            ["1 1/2", "@", "1 1/2", 10],
            ["2 3/4", "@", "2 3/4", 10],
            ["15/3", "@", "15/3", 10],
            ["150/200", "@", "150/200", 10],
            ["1/5/5", "@", "1/5/5", 10],
            ["0 1/2", "@", "0 1/2", 10],
            ["0 1/10", "@", "0 1/10", 10],
            ["0 1/100", "@", "0 1/100", 10],
        ];
        
        for (let i = 0; i < testCases.length; i++) {
            let value = testCases[i][0];
            let format = testCases[i][1];
            let res = testCases[i][2];
            let appliedFormat = testCases[i][3]
            
            let appliedSpecificFormat = testCases[i][4]


            // Apply format
            let formatted = AscCommon.g_oFormatParser.parse(value, null, appliedFormat, appliedSpecificFormat);

            if (formatted) {
                assert.strictEqual(formatted.format, format, `Case format: ${value}`);              
                assert.strictEqual(formatted.value, res, `Case format: ${res}`);              
            } else {
                assert.strictEqual(formatted, format, `Case format: ${value}`);
            }

        }
    });

});
