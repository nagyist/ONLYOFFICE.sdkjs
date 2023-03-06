/*
 * (c) Copyright Ascensio System SIA 2010-2023
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


(function (window)
{
	const {
		wbModel,
		wbView,
		executeTestWithCatchEvent,
		getFragments,
		getSelectionCellEditor,
		moveToStartCellEditor,
		moveToEndCellEditor,
		moveRight,
		moveToCell,
		selectToCell,
		checkOpenCellEditor,
		onKeyDown,
		remove,
		closeCellEditor,
		enterTextWithoutClose,
		setCheckOpenCellEditor,
		enterText,
		cellEditor,
		getCellText,
		getCellTextWithoutFormat,
		moveDown,
		wsView,
		ws,
		moveAndEnterText,
		createTest,
		moveAndGetCellText,
		goToSheet,
		createWorksheet,
		removeCurrentWorksheet,
		cleanCell,
		cleanRange,
		cleanSelection,
		cleanActiveCell,
		checkRange,
		openCellEditor,
		checkActiveCell,
		cleanCache,
		selectAll,
		cleanAll,
		setCellFormat,
		selectionInfo,
		xfs,
		undo,
		createEvent,
		selectAllCell,
		cellPosition,
		getCellEditMode,
		testPreventDefaultAndStopPropagation,
		controller,
		handlers,
		activeCell,
		selectionRange,
		activeCellRange
	} = window.AscTestShortcut;


	$(
		function ()
		{
			QUnit.module('test worksheet shortcuts');
			QUnit.test('Test refresh all connections', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test refresh selected connections', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});


			QUnit.test('Test change format table info', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test calculate all', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test calculate workbook', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});


			QUnit.test('Test calculate active sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test calculate only changed', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test focus on cell editor', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(113, false, false, false, false, false);
				equal(getCellEditMode(), true);
			});

			QUnit.test('Test add date', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				const oEvent = createEvent(186, true, false, false, false, false);
				moveToCell(0, 0);
				onKeyDown(oEvent);
				const oDate = new Asc.cDate();
				equal(getCellText(), oDate.getDateString(editor), 'Check insert current date');

			});
			QUnit.test('Test add time', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				const oEvent = createEvent(186, true, true, false, false, false);
				const oDate = new Asc.cDate();
				onKeyDown(oEvent);
				equal(getCellText(), oDate.getTimeString(editor).split(' ').join(':00 '), 'Check insert current time');

			});


			QUnit.test('Test remove active cell text', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				enterText('hello World');
				onKeyDown(createEvent(8, false, false, false, false, false));
				equal(getCellText(), '', 'Check remove active cell');
			});


			QUnit.test('Test empty range', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveAndEnterText('Hello World', 0, 0);
				moveAndEnterText('Hello World', 1, 1);
				moveAndEnterText('Hello World', 2, 2);
				moveToCell(0, 0);
				selectToCell(5, 5);
				onKeyDown(createEvent(46, false, false, false, false, false));
				const arrSteps = [];
				closeCellEditor();
				arrSteps.push(moveAndGetCellText(0, 0));
				arrSteps.push(moveAndGetCellText(1, 1));
				arrSteps.push(moveAndGetCellText(2, 2));
				deep(arrSteps, ['', '', ''], 'Check empty shortcut');
			});


			QUnit.test('Test move active cell to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 1);
				onKeyDown(9, false, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 0), 'Check move left active cell');
			});

			QUnit.test('Test move active cell to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(9, false, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(1, 0), 'Check move right active cell');
			});

			QUnit.test('Test move active cell to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(1, 0);
				onKeyDown(13, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 2), 'Check move down active cell');
			});

			QUnit.test('Test move active cell to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(1, 0);
				onKeyDown(13, false, true, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 0), 'Check move up active cell');
			});

			QUnit.test('Test reset', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test disable num lock', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				AscCommon.AscBrowser.isOpera = true;
				testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert);

				AscCommon.AscBrowser.isOpera = false;
				testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert, true);

			});

			QUnit.test('Test disable scroll lock', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				AscCommon.AscBrowser.isOpera = true;
				testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert);

				AscCommon.AscBrowser.isOpera = false;
				testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert, true);
			});

			QUnit.test('Test select column', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(32, true, false, false, false, false);
				deep(cleanSelection(), checkRange(0, 1048575, 0, 0), 'Check move up');
			});

			QUnit.test('Test select row', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				onKeyDown(32, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 0, 0, 16383), 'Check move up');
			});

			QUnit.test('Test select sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				onKeyDown(32, true, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 1048575, 0, 16383), 'Check move up');
			});

			QUnit.test('Test add separator', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(110, false, false, false, false, false);
				equal(getCellText(), '.');
			});


			QUnit.test('Test go to previous sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				const sName = createWorksheet();
				onKeyDown(33, false, false, true, false, false);
				equal(wbView().wsActive, 0, 'Check got to previous worksheet');
				goToSheet(1);
				removeCurrentWorksheet();
			});

			QUnit.test('Test move to top cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(38, 0);
				onKeyDown(33, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 1), 'Check move to top cell');
			});

			QUnit.test('Test move to next sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				const sName = createWorksheet();
				goToSheet(0);
				onKeyDown(33, false, false, true, false, false);
				equal(wbView().wsActive, 1, 'Check got to next worksheet');
				removeCurrentWorksheet();
			});


			QUnit.test('Test move to bottom cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				onKeyDown(33, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 37), 'Check move to top cell');
			});
			QUnit.test('Test move to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 39);
				onKeyDown(37, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 16), 'check move left');
			});

			QUnit.test('Test select to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 39);
				onKeyDown(37, true, false, false, false, false);
				deep(cleanSelection(), checkRange(0, 0, 0, 0), 'check move left');
			});

			QUnit.test('Test move to left cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 1);
				onKeyDown(37, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 0), 'check move left');
			});

			QUnit.test('Test select to left cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 1);
				onKeyDown(37, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 0, 0, 1), 'check move left');
			});

			QUnit.test('Test move to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				onKeyDown(39, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 23), 'check move left');
			});

			QUnit.test('Test select to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll()
				moveToCell(0, 0);
				onKeyDown(39, true, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 0, 0, 23), 'check move left');
			});

			QUnit.test('Test move to right cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(39, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 1), 'check move left');
			});

			QUnit.test('Test select to right cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(39, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 0, 0, 1), 'check move left');
			});

			QUnit.test('Test move to top cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(25, 0);
				enterText('Hello');
				moveToCell(27, 0);
				enterText('Hello');
				moveToCell(35, 0);

				onKeyDown(38, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(27, 0));

				onKeyDown(38, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(25, 0));
			});

			QUnit.test('Test select to top cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(25, 0);
				enterText('Hello');
				moveToCell(27, 0);
				enterText('Hello');
				moveToCell(35, 0);

				onKeyDown(38, true, true, false, false, false);
				deep(cleanSelection(), checkRange(27, 35, 0, 0));

				onKeyDown(38, true, true, false, false, false);
				deep(cleanSelection(), checkRange(25, 35, 0, 0));
			});

			QUnit.test('Test move to up cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(5, 0);
				onKeyDown(38, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(4, 0));
				onKeyDown(38, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(3, 0));
				onKeyDown(38, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(2, 0));
				onKeyDown(38, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(1, 0));
			});

			QUnit.test('Test select to up cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(5, 0);
				onKeyDown(38, false, true, false, false, false);
				deep(cleanSelection(), checkRange(5, 4, 0, 0));
				onKeyDown(38, false, true, false, false, false);
				deep(cleanSelection(), checkRange(5, 3, 0, 0));
				onKeyDown(38, false, true, false, false, false);
				deep(cleanSelection(), checkRange(5, 2, 0, 0));
				onKeyDown(38, false, true, false, false, false);
				deep(cleanSelection(), checkRange(5, 1, 0, 0));
			});

			QUnit.test('Test move to bottom cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(45, 0);
				enterText('Hello');
				moveToCell(47, 0);
				enterText('Hello');
				moveToCell(42, 0);

				onKeyDown(40, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(45, 0));
				onKeyDown(40, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(47, 0));
			});

			QUnit.test('Test select to bottom cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll();
				moveToCell(45, 0);
				enterText('Hello');
				moveToCell(47, 0);
				enterText('Hello');
				moveToCell(42, 0);

				onKeyDown(40, true, true, false, false, false);
				deep(cleanSelection(), checkRange(42, 45, 0, 0));
				onKeyDown(40, true, true, false, false, false);
				deep(cleanSelection(), checkRange(42, 47, 0, 0));
			});

			QUnit.test('Test move to down cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				onKeyDown(40, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(1, 0));
				onKeyDown(40, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(2, 0));
				onKeyDown(40, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(3, 0));
				onKeyDown(40, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(4, 0));
			});

			QUnit.test('Test select to down cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				onKeyDown(40, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 1, 0, 0));
				onKeyDown(40, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 2, 0, 0));
				onKeyDown(40, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 3, 0, 0));
				onKeyDown(40, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 4, 0, 0));
			});


			QUnit.test('Test move to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(5, 25);
				onKeyDown(36, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(5, 0));
			});

			QUnit.test('Test select to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll();
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(5, 25);

				onKeyDown(36, false, true, false, false, false);
				deep(cleanSelection(), checkRange(5, 5, 0, 25));
			});

			QUnit.test('Test move to left edge top', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(5, 25);

				onKeyDown(36, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 0));
			});

			QUnit.test('Test select to left edge top', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(5, 25);

				onKeyDown(36, true, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 5, 0, 25));
			});

			QUnit.test('Test move to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll()
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(0, 0);

				onKeyDown(35, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 23));

				moveToCell(4, 0);
				onKeyDown(35, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(4, 8));

				moveToCell(5, 0);
				onKeyDown(35, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(5, 5));
			});

			QUnit.test('Test select to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll()
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(0, 0);

				onKeyDown(35, false, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 0, 0, 23));

				moveToCell(4, 0);
				onKeyDown(35, false, true, false, false, false);
				deep(cleanSelection(), checkRange(4, 4, 0, 8));

				moveToCell(5, 0);
				onKeyDown(35, false, true, false, false, false);
				deep(cleanSelection(), checkRange(5, 5, 0, 5));
			});

			QUnit.test('Test move to right bottom edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll()
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(0, 0);

				onKeyDown(35, true, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(5, 8));
			});

			QUnit.test('Test select to right bottom edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll()
				moveAndEnterText('Hello', 5, 5);
				moveAndEnterText('Hello', 4, 8);
				moveToCell(0, 0);

				onKeyDown(35, true, true, false, false, false);
				deep(cleanSelection(), checkRange(0, 5, 0, 8));
			});

			QUnit.test('Test set number format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveAndEnterText('49990', 5, 5);
				onKeyDown(49, true, true, false, false, false);
				equal(getCellText(), '49990.00', 'set number format');
			});


			QUnit.test('Test set time format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(5, 5);
				setCellFormat(Asc.c_oAscNumFormatType.General);
				enterText('49990');
				onKeyDown(50, true, true, false, false, false);
				equal(getCellText(), '12:00:00 AM', 'set number format');
			});

			QUnit.test('Test set date format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(5, 5);
				setCellFormat(Asc.c_oAscNumFormatType.General);
				enterText('49990');
				onKeyDown(51, true, true, false, false, false);
				equal(getCellText(), '11/11/2036', 'set number format');
			});

			QUnit.test('Test set currency format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(5, 5);
				setCellFormat(Asc.c_oAscNumFormatType.General);
				enterText('49990');
				onKeyDown(52, true, true, false, false, false);
				equal(getCellText(), '$49,990.00', 'set number format');
			});

			QUnit.test('Test set percent format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll();
				moveToCell(5, 5);
				setCellFormat(Asc.c_oAscNumFormatType.General);
				enterText('0.1');
				onKeyDown(53, true, true, false, false, false);
				equal(getCellText(), '10.00%', 'set number format');
			});


			QUnit.test('Test strikethrough', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				onKeyDown(53, true, false, false, false, false);
				equal(xfs().asc_getFontStrikeout(), true);
				onKeyDown(53, true, false, false, false, false);
				equal(xfs().asc_getFontStrikeout(), false);
			});

			QUnit.test('Test set exponential format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(5, 5);
				setCellFormat(Asc.c_oAscNumFormatType.General);
				enterText('0.1');
				onKeyDown(54, true, true, false, false, false);
				equal(getCellText(), '1.00E-01', 'set number format');
			});

			QUnit.test('select all', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				onKeyDown(65, true, false, false, false, false);
				deep(cleanSelection(), checkRange(0, 1048575, 0, 16383), 'Check move up');
			});

			QUnit.test('Test bold', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				onKeyDown(66, true, false, false, false, false);
				equal(xfs().asc_getFontBold(), true);
				onKeyDown(66, true, false, false, false, false);
				equal(xfs().asc_getFontBold(), false);
			});

			QUnit.test('Test italic', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				onKeyDown(73, true, false, false, false, false);
				equal(xfs().asc_getFontItalic(), true);
				onKeyDown(73, true, false, false, false, false);
				equal(xfs().asc_getFontItalic(), false);
			});


			QUnit.test('Test save', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

			});

			QUnit.test('Test underline', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				onKeyDown(85, true, false, false, false, false);
				equal(xfs().asc_getFontUnderline(), true);
				onKeyDown(85, true, false, false, false, false);
				equal(xfs().asc_getFontUnderline(), false);
			});


			QUnit.test('Test set general format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll();
				moveToCell(5, 5);
				enterText('0.1');
				setCellFormat(Asc.c_oAscNumFormatType.Time);

				onKeyDown(192, true, true, false, false, false);
				equal(getCellText(), '0.1', 'set number format');
			});


			QUnit.test('Test redo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				undo();
				onKeyDown(85, true, false, false, false, false);
				equal(getCellText(), '0.1');
			});

			QUnit.test('Test undo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				onKeyDown(85, true, false, false, false, false);
				equal(getCellText(), '');
			});

			QUnit.test('Test print', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				const oEvent = createEvent(80, true, false, false, false, false);
				executeTestWithCatchEvent('asc_onPrint', () => true, true, oEvent, oAssert);
			});

			QUnit.test('Test add sum function', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll();
				moveToCell(6, 6);
				onKeyDown(85, true, false, false, false, false);
				equal(getCellText(), false);
			});

			QUnit.test('Test context menu', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				const oEvent = createEvent(93, false, false, false, false, false);
				executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);
			});


			QUnit.module("test cell editor shortcuts", {
				before    : function ()
				{
					goToSheet(0);
				},
				beforeEach: function ()
				{
					setCheckOpenCellEditor(false);
				},
				afterEach : function ()
				{
					if (!checkOpenCellEditor())
					{
						throw new Error('cell editor must be opened in cell editor module');
					}
				}
			});
			QUnit.test('Test close cell editor', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello');
				onKeyDown(27, false, false, false, false, false);
				equal(getCellText(), '');
			});
			QUnit.test('Test add new line', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello');
				onKeyDown(13, false, false, true, false, false);
				equal(cellEditor().textRender.getLinesCount(), 2);

				onKeyDown(13, false, false, true, false, false);
				equal(cellEditor().textRender.getLinesCount(), 3);

				onKeyDown(13, false, false, true, false, false);
				equal(cellEditor().textRender.getLinesCount(), 4);

				onKeyDown(13, false, false, true, false, false);
				equal(cellEditor().textRender.getLinesCount(), 5);
			});

			QUnit.test('Try close editor', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello');
				onKeyDown(13, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(1, 0));

				moveToCell(0, 0);
				equal(getCellText(), 'Hello');
			});

			QUnit.test('Test sync and close editor', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello');
				onKeyDown(9, false, false, false, false, false);
				deep(cleanActiveCell(), checkActiveCell(0, 1));
				moveToCell(0, 0);
				equal(getCellText(), 'Hello');

			});
			QUnit.test('Test remove char back', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello');
				onKeyDown(8, false, false, false, false, false);
				equal(getCellText(), 'Hell');
			});
			QUnit.test('Test remove word back', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World');
				onKeyDown(8, true, false, false, false, false);
				equal(getCellText(), 'Hello ');
			});

			QUnit.test('Test add space in cell editor', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				onKeyDown(32, true, false, false, false, false);
				equal(getCellText(), ' ');
			});

			QUnit.test('Test move cursor to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(35, false, false, false, false, false);
				equal(cellPosition(), 18);

				onKeyDown(35, true, false, false, false, false);
				equal(cellPosition(), 47);
			});

			QUnit.test('Test select to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(35, false, true, false, false, false);
				equal(getSelectionCellEditor(), 'Hello World Hello ');
				onKeyDown(35, true, true, false, false, false);
				equal(getSelectionCellEditor(), 'Hello World Hello World Hello World Hello World');
			});

			QUnit.test('Test move cursor to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToEndCellEditor();
				onKeyDown(36, false, false, false, false, false);
				equal(cellPosition(), 36);

				onKeyDown(36, true, false, false, false, false);
				equal(cellPosition(), 0);
			});
			QUnit.test('Test select to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToEndCellEditor();
				onKeyDown(36, false, true, false, false, false);
				equal(getSelectionCellEditor(), 'Hello World');

				onKeyDown(36, true, true, false, false, false);
				equal(getSelectionCellEditor(), 'Hello World Hello World Hello World Hello World');
			});
			QUnit.test('Test move cursor to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToEndCellEditor();
				onKeyDown(37, false, false, false, false, false);
				equal(cellPosition(), 46);

				onKeyDown(37, true, false, false, false, false);
				equal(cellPosition(), 42);
			});
			QUnit.test('Test select to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToEndCellEditor();
				onKeyDown(37, false, true, false, false, false);
				equal(getSelectionCellEditor(), 'd');

				onKeyDown(37, true, true, false, false, false);
				equal(getSelectionCellEditor(), 'World');

			});
			QUnit.test('Test move cursor to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				onKeyDown(38, false, false, false, false, false);
				equal(cellPosition(), 29);
			});
			QUnit.test('Test select to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				onKeyDown(38, false, true, false, false, false);
				equal(getSelectionCellEditor(), ' World Hello World');

			});
			QUnit.test('Test move cursor to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(39, false, false, false, false, false);
				equal(cellPosition(), 1);

				onKeyDown(39, true, false, false, false, false);
				equal(cellPosition(), 6);
			});
			QUnit.test('Test select to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(39, false, true, false, false, false);
				equal(getSelectionCellEditor(), 'H');

				onKeyDown(39, true, true, false, false, false);
				equal(getSelectionCellEditor(), 'Hello ');
			});

			QUnit.test('Test move cursor to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(40, false, false, false, false, false);
				equal(cellPosition(), 18);
			});
			QUnit.test('Test select to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				openCellEditor();
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(40, false, true, false, false, false);
				equal(getSelectionCellEditor(), 'Hello World Hello ');
			});

			QUnit.test('Test delete front', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(46, false, false, false, false, false);
				equal(getCellText(), 'ello World Hello World Hello World Hello World');

				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(46, true, false, false, false, false);
				equal(getCellText(), 'World Hello World Hello World Hello World');
			});

			QUnit.test('Test strikethrough', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(53, true, false, false, false, false);
				enterTextWithoutClose('hihih');
				const arrFragments = getFragments(0, 5);
				equal(arrFragments.length, 1);
				equal(arrFragments[0].format.getStrikeout(), true);
			});

			QUnit.test('Test select all text', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(65, true, false, false, false, false);

				equal(getSelectionCellEditor(), 'Hello World Hello World Hello World Hello World');
			});

			QUnit.test('Test bold', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(66, true, false, false, false, false);
				enterTextWithoutClose('hihih');
				const arrFragments = getFragments(0, 5);
				equal(arrFragments.length, 1);
				equal(arrFragments[0].format.getBold(), true);
			});

			QUnit.test('Test italic', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(73, true, false, false, false, false);
				enterTextWithoutClose('hihih');
				const arrFragments = getFragments(0, 5);
				equal(arrFragments.length, 1);
				equal(arrFragments[0].format.getItalic(), true);
			});

			QUnit.test('Test underline', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				enterTextWithoutClose('Hello World Hello World Hello World Hello World');
				moveToStartCellEditor();
				onKeyDown(85, true, false, false, false, false);
				enterTextWithoutClose('hihih');
				const arrFragments = getFragments(0, 5);
				equal(arrFragments.length, 1);
				equal(arrFragments[0].format.getUnderline(), Asc.EUnderline.underlineSingle);
			});

			QUnit.test('Test disable scroll lock', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				AscCommon.AscBrowser.isOpera = true;
				testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert);

				AscCommon.AscBrowser.isOpera = false;
				testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert, true);

			});
			QUnit.test('Test disable num lock', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				AscCommon.AscBrowser.isOpera = true;
				testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert);

				AscCommon.AscBrowser.isOpera = false;
				testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert, true);


			});

			QUnit.test('Test print', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				testPreventDefaultAndStopPropagation(createEvent(80, true, false, false, false, false), oAssert);
			});

			QUnit.test('Test undo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				enterTextWithoutClose('H');
				enterTextWithoutClose('e');
				enterTextWithoutClose('l');
				enterTextWithoutClose('l');
				enterTextWithoutClose('o');
				onKeyDown(90, true, false, false, false, false);
				equal(getCellText(), 'Hell');
			});

			QUnit.test('Test redo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				enterTextWithoutClose('H');
				enterTextWithoutClose('e');
				enterTextWithoutClose('l');
				enterTextWithoutClose('l');
				enterTextWithoutClose('o');
				cellEditor().undo();
				onKeyDown(89, true, false, false, false, false);
				equal(getCellText(), 'Hello');
			});

			QUnit.test('Test add separator', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				onKeyDown(110, false, false, false, false, false);
				equal(getCellText(), '.');
			});
			QUnit.test('Test disable F2', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();

				AscCommon.AscBrowser.isOpera = true;
				testPreventDefaultAndStopPropagation(createEvent(113, false, false, false, false, false), oAssert);

				AscCommon.AscBrowser.isOpera = false;
				testPreventDefaultAndStopPropagation(createEvent(113, false, false, false, false, false), oAssert, true);
			});

			QUnit.test('Test switch reference', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				enterTextWithoutClose('=F4');

				onKeyDown(115, false, false, false, false, false);
				selectAllCell();
				equal(getSelectionCellEditor(), '=$F$4');
			});

			QUnit.test('Test add time', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				moveToCell(0, 0);
				openCellEditor();
				onKeyDown(186, true, true, false, false);
				const oDate = new Asc.cDate();
				equal(getCellText(), oDate.getTimeString(editor).split(' ').join(':00 '), 'Check insert current time');
			});

			QUnit.test('Test add date', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				moveToCell(0, 0);
				openCellEditor();
				onKeyDown(186, true, false, false, false);
				const oDate = new Asc.cDate();
				equal(getCellText(), oDate.getDateString(editor), 'Check insert current date');
			});


			// QUnit.module('test shortcuts');
			// QUnit.test('Test common shortcuts', function (oAssert)
			// {
			// 	let oEvent;
			// 	oEvent = createEvent(9, false, false, false, false, false);
			// 	editor.onKeyDown(oEvent);
			// 	moveRight();
			// 	cleanActiveCell(checkActiveCell((2,  0));
			//
			// 	moveToCell(5, 5);
			// 	cleanActiveCell(checkActiveCell((5,  5));
			// 	//ch-eck events controller
			// 	oEvent = createEvent(116, true, false, true, false, false);
			// 	checkRefreshAllConnections(oEvent, oAssert);
			// 	oEvent = createEvent(116, false, false, false, true, false);
			// 	//checkRefreshSelectedConnections(oEvent, oAssert);
			// 	oEvent = createEvent(82, true, true, false, false, false);
			// 	//checkChangeFormatTableInfo(oEvent, oAssert);
			// 	oEvent = createEvent(120, true, true, true, false, false);
			// 	checkCalculateAll(oEvent, oAssert);
			// 	oEvent = createEvent(120, true, false, true, false, false);
			// 	checkCalculateWorkbook(oEvent, oAssert);
			// 	oEvent = createEvent(120, false, true, false, false, false);
			// 	checkCalculateActiveSheet(oEvent, oAssert);
			// 	oEvent = createEvent(120, false, false, false, false, false);
			// 	checkCalculateOnlyChanged(oEvent, oAssert);
			// 	oEvent = createEvent(113, false, false, false, false, false);
			// 	checkFocusOnEditor(oEvent, oAssert);
			// 	oEvent = createEvent(186, true, false, false, false, false);
			// 	checkAddDateOrTime(oEvent, oAssert);
			// 	oEvent = createEvent(8, false, false, false, false, false);
			// 	checkRemoveCellText(oEvent, oAssert);
			// 	oEvent = createEvent(46, false, false, false, false, false);
			// 	checkEmpty(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, true, false, false, false);
			// 	checkMoveToLeftCell(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, false, false, false, false);
			// 	checkMoveToRightCell(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, true, false, false, false);
			// 	checkMoveToUpCell(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkMoveToDownCell(oEvent, oAssert);
			// 	oEvent = createEvent(27, false, false, false, false, false);
			// 	checkReset(oEvent, oAssert);
			// 	oEvent = createEvent(144, false, false, false, false, false);
			// 	checkDisableNumlock(oEvent, oAssert);
			// 	oEvent = createEvent(145, false, false, false, false, false);
			// 	checkDisableScrollLock(oEvent, oAssert);
			// 	oEvent = createEvent(32, true, false, false, false, false);
			// 	checkSelectColumn(oEvent, oAssert);
			// 	oEvent = createEvent(32, false, true, false, false, false);
			// 	checkSelectRow(oEvent, oAssert);
			// 	oEvent = createEvent(32, true, true, false, false, false);
			// 	checkSelectSheet(oEvent, oAssert);
			// 	oEvent = createEvent(32, false, false, false, false, false);
			// 	checkAddSpace(oEvent, oAssert);
			// 	oEvent = createEvent(110, false, false, false, false, false);
			// 	checkAddRegionalDecimal(oEvent, oAssert);
			// 	oEvent = createEvent(33, true, false, false, false, false);
			// 	checkGoToPreviousSheet(oEvent, oAssert);
			// 	oEvent = createEvent(33, false, false, true, false, false);
			// 	checkGoToPreviousSheet(oEvent, oAssert);
			// 	oEvent = createEvent(33, false, false, false, false, false);
			// 	checkMoveToTopCell(oEvent, oAssert);
			// 	oEvent = createEvent(34, true, false, false, false, false);
			// 	checkMoveToNextSheet(oEvent, oAssert);
			// 	oEvent = createEvent(34, false, false, true, false, false);
			// 	checkMoveToNextSheet(oEvent, oAssert);
			// 	oEvent = createEvent(34, false, false, false, false, false);
			// 	checkMoveToBottomCell(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, false, false, false, false);
			// 	checkMoveToLeftEdgeCell(oEvent, oAssert);
			// 	oEvent = createEvent(37, false, false, false, false, false);
			// 	checkMoveToLeftCell(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, false, false, false, false);
			// 	checkMoveToRightEdgeCell(oEvent, oAssert);
			// 	oEvent = createEvent(39, false, false, false, false, false);
			// 	checkMoveToRightCell(oEvent, oAssert);
			// 	oEvent = createEvent(38, true, false, false, false, false);
			// 	checkMoveToTopCell(oEvent, oAssert);
			// 	oEvent = createEvent(38, false, false, false, false, false);
			// 	checkMoveToUpCell(oEvent, oAssert);
			// 	oEvent = createEvent(40, true, false, false, false, false);
			// 	checkMoveToBottomCell(oEvent, oAssert);
			// 	oEvent = createEvent(40, false, false, false, false, false);
			// 	checkMoveToDownCell(oEvent, oAssert);
			// 	oEvent = createEvent(36, false, false, false, false, false);
			// 	checkMoveToLeftEdgeCell(oEvent, oAssert);
			// 	oEvent = createEvent(36, true, false, false, false, false);
			// 	checkMoveToLeftEdgeTopCell(oEvent, oAssert);
			// 	oEvent = createEvent(35, false, false, false, false, false);
			// 	checkMoveToRightEdgeCell(oEvent, oAssert);
			// 	oEvent = createEvent(35, true, false, false, false, false);
			// 	checkMoveToRightEdgeBottomCell(oEvent, oAssert);
			// 	oEvent = createEvent(49, true, true, false, false, false);
			// 	checkSetNumberFormat(oEvent, oAssert);
			// 	oEvent = createEvent(50, true, true, false, false, false);
			// 	checkSetTimeFormat(oEvent, oAssert);
			// 	oEvent = createEvent(51, true, true, false, false, false);
			// 	checkSetDateTime(oEvent, oAssert);
			// 	oEvent = createEvent(52, true, true, false, false, false);
			// 	checkSetCurrencyFormat(oEvent, oAssert);
			// 	oEvent = createEvent(53, true, false, false, false, false);
			// 	checkStrikethrough(oEvent, oAssert);
			// 	oEvent = createEvent(54, true, true, false, false, false);
			// 	checkSetExponentialFormat(oEvent, oAssert);
			// 	oEvent = createEvent(66, true, false, false, false, false);
			// 	checkBold(oEvent, oAssert);
			// 	oEvent = createEvent(73, true, false, false, false, false);
			// 	checkItalic(oEvent, oAssert);
			// 	oEvent = createEvent(83, false, false, false, false, false);
			// 	checkSave(oEvent, oAssert);
			// 	oEvent = createEvent(85, true, false, false, false, false);
			// 	checkUnderline(oEvent, oAssert);
			// 	oEvent = createEvent(192, true, true, false, false, false);
			// 	checkSetGeneralFormat(oEvent, oAssert);
			// 	oEvent = createEvent(89, true, false, false, false, false);
			// 	checkRedo(oEvent, oAssert);
			// 	oEvent = createEvent(90, true, false, false, false, false);
			// 	checkUndo(oEvent, oAssert);
			// 	oEvent = createEvent(65, true, false, false, false, false);
			// 	checkSelectSheet(oEvent, oAssert);
			// 	oEvent = createEvent(80, true, false, false, false, false);
			// 	checkPrint(oEvent, oAssert);
			// 	oEvent = createEvent(61, false, false, true, false, false);
			// 	checkAddSumFunction(oEvent, oAssert);
			// 	oEvent = createEvent(187, false, false, true, false, false);
			// 	checkAddSumFunction(oEvent, oAssert);
			//
			// 	//check for mac os
			// 	oEvent = createEvent(61, true, false, true, false, false);
			// 	checkAddSumFunction(oEvent, oAssert);
			// 	oEvent = createEvent(187, true, false, true, false, false);
			// 	checkAddSumFunction(oEvent, oAssert);
			//
			// 	oEvent = createEvent(93, false, false, false, false, false);
			// 	checkContextMenu(oEvent, oAssert);
			//
			// 	//ch-eckChangeSelectionInOleEditor
			//
			// 	function checkAddRegionalDecimal()
			// 	{
			//
			// 	}
			//
			// 	function checkAddNewLineMenuEditorMode()
			// 	{
			//
			// 	}
			//
			// 	function checkCloseEditorAndMoveDown()
			// 	{
			//
			// 	}
			//
			// 	function checkCloseEditorAndMoveUp()
			// 	{
			//
			// 	}
			//
			// 	function checkCloseEditorAndMoveRight()
			// 	{
			//
			// 	}
			//
			// 	function checkAddRegionalDecimalCellEditor()
			// 	{
			//
			// 	}
			//
			// 	function checkRemoveCharShape()
			// 	{
			//
			// 	}
			//
			// 	function checkRemoveWordShape()
			// 	{
			//
			// 	}
			//
			// 	function checkMoveToStartLineContent()
			// 	{
			//
			// 	}
			//
			// 	function checkRemoveContentCharFrontShape()
			// 	{
			//
			// 	}
			//
			//
			// 	//ch-eck cell editor shortcuts
			// 	oEvent = createEvent(27, false, false, false, false, false);
			// 	checkCloseCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, true, false, false);
			// 	checkAddNewLine(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkAddNewLineMenuEditorMode(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkCloseEditorAndMoveDown(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, true, false, false, false);
			// 	checkCloseEditorAndMoveUp(oEvent, oAssert);
			// 	oEvent = createEvent(13, true, true, false, false, false);
			// 	checkCloseCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(13, true, false, false, false, false);
			// 	checkCloseCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, false, false, false, false);
			// 	checkCloseEditorAndMoveRight(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, true, false, false, false);
			// 	checkCloseEditorAndMoveRight(oEvent, oAssert);
			// 	oEvent = createEvent(8, false, false, false, false, false);
			// 	checkRemoveCharBack(oEvent, oAssert);
			// 	oEvent = createEvent(8, true, false, false, false, false);
			// 	checkRemoveWordBack(oEvent, oAssert);
			// 	oEvent = createEvent(32, false, false, false, false, false);
			// 	checkAddSpace(oEvent, oAssert);
			// 	oEvent = createEvent(35, false, false, false, false, false);
			// 	checkMoveToEndLine(oEvent, oAssert);
			// 	oEvent = createEvent(35, true, false, false, false, false);
			// 	checkMoveToEndPos(oEvent, oAssert);
			// 	oEvent = createEvent(35, false, true, false, false, false);
			// 	checkSelectToEndLine(oEvent, oAssert);
			// 	oEvent = createEvent(35, true, true, false, false, false);
			// 	checkSelectToEndPos(oEvent, oAssert);
			// 	oEvent = createEvent(36, false, false, false, false, false);
			// 	checkMoveToStartLine(oEvent, oAssert);
			// 	oEvent = createEvent(36, true, false, false, false, false);
			// 	checkMoveToStartPos(oEvent, oAssert);
			// 	oEvent = createEvent(36, false, true, false, false, false);
			// 	checkSelectToStartLine(oEvent, oAssert);
			// 	oEvent = createEvent(36, true, true, false, false, false);
			// 	checkSelectToStartPos(oEvent, oAssert);
			// 	oEvent = createEvent(37, false, false, false, false, false);
			// 	checkMoveToLeftChar(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, false, false, false, false);
			// 	checkMoveToLeftWord(oEvent, oAssert);
			// 	oEvent = createEvent(37, false, true, false, false, false);
			// 	checkSelectToLeftChar(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, true, false, false, false);
			// 	checkSelectToLeftWord(oEvent, oAssert);
			// 	oEvent = createEvent(38, false, false, false, false, false);
			// 	checkMoveToUpLine(oEvent, oAssert);
			// 	oEvent = createEvent(38, false, true, false, false, false);
			// 	checkSelectToUpLine(oEvent, oAssert);
			// 	oEvent = createEvent(39, false, false, false, false, false);
			// 	checkMoveToRightChar(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, false, false, false, false);
			// 	checkMoveToRightWord(oEvent, oAssert);
			// 	oEvent = createEvent(39, false, true, false, false, false);
			// 	checkSelectToRightChar(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, true, false, false, false);
			// 	checkSelectToRightWord(oEvent, oAssert);
			// 	oEvent = createEvent(40, false, false, false, false, false);
			// 	checkMoveToDownLine(oEvent, oAssert);
			// 	oEvent = createEvent(40, false, true, false, false, false);
			// 	checkSelectToDownLine(oEvent, oAssert);
			// 	oEvent = createEvent(46, false, false, false, false, false);
			// 	checkRemoveFrontChar(oEvent, oAssert);
			// 	oEvent = createEvent(46, true, false, false, false, false);
			// 	checkRemoveFrontWord(oEvent, oAssert);
			// 	oEvent = createEvent(53, true, false, false, false, false);
			// 	checkStrikethroughCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(65, true, false, false, false, false);
			// 	checkSelectAllCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(66, true, false, false, false, false);
			// 	checkBoldCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(73, true, false, false, false, false);
			// 	checkItalicCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(85, true, false, false, false, false);
			// 	checkUnderlineCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(144, false, false, false, false, false);
			// 	checkDisableScrollLockCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(145, false, false, false, false, false);
			// 	checkDisableNumLockCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(80, true, false, false, false, false);
			// 	checkPrintCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(90, false, false, false, false, false);
			// 	checkUndoCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(89, false, false, false, false, false);
			// 	checkRedoCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(110, false, false, false, false, false);
			// 	checkAddRegionalDecimalCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(113, false, false, false, false, false);
			// 	checkDisableF2InOperaCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(115, false, false, false, false, false);
			// 	checkSwitchReference(oEvent, oAssert);
			// 	oEvent = createEvent(186, true, true, false, false, false);
			// 	checkAddTimeCellEditor(oEvent, oAssert);
			// 	oEvent = createEvent(186, false, true, false, false, false);
			// 	checkAddDateCellEditor(oEvent, oAssert);
			//
			// 	// ch-eck common controllers shortcuts
			//
			// 	oEvent = createEvent(8, false, false, false, false, false);
			// 	checkRemoveShape(oEvent, oAssert);
			// 	oEvent = createEvent(8, false, false, false, false, false);
			// 	checkRemoveCharShape(oEvent, oAssert);
			// 	oEvent = createEvent(8, true, false, false, false, false);
			// 	checkRemoveWordShape(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, false, false, false, false);
			// 	checkAddTab(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, false, false, false, false);
			// 	checkSelectNextObject(oEvent, oAssert);
			// 	oEvent = createEvent(9, false, true, false, false, false);
			// 	checkSelectPreviousObject(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkVisitHyperlink(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkAddNewLineMath(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, true, false, false, false);
			// 	checkAddBreakLine(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkAddNewParagraph(oEvent, oAssert);
			//
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkCreateTxBoxContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkCreateTxBodyContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkMoveCursorToStartPosShape(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkSelectAllContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkMoveCursorToStartPosChartTitle(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkSelectAllContentChartTitle(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkRemoveSelectionCellGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(13, false, false, false, false, false);
			// 	checkMoveCursorToStartPosAndSelectAllTable(oEvent, oAssert);
			//
			// 	oEvent = createEvent(27, false, false, false, false, false);
			// 	checkStepRemoveSelection(oEvent, oAssert);
			// 	oEvent = createEvent(27, false, false, false, false, false);
			// 	checkResetAddShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(35, true, false, false, false, false);
			// 	checkMoveCursorToEndPositionContent(oEvent, oAssert);
			// 	oEvent = createEvent(35, false, false, false, false, false);
			// 	checkMoveCursorToEndLineContent(oEvent, oAssert);
			// 	oEvent = createEvent(35, false, true, false, false, false);
			// 	checkSelectToEndLineContent(oEvent, oAssert);
			// 	oEvent = createEvent(36, true, false, false, false, false);
			// 	checkMoveCursorToStartPositionContent(oEvent, oAssert);
			// 	oEvent = createEvent(36, false, false, false, false, false);
			// 	checkMoveToStartLineContent(oEvent, oAssert);
			// 	oEvent = createEvent(36, false, true, false, false, false);
			// 	checkSelectToStartLineContent(oEvent, oAssert);
			//
			// 	oEvent = createEvent(37, false, false, false, false, false);
			// 	checkMoveCursorLeftCharContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, false, false, false, false);
			// 	checkMoveCursorLeftWordContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(37, false, true, false, false, false);
			// 	checkSelectLeftCharContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, true, false, false, false);
			// 	checkSelectLeftWordContentGraphicFrame(oEvent, oAssert);
			//
			// 	oEvent = createEvent(37, false, false, false, false, false);
			// 	checkMoveCursorLeftCharContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, false, false, false, false);
			// 	checkMoveCursorLeftWordContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(37, false, true, false, false, false);
			// 	checkSelectLeftCharContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(37, true, true, false, false, false);
			// 	checkSelectLeftWordContentShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(37, true, false, false, false, false);
			// 	checkSmallMoveLeftShape(oEvent, oAssert);
			// 	oEvent = createEvent(37, false, false, false, false, false);
			// 	checkBigMoveLeftShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(39, false, false, false, false, false);
			// 	checkMoveCursorRightCharContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, false, false, false, false);
			// 	checkMoveCursorRightWordContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(39, false, true, false, false, false);
			// 	checkSelectRightCharContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, true, false, false, false);
			// 	checkSelectRightWordContentGraphicFrame(oEvent, oAssert);
			//
			// 	oEvent = createEvent(39, false, false, false, false, false);
			// 	checkMoveCursorRightCharContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, false, false, false, false);
			// 	checkMoveCursorRightWordContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(39, false, true, false, false, false);
			// 	checkSelectRightCharContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(39, true, true, false, false, false);
			// 	checkSelectRightWordContentShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(39, true, false, false, false, false);
			// 	checkSmallMoveRightShape(oEvent, oAssert);
			// 	oEvent = createEvent(39, false, false, false, false, false);
			// 	checkBigMoveRightShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(38, false, false, false, false, false);
			// 	checkMoveCursorUpCharContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(38, false, true, false, false, false);
			// 	checkSelectUpCharContentGraphicFrame(oEvent, oAssert);
			//
			// 	oEvent = createEvent(38, false, false, false, false, false);
			// 	checkMoveCursorUpCharContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(38, false, true, false, false, false);
			// 	checkSelectUpCharContentShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(38, true, false, false, false, false);
			// 	checkSmallMoveUpShape(oEvent, oAssert);
			// 	oEvent = createEvent(38, false, false, false, false, false);
			// 	checkBigMoveUpShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(40, false, false, false, false, false);
			// 	checkMoveCursorDownCharContentGraphicFrame(oEvent, oAssert);
			// 	oEvent = createEvent(40, false, true, false, false, false);
			// 	checkSelectDownCharContentGraphicFrame(oEvent, oAssert);
			//
			// 	oEvent = createEvent(40, false, false, false, false, false);
			// 	checkMoveCursorDownCharContentShape(oEvent, oAssert);
			// 	oEvent = createEvent(40, false, true, false, false, false);
			// 	checkSelectDownCharContentShape(oEvent, oAssert);
			//
			// 	oEvent = createEvent(40, true, false, false, false, false);
			// 	checkSmallMoveDownShape(oEvent, oAssert);
			// 	oEvent = createEvent(40, false, false, false, false, false);
			// 	checkBigMoveDownShape(oEvent, oAssert);
			//
			// 	function checkRemoveContentWordFrontShape()
			// 	{
			//
			// 	}
			//
			// 	//TODO: check remove more
			// 	oEvent = createEvent(46, false, false, false, false, false);
			// 	checkRemoveShape(oEvent, oAssert);
			// 	oEvent = createEvent(46, false, false, false, false, false);
			// 	checkRemoveContentCharFrontShape(oEvent, oAssert);
			// 	oEvent = createEvent(46, true, false, false, false, false);
			// 	checkRemoveContentWordFrontShape(oEvent, oAssert);
			// 	oEvent = createEvent(65, true, false, false, false, false);
			// 	checkSelectAllShape(oEvent, oAssert);
			// 	oEvent = createEvent(66, true, false, false, false, false);
			// 	checkBoldShape(oEvent, oAssert);
			// 	oEvent = createEvent(67, false, false, true, false, false);
			// 	checkClearSlicer(oEvent, oAssert);
			// 	//macOs
			// 	oEvent = createEvent(67, true, false, true, false, false);
			// 	checkClearSlicer(oEvent, oAssert);
			// 	oEvent = createEvent(69, true, false, false, false, false);
			// 	checkSetCenterAlign(oEvent, oAssert);
			// 	oEvent = createEvent(73, true, false, false, false, false);
			// 	checkSetItalicShape(oEvent, oAssert);
			// 	oEvent = createEvent(74, true, false, false, false, false);
			// 	checkSetJustifyAlign(oEvent, oAssert);
			// 	oEvent = createEvent(76, true, false, false, false, false);
			// 	checkSetLeftAlign(oEvent, oAssert);
			// 	oEvent = createEvent(82, true, false, false, false, false);
			// 	checkSetRightAlign(oEvent, oAssert);
			// 	oEvent = createEvent(83, false, false, true, false, false);
			// 	checkSlicerMultiSelect(oEvent, oAssert);
			// 	oEvent = createEvent(83, true, false, true, false, false);
			// 	checkSlicerMultiSelect(oEvent, oAssert);
			// 	oEvent = createEvent(85, true, false, false, false, false);
			// 	checkUnderlineShape(oEvent, oAssert);
			// 	oEvent = createEvent(187, true, false, false, false, false);
			// 	checkSubscriptShape(oEvent, oAssert);
			// 	oEvent = createEvent(187, true, true, false, false, false);
			// 	checkSuperscriptShape(oEvent, oAssert);
			// 	oEvent = createEvent(188, true, false, false, false, false);
			// 	checkSuperscriptShape(oEvent, oAssert);
			// 	oEvent = createEvent(189, true, true, false, false, false);
			// 	checkAddEnDash(oEvent, oAssert);
			// 	oEvent = createEvent(189, false, true, false, false, false);
			// 	checkAddLowLine(oEvent, oAssert);
			// 	oEvent = createEvent(189, false, false, false, false, false);
			// 	checkAddHyphenMinus(oEvent, oAssert);
			// 	oEvent = createEvent(190, true, false, false, false, false);
			// 	checkSubscriptShape(oEvent, oAssert);
			// 	oEvent = createEvent(219, true, false, false, false, false);
			// 	checkDecreaseFontSize(oEvent, oAssert);
			// 	oEvent = createEvent(221, true, false, false, false, false);
			// 	checkIncreaseFontSize(oEvent, oAssert);
			// });
		}
	)
})(window);

function checkRemoveShape(oEvent, oAssert)
{

}

function checkAddTab(oEvent, oAssert)
{

}

function checkSelectNextObject(oEvent, oAssert)
{

}

function checkSelectPreviousObject(oEvent, oAssert)
{

}

function checkVisitHyperlink(oEvent, oAssert)
{

}

function checkAddNewLineMath(oEvent, oAssert)
{

}

function checkAddBreakLine(oEvent, oAssert)
{

}

function checkAddNewLine(oEvent, oAssert)
{

}

function checkCreateTxBoxContentShape(oEvent, oAssert)
{

}

function checkCreateTxBodyContentShape(oEvent, oAssert)
{

}

function checkMoveCursorToStartPosShape(oEvent, oAssert)
{

}

function checkSelectAllContentShape(oEvent, oAssert)
{

}

function checkMoveCursorToStartPosChartTitle(oEvent, oAssert)
{

}

function checkSelectAllContentChartTitle(oEvent, oAssert)
{

}

function checkRemoveSelectionCellGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorToStartPosAndSelectAllTable(oEvent, oAssert)
{

}

function checkStepRemoveSelection(oEvent, oAssert)
{

}

function checkResetAddShape(oEvent, oAssert)
{

}

function checkMoveCursorToEndPositionContent(oEvent, oAssert)
{

}

function checkMoveCursorToEndLineContent(oEvent, oAssert)
{

}

function checkSelectToEndLineContent(oEvent, oAssert)
{

}

function checkMoveCursorToStartPositionContent(oEvent, oAssert)
{

}

function checkSelectToStartLineContent(oEvent, oAssert)
{

}

function checkMoveCursorLeftCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorLeftWordContentGraphicFrame(oEvent, oAssert)
{

}

function checkSelectLeftCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkSelectLeftWordContentGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorLeftCharContentShape(oEvent, oAssert)
{

}

function checkMoveCursorLeftWordContentShape(oEvent, oAssert)
{

}

function checkSelectLeftCharContentShape(oEvent, oAssert)
{

}

function checkSelectLeftWordContentShape(oEvent, oAssert)
{

}

function checkSmallMoveLeftShape(oEvent, oAssert)
{

}

function checkBigMoveLeftShape(oEvent, oAssert)
{

}

function checkMoveCursorRightCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorRightWordContentGraphicFrame(oEvent, oAssert)
{

}

function checkSelectRightCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkSelectRightWordContentGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorRightCharContentShape(oEvent, oAssert)
{

}

function checkMoveCursorRightWordContentShape(oEvent, oAssert)
{

}

function checkSelectRightCharContentShape(oEvent, oAssert)
{

}

function checkSelectRightWordContentShape(oEvent, oAssert)
{

}

function checkSmallMoveRightShape(oEvent, oAssert)
{

}

function checkBigMoveRightShape(oEvent, oAssert)
{

}

function checkMoveCursorUpCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkSelectUpCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorUpCharContentShape(oEvent, oAssert)
{

}

function checkSelectUpCharContentShape(oEvent, oAssert)
{

}

function checkSmallMoveUpShape(oEvent, oAssert)
{

}

function checkBigMoveUpShape(oEvent, oAssert)
{

}

function checkMoveCursorDownCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkSelectDownCharContentGraphicFrame(oEvent, oAssert)
{

}

function checkMoveCursorDownCharContentShape(oEvent, oAssert)
{

}

function checkSelectDownCharContentShape(oEvent, oAssert)
{

}

function checkSmallMoveDownShape(oEvent, oAssert)
{

}

function checkBigMoveDownShape(oEvent, oAssert)
{

}

function checkRemoveShape(oEvent, oAssert)
{

}

function checkSelectAllShape(oEvent, oAssert)
{

}

function checkBoldShape(oEvent, oAssert)
{

}

function checkClearSlicer(oEvent, oAssert)
{

}

function checkSetCenterAlign(oEvent, oAssert)
{

}

function checkSetItalicShape(oEvent, oAssert)
{

}

function checkSetJustifyAlign(oEvent, oAssert)
{

}

function checkSetLeftAlign(oEvent, oAssert)
{

}

function checkSetRightAlign(oEvent, oAssert)
{

}

function checkSlicerMultiSelect(oEvent, oAssert)
{

}

function checkUnderlineShape(oEvent, oAssert)
{

}

function checkSubscriptShape(oEvent, oAssert)
{

}

function checkSuperscriptShape(oEvent, oAssert)
{

}

function checkSuperscriptShape(oEvent, oAssert)
{

}

function checkAddEnDash(oEvent, oAssert)
{

}

function checkAddLowLine(oEvent, oAssert)
{

}

function checkAddHyphenMinus(oEvent, oAssert)
{

}

function checkSubscriptShape(oEvent, oAssert)
{

}

function checkDecreaseFontSize(oEvent, oAssert)
{

}

function checkIncreaseFontSize(oEvent, oAssert)
{

}