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
	const editor = Asc.editor;
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

	const pxToMm = AscCommon.g_dKoef_pix_to_mm;

	function addHyperlink(sLink, sText)
	{
		const oHyperlink = new AscCommonExcel.Hyperlink();
		oHyperlink.Hyperlink = sLink;

		const oAscHyperlink = new Asc.asc_CHyperlink(oHyperlink);
		oAscHyperlink.text = sText;
		editor.asc_insertHyperlink(oAscHyperlink);
	}

	function createMathInShape()
	{
		const {oShape, oParagraph} = createShapeWithContent('', true);
		selectOnlyObjects([oShape]);
		moveToParagraph(oParagraph, true);
		editor.asc_AddMath2(c_oAscMathType.FractionVertical);
		return {oShape, oParagraph};
	}

	function moveCursorRight(bAddToSelect, bWord)
	{
		graphicController().cursorMoveRight(bAddToSelect, bWord);
	}

	function moveCursorLeft(bAddToSelect, bWord)
	{
		graphicController().cursorMoveLeft(bAddToSelect, bWord);
	}

	function moveCursorUp(bAddToSelect, bWord)
	{
		graphicController().cursorMoveUp(bAddToSelect, bWord);
	}

	function moveCursorDown(bAddToSelect, bWord)
	{
		graphicController().cursorMoveDown(bAddToSelect, bWord);

	}

	function checkMoveContentShapeHelper(arrExpected, oEvent, oAssert)
	{
		const {deep, equal} = createTest(oAssert);
		const {oShape} = createShapeWithContent('Hello World Hello World Hello World Hello World Hello World');
		moveToShapeParagraph(oShape, 0);
		onKeyDown(oEvent);
		equal(contentPosition(), arrExpected[0]);

		moveToShapeParagraph(oShape, 0, true);
		onKeyDown(oEvent);
		equal(contentPosition(), arrExpected[1]);
	}

	function contentPosition()
	{
		const arrContentPosition = graphicController().getTargetDocContent().GetContentPosition();
		return arrContentPosition[arrContentPosition.length - 1].Position;
	}

	function selectedContent()
	{
		return graphicController().GetSelectedText(false, {TabSymbol: '\t'});
	}

	function checkSelectContentShapeHelper(arrExpected, oEvent, oAssert)
	{
		const {deep, equal} = createTest(oAssert);
		const {oShape} = createShapeWithContent('Hello World Hello World Hello World Hello World Hello World');
		moveToShapeParagraph(oShape, 0, true);
		onKeyDown(oEvent);
		equal(selectedContent(), arrExpected[0]);

		moveToShapeParagraph(oShape, 0);
		onKeyDown(oEvent);
		equal(selectedContent(), arrExpected[1]);
	}

	function round(nNumber, nAmount)
	{
		const nPower = Math.pow(10, nAmount);
		return Math.round(nNumber * nPower) / nPower;
	}

	function moveShapeHelper(oExpected, oEvent, oAssert)
	{
		const oShape = createShape();
		const {deep, equal} = createTest(oAssert);
		selectOnlyObjects([oShape]);

		onKeyDown(oEvent);
		oShape.recalculateTransform();
		deep({x: round(oShape.x, 12), y: round(oShape.y, 12)}, {x: round(oExpected.x, 12), y: round(oExpected.y, 12)});
	}

	function createTable()
	{
		moveToCell(0, 0);
		enterText('Hello');

		moveToCell(1, 0);
		enterText('1');
		moveToCell(2, 0);
		enterText('2');
		moveToCell(3, 0);
		enterText('3');
		moveToCell(4, 0);
		enterText('4');
		moveToCell(0, 0);
		selectToCell(4, 0);

		editor.asc_addAutoFilter('TableStyleMedium2', editor.asc_getAddFormatTableOptions());
	}
	function createSlicer()
	{
		createTable();
		editor.asc_insertSlicer(['Hello']);
		return graphicController().getSelectedArray()[0];
	}

	function moveToShapeParagraph(oShape, nParagraph, bIsMoveToStart)
	{
		const oParagraph = oShape.getDocContent().Content[nParagraph];
		moveToParagraph(oParagraph, bIsMoveToStart);

		return oParagraph;
	}


	function addProperty(oPr)
	{
		graphicController().paragraphAdd(new AscCommonWord.ParaTextPr(oPr), true);
	}

	function targetContent()
	{
		return graphicController().getTargetDocContent();
	}

	function textPr()
	{
		return graphicController().getParagraphTextPr();
	}

	function paraPr()
	{
		return graphicController().getParagraphParaPr();
	}

	function moveToParagraph(oParagraph, bToStart)
	{
		oParagraph.SetThisElementCurrent();
		if (bToStart)
		{
			oParagraph.MoveCursorToStartPos();
		} else
		{
			oParagraph.MoveCursorToEndPos();
		}

		graphicController().recalculateCurPos(true, true);
	}

	function selectedObjects()
	{
		return graphicController().getSelectedArray();
	}

	function cleanGraphic()
	{
		graphicController().resetSelection();
		graphicController().selectAll();
		graphicController().remove();
	}

	function graphicController()
	{
		return editor.getGraphicController();
	}

	function createShape()
	{
		const oGraphicController = graphicController();
		const oTrack = new AscFormat.NewShapeTrack('rect', 0, 0, wbModel().theme, null, null, null, 0, oGraphicController);
		const oShape = oTrack.getShape(false, drawingDocument(), oGraphicController.drawingObjects);
		oShape.setWorksheet(wbView().getWorksheet().model);
		oShape.spPr.xfrm.setOffX(100);
		oShape.spPr.xfrm.setOffY(100);
		oShape.spPr.xfrm.setExtX(100);
		oShape.spPr.xfrm.setExtY(100);
		oShape.setPaddings({Left:0,Right:0,Top:0,Bottom:0});
		oShape.setParent(oGraphicController.drawingObjects);
		oShape.setRecalculateInfo();

		oShape.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
		oShape.checkDrawingBaseCoords();
		oGraphicController.checkChartTextSelection();
		oGraphicController.resetSelection();
		oShape.select(oGraphicController, 0);
		oGraphicController.startRecalculate();
		oGraphicController.drawingObjects.sendGraphicObjectProps();
		oGraphicController.recalculate();
		oShape.recalculateTransform();

		return oShape;
	}

	function drawingDocument()
	{
		return wbModel().DrawingDocument;
	}

	function selectOnlyObjects(arrShapes)
	{
		const oController = graphicController();
		oController.resetSelection();
		for (let i = 0; i < arrShapes.length; i += 1)
		{
			const oObject = arrShapes[i];
			if (oObject.group)
			{
				const oMainGroup = oObject.group.getMainGroup();
				oMainGroup.select(oController, 0);
				oController.selection.groupSelection = oMainGroup;
			}
			oObject.select(oController, 0);
		}
	}

	function createShapeWithContent(sText)
	{
		const oShape = createShape();
		selectOnlyObjects([oShape])
		oShape.setTxBody(AscFormat.CreateTextBodyFromString(sText, drawingDocument(), oShape));
		const oParagraph = oShape.txBody.content.Content[0];
		oShape.recalculateContentWitCompiledPr();
		return {oShape, oParagraph};
	}

	function checkTextAfterKeyDownHelperHelloWorld(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, 'Hello World');
	}

	function getParagraphText(oParagraph)
	{
		return oParagraph.GetText({ParaEndToSpace: false});
	}

	function checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, sInitText)
	{
		const {oParagraph} = createShapeWithContent(sInitText);
		moveToParagraph(oParagraph);
		onKeyDown(oEvent);
		const sTextAfterKeyDown = getParagraphText(oParagraph);
		oAssert.strictEqual(sTextAfterKeyDown, sCheckText, sPrompt);
	}

	function checkDirectTextPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = createShapeWithContent('Hello World');
		let oTextPr = getDirectTextPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oTextPr), nExpectedValue, sPrompt);
		return function continueCheck(fCallback2, oEvent2, oAssert2, nExpectedValue2, sPrompt2)
		{
			oTextPr = getDirectTextPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oTextPr), nExpectedValue2, sPrompt2);
			return continueCheck;
		}
	}

	function checkDirectParaPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = createShapeWithContent('Hello World');
		let oParaPr = getDirectParaPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oParaPr), nExpectedValue, sPrompt);

		return function continueCheck(fCallback2, oEvent2, oAssert2, nExpectedValue2, sPrompt2)
		{
			oParaPr = getDirectParaPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oParaPr), nExpectedValue2, sPrompt2);
			return continueCheck;
		}
	}


	function getDirectTextPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		graphicController().selectAll();
		onKeyDown(oEvent);
		return textPr();
	}

	function getDirectParaPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		graphicController().selectAll();
		onKeyDown(oEvent);
		return paraPr();
	}

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
				createTable();
				moveToCell(4, 0);
				onKeyDown(82, true, true, false, false, false);
				moveToCell(5, 0);
				equal(getCellText(), '10');
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
				cleanAll();
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
				editor.asc_startAddShape('rect');
				onKeyDown(27, false, false, false, false);
				equal(graphicController().checkEndAddShape(), false);
				moveToCell(0, 0);
				editor.asc_SelectionCut();
				onKeyDown(27, false, false, false, false);
				equal(wsView().copyCutRange, null);

				moveToCell(0, 0);
				editor.asc_Copy();
				onKeyDown(27, false, false, false, false);
				equal(wsView().copyCutRange, null);
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

			QUnit.test('Test move to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				cleanAll();
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
				onKeyDown(89, true, false, false, false, false);
				equal(getCellText(), '0.1');
			});

			QUnit.test('Test undo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);

				cleanAll();
				moveToCell(6, 6);
				enterText('0.1');
				onKeyDown(90, true, false, false, false, false);
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
				goToSheet(0);
				moveToCell(0, 0);
				enterText('1');
				moveToCell(0, 1);
				enterText('2');
				moveToCell(7, 7);
				setCellFormat(Asc.c_oAscNumFormatType.General);
				onKeyDown(61, false, false, true, false, false);
				enterTextWithoutClose('A1,B1');
				equal(getCellText(), "3");
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
					goToSheet(0);
					cleanAll();
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

			QUnit.module('Test graphic objects shortcuts');
			QUnit.test('Test remove back text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkTextAfterKeyDownHelperHelloWorld('Hello Worl', oEvent, oAssert, 'Check delete with backspace');
				}, oTypes.removeBackChar);
			});
			QUnit.test('Test remove back text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkTextAfterKeyDownHelperHelloWorld('Hello ', oEvent, oAssert, 'Check delete word with backspace')
				}, oTypes.removeBackWord);
			});

			function createChart()
			{
				moveAndEnterText('1', 0, 0);
				moveAndEnterText('2', 1, 0);
				moveAndEnterText('3', 2, 0);

				moveAndEnterText('1', 0, 1);
				moveAndEnterText('2', 1, 1);
				moveAndEnterText('3', 2, 1);

				selectToCell(0, 0);

				const oProps = editor.asc_getChartObject(true);
				oProps.changeType(0);
				editor.asc_addChartDrawingObject(oProps);
				return selectedObjects()[0];
			}
			QUnit.test('Test remove chart', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oChart = createChart();
					selectOnlyObjects([oChart]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oChart, ws().Drawings.map((el) => el.graphicObject)), "Check remove group");
				}, oTypes.removeChart);
			});
			function checkRemoveObject(oShape, arrShapes)
			{
				return !!(arrShapes.indexOf(oShape) === -1 && oShape.bDeleted);
			}
			QUnit.test('Test remove shape', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('', true);
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oShape, ws().Drawings.map((el) => el.graphicObject)), 'Check remove shape');
				}, oTypes.removeShape);
			});
			function createGroup(arrShapes)
			{
				graphicController().resetSelection();
				selectOnlyObjects(arrShapes);
				return graphicController().createGroup();

			}
			QUnit.test('Test remove group', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oGroup = createGroup([createShape(), createShape()]);
					selectOnlyObjects([oGroup]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oGroup, ws().Drawings.map((el) => el.graphicObject)), 'Check remove group');
				}, oTypes.removeGroup);
			});
			QUnit.test('Test remove shape in group', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oGroupedGroup = createGroup([createShape(), createShape()]);
					const oRemovedShape = createShape();
					const oGroup = createGroup([oGroupedGroup, oRemovedShape]);
					selectOnlyObjects([oRemovedShape]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oRemovedShape, oGroup.spTree), 'Check remove shape in group');
				}, oTypes.removeShapeInGroup);
			});

			QUnit.test('Test add tab', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oParagraph} = createShapeWithContent('');
					moveToParagraph(oParagraph);
					onKeyDown(oEvent);
					selectAll();
					equal(selectedContent(), '\t');
				}, oTypes.addTab);
			});
			QUnit.test('Test select next object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oShape1 = createShape();
					const oShape2 = createShape();
					const oShape3 = createShape();
					selectOnlyObjects([oShape1]);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape2, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape3, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape1, true);
				}, oTypes.selectNextObject);
			});

			QUnit.test('Test select previous object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oShape1 = createShape();
					const oShape2 = createShape();
					const oShape3 = createShape();
					selectOnlyObjects([oShape1]);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape3, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape2, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape1, true);
				}, oTypes.selectPreviousObject);
			});

			QUnit.test('Test visit hyperlink', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('', 0, false);
					moveToParagraph(oParagraph);
					addHyperlink("https://www.onlyoffice.com/", "Hello");
					moveCursorLeft();
					moveCursorLeft();
					executeTestWithCatchEvent('asc_onHyperlinkClick', (sValue) => sValue, 'https://www.onlyoffice.com/', oEvent, oAssert);
					moveToParagraph(oParagraph);
					moveCursorLeft();
					const oSelectedInfo = new CSelectedElementsInfo();
					graphicController().getTargetDocContent().GetSelectedElementsInfo(oSelectedInfo);
					equal(oSelectedInfo.m_oHyperlink.Visited, true);
				}, oTypes.visitHyperink);
			});

			QUnit.test('Test add line in math', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oParagraph} = createMathInShape();
					moveToParagraph(oParagraph, true);
					moveCursorRight();
					moveCursorRight();
					enterText('fffffff');
					moveCursorLeft();
					moveCursorLeft();
					onKeyDown(oEvent);
					const oParaMath = oParagraph.GetAllParaMaths()[0];
					const oFraction = oParaMath.Root.GetFirstElement();
					const oNumerator = oFraction.getNumerator();
					const oEqArray = oNumerator.GetFirstElement();
					oAssert.strictEqual(oEqArray.getRowsCount(), 2, 'Check add new line math');
				}, oTypes.addLineInMath);
			});
			QUnit.test('Test add break line', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('');
					moveToParagraph(oParagraph);

					onKeyDown(oEvent);
					oAssert.true(oShape.getDocContent().Content.length === 1 && oParagraph.GetLinesCount() === 2, 'Check add break line');
				}, oTypes.addBreakLine);
			});

			QUnit.test('Test add paragraph', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('');
					moveToParagraph(oParagraph);

					onKeyDown(oEvent);
					oAssert.true(oShape.getDocContent().Content.length === 2, 'Check add new paragraph');
				}, oTypes.addParagraph);
			});

			QUnit.test('Test create text body', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oShape = createShape();
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.true(!!oShape.txBody, 'Check creating txBody');
				}, oTypes.createTxBody);
			});

			QUnit.test('Test move cursor to start position in empty content', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('', true);
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.true(oParagraph.IsCursorAtBegin(), 'Check move cursor to start position in shape');
				}, oTypes.moveToStartInEmptyContent);
			});
			QUnit.test('Test select all after enter', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello Word', true);
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.strictEqual(selectedContent(), 'Hello Word', 'Check select all content in shape');
				}, oTypes.selectAllAfterEnter);
			});
			QUnit.test('Test move cursor to start position in empty title', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oChart = createChart();
					selectOnlyObjects([oChart]);
					const oTitles = oChart.getAllTitles();
					const oController = graphicController();
					oController.selection.chartSelection = oChart;
					oChart.selectTitle(oTitles[0], 0);

					onKeyDown(oEvent);
					oAssert.strictEqual(selectedContent(), 'Diagram Title', 'Check select all title');
				}, oTypes.moveCursorToStartPositionInTitle);
			});
			QUnit.test('Test select all after enter in title', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oChart = createChart();
					const oTitles = oChart.getAllTitles();
					const oContent = AscFormat.CreateDocContentFromString('', graphicController().drawingObjects.getDrawingDocument(), oTitles[0].txBody);
					oTitles[0].txBody.content = oContent;
					selectOnlyObjects([oChart]);

					const oController = graphicController();
					oController.selection.chartSelection = oChart;
					oChart.selectTitle(oTitles[0], 0);

					onKeyDown(oEvent);
					oAssert.true(oContent.IsCursorAtBegin(), 'Check move cursor to begin pos in title');
				}, oTypes.selectAllTitleAfterEnter);
			});
			QUnit.test('Test reset text selection', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent("Hello world");
					const oParagraph = moveToShapeParagraph(oShape, 0);
					onKeyDown(oEvent);
					equal(!graphicController().selection.textSelection, true);
				}, oTypes.resetTextSelection);
			});

			QUnit.test('Test reset step selection', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					let oGroupedShape1 = createShape();
					let oGroupedShape2 = createShape();
					createGroup([oGroupedShape1, oGroupedShape2]);
					let oTestGroup = oGroupedShape1.group;
					const oController = graphicController();
					selectOnlyObjects([oTestGroup, oGroupedShape1]);
					onKeyDown(oEvent);
					oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oTestGroup && oTestGroup.selectedObjects.length === 0, 'Check reset step selection');

				}, oTypes.resetStepSelection);
			});
			QUnit.test('Test move cursor to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 18], oEvent, oAssert);
				}, oTypes.moveCursorToEndLine);
			});

			QUnit.test('Test move cursor to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 59], oEvent, oAssert);
				}, oTypes.moveCursorToEndDocument);
			});


			QUnit.test('Test select to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello World Hello ', ''], oEvent, oAssert);
				}, oTypes.selectToEndLine);
			});

			QUnit.test('Test select to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello World Hello World Hello World Hello World Hello World', ''], oEvent, oAssert);
				}, oTypes.selectToEndDocument);
			});

			QUnit.test('Test move cursor to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([54, 0], oEvent, oAssert);
				}, oTypes.moveCursorToStartLine);
			});

			QUnit.test('Test move cursor to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([0, 0], oEvent, oAssert);
				}, oTypes.moveCursorToStartDocument);
			});

			QUnit.test('Test select to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'World'], oEvent, oAssert);
				}, oTypes.selectToStartLine);
			});

			QUnit.test('Test select to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'Hello World Hello World Hello World Hello World Hello World'], oEvent, oAssert);
				}, oTypes.selectToStartDocument);
			});

			QUnit.test('Test move cursor to left char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([58, 0], oEvent, oAssert);

				}, oTypes.moveCursorLeftChar);
			});

			QUnit.test('Test select to left char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'd'], oEvent, oAssert);
				}, oTypes.selectCursorLeftChar);
			});

			QUnit.test('Test move cursor to left word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([54, 0], oEvent, oAssert);
				}, oTypes.moveCursorLeftWord);
			});

			QUnit.test('Test select to left word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'World'], oEvent, oAssert);
				}, oTypes.selectCursorLeftWord);
			});

			QUnit.test('Test move object to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100 - pxToMm, y: 100}, oEvent, oAssert);
				}, oTypes.littleMoveGraphicObjectLeft);
			});

			QUnit.test('Test move object to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100 - 5 * pxToMm, y: 100}, oEvent, oAssert);

				}, oTypes.bigMoveGraphicObjectLeft);
			});

			QUnit.test('Test move cursor to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([41, 0], oEvent, oAssert);

				}, oTypes.moveCursorUp);
			});

			QUnit.test('Test select to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', ' World Hello World'], oEvent, oAssert);
				}, oTypes.selectCursorUp);
			});

			QUnit.test('Test move object to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 - pxToMm}, oEvent, oAssert);
				}, oTypes.littleMoveGraphicObjectUp);
			});

			QUnit.test('Test move object to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 - 5 * pxToMm}, oEvent, oAssert);
				}, oTypes.bigMoveGraphicObjectUp);
			});

			QUnit.test('Test move cursor to right char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 1], oEvent, oAssert);

				}, oTypes.moveCursorRightChar);
			});

			QUnit.test('Test select to right char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['H', ''], oEvent, oAssert);
				}, oTypes.selectCursorRightChar);
			});

			QUnit.test('Test move cursor to right word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 6], oEvent, oAssert);
				}, oTypes.moveCursorRightWord);
			});

			QUnit.test('Test select to right word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello ', ''], oEvent, oAssert);
				}, oTypes.selectCursorRightWord);
			});

			QUnit.test('Test move object to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100 + pxToMm, y: 100}, oEvent, oAssert);

				}, oTypes.littleMoveGraphicObjectRight);
			});

			QUnit.test('Test move object to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100 + 5 * pxToMm, y: 100}, oEvent, oAssert);

				}, oTypes.bigMoveGraphicObjectRight);
			});

			QUnit.test('Test move cursor to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 18], oEvent, oAssert);
				}, oTypes.moveCursorDown);
			});

			QUnit.test('Test select to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello World Hello ', ''], oEvent, oAssert);
				}, oTypes.selectCursorDown);
			});

			QUnit.test('Test move object to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 + pxToMm}, oEvent, oAssert);

				}, oTypes.littleMoveGraphicObjectDown);
			});

			QUnit.test('Test move object to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 + 5 * pxToMm}, oEvent, oAssert);
				}, oTypes.bigMoveGraphicObjectDown);
			});

			function getParagraphText2()
			{
				const oParagraph = graphicController().getTargetDocContent().GetCurrentParagraph();
				return getParagraphText(oParagraph);
			}
			QUnit.test('Test remove front text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oParagraph} = createShapeWithContent('Hello World');
					moveToParagraph(oParagraph, true);
					onKeyDown(oEvent);
					equal(getParagraphText(oParagraph), 'ello World');
				}, oTypes.removeFrontChar);
			});

			QUnit.test('Test remove front text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oParagraph} = createShapeWithContent('Hello World');
					moveToParagraph(oParagraph, true);
					onKeyDown(oEvent);
					equal(getParagraphText(oParagraph), 'World');
				}, oTypes.removeFrontWord);
			});
			QUnit.test('Test select all content in shape', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello World');
					const oParagraph = moveToShapeParagraph(oShape, 0);
					onKeyDown(oEvent);
					equal(selectedContent(), 'Hello World');
				}, oTypes.selectAllContent);
			});
			QUnit.test('Test select all graphic objects', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oShape1 = createShape();
					const oShape2 = createShape();
					const oShape3 = createShape();
					selectOnlyObjects([oShape1]);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 3 && (oShape1.selected && selectedObjects().indexOf(oShape1) !== -1) && (oShape2.selected && selectedObjects().indexOf(oShape2) !== -1) && (oShape3.selected && selectedObjects().indexOf(oShape3) !== -1), true);
				}, oTypes.selectAllDrawings);
			});

			QUnit.test('Test bold', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetBold(), oEvent, oAssert, true);

				}, oTypes.bold);
			});

			QUnit.test('Test clear slicer', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oSlicer = createSlicer();
					oSlicer.data.values[0].asc_setVisible(false);
					onKeyDown(oEvent);
					equal(oSlicer.data.values[0].asc_getVisible(), true);
				}, oTypes.cleanSlicer);
			});

			QUnit.test('Test center align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectParaPrAfterKeyDown((oPara) => oPara.GetJc(), oEvent, oAssert, align_Center);

				}, oTypes.centerAlign);
			});
			QUnit.test('Test italic', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetItalic(), oEvent, oAssert, true);

				}, oTypes.italic);
			});
			QUnit.test('Test justify align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectParaPrAfterKeyDown((oPara) => oPara.GetJc(), oEvent, oAssert, align_Justify);

				}, oTypes.justifyAlign);
			});
			QUnit.test('Test left align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const continueCheck = checkDirectParaPrAfterKeyDown((oPara) => oPara.GetJc(), oEvent, oAssert, align_Justify);
					continueCheck((oPara) => oPara.GetJc(), oEvent, oAssert, align_Left);

				}, oTypes.leftAlign);
			});
			QUnit.test('Test right align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oEvent, oAssert, align_Right);
				}, oTypes.rightAlign);
			});

			QUnit.test('Test invert multiselect slicer', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const oSlicer = createSlicer();
					onKeyDown(oEvent);
					equal(oSlicer.isMultiSelect(), true);
				}, oTypes.invertMultiselectSlicer);
			});
			QUnit.test('Test underline', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetUnderline(), oEvent, oAssert, true);
				}, oTypes.underline);
			});

			QUnit.test('Test superscript vertical align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oEvent, oAssert, AscCommon.vertalign_SuperScript);

				}, oTypes.superscript);
			});

			QUnit.test('Test add en dash', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{

				}, oTypes.enDash);
			});

			QUnit.test('Test add hyphen', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{

				}, oTypes.hyphen);
			});

			QUnit.test('Test add underscore', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{

				}, oTypes.underscore);
			});

			QUnit.test('Test add subscript', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oEvent, oAssert, AscCommon.vertalign_SubScript, 'Check subscript shortcut');
				}, oTypes.subscript);
			});

			QUnit.test('Test decrease font size', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello');
					selectAll();
					addProperty({FontSize: 16});
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 14);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 12);

					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 11);

					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 10);
				}, oTypes.decreaseFontSize);
			});

			QUnit.test('Test increase font size', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello');
					moveToShapeParagraph(oShape, 0);
					selectAll();
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 11);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 12);

					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 14);

					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 16);
				}, oTypes.increaseFontSize);
			});
		}
	)

	const oTypes = {
		removeBackChar                  : 1,
		removeBackWord                  : 1.5,
		removeChart                     : 2,
		removeShape                     : 3,
		removeGroup                     : 4,
		removeShapeInGroup              : 5,
		addTab                          : 6,
		selectNextObject                : 7,
		selectPreviousObject            : 8,
		visitHyperink                   : 9,
		addLineInMath                   : 10,
		addBreakLine                    : 11,
		addParagraph                    : 12,
		createTxBody                    : 13,
		moveToStartInEmptyContent       : 14,
		selectAllAfterEnter             : 15,
		moveCursorToStartPositionInTitle: 16,
		selectAllTitleAfterEnter        : 17,
		resetTextSelection              : 18,
		resetStepSelection              : 18.5,
		moveCursorToEndDocument         : 19,
		moveCursorToEndLine             : 19.5,
		selectToEndDocument             : 20,
		selectToEndLine                 : 20.5,
		moveCursorToStartDocument       : 21,
		moveCursorToStartLine           : 21.5,
		selectToStartDocument           : 22,
		selectToStartLine               : 22.5,
		moveCursorLeftChar              : 23,
		selectCursorLeftChar            : 24,
		moveCursorLeftWord              : 25,
		selectCursorLeftWord            : 26,
		bigMoveGraphicObjectLeft        : 27,
		littleMoveGraphicObjectLeft     : 28,
		moveCursorRightChar             : 29,
		selectCursorRightChar           : 30,
		moveCursorRightWord             : 31,
		selectCursorRightWord           : 32,
		bigMoveGraphicObjectRight       : 33,
		littleMoveGraphicObjectRight    : 34,
		moveCursorUp                    : 35,
		selectCursorUp                  : 36,
		bigMoveGraphicObjectUp          : 37,
		littleMoveGraphicObjectUp       : 38,
		moveCursorDown                  : 39,
		selectCursorDown                : 40,
		bigMoveGraphicObjectDown        : 41,
		littleMoveGraphicObjectDown     : 42,
		removeFrontWord                 : 43,
		removeFrontChar                 : 44,
		selectAllContent                : 45,
		selectAllDrawings               : 46,
		bold                            : 47,
		cleanSlicer                     : 48,
		centerAlign                     : 49,
		italic                          : 50,
		justifyAlign                    : 51,
		leftAlign                       : 52,
		rightAlign                      : 53,
		invertMultiselectSlicer         : 54,
		underline                       : 55,
		superscriptAndSubscript         : 56,
		superscript                     : 57,
		enDash                          : 58,
		hyphen                          : 59,
		underscore                      : 60,
		subscript                       : 61,
		increaseFontSize                : 62,
		decreaseFontSize                : 63
	};


	const testAll = 0;
	const testWindows = 1;
	const testMacOs = 2;

	function CTestEvent(oEvent, nType)
	{
		this.type = nType || testAll;
		this.event = oEvent;
	}

	const oKeyCode =
		{
			BackSpace       : 8,
			Tab             : 9,
			Enter           : 13,
			Esc             : 27,
			End             : 35,
			Home            : 36,
			ArrowLeft       : 37,
			ArrowTop        : 38,
			ArrowRight      : 39,
			ArrowBottom     : 40,
			Delete          : 46,
			A               : 65,
			B               : 66,
			C               : 67,
			E               : 69,
			I               : 73,
			J               : 74,
			K               : 75,
			L               : 76,
			M               : 77,
			P               : 80,
			R               : 82,
			S               : 83,
			U               : 85,
			V               : 86,
			X               : 88,
			Y               : 89,
			Z               : 90,
			OperaContextMenu: 57351,
			F10             : 121,
			NumLock         : 144,
			ScrollLock      : 145,
			Equal           : 187,
			Comma           : 188,
			Minus           : 189,
			Period          : 190,
			BracketLeft     : 219,
			BracketRight    : 221,
			F2              : 113
		}

	const oTestEvents = {};
	oTestEvents[oTypes.removeBackChar] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, true, false, false)),
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, true, false, false))
	];
	oTestEvents[oTypes.removeBackWord] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, true, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.BackSpace, true, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.BackSpace, true, false, true, false, false))
	];
	oTestEvents[oTypes.removeChart] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oTestEvents[oTypes.removeShape] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oTestEvents[oTypes.removeGroup] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oTestEvents[oTypes.removeShapeInGroup] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oTestEvents[oTypes.addTab] = [
		new CTestEvent(createEvent(oKeyCode.Tab, false, false, false, false, false))
	];
	oTestEvents[oTypes.selectNextObject] = [
		new CTestEvent(createEvent(oKeyCode.Tab, false, false, false, false, false))
	];
	oTestEvents[oTypes.selectPreviousObject] = [
		new CTestEvent(createEvent(oKeyCode.Tab, false, true, false, false, false))
	];
	oTestEvents[oTypes.visitHyperink] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oTestEvents[oTypes.addLineInMath] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oTestEvents[oTypes.addBreakLine] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, true, false, false, false))
	];
	oTestEvents[oTypes.addParagraph] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oTestEvents[oTypes.createTxBody] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oTestEvents[oTypes.moveToStartInEmptyContent] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oTestEvents[oTypes.selectAllAfterEnter] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oTestEvents[oTypes.moveCursorToStartPositionInTitle] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectAllTitleAfterEnter] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))

	];
	oTestEvents[oTypes.resetTextSelection] = [
		new CTestEvent(createEvent(oKeyCode.Esc, false, false, false, false, false))

	];
	oTestEvents[oTypes.resetStepSelection] = [
		new CTestEvent(createEvent(oKeyCode.Esc, false, false, false, false, false))
	];
	oTestEvents[oTypes.moveCursorToEndDocument] = [
		new CTestEvent(createEvent(oKeyCode.End, true, false, false, false, false))

	];
	oTestEvents[oTypes.moveCursorToEndLine] = [
		new CTestEvent(createEvent(oKeyCode.End, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectToEndDocument] = [
		new CTestEvent(createEvent(oKeyCode.End, true, true, false, false, false))

	];
	oTestEvents[oTypes.selectToEndLine] = [
		new CTestEvent(createEvent(oKeyCode.End, false, true, false, false, false))

	];
	oTestEvents[oTypes.moveCursorToStartDocument] = [
		new CTestEvent(createEvent(oKeyCode.Home, true, false, false, false, false))

	];
	oTestEvents[oTypes.moveCursorToStartLine] = [
		new CTestEvent(createEvent(oKeyCode.Home, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectToStartDocument] = [
		new CTestEvent(createEvent(oKeyCode.Home, true, true, false, false, false))

	];
	oTestEvents[oTypes.selectToStartLine] = [
		new CTestEvent(createEvent(oKeyCode.Home, false, true, false, false, false))

	];
	oTestEvents[oTypes.moveCursorLeftChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, false, false, false, false))
	];
	oTestEvents[oTypes.selectCursorLeftChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, true, false, false, false))

	];
	oTestEvents[oTypes.moveCursorLeftWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, false, false, false, false))

	];
	oTestEvents[oTypes.selectCursorLeftWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, true, false, false, false))

	];
	oTestEvents[oTypes.bigMoveGraphicObjectLeft] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, false, false, false, false))

	];
	oTestEvents[oTypes.littleMoveGraphicObjectLeft] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, false, false, false, false))

	];
	oTestEvents[oTypes.moveCursorRightChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectCursorRightChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, true, false, false, false))

	];
	oTestEvents[oTypes.moveCursorRightWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, false, false, false, false))

	];
	oTestEvents[oTypes.selectCursorRightWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, true, false, false, false))

	];
	oTestEvents[oTypes.bigMoveGraphicObjectRight] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, false, false, false, false))

	];
	oTestEvents[oTypes.littleMoveGraphicObjectRight] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, false, false, false, false))

	];
	oTestEvents[oTypes.moveCursorUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectCursorUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, false, true, false, false, false))

	];
	oTestEvents[oTypes.bigMoveGraphicObjectUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, false, false, false, false, false))
	];
	oTestEvents[oTypes.littleMoveGraphicObjectUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, true, false, false, false, false))

	];
	oTestEvents[oTypes.moveCursorDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectCursorDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, false, true, false, false, false))

	];
	oTestEvents[oTypes.bigMoveGraphicObjectDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, false, false, false, false, false))

	];
	oTestEvents[oTypes.littleMoveGraphicObjectDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, true, false, false, false, false))

	];
	oTestEvents[oTypes.removeFrontWord] = [
		new CTestEvent(createEvent(oKeyCode.Delete, true, false, false, false, false))

	];
	oTestEvents[oTypes.removeFrontChar] = [
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))

	];
	oTestEvents[oTypes.selectAllContent] = [
		new CTestEvent(createEvent(oKeyCode.A, true, false, false, false, false))

	];
	oTestEvents[oTypes.selectAllDrawings] = [
		new CTestEvent(createEvent(oKeyCode.A, true, false, false, false, false))

	];
	oTestEvents[oTypes.bold] = [
		new CTestEvent(createEvent(oKeyCode.B, true, false, false, false, false))

	];
	oTestEvents[oTypes.cleanSlicer] = [
		new CTestEvent(createEvent(oKeyCode.C, true, false, true, false, false), testMacOs),
		new CTestEvent(createEvent(oKeyCode.C, false, false, true, false, false), testWindows)
	];
	oTestEvents[oTypes.centerAlign] = [
		new CTestEvent(createEvent(oKeyCode.E, true, false, false, false, false))

	];
	oTestEvents[oTypes.italic] = [
		new CTestEvent(createEvent(oKeyCode.I, true, false, false, false, false))

	];
	oTestEvents[oTypes.justifyAlign] = [
		new CTestEvent(createEvent(oKeyCode.J, true, false, false, false, false))

	];
	oTestEvents[oTypes.leftAlign] = [
		new CTestEvent(createEvent(oKeyCode.L, true, false, false, false, false))

	];
	oTestEvents[oTypes.rightAlign] = [
		new CTestEvent(createEvent(oKeyCode.R, true, false, false, false, false))
	];
	oTestEvents[oTypes.invertMultiselectSlicer] = [
		new CTestEvent(createEvent(oKeyCode.S, true, false, true, false, false), testMacOs),
		new CTestEvent(createEvent(oKeyCode.S, false, false, true, false, false), testWindows)
	];
	oTestEvents[oTypes.underline] = [
		new CTestEvent(createEvent(oKeyCode.U, true, false, false, false, false))

	];
	oTestEvents[oTypes.superscript] = [
		new CTestEvent(createEvent(oKeyCode.Equal, true, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Comma, true, false, false, false, false))

	];
	oTestEvents[oTypes.enDash] = [
		new CTestEvent(createEvent(oKeyCode.Minus, true, true, false, false, false))

	];
	oTestEvents[oTypes.hyphen] = [
		new CTestEvent(createEvent(oKeyCode.Minus, true, false, false, false, false))

	];
	oTestEvents[oTypes.underscore] = [
		new CTestEvent(createEvent(oKeyCode.Minus, false, true, false, false, false))

	];
	oTestEvents[oTypes.subscript] = [
		new CTestEvent(createEvent(oKeyCode.Equal, true, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Period, true, false, false, false, false))

	];
	oTestEvents[oTypes.increaseFontSize] = [
		new CTestEvent(createEvent(oKeyCode.BracketRight, true, false, false, false, false))

	];
	oTestEvents[oTypes.decreaseFontSize] = [
		new CTestEvent(createEvent(oKeyCode.BracketLeft, true, false, false, false, false))

	];

	function startTest(fCallback, nShortcutType)
	{
		cleanGraphic();
		const arrTestEvents = oTestEvents[nShortcutType];

		for (let i = 0; i < arrTestEvents.length; i += 1)
		{
			cleanGraphic();
			const nTestType = arrTestEvents[i].type;
			if (nTestType === testAll)
			{
				AscCommon.AscBrowser.isMacOs = true;
				fCallback(arrTestEvents[i].event);

				cleanGraphic();
				AscCommon.AscBrowser.isMacOs = false;
				fCallback(arrTestEvents[i].event);
			} else if (nTestType === testMacOs)
			{
				AscCommon.AscBrowser.isMacOs = true;
				fCallback(arrTestEvents[i].event);
				AscCommon.AscBrowser.isMacOs = false;
			} else if (nTestType === testWindows)
			{
				fCallback(arrTestEvents[i].event);
			}
		}
	}
})(window);
