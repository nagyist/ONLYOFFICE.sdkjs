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

'use strict';
(function (window)
{
	const {
		createMathInShape,
		moveCursorDown,
		moveCursorLeft,
		moveCursorRight,
		onInput,
		goToPageWithFocus,
		goToPage,
		checkDirectTextPrAfterKeyDown,
		checkDirectParaPrAfterKeyDown,
		createShape,
		addSlide,
		onKeyDown,
		executeTestWithCatchEvent,
		createNativeEvent,
		checkSelectedSlides,
		cleanPresentation,
		executeCheckMoveShape,
		addPropertyToDocument,
		testMoveHelper,
		createShapeWithTitlePlaceholder,
		createChart,
		createTable,
		checkRemoveObject,
		checkTextAfterKeyDownHelperHelloWorld,
		checkTextAfterKeyDownHelperEmpty,
		getDirectParaPrHelper,
		getDirectTextPrHelper,
		createEvent,
		getShapeWithParagraphHelper,
		moveToParagraph,
		getFirstSlide,
		selectOnlyObjects,
		remove,
		addToSelection,
		createGroup,
		getController,
		oGlobalLogicDocument,
		oGlobalShape
	} = window.AscTestShortcut;

	let oUndoShape;
	let oMockEvent = createNativeEvent();

	function checkEditUndo(oAssert)
	{
		oUndoShape = createShape();
		onKeyDown(oMockEvent);
		oAssert.strictEqual(getFirstSlide().cSld.spTree.length, 0, 'Check undo shortcut');
	}

	function checkEditRedo(oAssert)
	{
		onKeyDown(oMockEvent);
		oAssert.strictEqual(getFirstSlide().cSld.spTree.length === 1 && getFirstSlide().cSld.spTree[0] === oUndoShape, true, 'Check redo shortcut');
	}

	function checkEditSelectAll(oAssert)
	{
		onKeyDown(oMockEvent);
		const oController = getController();
		oAssert.strictEqual(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oUndoShape, true, 'check select all shortcut');
	}

	function checkDuplicate(oAssert)
	{
		const {oShape} = getShapeWithParagraphHelper('', true);
		const arrOldSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree.slice();
		selectOnlyObjects([oShape]);
		onKeyDown(oMockEvent);
		const arrUpdatedSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree;
		const oNewShape = arrUpdatedSpTree[arrUpdatedSpTree.length - 1];
		oAssert.true(arrOldSpTree.indexOf(oNewShape) === -1, 'Check duplicate shape');
	}

	function checkPrint(oAssert)
	{
		executeTestWithCatchEvent('asc_onPrint', () => true, true, oMockEvent, oAssert);
	}

	function checkSave(oAssert)
	{
		const fOldSave = editor._onSaveCallbackInner;
		let bCheck = false;
		editor._onSaveCallbackInner = function ()
		{
			bCheck = true;
			editor.canSave = true;
		};
		onKeyDown(oMockEvent);
		oAssert.strictEqual(bCheck, true, 'Check save shortcut');
		editor._onSaveCallbackInner = fOldSave;
	}

	function checkShowContextMenu(oAssert)
	{
		executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oMockEvent, oAssert, () =>
		{
			const {oParagraph} = getShapeWithParagraphHelper('');
			oParagraph.SetThisElementCurrent();
		});
	}

	function checkShowParaMarks(oAssert)
	{
		editor.put_ShowParaMarks(false);
		onKeyDown(oMockEvent);
		oAssert.true(!!editor.get_ShowParaMarks(), 'Check show para marks shortcut');
	}

	function checkBold(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Bold(), oMockEvent, oAssert, true, 'Check bold shortcut');
	}

	function checkCenterAlign(oAssert)
	{
		checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Center, 'Check center align shortcut');
	}

	function checkEuroSign(oAssert)
	{
		checkTextAfterKeyDownHelperEmpty('€', oMockEvent, oAssert, 'Check euro sign shortcut');
	}

	let oGroup;
	let oFirstShape;
	let oSecondShape;

	function checkGroup(oAssert)
	{
		oFirstShape = createShape();
		oSecondShape = createShape();
		selectOnlyObjects([oFirstShape, oSecondShape]);

		onKeyDown(oMockEvent);
		oGroup = oFirstShape.group;
		oAssert.true(oFirstShape.group && (oFirstShape.group === oSecondShape.group), 'Check group shortcut');
	}

	function checkUnGroup(oAssert)
	{
		selectOnlyObjects([oGroup]);

		onKeyDown(oMockEvent);
		oAssert.true(!oFirstShape.group && !oSecondShape.group && oGroup.bDeleted, 'Check ungroup shortcut');

	}

	function checkItalic(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Italic(), oMockEvent, oAssert, true, 'Check italic shortcut');
	}

	function checkJustifyAlign(oAssert)
	{
		checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Justify, 'check justify align shortcut');
	}

	function checkAddHyperlink(oAssert)
	{
		executeTestWithCatchEvent('asc_onDialogAddHyperlink', () => true, true, oMockEvent, oAssert, () =>
		{
			const {oParagraph} = getShapeWithParagraphHelper('Hello World');
			moveToParagraph(oParagraph);
			oGlobalLogicDocument.SelectAll();
		});
	}

	function checkBulletList(oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();

		onKeyDown(oMockEvent);
		const oBullet = oParagraph.Get_PresentationNumbering();
		oAssert.true(oBullet.m_nType === AscFormat.numbering_presentationnumfrmt_Char, 'Check bullet list shortcut');
	}

	function checkLeftAlign(oAssert)
	{
		checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Left, 'check right align shortcut');
	}

	function checkRightAlign(oAssert)
	{
		checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Right, 'check right align shortcut');

	}

	function checkUnderline(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Underline(), oMockEvent, oAssert, true, 'Check underline shortcut');
	}

	function checkStrikethrough(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Strikeout(), oMockEvent, oAssert, true, 'Check strikeout shortcut');
	}

	let oCopyParagraphTextPr;

	function checkCopyFormat(oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();
		addPropertyToDocument({Bold: true, Italic: true, Underline: true});

		onKeyDown(oMockEvent);
		oCopyParagraphTextPr = new AscCommonWord.CTextPr();
		oCopyParagraphTextPr.SetUnderline(true);
		oCopyParagraphTextPr.SetBold(true);
		oCopyParagraphTextPr.BoldCS = true;
		oCopyParagraphTextPr.SetItalic(true);
		oCopyParagraphTextPr.ItalicCS = true;
		oAssert.deepEqual(editor.getFormatPainterData().TextPr, oCopyParagraphTextPr, 'Check copy format shortcut');
	}

	function checkPasteFormat(oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();

		onKeyDown(oMockEvent);
		const oDirectTextPr = oParagraph.GetDirectTextPr();
		oAssert.deepEqual(oDirectTextPr, oCopyParagraphTextPr, 'check paste format shortcut');
	}

	function checkSuperscript(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oMockEvent, oAssert, AscCommon.vertalign_SuperScript, 'Check superscript shortcut');
	}

	function checkSubscript(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oMockEvent, oAssert, AscCommon.vertalign_SubScript, 'Check subscript shortcut');
	}

	function checkEnDash(oAssert)
	{
		checkTextAfterKeyDownHelperEmpty('–', oMockEvent, oAssert, 'Check en dash shortcut');
	}

	function checkDecreaseFont(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), oMockEvent, oAssert, 9, 'Check decrease font size shortcut');
	}

	function checkIncreaseFont(oAssert)
	{
		checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), oMockEvent, oAssert, 11, 'Check increase font size shortcut');
	}

	function checkDeleteBack(oEvent, oAssert)
	{
		checkTextAfterKeyDownHelperHelloWorld('Hello Worl', oEvent, oAssert, 'Check delete with backspace')
	}

	function checkDeleteWordBack(oEvent, oAssert)
	{
		checkTextAfterKeyDownHelperHelloWorld('Hello ', oEvent, oAssert, 'Check delete word with backspace')
	}

	function checkRemoveAnimation(oEvent, oAssert)
	{
		const {oShape} = getShapeWithParagraphHelper('', true);
		selectOnlyObjects([oShape]);
		oGlobalLogicDocument.AddAnimation(1, 1, 0, false, false);

		onKeyDown(oEvent);
		const oTiming = oGlobalLogicDocument.GetCurTiming();
		const arrEffects = oTiming.getObjectEffects(oShape.GetId());
		oAssert.true(arrEffects.length === 0, 'Check remove animation');
	}

	function checkRemoveChart(oEvent, oAssert)
	{
		const oChart = createChart(getFirstSlide());
		selectOnlyObjects([oChart]);
		onKeyDown(oEvent);
		oAssert.true(checkRemoveObject(oChart, getFirstSlide().cSld.spTree), "Check remove group");
	}

	function checkRemoveShape(oEvent, oAssert)
	{
		const {oShape} = getShapeWithParagraphHelper('', true);
		selectOnlyObjects([oShape]);

		onKeyDown(oEvent);
		const arrSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree;
		oAssert.true(checkRemoveObject(oShape, arrSpTree), 'Check remove shape');
	}

	function checkRemoveTable(oEvent, oAssert)
	{
		const oGraphicFrame = createTable(3, 3);
		selectOnlyObjects([oGraphicFrame]);
		onKeyDown(oEvent);
		const arrSpTree = getFirstSlide().cSld.spTree;
		oAssert.true(checkRemoveObject(oGraphicFrame, arrSpTree), "Check remove table");
	}

	function checkRemoveGroup(oEvent, oAssert)
	{
		const oGroup = createGroup([createShape(), createShape()]);
		selectOnlyObjects([oGroup]);
		onKeyDown(oEvent);
		const arrSpTree = getFirstSlide().cSld.spTree;
		oAssert.true(checkRemoveObject(oGroup, arrSpTree), 'Check remove group');
	}

	function checkRemoveShapeInGroup(oEvent, oAssert)
	{
		const oGroupedGroup = createGroup([createShape(), createShape()]);
		const oRemovedShape = createShape();
		const oGroup = createGroup([oGroupedGroup, oRemovedShape]);
		selectOnlyObjects([oRemovedShape]);
		onKeyDown(oEvent);
		oAssert.true(checkRemoveObject(oRemovedShape, oGroup.spTree), 'Check remove shape in group');
	}

	let oTable

	function checkMoveToNextCell(oEvent, oAssert)
	{
		const oGraphicFrame = createTable(3, 3);
		oTable = oGraphicFrame.graphicObject;
		onKeyDown(oEvent);
		oAssert.strictEqual(oTable.CurCell.Index, 1, 'check go to next cell shortcut');
	}

	function checkMoveToPreviousCell(oEvent, oAssert)
	{
		onKeyDown(oEvent);
		oAssert.strictEqual(oTable.CurCell.Index, 0, 'check go to previous cell shortcut');
	}

	let oBulletParagraph;

	function checkIncreaseBulletIndent(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello');
		const oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
		oParagraph.Add_PresentationNumbering(oBullet);
		moveToParagraph(oParagraph, true);
		oParagraph.Set_Ind({Left: 0});
		onKeyDown(oEvent);
		oBulletParagraph = oParagraph;
		oAssert.strictEqual(oParagraph.Pr.Get_IndLeft(), 11.1125, 'Check bullet indent shortcut');
	}

	function checkDecreaseBulletIndent(oEvent, oAssert)
	{
		moveToParagraph(oBulletParagraph, true);
		onKeyDown(oEvent);
		oAssert.strictEqual(oBulletParagraph.Pr.Get_IndLeft(), 0, 'Check bullet indent shortcut');
	}

	function checkAddTab(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('');
		moveToParagraph(oParagraph);
		onKeyDown(oEvent);
		let bCheck = false;
		for (let i = oParagraph.Content.length - 2; i >= 0; --i)
		{
			const oRun = oParagraph.Content[i];
			if (oRun.Content.length && !oRun.IsParaEndRun())
			{
				bCheck = oRun.Content[oRun.Content.length - 1].Type === para_Tab;
				break;
			}
		}
		oAssert.true(bCheck, 'Check add tab');
	}

	function checkSelectNextObject(oEvent, oAssert)
	{
		const {oShape, oController} = getShapeWithParagraphHelper('', true);
		selectOnlyObjects([oShape]);
		onKeyDown(oEvent);
		const arrSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree;
		let oSelectedShape;
		for (let i = 0; i < arrSpTree.length; i += 1)
		{
			if (arrSpTree[i] === oShape)
			{
				oSelectedShape = arrSpTree[i < arrSpTree.length - 1 ? i + 1 : 0];
			}
		}
		oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oSelectedShape && oController.selectedObjects[0] !== oShape, 'Check select next object');

	}

	function checkSelectPreviousObject(oEvent, oAssert)
	{
		const {oShape, oController} = getShapeWithParagraphHelper('', true);

		selectOnlyObjects([oShape]);
		onKeyDown(oEvent);
		const arrSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree;
		let oSelectedShape;
		for (let i = 0; i < arrSpTree.length; i += 1)
		{
			if (arrSpTree[i] === oShape)
			{
				oSelectedShape = arrSpTree[i > 0 ? i - 1 : arrSpTree.length - 1];
			}
		}
		oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oSelectedShape && oController.selectedObjects[0] !== oShape, 'Check select previous object');

	}

	function checkVisitHyperlink(oEvent, oAssert)
	{
		goToPage(1);
		const {oParagraph} = getShapeWithParagraphHelper('Hello');
		moveToParagraph(oParagraph);
		oGlobalLogicDocument.AddHyperlink({
			Text   : 'abcd',
			ToolTip: 'abcd',
			Value  : 'ppaction://hlinkshowjump?jump=firstslide'
		});
		moveCursorLeft();
		moveCursorLeft();
		onKeyDown(oEvent);
		const oSelectedInfo = oGlobalLogicDocument.IsCursorInHyperlink();
		oAssert.true(oSelectedInfo.Visited && oGlobalLogicDocument.GetSelectedSlides()[0] === 0, 'Check visit hyperlink');
		goToPage(0);
	}

	function checkSelectNextObjectWithPlaceholder(oEvent, oAssert)
	{
		const oFirstShapeWithPlaceholder = createShapeWithTitlePlaceholder();
		const oSecondShapeWithPlaceholder = createShapeWithTitlePlaceholder();

		const oController = getController();
		oController.resetSelection();
		onKeyDown(oEvent);
		oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oFirstShapeWithPlaceholder && oFirstShapeWithPlaceholder.selected, 'Check select first shape with placeholder');

		onKeyDown(oEvent);
		oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oSecondShapeWithPlaceholder && oSecondShapeWithPlaceholder.selected, 'Check select second shape with placeholder');

	}

	function checkAddNextSlideAfterSelectLastPlaceholderObject(oEvent, oAssert)
	{
		const arrOldSlides = oGlobalLogicDocument.Slides.slice();
		onKeyDown(oEvent);
		const arrSelectedSlides = oGlobalLogicDocument.GetSelectedSlides();
		oAssert.true(arrSelectedSlides.length === 1 && arrSelectedSlides[0] === 1 && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check add next slide after selecting last placeholder on current slide');
		goToPage(0);
	}

	function checkAddBreakLine(oEvent, oAssert)
	{
		const {oShape, oParagraph} = getShapeWithParagraphHelper('');
		moveToParagraph(oParagraph);

		onKeyDown(oEvent);
		oAssert.true(oShape.getDocContent().Content.length === 1 && oParagraph.GetLinesCount() === 2, 'Check add break line');
	}

	function checkAddTitleBreakLine(oEvent, oAssert)
	{
		const oShapeWithPlaceholder = createShapeWithTitlePlaceholder();
		const oContent = oShapeWithPlaceholder.getDocContent();
		const oParagraph = oContent.GetAllParagraphs()[0];
		oParagraph.SetThisElementCurrent();
		onKeyDown(oEvent);
		oAssert.true(oContent.Content.length === 1 && oParagraph.GetLinesCount() === 2, 'Check add break line in title');
	}

	function checkAddMathBreakLine(oEvent, oAssert)
	{
		const {oParagraph} = createMathInShape();
		oGlobalLogicDocument.MoveCursorToStartPos();
		moveCursorRight();
		moveCursorRight();
		onInput([56, 56, 56, 56, 56, 56, 56]);
		moveCursorLeft();
		moveCursorLeft();
		onKeyDown(oEvent);
		const oParaMath = oParagraph.GetAllParaMaths()[0];
		const oFraction = oParaMath.Root.GetFirstElement();
		const oNumerator = oFraction.getNumerator();
		const oEqArray = oNumerator.GetFirstElement();
		oAssert.strictEqual(oEqArray.getRowsCount(), 2, 'Check add new line math');
	}

	function checkAddParagraph(oEvent, oAssert)
	{
		const {oShape, oParagraph} = getShapeWithParagraphHelper('');
		moveToParagraph(oParagraph);

		onKeyDown(oEvent);
		oAssert.true(oShape.getDocContent().Content.length === 2, 'Check add new paragraph');
	}

	function checkAddTxBodyShape(oEvent, oAssert)
	{
		const oShape = createShape();
		selectOnlyObjects([oShape]);
		onKeyDown(oEvent);
		oAssert.true(!!oShape.txBody, 'Check creating txBody');
	}

	function checkMoveCursorToStartPosShape(oEvent, oAssert)
	{
		const {oShape, oParagraph} = getShapeWithParagraphHelper('', true);
		selectOnlyObjects([oShape]);
		onKeyDown(oEvent);
		oAssert.true(oParagraph.IsCursorAtBegin(), 'Check move cursor to start position in shape');
	}

	function checkSelectAllContentShape(oEvent, oAssert)
	{
		const {oShape} = getShapeWithParagraphHelper('Hello Word', true);
		selectOnlyObjects([oShape]);
		onKeyDown(oEvent);
		oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(), 'Hello Word', 'Check select all content in shape');

	}

	let oChart;

	function checkSelectAllContentChartTitle(oEvent, oAssert)
	{
		oChart = createChart();
		selectOnlyObjects([oChart]);
		const oTitles = oChart.getAllTitles();
		const oController = getController();
		oController.selection.chartSelection = oChart;
		oChart.selectTitle(oTitles[0], 0);

		onKeyDown(oEvent);
		oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(), 'Diagram Title', 'Check select all title');
	}

	function checkMoveCursorToStartPosChartTitle(oEvent, oAssert)
	{
		const oTitles = oChart.getAllTitles();
		const oContent = AscFormat.CreateDocContentFromString('', editor.WordControl.m_oDrawingDocument, oTitles[0].txBody);
		oTitles[0].txBody.content = oContent;
		selectOnlyObjects([oChart]);

		const oController = getController();
		oController.selection.chartSelection = oChart;
		oChart.selectTitle(oTitles[0], 0);

		onKeyDown(oEvent);
		oAssert.true(oContent.IsCursorAtBegin(), 'Check move cursor to begin pos in title');
	}


	function checkRemoveAndMoveToStartPosTable(oEvent, oAssert)
	{
		const arrSteps = [];
		const oFrame = createTable(3, 3);
		oFrame.Set_CurrentElement();
		const oTable = oFrame.graphicObject;
		oTable.MoveCursorToStartPos();
		// First cell
		moveCursorRight(true, true);
		moveCursorRight(true, true);
		// Second cell
		moveCursorRight(true, true);
		// Third cell
		moveCursorRight(true, true);

		onKeyDown(oEvent);
		arrSteps.push(oTable.IsCursorAtBegin());
		moveCursorRight(true, true);
		moveCursorRight(true, true);
		moveCursorRight(true, true);
		arrSteps.push(oGlobalLogicDocument.GetSelectedText());
		oAssert.deepEqual(arrSteps, [true, ''], 'Check remove and move to start position in table');
	}

	function checkSelectFirstCellContent(oEvent, oAssert)
	{
		const oFrame = createTable(3, 3);
		selectOnlyObjects([oFrame]);

		onKeyDown(oEvent);
		oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(), 'Cell0x0', 'Check select first cell content');
	}

	function checkResetAddShape(oEvent, oAssert)
	{
		const oController = getController();
		oController.changeCurrentState(new AscFormat.StartAddNewShape(oController, 'rect'));

		onKeyDown(oEvent);
		oAssert.true(oController.curState instanceof AscFormat.NullState, 'Check reset add new shape');
	}

	let oGroupedShape1;
	let oGroupedShape2;
	let oTestGroup;

	function checkResetAllDrawingSelection(oEvent, oAssert)
	{
		const oController = getController();
		oController.resetSelection();
		oGroupedShape1 = createShape();
		oGroupedShape2 = createShape();
		createGroup([oGroupedShape1, oGroupedShape2]);
		addToSelection(oGroupedShape1);
		oTestGroup = oGroupedShape1.group;
		onKeyDown(oEvent);
		oAssert.true(oController.selectedObjects.length === 0, 'Check reset all selection');

	}

	function checkResetStepDrawingSelection(oEvent, oAssert)
	{
		const oController = getController();

		selectOnlyObjects([oTestGroup, oGroupedShape1]);
		onKeyDown(oEvent);
		oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oTestGroup && oTestGroup.selectedObjects.length === 0, 'Check reset step selection');
	}

	function checkNonBreakingSpace(oEvent, oAssert)
	{
		checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00A0), oEvent, oAssert, 'Check add non breaking space');
	}

	function checkClearParagraphFormatting(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();
		addPropertyToDocument({Bold: true, Italic: true, Underline: true});

		onKeyDown(oEvent);
		const oTextPr = oGlobalLogicDocument.GetDirectTextPr();
		oAssert.true(!(oTextPr.GetBold() || oTextPr.GetItalic() || oTextPr.GetUnderline()), 'Check clear paragraph formatting');
	}

	function checkAddSpace(oEvent, oAssert)
	{
		checkTextAfterKeyDownHelperEmpty(' ', oEvent, oAssert, 'Check add space')
	}

	function checkMoveToEndPosContent(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, false, true, false);
		oAssert.true(oPos.X === 25 && oPos.Y === 75, 'Check move cursor to end position shortcut');
	}

	function checkMoveToEndLineContent(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, false, true, false);
		oAssert.true(oPos.X === 100 && oPos.Y === 15, 'Check move cursor to end line shortcut');
	}

	function checkSelectToEndLineContent(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
		oAssert.strictEqual(sSelectedText, 'HelloworldHelloworld', 'Check select text to end line shortcut');
	}

	function checkMoveToStartPosContent(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, true, true, false);
		oAssert.true(oPos.X === 0 && oPos.Y === 15, 'Check move to start position shortcut');
	}

	function checkMoveToStartLineContent(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, true, true, false);
		oAssert.true(oPos.X === 0 && oPos.Y === 75, 'Check move to start line shortcut');
	}

	function checkSelectToStartLineContent(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
		oAssert.strictEqual(sSelectedText, 'Hello', 'Check select to start line shortcut');
	}

	function checkMoveCursorLeft(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, true, true, false);
		oAssert.true(oPos.X === 20 && oPos.Y === 75, 'Check move cursor to end position shortcut');
	}

	function checkSelectCursorLeft(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
		oAssert.strictEqual(sSelectedText, 'o', 'Check select text to left position shortcut');
	}

	function checkSelectWordCursorLeft(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
		oAssert.strictEqual(sSelectedText, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', 'Check select word text to left position shortcut');
	}

	function checkMoveCursorWordLeft(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, true, true, false);
		oAssert.true(oPos.X === 0 && oPos.Y === 15, 'Check move cursor to left word position shortcut');
	}

	function checkMoveCursorLeftTable(oEvent, oAssert)
	{
		const oFrame = createTable(3, 3);
		oFrame.Set_CurrentElement();
		const oTable = oFrame.graphicObject;
		oTable.MoveCursorToStartPos();
		moveCursorRight(true);
		moveCursorRight(true);
		onKeyDown(oEvent);
		oAssert.deepEqual([oTable.CurCell.Row.Index, oTable.CurCell.Index], [0, 0], 'Check move left in table');
	}

	function checkMoveCursorRight(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, false, true, false);
		oAssert.true(oPos.X === 5 && oPos.Y === 15, 'Check move cursor to right position shortcut');
	}

	function checkSelectCursorRight(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
		oAssert.strictEqual(sSelectedText, 'H', 'Check select text to right position shortcut');
	}

	function checkSelectWordCursorRight(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
		oAssert.strictEqual(sSelectedText, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', 'Check select word text to right position shortcut');
	}

	function checkMoveCursorWordRight(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, true, true, false);
		oAssert.true(oPos.X === 25 && oPos.Y === 75, 'Check move cursor to right word position shortcut');
	}

	function checkMoveCursorRightTable(oEvent, oAssert)
	{
		const oFrame = createTable(3, 3);
		oFrame.Set_CurrentElement();
		const oTable = oFrame.graphicObject;
		oTable.MoveCursorToStartPos();
		moveCursorRight(true);
		onKeyDown(oEvent);
		oAssert.deepEqual([oTable.CurCell.Row.Index, oTable.CurCell.Index], [0, 1], 'Check move right in table');
	}

	function checkMoveCursorTop(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, true, true, false);
		oAssert.true(oPos.X === 25 && oPos.Y === 55, 'Check move cursor to top position shortcut');
	}

	function checkSelectCursorTop(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
		oAssert.strictEqual(sSelectedText, 'worldHelloworldHello', 'Check select text to top position shortcut');
	}

	function checkMoveCursorTopTable(oEvent, oAssert)
	{
		const oFrame = createTable(3, 3);
		oFrame.Set_CurrentElement();
		const oTable = oFrame.graphicObject;
		oTable.MoveCursorToStartPos();
		moveCursorDown();
		onKeyDown(oEvent);
		oAssert.deepEqual([oTable.CurCell.Row.Index, oTable.CurCell.Index], [0, 0], 'Check move top in table');
	}

	function checkMoveCursorBottom(oEvent, oAssert)
	{
		const {oPos} = testMoveHelper(oEvent, false, true, false);
		oAssert.true(oPos.X === 0 && oPos.Y === 35, 'Check move cursor to bottom position shortcut');
	}

	function checkSelectCursorBottom(oEvent, oAssert)
	{
		const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
		oAssert.strictEqual(sSelectedText, 'HelloworldHelloworld', 'Check select text to bottom position shortcut');
	}

	function checkMoveCursorBottomTable(oEvent, oAssert)
	{
		const oFrame = createTable(3, 3);
		oFrame.Set_CurrentElement();
		const oTable = oFrame.graphicObject;
		oTable.MoveCursorToStartPos();
		onKeyDown(oEvent);
		oAssert.deepEqual([oTable.CurCell.Row.Index, oTable.CurCell.Index], [1, 0], 'Check move bottom in table');
	}

	function checkMoveShapeBottom(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.y, 5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape bottom');
	}

	function checkLittleMoveShapeBottom(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.y, 1 * AscCommon.g_dKoef_pix_to_mm, 'Check little move shape bottom');
	}

	function checkMoveShapeTop(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.y, -5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape top');
	}

	function checkLittleMoveShapeTop(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.y, -1 * AscCommon.g_dKoef_pix_to_mm, 'Check  move shape top');
	}

	function checkMoveShapeRight(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.x, 5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape right');
	}

	function checkLittleMoveShapeRight(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.x, 1 * AscCommon.g_dKoef_pix_to_mm, 'Check little move shape right');
	}

	function checkMoveShapeLeft(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.x, -5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape left');
	}

	function checkLittleMoveShapeLeft(oEvent, oAssert)
	{
		const oShape = executeCheckMoveShape(oEvent);
		oAssert.strictEqual(oShape.x, -1 * AscCommon.g_dKoef_pix_to_mm, 'Check little move shape left');
	}

	function checkDeleteFront(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello world');
		moveToParagraph(oParagraph, true);

		onKeyDown(oEvent);
		oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'ello world', 'Check delete front shortcut');
	}

	function checkDeleteWordFront(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello world');
		moveToParagraph(oParagraph, true);

		onKeyDown(oEvent);
		oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'world', 'Check delete front word shortcut');
	}

	function checkIncreaseIndent(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello');
		oParagraph.Pr.SetInd(0, 0, 0);
		oParagraph.Set_PresentationLevel(0);
		moveToParagraph(oParagraph, true);

		onKeyDown(oEvent);
		const oParaPr = oGlobalLogicDocument.GetDirectParaPr();
		oAssert.strictEqual(oParaPr.GetIndLeft(), 11.1125, 'Check increase indent');
	}

	function checkDecreaseIndent(oEvent, oAssert)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello');
		oParagraph.Pr.SetInd(0, 12, 0);
		oParagraph.Set_PresentationLevel(1);
		moveToParagraph(oParagraph, true);

		onKeyDown(oEvent);
		const oParaPr = oGlobalLogicDocument.GetDirectParaPr();
		oAssert.true(AscFormat.fApproxEqual(oParaPr.GetIndLeft(), 0.8875), 'Check decrease indent');
	}

	function checkNumLock(oEvent, oAssert)
	{
		onKeyDown(oEvent);
		oAssert.true(oEvent.isDefaultPrevented, 'Check prevent default on num lock');
	}

	function checkScrollLock(oEvent, oAssert)
	{
		onKeyDown(oEvent);
		oAssert.true(oEvent.isDefaultPrevented, 'Check prevent default on scroll lock');
	}

	$(function ()
	{
		QUnit.module('Check shortcut focus', {
			before: function ()
			{
				addSlide();
			},
			after : function ()
			{
				cleanPresentation();
			}
		});
		QUnit.test('check shortcut focus', (oAssert) =>
		{
			editor.StartDemonstration("presentation-preview", 0);
			let bCheck = false;
			let fOldKeyDown;
			fOldKeyDown = editor.WordControl.DemonstrationManager.onKeyDown;
			editor.WordControl.DemonstrationManager.onKeyDown = function ()
			{
				bCheck = true;
			}
			editor.WordControl.onKeyDown(createNativeEvent());
			oAssert.true(bCheck, 'Check demonstration onKeyDown');
			editor.WordControl.DemonstrationManager.onKeyDown = fOldKeyDown;
			editor.EndDemonstration();

			bCheck = false;
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			fOldKeyDown = editor.WordControl.Thumbnails.onKeyDown;
			editor.WordControl.Thumbnails.onKeyDown = function ()
			{
				bCheck = true;
			}
			editor.WordControl.onKeyDown(createNativeEvent());
			oAssert.true(bCheck, 'Check thumbnails onKeyDown');
			editor.WordControl.Thumbnails.onKeyDown = fOldKeyDown;

			bCheck = false;
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			fOldKeyDown = editor.WordControl.m_oLogicDocument.OnKeyDown;
			editor.WordControl.m_oLogicDocument.OnKeyDown = function ()
			{
				bCheck = true;
			}
			editor.WordControl.onKeyDown(createNativeEvent());
			oAssert.true(bCheck, 'Check logic document onKeyDown');
			editor.WordControl.m_oLogicDocument.OnKeyDown = fOldKeyDown;
		});

		QUnit.module("Test thumbnails shortcuts", {
			beforeEach: function ()
			{
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			},
			afterEach : function ()
			{
				cleanPresentation();
			}
		});

		QUnit.test('test thumbnails shortcuts', (oAssert) =>
		{
			let oEvent;
			oEvent = createNativeEvent(77, true, false, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			const arrOldSlides = oGlobalLogicDocument.Slides.slice();
			onKeyDown(oEvent);
			const arrSelectedSlides = oGlobalLogicDocument.GetSelectedSlides();
			oAssert.true(checkSelectedSlides([1]) && (arrOldSlides.indexOf(oGlobalLogicDocument.Slides[arrSelectedSlides[0]]) === -1), 'check add next slide');

			oEvent = createNativeEvent(38, false, false, false, false, false, false);
			goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([3]), 'Check move to previous slide');

			oEvent = createNativeEvent(39, false, false, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([1]), 'Check move to next slide');

			oEvent = createNativeEvent(36, false, false, false, false, false, false);
			goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0]), 'Check move to first slide');

			oEvent = createNativeEvent(36, false, true, false, false, false, false);
			goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0, 1, 2, 3, 4]), 'Check select to first slide');

			oEvent = createNativeEvent(37, false, false, false, false, false, false);
			goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([3]), 'Check move to previous slide');

			oEvent = createNativeEvent(40, false, false, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([1]), 'Check move to next slide');

			oEvent = createNativeEvent(35, false, false, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([4]), 'Check move to last slide');

			oEvent = createNativeEvent(35, false, true, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0, 1, 2, 3, 4]), 'Check select to last slide');

			oEvent = createNativeEvent(33, false, false, false, false, false, false);
			goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([3]), 'Check move to previous slide');

			oEvent = createNativeEvent(34, false, false, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([1]), 'Check move to next slide');
		});

		QUnit.test('Test thumbnails shortcut actions', (oAssert) =>
		{
			const fOldShortcut = editor.getShortcut;
			goToPage(0);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditSelectAll;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			onKeyDown(createNativeEvent());
			oAssert.true(checkSelectedSlides([0, 1, 2, 3]), 'Check select all slides');

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Duplicate;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			const arrOldSlides = oGlobalLogicDocument.Slides.slice();
			onKeyDown(createNativeEvent());
			oAssert.true(checkSelectedSlides([1]) && oGlobalLogicDocument.Slides.length === 5 && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check duplicate slides');

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Print;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			executeTestWithCatchEvent('asc_onPrint', () => true, true, createNativeEvent(), oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Save;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			checkSave(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.ShowContextMenu;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oMockEvent, oAssert);

			editor.getShortcut = fOldShortcut;
		});


		QUnit.test('Test thumbnails hotkeys', (oAssert) =>
		{
			let oEvent;
			let arrOldSlides;
			let oOldSlide;

			oEvent = createNativeEvent(13, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			arrOldSlides = oGlobalLogicDocument.Slides.slice();
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([1]) && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check add next slide');

			oEvent = createNativeEvent(46, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			oOldSlide = getFirstSlide();
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0]) && oGlobalLogicDocument.Slides.indexOf(oOldSlide) === -1, 'Check remove selected slides');

			oEvent = createNativeEvent(8, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			oOldSlide = getFirstSlide();
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0]) && oGlobalLogicDocument.Slides.indexOf(oOldSlide) === -1, 'Check remove selected slides');

			const checkMoveSelectedSlidesToEnd = () =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				const oFirstSlide = getFirstSlide();
				onKeyDown(oEvent);
				oAssert.true(oGlobalLogicDocument.Slides[2] === oFirstSlide, 'Check move selected slides to end');
			};
			oEvent = createNativeEvent(34, true, true, false, false);
			checkMoveSelectedSlidesToEnd();
			oEvent = createNativeEvent(40, true, true, false, false);
			checkMoveSelectedSlidesToEnd();

			const checkMoveSelectedSlidesToNextPos = () =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				const oFirstSlide = getFirstSlide();
				onKeyDown(oEvent);
				oAssert.true(oGlobalLogicDocument.Slides[1] === oFirstSlide, 'Check move selected slides to next pos');
			}
			oEvent = createNativeEvent(34, true, false, false, false);
			checkMoveSelectedSlidesToNextPos();
			oEvent = createNativeEvent(40, true, false, false, false);
			checkMoveSelectedSlidesToNextPos();

			const checkSelectNextSlide = () =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0, 1]), 'Check select next slide');
			}
			oEvent = createNativeEvent(34, false, true, false, false);
			checkSelectNextSlide();
			oEvent = createNativeEvent(40, false, true, false, false);
			checkSelectNextSlide();

			const checkMoveToNextSlide = () =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1]), 'Check move to next slide');
			};
			oEvent = createNativeEvent(34, true, false, false, false);
			checkMoveToNextSlide();
			oEvent = createNativeEvent(40, true, false, false, false);
			checkMoveToNextSlide();

			oEvent = createNativeEvent(36, false, false, false, false);
			goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0]), 'Check move to first slide');

			oEvent = createNativeEvent(36, false, true, false, false);
			goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0, 1, 2]), 'Check select from current position to first slide');

			oEvent = createNativeEvent(35, false, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([2]), 'Check move to last slide');

			oEvent = createNativeEvent(35, false, true, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([0, 1, 2]), 'Check select from current position to last slide');


			const checkMoveSelectedSlidesToStart = () =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				const oLastSlide = oGlobalLogicDocument.Slides[2];
				onKeyDown(oEvent);
				oAssert.true(getFirstSlide() === oLastSlide, 'Check move selected slides to start');
			}
			oEvent = createNativeEvent(33, true, true, false, false);
			checkMoveSelectedSlidesToStart();
			oEvent = createNativeEvent(38, true, true, false, false);
			checkMoveSelectedSlidesToStart();


			const checkMoveSelectedSlidesToPreviousPosition = () =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);

				const oLastSlide = oGlobalLogicDocument.Slides[2];
				onKeyDown(oEvent);
				oAssert.true(oGlobalLogicDocument.Slides[1] === oLastSlide, 'Check move selected slides to previous pos');
			};
			oEvent = createNativeEvent(33, true, false, false, false);
			checkMoveSelectedSlidesToPreviousPosition();
			oEvent = createNativeEvent(38, true, false, false, false);
			checkMoveSelectedSlidesToPreviousPosition();

			const checkSelectPreviousSlide = () =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);

				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1, 2]), 'Check select previous slide');
			}
			oEvent = createNativeEvent(33, false, true, false, false);
			checkSelectPreviousSlide();
			oEvent = createNativeEvent(38, false, true, false, false);
			checkSelectPreviousSlide();

			const checkMoveToPreviousSlide = () =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1]), 'Check move to previous slide');
			};
			oEvent = createNativeEvent(33, true, false, false, false);
			checkMoveToPreviousSlide();
			oEvent = createNativeEvent(38, true, false, false, false);
			checkMoveToPreviousSlide();

			oEvent = createNativeEvent(77, true, false, false, false);
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			arrOldSlides = oGlobalLogicDocument.Slides.slice();
			onKeyDown(oEvent);
			oAssert.true(checkSelectedSlides([1]) && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check add next slide');
		});

		QUnit.module('Test demonstration mode shortcuts', {
			beforeEach: function ()
			{
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
			},
			afterEach : function ()
			{
				cleanPresentation();
			}
		});

		QUnit.test('Test demonstration mode shortcuts', (oAssert) =>
		{
			let oEvent;

			editor.StartDemonstration("presentation-preview", 0);
			oEvent = createNativeEvent(13, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 1, oEvent, oAssert);

			oEvent = createNativeEvent(32, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 2, oEvent, oAssert);

			oEvent = createNativeEvent(34, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 3, oEvent, oAssert);

			oEvent = createNativeEvent(39, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, oEvent, oAssert);

			oEvent = createNativeEvent(40, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 5, oEvent, oAssert);

			oEvent = createNativeEvent(33, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, oEvent, oAssert);

			oEvent = createNativeEvent(37, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 3, oEvent, oAssert);

			oEvent = createNativeEvent(38, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 2, oEvent, oAssert);

			oEvent = createNativeEvent(36, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 0, oEvent, oAssert);

			oEvent = createNativeEvent(35, false, false, false, false);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 5, oEvent, oAssert);

			oEvent = createNativeEvent(27, false, false, false, false);
			executeTestWithCatchEvent('asc_onEndDemonstration', () => true, true, oEvent, oAssert);

			editor.EndDemonstration();
		});

		QUnit.module("Test main focus shortcuts", {
			beforeEach: function ()
			{
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			},
			afterEach : function ()
			{
				cleanPresentation();
			}
		});
		QUnit.test('Test if the desired action is received by the keyboard shortcut.', (oAssert) =>
		{
			let oEvent;
			oEvent = createEvent(65, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EditSelectAll, 'Check getting select all shortcut action');

			oEvent = createEvent(90, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EditUndo, 'Check getting undo shortcut action');

			oEvent = createEvent(89, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EditRedo, 'Check getting redo shortcut action');

			oEvent = createEvent(88, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Cut, 'Check getting cut shortcut action');

			oEvent = createEvent(67, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Copy, 'Check getting copy shortcut action');

			oEvent = createEvent(86, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Paste, 'Check getting paste shortcut action');

			oEvent = createEvent(68, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Duplicate, 'Check getting duplicate shortcut action');

			oEvent = createEvent(80, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Print, 'Check getting print shortcut action');

			oEvent = createEvent(83, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Save, 'Check getting save shortcut action');

			oEvent = createEvent(93, false, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowContextMenu, 'Check getting show context menu shortcut action');

			oEvent = createEvent(121, false, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowContextMenu, 'Check getting show context menu shortcut action');

			oEvent = createEvent(57351, false, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowContextMenu, 'Check getting show context menu shortcut action');

			oEvent = createEvent(56, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowParaMarks, 'Check getting show paragraph marks shortcut action');

			oEvent = createEvent(66, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Bold, 'Check getting bold shortcut action');

			oEvent = createEvent(67, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.CopyFormat, 'Check getting copy format shortcut action');

			oEvent = createEvent(69, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.CenterAlign, 'Check getting center align shortcut action');

			oEvent = createEvent(69, true, false, true, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EuroSign, 'Check getting euro sign shortcut action');

			oEvent = createEvent(71, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Group, 'Check getting group shortcut action');

			oEvent = createEvent(71, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.UnGroup, 'Check getting ungroup shortcut action');

			oEvent = createEvent(73, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Italic, 'Check getting italic shortcut action');

			oEvent = createEvent(74, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.JustifyAlign, 'Check getting justify align shortcut action');

			oEvent = createEvent(75, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.AddHyperlink, 'Check getting add hyperlink shortcut action');

			oEvent = createEvent(76, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.BulletList, 'Check getting bullet list shortcut action');

			oEvent = createEvent(76, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.LeftAlign, 'Check getting left align shortcut action');

			oEvent = createEvent(82, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.RightAlign, 'Check getting right align shortcut action');

			oEvent = createEvent(85, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Underline, 'Check getting underline shortcut action');

			oEvent = createEvent(53, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Strikethrough, 'Check getting strikethrough shortcut action');

			oEvent = createEvent(83, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.PasteFormat, 'Check getting paste format shortcut action');

			oEvent = createEvent(187, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Superscript, 'Check getting superscript shortcut action');

			oEvent = createEvent(188, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Superscript, 'Check getting superscript shortcut action');

			oEvent = createEvent(187, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Subscript, 'Check getting subscript shortcut action');

			oEvent = createEvent(190, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Subscript, 'Check getting subscript shortcut action');

			oEvent = createEvent(189, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EnDash, 'Check getting en dash shortcut action');

			oEvent = createEvent(219, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.DecreaseFont, 'Check getting decrease font size shortcut action');

			oEvent = createEvent(221, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.IncreaseFont, 'Check getting increase font size shortcut action');
		});

		QUnit.test('Test main shortcut actions', (oAssert) =>
		{
			const fOldGetShortcut = editor.getShortcut;
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditUndo;};
			checkEditUndo(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditRedo;};
			checkEditRedo(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditSelectAll;};
			checkEditSelectAll(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Duplicate;};
			checkDuplicate(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Print;};
			checkPrint(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Save;};
			checkSave(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.ShowContextMenu;};
			checkShowContextMenu(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.ShowParaMarks;};
			checkShowParaMarks(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Bold;};
			checkBold(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.CopyFormat;};
			checkCopyFormat(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.PasteFormat;};
			checkPasteFormat(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.CenterAlign;};
			checkCenterAlign(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EuroSign;};
			checkEuroSign(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Group;};
			checkGroup(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.UnGroup;};
			checkUnGroup(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Italic;};
			checkItalic(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.JustifyAlign;};
			checkJustifyAlign(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.AddHyperlink;};
			checkAddHyperlink(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.BulletList;};
			checkBulletList(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.LeftAlign;};
			checkLeftAlign(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.RightAlign;};
			checkRightAlign(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Underline;};
			checkUnderline(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Strikethrough;};
			checkStrikethrough(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Superscript;};
			checkSuperscript(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Subscript;};
			checkSubscript(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EnDash;};
			checkEnDash(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.DecreaseFont;};
			checkDecreaseFont(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.IncreaseFont;};
			checkIncreaseFont(oAssert);

			editor.getShortcut = fOldGetShortcut;
		});

		QUnit.test('Test common shortcuts', function (oAssert)
		{
			editor.initDefaultShortcuts();
			let oEvent;
			oEvent = createNativeEvent(8, false, false, false, false, false, false);
			checkDeleteBack(oEvent, oAssert);

			oEvent = createNativeEvent(8, true, false, false, false, false, false);
			checkDeleteWordBack(oEvent, oAssert);

			oEvent = createNativeEvent(8, false, false, false, false, false);
			checkRemoveAnimation(oEvent, oAssert);

			oEvent = createNativeEvent(8, false, false, false, false, false);
			checkRemoveChart(oEvent, oAssert);

			oEvent = createNativeEvent(8, false, false, false, false, false);
			checkRemoveShape(oEvent, oAssert);

			oEvent = createNativeEvent(8, false, false, false, false, false);
			checkRemoveTable(oEvent, oAssert);

			oEvent = createNativeEvent(8, false, false, false, false, false);
			checkRemoveGroup(oEvent, oAssert);

			oEvent = createNativeEvent(8, false, false, false, false, false);
			checkRemoveShapeInGroup(oEvent, oAssert);

			//Tab
			oEvent = createNativeEvent(9, false, false, false, false, false, false);
			checkMoveToNextCell(oEvent, oAssert);

			oEvent = createNativeEvent(9, false, true, false, false, false, false);
			checkMoveToPreviousCell(oEvent, oAssert);

			oEvent = createNativeEvent(9, false, false, false, false, false, false);
			checkIncreaseBulletIndent(oEvent, oAssert);

			oEvent = createNativeEvent(9, false, true, false, false, false, false);
			checkDecreaseBulletIndent(oEvent, oAssert);

			oEvent = createNativeEvent(9, false, false, false, false, false, false);
			checkAddTab(oEvent, oAssert);

			oEvent = createNativeEvent(9, false, false, false, false, false, false);
			checkSelectNextObject(oEvent, oAssert);

			oEvent = createNativeEvent(9, false, true, false, false, false, false);
			checkSelectPreviousObject(oEvent, oAssert);
			// Enter
			oEvent = createNativeEvent(13, false, false, false, false, false, false);
			checkVisitHyperlink(oEvent, oAssert);

			oEvent = createNativeEvent(13, true, false, false, false, false, false);
			checkSelectNextObjectWithPlaceholder(oEvent, oAssert);

			oEvent = createNativeEvent(13, true, false, false, false, false, false);
			checkAddNextSlideAfterSelectLastPlaceholderObject(oEvent, oAssert);

			oEvent = createNativeEvent(13, false, true, false, false, false, false);
			checkAddBreakLine(oEvent, oAssert);

			oEvent = createNativeEvent(13, false, true, false, false, false, false);
			checkAddMathBreakLine(oEvent, oAssert);

			oEvent = createNativeEvent(13, false, false, false, false, false, false);
			checkAddTitleBreakLine(oEvent, oAssert);

			oEvent = createNativeEvent(13, false, false, false, false, false, false);
			checkAddMathBreakLine(oEvent, oAssert);

			oEvent = createNativeEvent(13, false, false, false, false, false, false);
			checkAddParagraph(oEvent, oAssert);

			oEvent = createNativeEvent(13, false, false, false, false, false, false);
			checkAddTxBodyShape(oEvent, oAssert);
			checkMoveCursorToStartPosShape(oEvent, oAssert);
			checkSelectAllContentShape(oEvent, oAssert);
			checkSelectAllContentChartTitle(oEvent, oAssert);
			checkMoveCursorToStartPosChartTitle(oEvent, oAssert);
			checkRemoveAndMoveToStartPosTable(oEvent, oAssert);
			checkSelectFirstCellContent(oEvent, oAssert);
			// Esc
			oEvent = createNativeEvent(27, false, false, false, false, false, false);
			checkResetAddShape(oEvent, oAssert);

			oEvent = createNativeEvent(27, false, true, false, false, false, false);
			checkResetAllDrawingSelection(oEvent, oAssert);

			oEvent = createNativeEvent(27, false, false, false, false, false, false);
			checkResetStepDrawingSelection(oEvent, oAssert);

			// Space
			oEvent = createNativeEvent(32, true, true, false, false, false, false);
			checkNonBreakingSpace(oEvent, oAssert);

			oEvent = createNativeEvent(32, true, false, false, false, false, false);
			checkClearParagraphFormatting(oEvent, oAssert);

			oEvent = createNativeEvent(32, false, false, false, false, false, false);
			checkAddSpace(oEvent, oAssert);
			//pgUp

			//End
			oEvent = createNativeEvent(35, true, false, false, false, false, false);
			checkMoveToEndPosContent(oEvent, oAssert);

			oEvent = createNativeEvent(35, false, false, false, false, false, false);
			checkMoveToEndLineContent(oEvent, oAssert);

			oEvent = createNativeEvent(35, false, true, false, false, false, false);
			checkSelectToEndLineContent(oEvent, oAssert);

			// Home
			oEvent = createNativeEvent(36, true, false, false, false, false, false);
			checkMoveToStartPosContent(oEvent, oAssert);

			oEvent = createNativeEvent(36, false, false, false, false, false, false);
			checkMoveToStartLineContent(oEvent, oAssert);

			oEvent = createNativeEvent(36, false, true, false, false, false, false);
			checkSelectToStartLineContent(oEvent, oAssert);

			//Left arrow
			oEvent = createNativeEvent(37, false, false, false, false, false, false);
			checkMoveCursorLeft(oEvent, oAssert);

			oEvent = createNativeEvent(37, false, true, false, false, false, false);
			checkSelectCursorLeft(oEvent, oAssert);

			oEvent = createNativeEvent(37, true, true, false, false, false, false);
			checkSelectWordCursorLeft(oEvent, oAssert);

			oEvent = createNativeEvent(37, true, false, false, false, false, false);
			checkMoveCursorWordLeft(oEvent, oAssert);
			checkMoveCursorLeftTable(oEvent, oAssert);
			//Right arrow
			oEvent = createNativeEvent(39, false, false, false, false, false, false);
			checkMoveCursorRight(oEvent, oAssert);
			checkMoveCursorRightTable(oEvent, oAssert);

			oEvent = createNativeEvent(39, false, true, false, false, false, false);
			checkSelectCursorRight(oEvent, oAssert);

			oEvent = createNativeEvent(39, true, true, false, false, false, false);
			checkSelectWordCursorRight(oEvent, oAssert);

			oEvent = createNativeEvent(39, true, false, false, false, false, false);
			checkMoveCursorWordRight(oEvent, oAssert);
			//Top arrow
			oEvent = createNativeEvent(38, false, false, false, false, false, false);
			checkMoveCursorTop(oEvent, oAssert);
			checkMoveCursorTopTable(oEvent, oAssert);

			oEvent = createNativeEvent(38, false, true, false, false, false, false);
			checkSelectCursorTop(oEvent, oAssert);
			// Bottom arrow
			oEvent = createNativeEvent(40, false, false, false, false, false, false);
			checkMoveCursorBottom(oEvent, oAssert);
			checkMoveCursorBottomTable(oEvent, oAssert);

			oEvent = createNativeEvent(40, false, true, false, false, false, false);
			checkSelectCursorBottom(oEvent, oAssert);

			// Check move shape
			oEvent = createNativeEvent(40, false, false, false, false, false, false);
			checkMoveShapeBottom(oEvent, oAssert);

			oEvent = createNativeEvent(40, true, false, false, false, false, false);
			checkLittleMoveShapeBottom(oEvent, oAssert);

			oEvent = createNativeEvent(38, false, false, false, false, false, false);
			checkMoveShapeTop(oEvent, oAssert);

			oEvent = createNativeEvent(38, true, false, false, false, false, false);
			checkLittleMoveShapeTop(oEvent, oAssert);

			oEvent = createNativeEvent(39, false, false, false, false, false, false);
			checkMoveShapeRight(oEvent, oAssert);

			oEvent = createNativeEvent(39, true, false, false, false, false, false);
			checkLittleMoveShapeRight(oEvent, oAssert);

			oEvent = createNativeEvent(37, false, false, false, false, false, false);
			checkMoveShapeLeft(oEvent, oAssert);

			oEvent = createNativeEvent(37, true, false, false, false, false, false);
			checkLittleMoveShapeLeft(oEvent, oAssert);

			//Delete
			oEvent = createNativeEvent(46, false, false, false, false, false, false);
			checkDeleteFront(oEvent, oAssert);

			oEvent = createNativeEvent(46, true, false, false, false, false, false);
			checkDeleteWordFront(oEvent, oAssert);
			checkRemoveAnimation(oEvent, oAssert);
			checkRemoveChart(oEvent, oAssert);
			checkRemoveShape(oEvent, oAssert);
			checkRemoveTable(oEvent, oAssert);
			checkRemoveGroup(oEvent, oAssert);
			checkRemoveShapeInGroup(oEvent, oAssert);

			oEvent = createNativeEvent(77, true, false, false, false, false, false);
			checkIncreaseIndent(oEvent, oAssert);

			oEvent = createNativeEvent(77, true, true, false, false, false, false);
			checkDecreaseIndent(oEvent, oAssert);

			oEvent = createNativeEvent(144, false, false, false, false, false, false);
			checkNumLock(oEvent, oAssert);

			oEvent = createNativeEvent(145, false, false, false, false, false, false);
			checkScrollLock(oEvent, oAssert);
		});
	});
})(window);
