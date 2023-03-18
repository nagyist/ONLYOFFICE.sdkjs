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

	AscCommon.CGraphics.prototype.SetFontSlot = function () {};
	AscCommon.CGraphics.prototype.SetFont = function () {};
	AscCommon.CGraphics.prototype.SetFontInternal = function () {};
	window.AscFonts = window.AscFonts || {};
	AscFonts.g_fontApplication = {
		GetFontInfo    : function (sFontName)
		{
			if (sFontName === 'Cambria Math')
			{
				return new AscFonts.CFontInfo('Cambria Math', 40, 1, 433, 1, -1, -1, -1, -1, -1, -1);
			}
		},
		Init           : function () {},
		LoadFont       : function () {},
		GetFontInfoName: function () {}
	}

	window.g_fontApplication = AscFonts.g_fontApplication;

	AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback)
	{
		callback({'saveLock': false});
	};
	let oGlobalShape
	const oGlobalLogicDocument = AscTest.CreateLogicDocument()
	editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState = function ()
	{
	};

	function getController()
	{
		return oGlobalLogicDocument.GetCurrentController();
	}

	function createGroup(arrObjects)
	{
		const oController = getController();
		oController.resetSelection();
		for (let i = 0; i < arrObjects.length; i += 1)
		{
			arrObjects[i].select(oController, 0);
		}

		return oController.createGroup();
	}

	function addToSelection(oObject)
	{
		const oController = getController();
		if (oObject.group)
		{
			const oMainGroup = oObject.group.getMainGroup();
			oMainGroup.select(oController, 0);
			oController.selection.groupSelection = oMainGroup;
		}
		oObject.select(oController, 0);
	}

	function remove()
	{
		oGlobalLogicDocument.Remove();
	}

	function selectOnlyObjects(arrObjects)
	{
		const oController = getController();
		oController.resetSelection();
		for (let i = 0; i < arrObjects.length; i += 1)
		{
			const oObject = arrObjects[i];
			if (oObject.group)
			{
				const oMainGroup = oObject.group.getMainGroup();
				oMainGroup.select(oController, 0);
				oController.selection.groupSelection = oMainGroup;
			}
			oObject.select(oController, 0);
		}
	}

	function getFirstSlide()
	{
		return oGlobalLogicDocument.Slides[0];
	}

	function moveToParagraph(oParagraph, bIsStart)
	{
		oParagraph.SetThisElementCurrent();
		if (bIsStart)
		{
			oParagraph.MoveCursorToStartPos();
		} else
		{
			oParagraph.MoveCursorToEndPos();
		}
	}

	function getShapeWithParagraphHelper(sTextIntoShape, bResetSelection)
	{
		const oController = oGlobalLogicDocument.GetCurrentController();
		if (bResetSelection)
		{
			oController.resetSelection();
		}
		oGlobalShape = AscTest.createShape(oGlobalLogicDocument.Slides[0]);


		oGlobalShape.setTxBody(AscFormat.CreateTextBodyFromString(sTextIntoShape, editor.WordControl.m_oDrawingDocument, oGlobalShape));
		const oContent = oGlobalShape.getDocContent();
		const oParagraph = oContent.Content[0];

		return {oLogicDocument: oGlobalLogicDocument, oShape: oGlobalShape, oParagraph, oController, oContent};
	}

	function createEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsAltGr, bIsMacCmdKey)
	{
		const oKeyBoardEvent = new AscCommon.CKeyboardEvent();
		oKeyBoardEvent.KeyCode = nKeyCode;
		oKeyBoardEvent.ShiftKey = bIsShift;
		oKeyBoardEvent.AltKey = bIsAlt;
		oKeyBoardEvent.CtrlKey = bIsCtrl;
		oKeyBoardEvent.MacCmdKey = bIsMacCmdKey;
		oKeyBoardEvent.AltGr = bIsAltGr;
		return oKeyBoardEvent;
	}

	function getDirectTextPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();
		onKeyDown(oEvent);
		return oGlobalLogicDocument.GetDirectTextPr();
	}

	function getDirectParaPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();
		onKeyDown(oEvent);
		return oGlobalLogicDocument.GetDirectParaPr();
	}

	function checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, sInitText)
	{
		const {oLogicDocument, oParagraph} = getShapeWithParagraphHelper(sInitText);
		oParagraph.SetThisElementCurrent();
		oLogicDocument.MoveCursorToEndPos();
		onKeyDown(oEvent);
		const sTextAfterKeyDown = AscTest.GetParagraphText(oParagraph);
		oAssert.strictEqual(sTextAfterKeyDown, sCheckText, sPrompt);
	}

	function checkTextAfterKeyDownHelperEmpty(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, '');
	}

	function checkTextAfterKeyDownHelperHelloWorld(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, 'Hello World');
	}

	function checkRemoveObject(oObject, arrSpTree)
	{
		let bCheckRemoveFromSpTree = true;
		for (let nDrawingIndex = 0; nDrawingIndex < arrSpTree.length; nDrawingIndex += 1)
		{
			if (arrSpTree[nDrawingIndex] === oObject)
			{
				bCheckRemoveFromSpTree = false;
			}
		}
		return bCheckRemoveFromSpTree && oObject.bDeleted;
	}

	function createTable(nRows, nColumns)
	{
		const oGraphicFrame = oGlobalLogicDocument.Add_FlowTable(nColumns, nRows);
		const oTable = oGraphicFrame.graphicObject;
		//oTable.Resize(nColumns * 300, nRows * 200);
		for (let nRow = 0; nRow < nRows; nRow += 1)
		{
			for (let nColumn = 0; nColumn < nColumns; nColumn += 1)
			{
				const oCell = oTable.Content[nRow].Get_Cell(nColumn);
				const oContent = oCell.GetContent();
				AscFormat.AddToContentFromString(oContent, 'Cell' + nRow + 'x' + nColumn);
			}
		}
		oGlobalLogicDocument.Recalculate();
		return oGraphicFrame;
	}

	function createChart()
	{
		const oChart = AscCommon.getChartByType(Asc.c_oAscChartTypeSettings.lineNormal);
		oChart.setParent(oGlobalLogicDocument.Slides[0]);

		oChart.addToDrawingObjects();
		oChart.spPr.setXfrm(new AscFormat.CXfrm());
		oChart.spPr.xfrm.setOffX(0);
		oChart.spPr.xfrm.setOffY(0);
		oChart.spPr.xfrm.setExtX(100);
		oChart.spPr.xfrm.setExtY(100);
		oGlobalLogicDocument.Recalculate();
		oGlobalLogicDocument.Document_UpdateInterfaceState();
		oGlobalLogicDocument.CheckEmptyPlaceholderNotes();

		oGlobalLogicDocument.DrawingDocument.m_oWordControl.OnUpdateOverlay();
		return oChart;
	}

	function createShapeWithTitlePlaceholder()
	{
		const oShape = AscTest.createShape(oGlobalLogicDocument.Slides[0]);
		oShape.setNvSpPr(new AscFormat.UniNvPr());
		let oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_title);
		oShape.nvSpPr.nvPr.setPh(oPh);
		oShape.txBody = AscFormat.CreateTextBodyFromString('', oShape.getDrawingDocument(), oShape);

		oShape.recalculateContentWitCompiledPr();
		return oShape;
	}

	function testMoveHelper(oEvent, bMoveToEndPosition, bGetPos, bGetSelectedText)
	{
		const {
			oShape,
			oParagraph,
			oLogicDocument
		} = getShapeWithParagraphHelper('HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', true);
		oShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});
		oParagraph.SetThisElementCurrent();
		oParagraph.Pr.SetInd(0, 0, 0);
		oParagraph.Set_Align(AscCommon.align_Left);
		if (bMoveToEndPosition)
		{
			oLogicDocument.MoveCursorToEndPos();
		} else
		{
			oLogicDocument.MoveCursorToStartPos();
		}
		oShape.recalculateContentWitCompiledPr();
		oLogicDocument.RecalculateCurPos(true, true);

		onKeyDown(oEvent);

		let oPos;
		oLogicDocument.RecalculateCurPos(true, true);
		if (bGetPos)
		{

			oPos = oParagraph.GetCurPosXY(true, true);
		}
		let sSelectedText;
		if (bGetSelectedText)
		{
			sSelectedText = oParagraph.GetSelectedText();
		}
		return {oPos, sSelectedText};
	}

	function addPropertyToDocument(oPr)
	{
		oGlobalLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(oPr), true);
	}

	function executeCheckMoveShape(oEvent)
	{
		const oController = oGlobalLogicDocument.GetCurrentController();
		oController.resetSelection();
		oGlobalShape.spPr.xfrm.setOffX(0);
		oGlobalShape.spPr.xfrm.setOffY(0);
		oGlobalShape.select(oController, 0);
		oGlobalShape.recalculateTransform();
		editor.zoom100();
		onKeyDown(oEvent);
		return oGlobalShape;
	}

	function cleanPresentation()
	{
		goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
		editor.WordControl.Thumbnails.SelectAll();
		const arrSelectedArray = oGlobalLogicDocument.GetSelectedSlides();
		oGlobalLogicDocument.deleteSlides(arrSelectedArray);
	}

	function checkSelectedSlides(arrSelectedSlides)
	{
		const arrPresentationSelectedSlides = oGlobalLogicDocument.GetSelectedSlides();
		return arrSelectedSlides.length === arrPresentationSelectedSlides.length && arrSelectedSlides.every((el, ind) => el === arrPresentationSelectedSlides[ind]);
	}

	const arrCheckCodes = [48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 189, 187, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83,
		84, 85, 86, 87, 88, 89, 90, 219, 221, 186, 222, 220, 188, 190, 191, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 111, 106,
		109, 110, 107];

	function createNativeEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsMetaKey)
	{
		const bIsMacOs = AscCommon.AscBrowser.isMacOs;
		const oEvent = {};
		oEvent.isDefaultPrevented = false;
		oEvent.isPropagationStopped = false;
		oEvent.preventDefault = function ()
		{
			if (bIsMacOs && oEvent.altKey && !(oEvent.ctrlKey || oEvent.metaKey) && (arrCheckCodes.indexOf(nKeyCode) !== -1))
			{
				throw new Error('Alt key must not be disabled on macOS');
			}
			oEvent.isDefaultPrevented = true;
		};
		oEvent.stopPropagation = function ()
		{
			oEvent.isPropagationStopped = true;
		};

		oEvent.keyCode = nKeyCode;
		oEvent.ctrlKey = bIsCtrl;
		oEvent.shiftKey = bIsShift;
		oEvent.altKey = bIsAlt;
		oEvent.metaKey = bIsMetaKey;
		return oEvent;
	}

	function executeTestWithCatchEvent(sSendEvent, fCustomCheck, customExpectedValue, oEvent, oAssert, fBeforeCallback)
	{
		fBeforeCallback && fBeforeCallback();

		let bCheck = false;

		const fCheck = function (...args)
		{
			if (fCustomCheck)
			{
				bCheck = fCustomCheck(...args);
			} else
			{
				bCheck = true;
			}
		}
		editor.asc_registerCallback(sSendEvent, fCheck);

		onKeyDown(oEvent);
		oAssert.strictEqual(bCheck, customExpectedValue === undefined ? true : customExpectedValue, 'Check catch ' + sSendEvent + ' event');
		editor.asc_unregisterCallback(sSendEvent, fCheck);
	}

	function onKeyDown(oEvent)
	{
		editor.WordControl.onKeyDown(oEvent);
	}

	function addSlide()
	{
		oGlobalLogicDocument.addNextSlide(0);
	}

	function createShape()
	{
		return AscTest.createShape(oGlobalLogicDocument.Slides[0]);
	}

	function goToPage(nPage)
	{
		editor.WordControl.GoToPage(nPage);
	}

	function goToPageWithFocus(nPage, eFocus)
	{
		goToPage(nPage);
		oGlobalLogicDocument.SetThumbnailsFocusElement(eFocus);
	}

	function onInput(arrCodes)
	{
		for (let i = 0; i < arrCodes.length; i += 1)
		{
			oGlobalLogicDocument.OnKeyPress(createEvent(arrCodes[i], false, false, false, false, false));
		}
	}

	function moveCursorRight(bWord, bAddToSelect)
	{
		oGlobalLogicDocument.MoveCursorRight(!!bAddToSelect, !!bWord);
	}

	function moveCursorLeft(bWord, bAddToSelect)
	{
		oGlobalLogicDocument.MoveCursorLeft(!!bAddToSelect, !!bWord);
	}

	function moveCursorDown(bCtrlKey, bAddToSelect)
	{
		oGlobalLogicDocument.MoveCursorDown(!!bAddToSelect, !!bCtrlKey);
	}

	function createMathInShape()
	{
		const {oShape, oParagraph} = getShapeWithParagraphHelper('', true);
		selectOnlyObjects([oShape]);
		moveToParagraph(oParagraph, true);
		editor.asc_AddMath2(c_oAscMathType.FractionVertical);
		return {oShape, oParagraph};
	}

	function checkDirectTextPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		const oTextPr = getDirectTextPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oTextPr), nExpectedValue, sPrompt);
	}

	function checkDirectParaPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		const oParaPr = getDirectParaPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oParaPr), nExpectedValue, sPrompt);
	}

	const testAll = 0;
	const testMacOs = 1;
	const testWindows = 2;

	function CTestEvent(oEvent, nType)
	{
		this.type = nType || testAll;
		this.event = oEvent;
	}

	const oMainShortcutTypes = {
		checkDeleteBack                                  : 0,
		checkDeleteWordBack                              : 1,
		checkRemoveAnimation                             : 2,
		checkRemoveChart                                 : 3,
		checkRemoveShape                                 : 4,
		checkRemoveTable                                 : 5,
		checkRemoveGroup                                 : 6,
		checkRemoveShapeInGroup                          : 7,
		checkMoveToNextCell                              : 8,
		checkMoveToPreviousCell                          : 9,
		checkIncreaseBulletIndent                        : 10,
		checkDecreaseBulletIndent                        : 11,
		checkAddTab                                      : 12,
		checkSelectNextObject                            : 13,
		checkSelectPreviousObject                        : 14,
		checkVisitHyperlink                              : 15,
		checkSelectNextObjectWithPlaceholder             : 16,
		checkAddNextSlideAfterSelectLastPlaceholderObject: 17,
		checkAddBreakLine                                : 18,
		checkAddTitleBreakLine                           : 19,
		checkAddMathBreakLine                            : 20,
		checkAddParagraph                                : 21,
		checkAddTxBodyShape                              : 22,
		checkMoveCursorToStartPosShape                   : 23,
		checkSelectAllContentShape                       : 24,
		checkSelectAllContentChartTitle                  : 25,
		checkMoveCursorToStartPosChartTitle              : 26,
		checkRemoveAndMoveToStartPosTable                : 27,
		checkSelectFirstCellContent                      : 28,
		checkResetAddShape                               : 29,
		checkResetAllDrawingSelection                    : 30,
		checkResetStepDrawingSelection                   : 31,
		checkNonBreakingSpace                            : 32,
		checkClearParagraphFormatting                    : 33,
		checkAddSpace                                    : 34,
		checkMoveToEndPosContent                         : 35,
		checkMoveToEndLineContent                        : 36,
		checkSelectToEndLineContent                      : 37,
		checkMoveToStartPosContent                       : 38,
		checkMoveToStartLineContent                      : 39,
		checkSelectToStartLineContent                    : 40,
		checkMoveCursorLeft                              : 41,
		checkSelectCursorLeft                            : 42,
		checkSelectWordCursorLeft                        : 43,
		checkMoveCursorWordLeft                          : 44,
		checkMoveCursorLeftTable                         : 45,
		checkMoveCursorRight                             : 46,
		checkMoveCursorRightTable                        : 47,
		checkSelectCursorRight                           : 48,
		checkSelectWordCursorRight                       : 49,
		checkMoveCursorWordRight                         : 50,
		checkMoveCursorTop                               : 51,
		checkMoveCursorTopTable                          : 52,
		checkSelectCursorTop                             : 53,
		checkMoveCursorBottom                            : 54,
		checkMoveCursorBottomTable                       : 55,
		checkSelectCursorBottom                          : 56,
		checkMoveShapeBottom                             : 57,
		checkLittleMoveShapeBottom                       : 58,
		checkMoveShapeTop                                : 59,
		checkLittleMoveShapeTop                          : 60,
		checkMoveShapeRight                              : 61,
		checkLittleMoveShapeRight                        : 62,
		checkMoveShapeLeft                               : 63,
		checkLittleMoveShapeLeft                         : 64,
		checkDeleteFront                                 : 65,
		checkDeleteWordFront                             : 66,
		checkIncreaseIndent                              : 67,
		checkDecreaseIndent                              : 68,
		checkNumLock                                     : 69,
		checkScrollLock                                  : 70
	};
	const oMainEvents = {};
	oMainEvents[oMainShortcutTypes.checkDeleteBack] = [new CTestEvent(createNativeEvent(8, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDeleteWordBack] = [new CTestEvent(createNativeEvent(8, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveAnimation] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveChart] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveShape] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveTable] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveGroup] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveShapeInGroup] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToNextCell] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToPreviousCell] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkIncreaseBulletIndent] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDecreaseBulletIndent] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddTab] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectNextObject] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectPreviousObject] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkVisitHyperlink] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectNextObjectWithPlaceholder] = [new CTestEvent(createNativeEvent(13, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddNextSlideAfterSelectLastPlaceholderObject] = [new CTestEvent(createNativeEvent(13, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddBreakLine] = [new CTestEvent(createNativeEvent(13, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddMathBreakLine] = [
		new CTestEvent(createNativeEvent(13, false, true, false, false, false, false)),
		new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddTitleBreakLine] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddParagraph] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddTxBodyShape] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorToStartPosShape] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectAllContentShape] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectAllContentChartTitle] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorToStartPosChartTitle] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveAndMoveToStartPosTable] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectFirstCellContent] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkResetAddShape] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkResetAllDrawingSelection] = [new CTestEvent(createNativeEvent(27, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkResetStepDrawingSelection] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkNonBreakingSpace] = [new CTestEvent(createNativeEvent(32, true, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkClearParagraphFormatting] = [new CTestEvent(createNativeEvent(32, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddSpace] = [new CTestEvent(createNativeEvent(32, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToEndPosContent] = [new CTestEvent(createNativeEvent(35, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToEndLineContent] = [new CTestEvent(createNativeEvent(35, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectToEndLineContent] = [new CTestEvent(createNativeEvent(35, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToStartPosContent] = [new CTestEvent(createNativeEvent(36, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToStartLineContent] = [new CTestEvent(createNativeEvent(36, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectToStartLineContent] = [new CTestEvent(createNativeEvent(36, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorLeft] = [new CTestEvent(createNativeEvent(37, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorLeft] = [new CTestEvent(createNativeEvent(37, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectWordCursorLeft] = [new CTestEvent(createNativeEvent(37, true, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorWordLeft] = [new CTestEvent(createNativeEvent(37, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorLeftTable] = [new CTestEvent(createNativeEvent(37, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorRight] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorRightTable] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorRight] = [new CTestEvent(createNativeEvent(39, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectWordCursorRight] = [new CTestEvent(createNativeEvent(39, true, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorWordRight] = [new CTestEvent(createNativeEvent(39, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorTop] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorTopTable] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorTop] = [new CTestEvent(createNativeEvent(38, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorBottom] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorBottomTable] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorBottom] = [new CTestEvent(createNativeEvent(40, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeBottom] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeBottom] = [new CTestEvent(createNativeEvent(40, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeTop] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeTop] = [new CTestEvent(createNativeEvent(38, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeRight] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeRight] = [new CTestEvent(createNativeEvent(39, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeLeft] = [new CTestEvent(createNativeEvent(37, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeLeft] = [new CTestEvent(createNativeEvent(37, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDeleteFront] = [new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDeleteWordFront] = [new CTestEvent(createNativeEvent(46, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkIncreaseIndent] = [new CTestEvent(createNativeEvent(77, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDecreaseIndent] = [new CTestEvent(createNativeEvent(77, true, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkNumLock] = [new CTestEvent(createNativeEvent(144, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkScrollLock] = [new CTestEvent(createNativeEvent(145, false, false, false, false, false, false))];

	const oDemonstrationTypes = {
		moveToNextSlide          : 0,
		moveToPreviousSlide      : 1,
		moveToFirstSlide         : 2,
		moveToLastSlide          : 3,
		exitFromDemonstrationMode: 4
	};
	const oDemonstrationEvents = {};
	oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide] = [
		new CTestEvent(createNativeEvent(13, false, false, false, false)),
		new CTestEvent(createNativeEvent(32, false, false, false, false)),
		new CTestEvent(createNativeEvent(34, false, false, false, false)),
		new CTestEvent(createNativeEvent(39, false, false, false, false)),
		new CTestEvent(createNativeEvent(40, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.moveToPreviousSlide] = [
		new CTestEvent(createNativeEvent(33, false, false, false, false)),
		new CTestEvent(createNativeEvent(37, false, false, false, false)),
		new CTestEvent(createNativeEvent(38, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.moveToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.moveToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.exitFromDemonstrationMode] = [
		new CTestEvent(createNativeEvent(27, false, false, false, false))
	];

	const oThumbnailsTypes = {
		addNextSlide                        : 0,
		removeSelectedSlides                : 1,
		moveSelectedSlidesToEnd             : 2,
		moveSelectedSlidesToNextPosition    : 3,
		selectNextSlide                     : 4,
		moveToNextSlide                     : 5,
		moveToFirstSlide                    : 6,
		selectToFirstSlide                  : 7,
		moveToLastSlide                     : 8,
		selectToLastSlide                   : 9,
		moveSelectedSlidesToStart           : 10,
		moveSelectedSlidesToPreviousPosition: 11,
		selectPreviousSlide                 : 12,
		moveToPreviousSlide                 : 13
	};
	const oThumbnailsEvents = {};
	oThumbnailsEvents[oThumbnailsTypes.addNextSlide] = [
		new CTestEvent(createNativeEvent(13, false, false, false, false)),
		new CTestEvent(createNativeEvent(77, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.removeSelectedSlides] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToEnd] = [
		new CTestEvent(createNativeEvent(40, true, true, false, false)),
		new CTestEvent(createNativeEvent(34, true, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToNextPosition] = [
		new CTestEvent(createNativeEvent(40, true, false, false, false)),
		new CTestEvent(createNativeEvent(34, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectNextSlide] = [
		new CTestEvent(createNativeEvent(40, false, true, false, false)),
		new CTestEvent(createNativeEvent(34, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToNextSlide] = [
		new CTestEvent(createNativeEvent(40, true, false, false, false)),
		new CTestEvent(createNativeEvent(34, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToStart] = [
		new CTestEvent(createNativeEvent(33, true, true, false, false)),
		new CTestEvent(createNativeEvent(38, true, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToPreviousPosition] = [
		new CTestEvent(createNativeEvent(33, true, false, false, false)),
		new CTestEvent(createNativeEvent(38, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectPreviousSlide] = [
		new CTestEvent(createNativeEvent(38, false, true, false, false)),
		new CTestEvent(createNativeEvent(33, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToPreviousSlide] = [
		new CTestEvent(createNativeEvent(33, true, false, false, false)),
		new CTestEvent(createNativeEvent(38, true, false, false, false))
	];

	const oThumbnailsMainFocusTypes = {
		addNextSlide                    : 0,
		moveToPreviousSlide             : 1,
		moveToNextSlide                 : 2,
		moveToFirstSlide                : 3,
		selectToFirstSlide              : 4,
		moveSelectedSlidesToEnd         : 5,
		moveSelectedSlidesToNextPosition: 6,
		moveToLastSlide                 : 7,
		selectToLastSlide               : 8
	};
	const oThumbnailsMainFocusEvents = {};
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.addNextSlide] = [
		new CTestEvent(createNativeEvent(77, true, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToPreviousSlide] = [
		new CTestEvent(createNativeEvent(38, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(37, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(33, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToNextSlide] = [
		new CTestEvent(createNativeEvent(39, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(40, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(34, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.selectToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, true, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.selectToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, true, false, false, false, false))
	];

	function privateStartTest(fCallback, nShortcutType, oTestEvents)
	{
		const arrTestEvents = oTestEvents[nShortcutType];

		for (let i = 0; i < arrTestEvents.length; i += 1)
		{
			const nTestType = arrTestEvents[i].type;
			if (nTestType === testAll)
			{
				AscCommon.AscBrowser.isMacOs = true;
				fCallback(arrTestEvents[i].event);

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

	function startThumbnailsMainFocusTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oThumbnailsMainFocusEvents);
	}

	function startMainTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oMainEvents);
	}

	function startThumbnailsFocusTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oThumbnailsEvents);
	}

	const AscTestShortcut = window.AscTestShortcut = {};
	AscTestShortcut.createMathInShape = createMathInShape;
	AscTestShortcut.moveCursorDown = moveCursorDown;
	AscTestShortcut.moveCursorLeft = moveCursorLeft;
	AscTestShortcut.moveCursorRight = moveCursorRight;
	AscTestShortcut.onInput = onInput;
	AscTestShortcut.goToPageWithFocus = goToPageWithFocus;
	AscTestShortcut.goToPage = goToPage;
	AscTestShortcut.createShape = createShape;
	AscTestShortcut.addSlide = addSlide;
	AscTestShortcut.onKeyDown = onKeyDown;
	AscTestShortcut.executeTestWithCatchEvent = executeTestWithCatchEvent;
	AscTestShortcut.createNativeEvent = createNativeEvent;
	AscTestShortcut.checkSelectedSlides = checkSelectedSlides;
	AscTestShortcut.cleanPresentation = cleanPresentation;
	AscTestShortcut.executeCheckMoveShape = executeCheckMoveShape;
	AscTestShortcut.addPropertyToDocument = addPropertyToDocument;
	AscTestShortcut.testMoveHelper = testMoveHelper;
	AscTestShortcut.createShapeWithTitlePlaceholder = createShapeWithTitlePlaceholder;
	AscTestShortcut.createChart = createChart;
	AscTestShortcut.createTable = createTable;
	AscTestShortcut.checkRemoveObject = checkRemoveObject;
	AscTestShortcut.checkTextAfterKeyDownHelperHelloWorld = checkTextAfterKeyDownHelperHelloWorld;
	AscTestShortcut.checkTextAfterKeyDownHelperEmpty = checkTextAfterKeyDownHelperEmpty;
	AscTestShortcut.getDirectParaPrHelper = getDirectParaPrHelper;
	AscTestShortcut.getDirectTextPrHelper = getDirectTextPrHelper;
	AscTestShortcut.createEvent = createEvent;
	AscTestShortcut.getShapeWithParagraphHelper = getShapeWithParagraphHelper;
	AscTestShortcut.moveToParagraph = moveToParagraph;
	AscTestShortcut.getFirstSlide = getFirstSlide;
	AscTestShortcut.selectOnlyObjects = selectOnlyObjects;
	AscTestShortcut.remove = remove;
	AscTestShortcut.addToSelection = addToSelection;
	AscTestShortcut.createGroup = createGroup;
	AscTestShortcut.getController = getController;
	AscTestShortcut.oGlobalLogicDocument = oGlobalLogicDocument;
	AscTestShortcut.oGlobalShape = oGlobalShape;
	AscTestShortcut.checkDirectTextPrAfterKeyDown = checkDirectTextPrAfterKeyDown;
	AscTestShortcut.checkDirectParaPrAfterKeyDown = checkDirectParaPrAfterKeyDown;
	AscTestShortcut.oMainEvents = oMainEvents;
	AscTestShortcut.oMainShortcutTypes = oMainShortcutTypes;
	AscTestShortcut.oDemonstrationEvents = oDemonstrationEvents;
	AscTestShortcut.oDemonstrationTypes = oDemonstrationTypes;
	AscTestShortcut.oThumbnailsEvents = oThumbnailsEvents;
	AscTestShortcut.oThumbnailsTypes = oThumbnailsTypes;
	AscTestShortcut.oThumbnailsMainFocusTypes = oThumbnailsMainFocusTypes;
	AscTestShortcut.oThumbnailsMainFocusEvents = oThumbnailsMainFocusEvents;
	AscTestShortcut.startThumbnailsFocusTest = startThumbnailsFocusTest;
	AscTestShortcut.startThumbnailsMainFocusTest = startThumbnailsMainFocusTest;
	AscTestShortcut.startMainTest = startMainTest;
})(window)
