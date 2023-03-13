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
	window.setTimeout = function (callback)
	{
		callback();
	}
	AscCommon.CTableId = Object;
	const AscTestShortcut = window.AscTestShortcut = {};

	let editor = new Asc.asc_docs_api({'id-view': 'editor_sdk'});
	window.editor = editor;
	AscCommon.loadSdk = function ()
	{
		editor._onEndLoadSdk();
	}


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
		Init           : function ()
		{

		},
		LoadFont       : function ()
		{

		},
		GetFontInfoName: function () {}
	}

	window.g_fontApplication = AscFonts.g_fontApplication;

	Asc.createPluginsManager = function ()
	{

	};

	AscFonts.FontPickerByCharacter = {
		checkText: function (text, _this, callback)
		{
			callback.call(_this);
		},
		getFontBySymbol: function ()
		{
			
		}
	};

	AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback)
	{
		callback({"saveLock": false});
	};

	let oGlobalLogicDocument;

	function createLogicDocument()
	{
		if (oGlobalLogicDocument)
			return oGlobalLogicDocument;

		editor.InitEditor();
		editor.bInit_word_control = true;
		editor.WordControl.StartMainTimer = function ()
		{

		};
		editor.WordControl.InitControl();

		oGlobalLogicDocument = editor.WordControl.m_oLogicDocument;
		editor.WordControl.m_oDrawingDocument.m_oLogicDocument = oGlobalLogicDocument;
		oGlobalLogicDocument.UpdateAllSectionsInfo();
		oGlobalLogicDocument.Set_DocumentPageSize(100, 200);
		var props = new Asc.CDocumentSectionProps();
		props.put_TopMargin(0);
		props.put_LeftMargin(0);
		props.put_BottomMargin(0);
		props.put_RightMargin(0);
		oGlobalLogicDocument.Set_SectionProps(props);
		oGlobalLogicDocument.private_IsStartTimeoutOnRecalc = function ()
		{
			return false;
		}
		return oGlobalLogicDocument;
	}

	createLogicDocument();
	oGlobalLogicDocument.UpdateAllSectionsInfo();

	AscCommon.g_font_loader.LoadFont = function ()
	{
		return false;
	}
	editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState = function ()
	{
	};
	AscTest.CreateParagraph = function ()
	{
		return new AscWord.CParagraph(editor.WordControl.m_oDrawingDocument);
	}

	function addPropertyToDocument(oPr)
	{
		oGlobalLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(oPr), true);
	}
	function checkTextAfterKeyDownHelperEmpty(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, '');
	}

	function moveCursorDown(AddToSelect, CtrlKey)
	{
		oGlobalLogicDocument.MoveCursorDown(AddToSelect, CtrlKey);
	}
	function moveCursorUp(AddToSelect, CtrlKey)
	{
		oGlobalLogicDocument.MoveCursorUp(AddToSelect, CtrlKey);
	}

	function moveToParagraph(oParagraph, bIsStart, bSkipRemoveSelection)
	{
		if (!bSkipRemoveSelection)
		{
			oGlobalLogicDocument.RemoveSelection();
		}
		oParagraph.SetThisElementCurrent();
		if (bIsStart)
		{
			oParagraph.MoveCursorToStartPos();
		} else
		{
			oParagraph.MoveCursorToEndPos();
		}
		oGlobalLogicDocument.private_UpdateCursorXY(true, true);
	}

	function resetLogicDocument(oLogicDocument)
	{
		oLogicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
	}
	function clean()
	{
		oGlobalLogicDocument.RemoveFromContent(0, oGlobalLogicDocument.GetElementsCount(), false);
	}
	function getLogicDocumentWithParagraphs(arrText, bRecalculate)
	{
		resetLogicDocument(oGlobalLogicDocument);
		if (!oGlobalLogicDocument.TurnOffRecalc)
		{
			oGlobalLogicDocument.Start_SilentMode();
		}
		clean();
		if (Array.isArray(arrText))
		{
			for (let i = 0; i < arrText.length; i += 1)
			{
				addParagraphToDocumentWithText(arrText[i]);
			}
		}
		if (oGlobalLogicDocument.TurnOffRecalc && bRecalculate)
		{
			oGlobalLogicDocument.End_SilentMode(true);
			recalculate();
			oGlobalLogicDocument.private_UpdateCursorXY(true, true);
		}
		recalculate();

		//oGlobalLogicDocument.MoveCursorToEndPos();
		const oFirstParagraph = oGlobalLogicDocument.Content[0];
		return {oLogicDocument: oGlobalLogicDocument, oParagraph: oFirstParagraph};
	}
	function recalculate()
	{
		oGlobalLogicDocument.RecalculateFromStart(false);
	}
	function addParagraphToDocumentWithText(sText)
	{
		const oParagraph = AscTest.CreateParagraph();
		oParagraph.Set_Ind({FirstLine: 0, Left: 0, Right: 0});
		oGlobalLogicDocument.Internal_Content_Add(oGlobalLogicDocument.Content.length, oParagraph);
		oParagraph.MoveCursorToEndPos();
		const oRun = new AscWord.CRun();
		oParagraph.AddToContent(0, oRun);
		oRun.AddText(sText);
		return oParagraph;
	}

	function remove()
	{
		oGlobalLogicDocument.Remove();
	}

	function checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, sInitText)
	{
		const {oLogicDocument, oParagraph} = getLogicDocumentWithParagraphs([sInitText]);
		oParagraph.SetThisElementCurrent();
		oLogicDocument.MoveCursorToEndPos();
		onKeyDown(oEvent);
		const sTextAfterKeyDown = AscTest.GetParagraphText(oParagraph);
		oAssert.strictEqual(sTextAfterKeyDown, sCheckText, sPrompt);
	}

	function onKeyDown(oEvent)
	{
		editor.WordControl.onKeyDown(oEvent);
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

	function checkDirectTextPrAfterKeyDown(fCallback, nExpectedValue, sPrompt, oEvent, oAssert)
	{
		const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
		let oTextPr = getDirectTextPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oTextPr), nExpectedValue, sPrompt);
		return function recursive(fCallback2, nExpectedValue2, sPrompt2, oEvent2, oAssert2)
		{
			oTextPr = getDirectTextPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oTextPr), nExpectedValue2, sPrompt2);
			return recursive;
		}
	}

	function checkDirectParaPrAfterKeyDown(fCallback, nExpectedValue, sPrompt, oEvent, oAssert)
	{
		const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
		let oParaPr = getDirectParaPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oParaPr), nExpectedValue, sPrompt);
		return function recursive(fCallback2, nExpectedValue2, sPrompt2, oEvent2, oAssert2)
		{
			oParaPr = getDirectParaPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oParaPr), nExpectedValue2, sPrompt2);
			return recursive;
		}
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
	function getParagraphText(oParagraph)
	{
		return AscTest.GetParagraphText(oParagraph);
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

	function moveCursorRight(AddToSelect, Word)
	{
		oGlobalLogicDocument.MoveCursorRight(AddToSelect, Word);
	}
	function moveCursorLeft(AddToSelect, Word)
	{
		oGlobalLogicDocument.MoveCursorLeft(AddToSelect, Word);
	}

	function selectAll()
	{
		oGlobalLogicDocument.SelectAll();
	}

	function getSelectedText()
	{
		return oGlobalLogicDocument.GetSelectedText(false, {TabSymbol:'\t'});
	}

	AscTestShortcut.addPropertyToDocument = addPropertyToDocument;
	AscTestShortcut.checkTextAfterKeyDownHelperEmpty = checkTextAfterKeyDownHelperEmpty;
	AscTestShortcut.getLogicDocumentWithParagraphs = getLogicDocumentWithParagraphs;
	AscTestShortcut.resetLogicDocument = resetLogicDocument;
	AscTestShortcut.oGlobalLogicDocument = oGlobalLogicDocument;
	AscTestShortcut.onKeyDown = onKeyDown;
	AscTestShortcut.moveToParagraph = moveToParagraph;
	AscTestShortcut.createNativeEvent = createNativeEvent;
	AscTestShortcut.addParagraphToDocumentWithText = addParagraphToDocumentWithText;

	AscTestShortcut.checkDirectTextPrAfterKeyDown = checkDirectTextPrAfterKeyDown;
	AscTestShortcut.checkDirectParaPrAfterKeyDown = checkDirectParaPrAfterKeyDown;
	AscTestShortcut.executeTestWithCatchEvent = executeTestWithCatchEvent;
	AscTestShortcut.remove = remove;
	AscTestShortcut.recalculate = recalculate;
	AscTestShortcut.clean = clean;
	AscTestShortcut.moveCursorDown = moveCursorDown;
	AscTestShortcut.moveCursorUp = moveCursorUp;
	AscTestShortcut.moveCursorLeft = moveCursorLeft;
	AscTestShortcut.moveCursorRight = moveCursorRight;
	AscTestShortcut.selectAll = selectAll;
	AscTestShortcut.getSelectedText = getSelectedText;
})(window);
