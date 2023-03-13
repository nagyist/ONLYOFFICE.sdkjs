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
		addPropertyToDocument,
		getLogicDocumentWithParagraphs,
		checkTextAfterKeyDownHelperEmpty,
		checkDirectTextPrAfterKeyDown,
		checkDirectParaPrAfterKeyDown,
		oGlobalLogicDocument,
		addParagraphToDocumentWithText,
		remove,
		recalculate,
		clean,
		onKeyDown,
		moveToParagraph,
		createNativeEvent,
		moveCursorDown,
		moveCursorUp,
		moveCursorLeft,
		moveCursorRight,
		selectAll,
		getSelectedText,
		executeTestWithCatchEvent
	} = AscTestShortcut;


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


	function checkInsertElementByType(nType, sPrompt, oAssert, oEvent)
	{
		const {oParagraph} = getLogicDocumentWithParagraphs(['']);
		onKeyDown(oEvent);
		let bCheck = false;
		for (let i = 0; i < oParagraph.Content.length; i += 1)
		{
			const oRun = oParagraph.Content[i];
			for (let j = 0; j < oRun.Content.length; j += 1)
			{
				if (oRun.Content[j].Type === nType)
				{
					bCheck = true;
				}
			}
		}
		oAssert.true(bCheck, sPrompt);
	}

	function createParagraphWithText(sText)
	{
		const oParagraph = AscTest.CreateParagraph();
		const oRun = new AscWord.CRun();
		oParagraph.AddToContent(0, oRun);
		oRun.AddText(sText);
		return oParagraph;
	}

	function checkApplyParagraphStyle(sStyleName, sPrompt, oEvent, oAssert)
	{
		const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
		oLogicDocument.SelectAll();
		onKeyDown(oEvent);
		const oParagraphPr = oLogicDocument.GetDirectParaPr();
		const sPStyleName = oLogicDocument.Styles.Get_Name(oParagraphPr.Get_PStyle());
		oAssert.strictEqual(sPStyleName, sStyleName, sPrompt);
	}


	$(function ()
	{
		let fOldGetShortcut;
		QUnit.module("Test shortcut actions", {
			before    : function ()
			{
				editor.initDefaultShortcuts();
			},
			beforeEach: function ()
			{
				fOldGetShortcut = editor.getShortcut;
			},
			afterEach : function ()
			{
				editor.getShortcut = fOldGetShortcut;
			},
			after     : function ()
			{
				editor.Shortcuts = new AscCommon.CShortcuts();
			}
		});

		QUnit.test('Check page break shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertPageBreak};
			let oEvent = createNativeEvent();
			const {oLogicDocument} = getLogicDocumentWithParagraphs([''], true);
			oLogicDocument.OnKeyDown(oEvent);
			oAssert.strictEqual(oLogicDocument.GetPagesCount(), 2, 'Check page break shortcut');
			oLogicDocument.OnKeyDown(oEvent);
			oAssert.strictEqual(oLogicDocument.GetPagesCount(), 3, 'Check page break shortcut');
			oLogicDocument.OnKeyDown(oEvent);
			oAssert.strictEqual(oLogicDocument.GetPagesCount(), 4, 'Check page break shortcut');
		});

		QUnit.test('Check line break shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertLineBreak};
			const {oLogicDocument, oParagraph} = getLogicDocumentWithParagraphs([''], true);
			oLogicDocument.OnKeyDown(createNativeEvent());
			oAssert.strictEqual(oParagraph.GetLinesCount(), 2, 'Check line break shortcut');
			oLogicDocument.OnKeyDown(createNativeEvent());
			oAssert.strictEqual(oParagraph.GetLinesCount(), 3, 'Check line break shortcut');
			oLogicDocument.OnKeyDown(createNativeEvent());
			oAssert.strictEqual(oParagraph.GetLinesCount(), 4, 'Check line break shortcut');
		});

		QUnit.test('Check column break shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertColumnBreak};
			let oColumnProps = new Asc.CDocumentColumnsProps();
			oColumnProps.put_Num(2);
			editor.asc_SetColumnsProps(oColumnProps);
			const {oLogicDocument} = getLogicDocumentWithParagraphs([''], true);

			oLogicDocument.OnKeyDown(createNativeEvent());
			const oParagraph = oLogicDocument.GetCurrentParagraph();

			oAssert.strictEqual(oParagraph.Get_CurrentColumn(), 1, 'Check column break shortcut');
			oColumnProps = new Asc.CDocumentColumnsProps();
			oColumnProps.put_Num(1);
			editor.asc_SetColumnsProps(oColumnProps);
		});

		QUnit.test('Check reset char shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ResetChar};

			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello world']);
			oLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true, Underline: true});
			onKeyDown(createNativeEvent());
			const oDirectTextPr = oGlobalLogicDocument.GetDirectTextPr();
			oAssert.true(!(oDirectTextPr.Get_Bold() || oDirectTextPr.Get_Italic() || oDirectTextPr.Get_Underline()), 'Check reset char shortcut');
		});

		QUnit.test('Check add non breaking space shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.NonBreakingSpace};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00A0), createNativeEvent(), oAssert, 'Check add non breaking space shortcut');
		});

		QUnit.test('Check add strikeout shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Strikeout};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Strikeout(), true, 'Check add strikeout shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Strikeout(), false, 'Check add strikeout shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check show non printing characters shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ShowAll};
			editor.put_ShowParaMarks(false);
			onKeyDown(createNativeEvent());
			oAssert.true(editor.get_ShowParaMarks(), 'Check show non printing characters shortcut');
		});

		QUnit.test('Check select all shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EditSelectAll};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(oLogicDocument.GetSelectedText(), 'Hello World', 'Check select all shortcut');
		});

		QUnit.test('Check bold shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Bold};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Bold(), true, 'Check bold shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Bold(), false, 'Check bold shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check copy format shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.CopyFormat};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true, Underline: true});

			onKeyDown(createNativeEvent());
			const oCopyParagraphTextPr = new AscCommonWord.CTextPr();
			oCopyParagraphTextPr.SetUnderline(true);
			oCopyParagraphTextPr.SetBold(true);
			oCopyParagraphTextPr.BoldCS = true;
			oCopyParagraphTextPr.SetItalic(true);
			oCopyParagraphTextPr.ItalicCS = true;
			oAssert.deepEqual(editor.getFormatPainterData().TextPr, oCopyParagraphTextPr, 'Check copy format shortcut');
		});

		QUnit.test('Check insert copyright shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.CopyrightSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00A9), createNativeEvent(), oAssert, 'Check add non breaking space shortcut');
		});

		QUnit.test('Check insert endnote shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertEndnoteNow};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello']);
			oLogicDocument.SelectAll();
			oLogicDocument.OnKeyDown(createNativeEvent());
			const arrEndnotes = oLogicDocument.GetEndnotesList();
			oAssert.deepEqual(arrEndnotes.length, 1, 'Check insert endnote shortcut');
		});

		QUnit.test('Check center para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.CenterPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Center, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check insert euro sign shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EuroSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x20AC), createNativeEvent(), oAssert, 'Check add non breaking space shortcut');
		});

		QUnit.test('Check italic shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Italic};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Italic(), true, 'Check add italic shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Italic(), false, 'Check add italic shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check justify para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.JustifyPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Justify, 'Check justify para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check justify para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check bullet list shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyListBullet};
			const {oLogicDocument, oParagraph} = getLogicDocumentWithParagraphs(['Hello']);
			oLogicDocument.SelectAll();
			onKeyDown(createNativeEvent());

			oAssert.true(oParagraph.IsBulletedNumbering(), 'check apply bullet list');
		});

		QUnit.test('Check left para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.LeftPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Justify, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check indent shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Indent};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello']);
			oParagraph.Pr.SetInd(0, 0, 0);
			moveToParagraph(oParagraph, true);

			onKeyDown(createNativeEvent());
			const oParaPr = oGlobalLogicDocument.GetDirectParaPr();
			oAssert.strictEqual(oParaPr.GetIndLeft(), 12.5, 'Check increase indent');
		});

		QUnit.test('Check unindent shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.UnIndent};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello']);
			oParagraph.Pr.SetInd(0, 12.5, 0);
			moveToParagraph(oParagraph, true);

			onKeyDown(createNativeEvent());
			const oParaPr = oGlobalLogicDocument.GetDirectParaPr();
			oAssert.true(AscFormat.fApproxEqual(oParaPr.GetIndLeft(), 0), 'Check decrease indent');
		});

		QUnit.test('Check insert page number shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertPageNumber};
			checkInsertElementByType(para_PageNum, 'Check insert page number shortcut', oAssert, createNativeEvent());
		});

		QUnit.test('Check right para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.RightPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Right, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check registered sign shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.RegisteredSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00AE), createNativeEvent(), oAssert, 'Check registered sign shortcut');
		});

		QUnit.test('Check trademark sign shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.TrademarkSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2122), createNativeEvent(), oAssert, 'Check registered sign shortcut');
		});

		QUnit.test('Check underline shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Underline};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Underline(), true, 'Check underline shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Underline(), false, 'Check underline shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check paste format shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.PasteFormat};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true});
			oGlobalLogicDocument.Document_Format_Copy();
			remove();
			addParagraphToDocumentWithText('Hello');
			oGlobalLogicDocument.SelectAll();
			onKeyDown(createNativeEvent());
			const oDirectTextPr = oGlobalLogicDocument.GetDirectTextPr();
			oAssert.true(oDirectTextPr.Get_Bold() && oDirectTextPr.Get_Italic(), 'Check paste format shortcut');
		});

		QUnit.test('Check redo shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EditRedo};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
			oLogicDocument.SelectAll();
			oLogicDocument.Remove(undefined, undefined, true);
			oLogicDocument.Document_Undo();
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(AscTest.GetParagraphText(oLogicDocument.Content[0]), '', 'Check redo shortcut');
		});

		QUnit.test('Check undo shortcut', (oAssert) =>
		{
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
			selectAll();
			editor.asc_Remove();

			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EditUndo};
			onKeyDown(createNativeEvent());
			selectAll();
			oAssert.strictEqual(getSelectedText(), 'Hello World', 'Check redo shortcut');
		});

		QUnit.test('Check en dash shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EnDash};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2013), createNativeEvent(), oAssert, 'Check en dash shortcut');
		});

		QUnit.test('Check em dash shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EmDash};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2014), createNativeEvent(), oAssert, 'Check em dash shortcut');
		});

		QUnit.test('Check update fields shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.UpdateFields};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello', 'Hello', 'Hello'], true);
			AscTest.Recalculate();
			for (let i = 0; i < oLogicDocument.Content.length; i += 1)
			{
				oLogicDocument.Set_CurrentElement(i, true);
				oLogicDocument.SetParagraphStyle("Heading 1");
			}
			AscTest.Recalculate();
			oLogicDocument.MoveCursorToStartPos();
			const props = new Asc.CTableOfContentsPr();
			props.put_OutlineRange(1, 9);
			props.put_Hyperlink(true);
			props.put_ShowPageNumbers(true);
			props.put_RightAlignTab(true);
			props.put_TabLeader(Asc.c_oAscTabLeader.Dot);
			editor.asc_AddTableOfContents(null, props);

			oLogicDocument.MoveCursorToEndPos();
			const oParagraph = createParagraphWithText('Hello');
			oLogicDocument.AddToContent(oLogicDocument.Content.length, oParagraph);
			oParagraph.MoveCursorToEndPos();
			oParagraph.SetThisElementCurrent();
			oLogicDocument.SetParagraphStyle("Heading 1");

			oLogicDocument.Content[0].SetThisElementCurrent();
			AscTest.Recalculate();
			oLogicDocument.OnKeyDown(createNativeEvent());
			oAssert.strictEqual(oLogicDocument.Content[0].Content.Content.length, 5, 'Check update fields shortcut');
		});

		QUnit.test('Check superscript shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Superscript};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_SuperScript, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_Baseline, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check non breaking hyphen shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.NonBreakingHyphen};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x002D), createNativeEvent(), oAssert, 'Check non breaking hyphen shortcut');
		});

		QUnit.test('Check horizontal ellipsis shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.HorizontalEllipsis};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2026), createNativeEvent(), oAssert, 'Check add horizontal ellipsis shortcut');
		});

		QUnit.test('Check subscript shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Subscript};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_SubScript, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_Baseline, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check show hyperlink menu shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertHyperlink};
			executeTestWithCatchEvent('asc_onDialogAddHyperlink', () => true, true, createNativeEvent(), oAssert, () =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
				moveToParagraph(oParagraph);
				oGlobalLogicDocument.SelectAll();
			});
		});

		QUnit.test('Check print shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.PrintPreviewAndPrint};
			executeTestWithCatchEvent('asc_onPrint', () => true, true, createNativeEvent(), oAssert);
		});

		QUnit.test('Check save shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Save};
			const fOldSave = editor._onSaveCallbackInner;
			let bCheck = false;
			editor._onSaveCallbackInner = function ()
			{
				bCheck = true;
				editor.canSave = true;
			};
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(bCheck, true, 'Check save shortcut');
			editor._onSaveCallbackInner = fOldSave;
		});

		QUnit.test('Check increase font size shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.IncreaseFontSize};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), 11, 'Check increase font size shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_FontSize(), 12, 'Check increase font size shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check decrease font size shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.DecreaseFontSize};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), 9, 'Check decrease font size shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_FontSize(), 8, 'Check decrease font size shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check apply heading 1', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyHeading1};
			checkApplyParagraphStyle('Heading 1', 'Check apply heading 1 shortcut', createNativeEvent(), oAssert);
		});
		QUnit.test('Check apply heading 2', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyHeading2};
			checkApplyParagraphStyle('Heading 2', 'Check apply heading 2 shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check apply heading 3', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyHeading3};
			checkApplyParagraphStyle('Heading 3', 'Check apply heading 3 shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check insert footnotes now', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertFootnoteNow};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello']);
			oLogicDocument.SelectAll();
			oLogicDocument.OnKeyDown(createNativeEvent());
			const arrFootnotes = oLogicDocument.GetFootnotesList();
			oAssert.deepEqual(arrFootnotes.length, 1, 'Check insert footnote shortcut');
		});

		QUnit.test('Check insert equation', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertEquation};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['']);
			onKeyDown(createNativeEvent());
			const oMath = oLogicDocument.GetCurrentMath();
			oAssert.true(!!oMath, 'Check insert equation shortcut');
		});

		function createHyperlink()
		{
			const oProps = new Asc.CHyperlinkProperty({Anchor: '_top', Text: "Beginning of document"});
			editor.add_Hyperlink(oProps);
		}

		QUnit.module("Test getting desired action by event")
		QUnit.test("Test getting common desired action by event", (oAssert) =>
		{
			editor.initDefaultShortcuts();
			oAssert.strictEqual(editor.getShortcut(createEvent(13, true, false, false, false, false, false)), c_oAscDocumentShortcutType.InsertPageBreak, 'Check getting c_oAscDocumentShortcutType.InsertPageBreak action');
			oAssert.strictEqual(editor.getShortcut(createEvent(13, false, true, false, false, false, false)), c_oAscDocumentShortcutType.InsertLineBreak, 'Check getting c_oAscDocumentShortcutType.InsertLineBreak action');
			oAssert.strictEqual(editor.getShortcut(createEvent(13, true, true, false, false, false, false)), c_oAscDocumentShortcutType.InsertColumnBreak, 'Check getting c_oAscDocumentShortcutType.InsertColumnBreak action');
			oAssert.strictEqual(editor.getShortcut(createEvent(32, true, false, false, false, false, false)), c_oAscDocumentShortcutType.ResetChar, 'Check getting c_oAscDocumentShortcutType.ResetChar action');
			oAssert.strictEqual(editor.getShortcut(createEvent(32, true, true, false, false, false, false)), c_oAscDocumentShortcutType.NonBreakingSpace, 'Check getting c_oAscDocumentShortcutType.NonBreakingSpace action');
			oAssert.strictEqual(editor.getShortcut(createEvent(53, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Strikeout, 'Check getting c_oAscDocumentShortcutType.Strikeout action');
			oAssert.strictEqual(editor.getShortcut(createEvent(56, true, true, false, false, false, false)), c_oAscDocumentShortcutType.ShowAll, 'Check getting c_oAscDocumentShortcutType.ShowAll action');
			oAssert.strictEqual(editor.getShortcut(createEvent(65, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EditSelectAll, 'Check getting c_oAscDocumentShortcutType.EditSelectAll action');
			oAssert.strictEqual(editor.getShortcut(createEvent(66, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Bold, 'Check getting c_oAscDocumentShortcutType.Bold action');
			oAssert.strictEqual(editor.getShortcut(createEvent(67, true, true, false, false, false, false)), c_oAscDocumentShortcutType.CopyFormat, 'Check getting c_oAscDocumentShortcutType.CopyFormat action');
			oAssert.strictEqual(editor.getShortcut(createEvent(67, true, false, true, false, false, false)), c_oAscDocumentShortcutType.CopyrightSign, 'Check getting c_oAscDocumentShortcutType.CopyrightSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(68, true, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertEndnoteNow, 'Check getting c_oAscDocumentShortcutType.InsertEndnoteNow action');
			oAssert.strictEqual(editor.getShortcut(createEvent(69, true, false, false, false, false, false)), c_oAscDocumentShortcutType.CenterPara, 'Check getting c_oAscDocumentShortcutType.CenterPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(69, true, false, true, false, false, false)), c_oAscDocumentShortcutType.EuroSign, 'Check getting c_oAscDocumentShortcutType.EuroSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(73, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Italic, 'Check getting c_oAscDocumentShortcutType.Italic action');
			oAssert.strictEqual(editor.getShortcut(createEvent(74, true, false, false, false, false, false)), c_oAscDocumentShortcutType.JustifyPara, 'Check getting c_oAscDocumentShortcutType.JustifyPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(75, true, false, false, false, false, false)), c_oAscDocumentShortcutType.InsertHyperlink, 'Check getting c_oAscDocumentShortcutType.InsertHyperlink action');
			oAssert.strictEqual(editor.getShortcut(createEvent(76, true, true, false, false, false, false)), c_oAscDocumentShortcutType.ApplyListBullet, 'Check getting c_oAscDocumentShortcutType.ApplyListBullet action');
			oAssert.strictEqual(editor.getShortcut(createEvent(76, true, false, false, false, false, false)), c_oAscDocumentShortcutType.LeftPara, 'Check getting c_oAscDocumentShortcutType.LeftPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(77, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Indent, 'Check getting c_oAscDocumentShortcutType.Indent action');
			oAssert.strictEqual(editor.getShortcut(createEvent(77, true, true, false, false, false, false)), c_oAscDocumentShortcutType.UnIndent, 'Check getting c_oAscDocumentShortcutType.UnIndent action');
			oAssert.strictEqual(editor.getShortcut(createEvent(80, true, false, false, false, false, false)), c_oAscDocumentShortcutType.PrintPreviewAndPrint, 'Check getting c_oAscDocumentShortcutType.PrintPreviewAndPrint action');
			oAssert.strictEqual(editor.getShortcut(createEvent(80, true, true, false, false, false, false)), c_oAscDocumentShortcutType.InsertPageNumber, 'Check getting c_oAscDocumentShortcutType.InsertPageNumber action');
			oAssert.strictEqual(editor.getShortcut(createEvent(82, true, false, false, false, false, false)), c_oAscDocumentShortcutType.RightPara, 'Check getting c_oAscDocumentShortcutType.RightPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(82, true, false, true, false, false, false)), c_oAscDocumentShortcutType.RegisteredSign, 'Check getting c_oAscDocumentShortcutType.RegisteredSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(83, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Save, 'Check getting c_oAscDocumentShortcutType.Save action');
			oAssert.strictEqual(editor.getShortcut(createEvent(84, true, false, true, false, false, false)), c_oAscDocumentShortcutType.TrademarkSign, 'Check getting c_oAscDocumentShortcutType.TrademarkSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(85, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Underline, 'Check getting c_oAscDocumentShortcutType.Underline action');
			oAssert.strictEqual(editor.getShortcut(createEvent(86, true, true, false, false, false, false)), c_oAscDocumentShortcutType.PasteFormat, 'Check getting c_oAscDocumentShortcutType.PasteFormat action');
			oAssert.strictEqual(editor.getShortcut(createEvent(89, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EditRedo, 'Check getting c_oAscDocumentShortcutType.EditRedo action');
			oAssert.strictEqual(editor.getShortcut(createEvent(90, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EditUndo, 'Check getting c_oAscDocumentShortcutType.EditUndo action');
			oAssert.strictEqual(editor.getShortcut(createEvent(109, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EnDash, 'Check getting c_oAscDocumentShortcutType.EnDash action');
			oAssert.strictEqual(editor.getShortcut(createEvent(109, true, false, true, false, false, false)), c_oAscDocumentShortcutType.EmDash, 'Check getting c_oAscDocumentShortcutType.EmDash action');
			oAssert.strictEqual(editor.getShortcut(createEvent(120, false, false, false, false, false, false)), c_oAscDocumentShortcutType.UpdateFields, 'Check getting c_oAscDocumentShortcutType.UpdateFields action');
			oAssert.strictEqual(editor.getShortcut(createEvent(188, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Superscript, 'Check getting c_oAscDocumentShortcutType.Superscript action');
			oAssert.strictEqual(editor.getShortcut(createEvent(189, true, true, false, false, false, false)), c_oAscDocumentShortcutType.NonBreakingHyphen, 'Check getting c_oAscDocumentShortcutType.NonBreakingHyphen action');
			oAssert.strictEqual(editor.getShortcut(createEvent(190, true, false, true, false, false, false)), c_oAscDocumentShortcutType.HorizontalEllipsis, 'Check getting c_oAscDocumentShortcutType.HorizontalEllipsis action');
			oAssert.strictEqual(editor.getShortcut(createEvent(190, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Subscript, 'Check getting c_oAscDocumentShortcutType.Subscript action');
			oAssert.strictEqual(editor.getShortcut(createEvent(219, true, false, false, false, false, false)), c_oAscDocumentShortcutType.DecreaseFontSize, 'Check getting c_oAscDocumentShortcutType.DecreaseFontSize action');
			oAssert.strictEqual(editor.getShortcut(createEvent(221, true, false, false, false, false, false)), c_oAscDocumentShortcutType.IncreaseFontSize, 'Check getting c_oAscDocumentShortcutType.IncreaseFontSize action');
			editor.Shortcuts = new AscCommon.CShortcuts();
		});

		QUnit.test("Test getting windows desired action by event", (oAssert) =>
		{
			editor.initDefaultShortcuts();
			oAssert.strictEqual(editor.getShortcut(createEvent(49, false, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading1, 'Check getting c_oAscDocumentShortcutType.ApplyHeading1 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(50, false, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading2, 'Check getting c_oAscDocumentShortcutType.ApplyHeading2 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(51, false, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading3, 'Check getting c_oAscDocumentShortcutType.ApplyHeading3 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(70, true, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertFootnoteNow, 'Check getting c_oAscDocumentShortcutType.InsertFootnoteNow shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(187, false, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertEquation, 'Check getting c_oAscDocumentShortcutType.InsertEquation shortcut type');
			editor.Shortcuts = new AscCommon.CShortcuts();
		});

		QUnit.test("Test getting macOs desired action by event", (oAssert) =>
		{
			const bOldMacOs = AscCommon.AscBrowser.isMacOs;
			AscCommon.AscBrowser.isMacOs = true;
			editor.initDefaultShortcuts();
			oAssert.strictEqual(editor.getShortcut(createEvent(49, true, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading1, 'Check getting c_oAscDocumentShortcutType.ApplyHeading1 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(50, true, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading2, 'Check getting c_oAscDocumentShortcutType.ApplyHeading2 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(51, true, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading3, 'Check getting c_oAscDocumentShortcutType.ApplyHeading3 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(187, true, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertEquation, 'Check getting c_oAscDocumentShortcutType.InsertEquation shortcut type');
			editor.Shortcuts = new AscCommon.CShortcuts();
			AscCommon.AscBrowser.isMacOs = bOldMacOs;
		});

		function createTable(nRows, nCols)
		{
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['']);
			return oLogicDocument.AddInlineTable(nCols, nRows);
		}

		function moveToTable(oTable, bToStart)
		{
			oTable.Document_SetThisElementCurrent();
			if (bToStart)
			{
				oTable.MoveCursorToStartPos();
			} else
			{
				oTable.MoveCursorToEndPos();
			}
		}

		function createShape()
		{
			AscCommon.History.Create_NewPoint();
			const oDrawing = new ParaDrawing(200, 100, null, oGlobalLogicDocument.GetDrawingDocument(), oGlobalLogicDocument, null);
			const oShapeTrack = new AscFormat.NewShapeTrack('rect', 0, 0, oGlobalLogicDocument.theme, null, null, null, 0);
			oShapeTrack.track({}, 0, 0);
			const oShape = oShapeTrack.getShape(true, oGlobalLogicDocument.GetDrawingDocument(), null);
			oShape.setBDeleted(false);
			oShape.setParent(oDrawing);
			oDrawing.Set_GraphicObject(oShape);
			oDrawing.Set_DrawingType(drawing_Anchor);
			oDrawing.Set_WrappingType(WRAPPING_TYPE_NONE);
			oDrawing.Set_Distance(0, 0, 0, 0);
			const oNearestPos = oGlobalLogicDocument.Get_NearestPos(0, oShape.x, oShape.y, true, oDrawing);
			oDrawing.Set_XYForAdd(oShape.x, oShape.y, oNearestPos, 0);
			oDrawing.AddToDocument(oNearestPos);
			oDrawing.CheckWH();
			recalculate();
			return oDrawing;
		}

		function getShapeWithText(sText, bStartRecalculate)
		{
			const oParaDrawing = createShape();
			selectParaDrawing(oParaDrawing);
			const oShape = oParaDrawing.GraphicObj;
			oShape.createTextBoxContent();
			const oParagraph = oShape.getDocContent().Content[0];
			moveToParagraph(oParagraph);
			addText(sText)
			if (bStartRecalculate)
			{
				startRecalculate();
			}
			return {oParagraph};
		}

		function selectParaDrawing(oParaDrawing)
		{
			oGlobalLogicDocument.SelectDrawings([oParaDrawing], oGlobalLogicDocument);
		}

		QUnit.test("Test remove back", (oAssert) =>
		{
			let oEvent;
			oEvent = createNativeEvent(8, false, false, false, false);
			let {oParagraph} = getLogicDocumentWithParagraphs(['Hello World'], true);
			moveToParagraph(oParagraph);
			onKeyDown(oEvent);
			selectAll();
			oAssert.strictEqual(getSelectedText(), 'Hello Worl', 'Test remove back symbol');

			oEvent = createNativeEvent(8, true, false, false, false);
			moveToParagraph(oParagraph);
			onKeyDown(oEvent);
			selectAll();
			oAssert.strictEqual(getSelectedText(), 'Hello ', 'Test remove back word');

			// todo
			moveToParagraph(oParagraph);
			const oDrawing = createShape();
			selectParaDrawing(oDrawing);
			oEvent = createNativeEvent(8, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oGlobalLogicDocument.Content[0].GetRunByElement(oDrawing), null, 'Test remove shape');

			const oInlineLvlSdt = createComboBox();
			oEvent = createNativeEvent(8, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oParagraph.GetRunByElement(oInlineLvlSdt), null, 'Test remove form');
		});

		QUnit.test("Test move to next form", (oAssert) =>
		{
			getLogicDocumentWithParagraphs(['']);
			const oInlineSdt1 = createComboBox();
			moveCursorRight();
			const oInlineSdt2 = createComboBox();
			moveCursorRight();
			const oInlineSdt3 = createComboBox();
			setFillingFormsMode(true);

			onKeyDown(createNativeEvent(9, false, false, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt1, 'Test move to next form');

			onKeyDown(createNativeEvent(9, false, false, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt2, 'Test move to next form');

			onKeyDown(createNativeEvent(9, false, false, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt3, 'Test move to next form');

			setFillingFormsMode(false);

		});

		QUnit.test("Test move to previous form", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(['']);
			const oInlineSdt1 = createComboBox();
			moveCursorRight();
			const oInlineSdt2 = createComboBox();
			moveCursorRight();
			const oInlineSdt3 = createComboBox();
			setFillingFormsMode(true);

			onKeyDown(createNativeEvent(9, false, true, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt2, 'Test move to next form');

			onKeyDown(createNativeEvent(9, false, true, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt1, 'Test move to next form');

			onKeyDown(createNativeEvent(9, false, true, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt3, 'Test move to next form');

			setFillingFormsMode(false);

		});

		QUnit.test("Test handle tab in math", (oAssert) =>
		{
			oAssert.strictEqual(true, false, 'Test move to next form');
		});

		function createChart()
		{
			var oDrawingDocument = editor.WordControl.m_oDrawingDocument;

			var oDrawing = new ParaDrawing(100, 100, null, oDrawingDocument, null, null);
			const oChartSpace = AscCommon.getChartByType(Asc.c_oAscChartTypeSettings.lineNormal);
			oChartSpace.spPr.setXfrm(new AscFormat.CXfrm());
			oChartSpace.spPr.xfrm.setOffX(0);
			oChartSpace.spPr.xfrm.setOffY(0);
			oChartSpace.spPr.xfrm.setExtX(100);
			oChartSpace.spPr.xfrm.setExtY(100);

			oChartSpace.setParent(oDrawing);
			oDrawing.Set_GraphicObject(oChartSpace);
			oDrawing.setExtent(oChartSpace.spPr.xfrm.extX, oChartSpace.spPr.xfrm.extY);

			oDrawing.Set_DrawingType(drawing_Anchor);
			oDrawing.Set_WrappingType(WRAPPING_TYPE_NONE);
			oDrawing.Set_Distance(0, 0, 0, 0);
			const oNearestPos = oGlobalLogicDocument.Get_NearestPos(0, oChartSpace.x, oChartSpace.y, true, oDrawing);
			oDrawing.Set_XYForAdd(oChartSpace.x, oChartSpace.y, oNearestPos, 0);
			oDrawing.AddToDocument(oNearestPos);
			oDrawing.CheckWH();
			recalculate();

			return oDrawing;
		}

		QUnit.test("Test move to cell", (oAssert) =>
		{
			let oEvent;
			oEvent = createNativeEvent(9, false, false, false, false);
			const oTable = createTable(3, 3);
			moveToTable(oTable, true);
			onKeyDown(oEvent);

			oAssert.strictEqual(oTable.CurCell.Index, 1, 'Test move to next cell');

			oEvent = createNativeEvent(9, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oTable.CurCell.Index, 0, 'Test move to previous cell');
		});

		function drawingObjects()
		{
			return oGlobalLogicDocument.DrawingObjects;
		}

		QUnit.test("Test select object", (oAssert) =>
		{
			clean();
			getLogicDocumentWithParagraphs([''], true);
			recalculate();
			let oEvent;
			oEvent = createNativeEvent(9, false, false, false, false);
			const oFirstParaDrawing = createShape();
			const oSecondParaDrawing = createShape();
			selectParaDrawing(oFirstParaDrawing);
			onKeyDown(oEvent);
			oAssert.strictEqual(drawingObjects().selectedObjects.length === 1 && drawingObjects().selectedObjects[0] === oSecondParaDrawing.GraphicObj, true, 'Test select next object');
			oEvent = createNativeEvent(9, false, true, false, false);
			onKeyDown(oEvent);

			oAssert.strictEqual(drawingObjects().selectedObjects.length === 1 && drawingObjects().selectedObjects[0] === oFirstParaDrawing.GraphicObj, true, 'Test select previous object');
		});

		function logicContent()
		{
			return oGlobalLogicDocument.Content;
		}

		function directParaPr()
		{
			return oGlobalLogicDocument.GetDirectParaPr();
		}

		function directTextPr()
		{
			return oGlobalLogicDocument.GetDirectTextPr();
		}

		QUnit.test("Test working with indent", (oAssert) =>
		{
			let oEvent;
			getLogicDocumentWithParagraphs(['Hello world', "Hello world"]);
			selectAll();
			oEvent = createNativeEvent(9, false, false, false, false);
			onKeyDown(oEvent);
			let arrSteps = [];
			moveToParagraph(logicContent()[0]);
			arrSteps.push(directParaPr().GetIndLeft());
			moveToParagraph(logicContent()[1]);
			arrSteps.push(directParaPr().GetIndLeft());
			oAssert.deepEqual(arrSteps, [12.5, 12.5], 'Test indent');

			selectAll();
			oEvent = createNativeEvent(9, false, true, false, false);
			onKeyDown(oEvent);

			arrSteps = [];
			moveToParagraph(logicContent()[0]);
			arrSteps.push(directParaPr().GetIndLeft());
			moveToParagraph(logicContent()[1]);
			arrSteps.push(directParaPr().GetIndLeft());

			oAssert.deepEqual(arrSteps, [0, 0], 'Test unindent');
		});

		QUnit.test("Test add tab to paragraph", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
			moveToParagraph(oParagraph, true);
			moveCursorRight();
			onKeyDown(createNativeEvent(9, false, false, false));
			selectAll();

			oAssert.strictEqual(getSelectedText(), 'H\tello World', 'Test indent');
		});

		function addToParagraph(oElement)
		{
			oGlobalLogicDocument.AddToParagraph(oElement);
		}

		function addBreakPage()
		{
			addToParagraph(new AscWord.CRunBreak(AscWord.break_Page));

		}

		QUnit.test("Test visit hyperlink", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(['']);
			addBreakPage();
			createHyperlink();
			moveCursorLeft();
			moveCursorLeft();
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(oGlobalLogicDocument.GetCurrentParagraph(), oGlobalLogicDocument.Content[0]);
			//oAssert.strictEqual(contentPosition(), 0);
			oAssert.strictEqual(oGlobalLogicDocument.Get_CurPage(), 0);
		});

		// QUnit.test("Test go to bookmark", (oAssert) =>
		// {
		// 	oAssert.strictEqual(true, false, 'Test indent');
		// 	oAssert.strictEqual(true, false, 'Test unindent');
		// 	oAssert.strictEqual(true, false, 'Test add tab');
		// });

		function createInlineSdt()
		{
			return editor.asc_AddContentControl(c_oAscSdtLevelType.Inline).CC;
		}

		QUnit.test("Test add break line to inlinelvlsdt", (oAssert) =>
		{
			getLogicDocumentWithParagraphs([''], true);
			const oInlineSdt = createComplexForm();
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(oInlineSdt.Lines[0], 2);
		});

		QUnit.test("Test create textBoxContent", (oAssert) =>
		{
			startRecalculate();
			const oParaDrawing = createShape();
			selectParaDrawing(oParaDrawing);
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(!!oParaDrawing.GraphicObj.textBoxContent, true);
		});

		QUnit.test("Test create txBody", (oAssert) =>
		{
			startRecalculate();
			const oParaDrawing = createShape();
			oParaDrawing.GraphicObj.setWordShape(false);
			selectParaDrawing(oParaDrawing);
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(!!oParaDrawing.GraphicObj.txBody, true);
		});

		QUnit.test("Test add new line to math", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(['']);
			createMath(c_oAscMathType.FractionVertical);
			moveCursorLeft();
			moveCursorLeft();
			addText('Hello');
			moveCursorLeft();
			moveCursorLeft();
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			const oParaMath = oParagraph.GetAllParaMaths()[0];
			const oFraction = oParaMath.Root.GetFirstElement();
			const oNumerator = oFraction.getNumerator();
			const oEqArray = oNumerator.GetFirstElement();
			oAssert.strictEqual(oEqArray.getRowsCount(), 2, 'Check add new line math');

		});

		QUnit.test("Test move cursor to start position shape", (oAssert) =>
		{
			const oParaDrawing = createShape();
			const oShape = oParaDrawing.GraphicObj;
			oShape.createTextBoxContent();
			selectParaDrawing(oParaDrawing);
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(oShape.getDocContent().IsCursorAtBegin(), true);
		});



		QUnit.test("Test select all in shape", (oAssert) =>
		{
			getLogicDocumentWithParagraphs([''], true)
			const oParaDrawing = createShape();
			const oShape = oParaDrawing.GraphicObj;
			oShape.createTextBoxContent();
			moveToParagraph(oShape.getDocContent().Content[0]);
			addText('Hello');
			selectParaDrawing(oParaDrawing);
			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(getSelectedText(), 'Hello');
		});
function startRecalculate()
{
	if (oGlobalLogicDocument.TurnOffRecalc)
	{
		oGlobalLogicDocument.End_SilentMode(true);
		recalculate();
		oGlobalLogicDocument.private_UpdateCursorXY(true, true);
	}
}
		QUnit.test("Test move cursor to start position chart title", (oAssert) =>
		{
			startRecalculate();
			const oParaDrawing = createChart();
			const oChart = oParaDrawing.GraphicObj;
			const oTitles = oChart.getAllTitles();
			const oContent = AscFormat.CreateDocContentFromString('', drawingObjects().getDrawingDocument(), oTitles[0].txBody);
			oTitles[0].txBody.content = oContent;
			selectParaDrawing(oParaDrawing);

			const oController = drawingObjects();
			oController.selection.chartSelection = oChart;
			oChart.selectTitle(oTitles[0], 0);

			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.true(oContent.IsCursorAtBegin(), 'Check move cursor to begin pos in title');

		});

		QUnit.test("Test select all in chart title", (oAssert) =>
		{
			getLogicDocumentWithParagraphs([''], true);
			const oParaDrawing = createChart();
			const oChart = oParaDrawing.GraphicObj;
			selectParaDrawing(oParaDrawing);
			const oTitles = oChart.getAllTitles();
			const oController = drawingObjects();
			oController.selection.chartSelection = oChart;
			oChart.selectTitle(oTitles[0], 0);

			onKeyDown(createNativeEvent(13, false, false, false, false, false));
			oAssert.strictEqual(getSelectedText(), 'Diagram Title', 'Check select all title');
		});

		function createMath(nType)
		{
			return editor.asc_AddMath(nType);
		}

		function addText(sText)
		{
			oGlobalLogicDocument.AddTextWithPr(sText);
		}

		QUnit.test("Test add new paragraph", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello Text']);
			moveToParagraph(oParagraph);
			let oEvent = createNativeEvent(13, false, false, false, false);

			onKeyDown(oEvent);

			oAssert.strictEqual(logicContent().length, 2, 'Test add new paragraph to content');
			createMath();
			addText('abcd');
			moveCursorLeft();
			oEvent = createNativeEvent(13, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(logicContent().length, 3, 'Test add new paragraph with math');


		});

		QUnit.test("Test close all window popups", (oAssert) =>
		{
			const oEvent = createNativeEvent(27, false, false, false, false, false)
			executeTestWithCatchEvent('asc_onMouseMoveStart', () => true, true, oEvent, oAssert);
			executeTestWithCatchEvent('asc_onMouseMove', () => true, true, oEvent, oAssert);
			executeTestWithCatchEvent('asc_onMouseMoveEnd', () => true, true, oEvent, oAssert);
		});

		QUnit.test("Test reset", (oAssert) =>
		{

			oAssert.strictEqual(true, false, "Test reset drag'n'drop");
			oAssert.strictEqual(true, false, "Test reset marker");
			oAssert.strictEqual(true, false, "Test reset formatting by example");
			oAssert.strictEqual(true, false, "Test reset shape selection");
			oAssert.strictEqual(true, false, "Test reset add shape");

		});


		QUnit.test("Test end editing", (oAssert) =>
		{
			oAssert.strictEqual(true, false, "Test end editing footer");
			oAssert.strictEqual(true, false, "Test end editing drawing");
			oAssert.strictEqual(true, false, "Test end editing form");
		});

		QUnit.test("Test toggle checkbox", (oAssert) =>
		{
			const oInlineSdt = createCheckBox();
			setFillingFormsMode(true);
			onKeyDown(createNativeEvent(32, false, false, false, false, false));
			oAssert.strictEqual(oInlineSdt.IsCheckBoxChecked(), true);
			setFillingFormsMode(false);
		});


		QUnit.test("Test actions to page up", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World ']);
			console.log(oGlobalLogicDocument.Pages.length)
			moveToParagraph(oParagraph);
			onKeyDown(33)
			oAssert.strictEqual(true, false, "Test move to begin of previous page");
			oAssert.strictEqual(true, false, "Test move to previous page");
			oAssert.strictEqual(true, false, "Test select to previous page");
			oAssert.strictEqual(true, false, "Test select to begin of previous page");
			oAssert.strictEqual(true, false, "Test move to previous header/footer");

		});

		QUnit.test("Test actions to page down", (oAssert) =>
		{
			oAssert.strictEqual(true, false, "Test move to next header/footer");
			oAssert.strictEqual(true, false, "Test move to begin of next page");
			oAssert.strictEqual(true, false, "Test move to next page");
			oAssert.strictEqual(true, false, "Test select to next page");
			oAssert.strictEqual(true, false, "Test select to begin of next page");
		});

		QUnit.test("Test actions to end", (oAssert) =>
		{
			let oEvent;
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
			moveToParagraph(oParagraph, true);

			recalculate();
			oEvent = createNativeEvent(35, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 107, "Test move to end of document");
			moveToParagraph(oParagraph, true);
			oEvent = createNativeEvent(35, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 18, "Test move to end of line");
			moveToParagraph(oParagraph, true);
			oEvent = createNativeEvent(35, true, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), "Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World", "Test select to end of document");
			moveToParagraph(oParagraph, true);
			oEvent = createNativeEvent(35, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), "Hello World Hello ", "Test select to end of line");
		});

		QUnit.test("Test actions to home", (oAssert) =>
		{
			let oEvent;
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
			moveToParagraph(oParagraph);

			recalculate();
			oEvent = createNativeEvent(36, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 0, "Test move to home of document");
			moveToParagraph(oParagraph);
			oEvent = createNativeEvent(36, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 90, "Test move to home of line");
			moveToParagraph(oParagraph);
			oEvent = createNativeEvent(36, true, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), "Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World", "Test select to home of document");
			moveToParagraph(oParagraph);
			oEvent = createNativeEvent(36, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), "World Hello World", "Test select to home of line");
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			let oEvent;
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World"], true);

			moveToParagraph(oParagraph);

			oEvent = createNativeEvent(37, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 22, "Test move to previous symbol");

			oEvent = createNativeEvent(37, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), 'l', "Test select to previous symbol");

			oEvent = createNativeEvent(37, true, false, false, false);
			moveCursorLeft();
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 18, "Test move to previous word");
			oEvent = createNativeEvent(37, true, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), 'Hello ', "Test select to previous word");
		});


		QUnit.test("Test actions to right", (oAssert) =>
		{
			let oEvent
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);

			moveToParagraph(oParagraph, true);

			oEvent = createNativeEvent(39, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(contentPosition(), 1, "Test move to next symbol");

			oEvent = createNativeEvent(39, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), "e", "Test select to next symbol");

			moveCursorRight();
			oEvent = createNativeEvent(39, true, false, false, false);
			onKeyDown(oEvent);
			oAssert.deepEqual(contentPosition(), 6, "Test move to next word");

			oEvent = createNativeEvent(39, true, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), 'World ', "Test select to next word");
		});

		// QUnit.test("Test actions to right in shape", (oAssert) =>
		// {
		// 	let oEvent
		// 	const {oParagraph} = getShapeWithText("Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World", true);
		// 	moveToParagraph(oParagraph, true, true);
		//
		// 	oEvent = createNativeEvent(39, false, false, false, false);
		// 	onKeyDown(oEvent);
		// 	oAssert.strictEqual(contentPosition(), 1, "Test move to next symbol");
		//
		// 	oEvent = createNativeEvent(39, false, true, false, false);
		// 	onKeyDown(oEvent);
		// 	oAssert.strictEqual(getSelectedText(), "e", "Test select to next symbol");
		//
		// 	moveCursorRight();
		// 	oEvent = createNativeEvent(39, true, false, false, false);
		// 	onKeyDown(oEvent);
		// 	oAssert.deepEqual(contentPosition(), 6, "Test move to next word");
		//
		// 	oEvent = createNativeEvent(39, true, true, false, false);
		// 	onKeyDown(oEvent);
		// 	oAssert.strictEqual(getSelectedText(), 'World ', "Test select to next word");
		// });

		QUnit.test("Test actions to up", (oAssert) =>
		{
			let oEvent;
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);

			moveToParagraph(oParagraph, true);

			moveCursorDown();
			oEvent = createNativeEvent(38, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.deepEqual(contentPosition(), 0, "Test move to upper line");

			moveCursorDown();
			oEvent = createNativeEvent(38, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), 'Hello World Hello ', "Test select to upper line");

			clean();
			getLogicDocumentWithParagraphs(['']);
			createComboBox();
			setFillingFormsMode(true);
			oEvent = createNativeEvent(38, false, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(AscTest.GetParagraphText(oGlobalLogicDocument.Content[0]), 'Hello', "Test select next option in combo box");
			onKeyDown(oEvent);
			oAssert.strictEqual(AscTest.GetParagraphText(oGlobalLogicDocument.Content[0]), 'World', "Test select next option in combo box");
			setFillingFormsMode(false);
		});

		function setFillingFormsMode(bState)
		{
			var oRole = new AscCommon.CRestrictionSettings();
			oRole.put_OFormRole("Anyone");
			editor.asc_setRestriction(bState ? Asc.c_oAscRestrictionType.OnlyForms : Asc.c_oAscRestrictionType.None, oRole);
			editor.asc_SetPerformContentControlActionByClick(bState);
			editor.asc_SetHighlightRequiredFields(bState);
		}

		let nKeyId = 0;

		function createCheckBox()
		{
			const oCheckBox = oGlobalLogicDocument.AddContentControlCheckBox();
			var props = new AscCommon.CContentControlPr();
			var specProps = new AscCommon.CSdtCheckBoxPr();
			var oFormProps = new AscCommon.CSdtFormPr('key' + nKeyId++, '', '', false);
			props.SetFormPr(oFormProps);
			props.put_CheckBoxPr(specProps);
			editor.asc_SetContentControlProperties(props, oCheckBox.GetId());
			return oCheckBox;
		}
		function createComboBox()
		{
			const oComboBox = oGlobalLogicDocument.AddContentControlComboBox();
			var props = new AscCommon.CContentControlPr();
			var specProps = new AscCommon.CSdtComboBoxPr();
			var oFormProps = new AscCommon.CSdtFormPr('key' + nKeyId++, '', '', false);
			props.SetFormPr(oFormProps);
			specProps.clear();
			specProps.add_Item('Hello', 'Hello');
			specProps.add_Item('World', 'World');
			props.put_ComboBoxPr(specProps);
			editor.asc_SetContentControlProperties(props, oComboBox.GetId());
			return oComboBox;
		}
		function createComplexForm()
		{
			const oComplexForm = oGlobalLogicDocument.AddComplexForm();
			var props   = new AscCommon.CContentControlPr();
			var formTextPr = new AscCommon.CSdtTextFormPr();
			formTextPr.put_MultiLine(true);
			props.put_TextFormPr(formTextPr);
			editor.asc_SetContentControlProperties(props, oComplexForm.GetId());
			return oComplexForm;
		}

		function contentPosition()
		{
			const oPos = oGlobalLogicDocument.GetContentPosition();
			return oPos[oPos.length - 1].Position;
		}

		QUnit.test("Test actions to down", (oAssert) =>
		{
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
			moveToParagraph(oParagraph, true);

			recalculate();
			let oEvent;
			oEvent = createNativeEvent(40, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.deepEqual(contentPosition(), 18, "Test move to down line");

			oEvent = createNativeEvent(40, false, true, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), 'World Hello World ', "Test select to down line");

			clean();
			getLogicDocumentWithParagraphs(['']);
			createComboBox();
			setFillingFormsMode(true);
			oEvent = createNativeEvent(40, false, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(AscTest.GetParagraphText(oGlobalLogicDocument.Content[0]), 'Hello', "Test select next option in combo box");
			onKeyDown(oEvent);
			oAssert.strictEqual(AscTest.GetParagraphText(oGlobalLogicDocument.Content[0]), 'World', "Test select next option in combo box");
			setFillingFormsMode(false);
		});

		QUnit.test("Test remove front", (oAssert) =>
		{
			let oEvent;
			oEvent = createNativeEvent(46, false, false, false, false);
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World"], true);
			moveToParagraph(oParagraph, true);
			onKeyDown(oEvent);
			selectAll();
			oAssert.strictEqual(getSelectedText(), 'ello World', 'Test remove front symbol');

			oEvent = createNativeEvent(46, true, false, false, false);
			moveToParagraph(oParagraph, true);
			onKeyDown(oEvent);
			selectAll();
			oAssert.strictEqual(getSelectedText(), 'World', 'Test remove front word');

			// todo
			moveToParagraph(oParagraph);
			const oDrawing = createShape();
			selectParaDrawing(oDrawing);
			oEvent = createNativeEvent(46, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oParagraph.GetRunByElement(oDrawing), null, 'Test remove shape');

			const oInlineLvlSdt = createComboBox();
			oEvent = createNativeEvent(46, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oParagraph.GetRunByElement(oInlineLvlSdt), null, 'Test remove form');
		});

		QUnit.test("Test replace unicode code to symbol", (oAssert) =>
		{
			let oEvent = createNativeEvent(88, false, false, true, false);
			const {oParagraph} = getLogicDocumentWithParagraphs(["2601"]);
			moveToParagraph(oParagraph, true);
			moveCursorRight(true, true);
			onKeyDown(oEvent);
			oAssert.strictEqual(getSelectedText(), '☁', 'Test replace unicode code to symbol');
		});
		QUnit.test("Test show context menu", (oAssert) =>
		{
			let oEvent;
			const {oParagraph} = getLogicDocumentWithParagraphs(["Hello Text"]);
			moveToParagraph(oParagraph, true);

			oEvent = createNativeEvent(93, false, false, false, false);
			executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);

			AscCommon.AscBrowser.isOpera = true;
			oEvent = createNativeEvent(57351, false, false, false, false);
			executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);
			AscCommon.AscBrowser.isOpera = false;

			oEvent = createNativeEvent(121, false, true, false, false);
			executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);
		});
		QUnit.test("Test disable numlock", (oAssert) =>
		{
			let oEvent = createNativeEvent(144, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oEvent.isDefaultPrevented, true, 'Test prevent default on numlock');
		});
		QUnit.test("Test disable scroll lock", (oAssert) =>
		{
			let oEvent = createNativeEvent(145, false, false, false, false);
			onKeyDown(oEvent);
			oAssert.strictEqual(oEvent.isDefaultPrevented, true, 'Test prevent default on scroll lock');
		});
		QUnit.test("Test add SJK test", (oAssert) =>
		{
			let oEvent = createNativeEvent(12288, false, false, false);
			checkTextAfterKeyDownHelperEmpty(' ', oEvent, oAssert, 'Check add space after SJK space');
		});
	});
})(window);
