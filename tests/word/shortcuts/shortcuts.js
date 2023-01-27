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


(function(window) {
  window.AscFonts = window.AscFonts || {};
  AscFonts.g_fontApplication = {
    GetFontInfo: function (sFontName) {
      if (sFontName === 'Cambria Math') {
        return new AscFonts.CFontInfo('Cambria Math', 40, 1, 433, 1,-1,-1,-1,-1,-1,-1);
      }
    },
    Init: function () {

    },
  }

  window.g_fontApplication = AscFonts.g_fontApplication;

  Asc.createPluginsManager = function () {

  };
  const editor = window.editor;

  editor.initDefaultShortcuts = Asc.asc_docs_api.prototype.initDefaultShortcuts;
  editor._InitCommonShortcuts = Asc.asc_docs_api.prototype._InitCommonShortcuts;
  editor._InitWindowsShortcuts = Asc.asc_docs_api.prototype._InitWindowsShortcuts;
  editor._InitMacOsShortcuts = Asc.asc_docs_api.prototype._InitMacOsShortcuts;
  editor.put_ShowParaMarks = Asc.asc_docs_api.prototype.put_ShowParaMarks;
  editor.get_ShowParaMarks = Asc.asc_docs_api.prototype.get_ShowParaMarks;
  editor.sync_ShowParaMarks = Asc.asc_docs_api.prototype.sync_ShowParaMarks;
  editor.FontSizeOut = Asc.asc_docs_api.prototype.FontSizeOut;
  editor.FontSizeIn = Asc.asc_docs_api.prototype.FontSizeIn;
  editor.asc_AddMath = Asc.asc_docs_api.prototype.asc_AddMath;
  editor._InitVariablesOnEndLoadSdk = Asc.asc_docs_api.prototype._InitVariablesOnEndLoadSdk;
  editor.asc_AddMath2 = Asc.asc_docs_api.prototype.asc_AddMath2;

  editor._InitVariablesOnEndLoadSdk();
  AscCommon.g_font_loader = {
    LoadFont: function () {
      return false;
    }
  }
  function getLogicDocumentWithParagraphs(arrText) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    resetLogicDocument(oLogicDocument);
    oLogicDocument.Start_SilentMode();
    oLogicDocument.RemoveFromContent(0, oLogicDocument.GetElementsCount(), false);
    if (Array.isArray(arrText)) {
      for (let i = 0; i < arrText.length; i += 1) {
        const oParagraph = AscTest.CreateParagraph();
        oLogicDocument.AddToContent(oLogicDocument.Content.length, oParagraph);
        oParagraph.MoveCursorToEndPos();
        const oRun = new AscWord.CRun();
        oParagraph.AddToContent(0, oRun);
        oRun.AddText(arrText[i]);
      }
    }


    return oLogicDocument;
  }
  function createEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsAltGr, bIsMacCmdKey) {
    const oKeyBoardEvent = new AscCommon.CKeyboardEvent();
    oKeyBoardEvent.KeyCode = nKeyCode;
    oKeyBoardEvent.ShiftKey = bIsShift;
    oKeyBoardEvent.AltKey = bIsAlt;
    oKeyBoardEvent.CtrlKey = bIsCtrl;
    oKeyBoardEvent.MacCmdKey = bIsMacCmdKey;
    oKeyBoardEvent.AltGr = bIsAltGr;
    return oKeyBoardEvent;
  }

  function checkInsertPageBreak(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);

    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].Check_PageBreak();
  }

  function checkInsertLineBreak(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);

    oLogicDocument.OnKeyDown(event);
    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].IsLineBreak && oRun.Content[j].IsLineBreak()) {
          return true;
        }
      }
    }
    return false;
  }
  function checkInsertColumnBreak(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);

    oLogicDocument.OnKeyDown(event);
    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].IsColumnBreak && oRun.Content[j].IsColumnBreak()) {
          return true;
        }
      }
    }
    return false;
  }

  function checkResetChar(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);
    const oParagraph = oLogicDocument.Content[0];
    const oRun = new AscWord.CRun();
    oParagraph.AddToContent(0, oRun);
    oRun.AddText("Hello Word!");
    oLogicDocument.ApplyToAll = true;
    oLogicDocument.SelectAll();
    oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Bold: true, Italic: true, Underline: true}));

    oLogicDocument.OnKeyDown(event);
    return !(oRun.Get_Bold() || oRun.Get_Italic() || oRun.Get_Underline());
  }

  function checkNonBreakingSpace(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);
    oLogicDocument.OnKeyDown(event);
    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x00A0) {
          return true;
        }
      }
    }
    return false;
  }

  function checkStrikeout(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oRun = oLogicDocument.Content[0].Content[0];

    return !!oRun.Get_Strikeout();
  }

  function checkShowNonPrintingCharacters(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);
    editor.put_ShowParaMarks(false);
    oLogicDocument.OnKeyDown(event);
    return !!editor.get_ShowParaMarks();
  }

  function checkSelectAll(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.GetSelectedText() === 'Hello World';
  }

  function checkBold(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oRun = oLogicDocument.Content[0].Content[0];

    return !!oRun.Get_Bold();
  }

  function checkCopyFormat(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    const oTextRun = oLogicDocument.Content[0].Content[0];
    oTextRun.SetBold(true);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.CopyTextPr.IsEqual(oTextRun.Pr);
  }

  function checkInsertCopyright(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);
    oLogicDocument.OnKeyDown(event);
    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x00A9) {
          return true;
        }
      }
    }
    return false;
  }
  function resetLogicDocument(oLogicDocument) {
    oLogicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
  }
  function checkInsertEndNote(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return !!oLogicDocument.Endnotes.CurEndnote;
  }

  function checkCenterPara(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].GetParagraphAlign() === AscCommon.align_Center;
  }

  function checkEuroSign(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x20AC) {
          return true;
        }
      }
    }
    return false;
  }

  function checkItalic(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oRun = oLogicDocument.Content[0].Content[0];

    return !!oRun.Get_Italic();
  }

  function checkJustifyPara(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].GetParagraphAlign() === AscCommon.align_Justify;
  }
// in our editors, we send an event to open a window with hyperlink settings, check if the event was sent
  function checkHyperlink(event, assert) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    let bCheck = false;
    const fOldSyncDialogAddHyperlink = editor.sync_DialogAddHyperlink;
    editor.sync_DialogAddHyperlink = function () {
      assert.true(true, 'Check hyperlink shortcut');
      bCheck = true;
    };
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    if (!bCheck) {
      assert.true(false, 'Check hyperlink shortcut');
    }
    editor.sync_DialogAddHyperlink = fOldSyncDialogAddHyperlink;
  }

  function checkBulletList(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    return oParagraph.IsBulletedNumbering();
  }

  function checkLeftPara(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.private_ToggleParagraphAlignByHotkey(AscCommon.align_Justify);
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].GetParagraphAlign() === AscCommon.align_Left;
  }

  function checkIndent(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    const nOldIndent = oLogicDocument.Content[0].Pr.Get_IndLeft();
    oLogicDocument.OnKeyDown(event);
    const nNewIndent = oLogicDocument.Content[0].Pr.Get_IndLeft();
    return nNewIndent !== nOldIndent;
  }

  function checkUnIndent(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.Content[0].Pr.SetInd(undefined, 12.5);
    const nOldIndent = oLogicDocument.Content[0].Pr.Get_IndLeft();
    oLogicDocument.OnKeyDown(event);
    const nNewIndent = oLogicDocument.Content[0].Pr.Get_IndLeft();
    return nNewIndent !== nOldIndent;
  }

  function checkPrintPreviewAndPrint(event, assert) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    let bCheck = false;
    const fOldOnPrint = editor.onPrint;
    editor.onPrint = function () {
      assert.true(true, 'Check hyperlink shortcut');
      bCheck = true;
    };
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    if (!bCheck) {
      assert.true(false, 'Check hyperlink shortcut');
    }
    editor.onPrint = fOldOnPrint;
  }

  function checkInsertPageNumber(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);

    oLogicDocument.OnKeyDown(event);
    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Type === para_PageNum) {
          return true;
        }
      }
    }
    return false;
  }

  function checkRightPara(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].GetParagraphAlign() === AscCommon.align_Right;
  }


  function checkRegisteredSign(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x00AE) {
          return true;
        }
      }
    }
    return false;
  }

  function checkSave(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs();
    const fOldSave = editor.asc_Save;
    let bCheck = false;
    editor.asc_Save = function () {
      bCheck = true;
    }
    oLogicDocument.OnKeyDown(event);
    editor.asc_Save = fOldSave;
    return bCheck;
  }

  function checkTradeMarkSign(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x2122) {
          return true;
        }
      }
    }
    return false;
  }

  function checkUnderline(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oRun = oLogicDocument.Content[0].Content[0];

    return !!oRun.Get_Underline();
  }

  function checkPasteFormat(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World', 'Hello word']);
    const oFirstParagraph = oLogicDocument.Content[0];
    const oFirstRun = oFirstParagraph.Content[0];
    oFirstRun.SetBold(true);
    oLogicDocument.MoveCursorToStartPos();
    oLogicDocument.SelectCurrentWord();
    oLogicDocument.Document_Format_Copy();
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oSecondParagraph = oLogicDocument.Content[1];
    const oSecondRun = oSecondParagraph.Content[0];
    return !!oSecondRun.Get_Bold();
  }

  function checkRedo(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    //oLogicDocument.OnKeyDown(event);
  }

  function checkUndo(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    //oLogicDocument.OnKeyDown(event);
  }

  function checkEnDash(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x2013) {
          return true;
        }
      }
    }
    return false;
  }

  function checkEmDash(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x2014) {
          return true;
        }
      }
    }
    return false;
  }

  function checkUpdateFields(event) {

  }

  function checkSuperscript(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].Content[0].Get_VertAlign() === AscCommon.vertalign_SuperScript;
  }

  function checkNonBreakingHyphen(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x002D) {
          return true;
        }
      }
    }
    return false;
  }

  function checkHorizontalEllipsis(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.OnKeyDown(event);

    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      for (let j = 0; j < oRun.Content.length; j += 1) {
        if (oRun.Content[j].Value === 0x2026) {
          return true;
        }
      }
    }
    return false;
  }

  function checkSubscript(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return oLogicDocument.Content[0].Content[0].Get_VertAlign() === AscCommon.vertalign_SubScript;
  }

  function checkIncreaseFontSize(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    const oRun = oLogicDocument.Content[0].Content[0];
    const nOldFontSize = oRun.Get_FontSize();
    oLogicDocument.OnKeyDown(event);
    const nNewFontSize = oRun.Get_FontSize();

    return nOldFontSize < nNewFontSize;
  }

  function checkDecreaseFontSize(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    const oRun = oLogicDocument.Content[0].Content[0];
    const nOldFontSize = oRun.Get_FontSize();
    oLogicDocument.OnKeyDown(event);
    const nNewFontSize = oRun.Get_FontSize();

    return nOldFontSize > nNewFontSize;
  }

  function checkApplyHeading1(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oParagraphPr = oLogicDocument.Content[0].Pr;
    const sPStyleName = oLogicDocument.Styles.Get_Name(oParagraphPr.Get_PStyle());
    return sPStyleName === 'Heading 1';
  }

  function checkApplyHeading2(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oParagraphPr = oLogicDocument.Content[0].Pr;
    const sPStyleName = oLogicDocument.Styles.Get_Name(oParagraphPr.Get_PStyle());
    return sPStyleName === 'Heading 2';
  }

  function checkApplyHeading3(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello World']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    const oParagraphPr = oLogicDocument.Content[0].Pr;
    const sPStyleName = oLogicDocument.Styles.Get_Name(oParagraphPr.Get_PStyle());
    return sPStyleName === 'Heading 3';
  }

  function checkInsertFootnote(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['Hello']);
    oLogicDocument.SelectAll();
    oLogicDocument.OnKeyDown(event);
    return !!oLogicDocument.Footnotes.CurFootnote;
  }

  function checkInsertEquation(event) {
    const oLogicDocument = getLogicDocumentWithParagraphs(['']);
    oLogicDocument.OnKeyDown(event);
    const oParagraph = oLogicDocument.Content[0];
    for (let i = 0; i < oParagraph.Content.length; i += 1) {
      const oRun = oParagraph.Content[i];
      if (oRun instanceof AscCommonWord.ParaMath) {
        return true;
      }
    }
    return false;
  }

  function checkEnter() {

  }

  function checkEsc() {

  }

  function checkSpace() {

  }

  function checkPgUp() {

  }

  function checkPgDn() {

  }

  function checkEnd() {

  }

  function checkHome() {

  }

  function checkLeftArrow() {

  }

  function checkTopArrow() {

  }

  function checkRightArrow() {

  }

  function checkBottomArrow() {

  }

  function checkDelete() {

  }

  function checkX() {

  }

  function checkContextMenu() {

  }

  function checkNumLock() {

  }

  function checkScrollLock() {

  }

  function checkCJKSpace() {

  }

  $(function () {

    QUnit.module("Unit-tests for Shortcuts");


    QUnit.test("Test common shortcuts", function (assert)
    {
      editor.initDefaultShortcuts();

      let event = createEvent(13, true, false, false, false, false);
      assert.strictEqual(checkInsertPageBreak(event), true, 'Check page break shortcut');

      event = createEvent(13, false, true, false,false,false);
      assert.strictEqual(checkInsertLineBreak(event), true, 'Check line break shortcut');

      event = createEvent(13, true, true, false,false,false);
      assert.strictEqual(checkInsertColumnBreak(event), true, 'Check column break shortcut');


      event = createEvent(32, true, false, false,false,false);
      assert.strictEqual(checkResetChar(event), true, 'Check reset char shortcut');


      event = createEvent(32, true, true, false,false,false);
      assert.strictEqual(checkNonBreakingSpace(event), true, 'Check add non breaking space shortcut');


      event = createEvent(53, true, false, false,false,false);
      assert.strictEqual(checkStrikeout(event), true, 'Check add strikeout shortcut');


      event = createEvent(56, true, true, false,false,false);
      assert.strictEqual(checkShowNonPrintingCharacters(event), true, 'Check show non printing characters shortcut');

      event = createEvent(65, true, false, false,false,false);
      assert.strictEqual(checkSelectAll(event), true, 'Check select all shortcut');

      event = createEvent(66, true, false, false,false,false);
      assert.strictEqual(checkBold(event), true, 'Check bold shortcut');

      event = createEvent(67, true, true, false,false,false);
      assert.strictEqual(checkCopyFormat(event), true, 'Check copy format shortcut');

      event = createEvent(67, true, false, true,false,false);
      assert.strictEqual(checkInsertCopyright(event), true, 'Check insert copyright shortcut');

      event = createEvent(68, true, false, true,false,false);
      assert.strictEqual(checkInsertEndNote(event), true, 'Check insert endnote shortcut');

      event = createEvent(69, true, false, false,false,false);
      assert.strictEqual(checkCenterPara(event), true, 'Check center para shortcut');

      event = createEvent(69, true, false, true,false,false);
      assert.strictEqual(checkEuroSign(event), true, 'Check insert euro sign shortcut');

      event = createEvent(73, true, false, false,false,false);
      assert.strictEqual(checkItalic(event), true, 'Check italic shortcut');


      event = createEvent(74, true, false, false,false,false);
      assert.strictEqual(checkJustifyPara(event), true, 'Check justify para shortcut');

      event = createEvent(75, true, false, false,false,false);
      checkHyperlink(event, assert);

      event = createEvent(76, true, true, false,false,false);
      assert.strictEqual(checkBulletList(event), true, 'Check bullet list shortcut');

      event = createEvent(76, true, false, false,false,false);
      assert.strictEqual(checkLeftPara(event), true, 'Check left para shortcut');

      event = createEvent(77, true, false, false,false,false);
      assert.strictEqual(checkIndent(event), true, 'Check indent shortcut');

      event = createEvent(77, true, true, false,false,false);
      assert.strictEqual(checkUnIndent(event), true, 'Check indent shortcut');

      event = createEvent(80, true, false, false,false,false);
      checkPrintPreviewAndPrint(event, assert);

      event = createEvent(80, true, true, false,false,false);
      assert.strictEqual(checkInsertPageNumber(event), true, 'Check insert page number shortcut');


      event = createEvent(82, true, false, false,false,false);
      assert.strictEqual(checkRightPara(event), true, 'Check right para shortcut');


      event = createEvent(82, true, false, true,false,false);
      assert.strictEqual(checkRegisteredSign(event), true, 'Check registered sign shortcut');


      event = createEvent(83, true, false, false,false,false);
      assert.strictEqual(checkSave(event), true, 'Check save shortcut');


      event = createEvent(84, true, false, true,false,false);
      assert.strictEqual(checkTradeMarkSign(event), true, 'Check trademark sign shortcut');

      event = createEvent(85, true, false, false,false,false);
      assert.strictEqual(checkUnderline(event), true, 'Check underline shortcut');


      event = createEvent(86, true, true, false,false,false);
      assert.strictEqual(checkPasteFormat(event), true, 'Check paste format shortcut');


      event = createEvent(89, true, false, false,false,false);
      assert.strictEqual(checkRedo(event), true, 'Check redo shortcut');


      event = createEvent(90, true, false, false,false,false);
      assert.strictEqual(checkUndo(event), true, 'Check undo shortcut');


      event = createEvent(109, true, false, false,false,false);
      assert.strictEqual(checkEnDash(event), true, 'Check en dash shortcut');


      event = createEvent(109, true, false, true,false,false);
      assert.strictEqual(checkEmDash(event), true, 'Check em dash shortcut');


      event = createEvent(120, false, false, false,false,false);
      assert.strictEqual(checkUpdateFields(event), true, 'Check update fields shortcut');


      event = createEvent(188, true, false, false,false,false);
      assert.strictEqual(checkSuperscript(event), true, 'Check superscript shortcut');


      event = createEvent(189, true, true, false,false,false);
      assert.strictEqual(checkNonBreakingHyphen(event), true, 'Check non breaking hyphen shortcut');


      event = createEvent(190, true, false, true,false,false);
      assert.strictEqual(checkHorizontalEllipsis(event), true, 'Check horizontal ellipsis shortcut');


      event = createEvent(190, true, false, false,false,false);
      assert.strictEqual(checkSubscript(event), true, 'Check subscript shortcut');


      event = createEvent(219, true, false, false,false,false);
      assert.strictEqual(checkIncreaseFontSize(event), true, 'Check increase font size shortcut');


      event = createEvent(221, true, false, false,false,false);
      assert.strictEqual(checkDecreaseFontSize(event), true, 'Check decrease font size shortcut');


      editor.Shortcuts = new AscCommon.CShortcuts();
    });

    QUnit.test("Test windows shortcuts", function (assert)
    {
      editor.initDefaultShortcuts();
      let event;
      event = createEvent(49, false, false, true, false, false);
      assert.strictEqual(checkApplyHeading1(event), true, 'Check apply heading1 shortcut');

      event = createEvent(50, false, false, true, false, false);
      assert.strictEqual(checkApplyHeading2(event), true, 'Check apply heading2 shortcut');

      event = createEvent(51, false, false, true, false, false);
      assert.strictEqual(checkApplyHeading3(event), true, 'Check apply heading3 shortcut');

      event = createEvent(70, true, false, true, false, false);
      assert.strictEqual(checkInsertFootnote(event), true, 'Check insert footnote shortcut');

      event = createEvent(187, false, false, true, false, false);
      assert.strictEqual(checkInsertEquation(event), true, 'Check insert equation shortcut');

      editor.Shortcuts = new AscCommon.CShortcuts();
    });

    QUnit.test("Test macOs shortcuts", function (assert)
    {
      AscCommon.AscBrowser.isMacOs = true;
      editor.initDefaultShortcuts();
      let event;
      event = createEvent(49, true, false, true, false, false);
      assert.strictEqual(checkApplyHeading1(event), true, 'Check apply heading1 shortcut');

      event = createEvent(50, true, false, true, false, false);
      assert.strictEqual(checkApplyHeading2(event), true, 'Check apply heading2 shortcut');


      event = createEvent(51, true, false, true, false, false);
      assert.strictEqual(checkApplyHeading2(event), true, 'Check apply heading3 shortcut');


      event = createEvent(187, true, false, true, false, false);
      assert.strictEqual(checkInsertEquation(event), true, 'Check insert equation shortcut');

      editor.Shortcuts = new AscCommon.CShortcuts();
    });

    QUnit.test("Test common hotkeys", function (assert)
    {
      editor.initDefaultShortcuts();

      const event = createEvent(13, true, false, false, false, false);
      assert.strictEqual(checkInsertPageBreak(event), true);

      editor.Shortcuts = new AscCommon.CShortcuts();
    });

    QUnit.test("Test windows hotkeys", function (assert)
    {
      editor.initDefaultShortcuts();

      const event = createEvent(13, true, false, false, false, false);
      assert.strictEqual(checkInsertPageBreak(event), true);

      editor.Shortcuts = new AscCommon.CShortcuts();
    });

    QUnit.test("Test macOs hotkeys", function (assert)
    {
      editor.initDefaultShortcuts();

      const event = createEvent(13, true, false, false, false, false);
      assert.strictEqual(checkInsertPageBreak(event), true);

      editor.Shortcuts = new AscCommon.CShortcuts();
    });

    function checkNonBlockedAlt(event) {
      const oLogicDocument = AscTest.CreateLogicDocument();
      oLogicDocument.Start_SilentMode();
      oLogicDocument.RemoveFromContent(0, oLogicDocument.GetElementsCount(), false);
      const oParagraph = AscTest.CreateParagraph();
      oLogicDocument.AddToContent(0, oParagraph);
      oParagraph.MoveCursorToEndPos();

      const nRetMouseDown = oLogicDocument.OnKeyDown(event);
      return (nRetMouseDown & keydownresult_PreventDefault) === 0;
    }

    QUnit.test("Test unlocked alt button for mac", function (assert)
    {
      const bOldIsMacOs = AscCommon.AscBrowser.isMacOs;
      AscCommon.AscBrowser.isMacOs = true;
      editor.initDefaultShortcuts();

      const arrCheckCodes = [48,49,50,51,52,53,54,55,56,57,189,187,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,
        84,85,86,87,88,89,90,219,221,186,222,220,188,190,191,96,97,98,99,100,101,102,103,104,105,111,106,
        109,110,107];
      for (let nKeyCodeIndex = 0, nLength = arrCheckCodes.length; nKeyCodeIndex < nLength; ++nKeyCodeIndex) {
        const oAltEvent = createEvent(arrCheckCodes[nKeyCodeIndex], false, false, true, false, false);
        assert.strictEqual(checkNonBlockedAlt(oAltEvent), true, `Check (${arrCheckCodes[nKeyCodeIndex]}) key code with Alt`);

        const oAltShiftEvent = createEvent(arrCheckCodes[nKeyCodeIndex], false, true, true, false, false);
        assert.strictEqual(checkNonBlockedAlt(oAltShiftEvent), true, `Check (${arrCheckCodes[nKeyCodeIndex]}) key code with Shift and Alt`);
      }

      editor.Shortcuts = new AscCommon.CShortcuts();
      AscCommon.AscBrowser.isMacOs = bOldIsMacOs;
    });

  });
})(window);
