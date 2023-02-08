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
(function (window) {

  // editor.sync_slidePropCallback = Asc.asc_docs_api.prototype.sync_slidePropCallback;
  // editor.sync_BeginCatchSelectedElements = Asc.asc_docs_api.prototype.sync_BeginCatchSelectedElements;
  // editor.sync_PrLineSpacingCallBack = Asc.asc_docs_api.prototype.sync_PrLineSpacingCallBack;
  // editor.sync_EndCatchSelectedElements = Asc.asc_docs_api.prototype.sync_EndCatchSelectedElements;
  // editor.UpdateParagraphProp = Asc.asc_docs_api.prototype.UpdateParagraphProp;
  // editor.sync_ParaSpacingLine = Asc.asc_docs_api.prototype.sync_ParaSpacingLine;
  // editor.Update_ParaInd = Asc.asc_docs_api.prototype.Update_ParaInd;
  // editor.sync_PrAlignCallBack = Asc.asc_docs_api.prototype.sync_PrAlignCallBack;
  // editor.sync_ParaStyleName = Asc.asc_docs_api.prototype.sync_ParaStyleName;
  // editor.sync_ListType = Asc.asc_docs_api.prototype.sync_ListType;
  // editor.sync_PrPropCallback = Asc.asc_docs_api.prototype.sync_PrPropCallback;
  // editor.Internal_Update_Ind_Left = Asc.asc_docs_api.prototype.Internal_Update_Ind_Left;
  // editor.Internal_Update_Ind_FirstLine = Asc.asc_docs_api.prototype.Internal_Update_Ind_FirstLine;
  // editor.Internal_Update_Ind_Right = Asc.asc_docs_api.prototype.Internal_Update_Ind_Right;
  // editor.ClearPropObjCallback = Asc.asc_docs_api.prototype.ClearPropObjCallback;
  // editor.sync_CanAddHyperlinkCallback = Asc.asc_docs_api.prototype.sync_CanAddHyperlinkCallback;
  // editor.textArtPreviewManager = {clear: function () {}};
  // editor.initDefaultShortcuts = Asc.asc_docs_api.prototype.initDefaultShortcuts;
  // editor.sync_shapePropCallback = Asc.asc_docs_api.prototype.sync_shapePropCallback;
  // editor.sync_VerticalTextAlign = Asc.asc_docs_api.prototype.sync_VerticalTextAlign;
  // editor.sync_Vert = Asc.asc_docs_api.prototype.sync_Vert;
  // editor.asc_registerCallback = Asc.asc_docs_api.prototype.asc_registerCallback;
  // editor.asc_unregisterCallback = Asc.asc_docs_api.prototype.asc_unregisterCallback;
  // editor.sendEvent = Asc.asc_docs_api.prototype.sendEvent;
  // editor._saveCheck = Asc.asc_docs_api.prototype._saveCheck;
  // editor.sync_ContextMenuCallback = Asc.asc_docs_api.prototype.sync_ContextMenuCallback;
  // editor.put_ShowParaMarks = Asc.asc_docs_api.prototype.put_ShowParaMarks;
  // editor.get_ShowParaMarks = Asc.asc_docs_api.prototype.get_ShowParaMarks;
  // editor.sync_ShowParaMarks = Asc.asc_docs_api.prototype.sync_ShowParaMarks;
  // editor.sync_DialogAddHyperlink = Asc.asc_docs_api.prototype.sync_DialogAddHyperlink;
  // editor.FontSizeOut = Asc.asc_docs_api.prototype.FontSizeOut;
  // editor.FontSizeIn = Asc.asc_docs_api.prototype.FontSizeIn;
  // editor.sync_HyperlinkClickCallback = Asc.asc_docs_api.prototype.sync_HyperlinkClickCallback;
  // editor.sync_EndAddShape = Asc.asc_docs_api.prototype.sync_EndAddShape;
  // editor.sync_PaintFormatCallback = Asc.asc_docs_api.prototype.sync_PaintFormatCallback;
  // editor.sync_MouseMoveStartCallback  = Asc.asc_docs_api.prototype.sync_MouseMoveStartCallback ;
  // editor.sync_MouseMoveEndCallback = Asc.asc_docs_api.prototype.sync_MouseMoveEndCallback;
  // editor.asc_hideComments = Asc.asc_docs_api.prototype.asc_hideComments;
  // editor.sync_HideComment = Asc.asc_docs_api.prototype.sync_HideComment;

  window.AscFonts = window.AscFonts || {};
  AscFonts.g_fontApplication = {
    GetFontInfo: function (sFontName) {
      if (sFontName === 'Cambria Math') {
        return new AscFonts.CFontInfo('Cambria Math', 40, 1, 433, 1,-1,-1,-1,-1,-1,-1);
      }
    },
    Init: function () {

    },
    LoadFont: function () {

    },
    GetFontInfoName: function () {
      
    }
  }

  window.g_fontApplication = AscFonts.g_fontApplication;

  AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback) {
    callback({'saveLock': false});
  };

  AscTest.CreateLogicDocument();
  editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState = function () {
  };

  function getThumbnails() {
    return {
      GetSelectedArray: function () {
        const oLogicDocument = editor.WordControl.m_oLogicDocument;
        return [oLogicDocument.CurPage];
      },
      IsSlideHidden: function () {
        return false;
      },
      SelectSlides: function () {
      },
    }
  }


  let oGlobalShape;
  const oGlobalLogicDocument = AscTest.CreateLogicDocument();

  function executeTestWithParams(fCallback, sTextIntoShape, bGlobalShape, bResetSelection) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    if (bResetSelection) {
      oController.resetSelection();
    }
    if (!oGlobalShape || oGlobalShape.bDeleted || !bGlobalShape) {
      oGlobalShape = AscTest.createShape(oGlobalLogicDocument.Slides[0]);
    }

    oGlobalShape.setTxBody(AscFormat.CreateTextBodyFromString(sTextIntoShape, editor.WordControl.m_oDrawingDocument, oGlobalShape));
    const oContent = oGlobalShape.txBody.content;
    const oParagraph = oContent.Content[0];

    fCallback({oLogicDocument: oGlobalLogicDocument, oShape: oGlobalShape, oParagraph, oController});
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

  let oUndoShape;

  function checkEditUndo(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    oUndoShape = AscTest.createShape(oLogicDocument.Slides[0]);

    oLogicDocument.OnKeyDown(oEvent);

    oAssert.strictEqual(oLogicDocument.Slides[0].cSld.spTree.length, 0, 'Check undo shortcut');
  }

  function checkEditRedo(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();

    oLogicDocument.OnKeyDown(oEvent);

    oAssert.strictEqual(oLogicDocument.Slides[0].cSld.spTree.length === 1 && oLogicDocument.Slides[0].cSld.spTree[0] === oUndoShape, true, 'Check redo shortcut');
  }

  function checkEditSelectAll(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();

    oLogicDocument.OnKeyDown(oEvent);

    const oController = oLogicDocument.GetCurrentController();
    oAssert.strictEqual(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oUndoShape, true, 'check select all shortcut');
  }

  function checkDuplicate(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    const oController = oLogicDocument.GetCurrentController();
    oController.resetSelection();
    
    editor.WordControl.m_oThumbnailsContainer.AbsolutePosition.SetParams(0, 0, 20, 20, 0, 0, 20, 20, 20, 20);
    editor.WordControl.m_oThumbnails.AbsolutePosition.SetParams(0, 0, 20, 20, 0, 0, 20, 20, 20, 20);
    editor.WordControl.Thumbnails.m_bIsVisible = true;
    editor.zoom100();
    editor.WordControl.Thumbnails.SelectSlides([0]);
    oLogicDocument.OnKeyDown(oEvent);

    const bCheck = oLogicDocument.Slides.length === 2 && oLogicDocument.Slides[0].cSld.spTree.length === oLogicDocument.Slides[1].cSld.spTree.length;
    oAssert.true(bCheck, 'check duplicate slides  shortcut');
  }

  function checkPrint(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    let bCheck = false;
    editor.asc_registerCallback('asc_onPrint', function () {
      bCheck = true;
    });

    oLogicDocument.OnKeyDown(oEvent);

    oAssert.true(bCheck, 'Check print shortcut');

    editor.asc_unregisterCallback('asc_onPrint');
  }

  function checkSave(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    const fOldSave = editor._onSaveCallbackInner;
    let bCheck = false;
    editor._onSaveCallbackInner = function () {
      bCheck = true;
      editor._onSaveCallbackInner = fOldSave;
    };

    oLogicDocument.OnKeyDown(oEvent);

    oAssert.strictEqual(bCheck, true, 'Check save shortcut');
  }

  function checkShowContextMenu(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      let bCheck = false;
      const fCheck = function () {
        bCheck = true;
      }
      editor.asc_registerCallback('asc_onContextMenu', fCheck);

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(bCheck, 'Check context menu shortcut');

      editor.asc_unregisterCallback('asc_onContextMenu', fCheck);
    }, '', true);
  }

  function checkShowParaMarks(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    editor.put_ShowParaMarks(false);

    oLogicDocument.OnKeyDown(oEvent);
    oAssert.true(!!editor.get_ShowParaMarks(), 'Check show para marks  shortcut');
  }

  function checkBold(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Bold(), 'Check bold shortcut');
    }, 'Hello World', true);
  }

  function checkCenterAlign(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Center, 'Check center align shortcut');
    }, 'Hello World', true);
  }

  function checkEuroSign(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      let bCheck = false;
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();

      oLogicDocument.OnKeyDown(oEvent);
      const sText = AscTest.GetParagraphText(oParagraph);
      oAssert.strictEqual(sText[sText.length - 1], '€', 'Check euro sign shortcut');
    }, 'Hello World', true);
  }

  let oGroup;
  let oFirstShape;
  let oSecondShape;

  function checkGroup(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    oFirstShape = AscTest.createShape(oLogicDocument.Slides[0]);
    oSecondShape = AscTest.createShape(oLogicDocument.Slides[0]);
    const oController = oLogicDocument.GetCurrentController();
    oFirstShape.select(oController, 0);
    oSecondShape.select(oController, 0);

    oLogicDocument.OnKeyDown(oEvent);
    oGroup = oFirstShape.group;
    oAssert.true(oFirstShape.group && (oFirstShape.group === oSecondShape.group), 'Check group shortcut');
  }

  function checkUnGroup(oEvent, oAssert) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    const oController = oLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGroup.select(oController);

    oLogicDocument.OnKeyDown(oEvent);
    oAssert.true(!oFirstShape.group && !oSecondShape.group && oGroup.bDeleted, 'Check ungroup shortcut');

  }

  function checkItalic(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Italic(), 'Check italic shortcut');
    }, 'Hello World', true);
  }

  function checkJustifyAlign(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Justify, 'check justify align shortcut');
    }, 'Hello World', true);
  }

  function checkAddHyperlink(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      let bCheck = false;
      let fCheck = function () {
        bCheck = true;
      };
      editor.asc_registerCallback('asc_onDialogAddHyperlink', fCheck);
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(bCheck, 'Check hyperlink shortcut');
      editor.asc_registerCallback('asc_onDialogAddHyperlink', fCheck);
    }, 'Hello World', true);
  }

  function checkBulletList(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();

      oLogicDocument.OnKeyDown(oEvent);
      const oBullet = oParagraph.Get_PresentationNumbering();
      oAssert.true(oBullet.m_nType === AscFormat.numbering_presentationnumfrmt_Char, 'Check bullet list shortcut');
    }, 'Hello World', true);
  }

  function checkLeftAlign(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Left, 'Check left align shortcut');
    }, 'Hello World', true);
  }

  function checkRightAlign(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Right, 'Check right align shortcut');
    }, 'Hello World', true);
  }

  function checkUnderline(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Underline(), 'Check underline shortcut');
    }, 'Hello World', true);
  }

  function checkStrikethrough(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Strikeout(), 'Check strikeout shortcut');
    }, 'Hello World', true);
  }

  let oCopyParagraphTextPr;

  function checkCopyFormat(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      const oTextPr = oRun.GetTextPr();
      oTextPr.SetBold(true);
      oTextPr.SetItalic(true);
      oTextPr.SetUnderline(true);
      oLogicDocument.SelectAll();
      oCopyParagraphTextPr = oParagraph.GetCalculatedTextPr();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.deepEqual(oLogicDocument.CopyTextPr, oCopyParagraphTextPr, 'Check copy format shortcut');
    }, 'Hello World', true);
  }

  function checkPasteFormat(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.deepEqual(oParagraph.GetCalculatedTextPr(), oCopyParagraphTextPr, 'check paste format shortcut');
    }, 'Hello World', true);
  }

  function checkSuperscript(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.OnKeyDown(oEvent);
      const oTextPr = oParagraph.GetCalculatedTextPr();
      oAssert.strictEqual(oTextPr.VertAlign, AscCommon.vertalign_SuperScript, 'Check superscript shortcut');
    }, 'Hello World', true);
  }

  function checkSubscript(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];

      oLogicDocument.OnKeyDown(oEvent);
      const oTextPr = oParagraph.GetCalculatedTextPr();
      oAssert.strictEqual(oTextPr.VertAlign, AscCommon.vertalign_SubScript, 'Check subscript shortcut');
    }, 'Hello World', true);
  }

  function checkEnDash(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();

      oLogicDocument.OnKeyDown(oEvent);
      const sText = AscTest.GetParagraphText(oParagraph);
      oAssert.strictEqual(sText.charCodeAt(sText.length - 1), 0x2013, 'Check en dash shortcut');
    }, 'Hello World', true);
  }

  function checkDecreaseFont(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      oRun.SetFontSize(10);
      oLogicDocument.SelectAll();
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.strictEqual(oRun.Get_FontSize(), 9, 'Check decrease font size shortcut');
    }, 'Hello World', true);
  }

  function checkIncreaseFont(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      oRun.SetFontSize(10);
      oLogicDocument.SelectAll();
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.strictEqual(oRun.Get_FontSize(), 11, 'Check increase font size shortcut');
    }, 'Hello World', true);
  }


  function checkDeleteBack(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'Hello Worl', 'Check delete with backspace');
    }, 'Hello World', true);
  }

  function checkDeleteWordBack(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'Hello ', 'Check delete word with backspace');
    }, 'Hello World', true);
  }

  let oTable
  function checkMoveToNextCell(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument}) {
      const oGraphicFrame = oLogicDocument.Add_FlowTable(3, 3);
      oTable = oGraphicFrame.graphicObject;
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oTable.CurCell.Index, 1, 'check go to next cell shortcut');
    }, '', true, true);
  }

  function checkMoveToPreviousCell(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument}) {
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oTable.CurCell.Index, 0, 'check go to previous cell shortcut');
    }, '', true);
  }
  
  let oBulletParagraph;
  function checkIncreaseBulletIndent(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      const oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
      oParagraph.Add_PresentationNumbering(oBullet);
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToStartPos();
      oParagraph.Set_Ind({Left: 0});
      oLogicDocument.OnKeyDown(oEvent);
      oBulletParagraph = oParagraph;
      oAssert.strictEqual(oParagraph.Pr.Get_IndLeft(), 11.1125, 'Check bullet indent shortcut');
    }, 'Hello', true);
  }

  function checkDecreaseBulletIndent(oEvent, oAssert) {
    oBulletParagraph.SetThisElementCurrent();
    oBulletParagraph.MoveCursorToStartPos();
    oGlobalLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oBulletParagraph.Pr.Get_IndLeft(), 0, 'Check bullet indent shortcut');
  }

  function checkAddTab(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      oLogicDocument.OnKeyDown(oEvent);
      let bCheck = false;
      for (let i = oParagraph.Content.length - 2; i >= 0; --i) {
        const oRun = oParagraph.Content[i];
        if (oRun.Content.length && !oRun.IsParaEndRun()) {
          bCheck = oRun.Content[oRun.Content.length - 1].Type === para_Tab;
          break;
        }
      }
      oAssert.true(bCheck, 'Check add tab');
    }, 'Hello text', true);
  }

  function checkSelectNextObject(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oController}) {
      oShape.select(oController, 0);
      oLogicDocument.OnKeyDown(oEvent);
      const arrSpTree = oLogicDocument.Slides[0].cSld.spTree;
      let oSelectedShape;
      for (let i = 0; i < arrSpTree.length; i += 1) {
        if (arrSpTree[i] === oShape) {
          oSelectedShape = arrSpTree[i < arrSpTree.length - 1 ? i + 1 : 0];
        }
      }
      oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oSelectedShape && oController.selectedObjects[0] !== oShape, 'Check select next object');
    }, '', false, true);
  }

  function checkSelectPreviousObject(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oShape, oController}) {
      oShape.select(oController, 0);
      oLogicDocument.OnKeyDown(oEvent);
      const arrSpTree = oLogicDocument.Slides[0].cSld.spTree;
      let oSelectedShape;
      for (let i = 0; i < arrSpTree.length; i += 1) {
        if (arrSpTree[i] === oShape) {
          oSelectedShape = arrSpTree[i > 0 ? i - 1 : arrSpTree.length - 1];
        }
      }
      oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oSelectedShape && oController.selectedObjects[0] !== oShape, 'Check select previous object');
    }, '', false, true);
  }

  function checkVisitHyperlink(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      oLogicDocument.AddHyperlink({Text: 'abcd', ToolTip: 'abcd', Value: 'ppaction://hlinkshowjump?jump=firstslide'});
      oLogicDocument.OnKeyDown(oEvent);

    }, 'Hello', true);
  }

  function checkSelectNextObjectWithPlaceholder(oEvent, oAssert) {

  }

  function checkAddNextSlide(oEvent, oAssert) {

  }

  function checkAddBreakLine(oEvent, oAssert) {

  }

  function checkAddTitleBreakLine(oEvent, oAssert) {

  }

  function checkAddMathBreakLine(oEvent, oAssert) {

  }

  function checkAddParagraph(oEvent, oAssert) {

  }

  function checkHandleEnter(oEvent, oAssert) {

  }

  function checkResetAddShape(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.changeCurrentState(new AscFormat.StartAddNewShape(oController, 'rect'));

    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.true(oController.curState instanceof AscFormat.NullState, 'Check reset add new shape');
  }

  let oGroupedShape1;
  let oGroupedShape2;
  let oTestGroup;

  function checkResetAllDrawingSelection(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGroupedShape1 = AscTest.createShape(oGlobalLogicDocument.Slides[0]);
    oGroupedShape2 = AscTest.createShape(oGlobalLogicDocument.Slides[0]);
    oGroupedShape1.select(oController, 0);
    oGroupedShape2.select(oController, 0);
    oController.createGroup();
    oGroupedShape1.select(oController, 0);
    oTestGroup = oGroupedShape1.group;
    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.true(oController.selectedObjects.length === 0, 'Check reset all selection');

  }

  function checkResetStepDrawingSelection(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oTestGroup.select(oController, 0);
    oGroupedShape1.select(oController, 0);
    oController.selection.groupSelection = oTestGroup;
    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oTestGroup && oTestGroup.selectedObjects.length === 0, 'Check reset step selection');
  }

  function checkNonBreakingSpace(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      oLogicDocument.OnKeyDown(oEvent);
      let bCheck = false;
      for (let i = oParagraph.Content.length - 2; i >= 0; --i) {
        const oRun = oParagraph.Content[i];
        if (oRun.Content.length && !oRun.IsParaEndRun()) {
          bCheck = oRun.Content[oRun.Content.length - 1].Value === 0x00A0;
          break;
        }
      }
      oAssert.true(bCheck, 'Check add non breaking space');
    }, 'Hello text', true);
  }

  function checkClearParagraphFormatting(oEvent, oAssert) {

  }

  function checkAddSpace(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      oLogicDocument.OnKeyDown(oEvent);
      let bCheck = false;
      for (let i = oParagraph.Content.length - 2; i >= 0; --i) {
        const oRun = oParagraph.Content[i];
        if (oRun.Content.length && !oRun.IsParaEndRun()) {
          bCheck = oRun.Content[oRun.Content.length - 1].Type === para_Space;
          break;
        }
      }
      oAssert.true(bCheck, 'Check add space');
    }, 'Hello text', true);
  }

  function checkMoveToUpperSlide(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGlobalLogicDocument.addNextSlide();
    oGlobalLogicDocument.Set_CurPage(1);
    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oGlobalLogicDocument.CurPage, 0);
  }

  function checkMoveToDownSlide(oEvent, oAssert) {
    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oGlobalLogicDocument.CurPage, 1);
  }

  function checkMoveToEndPosContent(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 25 && oPos.Y === 75, 'Check move cursor to end position shortcut');
    }, oEvent, false, true, false);
  }

  function checkMoveToEndLineContent(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 100 && oPos.Y === 15, 'Check move cursor to end line shortcut');
    }, oEvent, false, true, false);
  }

  function checkSelectToEndLineContent(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'HelloworldHelloworld', 'Check select text to end line shortcut');
    }, oEvent, false, false, true);
  }

  function checkMoveToEndSlide(oEvent, oAssert) {

  }

  function checkSelectToEndSlide(oEvent, oAssert) {

  }

  function checkMoveToStartPosContent(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 0 && oPos.Y === 15, 'Check move to start position shortcut');
    }, oEvent, true, true, false);
  }

  function checkMoveToStartLineContent(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 0 && oPos.Y === 75, 'Check move to start line shortcut');
    }, oEvent, true, true, false);
  }

  function testMoveHelper(fTestCallback, oEvent, bMoveToEndPosition, bGetPos, bGetSelectedText) {
    executeTestWithParams(function ({oShape, oParagraph, oLogicDocument}) {
      oShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});
      oParagraph.SetThisElementCurrent();
      oParagraph.Pr.SetInd(0, 0, 0);
      oParagraph.Set_Align(AscCommon.align_Left);
      if (bMoveToEndPosition) {
        oParagraph.MoveCursorToEndPos();
      } else {
        oParagraph.MoveCursorToStartPos();
      }
      oShape.recalculateContentWitCompiledPr();
      oParagraph.RecalculateCurPos(true, true);
      
      oLogicDocument.OnKeyDown(oEvent);

      let oPos;
      oParagraph.RecalculateCurPos(true, true);
      if (bGetPos) {
        
        oPos = oParagraph.GetCurPosXY(true, true);
      }
      let sSelectedText;
      if (bGetSelectedText) {
        sSelectedText = oParagraph.GetSelectedText();
      }
      fTestCallback({oPos, sSelectedText});
    }, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', true, true);
  }

  function checkSelectToStartLineContent(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText}) {
      oAssert.strictEqual(sSelectedText, 'Hello', 'Check select to start line shortcut');
    }, oEvent, true, false, true);
  }

  function checkMoveToStartSlide(oEvent, oAssert) {

  }

  function checkSelectToStartSlide(oEvent, oAssert) {

  }

  function checkMoveCursorLeft(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 20 && oPos.Y === 75, 'Check move cursor to end position shortcut');
    }, oEvent, true, true, false);
  }



  function checkGoToSlideUpper(oEvent, oAssert) {

  }

  function checkSelectCursorLeft(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'o', 'Check select text to left position shortcut');
    }, oEvent, true, false, true);
  }

  function checkSelectWordCursorLeft(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', 'Check select word text to left position shortcut');
    }, oEvent, true, false, true);
  }

  function checkMoveCursorWordLeft(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 0 && oPos.Y === 15, 'Check move cursor to left word position shortcut');
    }, oEvent, true, true, false);
  }

  function checkMoveCursorRight(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 5 && oPos.Y === 15, 'Check move cursor to right position shortcut');
    }, oEvent, false, true, false);
  }



  function checkGoToSlideDown(oEvent, oAssert) {

  }

  function checkSelectCursorRight(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'H', 'Check select text to right position shortcut');
    }, oEvent, false, false, true);
  }

  function checkSelectWordCursorRight(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', 'Check select word text to right position shortcut');
    }, oEvent, false, false, true);
  }

  function checkMoveCursorWordRight(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 25 && oPos.Y === 75, 'Check move cursor to right word position shortcut');
    }, oEvent, true, true, false);
  }

  function checkMoveCursorTop(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 25 && oPos.Y === 55, 'Check move cursor to top position shortcut');
    }, oEvent, true, true, false);
  }

  function checkGoToSlideUpper(oEvent, oAssert) {

  }

  function checkSelectCursorTop(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'worldHelloworldHello', 'Check select text to top position shortcut');
    }, oEvent, true, false, true);
  }

  function checkMoveCursorBottom(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.true(oPos.X === 0 && oPos.Y === 35, 'Check move cursor to bottom position shortcut');
    }, oEvent, false, true, false);
  }

  function checkSelectCursorBottom(oEvent, oAssert) {
    testMoveHelper(function ({sSelectedText, oPos}) {
      oAssert.strictEqual(sSelectedText, 'HelloworldHelloworld', 'Check select text to bottom position shortcut');
    }, oEvent, false, false, true);
  }

  function checkMoveShapeBottom(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGlobalShape.spPr.xfrm.setOffY(0);
    oGlobalShape.select(oController, 0);
    oGlobalShape.recalculateTransform();
    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oGlobalShape.y, 5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape bottom');
  }

  function checkMoveShapeTop(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGlobalShape.spPr.xfrm.setOffY(0);
    oGlobalShape.select(oController, 0);
    oGlobalShape.recalculateTransform();
    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oGlobalShape.y, -5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape top');
  }

  function checkMoveShapeRight(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGlobalShape.spPr.xfrm.setOffX(0);
    oGlobalShape.select(oController, 0);
    oGlobalShape.recalculateTransform();

    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oGlobalShape.x, 5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape right');
  }

  function checkMoveShapeLeft(oEvent, oAssert) {
    const oController = oGlobalLogicDocument.GetCurrentController();
    oController.resetSelection();
    oGlobalShape.spPr.xfrm.setOffX(0);
    oGlobalShape.select(oController, 0);
    oGlobalShape.recalculateTransform();

    oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oGlobalShape.x, -5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape left');
  }

  function checkDeleteFront(oEvent, oAssert) {
    executeTestWithParams(function ({oParagraph, oLogicDocument}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToStartPos();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'ello world', 'Check delete front shortcut');
    }, 'Hello world', true);
  }

  function checkDeleteWordFront(oEvent, oAssert) {
    executeTestWithParams(function ({oParagraph, oLogicDocument}) {
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToStartPos();

      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'world', 'Check delete front word shortcut');
    }, 'Hello world', true);
  }

  function checkIncreaseIndent(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph, oShape}) {
      oParagraph.Pr.SetInd(0, 0, 0);
      oParagraph.Set_PresentationLevel(0);
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.strictEqual(oParagraph.Pr.GetIndLeft(), 11.1125, 'Check increase indent');
    }, 'Hello', true);
  }

  function checkDecreaseIndent(oEvent, oAssert) {
    executeTestWithParams(function ({oLogicDocument, oParagraph, oShape}) {
      oParagraph.Pr.SetInd(0, 12, 0);
      oParagraph.Set_PresentationLevel(1);
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToStartPos();
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.true(AscFormat.fApproxEqual(oParagraph.Pr.GetIndLeft(), 0.8875), 'Check decrease indent');
    }, 'Hello', true);
  }

  function checkAddNextSlide(oEvent, oAssert) {

  }

  function checkNumLock(oEvent, oAssert) {
    const oRes = oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oRes & keydownresult_PreventDefault, keydownresult_PreventDefault, 'Check prevent default on num lock');
  }

  function checkScrollLock(oEvent, oAssert) {
    const oRes = oGlobalLogicDocument.OnKeyDown(oEvent);
    oAssert.strictEqual(oRes & keydownresult_PreventDefault, keydownresult_PreventDefault, 'Check prevent default on scroll lock');
  }

  $(function () {
    QUnit.module('Unit-tests for Shortcuts');

    QUnit.test('Test common shortcuts', function (oAssert) {

      editor.initDefaultShortcuts();
      let oEvent;

      // add shape and do undo
      oEvent = createEvent(90, true, false, false, false, false);
      checkEditUndo(oEvent, oAssert);

      //do redo after undo
      oEvent = createEvent(89, true, false, false, false, false);
      checkEditRedo(oEvent, oAssert);

      // add shape and check select
      oEvent = createEvent(65, true, false, false, false, false);
      checkEditSelectAll(oEvent, oAssert);

      oEvent = createEvent(68, true, false, false, false, false);
      checkDuplicate(oEvent, oAssert);

      oEvent = createEvent(80, true, false, false, false, false);
      checkPrint(oEvent, oAssert);

      oEvent = createEvent(83, true, false, false, false, false);
      checkSave(oEvent, oAssert);

      oEvent = createEvent(93, false, false, false, false, false);
      checkShowContextMenu(oEvent, oAssert);

      oEvent = createEvent(121, false, true, false, false, false);
      checkShowContextMenu(oEvent, oAssert);

      oEvent = createEvent(57351, false, false, false, false, false);
      checkShowContextMenu(oEvent, oAssert);

      oEvent = createEvent(56, true, true, false, false, false);
      checkShowParaMarks(oEvent, oAssert);

      oEvent = createEvent(66, true, false, false, false, false);
      checkBold(oEvent, oAssert);

      oEvent = createEvent(67, true, true, false, false, false);
      checkCopyFormat(oEvent, oAssert);

      oEvent = createEvent(69, true, false, false, false, false);
      checkCenterAlign(oEvent, oAssert);

      oEvent = createEvent(69, true, false, true, false, false);
      checkEuroSign(oEvent, oAssert);

      // group shapes
      oEvent = createEvent(71, true, false, false, false, false);
      checkGroup(oEvent, oAssert);

      // then ungroup
      oEvent = createEvent(71, true, true, false, false, false);
      checkUnGroup(oEvent, oAssert);

      oEvent = createEvent(73, true, false, false, false, false);
      checkItalic(oEvent, oAssert);

      oEvent = createEvent(74, true, false, false, false, false);
      checkJustifyAlign(oEvent, oAssert);

      oEvent = createEvent(75, true, false, false, false, false);
      checkAddHyperlink(oEvent, oAssert);

      oEvent = createEvent(76, true, true, false, false, false);
      checkBulletList(oEvent, oAssert);

      oEvent = createEvent(76, true, false, false, false, false);
      checkLeftAlign(oEvent, oAssert);

      oEvent = createEvent(82, true, false, false, false, false);
      checkRightAlign(oEvent, oAssert);

      oEvent = createEvent(85, true, false, false, false, false);
      checkUnderline(oEvent, oAssert);

      oEvent = createEvent(53, true, false, false, false, false);
      checkStrikethrough(oEvent, oAssert);

      oEvent = createEvent(83, true, true, false, false, false);
      checkPasteFormat(oEvent, oAssert);

      oEvent = createEvent(187, true, true, false, false, false);
      checkSuperscript(oEvent, oAssert);

      oEvent = createEvent(188, true, false, false, false, false);
      checkSuperscript(oEvent, oAssert);

      oEvent = createEvent(187, true, false, false, false, false);
      checkSubscript(oEvent, oAssert);

      oEvent = createEvent(190, true, false, false, false, false);
      checkSubscript(oEvent, oAssert);

      oEvent = createEvent(189, true, true, false, false, false);
      checkEnDash(oEvent, oAssert);

      oEvent = createEvent(219, true, false, false, false, false);
      checkDecreaseFont(oEvent, oAssert);

      oEvent = createEvent(221, true, false, false, false, false);
      checkIncreaseFont(oEvent, oAssert);

      oEvent = createEvent(8, false, false, false, false, false, false);
      checkDeleteBack(oEvent, oAssert);
      oEvent = createEvent(8, true, false, false, false, false, false);
      checkDeleteWordBack(oEvent, oAssert);
      //Tab
      oEvent = createEvent(9, false, false, false, false, false, false);
      checkMoveToNextCell(oEvent, oAssert);
      oEvent = createEvent(9, false, true, false, false, false, false);
      checkMoveToPreviousCell(oEvent, oAssert);
      oEvent = createEvent(9, false, false, false, false, false, false);
      checkIncreaseBulletIndent(oEvent, oAssert);
      oEvent = createEvent(9, false, true, false, false, false, false);
      checkDecreaseBulletIndent(oEvent, oAssert);
      oEvent = createEvent(9, false, false, false, false, false, false);
      checkAddTab(oEvent, oAssert);
      oEvent = createEvent(9, false, false, false, false, false, false);
      checkSelectNextObject(oEvent, oAssert);
      oEvent = createEvent(9, false, true, false, false, false, false);
      checkSelectPreviousObject(oEvent, oAssert);
      // Enter
      oEvent = createEvent(13, false, false, false, false, false, false);
      checkVisitHyperlink(oEvent, oAssert);
      oEvent = createEvent(13, false, true, false, false, false, false);
      checkSelectNextObjectWithPlaceholder(oEvent, oAssert);
      oEvent = createEvent(13, true, false, false, false, false, false);
      checkAddNextSlide(oEvent, oAssert);
      oEvent = createEvent(13, false, true, false, false, false, false);
      checkAddBreakLine(oEvent, oAssert);
      oEvent = createEvent(13, false, true, false, false, false, false);
      checkAddMathBreakLine(oEvent, oAssert);
      oEvent = createEvent(13, false, false, false, false, false, false);
      checkAddTitleBreakLine(oEvent, oAssert);
      oEvent = createEvent(13, false, false, false, false, false, false);
      checkAddMathBreakLine(oEvent, oAssert);
      oEvent = createEvent(13, false, false, false, false, false, false);
      checkAddParagraph(oEvent, oAssert);
      oEvent = createEvent(13, false, false, false, false, false, false);
      checkHandleEnter(oEvent, oAssert);
      // Esc
      oEvent = createEvent(27, false, false, false, false, false, false);
      checkResetAddShape(oEvent, oAssert);
      oEvent = createEvent(27, false, true, false, false, false, false);
      checkResetAllDrawingSelection(oEvent, oAssert);
      oEvent = createEvent(27, false, false, false, false, false, false);
      checkResetStepDrawingSelection(oEvent, oAssert);

      // Space
      oEvent = createEvent(32, true, true, false, false, false, false);
      checkNonBreakingSpace(oEvent, oAssert);
      oEvent = createEvent(32, true, false, false, false, false, false);
      checkClearParagraphFormatting(oEvent, oAssert);
      oEvent = createEvent(32, false, false, false, false, false, false);
      checkAddSpace(oEvent, oAssert);
      //pgUp
      oEvent = createEvent(33, false, false, false, false, false, false);
      checkMoveToUpperSlide(oEvent, oAssert);
      //pgDn
      oEvent = createEvent(34, false, false, false, false, false, false);
      checkMoveToDownSlide(oEvent, oAssert);
      //End
      oEvent = createEvent(35, true, false, false, false, false, false);
      checkMoveToEndPosContent(oEvent, oAssert);
      oEvent = createEvent(35, false, false, false, false, false, false);
      checkMoveToEndLineContent(oEvent, oAssert);
      oEvent = createEvent(35, false, true, false, false, false, false);
      checkSelectToEndLineContent(oEvent, oAssert);
      oEvent = createEvent(35, false, false, false, false, false, false);
      checkMoveToEndSlide(oEvent, oAssert);
      oEvent = createEvent(35, false, true, false, false, false, false);
      checkSelectToEndSlide(oEvent, oAssert);
      // Home
      oEvent = createEvent(36, true, false, false, false, false, false);
      checkMoveToStartPosContent(oEvent, oAssert);
      oEvent = createEvent(36, false, false, false, false, false, false);
      checkMoveToStartLineContent(oEvent, oAssert);
      oEvent = createEvent(36, false, true, false, false, false, false);
      checkSelectToStartLineContent(oEvent, oAssert);
      oEvent = createEvent(36, false, false, false, false, false, false);
      checkMoveToStartSlide(oEvent, oAssert);
      oEvent = createEvent(36, false, true, false, false, false, false);
      checkSelectToStartSlide(oEvent, oAssert);

      //Left arrow
      oEvent = createEvent(37, false, false, false, false, false, false);
      checkMoveCursorLeft(oEvent, oAssert);

      oEvent = createEvent(37, false, false, false, false, false, false);
      checkGoToSlideUpper(oEvent, oAssert);
      oEvent = createEvent(37, false, true, false, false, false, false);
      checkSelectCursorLeft(oEvent, oAssert);
      oEvent = createEvent(37, true, true, false, false, false, false);
      checkSelectWordCursorLeft(oEvent, oAssert);
      oEvent = createEvent(37, true, false, false, false, false, false);
      checkMoveCursorWordLeft(oEvent, oAssert);
      //Right arrow
      oEvent = createEvent(39, false, false, false, false, false, false);
      checkMoveCursorRight(oEvent, oAssert);

      oEvent = createEvent(39, false, false, false, false, false, false);
      checkGoToSlideDown(oEvent, oAssert);
      oEvent = createEvent(39, false, true, false, false, false, false);
      checkSelectCursorRight(oEvent, oAssert);
      oEvent = createEvent(39, true, true, false, false, false, false);
      checkSelectWordCursorRight(oEvent, oAssert);
      oEvent = createEvent(39, true, false, false, false, false, false);
      checkMoveCursorWordRight(oEvent, oAssert);
      //Top arrow
      oEvent = createEvent(38, false, false, false, false, false, false);
      checkMoveCursorTop(oEvent, oAssert);

      oEvent = createEvent(38, false, false, false, false, false, false);
      checkGoToSlideUpper(oEvent, oAssert);
      oEvent = createEvent(38, false, true, false, false, false, false);
      checkSelectCursorTop(oEvent, oAssert);
      // Bottom arrow
      oEvent = createEvent(40, false, false, false, false, false, false);
      checkMoveCursorBottom(oEvent, oAssert);

      oEvent = createEvent(40, false, false, false, false, false, false);
      checkGoToSlideDown(oEvent, oAssert);
      oEvent = createEvent(40, false, true, false, false, false, false);
      checkSelectCursorBottom(oEvent, oAssert);

      // Check move shape
      oEvent = createEvent(40, false, false, false, false, false, false);
      checkMoveShapeBottom(oEvent, oAssert);
      oEvent = createEvent(38, false, false, false, false, false, false);
      checkMoveShapeTop(oEvent, oAssert);
      oEvent = createEvent(39, false, false, false, false, false, false);
      checkMoveShapeRight(oEvent, oAssert);
      oEvent = createEvent(37, false, false, false, false, false, false);
      checkMoveShapeLeft(oEvent, oAssert);

      //Delete
      oEvent = createEvent(46, false, false, false, false, false, false);
      checkDeleteFront(oEvent, oAssert);
      oEvent = createEvent(46, true, false, false, false, false, false);
      checkDeleteWordFront(oEvent, oAssert);

      oEvent = createEvent(77, true, false, false, false, false, false);
      checkIncreaseIndent(oEvent, oAssert);
      oEvent = createEvent(77, true, true, false, false, false, false);
      checkDecreaseIndent(oEvent, oAssert);
      oEvent = createEvent(77, true, false, false, false, false, false);
      checkAddNextSlide(oEvent, oAssert);

      oEvent = createEvent(144, false, false, false, false, false, false);
      checkNumLock(oEvent, oAssert);
      oEvent = createEvent(145, false, false, false, false, false, false);
      checkScrollLock(oEvent, oAssert);
    });
  });

})(window);
