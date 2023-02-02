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

  editor.sync_slidePropCallback = Asc.asc_docs_api.prototype.sync_slidePropCallback;
  editor.sync_BeginCatchSelectedElements = Asc.asc_docs_api.prototype.sync_BeginCatchSelectedElements;
  editor.sync_PrLineSpacingCallBack = Asc.asc_docs_api.prototype.sync_PrLineSpacingCallBack;
  editor.sync_EndCatchSelectedElements = Asc.asc_docs_api.prototype.sync_EndCatchSelectedElements;
  editor.UpdateParagraphProp = Asc.asc_docs_api.prototype.UpdateParagraphProp;
  editor.sync_ParaSpacingLine = Asc.asc_docs_api.prototype.sync_ParaSpacingLine;
  editor.Update_ParaInd = Asc.asc_docs_api.prototype.Update_ParaInd;
  editor.sync_PrAlignCallBack = Asc.asc_docs_api.prototype.sync_PrAlignCallBack;
  editor.sync_ParaStyleName = Asc.asc_docs_api.prototype.sync_ParaStyleName;
  editor.sync_ListType = Asc.asc_docs_api.prototype.sync_ListType;
  editor.sync_PrPropCallback = Asc.asc_docs_api.prototype.sync_PrPropCallback;
  editor.Internal_Update_Ind_Left = Asc.asc_docs_api.prototype.Internal_Update_Ind_Left;
  editor.Internal_Update_Ind_FirstLine = Asc.asc_docs_api.prototype.Internal_Update_Ind_FirstLine;
  editor.Internal_Update_Ind_Right = Asc.asc_docs_api.prototype.Internal_Update_Ind_Right;
  editor.ClearPropObjCallback = Asc.asc_docs_api.prototype.ClearPropObjCallback;
  editor.sync_CanAddHyperlinkCallback = Asc.asc_docs_api.prototype.sync_CanAddHyperlinkCallback;
  editor.textArtPreviewManager = {clear: function () {}};
  editor.initDefaultShortcuts = Asc.asc_docs_api.prototype.initDefaultShortcuts;
  editor.sync_shapePropCallback = Asc.asc_docs_api.prototype.sync_shapePropCallback;
  editor.sync_VerticalTextAlign = Asc.asc_docs_api.prototype.sync_VerticalTextAlign;
  editor.sync_Vert = Asc.asc_docs_api.prototype.sync_Vert;
  editor.asc_registerCallback = Asc.asc_docs_api.prototype.asc_registerCallback;
  editor.asc_unregisterCallback = Asc.asc_docs_api.prototype.asc_unregisterCallback;
  editor.sendEvent = Asc.asc_docs_api.prototype.sendEvent;
  editor._saveCheck = Asc.asc_docs_api.prototype._saveCheck;
  editor.sync_ContextMenuCallback = Asc.asc_docs_api.prototype.sync_ContextMenuCallback;
  editor.put_ShowParaMarks = Asc.asc_docs_api.prototype.put_ShowParaMarks;
  editor.get_ShowParaMarks = Asc.asc_docs_api.prototype.get_ShowParaMarks;
  editor.sync_ShowParaMarks = Asc.asc_docs_api.prototype.sync_ShowParaMarks;
  editor.sync_DialogAddHyperlink = Asc.asc_docs_api.prototype.sync_DialogAddHyperlink;
  editor.FontSizeOut = Asc.asc_docs_api.prototype.FontSizeOut;
  editor.FontSizeIn = Asc.asc_docs_api.prototype.FontSizeIn;

  AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback) {
    callback({'saveLock': false});
  };
  
  AscTest.CreateLogicDocument();
  editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState = function () {};
  
  let oGlobalShape;
  function executeTestWithParams(fCallback, sTextIntoShape, bGlobalShape) {
    const oLogicDocument = AscTest.CreateLogicDocument();
    if (!oGlobalShape || oGlobalShape.bDeleted || !bGlobalShape) {
      oGlobalShape = AscTest.createShape(oLogicDocument.Slides[0]);
    }

    oGlobalShape.setTxBody(AscFormat.CreateTextBodyFromString(sTextIntoShape, editor.WordControl.m_oDrawingDocument, oGlobalShape));
    const oContent = oGlobalShape.txBody.content;
    const oParagraph = oContent.Content[0];
    oParagraph.SetThisElementCurrent();
    
    fCallback(oLogicDocument, oGlobalShape, oParagraph);
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
    
    oLogicDocument.OnKeyDown(oEvent);
    
    const bCheck = oLogicDocument.Slides.length === 2 && oLogicDocument.Slides[0].cSld.spTree.length === oLogicDocument.Slides[1].cSld.spTree;
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
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
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
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Bold(), 'Check bold shortcut');
    }, 'Hello World', true);
  }

  function checkCenterAlign(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Center, 'Check center align shortcut');
    }, 'Hello World', true);
  }
  
  function checkEuroSign(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
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
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Italic(), 'Check italic shortcut');
    }, 'Hello World', true);
  }
  
  function checkJustifyAlign(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Justify, 'check justify align shortcut');
    }, 'Hello World', true);
  }
  
  function checkAddHyperlink(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
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
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      
      oLogicDocument.OnKeyDown(oEvent);
      const oBullet = oParagraph.Get_PresentationNumbering();
      oAssert.true(oBullet.m_nType === AscFormat.numbering_presentationnumfrmt_Char, 'Check bullet list shortcut');
    }, 'Hello World', true);
  }
  
  function checkLeftAlign(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Left, 'Check left align shortcut');
    }, 'Hello World', true);
  }
  
  function checkRightAlign(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.strictEqual(oParagraph.GetParagraphAlign(), AscCommon.align_Right, 'Check right align shortcut');
    }, 'Hello World', true);
  }
  
  function checkUnderline(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Underline(), 'Check underline shortcut');
    }, 'Hello World', true);
  }
  
  function checkStrikethrough(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.true(oRun.Get_Strikeout(), 'Check strikeout shortcut');
    }, 'Hello World', true);
  }
  
  let oCopyParagraphTextPr;
  function checkCopyFormat(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
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
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      oLogicDocument.SelectAll();
      
      oLogicDocument.OnKeyDown(oEvent);
      oAssert.deepEqual(oParagraph.GetCalculatedTextPr(), oCopyParagraphTextPr, 'check paste format shortcut');
    }, 'Hello World', true);
  }
  
  function checkSuperscript(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      const oTextPr = oParagraph.GetCalculatedTextPr();
      oAssert.strictEqual(oTextPr.VertAlign, AscCommon.vertalign_SuperScript, 'Check superscript shortcut');
    }, 'Hello World', true);
  }
  
  function checkSubscript(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      
      oLogicDocument.OnKeyDown(oEvent);
      const oTextPr = oParagraph.GetCalculatedTextPr();
      oAssert.strictEqual(oTextPr.VertAlign, AscCommon.vertalign_SubScript, 'Check subscript shortcut');
    }, 'Hello World', true);
  }
  
  function checkEnDash(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      let bCheck = false;
      oParagraph.SetThisElementCurrent();
      oParagraph.MoveCursorToEndPos();
      
      oLogicDocument.OnKeyDown(oEvent);
      const sText = AscTest.GetParagraphText(oParagraph);
      oAssert.strictEqual(sText.charCodeAt(sText.length - 1), 0x2013, 'Check en dash shortcut');
    }, 'Hello World', true);
  }
  
  function checkDecreaseFont(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      oRun.SetFontSize(10);
      oLogicDocument.SelectAll();
      oLogicDocument.OnKeyDown(oEvent);

      oAssert.strictEqual(oRun.Get_FontSize(), 9, 'Check decrease font size shortcut');
    }, 'Hello World', true);
  }
  
  function checkIncreaseFont(oEvent, oAssert) {
    executeTestWithParams(function (oLogicDocument, oShape, oParagraph) {
      oParagraph.SetThisElementCurrent();
      const oRun = oParagraph.Content[0];
      oRun.SetFontSize(10);
      oLogicDocument.SelectAll();
      oLogicDocument.OnKeyDown(oEvent);
      
      oAssert.strictEqual(oRun.Get_FontSize(), 11, 'Check increase font size shortcut');
    }, 'Hello World', true);
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
      
      oEvent = createEvent( 56, true, true, false,false,false);
      checkShowParaMarks(oEvent, oAssert);
      
      oEvent = createEvent( 66, true, false, false,false,false);
      checkBold(oEvent, oAssert);
      
      oEvent = createEvent( 67, true, true, false,false,false);
      checkCopyFormat(oEvent, oAssert);
      
      oEvent = createEvent( 69, true, false, false,false,false);
      checkCenterAlign(oEvent, oAssert);
      
      oEvent = createEvent( 69, true, false, true,false,false);
      checkEuroSign(oEvent, oAssert);
      
      // group shapes
      oEvent = createEvent( 71, true, false, false,false,false);
      checkGroup(oEvent, oAssert);
      
      // then ungroup
      oEvent = createEvent( 71, true, true, false,false,false);
      checkUnGroup(oEvent, oAssert);
      
      oEvent = createEvent( 73, true, false, false,false,false);
      checkItalic(oEvent, oAssert);
      
      oEvent = createEvent( 74, true, false, false,false,false);
      checkJustifyAlign(oEvent, oAssert);
      
      oEvent = createEvent( 75, true, false, false,false,false);
      checkAddHyperlink(oEvent, oAssert);
      
      oEvent = createEvent( 76, true, true, false,false,false);
      checkBulletList(oEvent, oAssert);
      
      oEvent = createEvent( 76, true, false, false,false,false);
      checkLeftAlign(oEvent, oAssert);
      
      oEvent = createEvent( 82, true, false, false,false,false);
      checkRightAlign(oEvent, oAssert);
      
      oEvent = createEvent( 85, true, false, false,false,false);
      checkUnderline(oEvent, oAssert);
      
      oEvent = createEvent( 53, true, false, false,false,false);
      checkStrikethrough(oEvent, oAssert);
      
      oEvent = createEvent( 83, true, true, false,false,false);
      checkPasteFormat(oEvent, oAssert);
      
      oEvent = createEvent( 187, true, true, false,false,false);
      checkSuperscript(oEvent, oAssert);
      
      oEvent = createEvent( 188, true, false, false,false,false);
      checkSuperscript(oEvent, oAssert);
      
      oEvent = createEvent( 187, true, false, false,false,false);
      checkSubscript(oEvent, oAssert);
      
      oEvent = createEvent( 190, true, false, false,false,false);
      checkSubscript(oEvent, oAssert);
      
      oEvent = createEvent( 189, true, true, false,false,false);
      checkEnDash(oEvent, oAssert);
      
      oEvent = createEvent( 219, true, false, false,false,false);
      checkDecreaseFont(oEvent, oAssert);
      
      oEvent = createEvent( 221, true, false, false,false,false);
      checkIncreaseFont(oEvent, oAssert);
    });
  });

})(window);
