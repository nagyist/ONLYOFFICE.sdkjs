

window.AscFonts = AscFonts || {};
AscFonts.CFontManager.prototype.MeasureChar = function () {
    return {fAdvanceX: 5, oBBox: {fMaxX: 0, fMinX: 0}}
};
delete AscCommon.EncryptionWorker;
AscCommon.ZLib = function () {
    this.open = function () {
        return false;
    }
};

Asc.DrawingContext.prototype.setFont = function () {};
Asc.DrawingContext.prototype.fillText = function () {};
Asc.DrawingContext.prototype.getFontMetrics = function () {
    return {ascender:15, descender:4, lineGap: 1,nat_scale: 1000,nat_y1: 1000,nat_y2:-1000};
};
// AscCommonExcel.StringRender.prototype.measureString = function (fragments, flags, maxWidth) {
//     return new Asc.TextMetrics(fragments.length * 5, 20, 0,15,0);
// }

window.editor = new Asc.spreadsheet_api({'id-view': 'editor_sdk', 'id-input': 'ce-cell-content'});




editor.FontLoader.LoadDocumentFonts = function () {
    editor.ServerIdWaitComplete = true;
    editor._coAuthoringInitEnd();
    editor.asyncFontsDocumentEndLoaded();
}



const oOleObjectInfo = {"binary": AscCommon.getEmpty()};
const sStream = oOleObjectInfo["binary"];
const oFile = new AscCommon.OpenFileResult();
oFile.bSerFormat = AscCommon.checkStreamSignature(sStream, AscCommon.c_oSerFormat.Signature);
oFile.data = sStream;

editor.openDocument(oFile);


window.Asc.editor = window.editor;
let bIsPrevent = false;
let bIsStopPropogation
function createEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsAltGr, bIsMacCmdKey) {
    bIsPrevent = false;
    bIsStopPropogation = false;
    const oKeyBoardEvent = {
        preventDefault: function () {
            bIsPrevent = true;
        },
        stopPropagation: function () {
            bIsStopPropogation = true;
        }
    };
    oKeyBoardEvent.which = nKeyCode;
    oKeyBoardEvent.shiftKey = bIsShift;
    oKeyBoardEvent.altKey = bIsAlt;
    oKeyBoardEvent.ctrlKey = bIsCtrl;
    oKeyBoardEvent.metaKey = bIsMacCmdKey;
    oKeyBoardEvent.altGr = bIsAltGr;
    return oKeyBoardEvent;
}

$(
function () {
QUnit.module('test shortcuts');
QUnit.test('Test common shortcuts', function (oAssert) {
    
    oAssert.strictEqual(true, true);
    let oEvent;
    oEvent = createEvent(9, false, false, false, false, false);
    editor.onKeyDown(oEvent);
    const activeCell = editor.wb.controller.handlers.trigger('getActiveCell');
    oAssert.deepEqual(activeCell, new Asc.Range(1, 0, 1, 0));
    
    
    //check events controller
    checkRefreshConnections
    checkChangeFormatTableInfo
    checkCalculateAll
    checkCalculateWorkbook
    checkCalculateActiveSheet
    checkCalculateOnlyChanged
    checkFocusOnEditor
    checkAddDateOrTime
    checkRemoveCellText
    checkEmpty
    checkMoveToLeftCell
    checkMoveToRightCell
    checkMoveToUpCell
    checkMoveToDownCell
    checkReset
    checkDisableNumlock
    checkDisableScrollLock
    checkSelectColumn
    checkSelectRow
    checkSelectSheet
    checkAddSpace
    checkGoToPreviousSheet
    checkMoveToTopCell
    checkMoveToNextSheet
    checkMoveToBottomCell
    checkMoveToLeftEdgeCell
    checkMoveToLeftCell
    checkMoveToRightEdgeCell
    checkMoveToRightCell
    checkMoveToTopCell
    checkMoveToUpCell
    checkMoveToBottomCell
    checkMoveToDownCell
    checkMoveToLeftEdgeCell
    checkMoveToLeftEdgeTopCell
    checkMoveToRightEdgeCell
    checkMoveToRightEdgeBottomCell
    checkSetNumberFormat
    checkSetTimeFormat
    checkSetDateTime
    checkSetCurrencyFormat
    checkStrikethrough
    checkSetExponentialFormat
    checkBold
    checkItalic
    checkSave
    checkUnderline
    checkSetGeneralFormat
    checkRedo
    checkUndo
    checkSelectSheet
    checkPrint
    checkAddSumFunction
    checkContextMenu
    
    //ccheckChangeSelectionInOleEditor
    
    //check cell editor shortcuts
    checkCloseCellEditor
    checkAddNewLine
    checkTryCloseEditor
    checkTryCloseEditor
    checkRemoveCharBack
    checkRemoveWordBack
    checkAddSpace
    checkMoveToEndLine
    checkMoveToEndPos
    checkSelectToEndLine
    checkSelectToEndPos
    checkMoveToStartLine
    checkMoveToStartPos
    checkSelectToStartLine
    checkSelectToStartPos
    checkMoveToLeftChar
    checkMoveToLeftWord
    checkSelectToLeftChar
    checkSelectToLeftWord
    checkMoveToUpLine
    checkSelectToUpLine
    checkMoveToRightChar
    checkMoveToRightWord
    checkSelectToRightChar
    checkSelectToRightWord
    checkMoveToDownLine
    checkSelectToDownLine
    checkRemoveFrontChar
    checkRemoveFrontWord
    checkStrikethroughCellEditor
    checkSelectAllCellEditor
    checkBoldCellEditor
    checkItalicCellEditor
    checkUnderlineCellEditor
    checkDisableScrollLockCellEditor
    checkDisableNumLockCellEditor
    checkPrintCellEditor
    checkUndoCellEditor
    checkRedoCellEditor
    checkDisableF2InOperaCellEditor
    checkSwitchReference
    checkAddTimeCellEditor
    checkAddDateCellEditor

    // check common controllers shortcuts

    checkRemoveShape
    checkAddTab
    checkSelectNextObject
    checkSelectPreviousObject
    checkVisitHyperlink
    checkAddNewLineMath
    checkAddBreakLine
    checkAddNewLine

    checkCreateTxBoxContentShape
    checkCreateTxBodyContentShape
    checkMoveCursorToStartPosShape
    checkSelectAllContentShape
    checkMoveCursorToStartPosChartTitle
    checkSelectAllContentChartTitle
    checkRemoveSelectionCellGraphicFrame
    checkMoveCursorToStartPosAndSelectAllTable

    checkStepRemoveSelection
    checkResetAddShape
    moveCursorToEndPositionContent
    moveCursorToEndLineContent
    moveCursorToStartPositionContent
    moveCursorToStartLineContent

    checkMoveCursorLeftCharContentGraphicFrame
    checkMoveCursorLeftWordContentGraphicFrame
    checkSelectLeftCharContentGraphicFrame
    checkSelectLeftWordContentGraphicFrame

    checkMoveCursorLeftCharContentShape
    checkMoveCursorLeftWordContentShape
    checkSelectLeftCharContentShape
    checkSelectLeftWordContentShape

    checkSmallMoveLeftShape
    checkBigMoveLeftShape

    checkMoveCursorRightCharContentGraphicFrame
    checkMoveCursorRightWordContentGraphicFrame
    checkSelectRightCharContentGraphicFrame
    checkSelectRightWordContentGraphicFrame

    checkMoveCursorRightCharContentShape
    checkMoveCursorRightWordContentShape
    checkSelectRightCharContentShape
    checkSelectRightWordContentShape

    checkSmallMoveRightShape
    checkBigMoveRightShape

    checkMoveCursorUpCharContentGraphicFrame
    checkSelectUpCharContentGraphicFrame

    checkMoveCursorUpCharContentShape
    checkSelectUpCharContentShape

    checkSmallMoveUpShape
    checkBigMoveUpShape

    checkMoveCursorDownCharContentGraphicFrame
    checkSelectDownCharContentGraphicFrame

    checkMoveCursorDownCharContentShape
    checkSelectDownCharContentShape

    checkSmallMoveDownShape
    checkBigMoveDownShape

    checkRemoveShape
    checkSelectAllShape
    checkBoldShape
    checkClearSlicer
    checkSetCenterAlign
    checkSetItalicShape
    checkSetJustifyAlign
    checkSetLeftAlign
    checkSetRightAlign
    checkSlicerMultiSelect
    checkUnderlineShape
    checkSubscriptShape
    checkSuperscriptShape
    checkSuperscriptShape
    checkAddEnDash
    checkAddLowLine
    checkAddHyphenMinus
    checkSubscriptShape
    checkDecreaseFontSize
    checkIncreaseFontSize
});
}
)