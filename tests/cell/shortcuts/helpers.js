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
	let bIsCellEditorOpened = false;
	function setCheckOpenCellEditor(oPr)
	{
		bIsCellEditorOpened = oPr;
	}
	function checkOpenCellEditor()
	{
		return bIsCellEditorOpened;
	}
	window.AscFonts = AscFonts || {};
	AscCommonExcel.WorkbookView.prototype.sendCursor = function ()
	{

	}
	const fOldCellEditor = AscCommonExcel.CellEditor.prototype.open;
	AscCommonExcel.CellEditor.prototype.open = function (options)
	{
		options.getSides = function ()
		{
			return {l: [0], r: [100], b: [10], cellX: 0, cellY: 0, ri: 0, bi: 0};
		}
		fOldCellEditor.call(this, options);
	};
	Asc.DrawingContext.prototype.measureChar = function ()
	{
		return Asc.TextMetrics(5, 9, 10, 1, 1, 10, 5);
	}
	window.setTimeout = function (callback)
	{
		callback();
	}
	AscFonts.CFontManager.prototype.MeasureChar = function ()
	{
		return {fAdvanceX: 5, oBBox: {fMaxX: 0, fMinX: 0}};
	};
	delete AscCommon.EncryptionWorker;
	AscCommon.ZLib = function ()
	{
		this.open = function ()
		{
			return false;
		}
	};

	Asc.DrawingContext.prototype.setFont = function ()
	{
	};
	Asc.DrawingContext.prototype.fillText = function ()
	{
	};
	Asc.DrawingContext.prototype.getFontMetrics = function ()
	{
		return {ascender: 15, descender: 4, lineGap: 1, nat_scale: 1000, nat_y1: 1000, nat_y2: -1000};
	};
// AscCommonExcel.StringRender.prototype.measureString = function (fragments, flags, maxWidth) {
//     return new Asc.TextMetrics(fragments.length * 5, 20, 0,15,0);
// }

	const editor = new Asc.spreadsheet_api({'id-view': 'editor_sdk', 'id-input': 'ce-cell-content'});
	editor.FontLoader.LoadDocumentFonts = function ()
	{
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
	editor.asc_setZoom(1);

	function createEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsAltGr, bIsMacCmdKey)
	{
		bIsPrevent = false;
		bIsStopPropogation = false;
		const oKeyBoardEvent = {
			preventDefault : function ()
			{
				this.isDefaultPrevented = true;
			},
			stopPropagation: function ()
			{
				this.isPropagationStopped = true;
			}
		};
		oKeyBoardEvent.isDefaultPrevented = false;
		oKeyBoardEvent.isPropagationStopped = false;
		oKeyBoardEvent.which = nKeyCode;
		oKeyBoardEvent.keyCode = nKeyCode;
		oKeyBoardEvent.shiftKey = bIsShift;
		oKeyBoardEvent.altKey = bIsAlt;
		oKeyBoardEvent.ctrlKey = bIsCtrl;
		oKeyBoardEvent.metaKey = bIsMacCmdKey;
		oKeyBoardEvent.altGr = bIsAltGr;
		return oKeyBoardEvent;
	}

	function wbModel()
	{
		return editor.wbModel;
	}

	function wbView()
	{
		return editor.wb;
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

	function getFragments(start, length)
	{
		return cellEditor()._getFragments(start, length);
	}

	function getSelectionCellEditor()
	{
		return cellEditor().copySelection().map((e) => e.getFragmentText()).join('');
	}

	function moveToStartCellEditor()
	{
		cellEditor()._moveCursor(-2);
	}

	function moveToEndCellEditor()
	{
		cellEditor()._moveCursor(-4);
	}

	function moveRight()
	{
		wbView()._onChangeSelection(true, 1, 0, false, false);
	}

	function moveToCell(nRow, nCol)
	{
		const nCurrentCell = activeCell();
		wbView()._onChangeSelection(true, nCol - nCurrentCell.c1, nRow - nCurrentCell.r1, false, false);
	}

	function selectToCell(nRow, nCol)
	{
		const nCurrentCell = activeCell();
		wbView()._onChangeSelection(false, nCol - nCurrentCell.c1, nRow - nCurrentCell.r1, false, false);
	}

	function onKeyDown(oEvent)
	{
		if (oEvent instanceof Object)
		{
			editor.onKeyDown(oEvent);
		} else
		{
			const oRetEvent = createEvent.apply(null, arguments);
			editor.onKeyDown(oRetEvent);
			return oRetEvent;
		}
	}

	function remove()
	{
		editor.asc_Remove();
	}

	function closeCellEditor(bSkip)
	{
		wbView().closeCellEditor();
	}

	function enterTextWithoutClose(sString)
	{
		wbView().EnterText(sString.split('').map((e) => e.charCodeAt(0)));
		setCheckOpenCellEditor(true);
	}

	function enterText(sString)
	{
		enterTextWithoutClose(sString);
		closeCellEditor();
	}

	function cellEditor()
	{
		return wbView().cellEditor;
	}

	function getCellText()
	{
		closeCellEditor(true);
		return activeCellRange().getValueWithFormat();
	}

	function getCellTextWithoutFormat()
	{
		closeCellEditor();
		return activeCellRange().getValueWithoutFormat();
	}

	function moveDown()
	{
		wbView()._onChangeSelection(true, 0, 1, false, false);
	}

	function wsView()
	{
		return wbView().getWorksheet();
	}

	function ws()
	{
		return wsView().model;
	}

	function moveAndEnterText(sText, nRow, nCol)
	{
		moveToCell(nRow, nCol);
		enterText(sText);
	}

	function createTest(oAssert)
	{
		const deep = (result, expected, sPrompt) => oAssert.deepEqual(result, expected, sPrompt);
		const equal = (result, expected, sPrompt) => oAssert.strictEqual(result, expected, sPrompt);
		return {deep, equal};
	}

	function moveAndGetCellText(nRow, nCol)
	{

		moveToCell(nRow, nCol);
		return getCellText();
	}

	function goToSheet(i)
	{
		wbView().showWorksheet(i);
	}

	let id = 0;

	function createWorksheet()
	{
		const sName = 'name' + id;
		editor.asc_addWorksheet(sName);
		id += 1;
		return sName;
	}

	function removeCurrentWorksheet()
	{
		editor.asc_deleteWorksheet();
	}

	function cleanCell(oRange)
	{
		return {r: oRange.r1, c: oRange.c1};
	}

	function cleanRange(oRange)
	{
		return {r1: oRange.r1, r2: oRange.r2, c1: oRange.c1, c2: oRange.c2};
	}

	function cleanSelection()
	{
		return cleanRange(selectionRange());
	}

	function cleanActiveCell()
	{
		return cleanCell(activeCell());
	}

	function checkRange(nRow1, nRow2, nCol1, nCol2)
	{
		return cleanRange(Asc.Range(nCol1, nRow1, nCol2, nRow2, true));
	}

	function openCellEditor()
	{
		var enterOptions = new AscCommonExcel.CEditorEnterOptions();
		enterOptions.newText = '';
		enterOptions.quickInput = true;
		enterOptions.focus = true;
		handlers().trigger('editCell', enterOptions);
		setCheckOpenCellEditor(true);
	}

	function checkActiveCell(nRow, nCol)
	{
		return cleanActiveCell(new Asc.Range(nCol, nRow, nCol, nRow, true));
	}

	function cleanCache()
	{
		wsView()._cleanCache();
	}

	function selectAll()
	{
		wbView().selectAll();
	}

	function cleanAll()
	{
		selectAll();
		handlers().trigger("empty");
		cleanCache();
		wsView().changeZoomResize();
		moveToCell(0, 0);
	}

	function setCellFormat(nFormat)
	{
		handlers().trigger('setCellFormat', nFormat);


	}

	function selectionInfo()
	{
		return wbView().getSelectionInfo();
	}

	function xfs()
	{
		return selectionInfo().asc_getXfs();
	}

	function undo()
	{
		handlers().trigger('undo');
	}

	function selectAllCell()
	{
		cellEditor()._moveCursor(-2);
		cellEditor()._selectChars(-4);
	}

	function cellPosition()
	{
		return cellEditor().cursorPos;
	}

	function getCellEditMode()
	{
		return wsView().getCellEditMode();
	}
	function testPreventDefaultAndStopPropagation(oEvent, oAssert, bInvert)
	{
		onKeyDown(oEvent);
		oAssert.true(!!bInvert ? !oEvent.isDefaultPrevented : oEvent.isDefaultPrevented);
		oAssert.true(!!bInvert ? !oEvent.isPropagationStopped : oEvent.isPropagationStopped);
	}

	function controller()
	{
		return editor.wb.controller;
	}

	function handlers()
	{
		return controller().handlers;
	}

	function activeCell()
	{
		return wsView().getActiveCell();
	}

	function selectionRange()
	{
		return wsView().getSelectedRange().bbox;
	}

	function activeCellRange()
	{
		return ws().getRange3(activeCell().r1, activeCell().c1, activeCell().r2, activeCell().c2);
	}

	window.AscTestShortcut = {
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
		onKeyDown,
		remove,
		closeCellEditor,
		enterTextWithoutClose,
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
		checkOpenCellEditor,
		setCheckOpenCellEditor,
		activeCellRange
	};
})(window);
