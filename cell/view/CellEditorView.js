/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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

"use strict";
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {


	/*
	 * Import
	 * -----------------------------------------------------------------------------
	 */
	var asc = window["Asc"];

	var AscBrowser = AscCommon.AscBrowser;

	var cElementType = AscCommonExcel.cElementType;
	var c_oAscCellEditorSelectState = AscCommonExcel.c_oAscCellEditorSelectState;
	var c_oAscCellEditorState = asc.c_oAscCellEditorState;
	var Fragment = AscCommonExcel.Fragment;

	var asc_getcvt = asc.getCvtRatio;
	var asc_round = asc.round;
	var asc_search = asc.search;
	var asc_lastidx = asc.lastIndexOf;

	var asc_HL = AscCommonExcel.HandlersList;
	var asc_incDecFonSize = asc.incDecFonSize;


	/** @const */
	var kBeginOfLine = -1;
	/** @const */
	var kBeginOfText = -2;
	/** @const */
	var kEndOfLine = -3;
	/** @const */
	var kEndOfText = -4;
	/** @const */
	var kNextChar = -5;
	/** @const */
	var kNextWord = -6;
	/** @const */
	var kNextLine = -7;
	/** @const */
	var kPrevChar = -8;
	/** @const */
	var kPrevWord = -9;
	/** @const */
	var kPrevLine = -10;
	/** @const */
	var kPosition = -11;
	/** @const */
	var kPositionLength = -12;

	/** @const */
	var codeNewLine = 0x00A;
	var codeEqually = 0x3D;
	var codeUnarPlus = 0x2B;
	var codeUnarMinus = 0x2D;


	/**
	 * CellEditor widget
	 * -----------------------------------------------------------------------------
	 * @constructor
	 * @param {Element} elem
	 * @param {Element} input
	 * @param {Array} fmgrGraphics
	 * @param {AscCommonExcel.Font} oFont
	 * @param {HandlersList} handlers
	 * @param {Number} padding
	 * @param {Boolean} menuEditor
	 */
	function CellEditor(elem, input, fmgrGraphics, oFont, handlers, padding, menuEditor) {
		this.element = elem;
		this.input = input;
		this.handlers = new asc_HL(handlers);
		this.options = {};
		this.sides = undefined;
		this.menuEditor = menuEditor;

		//---declaration---
		this.canvasOuter = undefined;
		this.canvasOuterStyle = undefined;
		this.canvas = undefined;
		this.canvasOverlay = undefined;
		this.cursor = undefined;
		this.cursorStyle = undefined;
		this.cursorTID = undefined;
		this.cursorPos = 0;
		this.beginCompositePos = -1;
		this.compositeLength = 0;
		this.topLineIndex = 0;
		this.m_oFont = oFont;
		this.fmgrGraphics = fmgrGraphics;
		this.drawingCtx = undefined;
		this.overlayCtx = undefined;
		this.textRender = undefined;
		this.textFlags = undefined;
		this.kx = 1;
		this.ky = 1;
		this.skipKeyPress = undefined;
		this.undoList = [];
		this.redoList = [];
		this.undoMode = false;
		this.noUpdateMode = false;
		this.selectionBegin = -1;
		this.selectionEnd = -1;
		this.isSelectMode = c_oAscCellEditorSelectState.no;
		this.hasCursor = false;
		this.hasFocus = false;
		this.newTextFormat = null;
		this.selectionTimer = undefined;
		this.enableKeyEvents = true;
		this.isTopLineActive = false;
		this.openFromTopLine = false;
		this.skipTLUpdate = true;
		this.loadFonts = false;
		this.isOpened = false;
		this.callTopLineMouseup = false;
		this.m_nEditorState = c_oAscCellEditorState.editEnd; // Editor's status

		// Features that we will disable
		this.fKeyMouseUp = null;
		this.fKeyMouseMove = null;
		//-----------------

		this.objAutoComplete = new Map();
		this.sAutoComplete = null;
		this.eventListeners = [];

		/** @type RegExp */
		this.rangeChars = ["=", "-", "+", "*", "/", "(", "{", "<", ">", "^", "!", "&", ":", " ", "."];
		this.reNotFormula = new XRegExp("[^\\p{L}\\\\_\\#\\]\\[\\p{N}\\.\"\@]", "i");
		this.reFormula = new XRegExp("^([\\p{L}\\\\_\\]\\[][\\p{L}\\\\_\\#\\]\\[\\p{N}\\.@]*)", "i");

		this.defaults = {
			padding: padding,
			selectColor: new AscCommon.CColor(190, 190, 255, 0.5),
			canvasZIndex: 500,
			blinkInterval: 500,
			cursorShape: "text"
		};

		this._formula = null;
		this._parseResult = null;
		this.needFindFirstFunction = null;
		this.lastRangePos = null;
		this.lastRangeLength = null;

		// Click handler
		this.clickCounter = new AscFormat.ClickCounter();

		//temporary - for safari rendering. remove after fixed
		this._originalCanvasWidth = null;

		this._init();

		return this;
	}

	CellEditor.prototype._init = function () {
		var t = this;
		var z = t.defaults.canvasZIndex;
		this.sAutoComplete = null;

		if (null != this.element) {
			var ceMenuEditor = this.getMenuEditorMode() ? '-menu' : ''
			var ceCanvasOuterId = "ce-canvas-outer" + ceMenuEditor;
			var ceCanvasId = "ce-canvas" + ceMenuEditor;
			var ceCanvasOverlay = "ce-canvas-overlay" + ceMenuEditor;
			var ceCursor = "ce-cursor" + ceMenuEditor;

			var canvasOuterEl = document.getElementById(ceCanvasOuterId);
			if (canvasOuterEl) {
				var parentNode = canvasOuterEl.parentNode;
				parentNode.removeChild(canvasOuterEl);
			}
			t.canvasOuter = document.createElement('div');
			t.canvasOuter.id = ceCanvasOuterId;
			t.canvasOuter.style.position = "absolute";
			t.canvasOuter.style.display = "none";
			t.canvasOuter.style.zIndex = z;
			var innerHTML = '<canvas id=' + ceCanvasId + ' style="z-index: ' + (z + 1) + '"></canvas>';
			innerHTML += '<canvas id=' + ceCanvasOverlay + ' style="z-index: ' + (z + 2) + '; cursor: ' + t.defaults.cursorShape +
				'"></canvas>';
			innerHTML += '<div id=' + ceCursor + ' style="display: none; z-index: ' + (z + 3) + '"></div>';
			t.canvasOuter.innerHTML = innerHTML;
			this.element.appendChild(t.canvasOuter);

			t.canvasOuterStyle = t.canvasOuter.style;
			t.canvas = document.getElementById(ceCanvasId);
			t.canvasOverlay = document.getElementById(ceCanvasOverlay);
			t.cursor = document.getElementById(ceCursor);
			t.cursorStyle = t.cursor.style;
		}

		// create text render
		t.drawingCtx = new asc.DrawingContext({
			canvas: t.canvas, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
		});
		t.overlayCtx = new asc.DrawingContext({
			canvas: t.canvasOverlay, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
		});
		t.textRender = new AscCommonExcel.CellTextRender(t.drawingCtx);

		// bind event handlers
		if (t.canvasOuter && t.canvasOuter.addEventListener) {
			var eventInfo = new AscCommon.CEventListenerInfo(t.canvasOuter, AscCommon.getPtrEvtName("down"), function () {
				return t._onMouseDown.apply(t, arguments);
			}, false);
			t.eventListeners.push(eventInfo);

			eventInfo = new AscCommon.CEventListenerInfo(t.canvasOuter, AscCommon.getPtrEvtName("up"), function () {
				return t._onMouseUp.apply(t, arguments);
			}, false);
			t.eventListeners.push(eventInfo);

			eventInfo = new AscCommon.CEventListenerInfo(t.canvasOuter, AscCommon.getPtrEvtName("move"), function () {
				return t._onMouseMove.apply(t, arguments);
			}, false);
			t.eventListeners.push(eventInfo);

			eventInfo = new AscCommon.CEventListenerInfo(t.canvasOuter, AscCommon.getPtrEvtName("leave"), function () {
				return t._onMouseLeave.apply(t, arguments);
			}, false);
			t.eventListeners.push(eventInfo);
		}

		// check input, it may have zero len, for mobile version
		if (t.input && t.input.addEventListener) {
			eventInfo = new AscCommon.CEventListenerInfo(t.input, "focus", function () {
				return t.isOpened ? t._topLineGotFocus.apply(t, arguments) : true;
			}, false);
			t.eventListeners.push(eventInfo);

			eventInfo = new AscCommon.CEventListenerInfo(t.input, AscCommon.getPtrEvtName("down"), function () {
				return t.isOpened ? (t.callTopLineMouseup = true) : true;
			}, false);
			t.eventListeners.push(eventInfo);

			eventInfo = new AscCommon.CEventListenerInfo(t.input, AscCommon.getPtrEvtName("up"), function () {
				return t.isOpened ? t._topLineMouseUp.apply(t, arguments) : true;
			}, false);
			t.eventListeners.push(eventInfo);

			eventInfo = new AscCommon.CEventListenerInfo(t.input, "input", function () {
				return t._onInputTextArea.apply(t, arguments);
			}, false);
			t.eventListeners.push(eventInfo);

			// We do not support drop to the top line
			eventInfo = new AscCommon.CEventListenerInfo(t.input, "drop", function (e) {
				e.preventDefault();
				return false;
			}, false);
			t.eventListeners.push(eventInfo);
		}

		this.fKeyMouseUp = function () {
			return t._onWindowMouseUp.apply(t, arguments);
		};
		this.fKeyMouseMove = function () {
			return t._onWindowMouseMove.apply(t, arguments);
		};
		t.addEventListeners();
	};

	CellEditor.prototype.destroy = function () {
	};

	CellEditor.prototype.removeEventListeners = function () {
		this.eventListeners.forEach(function (eventInfo) {
			eventInfo.listeningElement.removeEventListener(eventInfo.eventName, eventInfo.listener);
		});
	};

	CellEditor.prototype.addEventListeners = function () {
		this.eventListeners.forEach(function (eventInfo) {
			eventInfo.listeningElement.addEventListener(eventInfo.eventName, eventInfo.listener, eventInfo.useCapture);
		});
	};

	/**
	 * @param {Object} options
	 *   fragments  - text fragments
	 *   flags      - text flags (wrapText, textAlign)
	 *   font
	 *   background
	 *   saveValueCallback
	 */
	CellEditor.prototype.open = function (options) {
		this._setEditorState(c_oAscCellEditorState.editStart);
		
		var b = this.input.selectionStart;

		this.isOpened = true;
		if (window.addEventListener) {
			window.addEventListener(AscCommon.getPtrEvtName("up"), this.fKeyMouseUp, false);
			window.addEventListener(AscCommon.getPtrEvtName("move"), this.fKeyMouseMove, false);
		}
		this._setOptions(options);
		this._cleanLastRangeInfo();
		this._updateTopLineActive(true === this.input.isFocused, true);

		this._updateEditorState();
		this._draw();

		if (null !== options.enterOptions.newText) {
			this._selectChars(kEndOfText);
			this._addChars(options.enterOptions.newText);
		}

		if (this.isTopLineActive && typeof b !== "undefined") {
			if (this.cursorPos !== b) {
				this._moveCursor(kPosition, b);
			}
		} else if (options.enterOptions.cursorPos) {
			this._moveCursor(kPosition, options.enterOptions.cursorPos);
		} else if (options.enterOptions.eventPos) {
			this._onMouseDown(options.enterOptions.eventPos);
			this._onMouseUp(options.enterOptions.eventPos);
		} else {
			this._moveCursor(kEndOfText);
		}

		/*
			* Set focus when opening
			* When clicking a symbol, do not set focus
			* When F2 sets focus in the editor
			* When dbl clicking, set focus depending on the presence of text in the cell
		 */
		this.setFocus(this.isTopLineActive ? true : (null === options.enterOptions.focus) ? this._haveTextInEdit() : options.enterOptions.focus);
		this._updateUndoRedoChanged();

		AscCommon.StartIntervalDrawText(true);
		this.openAction();
	};

	CellEditor.prototype.close = function (saveValue, callback) {
		var opt = this.options;
		var t = this;

		let externalSelectionController = this.handlers.trigger("getExternalSelectionController");
		if (externalSelectionController && externalSelectionController.getExternalFormulaEditMode()) {
			if (!externalSelectionController.supportVisibilityChangeOption) {
				externalSelectionController.sendExternalCloseEditor(saveValue);
				saveValue = false;
			} else {
				callback && callback(false);
				return;
			}
		}

		var api = window["Asc"]["editor"];
		if (api && !api.canUndoRedoByRestrictions()) {
			saveValue = false;
		}

		var localSaveValueCallback = function (isSuccess) {
			if (!isSuccess) {
				t.setFocus(true);
				t.cleanSelectRange();
				if (callback) {
					callback(false);
				}
				return false;
			}

			t.isOpened = false;

			t._formula = null;
			t._parseResult = null;

			if (!window['IS_NATIVE_EDITOR']) {
				if (window.removeEventListener) {
					window.removeEventListener(AscCommon.getPtrEvtName("up"), t.fKeyMouseUp, false);
					window.removeEventListener(AscCommon.getPtrEvtName("move"), t.fKeyMouseMove, false);
				}
				if (api && api.isMobileVersion) {
					t.input.blur();
				}
				t._blur();
				t._updateTopLineActive(false);
				t.input.isFocused = false;
				t._updateCursor();
				// hide
				t._hideCanvas();
			}

			// delete autoComplete
			t.objAutoComplete.clear();

			// Reset editor state
			t._setEditorState(c_oAscCellEditorState.editEnd);
			t.handlers.trigger("closed");
			t.closeAction();

			if (callback) {
				callback(true);
			} else {
				return true;
			}
		};

		if (this.isStartCompositeInput()) {
			this.End_CompositeInput();
		}

		if (saveValue) {
			// We always recalculate for a non-empty cell or if there were changes. http://bugzilla.onlyoffice.com/show_bug.cgi?id=34864
			if (0 < this.undoList.length || 0 < AscCommonExcel.getFragmentsCharCodesLength(this.options.fragments)) {
				var isFormula = this.isFormula();
				// We replace the text with auto-completion if there is a select and the text matches completely.
				if (this.sAutoComplete && !isFormula) {
					this.selectionBegin = this.textRender.getBeginOfText();
					this.cursorPos = this.selectionEnd = this.textRender.getEndOfText();
					this.noUpdateMode = true;
					this._addChars(this.sAutoComplete);
					this.noUpdateMode = false;
				}

				for (var i = 0; i < opt.fragments.length; i++) {
					opt.fragments[i].initText();
				}
				return opt.saveValueCallback(opt.fragments, this.textFlags, localSaveValueCallback);
			}
		}

		this.isOpened = false;

		this._formula = null;
		this._parseResult = null;

		if (!window['IS_NATIVE_EDITOR']) {
			if (window.removeEventListener) {
				window.removeEventListener(AscCommon.getPtrEvtName("up"), this.fKeyMouseUp, false);
				window.removeEventListener(AscCommon.getPtrEvtName("move"), this.fKeyMouseMove, false);
			}
			this._blur();
			this._updateTopLineActive(false);
			this.input.isFocused = false;
			this._updateCursor();
			// hide
			this._hideCanvas();
		}

		// delete autoComplete
		this.objAutoComplete.clear();

		// Reset editor state
		this._setEditorState(c_oAscCellEditorState.editEnd);
		this.handlers.trigger("closed");
		t.closeAction();

		if (callback) {
			callback(true);
		}

		return true;
	};

	CellEditor.prototype._blur = function () {
		this.handlers.trigger("doEditorFocus");
	};

	CellEditor.prototype.setTextStyle = function (prop, val) {
		if (this.isFormula()) {
			return;
		}
		if (!this.options.fragments) {
			return;
		}

		this.startAction();

		var t = this, opt = t.options, begin, end, i, first, last;

		if (t.selectionBegin !== t.selectionEnd) {
			begin = Math.min(t.selectionBegin, t.selectionEnd);
			end = Math.max(t.selectionBegin, t.selectionEnd);

			// save info to undo/redo
			if (end - begin < 2) {
				t.undoList.push({fn: t._addChars, args: [t.textRender.getChars(begin, 1), begin]});
			} else {
				t.undoList.push({fn: t._addFragments, args: [t._getFragments(begin, end - begin), begin]});
			}

			t._extractFragments(begin, end - begin);

			first = t._findFragment(begin);
			last = t._findFragment(end - 1);

			if (first && last) {
				for (i = first.index; i <= last.index; ++i) {
					var valTmp = t._setFormatProperty(opt.fragments[i].format, prop, val);
					// For hotkeys only
					if (null === val) {
						val = valTmp;
					}
				}
				// merge fragments with equal formats
				t._mergeFragments();
				t._update();

				// Refreshing the selection
				t._cleanSelection();
				t._drawSelection();

				// save info to undo/redo
				t.undoList.push({fn: t._removeChars, args: [begin, end - begin]});
				t.redoList = [];
			}

		} else {
			first = t._findFragmentToInsertInto(t.cursorPos);
			if (first) {
				if (!t.newTextFormat) {
					t.newTextFormat = opt.fragments[first.index].format.clone();
				}
				t._setFormatProperty(t.newTextFormat, prop, val);
				t._update();
			}
		}
		this.endAction();
	};

	CellEditor.prototype.changeTextCase = function (val) {
		if (this.isFormula()) {
			return;
		}
		if (!this.options.fragments) {
			return;
		}
		var t = this, opt = t.options;

		if (t.selectionBegin !== t.selectionEnd) {
			let begin = Math.min(t.selectionBegin, t.selectionEnd);
			let end = Math.max(t.selectionBegin, t.selectionEnd);

			let oNewText = AscCommonExcel.changeTextCase(opt.fragments, val, begin, end);
			if (oNewText && oNewText.fragmentsMap) {
				this._changeFragments(oNewText.fragmentsMap);
			}
		}
	};

	CellEditor.prototype._changeFragments = function (fragmentsMap) {
		let opt = this.options;
		if (!opt.fragments) {
			return;
		}
		this.startAction();
		if (fragmentsMap) {
			let _undoFragments = {};
			for (let i in fragmentsMap) {
				if (fragmentsMap.hasOwnProperty(i)) {
					_undoFragments[i] = opt.fragments[i].clone();
					opt.fragments[i] = fragmentsMap[i];
				}
			}

			if (!this.undoMode) {
				// save info to undo/redo
				this.undoList.push({fn: this._changeFragments, args: [_undoFragments]});
			}

			this._update();
			// Refreshing the selection
			this._cleanSelection();
			this._drawSelection();
		}
		this.endAction();
	};

	CellEditor.prototype.empty = function (options) {
		// Clean for editing only All
		if (Asc.c_oAscCleanOptions.All !== options) {
			return;
		}

		// We delete only the selection
		this._removeChars();
	};

	CellEditor.prototype.undo = function () {
		var api = window["Asc"]["editor"];
		if (api && !api.canUndoRedoByRestrictions()) {
			return;
		}
		api.sendEvent("asc_onBeforeUndoRedo");
		this._performAction(this.undoList, this.redoList);
		api.sendEvent("asc_onUndoRedo");
	};

	CellEditor.prototype.redo = function () {
		var api = window["Asc"]["editor"];
		if (api && !api.canUndoRedoByRestrictions()) {
			return;
		}
		api.sendEvent("asc_onBeforeUndoRedo");
		this._performAction(this.redoList, this.undoList);
		api.sendEvent("asc_onUndoRedo");
	};

	CellEditor.prototype.getZoom = function () {
		return this.drawingCtx.getZoom();
	};

	CellEditor.prototype.changeZoom = function (factor) {
		this.drawingCtx.changeZoom(factor);
		this.overlayCtx.changeZoom(factor);
	};

	CellEditor.prototype.canEnterCellRange = function () {
		if (this.lastRangePos !== null || this.handlers.trigger('getWizard')) {
			return true;
		}

		var res = false;
		var isSelection = this.selectionBegin !== this.selectionEnd;
		var curPos = isSelection ? (this.selectionBegin < this.selectionEnd ? this.selectionBegin : this.selectionEnd) : this.cursorPos;
		var prevChar = this.textRender.getChars(curPos - 1, 1);
		if (this.checkSymbolBeforeRange(prevChar)) {
			this.lastRangePos = curPos;
			if (isSelection) {
				this.lastRangeLength = Math.abs(this.selectionEnd - this.selectionBegin);
			}
			res = true;
		}

		return res;
	};

	CellEditor.prototype.checkSymbolBeforeRange = function (char) {
		if (char && !char.trim) {
			char = AscCommon.convertUnicodeToUTF16(char);
		}
		return (this.rangeChars && this.rangeChars.indexOf(char) >= 0) || char === AscCommon.FormulaSeparators.functionArgumentSeparator;
	};

	CellEditor.prototype.changeCellRange = function (range, moveEndOfText) {
		this.skipTLUpdate = false;
		this._moveCursor(kPosition, range.cursorePos);
		this._selectChars(kPositionLength, range.formulaRangeLength);
		this._addChars(range.getName(), undefined, /*isRange*/true);
		if (moveEndOfText) {
			this._moveCursor(kEndOfText);
		}
		this.skipTLUpdate = true;
	};

	CellEditor.prototype.changeCellText = function (str) {
		this.skipTLUpdate = false;
		this._moveCursor(kPosition, this.lastRangePos);
		if (this.lastRangeLength) {
			this._selectChars(kPositionLength, this.lastRangeLength);
		}
		this._addChars(str, undefined, /*isRange*/true);
		this.lastRangeLength = str.length;
		this.skipTLUpdate = true;

		let externalSelectionController = this.handlers.trigger("getExternalSelectionController");
		externalSelectionController && externalSelectionController.sendExternalChangeSelection();
	};

	CellEditor.prototype.insertFormula = function (functionName, isDefName, sRange) {
		this.skipTLUpdate = false;

		// ToDo check selection formula in wizard for delete
		if (this.selectionBegin !== this.selectionEnd) {
			this._removeChars(undefined, undefined, true);
		}

		var addText = '';
		var text = AscCommonExcel.getFragmentsText(this.options.fragments);
		if (!this.isFormula() && 0 === this.cursorPos) {
			addText = '=';
		} else if (functionName && !this.checkSymbolBeforeRange(text[this.cursorPos - 1])) {
			addText = '+';
		}

		if (functionName) {
			addText += functionName;
			if (!isDefName) {
				addText += sRange ? '(' + sRange + ')' : '()';
			}
		}

		if (addText) {
			this._addChars(addText);
			if (functionName && !isDefName) {
				this._moveCursor(kPosition, this.cursorPos - 1);

				// ToDo move this code to moveCursor

				this.lastRangePos = this._parseResult && this._parseResult.argPosArr && this._parseResult.argPosArr.length
					? this._parseResult.argPosArr[0].start
					: this.cursorPos;

				this.lastRangeLength = this._parseResult && this._parseResult.argPosArr && this._parseResult.argPosArr.length
					? this._parseResult.argPosArr[this._parseResult.argPosArr.length - 1].end - this._parseResult.argPosArr[0].start
					: 0;
			}
		}

		this.skipTLUpdate = true;
	};

	CellEditor.prototype.updateWizardMode = function (mode) {
		this._updateCursorStyle(mode ? AscCommon.Cursors.CellCur : this.defaults.cursorShape);
	};

	CellEditor.prototype.move = function () {
		if (!this.isOpened) {
			return;
		}
		if (this.handlers.trigger('isActive') && this.options.checkVisible()) {
			this.textFlags.wrapOnlyCE = false;
			this.sides = this.options.getSides();
			this.left = this.sides.cellX;
			this.top = this.sides.cellY;
			this.right = this.sides.r[this.sides.ri];
			this.bottom = this.sides.b[this.sides.bi];

			this._expand();
			this._adjustCanvas();
			this._showCanvas();
			this._calculateCanvasSize();
			this._renderText();
			this.topLineIndex = 0;
			this._updateCursorPosition();
			this._updateCursor();
			this._drawSelection();
		} else {
			// hide
			this._hideCanvas();
		}
	};

	CellEditor.prototype.setFocus = function (hasFocus) {
		this.hasFocus = !!hasFocus;
		this.handlers.trigger("gotFocus", this.hasFocus);
	};

	CellEditor.prototype.restoreFocus = function () {
		if (this.isTopLineActive) {
			this.input.focus();
		}
	};

	CellEditor.prototype.copySelection = function () {
		var t = this;
		var res = null;
		if (t.selectionBegin !== t.selectionEnd) {
			var start = t.selectionBegin;
			var end = t.selectionEnd;
			if (start > end) {
				var temp = start;
				start = end;
				end = temp;
			}
			res = t._getFragments(start, end - start);
		}
		return res;
	};

	CellEditor.prototype.cutSelection = function () {
		var t = this;
		var f = null;
		if (t.selectionBegin !== t.selectionEnd) {
			var start = t.selectionBegin;
			var end = t.selectionEnd;
			if (start > end) {
				var temp = start;
				start = end;
				end = temp;
			}
			f = t._getFragments(start, end - start);
			t._removeChars();
		}
		return f;
	};

	CellEditor.prototype.pasteText = function (text) {
		text = text.replace(/\t/g, " ");
		text = text.replace(/\r/g, "");
		text = text.replace(/^\n+|\n+$/g, "");

		if (0 === text.length) {
			return;
		}

		this._addChars(text);
	};

	CellEditor.prototype.paste = function (fragments, cursorPos) {
		if (!(fragments.length > 0)) {
			return;
		}
		this.startAction();

		var noUpdateMode = this.noUpdateMode;
		this.noUpdateMode = true;

		if (this.selectionBegin !== this.selectionEnd) {
			this._removeChars();
		}

		// limit count characters
		var length = AscCommonExcel.getFragmentsLength(fragments);
		var excess = this._checkMaxCellLength(length);
		if (excess) {
			length -= excess;
			if (0 === length) {
				this.noUpdateMode = noUpdateMode;
				return false;
			}
			this._extractFragments(0, length, fragments);
		}

		this._cleanFragments(fragments);

		// save info to undo/redo
		this.undoList.push({fn: this._removeChars, args: [this.cursorPos, length]});
		this.redoList = [];

		this.noUpdateMode = noUpdateMode;
		this._addFragments(fragments, this.cursorPos);

		// Made only for inserting a formula into a cell (when the editor is not open)
		if (undefined !== cursorPos) {
			this._moveCursor(kPosition, cursorPos);
		}
		this.endAction();
	};

	/** @param flag {Boolean} */
	CellEditor.prototype.enableKeyEventsHandler = function (flag) {
		var oldValue = this.enableKeyEvents;
		this.enableKeyEvents = !!flag;
		if (this.isOpened && oldValue !== this.enableKeyEvents) {
			this._updateCursor();
		}
	};

	CellEditor.prototype._isFormula = function () {
		let fragments = this.options.fragments;
		if (fragments && fragments.length > 0 && fragments[0].getCharCodesLength() > 0) {
			let firstSymbolCode = fragments[0].getCharCode(0);
			let isEqualSign = firstSymbolCode === codeEqually;
			let unarSign = isEqualSign ? false : (firstSymbolCode === codeUnarPlus || firstSymbolCode === codeUnarMinus);

			return isEqualSign || unarSign;
		}

		return false;
	};
	CellEditor.prototype.isFormula = function () {
		return c_oAscCellEditorState.editFormula === this.m_nEditorState;
	};

	CellEditor.prototype._updateTextAlign = function () {
		this.textFlags.textAlign = (this.options.flags.textAlign === AscCommon.align_Justify || this.isFormula()) ?
			AscCommon.align_Left : this.options.flags.textAlign;
	};

	CellEditor.prototype.replaceText = function (pos, len, newText) {
		this._moveCursor(kPosition, pos);
		this._selectChars(kPosition, pos + len);
		return this._addChars(newText);
	};

	CellEditor.prototype.setFontRenderingMode = function () {
		if (this.isOpened) {
			this._draw();
		}
	};

	CellEditor.prototype.cleanSelectRange = function () {
		this._cleanLastRangeInfo();
		this.handlers.trigger("cleanSelectRange");
		this.handlers.trigger("onSelectionEnd");
	};

	// Private

	CellEditor.prototype._setOptions = function (options) {
		var opt = this.options = options;
		var ctx = this.drawingCtx;
		var u = ctx.getUnits();

		this.textFlags = opt.flags.clone();
		this._updateTextAlign();
		this.textFlags.shrinkToFit = false;

		this._cleanFragments(opt.fragments);
		this.textRender.setString(opt.fragments, this.textFlags);
		this.newTextFormat = null;

		if (opt.zoom > 0) {
			this.overlayCtx.setFont(this.drawingCtx.getFont());
			this.changeZoom(opt.zoom);
		}

		this.kx = asc_getcvt(u, 0/*px*/, ctx.getPPIX());
		this.ky = asc_getcvt(u, 0/*px*/, ctx.getPPIY());

		this.sides = opt.getSides();

		this.left = this.sides.cellX;
		this.top = this.sides.cellY;
		this.right = this.sides.r[this.sides.ri];
		this.bottom = this.sides.b[this.sides.bi];

		this.cursorPos = 0;
		this.topLineIndex = 0;
		this.selectionBegin = -1;
		this.selectionEnd = -1;
		this.isSelectMode = c_oAscCellEditorSelectState.no;
		this.hasCursor = false;

		this.undoList = [];
		this.redoList = [];
		this.undoMode = false;
		this._setSkipKeyPress(false);

		this.updateWizardMode(false);
	};

	CellEditor.prototype._parseRangeStr = function (s) {
		var range = AscCommonExcel.g_oRangeCache.getAscRange(s);
		return range ? range.clone() : null;
	};

	CellEditor.prototype._parseFormulaRanges = function () {
		//I get a string without double-byte characters
		var s = this.options.fragments.reduce(function (pv, cv) {
			return pv + AscCommonExcel.convertUnicodeToSimpleString(cv.getCharCodes());
		}, "");
		var ws = this.handlers.trigger("getActiveWS");

		var bbox = this.options.bbox;
		this._parseResult = new AscCommonExcel.ParseResult([], []);
		this._parseResult.cursorPos = this.needFindFirstFunction ? undefined : this.cursorPos - 1;
		var cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		this._formula = new AscCommonExcel.parserFormula(s.substr(1), cellWithFormula, ws);
		this._formula.parse(true, true, this._parseResult, true);
		if (this.needFindFirstFunction) {
			this.argPosArr = this._parseResult.argPosArr;
			this.needFindFirstFunction = null;
		}

		var r, oper, wsName = null, bboxOper, range, isName = false;

		var oSelectionRange = new AscCommonExcel.SelectionRange(ws);
		// ToDo change create SelectionRange
		oSelectionRange.ranges = [];

		if (this._parseResult.refPos && this._parseResult.refPos.length > 0) {
			for (var index = 0; index < this._parseResult.refPos.length; index++) {
				wsName = null;
				isName = false;
				bboxOper = null;
				r = this._parseResult.refPos[index];
				oper = r.oper;
				if ((cElementType.table === oper.type || cElementType.name === oper.type ||
					cElementType.name3D === oper.type) && oper.externalLink == null) {
					oper = r.oper.toRef(bbox);
					if (oper instanceof AscCommonExcel.cError) {
						continue;
					}
					isName = true;
				}
				if ((cElementType.cell === oper.type || cElementType.cellsRange === oper.type || cElementType.cell3D === oper.type) && oper.externalLink == null) {
					wsName = oper.getWS().getName();
					bboxOper = oper.getBBox0();
				} else if ((cElementType.cellsRange3D === oper.type) && oper.externalLink == null) {
					if (oper.isSingleSheet()) {
						wsName = oper.getWS().getName();
						bboxOper = oper.getBBox0NoCheck();
					} else if (oper.isBetweenSheet(ws)) {
						wsName = ws.getName();
						bboxOper = oper.getBBox0NoCheck();
					}
				}
				if (bboxOper) {
					if (wsName && ws && ws.getName() !== wsName) {
						continue;
					}
					oSelectionRange.addRange();
					range = oSelectionRange.getLast();
					if (bboxOper.isOneCell()) {
						var isMerged = ws.getMergedByCell(bboxOper.r1, bboxOper.c1);
						if (isMerged) {
							bboxOper.r2 = isMerged.r2;
							bboxOper.c2 = isMerged.c2;
						}
					}
					range.assign2(bboxOper);
					range.cursorePos = range.colorRangePos = r.start + 1;
					range.formulaRangeLength = r.end - r.start;
					range.isName = isName;
				}
			}
		}

		this.handlers.trigger("newRanges", 0 !== oSelectionRange.ranges.length ? oSelectionRange : null);
	};

	CellEditor.prototype._findRangeUnderCursor = function () {
		// Get character string
		let s = this.textRender.getChars(0, this.textRender.getCharsCount());
		s = AscCommonExcel.convertUnicodeToSimpleString(s);
		let arrFR = this.handlers.trigger("getFormulaRanges");

		// Check cached formula ranges first
		if (arrFR) {
			let ranges = arrFR.ranges;
			// Check if cursor is over any existing ranges before re-parsing formula
			// Needed for cases like sumnas2:K2 where sumnas2 is invalid reference
			for (let i = 0, l = ranges.length; i < l; ++i) {
				let a = ranges[i];
				if (this.cursorPos >= a.cursorePos && this.cursorPos <= a.cursorePos + a.formulaRangeLength) {
					let range = a.clone(true);
					range.isName = a.isName;
					range.formulaRangeLength = a.formulaRangeLength;
					range.cursorePos = a.cursorePos;
					return {range: range};
				}
			}
		}

		// No ranges found under cursor, parse formula
		let r, offset, _e, _s, wsName = null, ret = false, refStr, isName = false;
		let _sColorPos, localStrObj;
		let ws = this.handlers.trigger("getActiveWS");

		let bbox = this.options.bbox;
		this._parseResult = new AscCommonExcel.ParseResult([], []);
		let cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		this._formula = new AscCommonExcel.parserFormula(s.substr(1), cellWithFormula, ws);
		this._formula.parse(true, true, this._parseResult, bbox);

		let range;
		if (this._parseResult.refPos && this._parseResult.refPos.length > 0) {
			for (let index = 0; index < this._parseResult.refPos.length; index++) {
				wsName = null;
				r = this._parseResult.refPos[index];

				offset = r.end;
				_e = r.end; 
				_sColorPos = _s = r.start;

				switch (r.oper.type) {
					case cElementType.cell: {
						wsName = ws.getName();
						refStr = r.oper.toLocaleString();
						ret = true;
						break;
					}
					case cElementType.cell3D: {
						localStrObj = r.oper.toLocaleStringObj();
						refStr = localStrObj[1];
						ret = true;
						wsName = r.oper.getWS().getName();
						_s = _e - localStrObj[1].length + 1;
						_sColorPos = _e - localStrObj[0].length;
						break;
					}
					case cElementType.cellsRange: {
						wsName = ws.getName();
						refStr = r.oper.toLocaleString();
						ret = true;
						break;
					}
					case cElementType.cellsRange3D: {
						if (!r.oper.isSingleSheet()) {
							continue;
						}
						ret = true;
						localStrObj = r.oper.toLocaleStringObj();
						refStr = localStrObj[1];
						wsName = r.oper.getWS().getName();
						_s = _e - localStrObj[1].length + 1;
						break;
					}
					case cElementType.table:
					case cElementType.name:
					case cElementType.name3D: {
						let nameRef = r.oper.toRef(bbox);
						if (nameRef instanceof AscCommonExcel.cError) {
							continue;
						}
						switch (nameRef.type) {
							case cElementType.cellsRange3D: {
								if (!nameRef.isSingleSheet()) {
									continue;
								}
							}
							case cElementType.cellsRange:
							case cElementType.cell3D: {
								ret = true;
								localStrObj = nameRef.toLocaleStringObj();
								refStr = localStrObj[1];
								wsName = nameRef.getWS().getName();
								_s = _e - localStrObj[1].length;
								break;
							}
						}
						isName = true;
						break;
					}
					default:
						continue;
				}

				if (ret && this.cursorPos > _s && this.cursorPos <= _s + refStr.length) {
					range = this._parseRangeStr(refStr);
					if (range) {
						range.isName = isName;
						range.formulaRangeLength = refStr.length;
						range.cursorePos = _s;
						return {range: range, wsName: wsName};
					}
				}
			}
		}
		range ? range.isName = isName : null;
		range ? range.formulaRangeLength = r.oper.value.length : null;
		range ? range.cursorePos = _s : null;
		return !range ? {range: null} : {range: range, wsName: wsName};
	};

	CellEditor.prototype._updateTopLineActive = function (state, isOpening) {
		if (state !== this.isTopLineActive) {
			this.isTopLineActive = state;
			this.openFromTopLine = isOpening && state;
			this.handlers.trigger("updateTopLine", this.isTopLineActive ? c_oAscCellEditorState.editInFormulaBar : c_oAscCellEditorState.editInCell);
		}
	};
	CellEditor.prototype._updateEditorState = function () {
		if (this.getMenuEditorMode()) {
			return;
		}
		var isFormula = this._isFormula();

		var editorState = isFormula ? c_oAscCellEditorState.editFormula : "" === AscCommonExcel.getFragmentsText(this.options.fragments) ? c_oAscCellEditorState.editEmptyCell : c_oAscCellEditorState.editText;
		this._setEditorState(editorState);

		this.handlers.trigger("updateFormulaEditMod", isFormula);
		if (isFormula) {
			this._parseFormulaRanges();
		}
	};
	CellEditor.prototype._cleanLastRangeInfo = function () {
		this.lastRangeLength = null;
		this.lastRangePos = null;
	};

	// Update Undo/Redo state
	CellEditor.prototype._updateUndoRedoChanged = function () {
		this.handlers.trigger("updateUndoRedoChanged", 0 < this.undoList.length, 0 < this.redoList.length);
	};

	CellEditor.prototype._haveTextInEdit = function () {
		var fragments = this.options.fragments;
		return fragments.length > 0 && fragments[0].getCharCodesLength() > 0;
	};

	CellEditor.prototype._setEditorState = function (editorState) {
		if (this.m_nEditorState !== editorState) {
			this.m_nEditorState = editorState;
			this.handlers.trigger("updateEditorState", this.m_nEditorState);
		}
	};

	CellEditor.prototype._getRenderFragments = function () {
		var opt = this.options, fragments = opt.fragments, i, k, l, first, last, val, lengthColors, tmpColors,
			colorIndex, uniqueColorIndex;
		if (this.isFormula()) {
			var ranges = this.handlers.trigger("getFormulaRanges");
			if (ranges) {
				fragments = [];
				for (i = 0; i < opt.fragments.length; ++i) {
					fragments.push(opt.fragments[i].clone());
				}

				lengthColors = AscCommonExcel.c_oAscFormulaRangeBorderColor.length;
				tmpColors = [];
				uniqueColorIndex = 0;
				ranges = ranges.ranges;
				for (i = 0, l = ranges.length; i < l; ++i) {
					val = ranges[i];
					colorIndex = asc.getUniqueRangeColor(ranges, i, tmpColors);
					if (null == colorIndex) {
						colorIndex = uniqueColorIndex++;
					}
					tmpColors.push(colorIndex);

					this._extractFragments(val.cursorePos, val.formulaRangeLength, fragments);
					first = this._findFragment(val.cursorePos, fragments);
					last = this._findFragment(val.cursorePos + val.formulaRangeLength - 1, fragments);
					if (first && last) {
						for (k = first.index; k <= last.index; ++k) {
							fragments[k].format.setColor(AscCommonExcel.c_oAscFormulaRangeBorderColor[colorIndex % lengthColors]);
						}
					}
				}
			}
		}

		return fragments;
	};

	// Rendering

	CellEditor.prototype._draw = function () {
		if (!this.options || !this.options.fragments) {
			return;
		}

		this._expand();
		this._cleanText();

		let externalSelectionController = this.handlers.trigger("getExternalSelectionController");
		if (!externalSelectionController || !externalSelectionController.getExternalFormulaEditMode()) {
			this._cleanSelection();
			this._adjustCanvas();
			this._showCanvas();
			this._calculateCanvasSize();
			this._renderText();
		}

		if (!this.getMenuEditorMode()) {
			for (var i = 0; i < this.options.fragments.length; i++) {
				this.options.fragments[i].initText();
			}
			this.input.value = AscCommonExcel.getFragmentsText((this.options.fragments));
		}
		this._updateCursorPosition();
		this._updateCursor();
	};

	CellEditor.prototype._update = function () {
		this._updateEditorState();

		let isExpand = this._expand();
		if (isExpand) {
			this._adjustCanvas();
			this._calculateCanvasSize();
		}

		// the call is needed to update the text of the top line, before updating the cursor position
		this.textRender.initStartX(0, this.textRender.lines[0], this._getContentLeft(), this._getContentWidth(), true);
		if (!this.getMenuEditorMode()) {
			this._fireUpdated();
		}
		this._updateCursorPosition(true, isExpand);
		this._updateCursor();

		this._updateUndoRedoChanged();

		if (window['IS_NATIVE_EDITOR']) {
			window['native']['onCellEditorChangeText'](AscCommonExcel.getFragmentsText(this.options.fragments));
		}
	};

	CellEditor.prototype._fireUpdated = function () {
		//TODO I save the text!
		var s = AscCommonExcel.getFragmentsText(this.options.fragments);
		var isFormula = -1 === this.beginCompositePos && (s.charAt(0) === "=" || s.charAt(0) === "+" || s.charAt(0) === "-");
		var api = window["Asc"]["editor"];
		var fPos, fName, match, fCurrent;

		if (!this.isTopLineActive || !this.skipTLUpdate || this.undoMode) {
			this.input.value = s;
		}

		//get a string without double-byte characters and pass it to the regular expression
		//positions of all functions must match
		//the question remains with arguments that can contain double-byte characters
		s = this.options.fragments ? this.options.fragments.reduce(function (pv, cv) {
			return pv + AscCommonExcel.convertUnicodeToSimpleString(cv.getCharCodes());
		}, "") : "";

		if (isFormula) {
			let obj = this._getFunctionByString(this.cursorPos, s);
			fPos = obj.fPos;
			fName = obj.fName;
			fCurrent = this._getEditableFunction(this._parseResult).func;
		}

		this.handlers.trigger("updated", s, this.cursorPos, fPos, fName);
		this.handlers.trigger("updatedEditableFunction", fCurrent, fPos !== undefined ? this.calculateOffset(fPos) : null);
		if (api && api.isMobileVersion) {
			this.restoreFocus();
		}
	};

	CellEditor.prototype._getFunctionByString = function (cursorPos, s) {
		let isInString = false;
		let isEscaped = false;

		if ('"' === s[cursorPos - 1]) {
			return {fPos: undefined, fName: undefined};
		}

		for (let i = 0; i < cursorPos; i++) {
			if (s[i] === '"') {
				if (!isEscaped) {
					isInString = !isInString;
				}
			}
		}

		if (isInString) {
			return {fPos: undefined, fName: undefined};
		}
		
		let fPos = asc_lastidx(s, this.reNotFormula, cursorPos) + 1;
		let match;
		if (fPos > 0) {
			match = s.slice(fPos, cursorPos).match(this.reFormula);
		}
		let fName;
		if (match) {
			fName = match[1];
		} else {
			fPos = undefined;
			fName = undefined;
		}

		return {fPos: fPos, fName: fName};
	};

	CellEditor.prototype._getEditableFunction = function (parseResult, bEndCurPos) {
		//TODO I save the text!
		var findOpenFunc = [], editableFunction = null, level = -1;
		if (!parseResult) {
			//in this case, I start parsing the formula up to the current position
			//I get a string without double-byte characters
			var s = this.options.fragments.reduce(function (pv, cv) {
				return pv + AscCommonExcel.convertUnicodeToSimpleString(cv.getCharCodes());
			}, "");
			var isFormula = -1 === this.beginCompositePos && s.charAt(0) === "=";
			if (isFormula) {
				var pos = this.cursorPos;
				var ws = this.handlers.trigger("getActiveWS");
				var bbox = this.options.bbox;

				var endPos = pos;
				if (!bEndCurPos) {
					for (var n = pos; n < s.length; n++) {
						if ("(" === s[n]) {
							endPos = n;
						}
					}
				}

				var formulaStr = s.substring(1, endPos);
				parseResult = new AscCommonExcel.ParseResult([], []);
				var cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
				var tempFormula = new AscCommonExcel.parserFormula(formulaStr, cellWithFormula, ws);
				tempFormula.parse(true, true, parseResult, true);
			}
		}

		var elements = parseResult ? parseResult.elems : null;
		if (elements) {
			for (var i = 0; i < elements.length; i++) {
				if (cElementType.func === elements[i].type && elements[i + 1] && "(" === elements[i + 1].name) {
					level++;
					findOpenFunc[level] = {elem: elements[i], counter: 1};
					i++;
				} else if (-1 !== level) {
					if ("(" === elements[i].name) {
						findOpenFunc[level].counter++;
					} else if (")" === elements[i].name) {
						findOpenFunc[level].counter--;
					}
				}
				if (level > -1 && findOpenFunc[level].counter === 0) {
					findOpenFunc.splice(level, 1);
					level--;
				}
			}
		}

		if (findOpenFunc) {
			for (var j = findOpenFunc.length - 1; j >= 0; j--) {
				if (findOpenFunc[j].counter > 0 && !(findOpenFunc[j].elem instanceof window['AscCommonExcel'].cUnknownFunction)) {
					editableFunction = findOpenFunc[j].elem.name;
					break;
				}
			}
		}

		return {func: editableFunction, argPos: parseResult ? parseResult.argPos : null};
	};

	CellEditor.prototype._expand = function () {
		var bottom, tm;
		var doAdjust = false, fragments = this._getRenderFragments();
		if (fragments && 0 < fragments.length) {
			bottom = this.bottom;
			this.bottom = this.sides.b[this.sides.bi];

			this._updateTextAlign();
			tm = this.textRender.measureString(fragments, this.textFlags, this._getContentWidth());

			if (!this.textFlags.wrapText && !this.textFlags.wrapOnlyCE) {
				while (tm.width > this._getContentWidth()) {
					if (!this._expandWidth()) {
						this.textFlags.wrapOnlyCE = true;
						tm = this.textRender.measureString(fragments, this.textFlags, this._getContentWidth());
						break;
					}
					doAdjust = true;
				}
			}
			while (tm.height > this._getContentHeight() && this._expandHeight()) {
			}
			if (bottom !== this.bottom) {
				if (bottom > this.bottom) {
					// Clear index when reduce size
					this.topLineIndex = 0;
				}
				doAdjust = true;
				// ToDo move this to _adjustCanvas
				if (this.getMenuEditorMode) {
					this.handlers.trigger("resizeEditorHeight");
				}
			}
		}
		return doAdjust;
	};
	CellEditor.prototype._expandWidth = function () {
		var i, l = -1, r = -1;

		if (AscCommon.align_Left === this.textFlags.textAlign || AscCommon.align_Center === this.textFlags.textAlign) {
			var rightSide = this.sides.r;
			for (i = 0; i < rightSide.length; ++i) {
				if (rightSide[i] > this.right) {
					r = rightSide[i];
					break;
				}
			}
		}
		if (AscCommon.align_Right === this.textFlags.textAlign || AscCommon.align_Center === this.textFlags.textAlign) {
			var leftSide = this.sides.l;
			for (i = 0; i < leftSide.length; ++i) {
				if (leftSide[i] < this.left) {
					l = leftSide[i];
					break;
				}
			}
		}

		if (AscCommon.align_Center === this.textFlags.textAlign) {
			if (-1 !== l && -1 !== r) {
				var min = Math.min(this.left - l, r - this.right);
				this.left -= min;
				this.right += min;
				return true;
			}
		} else {
			if (-1 !== l) {
				this.left = l;
				return true;
			} else if (-1 !== r) {
				this.right = r;
				return true;
			}
		}

		return false;
	};

	CellEditor.prototype._expandHeight = function () {
		var t = this, bottomSide = this.sides.b, i = asc_search(bottomSide, function (v) {
			return v > t.bottom;
		});
		if (i >= 0) {
			t.bottom = bottomSide[i];
			return true;
		}
		var val = bottomSide[bottomSide.length - 1];
		if (Math.abs(t.bottom - val) > 0.000001) { // bottom !== bottomSide[len-1]
			t.bottom = val;
		}
		return false;
	};

	CellEditor.prototype._cleanText = function () {
		this.drawingCtx.clear();
	};

	CellEditor.prototype._showCanvas = function () {
		this.canvasOuterStyle.display = 'block';
	};

	CellEditor.prototype._hideCanvas = function () {
		this.canvasOuterStyle.display = 'none';
	};

	CellEditor.prototype._adjustCanvas = function () {
		var z = this.defaults.canvasZIndex;
		var borderSize = AscBrowser.retinaPixelRatio === 1.5 ? 1 : AscCommon.AscBrowser.convertToRetinaValue(1, true);
		var left = this.left * this.kx;
		var top = this.top * this.ky;
		var width, height, widthStyle, heightStyle;

		width = widthStyle = (this.right - this.left) * this.kx - borderSize;
		height = heightStyle = (this.bottom - this.top) * this.ky - borderSize;

		left = AscCommon.AscBrowser.convertToRetinaValue(left);
		top = AscCommon.AscBrowser.convertToRetinaValue(top);
		widthStyle = AscCommon.AscBrowser.convertToRetinaValue(widthStyle);
		heightStyle = AscCommon.AscBrowser.convertToRetinaValue(heightStyle);

		// in safari with hardware acceleration enabled, there is a bug when entering text.
		// apparently they cache textures in a special way that are (w*h<5000) in size
		// the formula is accurate. not a pixel less. more - you can have as much as you like.
		// you need to check every safari update - and when they fix it - remove this stub
		// canvases are transparent and their increased size does not affect the result.
		//
		// in the new version of safari, we increase not only the canvases, but also the div.
		if (AscCommon.AscBrowser.isSafariMacOs) {
			if ((widthStyle * heightStyle) < 5000) {
				this._originalCanvasWidth = width;
				widthStyle = ((5000 / heightStyle) >> 0) + 1;
			} else {
				this._originalCanvasWidth = null;
			}
		}

		// Calculate canvas offset inside container
		let ws = this.handlers.trigger("getActiveWSView");
		let canvasTop = 0;
		let cellsTop = ws && AscCommon.AscBrowser.convertToRetinaValue(ws.cellsTop);
		if (ws && top < cellsTop) {
			// If editor position is above data area
			canvasTop = top < 0 ? -(cellsTop + Math.abs(top)) : top - cellsTop;
			// Fix container at headers level
			top = cellsTop;
		}

		this.canvasOuterStyle.left = left + 'px';
		this.canvasOuterStyle.top = top + 'px';
		this.canvasOuterStyle.width = widthStyle + 'px';
		this.canvasOuterStyle.height = heightStyle + 'px';
		if (!this.getMenuEditorMode()) {
			this.canvasOuterStyle.zIndex = /*this.top < 0 ? -1 :*/ z;
		}

		this.canvas.style.width = this.canvasOverlay.style.width = widthStyle + 'px';
		this.canvas.style.height = this.canvasOverlay.style.height = heightStyle + 'px';
		this.canvas.style.top = this.canvasOverlay.style.top = canvasTop + 'px';
	};

	CellEditor.prototype._calculateCanvasSize = function () {
		//this code is called after showCanvas because inside calculateCanvasSize getBoundingClientRect is used
		//if canvas has display = 'none' then the sizes will be returned as zero
		if (this.canvas) {
			AscCommon.calculateCanvasSize(this.canvas);
		}
		if (this.canvasOverlay) {
			AscCommon.calculateCanvasSize(this.canvasOverlay);
		}
	};

	CellEditor.prototype._renderText = function (dy, forceRender) {

		if (window.LOCK_DRAW && !forceRender)
		{
			this.textRender.initStartX(0, null, this._getContentLeft(), this._getContentWidth(), true);
			window.TEXT_DRAW_INSTANCE = this;
			window.TEXT_DRAW_INSTANCE_POS = dy;
			return;
		}

		if (forceRender) {
			window.TEXT_DRAW_INSTANCE = undefined;
		}

		var t = this, opt = t.options, ctx = t.drawingCtx;

		if (!window['IS_NATIVE_EDITOR']) {
			let _width = this._originalCanvasWidth ? this._originalCanvasWidth : ctx.getWidth();
			if (opt.background) {
				ctx.setFillStyle(opt.background);
			}
			ctx.fillRect(0, 0, _width, ctx.getHeight());
		}

		if (opt.fragments && opt.fragments.length > 0) {
			t.textRender.render(undefined, t._getContentLeft(), dy || 0, t._getContentWidth(), opt.font.getColor());
		}
	};

	CellEditor.prototype._cleanSelection = function () {
		this.overlayCtx.clear();
	};

	CellEditor.prototype._drawSelection = function () {
		var ctx = this.overlayCtx;
		var zoom = this.getZoom();
		var begPos, endPos, top, topLine, begInfo, endInfo, line, i, y, h, selection = [];

		function drawRect(x, y, w, h) {
			if (window['IS_NATIVE_EDITOR']) {
				selection.push([x, y, w, h]);
			} else {
				ctx.fillRect(x, y, w, h);
			}
		}

		begPos = this.selectionBegin;
		endPos = this.selectionEnd;

		if (!window['IS_NATIVE_EDITOR']) {
			ctx.setFillStyle(this.defaults.selectColor).clear();
		}

		if (begPos !== endPos && !this.isTopLineActive) {
			top = this.textRender.calcLineOffset(this.topLineIndex);
			begInfo = this.textRender.calcCharOffset(Math.min(begPos, endPos));
			line = this.textRender.getLineInfo(begInfo.lineIndex);
			topLine = this.textRender.calcLineOffset(begInfo.lineIndex);
			endInfo = this.textRender.calcCharOffset(Math.max(begPos, endPos));
			h = asc_round(line.th * zoom);
			y = topLine - top;
			if (begInfo.lineIndex === endInfo.lineIndex) {
				drawRect(begInfo.left, y, endInfo.left - begInfo.left, h);
			} else {
				drawRect(begInfo.left, y, line.tw - begInfo.left + line.startX, h);
				for (i = begInfo.lineIndex + 1, y += h; i < endInfo.lineIndex; ++i, y += h) {
					line = this.textRender.getLineInfo(i);
					h = asc_round(line.th * zoom);
					drawRect(line.startX, y, line.tw, h);
				}
				line = this.textRender.getLineInfo(endInfo.lineIndex);
				topLine = this.textRender.calcLineOffset(endInfo.lineIndex);
				if (line) {
					drawRect(line.startX, topLine - top, endInfo.left - line.startX, asc_round(line.th * zoom));
				}
			}
		}
		if (!this.isSelectMode) {
			this.handlers.trigger("onSelectionEnd");
		}

		let externalSelectionController = this.handlers.trigger("getExternalSelectionController");
		externalSelectionController && externalSelectionController.sendExternalChangeSelection();

		return selection;
	};

	CellEditor.prototype.calculateOffset = function (pos) {
		var left = 0;
		var top = 0;
		if (pos != null && this.textRender) {
			var _top = this.textRender.calcLineOffset(this.topLineIndex);
			var _begInfo = this.textRender.calcCharOffset(pos);
			var _topLine = _begInfo ? this.textRender.calcLineOffset(_begInfo.lineIndex) : null;

			left = _begInfo && _begInfo.left ? AscCommon.AscBrowser.convertToRetinaValue(_begInfo.left) : 0;
			top = _topLine != null && _top != null ? AscCommon.AscBrowser.convertToRetinaValue(_topLine - _top) : 0;
		}

		return [left, top];
	};

	// Cursor

	CellEditor.prototype._updateCursorStyle = function (cursor) {
		var newHtmlCursor = AscCommon.g_oHtmlCursor.value(cursor);
		if (this.canvasOverlay.style.cursor !== newHtmlCursor) {
			this.canvasOverlay.style.cursor = newHtmlCursor;
		}
	};

	CellEditor.prototype._updateCursor = function () {
		if (window['IS_NATIVE_EDITOR']) {
			return;
		}

		if (!this.isOpened || this.options.enterOptions.hideCursor || this.isTopLineActive
			|| !this.enableKeyEvents || this.handlers.trigger('getWizard')) {
			this._hideCursor();
		} else {
			this._showCursor();
		}
	};

	CellEditor.prototype.showCursor = function () {
		this.options.enterOptions.hideCursor = false;
		this._updateCursor();
	};

	CellEditor.prototype._showCursor = function () {
		var t = this;
		window.clearInterval(t.cursorTID);
		t.cursorStyle.display = "block";
		t.cursorTID = window.setInterval(function () {
			t.cursorStyle.display = ("none" === t.cursorStyle.display) ? "block" : "none";
		}, t.defaults.blinkInterval);
	};

	CellEditor.prototype._hideCursor = function () {
		window.clearInterval(this.cursorTID);
		this.cursorStyle.display = "none";
	};

	CellEditor.prototype._updateCursorPosition = function (redrawText, isExpand, lineIndex) {
		// ToDo should forward this function
		let h = this.canvas.height;
		let y = -this.textRender.calcLineOffset(this.topLineIndex);
		let cur = this.textRender.calcCharOffset(this.cursorPos, lineIndex);
		let charsCount = this.textRender.getCharsCount();
		let textAlign = this.textFlags && this.textFlags.textAlign;
		let curLeft = asc_round(
			((AscCommon.align_Right !== textAlign || this.cursorPos !== charsCount) && cur !== null &&
			cur.left !== null ? cur.left : this._getContentPosition()) * this.kx);
		let curTop = asc_round(((cur !== null ? cur.top : 0) + y) * this.ky);
		let curHeight = asc_round((cur !== null ? cur.height : this._getContentHeight()) * this.ky);
		let i, dy, nCount = this.textRender.getLinesCount();
		let zoom = this.getZoom();

		while (1 < nCount) {
			if (curTop + curHeight - 1 > h) {
				i = i === undefined ? 0 : i + 1;
				if (i === nCount) {
					break;
				}
				dy = asc_round(this.textRender.getLineInfo(i).th * zoom);
				y -= dy;
				curTop -= asc_round(dy * this.ky);
				++this.topLineIndex;
				continue;
			}
			if (curTop < 0) {
				--this.topLineIndex;
				if (this.textRender.lines && this.textRender.lines.length && this.topLineIndex >= this.textRender.lines.length) {
					this.topLineIndex = this.textRender.lines.length - 1;
				}
				dy = asc_round(this.textRender.getLineInfo(this.topLineIndex).th * zoom);
				y += dy;
				curTop += asc_round(dy * this.ky);
				continue;
			}
			break;
		}

		if (dy !== undefined || redrawText) {
			this._renderText(y, isExpand);
		}

		curLeft = AscCommon.AscBrowser.convertToRetinaValue(curLeft);
		curTop = AscCommon.AscBrowser.convertToRetinaValue(curTop);
		curHeight = AscCommon.AscBrowser.convertToRetinaValue(curHeight);

		this.curLeft = curLeft;
		this.curTop = curTop;
		this.curHeight = curHeight;

		if (!window['IS_NATIVE_EDITOR']) {

			// update cursor position
			let scrollDiff = parseFloat(this.canvas.style.top);

			this.cursorStyle.left = curLeft + "px";
			this.cursorStyle.top = curTop + scrollDiff + "px";

			this.cursorStyle.width = (((2 * AscCommon.AscBrowser.retinaPixelRatio) >> 0) / AscCommon.AscBrowser.retinaPixelRatio) + "px";
			this.cursorStyle.height = curHeight + "px";
		}

		if (AscCommon.g_inputContext) {
			AscCommon.g_inputContext.moveAccurate(this.left * this.kx + curLeft, this.top * this.ky + curTop);
		}

		if (cur) {
			this.input.scrollTop = this.input.clientHeight * cur.lineIndex;
		}
		if (this.isTopLineActive && !this.skipTLUpdate) {
			this._updateTopLineCurPos();
		}

		if (this.getMenuEditorMode()) {
			this.handlers.trigger("updateMenuEditorCursorPosition", curTop, curHeight);
		}

		//let fCurrent = this._getEditableFunction(null, true);
		//console.log("func: " + fCurrent.func + " arg: " + fCurrent.argPos);
		this._updateSelectionInfo();
	};

	CellEditor.prototype._moveCursor = function (kind, pos, lineIndex) {
		this.newTextFormat = null;
		var t = this;
		this.sAutoComplete = null;
		switch (kind) {
			case kPrevChar:
				t.cursorPos = t.textRender.getPrevChar(t.cursorPos);
				break;
			case kNextChar:
				t.cursorPos = t.textRender.getNextChar(t.cursorPos);
				break;
			case kPrevWord:
				t.cursorPos = t.textRender.getPrevWord(t.cursorPos);
				break;
			case kNextWord:
				t.cursorPos = t.textRender.getNextWord(t.cursorPos);
				break;
			case kBeginOfLine:
				t.cursorPos = t.textRender.getBeginOfLine(t.cursorPos);
				break;
			case kEndOfLine:
				t.cursorPos = t.textRender.getEndOfLine(t.cursorPos);
				break;
			case kBeginOfText:
				t.cursorPos = t.textRender.getBeginOfText(t.cursorPos);
				break;
			case kEndOfText:
				t.cursorPos = t.textRender.getEndOfText(t.cursorPos);
				break;
			case kPrevLine:
				t.cursorPos = t.textRender.getPrevLine(t.cursorPos);
				break;
			case kNextLine:
				t.cursorPos = t.textRender.getNextLine(t.cursorPos);
				break;
			case kPosition:
				t.cursorPos = pos;
				break;
			case kPositionLength:
				t.cursorPos += pos;
				break;
			default:
				return;
		}
		if (t.selectionBegin !== t.selectionEnd) {
			t.selectionBegin = t.selectionEnd = -1;
			t._cleanSelection();
		}
		t._updateCursorPosition(null, null, lineIndex);
		t._updateCursor();
	};

	CellEditor.prototype._findCursorPosition = function (coord) {
		return this.textRender.getCharPosByXY(coord.x, coord.y, this.topLineIndex, this.getZoom());
	};

	CellEditor.prototype._findLineIndex = function (coord) {
		return this.textRender.getLineByY(coord.y, this.topLineIndex, this.getZoom());
	};

	CellEditor.prototype._updateTopLineCurPos = function () {
		if (this.loadFonts) {
			return;
		}
		var isSelected = this.selectionBegin !== this.selectionEnd;
		var b = isSelected ? this.selectionBegin : this.cursorPos;
		var e = isSelected ? this.selectionEnd : this.cursorPos;
		if (this.input.setSelectionRange) {
			this.input.setSelectionRange(Math.min(b, e), Math.max(b, e));
		}
	};

	CellEditor.prototype._topLineGotFocus = function () {
		this._updateTopLineActive(true);
		this.input.isFocused = true;
		this.setFocus(true);
		this._updateCursor();
		this._cleanSelection();
	};

	CellEditor.prototype._topLineMouseUp = function () {
		this.callTopLineMouseup = false;
		// with this combination ctrl+a, click, ctrl+a, click selectionStart is not updated
		// therefore we perform processing after the system handler
		this._delayedUpdateCursorByTopLine();
	};
	CellEditor.prototype._delayedUpdateCursorByTopLine = function () {
		var t = this;
		setTimeout(function () {
			t._updateCursorByTopLine();
		});
	};
	CellEditor.prototype._updateCursorByTopLine = function () {
		var b = this.input.selectionStart;
		var e = this.input.selectionEnd;
		// ToDo replace code to input.selectionDirection after updating closure-compiler to version 20200719
		if ('backward' === this.input["selectionDirection"]) {
			var tmp = b;
			b = e;
			e = tmp;
		}
		if (typeof b !== "undefined") {
			if (this.cursorPos !== b || this.selectionBegin !== this.selectionEnd) {
				this._moveCursor(kPosition, b);
			}
			if (b !== e) {
				this._selectChars(kPosition, e);
			}

			//onSelectionEnd - used in plugins. It is needed to track the change of select.
			if (!this.isSelectMode) {
				this.handlers.trigger("onSelectionEnd");
			}
		}
	};

	CellEditor.prototype._syncEditors = function () {
		var t = this;
		var s1 = AscCommonExcel.getFragmentsCharCodes(t.options.fragments);
		var s2 = AscCommon.convertUTF16toUnicode(t.input.value);
		var l = Math.min(s1.length, s2.length);
		var i1 = 0, i2;

		while (i1 < l && s1[i1] === s2[i1]) {
			++i1;
		}
		i2 = i1 + 1;
		if (i2 >= l) {
			i2 = Math.max(s1.length, s2.length);
		} else {
			while (i2 < l && s1[i1] !== s2[i2]) {
				++i2;
			}
		}

		t._addChars(s2.slice(i1, i2), i1);
	};

	// Content

	CellEditor.prototype.getText = function () {
		return AscCommonExcel.getFragmentsText(this.options.fragments);
	};

	CellEditor.prototype._getContentLeft = function () {
		return this.defaults.padding;
	};

	CellEditor.prototype._getContentWidth = function () {
		//remove 1 px offset. without cell editor no 1 px offset
		return this.right - this.left - 2 * this.defaults.padding /*+ 1*//*px*/;
	};

	CellEditor.prototype._getContentHeight = function () {
		var t = this;
		return t.bottom - t.top;
	};

	CellEditor.prototype._getContentPosition = function () {
		if (!this.textFlags) {
			return this.defaults.padding;
		}
		switch (this.textFlags.textAlign) {
			case AscCommon.align_Right:
				return this.right - this.left - this.defaults.padding - 1;
			case AscCommon.align_Center:
				return 0.5 * (this.right - this.left);
		}
		return this.defaults.padding;
	};

	CellEditor.prototype._wrapText = function () {
		this.textFlags.wrapOnlyNL = true;
	};

	CellEditor.prototype._addCharCodes = function (arrCharCodes) {
		return this._addChars(arrCharCodes);
	};
	CellEditor.prototype._addChars = function (str, pos, isRange) {
		if (!isRange) {
			this.cleanSelectRange();
		}
		this.startAction();

		var opt = this.options, f, l, s;

		var noUpdateMode = this.noUpdateMode;
		this.noUpdateMode = true;

		this.sAutoComplete = null;

		if (!opt.fragments) {
			return;
		}

		if (this.selectionBegin !== this.selectionEnd) {
			var copyFragment = this._findFragmentToInsertInto(Math.min(this.selectionBegin, this.selectionEnd) + 1);
			if (copyFragment && !this.newTextFormat) {
				this.newTextFormat = opt.fragments[copyFragment.index].format.clone();
			}

			this._removeChars(undefined, undefined, isRange);
		}

		if (str.trim) {
			str = AscCommon.convertUTF16toUnicode(str);
		}
		var length = str.length;
		if (0 !== length) {
			// limit count characters
			var excess = this._checkMaxCellLength(length);
			if (excess) {
				length -= excess;
				if (0 === length) {
					this.noUpdateMode = noUpdateMode;
					return length;
				}
				str = str.slice(0, length);
			}

			if (pos === undefined) {
				pos = this.cursorPos;
			}

			if (!this.undoMode) {
				// save info to undo/redo
				this.undoList.push({fn: this._removeChars, args: [pos, length], isRange: isRange});
				this.redoList = [];
			}

			if (this.newTextFormat) {
				var oNewObj = new Fragment({format: this.newTextFormat, charCodes: str});
				this._addFragments([oNewObj], pos);
				this.newTextFormat = null;
			} else {
				f = this._findFragmentToInsertInto(pos);
				if (f) {
					l = pos - f.begin;
					s = opt.fragments[f.index].getCharCodes();

					opt.fragments[f.index].setCharCodes(s.slice(0, l).concat(str).concat(s.slice(l)));

					s = opt.fragments[f.index].getCharCodes();
				}
			}

			this.cursorPos = pos + str.length;
			if (-1 !== window["Asc"].search(str, function (val) {
				return val === codeNewLine
			})) {
				this._wrapText();
			}
		}

		this.noUpdateMode = noUpdateMode;
		if (!this.noUpdateMode) {
			this._update();
		}
		this.endAction();
		return length;
	};

	CellEditor.prototype._addNewLine = function () {
		this._wrapText();
		let sNewLine = "\n";
		this._addChars( /*codeNewLine*/sNewLine);
	};

	CellEditor.prototype._removeChars = function (pos, length, isRange) {
		var t = this, opt = t.options, b, e, l, first, last;

		if (!isRange) {
			this.cleanSelectRange();
		}

		this.sAutoComplete = null;

		if (t.selectionBegin !== t.selectionEnd) {
			b = Math.min(t.selectionBegin, t.selectionEnd);
			e = Math.max(t.selectionBegin, t.selectionEnd);
			t.selectionBegin = t.selectionEnd = -1;
			t._cleanSelection();
		} else if (length === undefined) {
			switch (pos) {
				case kPrevChar:
					b = t.textRender.getPrevChar(t.cursorPos, false);
					e = t.cursorPos;
					break;
				case kNextChar:
					b = t.cursorPos;
					e = t.textRender.getNextChar(t.cursorPos);
					break;
				case kPrevWord:
					b = t.textRender.getPrevWord(t.cursorPos);
					e = t.cursorPos;
					break;
				case kNextWord:
					b = t.cursorPos;
					e = t.textRender.getNextWord(t.cursorPos);
					break;
				default:
					return;
			}
		} else {
			b = pos;
			e = pos + length;
		}

		if (b === e) {
			return;
		}
		if (!opt.fragments) {
			return;
		}

		this.startAction();

		// search for begin and end positions
		first = t._findFragment(b);
		last = t._findFragment(e - 1);

		if (!t.undoMode) {
			// save info to undo/redo
			if (e - b < 2 && opt.fragments[first.index].getCharCodesLength() > 1) {
				t.undoList.push({fn: t._addChars, args: [t.textRender.getChars(b, 1), b], isRange: isRange});
			} else {
				t.undoList.push({fn: t._addFragments, args: [t._getFragments(b, e - b), b], isRange: isRange});
			}
			t.redoList = [];
		}

		if (first && last) {
			// remove chars
			if (first.index === last.index) {
				l = opt.fragments[first.index].getCharCodes();
				opt.fragments[first.index].setCharCodes(l.slice(0, b - first.begin).concat(l.slice(e - first.begin)));
			} else {
				opt.fragments[first.index].setCharCodes(opt.fragments[first.index].getCharCodes().slice(0, b - first.begin));
				opt.fragments[last.index].setCharCodes(opt.fragments[last.index].getCharCodes().slice(e - last.begin));
				l = last.index - first.index;
				if (l > 1) {
					opt.fragments.splice(first.index + 1, l - 1);
				}
			}
			// merge fragments with equal formats
			t._mergeFragments();
		}

		t.cursorPos = b;
		if (!t.noUpdateMode) {
			t._update();
		}
		this.endAction();
	};

	CellEditor.prototype._selectChars = function (kind, pos) {
		var t = this;
		var begPos, endPos;

		this.sAutoComplete = null;
		begPos = t.selectionBegin === t.selectionEnd ? t.cursorPos : t.selectionBegin;
		t._moveCursor(kind, pos);
		endPos = t.cursorPos;

		t.selectionBegin = begPos;
		t.selectionEnd = endPos;
		t._drawSelection();
		if (t.isTopLineActive && !t.skipTLUpdate) {
			t._updateTopLineCurPos();
		}
	};

	CellEditor.prototype._changeSelection = function (coord) {
		var t = this;

		function doChangeSelection(coordTmp) {
			// ToDo implement for the word.
			if (c_oAscCellEditorSelectState.word === t.isSelectMode) {
				return;
			}
			var pos = t._findCursorPosition(coordTmp);
			if (pos !== undefined) {
				pos >= 0 ? t._selectChars(kPosition, pos) : t._selectChars(pos);
			}
		}

		if (window['IS_NATIVE_EDITOR']) {
			doChangeSelection(coord);
		} else {
			window.clearTimeout(t.selectionTimer);
			t.selectionTimer = window.setTimeout(function () {
				doChangeSelection(coord);
			}, 0);
		}
	};

	CellEditor.prototype._findFragment = function (pos, fragments) {
		var i, begin, end;
		if (!fragments) {
			fragments = this.options.fragments;
		}
		if (!fragments) {
			return;
		}

		for (i = 0, begin = 0; i < fragments.length; ++i) {
			end = begin + fragments[i].getCharCodesLength();
			if (pos >= begin && pos < end) {
				return {index: i, begin: begin, end: end};
			}
			if (i < fragments.length - 1) {
				begin = end;
			}
		}
		return pos === end ? {index: i - 1, begin: begin, end: end} : undefined;
	};

	CellEditor.prototype._findFragmentToInsertInto = function (pos, fragments) {
		var i, begin, end;

		if (!fragments) {
			fragments = this.options.fragments;
		}
		if (!fragments) {
			return;
		}

		for (i = 0, begin = 0; i < fragments.length; ++i) {
			end = begin + fragments[i].getCharCodesLength();
			if (pos >= begin && pos <= end) {
				return {index: i, begin: begin, end: end};
			}
			if (i < fragments.length - 1) {
				begin = end;
			}
		}
		return undefined;
	};

	CellEditor.prototype._isWholeFragment = function (pos, len) {
		var fr = this._findFragment(pos);
		return fr && pos === fr.begin && len === fr.end - fr.begin;
	};

	CellEditor.prototype._splitFragment = function (f, pos, fragments) {
		var fr;
		if (!fragments) {
			fragments = this.options.fragments;
		}
		if (!fragments) {
			return;
		}

		if (pos > f.begin && pos < f.end) {
			fr = fragments[f.index];
			Array.prototype.splice.apply(fragments, [f.index, 1].concat([new Fragment({
				format: fr.format.clone(), charCodes: fr.getCharCodes().slice(0, pos - f.begin)
			}), new Fragment({format: fr.format.clone(), charCodes: fr.getCharCodes().slice(pos - f.begin)})]));
		}
	};

	CellEditor.prototype._getFragments = function (startPos, length) {
		var t = this, opt = t.options, endPos = startPos + length - 1, res = [], fr, i;
		var first = t._findFragment(startPos);
		var last = t._findFragment(endPos);

		if (!first || !last) {
			throw new Error("Can not extract fragment of text");
		}

		if (first.index === last.index) {
			fr = opt.fragments[first.index].clone();
			fr.charCodes = fr.getCharCodes().slice(startPos - first.begin, endPos - first.begin + 1);
			res.push(fr);
		} else {
			fr = opt.fragments[first.index].clone();
			fr.charCodes = fr.getCharCodes().slice(startPos - first.begin);
			res.push(fr);
			for (i = first.index + 1; i < last.index; ++i) {
				fr = opt.fragments[i].clone();
				res.push(fr);
			}
			fr = opt.fragments[last.index].clone();
			fr.charCodes = fr.getCharCodes().slice(0, endPos - last.begin + 1);
			res.push(fr);
		}

		return res;
	};

	CellEditor.prototype._extractFragments = function (startPos, length, fragments) {
		var fr;

		fr = this._findFragment(startPos, fragments);
		if (!fr) {
			throw new Error("Can not extract fragment of text");
		}
		this._splitFragment(fr, startPos, fragments);

		fr = this._findFragment(startPos + length, fragments);
		if (!fr) {
			throw new Error("Can not extract fragment of text");
		}
		this._splitFragment(fr, startPos + length, fragments);
	};

	CellEditor.prototype._addFragments = function (f, pos) {
		var t = this, opt = t.options, fr;

		if (!opt.fragments) {
			return;
		}

		fr = t._findFragment(pos);
		if (fr && pos < fr.end) {
			t._splitFragment(fr, pos);
			fr = t._findFragment(pos);
			Array.prototype.splice.apply(opt.fragments, [fr.index, 0].concat(f));
		} else {
			opt.fragments = opt.fragments.concat(f);
		}

		// merge fragments with equal formats
		t._mergeFragments();

		t.cursorPos = pos + AscCommonExcel.getFragmentsLength(f);
		if (!t.noUpdateMode) {
			t._update();
		}
	};

	CellEditor.prototype._mergeFragments = function (fragments) {
		var i;

		if (!fragments) {
			fragments = this.options.fragments;
		}
		if (!fragments) {
			return;
		}

		for (i = 0; i < fragments.length;) {
			if (fragments[i].getCharCodesLength() < 1 && fragments.length > 1) {
				fragments.splice(i, 1);
				continue;
			}
			if (i < fragments.length - 1) {
				var fr = fragments[i];
				var nextFr = fragments[i + 1];
				if (fr.format.isEqual(nextFr.format)) {
					fragments.splice(i, 2, new Fragment({
						format: fr.format,
						charCodes: fr.getCharCodes().concat(nextFr.getCharCodes())
					}));
					continue;
				}
			}
			++i;
		}
	};

	CellEditor.prototype._cleanFragments = function (fr) {
		var t = this, i, s, f, wrap = t.textFlags.wrapText || t.textFlags.wrapOnlyNL;

		if (!fr) {
			return;
		}

		for (i = 0; i < fr.length; ++i) {
			s = fr[i].getCharCodes();
			if (!wrap && -1 !== window["Asc"].search(s, function (val) {
				return val === codeNewLine
			})) {
				this._wrapText();
			}
			fr[i].setCharCodes(s);
			f = fr[i].format;
			if (f.getName() === "") {
				f.setName(t.options.font.getName());
			}
			if (f.getSize() === 0) {
				f.setSize(t.options.font.getSize());
			}
		}
	};

	CellEditor.prototype._setFormatProperty = function (format, prop, val) {
		switch (prop) {
			case "fn":
				format.setName(val);
				format.setScheme(null);
				break;
			case "fs":
				format.setSize(val);
				break;
			case "b":
				var bold = format.getBold();
				val = (null === val) ? ((bold) ? !bold : true) : val;
				format.setBold(val);
				break;
			case "i":
				var italic = format.getItalic();
				val = (null === val) ? ((italic) ? !italic : true) : val;
				format.setItalic(val);
				break;
			case "u":
				var underline = format.getUnderline();
				val = (null === val) ? ((Asc.EUnderline.underlineNone === underline) ? Asc.EUnderline.underlineSingle :
					Asc.EUnderline.underlineNone) : val;
				format.setUnderline(val);
				break;
			case "s":
				var strikeout = format.getStrikeout();
				val = (null === val) ? ((strikeout) ? !strikeout : true) : val;
				format.setStrikeout(val);
				break;
			case "fa":
				format.setVerticalAlign(val);
				break;
			case "c":
				format.setColor(val);
				break;
			case "changeFontSize":
				var newFontSize = asc_incDecFonSize(val, format.getSize());
				if (null !== newFontSize) {
					format.setSize(newFontSize);
				}
				break;
		}
		return val;
	};

	CellEditor.prototype._performAction = function (list1, list2) {
		var t = this, action, str, pos, len;

		if (list1.length < 1) {
			return;
		}

		action = list1.pop();

		if (action.fn === t._removeChars) {
			pos = action.args[0];
			len = action.args[1];
			if (len < 2 && !t._isWholeFragment(pos, len)) {
				list2.push({fn: t._addChars, args: [t.textRender.getChars(pos, len), pos], isRange: action.isRange});
			} else {
				list2.push({fn: t._addFragments, args: [t._getFragments(pos, len), pos], isRange: action.isRange});
			}
		} else if (action.fn === t._addChars) {
			str = action.args[0];
			pos = action.args[1];
			list2.push({fn: t._removeChars, args: [pos, str.length], isRange: action.isRange});
		} else if (action.fn === t._addFragments) {
			pos = action.args[1];
			len = AscCommonExcel.getFragmentsLength(action.args[0]);
			list2.push({fn: t._removeChars, args: [pos, len], isRange: action.isRange});
		} else if (action.fn === t._changeFragments) {
			let _fragments = action.args[0];
			let _redoFragments = {};
			for (let i in _fragments) {
				if (_fragments.hasOwnProperty(i)) {
					if (this.options.fragments && this.options.fragments[i]) {
						_redoFragments[i] = this.options.fragments[i].clone();
					}
				}
			}
			list2.push({fn: t._changeFragments, args: [_redoFragments]});
		} else {
			return;
		}

		t.undoMode = true;
		if (t.selectionBegin !== t.selectionEnd) {
			t.selectionBegin = t.selectionEnd = -1;
			t._cleanSelection();
		}
		action.fn.apply(t, action.args);
		t.undoMode = false;
	};

	CellEditor.prototype._tryCloseEditor = function (event) {
		var t = this;
		let nRetValue = keydownresult_PreventNothing;
		var callback = function (success) {
			// for the case when the user presses ctrl+shift+enter/crtl+enter the transition to a new line is not performed
			var applyByArray = t.textFlags && t.textFlags.ctrlKey;
			if (!applyByArray && success) {
				nRetValue = t.handlers.trigger("applyCloseEvent", event);
				AscCommon.StartIntervalDrawText(false);
			}
		};
		this.close(true, callback);
		return nRetValue;
	};

	CellEditor.prototype._getAutoComplete = function (str) {
		// ToDo can be sped up by searching each time not in a large array, but in a smaller one (by previous characters)
		//TODO I save the text!
		var oLastResult = this.objAutoComplete.get(str);
		if (oLastResult) {
			return oLastResult;
		}

		var arrAutoComplete = this.options.autoComplete;
		var arrAutoCompleteLC = this.options.autoCompleteLC;
		var i, length, arrResult = [];
		for (i = 0, length = arrAutoCompleteLC.length; i < length; ++i) {
			if (arrAutoCompleteLC[i].length !== str.length && 0 === arrAutoCompleteLC[i].indexOf(str)) {
				arrResult.push(arrAutoComplete[i]);
			}
		}
		this.objAutoComplete.set(str, arrResult);
		return arrResult;
	};

	CellEditor.prototype._updateSelectionInfo = function () {
		var f = this._findFragmentToInsertInto(this.cursorPos);
		if (!f) {
			return;
		}

		var xfs = new AscCommonExcel.CellXfs();
		xfs.setFont(this.newTextFormat || (this.options.fragments && this.options.fragments[f.index].format));
		this.handlers.trigger("updateEditorSelectionInfo", xfs);
	};

	CellEditor.prototype._checkMaxCellLength = function (length) {
		//TODO question, measure text length or number of characters
		var count = AscCommonExcel.getFragmentsCharCodesLength(this.options.fragments) + length - Asc.c_oAscMaxCellOrCommentLength;
		return 0 > count ? 0 : count;
	};

	// Event handlers
	CellEditor.prototype.executeShortcut = function(nShortcutAction) {
		let oResult = {keyResult: keydownresult_PreventAll};
		const oApi = window["Asc"]["editor"];
		const bHieroglyph = this.isTopLineActive && AscCommonExcel.getFragmentsLength(this.options.fragments) !== this.input.value.length;
		switch (nShortcutAction) {
			case Asc.c_oAscSpreadsheetShortcutType.Strikeout: {
				if (bHieroglyph) {
					this._syncEditors();
				}
				this.setTextStyle("s", null);
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.Bold: {
				if (bHieroglyph) {
					this._syncEditors();
				}
				this.setTextStyle("b", null);
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.Italic: {
				if (bHieroglyph) {
					this._syncEditors();
				}
				this.setTextStyle("i", null);
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.Underline: {
				if (bHieroglyph) {
					this._syncEditors();
				}
				this.setTextStyle("u", null);
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.EditSelectAll: {
				if (!this.hasFocus) {
					this.setFocus(true);
				}
				if (this.isTopLineActive) {
					oResult.keyResult = keydownresult_PreventNothing;
				}
				this._moveCursor(kBeginOfText);
				this._selectChars(kEndOfText);
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.EditUndo: {
				this.undo();
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.EditRedo: {
				this.redo();
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.CellInsertTime: {
				const oDate = new Asc.cDate();
				this._addChars(oDate.getTimeString(oApi));
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.CellInsertDate: {
				const oDate = new Asc.cDate();
				this._addChars(oDate.getDateString(oApi));
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.Print: {
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor: {
				if (!AscBrowser.isOpera) {
					oResult.keyResult = keydownresult_PreventNothing;
				}
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator: {
				this._addChars(oApi.asc_getDecimalSeparator());
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference: {
				const oRes = this._findRangeUnderCursor();
				if (oRes.range) {
					oRes.range.switchReference();
					// ToDo add change ref to other sheet
					this.changeCellRange(oRes.range);
				}
				break;
			}
			case Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize:
			case Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize: {
				if (bHieroglyph) {
					this._syncEditors();
				}
				this.setTextStyle("changeFontSize", nShortcutAction === Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
				break;
			}
			default: {
				const oCustom = oApi.getCustomShortcutAction(nShortcutAction);
				if (oCustom) {
					if (AscCommon.c_oAscCustomShortcutType.Symbol === oCustom.Type) {
						oApi["asc_insertSymbol"](oCustom.Font, oCustom.CharCode);
					}
				} else {
					oResult = null;
				}
				break;
			}
		}
		return oResult;
	};
	/**
	 *
	 * @param oEvent {AscCommon.CKeyboardEvent}
	 * @returns {number}
	 */
	CellEditor.prototype._onWindowKeyDown = function (oEvent) {
		const oThis = this;
		const oApi = window["Asc"]["editor"];

		let nRetValue = keydownresult_PreventNothing;

		if (this.handlers.trigger('getWizard') || !oThis.isOpened) {
			return nRetValue;
		}

		oThis._setSkipKeyPress(false);
		oThis.skipTLUpdate = false;

		// hieroglyph input definition
		const bHieroglyph = oThis.isTopLineActive && AscCommonExcel.getFragmentsLength(oThis.options.fragments) !== oThis.input.value.length;

		nRetValue = keydownresult_PreventKeyPress;
		const nShortcutAction = oApi.getShortcut(oEvent);
		const oShortcutRes = oThis.executeShortcut(nShortcutAction);
		if (oShortcutRes) {
			nRetValue = oShortcutRes.keyResult;
		} else {
			const bIsMacOs = AscCommon.AscBrowser.isMacOs;
			const bIsWordRemove = bIsMacOs ? oEvent.IsAlt() : oEvent.CtrlKey;
			switch (oEvent.GetKeyCode()) {
				case 27: { // "esc"

					if (oThis.handlers.trigger("isGlobalLockEditCell") || this.getMenuEditorMode()) {
						break;
					}
					oThis.close();
					nRetValue = keydownresult_PreventAll;
					break;
				}

				case 13: {  // "enter"
					if (window['IS_NATIVE_EDITOR']) {
						oThis._addNewLine();
					} else {
						if (!(oEvent.IsAlt() && oEvent.IsShift())) {
							if (oEvent.IsAlt()) {
								oThis._addNewLine();
							} else if (this.getMenuEditorMode()) {
								oThis._addNewLine();
							} else {
								if (false === oThis.handlers.trigger("isGlobalLockEditCell")) {
									if (oThis.textFlags) {
										oThis.textFlags.ctrlKey = oEvent.CtrlKey;
										oThis.textFlags.shiftKey = oEvent.IsShift();
									}
									oThis._tryCloseEditor(oEvent);
								}
							}
						}
					}
					nRetValue = keydownresult_PreventAll;
					break;
				}
				case 9: { // tab
					if (bHieroglyph) {
						oThis._syncEditors();
					}

					if (false === oThis.handlers.trigger("isGlobalLockEditCell")) {
						nRetValue = oThis._tryCloseEditor(oEvent);
					}
					break;
				}
				case 8: {  // "backspace"
					if (!this.enableKeyEvents) {
						break;
					}

					if (!window['IS_NATIVE_EDITOR']) {
						// Disable the browser's standard handling of pressing backspace
						nRetValue = keydownresult_PreventAll;
						if (bHieroglyph) {
							oThis._syncEditors();
						}
					}
					oThis._removeChars(bIsWordRemove ? kPrevWord : kPrevChar);
					break;
				}
				case 35: {  // "end"
					if (!this.enableKeyEvents) {
						break;
					}

					// Disable the browser's standard handling of pressing end
					nRetValue = keydownresult_PreventAll;
					if (!oThis.hasFocus) {
						break;
					}
					if (bHieroglyph) {
						oThis._syncEditors();
					}
					const nKind = oEvent.CtrlKey ? kEndOfText : kEndOfLine;
					oEvent.IsShift() ? oThis._selectChars(nKind) : oThis._moveCursor(nKind);
					break;
				}
				case 36: { // "home"
					if (!this.enableKeyEvents) {
						break;
					}

					// Disable the browser's standard handling of pressing home
					nRetValue = keydownresult_PreventAll;
					if (!oThis.hasFocus) {
						break;
					}
					if (bHieroglyph) {
						oThis._syncEditors();
					}
					const nKind = oEvent.CtrlKey ? kBeginOfText : kBeginOfLine;
					oEvent.IsShift() ? oThis._selectChars(nKind) : oThis._moveCursor(nKind);
					break;
				}
				case 37: { // "left"
					if (!this.enableKeyEvents) {
						this._delayedUpdateCursorByTopLine();
						break;
					}

					nRetValue = keydownresult_PreventAll;
					if (!oThis.hasFocus) {
						break;
					}
					if (bHieroglyph) {
						oThis._syncEditors();
					}
					if (bIsMacOs && oEvent.CtrlKey) {
						oEvent.IsShift() ? oThis._selectChars(kBeginOfLine) : oThis._moveCursor(kBeginOfLine);
					} else {
						const bWord = bIsMacOs ? oEvent.IsAlt() : oEvent.CtrlKey;
						const nKind = bWord ? kPrevWord : kPrevChar;
						oEvent.IsShift() ? oThis._selectChars(nKind) : oThis._moveCursor(nKind);
					}

					break;
				}
				case 38: {// "up"
					if (!this.enableKeyEvents) {
						this._delayedUpdateCursorByTopLine();
						break;
					}

					nRetValue = keydownresult_PreventAll;
					if (!oThis.hasFocus) {
						break;
					}
					if (bHieroglyph) {
						oThis._syncEditors();
					}
					oEvent.IsShift() ? oThis._selectChars(kPrevLine) : oThis._moveCursor(kPrevLine);
					break;
				}

				case 39: {// "right"
					if (!this.enableKeyEvents) {
						this._delayedUpdateCursorByTopLine();
						break;
					}

					nRetValue = keydownresult_PreventAll;
					if (!oThis.hasFocus) {
						break;
					}
					if (bHieroglyph) {
						oThis._syncEditors();
					}
					if (bIsMacOs && oEvent.CtrlKey) {
						oEvent.IsShift() ? oThis._selectChars(kEndOfLine) : oThis._moveCursor(kEndOfLine);
					} else {
						const bWord = bIsMacOs ? oEvent.IsAlt() : oEvent.CtrlKey;
						const nKind = bWord ? kNextWord : kNextChar;
						oEvent.IsShift() ? oThis._selectChars(nKind) : oThis._moveCursor(nKind);
					}
					break;
				}
				case 40: { // "down"
					if (!this.enableKeyEvents) {
						this._delayedUpdateCursorByTopLine();
						break;
					}

					nRetValue = keydownresult_PreventAll;
					if (!oThis.hasFocus) {
						break;
					}
					if (bHieroglyph) {
						oThis._syncEditors();
					}
					oEvent.IsShift() ? oThis._selectChars(kNextLine) : oThis._moveCursor(kNextLine);
					break;
				}
				case 46: {// "del"
					if (!this.enableKeyEvents || oEvent.IsShift()) {
						break;
					}

					if (bHieroglyph) {
						oThis._syncEditors();
					}
					nRetValue = keydownresult_PreventAll;
					oThis._removeChars(bIsWordRemove ? kNextWord : kNextChar);
					break;
				}
				case 144://Num Lock
				case 145: {//Scroll Lock
					if (AscBrowser.isOpera) {
						nRetValue = keydownresult_PreventAll;
					}
					break;
				}
				default: {
					nRetValue = keydownresult_PreventNothing;
					break;
				}
			}
		}

		if (nRetValue & keydownresult_PreventKeyPress) {
			oThis._setSkipKeyPress(true);
		}
		oThis.skipTLUpdate = true;
		return nRetValue;
	};

	/** @param event {KeyboardEvent} */
	CellEditor.prototype._onWindowKeyPress = function (event) {
		var t = this;

		if (!window['IS_NATIVE_EDITOR']) {
			if (event.KeyCode < 32 || t.skipKeyPress) {
				t._setSkipKeyPress(true);
				return true;
			}
		}

		let Code;
		if (null != event.Which) {
			Code = event.Which;
		} else if (event.KeyCode) {
			Code = event.KeyCode;
		} else {
			Code = 0;
		}

		return this.EnterText(Code);
	};

	CellEditor.prototype.EnterText = function (codePoints) {
		var t = this;

		if (!window['IS_NATIVE_EDITOR']) {
			if (!t.isOpened || !t.enableKeyEvents || this.handlers.trigger('getWizard')) {
				return true;
			}
			// hieroglyph input definition
			if (t.isTopLineActive && AscCommonExcel.getFragmentsLength(t.options.fragments) !== t.input.value.length) {
				t._syncEditors();
			}
		}

		t._setSkipKeyPress(false);

		//TODO Translation from code to symbols!
		var newChar;
		if (Array.isArray(codePoints)) {
			for (let nIdx = 0; nIdx < codePoints.length; ++nIdx) {
				newChar = String.fromCharCode(codePoints[nIdx]);
				t._addChars(newChar);
			}
		} else {
			newChar = String.fromCharCode(codePoints);
			t._addChars(newChar);
		}

		//TODO in case of adding an array - check - perhaps the part needs to be called every time after _addChars
		var tmpCursorPos;
		// The first time we enter quickly, we should add percentages at the end (for percentage format and only for numbers)
		if (t.options.isAddPersentFormat && AscCommon.isNumber(newChar)) {
			t.options.isAddPersentFormat = false;
			tmpCursorPos = t.cursorPos;
			t.undoMode = true;
			// add the percentage only to the line without a formula
			if (!t._formula) {
				t._addChars("%");
			}
			t.cursorPos = tmpCursorPos;
			t.undoMode = false;
			t._updateCursorPosition();
		}
		if (t.textRender.getEndOfText() === t.cursorPos && !t.isFormula()) {
			var s = AscCommonExcel.getFragmentsText(t.options.fragments);
			if (!AscCommon.isNumber(s) && s.length !== 0) {
				var arrAutoComplete = t._getAutoComplete(s.toLowerCase());
				var lengthInput = s.length;
				if (1 === arrAutoComplete.length) {
					var newValue = arrAutoComplete[0];
					tmpCursorPos = t.cursorPos;
					t._addChars(newValue.substring(lengthInput));
					t.selectionBegin = tmpCursorPos;
					t._selectChars(kEndOfText);
					this.sAutoComplete = newValue;
				}
			}
		}

		return t.isTopLineActive; // prevent event bubbling
	};

	/** @param event {KeyboardEvent} */
	CellEditor.prototype._onWindowKeyUp = function (event) {
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._onWindowMouseUp = function (event) {
		AscCommon.global_mouseEvent.UnLockMouse();
		if (c_oAscCellEditorSelectState.no !== this.isSelectMode) {
			this.cleanSelectRange();
		}
		this.isSelectMode = c_oAscCellEditorSelectState.no;
		if (this.callTopLineMouseup) {
			this._topLineMouseUp();
		}
		return true;
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._onWindowMouseMove = function (event) {
		if (c_oAscCellEditorSelectState.no !== this.isSelectMode && !this.hasCursor) {
			this._changeSelection(this._getCoordinates(event));
		}
		return true;
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._onMouseDown = function (event) {
		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus()) {
			return;
		}
		if (this.handlers.trigger('getWizard')) {
			return this.handlers.trigger('onMouseDown', event);
		}

		AscCommon.global_mouseEvent.LockMouse();

		var pos;
		var button = AscCommon.getMouseButton(event);
		var coord = this._getCoordinates(event);
		if (!window['IS_NATIVE_EDITOR']) {
			this.clickCounter.mouseDownEvent(coord.x, coord.y, button);
		}

		this.setFocus(true);

		this._updateTopLineActive(false);
		this.input.isFocused = false;

		if (0 === button) {
			if (1 === this.clickCounter.getClickCount() % 2) {
				this.isSelectMode = c_oAscCellEditorSelectState.char;
				if (!event.shiftKey) {
					this._updateCursor();
					pos = this._findCursorPosition(coord);
					if (pos !== undefined) {
						pos >= 0 ? this._moveCursor(kPosition, pos, this._findLineIndex(coord)) : this._moveCursor(pos, null, this._findLineIndex(coord));
					}
				} else {
					this._changeSelection(coord);
				}
			} else {
				// Dbl click
				this.isSelectMode = c_oAscCellEditorSelectState.word;

				let endWord, startWord;
				let fullString = AscCommonExcel.convertUnicodeToSimpleString(this.textRender.chars);
				let isNum = AscCommon.g_oFormatParser.isLocaleNumber(fullString);
				if (isNum) {
					// if we encounter a current numberDecimalSeparator in a number, we return the entire string as selection
					let splitIndex = fullString.indexOf(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);
					if (splitIndex !== -1) {
						endWord = fullString.length;
						startWord = 0;
					}
				}

				// End of the word
				endWord = endWord === undefined ? this.textRender.getNextWord(this.cursorPos) : endWord;
				// The beginning of the word (we look for the end, because we could get into a space)
				startWord = startWord === undefined ? this.textRender.getPrevWord(endWord) : startWord;

				this._moveCursor(kPosition, startWord);
				this._selectChars(kPosition, endWord);
			}
		} else if (2 === button) {
			this.handlers.trigger('onContextMenu', event);
		}
		return true;
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._onMouseUp = function (event) {
		var button = AscCommon.getMouseButton(event);
		AscCommon.global_mouseEvent.UnLockMouse();
		if (2 === button) {
			return true;
		}
		if (c_oAscCellEditorSelectState.no !== this.isSelectMode) {
			this.cleanSelectRange();
		}
		this.isSelectMode = c_oAscCellEditorSelectState.no;
		return true;
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._onMouseMove = function (event) {
		var coord = this._getCoordinates(event);
		this.clickCounter.mouseMoveEvent(coord.x, coord.y);
		this.hasCursor = true;
		if (c_oAscCellEditorSelectState.no !== this.isSelectMode) {
			this._changeSelection(coord);
		}
		return true;
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._onMouseLeave = function (event) {
		this.hasCursor = false;
		return true;
	};

	/** @param event {jQuery.Event} */
	CellEditor.prototype._onInputTextArea = function (event) {
		//TODO save the text!
		var t = this;
		if (!this.handlers.trigger("canEdit") || this.loadFonts) {
			return true;
		}
		if (this.handlers.trigger("isUserProtectActiveCell")) {
			this.handlers.trigger("asc_onError", Asc.c_oAscError.ID.ProtectedRangeByOtherUser, c_oAscError.Level.NoCritical);
			return true;
		}
		if (this.handlers.trigger("isProtectActiveCell")) {
			return true;
		}
		this.loadFonts = true;

		let checkedText = this.input.value.replace(/[\r\n]+/g, '');
		AscFonts.FontPickerByCharacter.checkText(checkedText, this, function () {
			t.loadFonts = false;
			t.skipTLUpdate = true;
			var length = t.replaceText(0, t.textRender.getEndOfText(), t.input.value);
			t._updateCursorByTopLine();

			if (length !== t.input.value.length) {
				t.input.value = AscCommonExcel.getFragmentsText((t.options.fragments));
				t._updateTopLineCurPos();
			}
		});
		return true;
	};

	/** @param event {MouseEvent} */
	CellEditor.prototype._getCoordinates = function (event) {
		if (window['IS_NATIVE_EDITOR']) {
			return {x: event.pageX, y: event.pageY};
		}

		var offs = AscCommon.UI.getBoundingClientRect(this.canvasOverlay);
		var x = (((event.pageX * AscBrowser.zoom) >> 0) - offs.left) / this.kx;
		var y = (((event.pageY * AscBrowser.zoom) >> 0) - offs.top) / this.ky;

		x *= AscCommon.AscBrowser.retinaPixelRatio;
		y *= AscCommon.AscBrowser.retinaPixelRatio;

		return {x: x, y: y};
	};

	CellEditor.prototype.getTextFromCharCodes = function (arrCharCodes) {
		//TODO save the text!
		var code, codePt, newText = '';
		for (var i = 0; i < arrCharCodes.length; ++i) {
			code = arrCharCodes[i];
			if (code < 0x10000) {
				newText += String.fromCharCode(code);
			} else {
				codePt = code - 0x10000;
				newText += String.fromCharCode(0xD800 + (codePt >> 10), 0xDC00 + (codePt & 0x3FF));
			}
		}
		return newText;
	};
	CellEditor.prototype.Begin_CompositeInput = function () {
		if (this.selectionBegin === this.selectionEnd) {
			this.beginCompositePos = this.cursorPos;
			this.compositeLength = 0;
		} else {
			this.beginCompositePos = Math.min(this.selectionBegin, this.selectionEnd);
			this.compositeLength = Math.max(this.selectionBegin, this.selectionEnd) - this.beginCompositePos;
		}
		this.setTextStyle('u', Asc.EUnderline.underlineSingle);
	};
	CellEditor.prototype.Replace_CompositeText = function (arrCharCodes) {
		if (!this.isOpened) {
			return;
		}

		var newText = this.getTextFromCharCodes(arrCharCodes);
		this.compositeLength = this.replaceText(this.beginCompositePos, this.compositeLength, newText);

		var tmpBegin = this.selectionBegin, tmpEnd = this.selectionEnd;

		this.selectionBegin = this.beginCompositePos;
		this.selectionEnd = this.beginCompositePos + this.compositeLength;
		this.setTextStyle('u', Asc.EUnderline.underlineSingle);

		this.selectionBegin = tmpBegin;
		this.selectionEnd = tmpEnd;

		// Refreshing the selection
		this._cleanSelection();
		this._drawSelection();
	};
	CellEditor.prototype.End_CompositeInput = function () {
		var tmpBegin = this.selectionBegin, tmpEnd = this.selectionEnd;

		//TODO linux(popOs + portuguese lang.) composite input - doesn't come Replace_CompositeText on remove chars
		let checkFragments = this._findFragment(this.beginCompositePos) && this._findFragment(this.beginCompositePos + this.compositeLength);
		if (checkFragments) {
			this.selectionBegin = this.beginCompositePos;
			this.selectionEnd = this.beginCompositePos + this.compositeLength;
		}
		this.setTextStyle('u', Asc.EUnderline.underlineNone);

		this.beginCompositePos = -1;
		this.compositeLength = 0;
		this.selectionBegin = tmpBegin;
		this.selectionEnd = tmpEnd;

		// Refreshing the selection
		this._cleanSelection();
		this._drawSelection();
	};
	CellEditor.prototype.isStartCompositeInput = function () {
		return this.beginCompositePos !== -1 && this.compositeLength !== 0;
	};
	CellEditor.prototype.Set_CursorPosInCompositeText = function (nPos) {
		if (-1 !== this.beginCompositePos) {
			nPos = Math.min(nPos, this.compositeLength);
			this._moveCursor(kPosition, this.beginCompositePos + nPos);
		}
	};
	CellEditor.prototype.Get_CursorPosInCompositeText = function () {
		return this.cursorPos - this.beginCompositePos;
	};
	CellEditor.prototype.Get_MaxCursorPosInCompositeText = function () {
		return this.compositeLength;
	};
	CellEditor.prototype.getMenuEditorMode = function () {
		return this.menuEditor;
	};
	CellEditor.prototype.selectAll = function () {
		//t.skipKeyPress
		var tmp = this.skipTLUpdate;
		this.skipTLUpdate = false;
		this._moveCursor(kBeginOfText);
		this._selectChars(kEndOfText);
		this.skipTLUpdate = tmp;
	};
	CellEditor.prototype._setSkipKeyPress = function (val) {
		this.skipKeyPress = val;
	};
	CellEditor.prototype.getText = function (start, len) {
		if (start == null) {
			start = 0;
		}
		if (len == null) {
			len = this.textRender.getCharsCount();
		}
		let chars = this.textRender.getChars(start, len);
		let res = "";
		for (let i in chars) {
			if (chars.hasOwnProperty(i)) {
				res += AscCommon.encodeSurrogateChar(chars[i]);
			}
		}
		return res;
	};

	CellEditor.prototype.getSelectionState = function () {
		return {start: this.selectionBegin, end: this.selectionEnd, cursor: this.cursorPos};
	};

	CellEditor.prototype.getSpeechDescription = function (prevState, curState, action) {
		if (curState.start === prevState.start && curState.end === prevState.end && prevState.cursor === curState.cursor) {
			return null;
		}

		let type = null, text = null, t = this;

		let compareSelection = function () {
			let _begin = Math.min(curState.start, curState.end);
			let _end = Math.max(curState.start, curState.end);
			let _start, _len;
			if (_end === _begin) {
				text = t.getText(t.cursorPos, 1);
				type = AscCommon.SpeechWorkerCommands.Text;
				return;
			}

			if (_end < prevState.start || prevState.end < _begin) {
				//no intersection
				//speech new select
				_start = _begin;
				_len = _end - _begin;
				type = AscCommon.SpeechWorkerCommands.Text;
			} else {
				if (_end !== prevState.end) {
					//changed end of text
					if (_end > prevState.end) {
						//added by select
						_start = prevState.end;
						_len = _end - prevState.end;
						type = AscCommon.SpeechWorkerCommands.TextSelected;
					} else {
						//deleted from select
						_start = _end;
						_len = prevState.end - _end;
						type = AscCommon.SpeechWorkerCommands.TextUnselected;
					}
				} else {
					if (_begin < prevState.start) {
						//added by select
						_start = _begin;
						_len = prevState.start - _begin;
						type = AscCommon.SpeechWorkerCommands.TextSelected;
					} else {
						//deleted from select
						_start = prevState.start;
						_len = _begin - prevState.start;
						type = AscCommon.SpeechWorkerCommands.TextUnselected;
					}
				}
			}

			text = t.getText(_start, _len);
		};

		let getWord = function () {
			let _cursorPos = t.cursorPos;
			type = AscCommon.SpeechWorkerCommands.Text;

			let _cursorPosNextWord = t.textRender.getNextWord(_cursorPos);
			text = t.getText(_cursorPos, _cursorPosNextWord - _cursorPos);
		};

		if (action) {
			let bWord = false;
			if (action.type !== AscCommon.SpeakerActionType.keyDown || action.event.KeyCode < 35 || action.event.KeyCode > 40) {
				return null;
			}

			if (!this.enableKeyEvents || !t.hasFocus) {
				return null;
			}

			let event = action.event;
			let isWizard = this.handlers.trigger('getWizard');
			if (!action.event || isWizard || !t.isOpened) {
				return null;
			}

			const bIsMacOs = AscCommon.AscBrowser.isMacOs;
			switch (event.GetKeyCode()) {
				case 8:   // "backspace"
					/*const bIsWord = bIsMacOs ? event.altKey : ctrlKey;
					t._removeChars(bIsWord ? kPrevWord : kPrevChar);*/
					break;
				case 35:  // "end"
					/*kind = ctrlKey ? kEndOfText : kEndOfLine;
					event.shiftKey ? t._selectChars(kind) : t._moveCursor(kind);
					return false;*/
					break;
				case 36:  // "home"
					/*kind = ctrlKey ? kBeginOfText : kBeginOfLine;
					event.shiftKey ? t._selectChars(kind) : t._moveCursor(kind);*/
					break;
				case 37:  // "left"
					if (bIsMacOs && event.CtrlKey) {
						//event.shiftKey ? t._selectChars(kBeginOfLine) : t._moveCursor(kBeginOfLine);
					} else {
						bWord = bIsMacOs ? event.AltKey : event.CtrlKey;
						/*kind = bWord ? kPrevWord : kPrevChar;
						event.shiftKey ? t._selectChars(kind) : t._moveCursor(kind);*/
					}

					break;
				case 38:  // "up"
					//event.shiftKey ? t._selectChars(kPrevLine) : t._moveCursor(kPrevLine);
					break;
				case 39:  // "right"
					if (bIsMacOs && event.CtrlKey) {
						//event.shiftKey ? t._selectChars(kEndOfLine) : t._moveCursor(kEndOfLine);
					} else {
						bWord = bIsMacOs ? event.AltKey : event.CtrlKey;
						/*kind = bWord ? kNextWord : kNextChar;
						event.shiftKey ? t._selectChars(kind) : t._moveCursor(kind);*/
					}
					break;

				case 40:  // "down"
					//event.shiftKey ? t._selectChars(kNextLine) : t._moveCursor(kNextLine);
					break;
			}

			if (bWord) {
				getWord();
			} else {
				compareSelection();
			}
		} else {
			compareSelection();
		}

		return type !== null ? {type: type, obj: {text: text}} : null;
	};

	CellEditor.prototype.setSelectionState = function (obj) {
		if (!obj) {
			return;
		}
		this.selectionBegin = obj.selectionBegin;
		this.selectionEnd = obj.selectionEnd;
		this.lastRangePos = obj.lastRangePos;
		this.lastRangeLength = obj.lastRangeLength;
		this.cursorPos = obj.cursorPos;
	};

	CellEditor.prototype.startAction = function () {
		var api = window["Asc"]["editor"];
		if (!api) {
			return;
		}
		api.sendEvent('asc_onUserActionStart');
	};

	CellEditor.prototype.endAction = function () {
		var api = window["Asc"]["editor"];
		if (!api) {
			return;
		}
		api.sendEvent('asc_onUserActionEnd');
	};

	CellEditor.prototype.openAction = function () {
		var api = window["Asc"]["editor"];
		if (!api) {
			return;
		}
		api.sendEvent('onOpenCellEditor');
	};

	CellEditor.prototype.closeAction = function () {
		var api = window["Asc"]["editor"];
		if (!api) {
			return;
		}
		api.sendEvent('onCloseCellEditor');
	};



	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window["AscCommonExcel"].CellEditor = CellEditor;
})(window);
