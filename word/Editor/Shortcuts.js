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

(function (undefined) {
	function AscShortcut(keyCode, isCtrl, isShift, isAlt, isCommand) {

	}
	const ShortcutActionKeycodes = {};
	const keyCodes = {
		Digit0      : 48,
		Digit1      : 49,
		Digit2      : 50,
		Digit3      : 51,
		Digit4      : 52,
		Digit5      : 53,
		Digit6      : 54,
		Digit7      : 55,
		Digit8      : 56,
		Digit9      : 57,
		KeyA        : 65,
		KeyB        : 66,
		KeyC        : 67,
		KeyD        : 68,
		KeyE        : 69,
		KeyF        : 70,
		KeyG        : 71,
		KeyH        : 72,
		KeyI        : 73,
		KeyJ        : 74,
		KeyK        : 75,
		KeyL        : 76,
		KeyM        : 77,
		KeyN        : 78,
		KeyO        : 79,
		KeyP        : 80,
		KeyQ        : 81,
		KeyR        : 82,
		KeyS        : 83,
		KeyT        : 84,
		KeyU        : 85,
		KeyV        : 86,
		KeyW        : 87,
		KeyX        : 88,
		KeyY        : 89,
		KeyZ        : 90,
		KeyMinus    : 63,
		KeyEqual    : 61,
		Tab         : 9,
		Escape      : 27,
		Enter       : 13,
		Backspace   : 8,
		Delete      : 46,
		Space       : 32,
		Home        : 36,
		End         : 35,
		PageUp      : 33,
		PageDown    : 34,
		NumpadPlus  : 107,
		NumpadMinus : 109,
		ArrowLeft   : 37,
		ArrowRight  : 39,
		ArrowUp     : 38,
		ArrowDown   : 40,
		Period      : 190,
		Comma       : 188,
		BracketRight: 221,
		BracketLeft : 219,
		Numpad8     : 104,
		F1          : 112,
		F9          : 120,
		F10         : 121
	};
	if (AscCommon.AscBrowser.isMacOs) {
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenFilePanel] = [new AscShortcut(keyCodes.KeyF, true, false, true, false)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenFindDialog] = [new AscShortcut(keyCodes.KeyF, false, false, false, true)];

		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenFindAndReplaceMenu] = [new AscShortcut(keyCodes.KeyH, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenCommentsPanel] = [new AscShortcut(keyCodes.KeyH, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Save] = [new AscShortcut(keyCodes.KeyS, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint] = [new AscShortcut(keyCodes.KeyP, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SaveAs] = [new AscShortcut(keyCodes.KeyS, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenExistingFile] = [new AscShortcut(keyCodes.KeyO, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NextFileTab] = [new AscShortcut(keyCodes.Tab, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PreviousFileTab] = [new AscShortcut(keyCodes.Tab, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CloseFile] = [new AscShortcut(keyCodes.KeyW, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CloseMenu] = [new AscShortcut(keyCodes.Escape, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Zoom100] = [new AscShortcut(keyCodes.Digit0, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartDocument] = [new AscShortcut(keyCodes.Home, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToEndDocument] = [new AscShortcut(keyCodes.End, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage] = [new AscShortcut(keyCodes.PageUp, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartNextPage] = [new AscShortcut(keyCodes.PageDown, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ZoomIn] = [new AscShortcut(keyCodes.NumpadPlus, false, false, false, true), new AscShortcut(keyCodes.KeyEqual, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ZoomOut] = [new AscShortcut(keyCodes.KeyMinus, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartWord] = [new AscShortcut(keyCodes.ArrowLeft, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToEndWord] = [new AscShortcut(keyCodes.ArrowRight, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertColumnBreak] = [new AscShortcut(keyCodes.Enter, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DeleteLeftWord] = [new AscShortcut(keyCodes.Backspace, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DeleteRightWord] = [new AscShortcut(keyCodes.Delete, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NonBreakingSpace] = [new AscShortcut(keyCodes.Space, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NonBreakingHyphen] = [new AscShortcut(keyCodes.KeyMinus, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditUndo] = [new AscShortcut(keyCodes.KeyZ, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditRedo] = [new AscShortcut(keyCodes.KeyY, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Cut] = [new AscShortcut(keyCodes.KeyX, false, false, false, true), new AscShortcut(keyCodes.Delete, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Copy] = [new AscShortcut(keyCodes.KeyC, false, false, false, true), new AscShortcut(keyCodes.KeyInsert, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Paste] = [new AscShortcut(keyCodes.KeyV, false, false, false, true), new AscShortcut(keyCodes.KeyInsert, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat] = [new AscShortcut(keyCodes.KeyV, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CopyFormat] = [new AscShortcut(keyCodes.KeyC, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PasteFormat] = [new AscShortcut(keyCodes.KeyV, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepSourceFormat] = [new AscShortcut(keyCodes.KeyK, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepTextOnly] = [new AscShortcut(keyCodes.KeyT, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsOverwriteCells] = [new AscShortcut(keyCodes.KeyO, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsNestTable] = [new AscShortcut(keyCodes.KeyN, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertHyperlink] = [new AscShortcut(keyCodes.KeyK, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToStartDocument] = [new AscShortcut(keyCodes.Home, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToEndDocument] = [new AscShortcut(keyCodes.End, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectRightWord] = [new AscShortcut(keyCodes.ArrowRight, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectLeftWord] = [new AscShortcut(keyCodes.ArrowLeft, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage] = [new AscShortcut(keyCodes.PageUp, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage] = [new AscShortcut(keyCodes.PageDown, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Bold] = [new AscShortcut(keyCodes.KeyB, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Italic] = [new AscShortcut(keyCodes.KeyI, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Underline] = [new AscShortcut(keyCodes.KeyU, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Strikeout] = [new AscShortcut(keyCodes.Digit5, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Subscript] = [new AscShortcut(keyCodes.Period, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Superscript] = [new AscShortcut(keyCodes.Comma, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ApplyListBullet] = [new AscShortcut(keyCodes.KeyL, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ResetChar] = [new AscShortcut(keyCodes.Space, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.IncreaseFontSize] = [new AscShortcut(keyCodes.BracketRight, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DecreaseFontSize] = [new AscShortcut(keyCodes.BracketLeft, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CenterPara] = [new AscShortcut(keyCodes.KeyE, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.JustifyPara] = [new AscShortcut(keyCodes.KeyJ, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.RightPara] = [new AscShortcut(keyCodes.KeyR, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LeftPara] = [new AscShortcut(keyCodes.KeyL, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertPageBreak] = [new AscShortcut(keyCodes.Enter, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Indent] = [new AscShortcut(keyCodes.KeyM, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.UnIndent] = [new AscShortcut(keyCodes.KeyM, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertPageNumber] = [new AscShortcut(keyCodes.KeyP, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ShowAll] = [new AscShortcut(keyCodes.Numpad8, false, true, false, true), new AscShortcut(keyCodes.Digit8, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft] = [new AscShortcut(keyCodes.ArrowLeft, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight] = [new AscShortcut(keyCodes.ArrowRight, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp] = [new AscShortcut(keyCodes.ArrowUp, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown] = [new AscShortcut(keyCodes.ArrowDown, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertEndnoteNow] = [new AscShortcut(keyCodes.KeyD, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertFootnoteNow] = [new AscShortcut(keyCodes.KeyF, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertTableBreak] = [new AscShortcut(keyCodes.Enter, false, true, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EmDash] = [new AscShortcut(keyCodes.NumpadMinus, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EnDash] = [new AscShortcut(keyCodes.NumpadMinus, false, false, false, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CopyrightSign] = [new AscShortcut(keyCodes.KeyG, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EuroSign] = [new AscShortcut(keyCodes.KeyE, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.RegisteredSign] = [new AscShortcut(keyCodes.KeyR, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.TrademarkSign] = [new AscShortcut(keyCodes.KeyT, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.HorizontalEllipsis] = [new AscShortcut(keyCodes.Period, false, false, true, true)];
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpeechWorker] = [new AscShortcut(keyCodes.KeyZ, false, false, true, true)];
	} else {
		ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenFilePanel] = [new AscShortcut(keyCodes.KeyF, false, false, true, false)];
	}

	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenFindDialog] = [new AscShortcut(keyCodes.KeyF, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenFindAndReplaceMenu] = [new AscShortcut(keyCodes.KeyH, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenCommentsPanel] = [new AscShortcut(keyCodes.KeyH, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenCommentField] = [new AscShortcut(keyCodes.KeyH, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenChatPanel] = [new AscShortcut(keyCodes.KeyQ, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Save] = [new AscShortcut(keyCodes.KeyS, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint] = [new AscShortcut(keyCodes.KeyP, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SaveAs] = [new AscShortcut(keyCodes.KeyS, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenHelpMenu] = [new AscShortcut(keyCodes.F1, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenExistingFile] = [new AscShortcut(keyCodes.KeyO, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NextFileTab] = [new AscShortcut(keyCodes.Tab, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PreviousFileTab] = [new AscShortcut(keyCodes.Tab, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CloseFile] = [new AscShortcut(keyCodes.KeyW, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.OpenContextMenu] = [new AscShortcut(keyCodes.F10, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CloseMenu] = [new AscShortcut(keyCodes.Escape, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Zoom100] = [new AscShortcut(keyCodes.Digit0, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.UpdateFields] = [new AscShortcut(keyCodes.F9, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartLine] = [new AscShortcut(keyCodes.Home, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartDocument] = [new AscShortcut(keyCodes.Home, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToEndLine] = [new AscShortcut(keyCodes.End, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToEndDocument] = [new AscShortcut(keyCodes.End, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage] = [new AscShortcut(keyCodes.PageUp, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartNextPage] = [new AscShortcut(keyCodes.PageDown, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ScrollDown] = [new AscShortcut(keyCodes.PageDown, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ScrollUp] = [new AscShortcut(keyCodes.PageUp, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ZoomIn] = [new AscShortcut(keyCodes.NumpadPlus, true, false, false, false), new AscShortcut(keyCodes.KeyEqual, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ZoomOut] = [new AscShortcut(keyCodes.KeyMinus, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToRightChar] = [new AscShortcut(keyCodes.ArrowRight, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToLeftChar] = [new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToUpLine] = [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToDownLine] = [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToStartWord] = [new AscShortcut(keyCodes.ArrowLeft, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToEndWord] = [new AscShortcut(keyCodes.ArrowRight, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NextModalControl] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PreviousModalControl] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToLowerHeaderFooter] = [new AscShortcut(keyCodes.PageDown, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToUpperHeaderFooter] = [new AscShortcut(keyCodes.PageUp, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToLowerHeader] = [new AscShortcut(keyCodes.PageDown, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToUpperHeader] = [new AscShortcut(keyCodes.PageUp, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EndParagraph] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertLineBreak] = [new AscShortcut(keyCodes.Enter, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertColumnBreak] = [new AscShortcut(keyCodes.Enter, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EquationAddPlaceholder] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentLeft] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentRight] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DeleteLeftChar] = [new AscShortcut(keyCodes.Backspace, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DeleteRightChar] = [new AscShortcut(keyCodes.Delete, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DeleteLeftWord] = [new AscShortcut(keyCodes.Backspace, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DeleteRightWord] = [new AscShortcut(keyCodes.Delete, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NonBreakingSpace] = [new AscShortcut(keyCodes.Space, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.NonBreakingHyphen] = [new AscShortcut(keyCodes.KeyMinus, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditUndo] = [new AscShortcut(keyCodes.KeyZ, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditRedo] = [new AscShortcut(keyCodes.KeyY, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Cut] = [new AscShortcut(keyCodes.KeyX, true, false, false, false), new AscShortcut(keyCodes.Delete, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Copy] = [new AscShortcut(keyCodes.KeyC, true, false, false, false), new AscShortcut(keyCodes.KeyInsert, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Paste] = [new AscShortcut(keyCodes.KeyV, true, false, false, false), new AscShortcut(keyCodes.KeyInsert, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat] = [new AscShortcut(keyCodes.KeyV, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CopyFormat] = [new AscShortcut(keyCodes.KeyC, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.PasteFormat] = [new AscShortcut(keyCodes.KeyV, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepSourceFormat] = [new AscShortcut(keyCodes.KeyK, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepTextOnly] = [new AscShortcut(keyCodes.KeyT, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsOverwriteCells] = [new AscShortcut(keyCodes.KeyO, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpecialOptionsNestTable] = [new AscShortcut(keyCodes.KeyN, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertHyperlink] = [new AscShortcut(keyCodes.KeyK, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.VisitHyperlink] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditSelectAll] = [new AscShortcut(keyCodes.KeyA, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToStartLine] = [new AscShortcut(keyCodes.Home, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToEndLine] = [new AscShortcut(keyCodes.End, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToStartDocument] = [new AscShortcut(keyCodes.Home, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToEndDocument] = [new AscShortcut(keyCodes.End, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectRightChar] = [new AscShortcut(keyCodes.ArrowRight, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectLeftChar] = [new AscShortcut(keyCodes.ArrowLeft, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectRightWord] = [new AscShortcut(keyCodes.ArrowRight, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectLeftWord] = [new AscShortcut(keyCodes.ArrowLeft, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectLineUp] = [new AscShortcut(keyCodes.ArrowUp, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectLineDown] = [new AscShortcut(keyCodes.ArrowDown, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectPageUp] = [new AscShortcut(keyCodes.PageUp, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectPageDown] = [new AscShortcut(keyCodes.PageDown, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage] = [new AscShortcut(keyCodes.PageUp, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage] = [new AscShortcut(keyCodes.PageDown, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Bold] = [new AscShortcut(keyCodes.KeyB, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Italic] = [new AscShortcut(keyCodes.KeyI, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Underline] = [new AscShortcut(keyCodes.KeyU, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Strikeout] = [new AscShortcut(keyCodes.Digit5, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Subscript] = [new AscShortcut(keyCodes.Period, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Superscript] = [new AscShortcut(keyCodes.Comma, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ApplyHeading1] = [new AscShortcut(keyCodes.Digit1, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ApplyHeading2] = [new AscShortcut(keyCodes.Digit2, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ApplyHeading3] = [new AscShortcut(keyCodes.Digit3, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ApplyListBullet] = [new AscShortcut(keyCodes.KeyL, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ResetChar] = [new AscShortcut(keyCodes.Space, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.IncreaseFontSize] = [new AscShortcut(keyCodes.BracketRight, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.DecreaseFontSize] = [new AscShortcut(keyCodes.BracketLeft, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CenterPara] = [new AscShortcut(keyCodes.KeyE, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.JustifyPara] = [new AscShortcut(keyCodes.KeyJ, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.RightPara] = [new AscShortcut(keyCodes.KeyR, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LeftPara] = [new AscShortcut(keyCodes.KeyL, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertPageBreak] = [new AscShortcut(keyCodes.Enter, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.Indent] = [new AscShortcut(keyCodes.KeyM, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.UnIndent] = [new AscShortcut(keyCodes.KeyM, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertPageNumber] = [new AscShortcut(keyCodes.KeyP, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ShowAll] = [new AscShortcut(keyCodes.Numpad8, true, true, false, false), new AscShortcut(keyCodes.Digit8, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.StartIndent] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.StartUnIndent] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertTab] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MixedIndent] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MixedUnIndent] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditShape] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EditChart] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft] = [new AscShortcut(keyCodes.ArrowLeft, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight] = [new AscShortcut(keyCodes.ArrowRight, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp] = [new AscShortcut(keyCodes.ArrowUp, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown] = [new AscShortcut(keyCodes.ArrowDown, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.BigMoveObjectLeft] = [new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.BigMoveObjectRight] = [new AscShortcut(keyCodes.ArrowRight, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.BigMoveObjectUp] = [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.BigMoveObjectDown] = [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveFocusToNextObject] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveFocusToPreviousObject] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertEndnoteNow] = [new AscShortcut(keyCodes.KeyD, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertFootnoteNow] = [new AscShortcut(keyCodes.KeyF, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToNextCell] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToPreviousCell] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToNextRow] = [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToPreviousRow] = [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EndParagraphCell] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.AddNewRow] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertTableBreak] = [new AscShortcut(keyCodes.Enter, true, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToNextForm] = [new AscShortcut(keyCodes.Tab, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.MoveToPreviousForm] = [new AscShortcut(keyCodes.Tab, false, true, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ChooseNextComboBoxOption] = [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ChoosePreviousComboBoxOption] = [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertLineBreakMultilineForm] = [new AscShortcut(keyCodes.Enter, false, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.InsertEquation] = [new AscShortcut(keyCodes.KeyEqual, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EmDash] = [new AscShortcut(keyCodes.NumpadMinus, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EnDash] = [new AscShortcut(keyCodes.NumpadMinus, true, false, false, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.CopyrightSign] = [new AscShortcut(keyCodes.KeyG, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.EuroSign] = [new AscShortcut(keyCodes.KeyE, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.RegisteredSign] = [new AscShortcut(keyCodes.KeyR, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.TrademarkSign] = [new AscShortcut(keyCodes.KeyT, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.HorizontalEllipsis] = [new AscShortcut(keyCodes.Period, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol] = [new AscShortcut(keyCodes.KeyX, false, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SpeechWorker] = [new AscShortcut(keyCodes.KeyZ, true, false, true, false)];
	ShortcutActionKeycodes[Asc.c_oAscDocumentShortcutType.SoftHyphen] = [];
})();
