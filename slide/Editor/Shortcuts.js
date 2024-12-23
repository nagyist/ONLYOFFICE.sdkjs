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

(function(undefined) {
	const AscShortcutAction = Asc.CAscShortcutAction;
	const AscShortcut = Asc.CAscShortcut;
	const keyCodes = Asc.c_oAscKeyCodes;
	const c_oAscDefaultShortcuts = {};

	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFindDialog] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenFindDialog, [new AscShortcut(keyCodes.KeyF, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu, [new AscShortcut(keyCodes.KeyH, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentsPanel] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenCommentsPanel, [new AscShortcut(keyCodes.KeyH, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentField] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenCommentField, [new AscShortcut(keyCodes.KeyH, false, false, true, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenChatPanel] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenChatPanel, [new AscShortcut(keyCodes.KeyQ, false, false, true, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Save] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Save, [new AscShortcut(keyCodes.KeyS, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Print] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Print, [new AscShortcut(keyCodes.KeyP, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SaveAs] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SaveAs, [new AscShortcut(keyCodes.KeyS, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenHelpMenu] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenHelpMenu, [new AscShortcut(keyCodes.F1, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NextFileTab] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.NextFileTab, [new AscShortcut(keyCodes.Tab, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PreviousFileTab] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.PreviousFileTab, [new AscShortcut(keyCodes.Tab, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ShowContextMenu] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.ShowContextMenu, [new AscShortcut(keyCodes.F10, false, true, false, false), new AscShortcut(keyCodes.ContextMenu, false, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseMenu] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.CloseMenu, [new AscShortcut(keyCodes.Esc, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Zoom100] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Zoom100, [new AscShortcut(keyCodes.Digit0, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToFirstSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.GoToFirstSlide, [new AscShortcut(keyCodes.Home, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToLastSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.GoToLastSlide, [new AscShortcut(keyCodes.End, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToNextSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.GoToNextSlide, [new AscShortcut(keyCodes.PageDown, false, false, false, false),new AscShortcut(keyCodes.ArrowDown, false, false, false, false), new AscShortcut(keyCodes.ArrowRight, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToPreviousSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.GoToPreviousSlide, [new AscShortcut(keyCodes.PageUp, false, false, false, false), new AscShortcut(keyCodes.ArrowUp, false, false, false, false),new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomIn] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.ZoomIn, [new AscShortcut(keyCodes.KeyEqual, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomOut] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.ZoomOut, [new AscShortcut(keyCodes.KeyMinus, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NextModalControl] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.NextModalControl, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PreviousModalControl] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.PreviousModalControl, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NewSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.NewSlide, [new AscShortcut(keyCodes.KeyM, true, false, false, false),new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.RemoveSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.RemoveSlide, [new AscShortcut(keyCodes.Delete, false, false, false, false), new AscShortcut(keyCodes.Backspace, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Duplicate] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Duplicate, [new AscShortcut(keyCodes.KeyD, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToBegin] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveSlideToBegin, [new AscShortcut(keyCodes.ArrowUp, true, true, false, false), new AscShortcut(keyCodes.PageUp, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToEnd] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveSlideToEnd, [new AscShortcut(keyCodes.ArrowDown, true, true, false, false), new AscShortcut(keyCodes.PageDown, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditShape] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EditShape, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditChart] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EditChart, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Group] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Group, [new AscShortcut(keyCodes.KeyG, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnGroup] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.UnGroup, [new AscShortcut(keyCodes.KeyG, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft, [new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectRight] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.BigMoveObjectRight, [new AscShortcut(keyCodes.ArrowRight, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectUp] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.BigMoveObjectUp, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectDown] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.BigMoveObjectDown, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToNextCell] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToNextCell, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToPreviousCell] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToPreviousCell, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToNextRow] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToNextRow, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToPreviousRow] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToPreviousRow, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EndParagraphCell] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EndParagraphCell, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.AddNewRow] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.AddNewRow, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide, [new AscShortcut(keyCodes.Enter, false, false, false, false), new AscShortcut(keyCodes.PageDown, false, false, false, false), new AscShortcut(keyCodes.ArrowRight, false, false, false, false), new AscShortcut(keyCodes.ArrowDown, false, false, false, false), new AscShortcut(keyCodes.Space, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide, [new AscShortcut(keyCodes.PageUp, false, false, false, false), new AscShortcut(keyCodes.ArrowLeft, false, false, false, false), new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide, [new AscShortcut(keyCodes.Home, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide, [new AscShortcut(keyCodes.End, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview, [new AscShortcut(keyCodes.Esc, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditUndo] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EditUndo, [new AscShortcut(keyCodes.KeyZ, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditRedo] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EditRedo, [new AscShortcut(keyCodes.KeyY, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CopyFormat] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.CopyFormat, [new AscShortcut(keyCodes.KeyC, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteFormat] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.PasteFormat, [new AscShortcut(keyCodes.KeyV, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UseDestinationTheme] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.UseDestinationTheme, [new AscShortcut(keyCodes.KeyH, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.KeepSourceFormat] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.KeepSourceFormat, [new AscShortcut(keyCodes.KeyK, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteAsPicture] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.PasteAsPicture, [new AscShortcut(keyCodes.KeyU, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.KeepTextOnly] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.KeepTextOnly, [new AscShortcut(keyCodes.KeyT, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.AddHyperlink] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.AddHyperlink, [new AscShortcut(keyCodes.KeyK, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.VisitHyperlink] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.VisitHyperlink, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditSelectAll] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EditSelectAll, [new AscShortcut(keyCodes.KeyA, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectNextSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectNextSlide, [new AscShortcut(keyCodes.PageDown, false, true, false, false), new AscShortcut(keyCodes.ArrowDown, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectPreviousSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectPreviousSlide, [new AscShortcut(keyCodes.PageUp, false, true, false, false), new AscShortcut(keyCodes.ArrowUp, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToFirstSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectToFirstSlide, [new AscShortcut(keyCodes.Home, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToLastSlide] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectToLastSlide, [new AscShortcut(keyCodes.End, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToStartLine] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectToStartLine, [new AscShortcut(keyCodes.Home, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToEndLine] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectToEndLine, [new AscShortcut(keyCodes.End, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectRightChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectRightChar, [new AscShortcut(keyCodes.ArrowRight, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLeftChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectLeftChar, [new AscShortcut(keyCodes.ArrowLeft, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLineUp] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectLineUp, [new AscShortcut(keyCodes.ArrowUp, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLineDown] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectLineDown, [new AscShortcut(keyCodes.ArrowDown, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditDeselectAll] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EditDeselectAll, [new AscShortcut(keyCodes.Escape, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ShowParaMarks] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.ShowParaMarks, [new AscShortcut(keyCodes.Digit8, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Bold] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Bold, [new AscShortcut(keyCodes.KeyB, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Italic] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Italic, [new AscShortcut(keyCodes.KeyI, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Underline] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Underline, [new AscShortcut(keyCodes.KeyU, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Strikethrough] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Strikethrough, [new AscShortcut(keyCodes.Digit5, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Subscript] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Subscript, [new AscShortcut(keyCodes.Period, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Superscript] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Superscript, [new AscShortcut(keyCodes.Comma, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BulletList] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.BulletList, [new AscShortcut(keyCodes.KeyL, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ResetChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.ResetChar, [new AscShortcut(keyCodes.Space, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.IncreaseFont] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.IncreaseFont, [new AscShortcut(keyCodes.BracketRight, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DecreaseFont] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DecreaseFont, [new AscShortcut(keyCodes.BracketLeft, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CenterAlign] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.CenterAlign, [new AscShortcut(keyCodes.KeyE, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.JustifyAlign] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.JustifyAlign, [new AscShortcut(keyCodes.KeyJ, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.RightAlign] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.RightAlign, [new AscShortcut(keyCodes.KeyR, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LeftAlign] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LeftAlign, [new AscShortcut(keyCodes.KeyL, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Indent] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Indent, [new AscShortcut(keyCodes.KeyM, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnIndent] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.UnIndent, [new AscShortcut(keyCodes.KeyM, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteLeftChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DeleteLeftChar, [new AscShortcut(keyCodes.Backspace, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteRightChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DeleteRightChar, [new AscShortcut(keyCodes.Delete, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.StartIndent] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.StartIndent, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.StartUnIndent] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.StartUnIndent, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.InsertTab] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.InsertTab, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.InsertLineBreak] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.InsertLineBreak, [new AscShortcut(keyCodes.Enter, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EndParagraph] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EndParagraph, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EuroSign] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EuroSign, [new AscShortcut(keyCodes.KeyE, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NonBreakingSpace] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.NonBreakingSpace, [new AscShortcut(keyCodes.Space, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToRightChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToRightChar, [new AscShortcut(keyCodes.ArrowRight, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToLeftChar] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToLeftChar, [new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToUpLine] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToUpLine, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToDownLine] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToDownLine, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder, [new AscShortcut(keyCodes.Enter, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartLine] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToStartLine, [new AscShortcut(keyCodes.Home, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndLine] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToEndLine, [new AscShortcut(keyCodes.End, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartContent] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToStartContent, [new AscShortcut(keyCodes.Home, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndContent] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToEndContent, [new AscShortcut(keyCodes.End, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SpeechWorker] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SpeechWorker, [new AscShortcut(keyCodes.KeyZ, true, false, true, false)]);

	if (AscCommon.AscBrowser.isMacOs) {
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFilePanel] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenFilePanel, [new AscShortcut(keyCodes.KeyF, true, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenExistingFile] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenExistingFile, [new AscShortcut(keyCodes.KeyO, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseFile] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.CloseFile, [new AscShortcut(keyCodes.KeyW, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideUp] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveSlideUp, [new AscShortcut(keyCodes.ArrowUp, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideDown] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveSlideDown, [new AscShortcut(keyCodes.ArrowDown, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft, [new AscShortcut(keyCodes.ArrowLeft, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight, [new AscShortcut(keyCodes.ArrowRight, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp, [new AscShortcut(keyCodes.ArrowUp, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown, [new AscShortcut(keyCodes.ArrowDown, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation, [new AscShortcut(keyCodes.Enter, false, true, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Cut] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Cut, [new AscShortcut(keyCodes.KeyX, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Copy] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Copy, [new AscShortcut(keyCodes.KeyC, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Paste] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Paste, [new AscShortcut(keyCodes.KeyV, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat, [new AscShortcut(keyCodes.KeyV, false, true, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectRightWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectRightWord, [new AscShortcut(keyCodes.ArrowRight, false, true, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLeftWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectLeftWord, [new AscShortcut(keyCodes.ArrowLeft, false, true, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteLeftWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DeleteLeftWord, [new AscShortcut(keyCodes.Backspace, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteRightWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DeleteRightWord, [new AscShortcut(keyCodes.Delete, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToStartWord, [new AscShortcut(keyCodes.ArrowLeft, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToEndWord, [new AscShortcut(keyCodes.ArrowRight, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EnDash] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EnDash, [new AscShortcut(keyCodes.KeyMinus, false, false, true, false)]);

		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFindDialog].initShortcuts([new AscShortcut(keyCodes.KeyF, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentsPanel].initShortcuts([new AscShortcut(keyCodes.KeyH, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentField].initShortcuts([new AscShortcut(keyCodes.KeyA, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenChatPanel].initShortcuts([new AscShortcut(keyCodes.KeyQ, true, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Save].initShortcuts([new AscShortcut(keyCodes.KeyS, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Print].initShortcuts([new AscShortcut(keyCodes.KeyP, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SaveAs].initShortcuts([new AscShortcut(keyCodes.KeyS, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Zoom100].initShortcuts([new AscShortcut(keyCodes.Digit0, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomIn].initShortcuts([new AscShortcut(keyCodes.KeyEqual, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomOut].initShortcuts([new AscShortcut(keyCodes.KeyMinus, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NewSlide].initShortcuts([new AscShortcut(keyCodes.KeyM, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Duplicate].initShortcuts([new AscShortcut(keyCodes.KeyD, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToBegin].initShortcuts([new AscShortcut(keyCodes.ArrowUp, false, true, false, true), new AscShortcut(keyCodes.PageUp, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToEnd].initShortcuts([new AscShortcut(keyCodes.ArrowDown, false, true, false, true), new AscShortcut(keyCodes.PageDown, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Group].initShortcuts([new AscShortcut(keyCodes.KeyG, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnGroup].initShortcuts([new AscShortcut(keyCodes.KeyG, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditUndo].initShortcuts([new AscShortcut(keyCodes.KeyZ, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditRedo].initShortcuts([new AscShortcut(keyCodes.KeyY, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CopyFormat].initShortcuts([new AscShortcut(keyCodes.KeyC, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteFormat].initShortcuts([new AscShortcut(keyCodes.KeyV, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.AddHyperlink].initShortcuts([new AscShortcut(keyCodes.KeyK, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditSelectAll].initShortcuts([new AscShortcut(keyCodes.KeyA, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ShowParaMarks].initShortcuts([new AscShortcut(keyCodes.Digit8, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Bold].initShortcuts([new AscShortcut(keyCodes.KeyB, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Italic].initShortcuts([new AscShortcut(keyCodes.KeyI, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Underline].initShortcuts([new AscShortcut(keyCodes.KeyU, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Strikethrough].initShortcuts([new AscShortcut(keyCodes.Digit5, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Subscript].initShortcuts([new AscShortcut(keyCodes.Period, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Superscript].initShortcuts([new AscShortcut(keyCodes.Comma, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BulletList].initShortcuts([new AscShortcut(keyCodes.KeyL, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ResetChar].initShortcuts([new AscShortcut(keyCodes.Space, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.IncreaseFont].initShortcuts([new AscShortcut(keyCodes.BracketRight, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DecreaseFont].initShortcuts([new AscShortcut(keyCodes.BracketLeft, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CenterAlign].initShortcuts([new AscShortcut(keyCodes.KeyE, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.JustifyAlign].initShortcuts([new AscShortcut(keyCodes.KeyJ, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.RightAlign].initShortcuts([new AscShortcut(keyCodes.KeyR, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LeftAlign].initShortcuts([new AscShortcut(keyCodes.KeyL, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Indent].initShortcuts([new AscShortcut(keyCodes.KeyM, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnIndent].initShortcuts([new AscShortcut(keyCodes.KeyM, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EuroSign].initShortcuts([new AscShortcut(keyCodes.KeyE, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NonBreakingSpace].initShortcuts([new AscShortcut(keyCodes.Space, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder].initShortcuts([new AscShortcut(keyCodes.Enter, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SpeechWorker].initShortcuts([new AscShortcut(keyCodes.KeyZ, false, false, true, true)]);
	} else {
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFilePanel] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenFilePanel, [new AscShortcut(keyCodes.KeyF, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenExistingFile] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.OpenExistingFile, [new AscShortcut(keyCodes.KeyO, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseFile] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.CloseFile, [new AscShortcut(keyCodes.KeyW, true, false, false, false), new AscShortcut(keyCodes.F4, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideUp] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveSlideUp, [new AscShortcut(keyCodes.ArrowUp, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideDown] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveSlideDown, [new AscShortcut(keyCodes.ArrowDown, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft, [new AscShortcut(keyCodes.ArrowLeft, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight, [new AscShortcut(keyCodes.ArrowRight, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp, [new AscShortcut(keyCodes.ArrowUp, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown, [new AscShortcut(keyCodes.ArrowDown, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation, [new AscShortcut(keyCodes.F5, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Cut] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Cut, [new AscShortcut(keyCodes.KeyX, true, false, false, false), new AscShortcut(keyCodes.Delete, false, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Copy] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Copy, [new AscShortcut(keyCodes.KeyC, true, false, false, false), new AscShortcut(keyCodes.Insert, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Paste] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.Paste, [new AscShortcut(keyCodes.KeyV, true, false, false, false), new AscShortcut(keyCodes.Insert, false, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat, [new AscShortcut(keyCodes.KeyV, true, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectRightWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectRightWord, [new AscShortcut(keyCodes.ArrowRight, true, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLeftWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.SelectLeftWord, [new AscShortcut(keyCodes.ArrowLeft, true, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteLeftWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DeleteLeftWord, [new AscShortcut(keyCodes.Backspace, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteRightWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.DeleteRightWord, [new AscShortcut(keyCodes.Delete, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToStartWord, [new AscShortcut(keyCodes.ArrowLeft, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndWord] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.MoveToEndWord, [new AscShortcut(keyCodes.ArrowRight, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EnDash] = new AscShortcutAction(Asc.c_oAscPresentationShortcutType.EnDash, [new AscShortcut(keyCodes.KeyMinus, true, true, false, false)]);
	}
	function getStringFromShortcutType(type) {
		switch (type) {
			case Asc.c_oAscPresentationShortcutType.OpenFilePanel:
				return "OpenFilePanel";
			case Asc.c_oAscPresentationShortcutType.OpenFindDialog:
				return "OpenFindDialog";
			case Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu:
				return "OpenFindAndReplaceMenu";
			case Asc.c_oAscPresentationShortcutType.OpenCommentsPanel:
				return "OpenCommentsPanel";
			case Asc.c_oAscPresentationShortcutType.OpenCommentField:
				return "OpenCommentField";
			case Asc.c_oAscPresentationShortcutType.OpenChatPanel:
				return "OpenChatPanel";
			case Asc.c_oAscPresentationShortcutType.Save:
				return "Save";
			case Asc.c_oAscPresentationShortcutType.Print:
				return "Print";
			case Asc.c_oAscPresentationShortcutType.SaveAs:
				return "SaveAs";
			case Asc.c_oAscPresentationShortcutType.OpenHelpMenu:
				return "OpenHelpMenu";
			case Asc.c_oAscPresentationShortcutType.OpenExistingFile:
				return "OpenExistingFile";
			case Asc.c_oAscPresentationShortcutType.NextFileTab:
				return "NextFileTab";
			case Asc.c_oAscPresentationShortcutType.PreviousFileTab:
				return "PreviousFileTab";
			case Asc.c_oAscPresentationShortcutType.CloseFile:
				return "CloseFile";
			case Asc.c_oAscPresentationShortcutType.ShowContextMenu:
				return "ShowContextMenu";
			case Asc.c_oAscPresentationShortcutType.CloseMenu:
				return "CloseMenu";
			case Asc.c_oAscPresentationShortcutType.Zoom100:
				return "Zoom100";
			case Asc.c_oAscPresentationShortcutType.GoToFirstSlide:
				return "GoToFirstSlide";
			case Asc.c_oAscPresentationShortcutType.GoToLastSlide:
				return "GoToLastSlide";
			case Asc.c_oAscPresentationShortcutType.GoToNextSlide:
				return "GoToNextSlide";
			case Asc.c_oAscPresentationShortcutType.GoToPreviousSlide:
				return "GoToPreviousSlide";
			case Asc.c_oAscPresentationShortcutType.ZoomIn:
				return "ZoomIn";
			case Asc.c_oAscPresentationShortcutType.ZoomOut:
				return "ZoomOut";
			case Asc.c_oAscPresentationShortcutType.NextModalControl:
				return "NextModalControl";
			case Asc.c_oAscPresentationShortcutType.PreviousModalControl:
				return "PreviousModalControl";
			case Asc.c_oAscPresentationShortcutType.NewSlide:
				return "NewSlide";
			case Asc.c_oAscPresentationShortcutType.RemoveSlide:
				return "RemoveSlide";
			case Asc.c_oAscPresentationShortcutType.Duplicate:
				return "Duplicate";
			case Asc.c_oAscPresentationShortcutType.MoveSlideUp:
				return "MoveSlideUp";
			case Asc.c_oAscPresentationShortcutType.MoveSlideDown:
				return "MoveSlideDown";
			case Asc.c_oAscPresentationShortcutType.MoveSlideToBegin:
				return "MoveSlideToBegin";
			case Asc.c_oAscPresentationShortcutType.MoveSlideToEnd:
				return "MoveSlideToEnd";
			case Asc.c_oAscPresentationShortcutType.EditShape:
				return "EditShape";
			case Asc.c_oAscPresentationShortcutType.EditChart:
				return "EditChart";
			case Asc.c_oAscPresentationShortcutType.Group:
				return "Group";
			case Asc.c_oAscPresentationShortcutType.UnGroup:
				return "UnGroup";
			case Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject:
				return "MoveFocusToNextObject";
			case Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject:
				return "MoveFocusToPreviousObject";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft:
				return "LittleMoveObjectLeft";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight:
				return "LittleMoveObjectRight";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp:
				return "LittleMoveObjectUp";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown:
				return "LittleMoveObjectDown";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft:
				return "BigMoveObjectLeft";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectRight:
				return "BigMoveObjectRight";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectUp:
				return "BigMoveObjectUp";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectDown:
				return "BigMoveObjectDown";
			case Asc.c_oAscPresentationShortcutType.MoveToNextCell:
				return "MoveToNextCell";
			case Asc.c_oAscPresentationShortcutType.MoveToPreviousCell:
				return "MoveToPreviousCell";
			case Asc.c_oAscPresentationShortcutType.MoveToNextRow:
				return "MoveToNextRow";
			case Asc.c_oAscPresentationShortcutType.MoveToPreviousRow:
				return "MoveToPreviousRow";
			case Asc.c_oAscPresentationShortcutType.EndParagraphCell:
				return "EndParagraphCell";
			case Asc.c_oAscPresentationShortcutType.AddNewRow:
				return "AddNewRow";
			case Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation:
				return "DemonstrationStartPresentation";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide:
				return "DemonstrationGoToNextSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide:
				return "DemonstrationGoToPreviousSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide:
				return "DemonstrationGoToFirstSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide:
				return "DemonstrationGoToLastSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview:
				return "DemonstrationClosePreview";
			case Asc.c_oAscPresentationShortcutType.EditUndo:
				return "EditUndo";
			case Asc.c_oAscPresentationShortcutType.EditRedo:
				return "EditRedo";
			case Asc.c_oAscPresentationShortcutType.Cut:
				return "Cut";
			case Asc.c_oAscPresentationShortcutType.Copy:
				return "Copy";
			case Asc.c_oAscPresentationShortcutType.Paste:
				return "Paste";
			case Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat:
				return "PasteTextWithoutFormat";
			case Asc.c_oAscPresentationShortcutType.CopyFormat:
				return "CopyFormat";
			case Asc.c_oAscPresentationShortcutType.PasteFormat:
				return "PasteFormat";
			case Asc.c_oAscPresentationShortcutType.UseDestinationTheme:
				return "UseDestinationTheme";
			case Asc.c_oAscPresentationShortcutType.KeepSourceFormat:
				return "KeepSourceFormat";
			case Asc.c_oAscPresentationShortcutType.PasteAsPicture:
				return "PasteAsPicture";
			case Asc.c_oAscPresentationShortcutType.KeepTextOnly:
				return "KeepTextOnly";
			case Asc.c_oAscPresentationShortcutType.AddHyperlink:
				return "AddHyperlink";
			case Asc.c_oAscPresentationShortcutType.VisitHyperlink:
				return "VisitHyperlink";
			case Asc.c_oAscPresentationShortcutType.EditSelectAll:
				return "EditSelectAll";
			case Asc.c_oAscPresentationShortcutType.SelectNextSlide:
				return "SelectNextSlide";
			case Asc.c_oAscPresentationShortcutType.SelectPreviousSlide:
				return "SelectPreviousSlide";
			case Asc.c_oAscPresentationShortcutType.SelectToFirstSlide:
				return "SelectToFirstSlide";
			case Asc.c_oAscPresentationShortcutType.SelectToLastSlide:
				return "SelectToLastSlide";
			case Asc.c_oAscPresentationShortcutType.SelectToStartLine:
				return "SelectToStartLine";
			case Asc.c_oAscPresentationShortcutType.SelectToEndLine:
				return "SelectToEndLine";
			case Asc.c_oAscPresentationShortcutType.SelectRightChar:
				return "SelectRightChar";
			case Asc.c_oAscPresentationShortcutType.SelectLeftChar:
				return "SelectLeftChar";
			case Asc.c_oAscPresentationShortcutType.SelectRightWord:
				return "SelectRightWord";
			case Asc.c_oAscPresentationShortcutType.SelectLeftWord:
				return "SelectLeftWord";
			case Asc.c_oAscPresentationShortcutType.SelectLineUp:
				return "SelectLineUp";
			case Asc.c_oAscPresentationShortcutType.SelectLineDown:
				return "SelectLineDown";
			case Asc.c_oAscPresentationShortcutType.EditDeselectAll:
				return "EditDeselectAll";
			case Asc.c_oAscPresentationShortcutType.ShowParaMarks:
				return "ShowParaMarks";
			case Asc.c_oAscPresentationShortcutType.Bold:
				return "Bold";
			case Asc.c_oAscPresentationShortcutType.Italic:
				return "Italic";
			case Asc.c_oAscPresentationShortcutType.Underline:
				return "Underline";
			case Asc.c_oAscPresentationShortcutType.Strikethrough:
				return "Strikethrough";
			case Asc.c_oAscPresentationShortcutType.Subscript:
				return "Subscript";
			case Asc.c_oAscPresentationShortcutType.Superscript:
				return "Superscript";
			case Asc.c_oAscPresentationShortcutType.BulletList:
				return "BulletList";
			case Asc.c_oAscPresentationShortcutType.ResetChar:
				return "ResetChar";
			case Asc.c_oAscPresentationShortcutType.IncreaseFont:
				return "IncreaseFont";
			case Asc.c_oAscPresentationShortcutType.DecreaseFont:
				return "DecreaseFont";
			case Asc.c_oAscPresentationShortcutType.CenterAlign:
				return "CenterAlign";
			case Asc.c_oAscPresentationShortcutType.JustifyAlign:
				return "JustifyAlign";
			case Asc.c_oAscPresentationShortcutType.RightAlign:
				return "RightAlign";
			case Asc.c_oAscPresentationShortcutType.LeftAlign:
				return "LeftAlign";
			case Asc.c_oAscPresentationShortcutType.Indent:
				return "Indent";
			case Asc.c_oAscPresentationShortcutType.UnIndent:
				return "UnIndent";
			case Asc.c_oAscPresentationShortcutType.DeleteLeftChar:
				return "DeleteLeftChar";
			case Asc.c_oAscPresentationShortcutType.DeleteRightChar:
				return "DeleteRightChar";
			case Asc.c_oAscPresentationShortcutType.DeleteLeftWord:
				return "DeleteLeftWord";
			case Asc.c_oAscPresentationShortcutType.DeleteRightWord:
				return "DeleteRightWord";
			case Asc.c_oAscPresentationShortcutType.StartIndent:
				return "StartIndent";
			case Asc.c_oAscPresentationShortcutType.StartUnIndent:
				return "StartUnIndent";
			case Asc.c_oAscPresentationShortcutType.InsertTab:
				return "InsertTab";
			case Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder:
				return "EquationAddPlaceholder";
			case Asc.c_oAscPresentationShortcutType.InsertLineBreak:
				return "InsertLineBreak";
			case Asc.c_oAscPresentationShortcutType.EndParagraph:
				return "EndParagraph";
			case Asc.c_oAscPresentationShortcutType.EuroSign:
				return "EuroSign";
			case Asc.c_oAscPresentationShortcutType.EnDash:
				return "EnDash";
			case Asc.c_oAscPresentationShortcutType.NonBreakingSpace:
				return "NonBreakingSpace";
			case Asc.c_oAscPresentationShortcutType.MoveToRightChar:
				return "MoveToRightChar";
			case Asc.c_oAscPresentationShortcutType.MoveToLeftChar:
				return "MoveToLeftChar";
			case Asc.c_oAscPresentationShortcutType.MoveToUpLine:
				return "MoveToUpLine";
			case Asc.c_oAscPresentationShortcutType.MoveToDownLine:
				return "MoveToDownLine";
			case Asc.c_oAscPresentationShortcutType.MoveToStartWord:
				return "MoveToStartWord";
			case Asc.c_oAscPresentationShortcutType.MoveToEndWord:
				return "MoveToEndWord";
			case Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder:
				return "GoToNextPlaceholder";
			case Asc.c_oAscPresentationShortcutType.MoveToStartLine:
				return "MoveToStartLine";
			case Asc.c_oAscPresentationShortcutType.MoveToEndLine:
				return "MoveToEndLine";
			case Asc.c_oAscPresentationShortcutType.MoveToStartContent:
				return "MoveToStartContent";
			case Asc.c_oAscPresentationShortcutType.MoveToEndContent:
				return "MoveToEndContent";
			case Asc.c_oAscPresentationShortcutType.SpeechWorker:
				return "SpeechWorker";
		}
	}
	function getShortcutTypeFromString(str) {
		switch (str) {
			case "OpenFilePanel":
				return Asc.c_oAscPresentationShortcutType.OpenFilePanel;
			case "OpenFindDialog":
				return Asc.c_oAscPresentationShortcutType.OpenFindDialog;
			case "OpenFindAndReplaceMenu":
				return Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu;
			case "OpenCommentsPanel":
				return Asc.c_oAscPresentationShortcutType.OpenCommentsPanel;
			case "OpenCommentField":
				return Asc.c_oAscPresentationShortcutType.OpenCommentField;
			case "OpenChatPanel":
				return Asc.c_oAscPresentationShortcutType.OpenChatPanel;
			case "Save":
				return Asc.c_oAscPresentationShortcutType.Save;
			case "Print":
				return Asc.c_oAscPresentationShortcutType.Print;
			case "SaveAs":
				return Asc.c_oAscPresentationShortcutType.SaveAs;
			case "OpenHelpMenu":
				return Asc.c_oAscPresentationShortcutType.OpenHelpMenu;
			case "OpenExistingFile":
				return Asc.c_oAscPresentationShortcutType.OpenExistingFile;
			case "NextFileTab":
				return Asc.c_oAscPresentationShortcutType.NextFileTab;
			case "PreviousFileTab":
				return Asc.c_oAscPresentationShortcutType.PreviousFileTab;
			case "CloseFile":
				return Asc.c_oAscPresentationShortcutType.CloseFile;
			case "ShowContextMenu":
				return Asc.c_oAscPresentationShortcutType.ShowContextMenu;
			case "CloseMenu":
				return Asc.c_oAscPresentationShortcutType.CloseMenu;
			case "Zoom100":
				return Asc.c_oAscPresentationShortcutType.Zoom100;
			case "GoToFirstSlide":
				return Asc.c_oAscPresentationShortcutType.GoToFirstSlide;
			case "GoToLastSlide":
				return Asc.c_oAscPresentationShortcutType.GoToLastSlide;
			case "GoToNextSlide":
				return Asc.c_oAscPresentationShortcutType.GoToNextSlide;
			case "GoToPreviousSlide":
				return Asc.c_oAscPresentationShortcutType.GoToPreviousSlide;
			case "ZoomIn":
				return Asc.c_oAscPresentationShortcutType.ZoomIn;
			case "ZoomOut":
				return Asc.c_oAscPresentationShortcutType.ZoomOut;
			case "NextModalControl":
				return Asc.c_oAscPresentationShortcutType.NextModalControl;
			case "PreviousModalControl":
				return Asc.c_oAscPresentationShortcutType.PreviousModalControl;
			case "NewSlide":
				return Asc.c_oAscPresentationShortcutType.NewSlide;
			case "RemoveSlide":
				return Asc.c_oAscPresentationShortcutType.RemoveSlide;
			case "Duplicate":
				return Asc.c_oAscPresentationShortcutType.Duplicate;
			case "MoveSlideUp":
				return Asc.c_oAscPresentationShortcutType.MoveSlideUp;
			case "MoveSlideDown":
				return Asc.c_oAscPresentationShortcutType.MoveSlideDown;
			case "MoveSlideToBegin":
				return Asc.c_oAscPresentationShortcutType.MoveSlideToBegin;
			case "MoveSlideToEnd":
				return Asc.c_oAscPresentationShortcutType.MoveSlideToEnd;
			case "EditShape":
				return Asc.c_oAscPresentationShortcutType.EditShape;
			case "EditChart":
				return Asc.c_oAscPresentationShortcutType.EditChart;
			case "Group":
				return Asc.c_oAscPresentationShortcutType.Group;
			case "UnGroup":
				return Asc.c_oAscPresentationShortcutType.UnGroup;
			case "MoveFocusToNextObject":
				return Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject;
			case "MoveFocusToPreviousObject":
				return Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject;
			case "LittleMoveObjectLeft":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft;
			case "LittleMoveObjectRight":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight;
			case "LittleMoveObjectUp":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp;
			case "LittleMoveObjectDown":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown;
			case "BigMoveObjectLeft":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft;
			case "BigMoveObjectRight":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectRight;
			case "BigMoveObjectUp":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectUp;
			case "BigMoveObjectDown":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectDown;
			case "MoveToNextCell":
				return Asc.c_oAscPresentationShortcutType.MoveToNextCell;
			case "MoveToPreviousCell":
				return Asc.c_oAscPresentationShortcutType.MoveToPreviousCell;
			case "MoveToNextRow":
				return Asc.c_oAscPresentationShortcutType.MoveToNextRow;
			case "MoveToPreviousRow":
				return Asc.c_oAscPresentationShortcutType.MoveToPreviousRow;
			case "EndParagraphCell":
				return Asc.c_oAscPresentationShortcutType.EndParagraphCell;
			case "AddNewRow":
				return Asc.c_oAscPresentationShortcutType.AddNewRow;
			case "DemonstrationStartPresentation":
				return Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation;
			case "DemonstrationGoToNextSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide;
			case "DemonstrationGoToPreviousSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide;
			case "DemonstrationGoToFirstSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide;
			case "DemonstrationGoToLastSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide;
			case "DemonstrationClosePreview":
				return Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview;
			case "EditUndo":
				return Asc.c_oAscPresentationShortcutType.EditUndo;
			case "EditRedo":
				return Asc.c_oAscPresentationShortcutType.EditRedo;
			case "Cut":
				return Asc.c_oAscPresentationShortcutType.Cut;
			case "Copy":
				return Asc.c_oAscPresentationShortcutType.Copy;
			case "Paste":
				return Asc.c_oAscPresentationShortcutType.Paste;
			case "PasteTextWithoutFormat":
				return Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat;
			case "CopyFormat":
				return Asc.c_oAscPresentationShortcutType.CopyFormat;
			case "PasteFormat":
				return Asc.c_oAscPresentationShortcutType.PasteFormat;
			case "UseDestinationTheme":
				return Asc.c_oAscPresentationShortcutType.UseDestinationTheme;
			case "KeepSourceFormat":
				return Asc.c_oAscPresentationShortcutType.KeepSourceFormat;
			case "PasteAsPicture":
				return Asc.c_oAscPresentationShortcutType.PasteAsPicture;
			case "KeepTextOnly":
				return Asc.c_oAscPresentationShortcutType.KeepTextOnly;
			case "AddHyperlink":
				return Asc.c_oAscPresentationShortcutType.AddHyperlink;
			case "VisitHyperlink":
				return Asc.c_oAscPresentationShortcutType.VisitHyperlink;
			case "EditSelectAll":
				return Asc.c_oAscPresentationShortcutType.EditSelectAll;
			case "SelectNextSlide":
				return Asc.c_oAscPresentationShortcutType.SelectNextSlide;
			case "SelectPreviousSlide":
				return Asc.c_oAscPresentationShortcutType.SelectPreviousSlide;
			case "SelectToFirstSlide":
				return Asc.c_oAscPresentationShortcutType.SelectToFirstSlide;
			case "SelectToLastSlide":
				return Asc.c_oAscPresentationShortcutType.SelectToLastSlide;
			case "SelectToStartLine":
				return Asc.c_oAscPresentationShortcutType.SelectToStartLine;
			case "SelectToEndLine":
				return Asc.c_oAscPresentationShortcutType.SelectToEndLine;
			case "SelectRightChar":
				return Asc.c_oAscPresentationShortcutType.SelectRightChar;
			case "SelectLeftChar":
				return Asc.c_oAscPresentationShortcutType.SelectLeftChar;
			case "SelectRightWord":
				return Asc.c_oAscPresentationShortcutType.SelectRightWord;
			case "SelectLeftWord":
				return Asc.c_oAscPresentationShortcutType.SelectLeftWord;
			case "SelectLineUp":
				return Asc.c_oAscPresentationShortcutType.SelectLineUp;
			case "SelectLineDown":
				return Asc.c_oAscPresentationShortcutType.SelectLineDown;
			case "EditDeselectAll":
				return Asc.c_oAscPresentationShortcutType.EditDeselectAll;
			case "ShowParaMarks":
				return Asc.c_oAscPresentationShortcutType.ShowParaMarks;
			case "Bold":
				return Asc.c_oAscPresentationShortcutType.Bold;
			case "Italic":
				return Asc.c_oAscPresentationShortcutType.Italic;
			case "Underline":
				return Asc.c_oAscPresentationShortcutType.Underline;
			case "Strikethrough":
				return Asc.c_oAscPresentationShortcutType.Strikethrough;
			case "Subscript":
				return Asc.c_oAscPresentationShortcutType.Subscript;
			case "Superscript":
				return Asc.c_oAscPresentationShortcutType.Superscript;
			case "BulletList":
				return Asc.c_oAscPresentationShortcutType.BulletList;
			case "ResetChar":
				return Asc.c_oAscPresentationShortcutType.ResetChar;
			case "IncreaseFont":
				return Asc.c_oAscPresentationShortcutType.IncreaseFont;
			case "DecreaseFont":
				return Asc.c_oAscPresentationShortcutType.DecreaseFont;
			case "CenterAlign":
				return Asc.c_oAscPresentationShortcutType.CenterAlign;
			case "JustifyAlign":
				return Asc.c_oAscPresentationShortcutType.JustifyAlign;
			case "RightAlign":
				return Asc.c_oAscPresentationShortcutType.RightAlign;
			case "LeftAlign":
				return Asc.c_oAscPresentationShortcutType.LeftAlign;
			case "Indent":
				return Asc.c_oAscPresentationShortcutType.Indent;
			case "UnIndent":
				return Asc.c_oAscPresentationShortcutType.UnIndent;
			case "DeleteLeftChar":
				return Asc.c_oAscPresentationShortcutType.DeleteLeftChar;
			case "DeleteRightChar":
				return Asc.c_oAscPresentationShortcutType.DeleteRightChar;
			case "DeleteLeftWord":
				return Asc.c_oAscPresentationShortcutType.DeleteLeftWord;
			case "DeleteRightWord":
				return Asc.c_oAscPresentationShortcutType.DeleteRightWord;
			case "StartIndent":
				return Asc.c_oAscPresentationShortcutType.StartIndent;
			case "StartUnIndent":
				return Asc.c_oAscPresentationShortcutType.StartUnIndent;
			case "InsertTab":
				return Asc.c_oAscPresentationShortcutType.InsertTab;
			case "EquationAddPlaceholder":
				return Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder;
			case "InsertLineBreak":
				return Asc.c_oAscPresentationShortcutType.InsertLineBreak;
			case "EndParagraph":
				return Asc.c_oAscPresentationShortcutType.EndParagraph;
			case "EuroSign":
				return Asc.c_oAscPresentationShortcutType.EuroSign;
			case "EnDash":
				return Asc.c_oAscPresentationShortcutType.EnDash;
			case "NonBreakingSpace":
				return Asc.c_oAscPresentationShortcutType.NonBreakingSpace;
			case "MoveToRightChar":
				return Asc.c_oAscPresentationShortcutType.MoveToRightChar;
			case "MoveToLeftChar":
				return Asc.c_oAscPresentationShortcutType.MoveToLeftChar;
			case "MoveToUpLine":
				return Asc.c_oAscPresentationShortcutType.MoveToUpLine;
			case "MoveToDownLine":
				return Asc.c_oAscPresentationShortcutType.MoveToDownLine;
			case "MoveToStartWord":
				return Asc.c_oAscPresentationShortcutType.MoveToStartWord;
			case "MoveToEndWord":
				return Asc.c_oAscPresentationShortcutType.MoveToEndWord;
			case "GoToNextPlaceholder":
				return Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder;
			case "MoveToStartLine":
				return Asc.c_oAscPresentationShortcutType.MoveToStartLine;
			case "MoveToEndLine":
				return Asc.c_oAscPresentationShortcutType.MoveToEndLine;
			case "MoveToStartContent":
				return Asc.c_oAscPresentationShortcutType.MoveToStartContent;
			case "MoveToEndContent":
				return Asc.c_oAscPresentationShortcutType.MoveToEndContent;
			case "SpeechWorker":
				return Asc.c_oAscPresentationShortcutType.SpeechWorker;
		}
	}
	window["Asc"]["c_oAscDefaultShortcuts"] = window["Asc"].c_oAscDefaultShortcuts = c_oAscDefaultShortcuts;
	window["AscCommon"].getStringFromShortcutType = getStringFromShortcutType;
	window["AscCommon"].getShortcutTypeFromString = getShortcutTypeFromString;
})();
