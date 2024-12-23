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
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenFindDialog] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenFindDialog, [new AscShortcut(keyCodes.KeyF, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenFindAndReplaceMenu] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenFindAndReplaceMenu, [new AscShortcut(keyCodes.KeyH, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenCommentsPanel] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenCommentsPanel, [new AscShortcut(keyCodes.KeyH, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Save] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Save, [new AscShortcut(keyCodes.KeyS, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint, [new AscShortcut(keyCodes.KeyP, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SaveAs] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SaveAs, [new AscShortcut(keyCodes.KeyS, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenHelpMenu] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenHelpMenu, [new AscShortcut(keyCodes.F1, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.NextFileTab] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.NextFileTab, [new AscShortcut(keyCodes.Tab, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PreviousFileTab] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.PreviousFileTab, [new AscShortcut(keyCodes.Tab, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenContextMenu] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenContextMenu, [new AscShortcut(keyCodes.F10, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CloseMenu] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.CloseMenu, [new AscShortcut(keyCodes.Escape, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Zoom100] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Zoom100, [new AscShortcut(keyCodes.Digit0, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.UpdateFields] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.UpdateFields, [new AscShortcut(keyCodes.F9, false, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartLine] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartLine, [new AscShortcut(keyCodes.Home, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartDocument] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartDocument, [new AscShortcut(keyCodes.Home, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToEndLine] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToEndLine, [new AscShortcut(keyCodes.End, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToEndDocument] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToEndDocument, [new AscShortcut(keyCodes.End, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ScrollDown] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ScrollDown, [new AscShortcut(keyCodes.PageDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ScrollUp] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ScrollUp, [new AscShortcut(keyCodes.PageUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ZoomIn] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ZoomIn, [new AscShortcut(keyCodes.KeyEqual, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ZoomOut] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ZoomOut, [new AscShortcut(keyCodes.KeyMinus, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToRightChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToRightChar, [new AscShortcut(keyCodes.ArrowRight, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToLeftChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToLeftChar, [new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToUpLine] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToUpLine, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToDownLine] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToDownLine, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.NextModalControl] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.NextModalControl, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PreviousModalControl] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.PreviousModalControl, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToLowerHeaderFooter] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToLowerHeaderFooter, [new AscShortcut(keyCodes.PageDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToUpperHeaderFooter] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToUpperHeaderFooter, [new AscShortcut(keyCodes.PageUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToLowerHeader] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToLowerHeader, [new AscShortcut(keyCodes.PageDown, false, false, true, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToUpperHeader] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToUpperHeader, [new AscShortcut(keyCodes.PageUp, false, false, true, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EndParagraph] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EndParagraph, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertLineBreak] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertLineBreak, [new AscShortcut(keyCodes.Enter, false, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertColumnBreak] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertColumnBreak, [new AscShortcut(keyCodes.Enter, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EquationAddPlaceholder] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EquationAddPlaceholder, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentLeft] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentLeft, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentRight] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentRight, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DeleteLeftChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DeleteLeftChar, [new AscShortcut(keyCodes.Backspace, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DeleteRightChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DeleteRightChar, [new AscShortcut(keyCodes.Delete, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.NonBreakingSpace] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.NonBreakingSpace, [new AscShortcut(keyCodes.Space, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.NonBreakingHyphen] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.NonBreakingHyphen, [new AscShortcut(keyCodes.KeyMinus, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditUndo] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EditUndo, [new AscShortcut(keyCodes.KeyZ, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditRedo] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EditRedo, [new AscShortcut(keyCodes.KeyY, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CopyFormat] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.CopyFormat, [new AscShortcut(keyCodes.KeyC, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PasteFormat] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.PasteFormat, [new AscShortcut(keyCodes.KeyV, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepSourceFormat] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepSourceFormat, [new AscShortcut(keyCodes.KeyK, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepTextOnly] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepTextOnly, [new AscShortcut(keyCodes.KeyT, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SpecialOptionsOverwriteCells] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SpecialOptionsOverwriteCells, [new AscShortcut(keyCodes.KeyO, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SpecialOptionsNestTable] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SpecialOptionsNestTable, [new AscShortcut(keyCodes.KeyN, true, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertHyperlink] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertHyperlink, [new AscShortcut(keyCodes.KeyK, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.VisitHyperlink] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.VisitHyperlink, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditSelectAll] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EditSelectAll, [new AscShortcut(keyCodes.KeyA, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToStartLine] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectToStartLine, [new AscShortcut(keyCodes.Home, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToEndLine] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectToEndLine, [new AscShortcut(keyCodes.End, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToStartDocument] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectToStartDocument, [new AscShortcut(keyCodes.Home, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToEndDocument] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectToEndDocument, [new AscShortcut(keyCodes.End, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectRightChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectRightChar, [new AscShortcut(keyCodes.ArrowRight, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectLeftChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectLeftChar, [new AscShortcut(keyCodes.ArrowLeft, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectLineUp] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectLineUp, [new AscShortcut(keyCodes.ArrowUp, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectLineDown] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectLineDown, [new AscShortcut(keyCodes.ArrowDown, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectPageUp] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectPageUp, [new AscShortcut(keyCodes.PageUp, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectPageDown] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectPageDown, [new AscShortcut(keyCodes.PageDown, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage, [new AscShortcut(keyCodes.PageUp, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage, [new AscShortcut(keyCodes.PageDown, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Bold] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Bold, [new AscShortcut(keyCodes.KeyB, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Italic] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Italic, [new AscShortcut(keyCodes.KeyI, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Underline] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Underline, [new AscShortcut(keyCodes.KeyU, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Subscript] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Subscript, [new AscShortcut(keyCodes.Period, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Superscript] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Superscript, [new AscShortcut(keyCodes.Comma, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyListBullet] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyListBullet, [new AscShortcut(keyCodes.KeyL, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ResetChar] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ResetChar, [new AscShortcut(keyCodes.Space, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.IncreaseFontSize] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.IncreaseFontSize, [new AscShortcut(keyCodes.BracketRight, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DecreaseFontSize] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DecreaseFontSize, [new AscShortcut(keyCodes.BracketLeft, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CenterPara] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.CenterPara, [new AscShortcut(keyCodes.KeyE, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.JustifyPara] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.JustifyPara, [new AscShortcut(keyCodes.KeyJ, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.RightPara] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.RightPara, [new AscShortcut(keyCodes.KeyR, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LeftPara] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LeftPara, [new AscShortcut(keyCodes.KeyL, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertPageBreak] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertPageBreak, [new AscShortcut(keyCodes.Enter, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Indent] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Indent, [new AscShortcut(keyCodes.KeyM, true, false, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.UnIndent] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.UnIndent, [new AscShortcut(keyCodes.KeyM, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertPageNumber] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertPageNumber, [new AscShortcut(keyCodes.KeyP, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ShowAll] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ShowAll, [new AscShortcut(keyCodes.Digit8, true, true, false, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.StartIndent] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.StartIndent, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.StartUnIndent] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.StartUnIndent, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertTab] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertTab, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MixedIndent] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MixedIndent, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MixedUnIndent] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MixedUnIndent, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditShape] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EditShape, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditChart] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EditChart, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.BigMoveObjectLeft] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.BigMoveObjectLeft, [new AscShortcut(keyCodes.ArrowLeft, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.BigMoveObjectRight] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.BigMoveObjectRight, [new AscShortcut(keyCodes.ArrowRight, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.BigMoveObjectUp] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.BigMoveObjectUp, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.BigMoveObjectDown] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.BigMoveObjectDown, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveFocusToNextObject] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveFocusToNextObject, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveFocusToPreviousObject] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveFocusToPreviousObject, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertEndnoteNow] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertEndnoteNow, [new AscShortcut(keyCodes.KeyD, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToNextCell] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToNextCell, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToPreviousCell] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToPreviousCell, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToNextRow] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToNextRow, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToPreviousRow] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToPreviousRow, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EndParagraphCell] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EndParagraphCell, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.AddNewRow] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.AddNewRow, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertTableBreak] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertTableBreak, [new AscShortcut(keyCodes.Enter, true, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToNextForm] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToNextForm, [new AscShortcut(keyCodes.Tab, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToPreviousForm] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToPreviousForm, [new AscShortcut(keyCodes.Tab, false, true, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ChooseNextComboBoxOption] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ChooseNextComboBoxOption, [new AscShortcut(keyCodes.ArrowDown, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ChoosePreviousComboBoxOption] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ChoosePreviousComboBoxOption, [new AscShortcut(keyCodes.ArrowUp, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertLineBreakMultilineForm] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertLineBreakMultilineForm, [new AscShortcut(keyCodes.Enter, false, false, false, false)], true);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CopyrightSign] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.CopyrightSign, [new AscShortcut(keyCodes.KeyG, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EuroSign] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EuroSign, [new AscShortcut(keyCodes.KeyE, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.RegisteredSign] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.RegisteredSign, [new AscShortcut(keyCodes.KeyR, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.TrademarkSign] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.TrademarkSign, [new AscShortcut(keyCodes.KeyT, true, false, true, false)]);
	c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SpeechWorker] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SpeechWorker, [new AscShortcut(keyCodes.KeyZ, true, false, true, false)]);

	if (AscCommon.AscBrowser.isMacOs) {
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenFilePanel] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenFilePanel, [new AscShortcut(keyCodes.KeyF, true, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenCommentField] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenCommentField, [new AscShortcut(keyCodes.KeyA, false, false, true, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenChatPanel] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenChatPanel, [new AscShortcut(keyCodes.KeyQ, true, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenExistingFile] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenExistingFile, [new AscShortcut(keyCodes.KeyO, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CloseFile] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.CloseFile, [new AscShortcut(keyCodes.KeyW, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage, [new AscShortcut(keyCodes.PageUp, false, false, true, false), new AscShortcut(keyCodes.PageUp, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartNextPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartNextPage, [new AscShortcut(keyCodes.PageDown, false, false, true, false), new AscShortcut(keyCodes.PageDown, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartWord, [new AscShortcut(keyCodes.ArrowLeft, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToEndWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToEndWord, [new AscShortcut(keyCodes.ArrowRight, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DeleteLeftWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DeleteLeftWord, [new AscShortcut(keyCodes.Backspace, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DeleteRightWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DeleteRightWord, [new AscShortcut(keyCodes.Delete, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Cut] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Cut, [new AscShortcut(keyCodes.KeyX, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Copy] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Copy, [new AscShortcut(keyCodes.KeyC, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Paste] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Paste, [new AscShortcut(keyCodes.KeyV, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat, [new AscShortcut(keyCodes.KeyV, false, true, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Strikeout] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Strikeout, [new AscShortcut(keyCodes.KeyX, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectRightWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectRightWord, [new AscShortcut(keyCodes.ArrowRight, false, true, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectLeftWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectLeftWord, [new AscShortcut(keyCodes.ArrowLeft, false, true, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyHeading1] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyHeading1, [new AscShortcut(keyCodes.Digit1, true, false, true, false), new AscShortcut(keyCodes.Digit1, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyHeading2] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyHeading2, [new AscShortcut(keyCodes.Digit2, true, false, true, false), new AscShortcut(keyCodes.Digit2, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyHeading3] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyHeading3, [new AscShortcut(keyCodes.Digit3, true, false, true, false), new AscShortcut(keyCodes.Digit3, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft, [new AscShortcut(keyCodes.ArrowLeft, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight, [new AscShortcut(keyCodes.ArrowRight, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp, [new AscShortcut(keyCodes.ArrowUp, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown, [new AscShortcut(keyCodes.ArrowDown, false, false, false, true)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertEquation] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertEquation, [new AscShortcut(keyCodes.KeyEqual, true, false, true, false), new AscShortcut(keyCodes.KeyEqual, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EmDash] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EmDash, [new AscShortcut(keyCodes.KeyMinus, false, true, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EnDash] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EnDash, [new AscShortcut(keyCodes.KeyMinus, false, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.HorizontalEllipsis] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.HorizontalEllipsis, [new AscShortcut(keyCodes.KeySemicolon, false, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol, [new AscShortcut(keyCodes.KeyX, false, false, true, true), new AscShortcut(keyCodes.KeyX, true, false, true, false)], true);
		// c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SoftHyphen] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SoftHyphen, [new AscShortcut(keyCodes.KeyMinus, true, false, true, false), new AscShortcut(keyCodes.KeyMinus, false, false, true, true)],false, true);

		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenFindDialog].initShortcuts([new AscShortcut(keyCodes.KeyF, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenCommentsPanel].initShortcuts([new AscShortcut(keyCodes.KeyH, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Save].initShortcuts([new AscShortcut(keyCodes.KeyS, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint].initShortcuts([new AscShortcut(keyCodes.KeyP, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SaveAs].initShortcuts([new AscShortcut(keyCodes.KeyS, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CloseMenu].initShortcuts([new AscShortcut(keyCodes.Escape, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Zoom100].initShortcuts([new AscShortcut(keyCodes.Digit0, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartDocument].initShortcuts([new AscShortcut(keyCodes.Home, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToEndDocument].initShortcuts([new AscShortcut(keyCodes.End, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ZoomIn].initShortcuts([new AscShortcut(keyCodes.KeyEqual, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ZoomOut].initShortcuts([new AscShortcut(keyCodes.KeyMinus, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertColumnBreak].initShortcuts([new AscShortcut(keyCodes.Enter, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.NonBreakingSpace].initShortcuts([new AscShortcut(keyCodes.Space, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.NonBreakingHyphen].initShortcuts([new AscShortcut(keyCodes.KeyMinus, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditUndo].initShortcuts([new AscShortcut(keyCodes.KeyZ, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditRedo].initShortcuts([new AscShortcut(keyCodes.KeyY, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CopyFormat].initShortcuts([new AscShortcut(keyCodes.KeyC, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PasteFormat].initShortcuts([new AscShortcut(keyCodes.KeyV, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertHyperlink].initShortcuts([new AscShortcut(keyCodes.KeyK, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EditSelectAll].initShortcuts([new AscShortcut(keyCodes.KeyA, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToStartDocument].initShortcuts([new AscShortcut(keyCodes.Home, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToEndDocument].initShortcuts([new AscShortcut(keyCodes.End, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage].initShortcuts([new AscShortcut(keyCodes.PageUp, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage].initShortcuts([new AscShortcut(keyCodes.PageDown, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Bold].initShortcuts([new AscShortcut(keyCodes.KeyB, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Italic].initShortcuts([new AscShortcut(keyCodes.KeyI, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Underline].initShortcuts([new AscShortcut(keyCodes.KeyU, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Subscript].initShortcuts([new AscShortcut(keyCodes.Period, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Superscript].initShortcuts([new AscShortcut(keyCodes.Comma, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyListBullet].initShortcuts([new AscShortcut(keyCodes.KeyL, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ResetChar].initShortcuts([new AscShortcut(keyCodes.Space, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.IncreaseFontSize].initShortcuts([new AscShortcut(keyCodes.BracketRight, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DecreaseFontSize].initShortcuts([new AscShortcut(keyCodes.BracketLeft, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CenterPara].initShortcuts([new AscShortcut(keyCodes.KeyE, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.JustifyPara].initShortcuts([new AscShortcut(keyCodes.KeyJ, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.RightPara].initShortcuts([new AscShortcut(keyCodes.KeyR, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LeftPara].initShortcuts([new AscShortcut(keyCodes.KeyL, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertPageBreak].initShortcuts([new AscShortcut(keyCodes.Enter, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Indent].initShortcuts([new AscShortcut(keyCodes.KeyM, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.UnIndent].initShortcuts([new AscShortcut(keyCodes.KeyM, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertPageNumber].initShortcuts([new AscShortcut(keyCodes.KeyP, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ShowAll].initShortcuts([new AscShortcut(keyCodes.Digit8, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertTableBreak].initShortcuts([new AscShortcut(keyCodes.Enter, false, true, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CopyrightSign].initShortcuts([new AscShortcut(keyCodes.KeyG, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EuroSign].initShortcuts([new AscShortcut(keyCodes.KeyE, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.RegisteredSign].initShortcuts([new AscShortcut(keyCodes.KeyR, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.TrademarkSign].initShortcuts([new AscShortcut(keyCodes.KeyT, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SpeechWorker].initShortcuts([new AscShortcut(keyCodes.KeyZ, false, false, true, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartLine].initShortcuts([new AscShortcut(keyCodes.ArrowLeft, false, false, false, true)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToEndLine].initShortcuts([new AscShortcut(keyCodes.ArrowRight, false, false, false, true)]);
	} else {
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenFilePanel] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenFilePanel, [new AscShortcut(keyCodes.KeyF, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenCommentField] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenCommentField, [new AscShortcut(keyCodes.KeyH, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenChatPanel] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenChatPanel, [new AscShortcut(keyCodes.KeyQ, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.OpenExistingFile] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.OpenExistingFile, [new AscShortcut(keyCodes.KeyO, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.CloseFile] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.CloseFile, [new AscShortcut(keyCodes.KeyW, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage, [new AscShortcut(keyCodes.PageUp, true, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartNextPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartNextPage, [new AscShortcut(keyCodes.PageDown, true, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToNextPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToNextPage, [new AscShortcut(keyCodes.PageDown, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToPreviousPage] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToPreviousPage, [new AscShortcut(keyCodes.PageUp, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToStartWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToStartWord, [new AscShortcut(keyCodes.ArrowLeft, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.MoveToEndWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.MoveToEndWord, [new AscShortcut(keyCodes.ArrowRight, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DeleteLeftWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DeleteLeftWord, [new AscShortcut(keyCodes.Backspace, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.DeleteRightWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.DeleteRightWord, [new AscShortcut(keyCodes.Delete, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Cut] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Cut, [new AscShortcut(keyCodes.KeyX, true, false, false, false), new AscShortcut(keyCodes.Delete, false, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Copy] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Copy, [new AscShortcut(keyCodes.KeyC, true, false, false, false), new AscShortcut(keyCodes.Insert, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Paste] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Paste, [new AscShortcut(keyCodes.KeyV, true, false, false, false), new AscShortcut(keyCodes.Insert, false, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat, [new AscShortcut(keyCodes.KeyV, true, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectRightWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectRightWord, [new AscShortcut(keyCodes.ArrowRight, true, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SelectLeftWord] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SelectLeftWord, [new AscShortcut(keyCodes.ArrowLeft, true, true, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.Strikeout] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.Strikeout, [new AscShortcut(keyCodes.Digit5, true, false, false, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyHeading1] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyHeading1, [new AscShortcut(keyCodes.Digit1, false, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyHeading2] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyHeading2, [new AscShortcut(keyCodes.Digit2, false, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ApplyHeading3] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ApplyHeading3, [new AscShortcut(keyCodes.Digit3, false, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft, [new AscShortcut(keyCodes.ArrowLeft, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight, [new AscShortcut(keyCodes.ArrowRight, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp, [new AscShortcut(keyCodes.ArrowUp, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown, [new AscShortcut(keyCodes.ArrowDown, true, false, false, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertFootnoteNow] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertFootnoteNow, [new AscShortcut(keyCodes.KeyF, true, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.InsertEquation] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.InsertEquation, [new AscShortcut(keyCodes.KeyEqual, false, false, true, false)]);
		// c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.SoftHyphen] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.SoftHyphen, [new AscShortcut(keyCodes.KeyMinus, false, false, true, false)],false, true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EmDash] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EmDash, [new AscShortcut(keyCodes.NumpadMinus, true, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.EnDash] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.EnDash, [new AscShortcut(keyCodes.NumpadMinus, true, false, false, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.HorizontalEllipsis] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.HorizontalEllipsis, [new AscShortcut(keyCodes.Period, true, false, true, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol] = new AscShortcutAction(Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol, [new AscShortcut(keyCodes.KeyX, false, false, true, false)], true);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ZoomIn].initShortcuts([new AscShortcut(keyCodes.NumpadPlus, true, false, false, false)]);
		c_oAscDefaultShortcuts[Asc.c_oAscDocumentShortcutType.ShowAll].initShortcuts([new AscShortcut(keyCodes.Numpad8, true, true, false, false)]);
	}

	function getStringFromShortcutType(type) {
		switch (type) {
			case Asc.c_oAscDocumentShortcutType.OpenFilePanel:
				return "OpenFilePanel";
			case Asc.c_oAscDocumentShortcutType.OpenFindDialog:
				return "OpenFindDialog";
			case Asc.c_oAscDocumentShortcutType.OpenFindAndReplaceMenu:
				return "OpenFindAndReplaceMenu";
			case Asc.c_oAscDocumentShortcutType.OpenCommentsPanel:
				return "OpenCommentsPanel";
			case Asc.c_oAscDocumentShortcutType.OpenCommentField:
				return "OpenCommentField";
			case Asc.c_oAscDocumentShortcutType.OpenChatPanel:
				return "OpenChatPanel";
			case Asc.c_oAscDocumentShortcutType.Save:
				return "Save";
			case Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint:
				return "PrintPreviewAndPrint";
			case Asc.c_oAscDocumentShortcutType.SaveAs:
				return "SaveAs";
			case Asc.c_oAscDocumentShortcutType.OpenHelpMenu:
				return "OpenHelpMenu";
			case Asc.c_oAscDocumentShortcutType.OpenExistingFile:
				return "OpenExistingFile";
			case Asc.c_oAscDocumentShortcutType.NextFileTab:
				return "NextFileTab";
			case Asc.c_oAscDocumentShortcutType.PreviousFileTab:
				return "PreviousFileTab";
			case Asc.c_oAscDocumentShortcutType.CloseFile:
				return "CloseFile";
			case Asc.c_oAscDocumentShortcutType.OpenContextMenu:
				return "OpenContextMenu";
			case Asc.c_oAscDocumentShortcutType.CloseMenu:
				return "CloseMenu";
			case Asc.c_oAscDocumentShortcutType.Zoom100:
				return "Zoom100";
			case Asc.c_oAscDocumentShortcutType.UpdateFields:
				return "UpdateFields";
			case Asc.c_oAscDocumentShortcutType.MoveToStartLine:
				return "MoveToStartLine";
			case Asc.c_oAscDocumentShortcutType.MoveToStartDocument:
				return "MoveToStartDocument";
			case Asc.c_oAscDocumentShortcutType.MoveToEndLine:
				return "MoveToEndLine";
			case Asc.c_oAscDocumentShortcutType.MoveToEndDocument:
				return "MoveToEndDocument";
			case Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage:
				return "MoveToStartPreviousPage";
			case Asc.c_oAscDocumentShortcutType.MoveToStartNextPage:
				return "MoveToStartNextPage";
			case Asc.c_oAscDocumentShortcutType.ScrollDown:
				return "ScrollDown";
			case Asc.c_oAscDocumentShortcutType.ScrollUp:
				return "ScrollUp";
			case Asc.c_oAscDocumentShortcutType.ZoomIn:
				return "ZoomIn";
			case Asc.c_oAscDocumentShortcutType.ZoomOut:
				return "ZoomOut";
			case Asc.c_oAscDocumentShortcutType.MoveToRightChar:
				return "MoveToRightChar";
			case Asc.c_oAscDocumentShortcutType.MoveToLeftChar:
				return "MoveToLeftChar";
			case Asc.c_oAscDocumentShortcutType.MoveToUpLine:
				return "MoveToUpLine";
			case Asc.c_oAscDocumentShortcutType.MoveToDownLine:
				return "MoveToDownLine";
			case Asc.c_oAscDocumentShortcutType.MoveToStartWord:
				return "MoveToStartWord";
			case Asc.c_oAscDocumentShortcutType.MoveToEndWord:
				return "MoveToEndWord";
			case Asc.c_oAscDocumentShortcutType.NextModalControl:
				return "NextModalControl";
			case Asc.c_oAscDocumentShortcutType.PreviousModalControl:
				return "PreviousModalControl";
			case Asc.c_oAscDocumentShortcutType.MoveToLowerHeaderFooter:
				return "MoveToLowerHeaderFooter";
			case Asc.c_oAscDocumentShortcutType.MoveToUpperHeaderFooter:
				return "MoveToUpperHeaderFooter";
			case Asc.c_oAscDocumentShortcutType.MoveToLowerHeader:
				return "MoveToLowerHeader";
			case Asc.c_oAscDocumentShortcutType.MoveToUpperHeader:
				return "MoveToUpperHeader";
			case Asc.c_oAscDocumentShortcutType.EndParagraph:
				return "EndParagraph";
			case Asc.c_oAscDocumentShortcutType.InsertLineBreak:
				return "InsertLineBreak";
			case Asc.c_oAscDocumentShortcutType.InsertColumnBreak:
				return "InsertColumnBreak";
			case Asc.c_oAscDocumentShortcutType.EquationAddPlaceholder:
				return "EquationAddPlaceholder";
			case Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentLeft:
				return "EquationChangeAlignmentLeft";
			case Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentRight:
				return "EquationChangeAlignmentRight";
			case Asc.c_oAscDocumentShortcutType.DeleteLeftChar:
				return "DeleteLeftChar";
			case Asc.c_oAscDocumentShortcutType.DeleteRightChar:
				return "DeleteRightChar";
			case Asc.c_oAscDocumentShortcutType.DeleteLeftWord:
				return "DeleteLeftWord";
			case Asc.c_oAscDocumentShortcutType.DeleteRightWord:
				return "DeleteRightWord";
			case Asc.c_oAscDocumentShortcutType.NonBreakingSpace:
				return "NonBreakingSpace";
			case Asc.c_oAscDocumentShortcutType.NonBreakingHyphen:
				return "NonBreakingHyphen";
			case Asc.c_oAscDocumentShortcutType.EditUndo:
				return "EditUndo";
			case Asc.c_oAscDocumentShortcutType.EditRedo:
				return "EditRedo";
			case Asc.c_oAscDocumentShortcutType.Cut:
				return "Cut";
			case Asc.c_oAscDocumentShortcutType.Copy:
				return "Copy";
			case Asc.c_oAscDocumentShortcutType.Paste:
				return "Paste";
			case Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat:
				return "PasteTextWithoutFormat";
			case Asc.c_oAscDocumentShortcutType.CopyFormat:
				return "CopyFormat";
			case Asc.c_oAscDocumentShortcutType.PasteFormat:
				return "PasteFormat";
			case Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepSourceFormat:
				return "SpecialOptionsKeepSourceFormat";
			case Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepTextOnly:
				return "SpecialOptionsKeepTextOnly";
			case Asc.c_oAscDocumentShortcutType.SpecialOptionsOverwriteCells:
				return "SpecialOptionsOverwriteCells";
			case Asc.c_oAscDocumentShortcutType.SpecialOptionsNestTable:
				return "SpecialOptionsNestTable";
			case Asc.c_oAscDocumentShortcutType.InsertHyperlink:
				return "InsertHyperlink";
			case Asc.c_oAscDocumentShortcutType.VisitHyperlink:
				return "VisitHyperlink";
			case Asc.c_oAscDocumentShortcutType.EditSelectAll:
				return "EditSelectAll";
			case Asc.c_oAscDocumentShortcutType.SelectToStartLine:
				return "SelectToStartLine";
			case Asc.c_oAscDocumentShortcutType.SelectToEndLine:
				return "SelectToEndLine";
			case Asc.c_oAscDocumentShortcutType.SelectToStartDocument:
				return "SelectToStartDocument";
			case Asc.c_oAscDocumentShortcutType.SelectToEndDocument:
				return "SelectToEndDocument";
			case Asc.c_oAscDocumentShortcutType.SelectRightChar:
				return "SelectRightChar";
			case Asc.c_oAscDocumentShortcutType.SelectLeftChar:
				return "SelectLeftChar";
			case Asc.c_oAscDocumentShortcutType.SelectRightWord:
				return "SelectRightWord";
			case Asc.c_oAscDocumentShortcutType.SelectLeftWord:
				return "SelectLeftWord";
			case Asc.c_oAscDocumentShortcutType.SelectLineUp:
				return "SelectLineUp";
			case Asc.c_oAscDocumentShortcutType.SelectLineDown:
				return "SelectLineDown";
			case Asc.c_oAscDocumentShortcutType.SelectPageUp:
				return "SelectPageUp";
			case Asc.c_oAscDocumentShortcutType.SelectPageDown:
				return "SelectPageDown";
			case Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage:
				return "SelectToBeginPreviousPage";
			case Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage:
				return "SelectToBeginNextPage";
			case Asc.c_oAscDocumentShortcutType.Bold:
				return "Bold";
			case Asc.c_oAscDocumentShortcutType.Italic:
				return "Italic";
			case Asc.c_oAscDocumentShortcutType.Underline:
				return "Underline";
			case Asc.c_oAscDocumentShortcutType.Strikeout:
				return "Strikeout";
			case Asc.c_oAscDocumentShortcutType.Subscript:
				return "Subscript";
			case Asc.c_oAscDocumentShortcutType.Superscript:
				return "Superscript";
			case Asc.c_oAscDocumentShortcutType.ApplyHeading1:
				return "ApplyHeading1";
			case Asc.c_oAscDocumentShortcutType.ApplyHeading2:
				return "ApplyHeading2";
			case Asc.c_oAscDocumentShortcutType.ApplyHeading3:
				return "ApplyHeading3";
			case Asc.c_oAscDocumentShortcutType.ApplyListBullet:
				return "ApplyListBullet";
			case Asc.c_oAscDocumentShortcutType.ResetChar:
				return "ResetChar";
			case Asc.c_oAscDocumentShortcutType.IncreaseFontSize:
				return "IncreaseFontSize";
			case Asc.c_oAscDocumentShortcutType.DecreaseFontSize:
				return "DecreaseFontSize";
			case Asc.c_oAscDocumentShortcutType.CenterPara:
				return "CenterPara";
			case Asc.c_oAscDocumentShortcutType.JustifyPara:
				return "JustifyPara";
			case Asc.c_oAscDocumentShortcutType.RightPara:
				return "RightPara";
			case Asc.c_oAscDocumentShortcutType.LeftPara:
				return "LeftPara";
			case Asc.c_oAscDocumentShortcutType.InsertPageBreak:
				return "InsertPageBreak";
			case Asc.c_oAscDocumentShortcutType.Indent:
				return "Indent";
			case Asc.c_oAscDocumentShortcutType.UnIndent:
				return "UnIndent";
			case Asc.c_oAscDocumentShortcutType.InsertPageNumber:
				return "InsertPageNumber";
			case Asc.c_oAscDocumentShortcutType.ShowAll:
				return "ShowAll";
			case Asc.c_oAscDocumentShortcutType.StartIndent:
				return "StartIndent";
			case Asc.c_oAscDocumentShortcutType.StartUnIndent:
				return "StartUnIndent";
			case Asc.c_oAscDocumentShortcutType.InsertTab:
				return "InsertTab";
			case Asc.c_oAscDocumentShortcutType.MixedIndent:
				return "MixedIndent";
			case Asc.c_oAscDocumentShortcutType.MixedUnIndent:
				return "MixedUnIndent";
			case Asc.c_oAscDocumentShortcutType.EditShape:
				return "EditShape";
			case Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft:
				return "LittleMoveObjectLeft";
			case Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight:
				return "LittleMoveObjectRight";
			case Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp:
				return "LittleMoveObjectUp";
			case Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown:
				return "LittleMoveObjectDown";
			case Asc.c_oAscDocumentShortcutType.BigMoveObjectLeft:
				return "BigMoveObjectLeft";
			case Asc.c_oAscDocumentShortcutType.BigMoveObjectRight:
				return "BigMoveObjectRight";
			case Asc.c_oAscDocumentShortcutType.BigMoveObjectUp:
				return "BigMoveObjectUp";
			case Asc.c_oAscDocumentShortcutType.BigMoveObjectDown:
				return "BigMoveObjectDown";
			case Asc.c_oAscDocumentShortcutType.MoveFocusToNextObject:
				return "MoveFocusToNextObject";
			case Asc.c_oAscDocumentShortcutType.MoveFocusToPreviousObject:
				return "MoveFocusToPreviousObject";
			case Asc.c_oAscDocumentShortcutType.InsertEndnoteNow:
				return "InsertEndnoteNow";
			case Asc.c_oAscDocumentShortcutType.InsertFootnoteNow:
				return "InsertFootnoteNow";
			case Asc.c_oAscDocumentShortcutType.MoveToNextCell:
				return "MoveToNextCell";
			case Asc.c_oAscDocumentShortcutType.MoveToPreviousCell:
				return "MoveToPreviousCell";
			case Asc.c_oAscDocumentShortcutType.MoveToNextRow:
				return "MoveToNextRow";
			case Asc.c_oAscDocumentShortcutType.MoveToPreviousRow:
				return "MoveToPreviousRow";
			case Asc.c_oAscDocumentShortcutType.EndParagraphCell:
				return "EndParagraphCell";
			case Asc.c_oAscDocumentShortcutType.AddNewRow:
				return "AddNewRow";
			case Asc.c_oAscDocumentShortcutType.InsertTableBreak:
				return "InsertTableBreak";
			case Asc.c_oAscDocumentShortcutType.MoveToNextForm:
				return "MoveToNextForm";
			case Asc.c_oAscDocumentShortcutType.MoveToPreviousForm:
				return "MoveToPreviousForm";
			case Asc.c_oAscDocumentShortcutType.ChooseNextComboBoxOption:
				return "ChooseNextComboBoxOption";
			case Asc.c_oAscDocumentShortcutType.ChoosePreviousComboBoxOption:
				return "ChoosePreviousComboBoxOption";
			case Asc.c_oAscDocumentShortcutType.InsertEquation:
				return "InsertEquation";
			case Asc.c_oAscDocumentShortcutType.EmDash:
				return "EmDash";
			case Asc.c_oAscDocumentShortcutType.EnDash:
				return "EnDash";
			case Asc.c_oAscDocumentShortcutType.CopyrightSign:
				return "CopyrightSign";
			case Asc.c_oAscDocumentShortcutType.EuroSign:
				return "EuroSign";
			case Asc.c_oAscDocumentShortcutType.RegisteredSign:
				return "RegisteredSign";
			case Asc.c_oAscDocumentShortcutType.TrademarkSign:
				return "TrademarkSign";
			case Asc.c_oAscDocumentShortcutType.HorizontalEllipsis:
				return "HorizontalEllipsis";
			case Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol:
				return "ReplaceUnicodeToSymbol";
			case Asc.c_oAscDocumentShortcutType.SoftHyphen:
				return "SoftHyphen";
			case Asc.c_oAscDocumentShortcutType.SpeechWorker:
				return "SpeechWorker";
			case Asc.c_oAscDocumentShortcutType.EditChart:
				return "EditChart";
			case Asc.c_oAscDocumentShortcutType.InsertLineBreakMultilineForm:
				return "InsertLineBreakMultilineForm";
			case Asc.c_oAscDocumentShortcutType.MoveToNextPage:
				return "MoveToNextPage";
			case Asc.c_oAscDocumentShortcutType.MoveToPreviousPage:
				return "MoveToPreviousPage";
		}
		return null;
	}

	function getShortcutTypeFromString(str) {
		switch (str) {
			case "OpenFilePanel":
				return Asc.c_oAscDocumentShortcutType.OpenFilePanel;
			case "OpenFindDialog":
				return Asc.c_oAscDocumentShortcutType.OpenFindDialog;
			case "OpenFindAndReplaceMenu":
				return Asc.c_oAscDocumentShortcutType.OpenFindAndReplaceMenu;
			case "OpenCommentsPanel":
				return Asc.c_oAscDocumentShortcutType.OpenCommentsPanel;
			case "OpenCommentField":
				return Asc.c_oAscDocumentShortcutType.OpenCommentField;
			case "OpenChatPanel":
				return Asc.c_oAscDocumentShortcutType.OpenChatPanel;
			case "Save":
				return Asc.c_oAscDocumentShortcutType.Save;
			case "PrintPreviewAndPrint":
				return Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint;
			case "SaveAs":
				return Asc.c_oAscDocumentShortcutType.SaveAs;
			case "OpenHelpMenu":
				return Asc.c_oAscDocumentShortcutType.OpenHelpMenu;
			case "OpenExistingFile":
				return Asc.c_oAscDocumentShortcutType.OpenExistingFile;
			case "NextFileTab":
				return Asc.c_oAscDocumentShortcutType.NextFileTab;
			case "PreviousFileTab":
				return Asc.c_oAscDocumentShortcutType.PreviousFileTab;
			case "CloseFile":
				return Asc.c_oAscDocumentShortcutType.CloseFile;
			case "OpenContextMenu":
				return Asc.c_oAscDocumentShortcutType.OpenContextMenu;
			case "CloseMenu":
				return Asc.c_oAscDocumentShortcutType.CloseMenu;
			case "Zoom100":
				return Asc.c_oAscDocumentShortcutType.Zoom100;
			case "UpdateFields":
				return Asc.c_oAscDocumentShortcutType.UpdateFields;
			case "MoveToStartLine":
				return Asc.c_oAscDocumentShortcutType.MoveToStartLine;
			case "MoveToStartDocument":
				return Asc.c_oAscDocumentShortcutType.MoveToStartDocument;
			case "MoveToEndLine":
				return Asc.c_oAscDocumentShortcutType.MoveToEndLine;
			case "MoveToEndDocument":
				return Asc.c_oAscDocumentShortcutType.MoveToEndDocument;
			case "MoveToStartPreviousPage":
				return Asc.c_oAscDocumentShortcutType.MoveToStartPreviousPage;
			case "MoveToStartNextPage":
				return Asc.c_oAscDocumentShortcutType.MoveToStartNextPage;
			case "ScrollDown":
				return Asc.c_oAscDocumentShortcutType.ScrollDown;
			case "ScrollUp":
				return Asc.c_oAscDocumentShortcutType.ScrollUp;
			case "ZoomIn":
				return Asc.c_oAscDocumentShortcutType.ZoomIn;
			case "ZoomOut":
				return Asc.c_oAscDocumentShortcutType.ZoomOut;
			case "MoveToRightChar":
				return Asc.c_oAscDocumentShortcutType.MoveToRightChar;
			case "MoveToLeftChar":
				return Asc.c_oAscDocumentShortcutType.MoveToLeftChar;
			case "MoveToUpLine":
				return Asc.c_oAscDocumentShortcutType.MoveToUpLine;
			case "MoveToDownLine":
				return Asc.c_oAscDocumentShortcutType.MoveToDownLine;
			case "MoveToStartWord":
				return Asc.c_oAscDocumentShortcutType.MoveToStartWord;
			case "MoveToEndWord":
				return Asc.c_oAscDocumentShortcutType.MoveToEndWord;
			case "NextModalControl":
				return Asc.c_oAscDocumentShortcutType.NextModalControl;
			case "PreviousModalControl":
				return Asc.c_oAscDocumentShortcutType.PreviousModalControl;
			case "MoveToLowerHeaderFooter":
				return Asc.c_oAscDocumentShortcutType.MoveToLowerHeaderFooter;
			case "MoveToUpperHeaderFooter":
				return Asc.c_oAscDocumentShortcutType.MoveToUpperHeaderFooter;
			case "MoveToLowerHeader":
				return Asc.c_oAscDocumentShortcutType.MoveToLowerHeader;
			case "MoveToUpperHeader":
				return Asc.c_oAscDocumentShortcutType.MoveToUpperHeader;
			case "EndParagraph":
				return Asc.c_oAscDocumentShortcutType.EndParagraph;
			case "InsertLineBreak":
				return Asc.c_oAscDocumentShortcutType.InsertLineBreak;
			case "InsertColumnBreak":
				return Asc.c_oAscDocumentShortcutType.InsertColumnBreak;
			case "EquationAddPlaceholder":
				return Asc.c_oAscDocumentShortcutType.EquationAddPlaceholder;
			case "EquationChangeAlignmentLeft":
				return Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentLeft;
			case "EquationChangeAlignmentRight":
				return Asc.c_oAscDocumentShortcutType.EquationChangeAlignmentRight;
			case "DeleteLeftChar":
				return Asc.c_oAscDocumentShortcutType.DeleteLeftChar;
			case "DeleteRightChar":
				return Asc.c_oAscDocumentShortcutType.DeleteRightChar;
			case "DeleteLeftWord":
				return Asc.c_oAscDocumentShortcutType.DeleteLeftWord;
			case "DeleteRightWord":
				return Asc.c_oAscDocumentShortcutType.DeleteRightWord;
			case "NonBreakingSpace":
				return Asc.c_oAscDocumentShortcutType.NonBreakingSpace;
			case "NonBreakingHyphen":
				return Asc.c_oAscDocumentShortcutType.NonBreakingHyphen;
			case "EditUndo":
				return Asc.c_oAscDocumentShortcutType.EditUndo;
			case "EditRedo":
				return Asc.c_oAscDocumentShortcutType.EditRedo;
			case "Cut":
				return Asc.c_oAscDocumentShortcutType.Cut;
			case "Copy":
				return Asc.c_oAscDocumentShortcutType.Copy;
			case "Paste":
				return Asc.c_oAscDocumentShortcutType.Paste;
			case "PasteTextWithoutFormat":
				return Asc.c_oAscDocumentShortcutType.PasteTextWithoutFormat;
			case "CopyFormat":
				return Asc.c_oAscDocumentShortcutType.CopyFormat;
			case "PasteFormat":
				return Asc.c_oAscDocumentShortcutType.PasteFormat;
			case "SpecialOptionsKeepSourceFormat":
				return Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepSourceFormat;
			case "SpecialOptionsKeepTextOnly":
				return Asc.c_oAscDocumentShortcutType.SpecialOptionsKeepTextOnly;
			case "SpecialOptionsOverwriteCells":
				return Asc.c_oAscDocumentShortcutType.SpecialOptionsOverwriteCells;
			case "SpecialOptionsNestTable":
				return Asc.c_oAscDocumentShortcutType.SpecialOptionsNestTable;
			case "InsertHyperlink":
				return Asc.c_oAscDocumentShortcutType.InsertHyperlink;
			case "VisitHyperlink":
				return Asc.c_oAscDocumentShortcutType.VisitHyperlink;
			case "EditSelectAll":
				return Asc.c_oAscDocumentShortcutType.EditSelectAll;
			case "SelectToStartLine":
				return Asc.c_oAscDocumentShortcutType.SelectToStartLine;
			case "SelectToEndLine":
				return Asc.c_oAscDocumentShortcutType.SelectToEndLine;
			case "SelectToStartDocument":
				return Asc.c_oAscDocumentShortcutType.SelectToStartDocument;
			case "SelectToEndDocument":
				return Asc.c_oAscDocumentShortcutType.SelectToEndDocument;
			case "SelectRightChar":
				return Asc.c_oAscDocumentShortcutType.SelectRightChar;
			case "SelectLeftChar":
				return Asc.c_oAscDocumentShortcutType.SelectLeftChar;
			case "SelectRightWord":
				return Asc.c_oAscDocumentShortcutType.SelectRightWord;
			case "SelectLeftWord":
				return Asc.c_oAscDocumentShortcutType.SelectLeftWord;
			case "SelectLineUp":
				return Asc.c_oAscDocumentShortcutType.SelectLineUp;
			case "SelectLineDown":
				return Asc.c_oAscDocumentShortcutType.SelectLineDown;
			case "SelectPageUp":
				return Asc.c_oAscDocumentShortcutType.SelectPageUp;
			case "SelectPageDown":
				return Asc.c_oAscDocumentShortcutType.SelectPageDown;
			case "SelectToBeginPreviousPage":
				return Asc.c_oAscDocumentShortcutType.SelectToBeginPreviousPage;
			case "SelectToBeginNextPage":
				return Asc.c_oAscDocumentShortcutType.SelectToBeginNextPage;
			case "Bold":
				return Asc.c_oAscDocumentShortcutType.Bold;
			case "Italic":
				return Asc.c_oAscDocumentShortcutType.Italic;
			case "Underline":
				return Asc.c_oAscDocumentShortcutType.Underline;
			case "Strikeout":
				return Asc.c_oAscDocumentShortcutType.Strikeout;
			case "Subscript":
				return Asc.c_oAscDocumentShortcutType.Subscript;
			case "Superscript":
				return Asc.c_oAscDocumentShortcutType.Superscript;
			case "ApplyHeading1":
				return Asc.c_oAscDocumentShortcutType.ApplyHeading1;
			case "ApplyHeading2":
				return Asc.c_oAscDocumentShortcutType.ApplyHeading2;
			case "ApplyHeading3":
				return Asc.c_oAscDocumentShortcutType.ApplyHeading3;
			case "ApplyListBullet":
				return Asc.c_oAscDocumentShortcutType.ApplyListBullet;
			case "ResetChar":
				return Asc.c_oAscDocumentShortcutType.ResetChar;
			case "IncreaseFontSize":
				return Asc.c_oAscDocumentShortcutType.IncreaseFontSize;
			case "DecreaseFontSize":
				return Asc.c_oAscDocumentShortcutType.DecreaseFontSize;
			case "CenterPara":
				return Asc.c_oAscDocumentShortcutType.CenterPara;
			case "JustifyPara":
				return Asc.c_oAscDocumentShortcutType.JustifyPara;
			case "RightPara":
				return Asc.c_oAscDocumentShortcutType.RightPara;
			case "LeftPara":
				return Asc.c_oAscDocumentShortcutType.LeftPara;
			case "InsertPageBreak":
				return Asc.c_oAscDocumentShortcutType.InsertPageBreak;
			case "Indent":
				return Asc.c_oAscDocumentShortcutType.Indent;
			case "UnIndent":
				return Asc.c_oAscDocumentShortcutType.UnIndent;
			case "InsertPageNumber":
				return Asc.c_oAscDocumentShortcutType.InsertPageNumber;
			case "ShowAll":
				return Asc.c_oAscDocumentShortcutType.ShowAll;
			case "StartIndent":
				return Asc.c_oAscDocumentShortcutType.StartIndent;
			case "StartUnIndent":
				return Asc.c_oAscDocumentShortcutType.StartUnIndent;
			case "InsertTab":
				return Asc.c_oAscDocumentShortcutType.InsertTab;
			case "MixedIndent":
				return Asc.c_oAscDocumentShortcutType.MixedIndent;
			case "MixedUnIndent":
				return Asc.c_oAscDocumentShortcutType.MixedUnIndent;
			case "EditShape":
				return Asc.c_oAscDocumentShortcutType.EditShape;
			case "LittleMoveObjectLeft":
				return Asc.c_oAscDocumentShortcutType.LittleMoveObjectLeft;
			case "LittleMoveObjectRight":
				return Asc.c_oAscDocumentShortcutType.LittleMoveObjectRight;
			case "LittleMoveObjectUp":
				return Asc.c_oAscDocumentShortcutType.LittleMoveObjectUp;
			case "LittleMoveObjectDown":
				return Asc.c_oAscDocumentShortcutType.LittleMoveObjectDown;
			case "BigMoveObjectLeft":
				return Asc.c_oAscDocumentShortcutType.BigMoveObjectLeft;
			case "BigMoveObjectRight":
				return Asc.c_oAscDocumentShortcutType.BigMoveObjectRight;
			case "BigMoveObjectUp":
				return Asc.c_oAscDocumentShortcutType.BigMoveObjectUp;
			case "BigMoveObjectDown":
				return Asc.c_oAscDocumentShortcutType.BigMoveObjectDown;
			case "MoveFocusToNextObject":
				return Asc.c_oAscDocumentShortcutType.MoveFocusToNextObject;
			case "MoveFocusToPreviousObject":
				return Asc.c_oAscDocumentShortcutType.MoveFocusToPreviousObject;
			case "InsertEndnoteNow":
				return Asc.c_oAscDocumentShortcutType.InsertEndnoteNow;
			case "InsertFootnoteNow":
				return Asc.c_oAscDocumentShortcutType.InsertFootnoteNow;
			case "MoveToNextCell":
				return Asc.c_oAscDocumentShortcutType.MoveToNextCell;
			case "MoveToPreviousCell":
				return Asc.c_oAscDocumentShortcutType.MoveToPreviousCell;
			case "MoveToNextRow":
				return Asc.c_oAscDocumentShortcutType.MoveToNextRow;
			case "MoveToPreviousRow":
				return Asc.c_oAscDocumentShortcutType.MoveToPreviousRow;
			case "EndParagraphCell":
				return Asc.c_oAscDocumentShortcutType.EndParagraphCell;
			case "AddNewRow":
				return Asc.c_oAscDocumentShortcutType.AddNewRow;
			case "InsertTableBreak":
				return Asc.c_oAscDocumentShortcutType.InsertTableBreak;
			case "MoveToNextForm":
				return Asc.c_oAscDocumentShortcutType.MoveToNextForm;
			case "MoveToPreviousForm":
				return Asc.c_oAscDocumentShortcutType.MoveToPreviousForm;
			case "ChooseNextComboBoxOption":
				return Asc.c_oAscDocumentShortcutType.ChooseNextComboBoxOption;
			case "ChoosePreviousComboBoxOption":
				return Asc.c_oAscDocumentShortcutType.ChoosePreviousComboBoxOption;
			case "InsertEquation":
				return Asc.c_oAscDocumentShortcutType.InsertEquation;
			case "EmDash":
				return Asc.c_oAscDocumentShortcutType.EmDash;
			case "EnDash":
				return Asc.c_oAscDocumentShortcutType.EnDash;
			case "CopyrightSign":
				return Asc.c_oAscDocumentShortcutType.CopyrightSign;
			case "EuroSign":
				return Asc.c_oAscDocumentShortcutType.EuroSign;
			case "RegisteredSign":
				return Asc.c_oAscDocumentShortcutType.RegisteredSign;
			case "TrademarkSign":
				return Asc.c_oAscDocumentShortcutType.TrademarkSign;
			case "HorizontalEllipsis":
				return Asc.c_oAscDocumentShortcutType.HorizontalEllipsis;
			case "ReplaceUnicodeToSymbol":
				return Asc.c_oAscDocumentShortcutType.ReplaceUnicodeToSymbol;
			case "SoftHyphen":
				return Asc.c_oAscDocumentShortcutType.SoftHyphen;
			case "SpeechWorker":
				return Asc.c_oAscDocumentShortcutType.SpeechWorker;
			case "EditChart":
				return Asc.c_oAscDocumentShortcutType.EditChart;
			case "InsertLineBreakMultilineForm":
				return Asc.c_oAscDocumentShortcutType.InsertLineBreakMultilineForm;
			case "MoveToNextPage":
				return Asc.c_oAscDocumentShortcutType.MoveToNextPage;
			case "MoveToPreviousPage":
				return Asc.c_oAscDocumentShortcutType.MoveToPreviousPage;
		}
	}

	window["Asc"]["c_oAscDefaultShortcuts"] = window["Asc"].c_oAscDefaultShortcuts = c_oAscDefaultShortcuts;
	window["AscCommon"].getStringFromShortcutType = getStringFromShortcutType;
	window["AscCommon"].getShortcutTypeFromString = getShortcutTypeFromString;
})();
