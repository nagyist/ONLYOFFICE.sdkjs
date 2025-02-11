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
	const AscShortcut = Asc.CAscShortcut;
	const keyCodes = Asc.c_oAscKeyCodes;

	const c_oAscSpreadsheetShortcutTypes = {
		OpenFilePanel:              1,
		OpenFindDialog:             2,
		OpenFindReplaceMenu:        3,
		OpenCommentsPanel:          4,
		OpenCommentField:           5,
		OpenChatPanel:              6,
		Save:                       7,
		Print:                      8,
		DownloadAs:                 9,
		HelpMenu:                   11,
		OpenExistingFile:           12,
		NextFileTab:                13,
		PreviousFileTab:            14,
		CloseFile:                  15,
		ElementContextualMenu:      16,
		CloseMenuModal:             17,
		ResetZoom:                  18,
		CellMoveUp:                 19,
		CellMoveDown:               20,
		CellMoveLeft:               21,
		CellMoveRight:              22,
		CellMoveActiveCellDown:     23,
		CellMoveActiveCellUp:       24,
		CellMoveActiveCellRight:    25,
		CellMoveActiveCellLeft:     26,
		CellMoveLeftNonBlank:       27,
		CellMoveFirstColumn:        28,
		CellMoveRightNonBlank:      29,
		CellMoveBottomNonBlank:     30,
		CellMoveBottomEdge:         31,
		CellMoveTopNonBlank:        32,
		CellMoveTopEdge:            33,
		CellMoveFirstCell:          35,
		CellMoveEndSpreadsheet:     37,
		PreviousWorksheet:          38,
		NextWorksheet:              39,
		ZoomIn:                     46,
		ZoomOut:                    47,
		NavigatePreviousControl:    48,
		NavigateNextControl:        49,
		SelectColumn:               51,
		SelectRow:                  52,
		SelectOneCellRight:         53,
		SelectOneCellLeft:          54,
		SelectOneCellUp:            55,
		SelectOneCellDown:          56,
		SelectCursorBeginningRow:   57,
		SelectCursorEndRow:         58,
		SelectNextNonblankRight:    59,
		SelectNextNonblankLeft:     60,
		SelectNextNonblankUp:       61,
		SelectNextNonblankDown:     62,
		SelectBeginningWorksheet:   63,
		SelectLastUsedCell:         64,
		SelectNearestNonblankRight: 65,
		SelectNonblankLeft:         66,
		SelectFirstColumn:          67,
		SelectNearestNonblankDown:  68,
		SelectNearestNonblankUp:    69,
		SelectDownOneScreen:        70,
		SelectUpOneScreen:          71,
		EditUndo:                   72,
		EditRedo:                   73,
		Cut:                        74,
		Copy:                       75,
		Paste:                      76,
		PasteOnlyFormula:           77,
		PasteFormulaNumberFormat:   78,
		PasteFormulaAllFormatting:  79,
		PasteFormulaNoBorders:      80,
		PasteFormulaColumnWidth:    81,
		Transpose:                  82,
		PasteOnlyValue:             83,
		PasteValueNumberFormat:     84,
		PasteValueAllFormatting:    85,
		PasteOnlyFormatting:        86,
		PasteLink:                  87,
		InsertHyperlink:            88,
		VisitHyperlink:             89,
		Bold:                       90,
		Italic:                     91,
		Underline:                  92,
		Strikeout:                  93,
		EditOpenCellEditor:         94,
		ToggleAutoFilter:           95,
		OpenFilterWindow:           96,
		FormatAsTableTemplate:      97,
		CompleteCellEntryMoveDown:  98,
		CompleteCellEntryMoveUp:    99,
		CompleteCellEntryMoveRight: 100,
		CompleteCellEntryMoveLeft:  101,
		CompleteCellEntryStay:      102,
		FillSelectedCellRange:      103,
		CellStartNewLine:           104,
		AddPlaceholderEquation:     105,
		CellEntryCancel:            106,
		RemoveCharLeft:             107,
		RemoveCharRight:            108,
		ClearActiveCellContent:     109,
		ClearSelectedCellsContent:  110,
		OpenInsertCellsWindow:      111,
		OpenDeleteCellsWindow:      112,
		CellInsertDate:             113,
		CellInsertTime:             114,
		CellAddSeparator:           115,
		AutoFill:                   116,
		RemoveWordLeft:             117,
		RemoveWordRight:            118,
		EditSelectAll:              119,
		MoveCharacterLeft:          120,
		MoveCharacterRight:         121,
		MoveCursorLineUp:           122,
		MoveCursorLineDown:         123,
		SelectCharacterRight:       124,
		SelectCharacterLeft:        125,
		MoveWordLeft:               126,
		MoveWordRight:              127,
		SelectWordLeft:             128,
		SelectWordRight:            129,
		MoveBeginningText:          130,
		MoveEndText:                131,
		SelectBeginningText:        132,
		SelectEndText:              133,
		MoveBeginningLine:          134,
		MoveEndLine:                135,
		SelectBeginningLine:        136,
		SelectEndLine:              137,
		SelectLineUp:               138,
		SelectLineDown:             139,
		RefreshSelectedPivots:      140,
		RefreshAllPivots:           141,
		SlicerClearSelectedValues:  142,
		SlicerSwitchMultiSelect:    143,
		FormatTableAddSummaryRow:   144,
		OpenInsertFunctionDialog:   145,
		CellInsertSumFunction:      146,
		RecalculateAll:             147,
		RecalculateActiveSheet:     148,
		DisplayFunctionsSheet:      149,
		CellEditorSwitchReference:  150,
		OpenNumberFormatDialog:     151,
		CellGeneralFormat:          152,
		CellCurrencyFormat:         153,
		CellPercentFormat:          154,
		CellExponentialFormat:      155,
		CellDateFormat:             156,
		CellTimeFormat:             157,
		CellNumberFormat:           158,
		EditShape:                  159,
		EditChart:                  160,
		MoveShapeLittleStepRight:   161,
		MoveShapeLittleStepLeft:    162,
		MoveShapeLittleStepUp:      163,
		MoveShapeLittleStepBottom:  164,
		MoveShapeBigStepLeft:       165,
		MoveShapeBigStepRight:      166,
		MoveShapeBigStepUp:         167,
		MoveShapeBigStepBottom:     168,
		MoveFocusNextObject:        169,
		MoveFocusPreviousObject:    170,
		DrawingAddTab:              172,
		DrawingSubscript:           173,
		DrawingSuperscript:         174,
		IncreaseFontSize:           175,
		DecreaseFontSize:           176,
		DrawingCenterPara:          177,
		DrawingJustifyPara:         178,
		DrawingRightPara:           179,
		DrawingLeftPara:            180,
		EndParagraph:               181,
		AddLineBreak:               182,
		RemoveGraphicalObject:      183,
		ExitAddingShapesMode:       184,
		SpeechWorker:               185,
		DrawingEnDash:              187
	};
	Asc.c_oAscSpreadsheetShortcutTypes = c_oAscSpreadsheetShortcutTypes;

	const c_oAscUnlockedShortcutActionTypes = {};
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.RefreshAllPivots] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.RefreshSelectedPivots] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.RecalculateAll] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.RecalculateActiveSheet] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.NextWorksheet] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.PreviousWorksheet] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.Strikeout] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.Italic] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.Bold] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.Underline] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.EditUndo] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.EditRedo] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.Print] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.EditOpenCellEditor] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellAddSeparator] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellNumberFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellTimeFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellDateFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellCurrencyFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellPercentFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellExponentialFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellGeneralFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.ShowFormulas] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.CellEditorSwitchReference] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash] = true;

	const c_oAscDefaultShortcuts = {};

	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenFindDialog] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenFindDialog, keyCodes.KeyF, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenFindReplaceMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenFindReplaceMenu, keyCodes.KeyH, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentsPanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentsPanel, keyCodes.KeyH, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Save] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Save, keyCodes.KeyS, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Print] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Print, keyCodes.KeyP, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DownloadAs] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DownloadAs, keyCodes.KeyS, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.HelpMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.HelpMenu, keyCodes.F1, false, false, false, false)];

	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ElementContextualMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ElementContextualMenu, keyCodes.F10, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CloseMenuModal] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CloseMenuModal, keyCodes.Esc, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ResetZoom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ResetZoom, keyCodes.Digit0, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveDown, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellDown, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellUp, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellRight, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellLeft, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstColumn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstColumn, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomEdge] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomEdge, keyCodes.PageDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopEdge] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopEdge, keyCodes.PageUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstCell] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstCell, keyCodes.Home, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveEndSpreadsheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveEndSpreadsheet, keyCodes.End, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PreviousWorksheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PreviousWorksheet, keyCodes.PageUp, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.NextWorksheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.NextWorksheet, keyCodes.PageDown, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn, keyCodes.KeyEqual, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn, keyCodes.NumpadPlus, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut, keyCodes.KeyMinus, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut, keyCodes.NumpadMinus, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.NavigatePreviousControl] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.NavigatePreviousControl, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.NavigateNextControl] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.NavigateNextControl, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectColumn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectColumn, keyCodes.Space, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectRow, keyCodes.Space, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellRight, keyCodes.ArrowRight, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellLeft, keyCodes.ArrowLeft, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellUp, keyCodes.ArrowUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellDown, keyCodes.ArrowDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorBeginningRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorBeginningRow, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorEndRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorEndRow, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankRight, keyCodes.ArrowRight, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankLeft, keyCodes.ArrowLeft, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankUp, keyCodes.ArrowUp, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankDown, keyCodes.ArrowDown, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningWorksheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningWorksheet, keyCodes.Home, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectLastUsedCell] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectLastUsedCell, keyCodes.End, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankRight, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectFirstColumn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectFirstColumn, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectDownOneScreen] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectDownOneScreen, keyCodes.PageDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectUpOneScreen] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectUpOneScreen, keyCodes.PageUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditUndo] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditUndo, keyCodes.KeyZ, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditRedo] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditRedo, keyCodes.KeyY, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormula] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormula, keyCodes.KeyF, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNumberFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNumberFormat, keyCodes.KeyO, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaAllFormatting] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaAllFormatting, keyCodes.KeyK, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNoBorders] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNoBorders, keyCodes.KeyB, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaColumnWidth] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaColumnWidth, keyCodes.KeyW, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Transpose] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Transpose, keyCodes.KeyT, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyValue] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyValue, keyCodes.KeyV, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteValueNumberFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteValueNumberFormat, keyCodes.KeyA, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteValueAllFormatting] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteValueAllFormatting, keyCodes.KeyE, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormatting] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormatting, keyCodes.KeyR, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.PasteLink] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.PasteLink, keyCodes.KeyN, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.InsertHyperlink] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.InsertHyperlink, keyCodes.KeyK, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.VisitHyperlink] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.VisitHyperlink, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Bold] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Bold, keyCodes.KeyB, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Underline] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Underline, keyCodes.KeyU, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Strikeout] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Strikeout, keyCodes.Digit5, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditOpenCellEditor] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditOpenCellEditor, keyCodes.F2, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ToggleAutoFilter] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ToggleAutoFilter, keyCodes.KeyL, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenFilterWindow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenFilterWindow, keyCodes.ArrowDown, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.FormatAsTableTemplate] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.FormatAsTableTemplate, keyCodes.KeyL, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveDown, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveUp, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveRight, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveLeft, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryStay] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryStay, keyCodes.Enter, true, true, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryStay, keyCodes.Enter, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.FillSelectedCellRange] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.FillSelectedCellRange, keyCodes.Enter, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellStartNewLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellStartNewLine, keyCodes.Enter, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.AddPlaceholderEquation] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.AddPlaceholderEquation, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellEntryCancel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellEntryCancel, keyCodes.Esc, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharLeft, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharRight, keyCodes.Delete, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ClearActiveCellContent] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ClearActiveCellContent, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ClearSelectedCellsContent] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ClearSelectedCellsContent, keyCodes.Delete, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow, keyCodes.KeyEqual, true, true, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow, keyCodes.NumpadPlus, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow, keyCodes.KeyMinus, true, true, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow, keyCodes.NumpadMinus, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate, keyCodes.Semicolon, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime, keyCodes.Semicolon, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.AutoFill] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.AutoFill, keyCodes.ArrowDown, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll, keyCodes.KeyA, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineDown, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterRight, keyCodes.ArrowRight, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterLeft, keyCodes.ArrowLeft, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningText, keyCodes.Home, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveEndText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveEndText, keyCodes.End, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningText, keyCodes.Home, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectEndText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectEndText, keyCodes.End, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningLine, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveEndLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveEndLine, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningLine, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectEndLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectEndLine, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectLineUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectLineUp, keyCodes.ArrowUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectLineDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectLineDown, keyCodes.ArrowDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RefreshSelectedPivots] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RefreshSelectedPivots, keyCodes.F5, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RefreshAllPivots] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RefreshAllPivots, keyCodes.F5, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.FormatTableAddSummaryRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.FormatTableAddSummaryRow, keyCodes.KeyR, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertFunctionDialog] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertFunctionDialog, keyCodes.F3, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RecalculateAll] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RecalculateAll, keyCodes.F9, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RecalculateActiveSheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RecalculateActiveSheet, keyCodes.F9, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DisplayFunctionsSheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DisplayFunctionsSheet, keyCodes.KeyBackquote, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellEditorSwitchReference] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellEditorSwitchReference, keyCodes.F4, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenNumberFormatDialog] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenNumberFormatDialog, keyCodes.Digit1, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellGeneralFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellGeneralFormat, keyCodes.KeyBackquote, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellCurrencyFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellCurrencyFormat, keyCodes.Digit4, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellPercentFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellPercentFormat, keyCodes.Digit5, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellExponentialFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellExponentialFormat, keyCodes.Digit6, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellDateFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellDateFormat, keyCodes.Digit3, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellTimeFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellTimeFormat, keyCodes.Digit2, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellNumberFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellNumberFormat, keyCodes.Digit1, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditShape] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditShape, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditChart] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditChart, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepBottom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepBottom, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusNextObject] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusNextObject, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusPreviousObject] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusPreviousObject, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingAddTab] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingAddTab, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript, keyCodes.Period, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript, keyCodes.Comma, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize, keyCodes.BracketRight, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize, keyCodes.BracketLeft, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara, keyCodes.KeyE, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara, keyCodes.KeyJ, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara, keyCodes.KeyR, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara, keyCodes.KeyL, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EndParagraph] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EndParagraph, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.AddLineBreak] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.AddLineBreak, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveGraphicalObject] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveGraphicalObject, keyCodes.Delete, false, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveGraphicalObject, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ExitAddingShapesMode] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ExitAddingShapesMode, keyCodes.Esc, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker, keyCodes.KeyZ, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash, keyCodes.KeyMinus, true, true, false, false)];

	if (AscCommon.AscBrowser.isMacOs) {
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenFilePanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenFilePanel, keyCodes.KeyF, true, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentField] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentField, keyCodes.KeyA, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenChatPanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenChatPanel, keyCodes.KeyQ, true, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeftNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeftNonBlank, keyCodes.ArrowLeft, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomNonBlank, keyCodes.ArrowDown, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopNonBlank, keyCodes.ArrowUp, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNonblankLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNonblankLeft, keyCodes.Key, false, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankDown, keyCodes.Key, false, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankUp, keyCodes.Key, false, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Cut] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Cut, keyCodes.KeyX, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Copy] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Copy, keyCodes.KeyC, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Paste] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Paste, keyCodes.KeyV, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Italic] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Italic, keyCodes.KeyI, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordLeft, keyCodes.Delete, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordRight, keyCodes.Delete, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveWordLeft, keyCodes.ArrowLeft, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveWordRight, keyCodes.ArrowRight, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectWordLeft, keyCodes.ArrowLeft, false, true, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectWordRight, keyCodes.ArrowRight, false, true, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues, keyCodes.KeyC, true, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues, keyCodes.KeyC, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect, keyCodes.KeyS, true, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect, keyCodes.KeyS, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction, keyCodes.KeyEqual, true, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepRight, keyCodes.ArrowRight, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepLeft, keyCodes.ArrowLeft, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepUp, keyCodes.ArrowUp, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepBottom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepBottom, keyCodes.ArrowDown, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenFindDialog].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenFindDialog, keyCodes.KeyF, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentsPanel].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentsPanel, keyCodes.KeyH, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Save].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Save, keyCodes.KeyS, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Print].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Print, keyCodes.KeyP, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DownloadAs].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DownloadAs, keyCodes.KeyS, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ResetZoom].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ResetZoom, keyCodes.Digit0, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank, keyCodes.ArrowRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstCell].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstCell, keyCodes.Home, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveEndSpreadsheet].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveEndSpreadsheet, keyCodes.End, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn, keyCodes.KeyEqual, false, false, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn, keyCodes.NumpadPlus, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut, keyCodes.KeyMinus, false, false, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut, keyCodes.NumpadMinus, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectColumn].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectColumn, keyCodes.Space, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankRight].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankRight, keyCodes.ArrowRight, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankLeft].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankLeft, keyCodes.ArrowLeft, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankUp].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankUp, keyCodes.ArrowUp, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankDown].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankDown, keyCodes.ArrowDown, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningWorksheet].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningWorksheet, keyCodes.Home, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectLastUsedCell].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectLastUsedCell, keyCodes.End, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditUndo].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditUndo, keyCodes.KeyZ, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditRedo].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditRedo, keyCodes.KeyY, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.InsertHyperlink].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.InsertHyperlink, keyCodes.KeyK, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Bold].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Bold, keyCodes.KeyB, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Underline].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Underline, keyCodes.KeyU, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Strikeout].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Strikeout, keyCodes.Digit5, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.ToggleAutoFilter].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.ToggleAutoFilter, keyCodes.KeyL, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.FormatAsTableTemplate].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.FormatAsTableTemplate, keyCodes.KeyL, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow, keyCodes.KeyEqual, false, true, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow, keyCodes.NumpadPlus, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow, keyCodes.KeyMinus, false, true, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow, keyCodes.NumpadMinus, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate, keyCodes.Semicolon, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime, keyCodes.Semicolon, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll, keyCodes.KeyA, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningText, keyCodes.Home, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveEndText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveEndText, keyCodes.End, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningText, keyCodes.Home, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectEndText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectEndText, keyCodes.End, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningLine].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningLine, keyCodes.ArrowLeft, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveEndLine].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveEndLine, keyCodes.ArrowRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.FormatTableAddSummaryRow].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.FormatTableAddSummaryRow, keyCodes.KeyR, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenNumberFormatDialog].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenNumberFormatDialog, keyCodes.Digit1, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript, keyCodes.Period, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript, keyCodes.Comma, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize, keyCodes.BracketRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize, keyCodes.BracketLeft, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara, keyCodes.KeyE, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara, keyCodes.KeyJ, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara, keyCodes.KeyR, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara, keyCodes.KeyL, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker, keyCodes.KeyZ, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash, keyCodes.KeyMinus, false, true, false, true));
	} else {
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenFilePanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenFilePanel, keyCodes.KeyF, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentField] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentField, keyCodes.KeyH, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.OpenChatPanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.OpenChatPanel, keyCodes.KeyQ, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeftNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeftNonBlank, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomNonBlank, keyCodes.ArrowDown, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopNonBlank, keyCodes.ArrowUp, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNonblankLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNonblankLeft, keyCodes.ArrowLeft, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankDown, keyCodes.ArrowDown, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankUp, keyCodes.ArrowUp, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Cut] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Cut, keyCodes.KeyX, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Cut, keyCodes.Delete, false, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Copy] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Copy, keyCodes.KeyC, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Copy, keyCodes.Insert, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Paste] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Paste, keyCodes.KeyV, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Paste, keyCodes.Insert, false, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.Italic] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.Italic, keyCodes.KeyI, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellAddSeparator] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellAddSeparator, keyCodes.NumpadDecimal, false, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordLeft, keyCodes.Backspace, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordRight, keyCodes.Delete, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveWordLeft, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveWordRight, keyCodes.ArrowRight, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectWordLeft, keyCodes.ArrowLeft, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectWordRight, keyCodes.ArrowRight, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues, keyCodes.KeyC, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect, keyCodes.KeyS, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction, keyCodes.KeyEqual, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepRight, keyCodes.ArrowRight, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepLeft, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepUp, keyCodes.ArrowUp, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepBottom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepBottom, keyCodes.ArrowDown, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank, keyCodes.ArrowRight, true, false, false, false));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankRight].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankRight, keyCodes.ArrowRight, true, true, false, false));
	}

	if (window["AscDesktopEditor"]) {
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NextFileTab] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.NextFileTab, keyCodes.Tab, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PreviousFileTab] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PreviousFileTab, keyCodes.Tab, true, true, false, false)];
		if (AscCommon.AscBrowser.isMacOs) {
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenExistingFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenExistingFile, keyCodes.KeyO, false, false, false, true)];
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseFile, keyCodes.KeyW, false, false, false, true)];
		} else {
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenExistingFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenExistingFile, keyCodes.KeyO, true, false, false, false)];
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseFile, keyCodes.KeyW, true, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseFile, keyCodes.F4, true, false, false, false)];
		}
	}

	function getStringFromShortcutType(type) {
		switch (type) {
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenFilePanel:
				return "OpenFilePanel";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenFindDialog:
				return "OpenFindDialog";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenFindReplaceMenu:
				return "OpenFindReplaceMenu";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentsPanel:
				return "OpenCommentsPanel";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentField:
				return "OpenCommentField";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenChatPanel:
				return "OpenChatPanel";
			case Asc.c_oAscSpreadsheetShortcutTypes.Save:
				return "Save";
			case Asc.c_oAscSpreadsheetShortcutTypes.Print:
				return "Print";
			case Asc.c_oAscSpreadsheetShortcutTypes.DownloadAs:
				return "DownloadAs";
			case Asc.c_oAscSpreadsheetShortcutTypes.HelpMenu:
				return "HelpMenu";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenExistingFile:
				return "OpenExistingFile";
			case Asc.c_oAscSpreadsheetShortcutTypes.NextFileTab:
				return "NextFileTab";
			case Asc.c_oAscSpreadsheetShortcutTypes.PreviousFileTab:
				return "PreviousFileTab";
			case Asc.c_oAscSpreadsheetShortcutTypes.CloseFile:
				return "CloseFile";
			case Asc.c_oAscSpreadsheetShortcutTypes.ElementContextualMenu:
				return "ElementContextualMenu";
			case Asc.c_oAscSpreadsheetShortcutTypes.CloseMenuModal:
				return "CloseMenuModal";
			case Asc.c_oAscSpreadsheetShortcutTypes.ResetZoom:
				return "ResetZoom";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveUp:
				return "CellMoveUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveDown:
				return "CellMoveDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeft:
				return "CellMoveLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRight:
				return "CellMoveRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellDown:
				return "CellMoveActiveCellDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellUp:
				return "CellMoveActiveCellUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellRight:
				return "CellMoveActiveCellRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellLeft:
				return "CellMoveActiveCellLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeftNonBlank:
				return "CellMoveLeftNonBlank";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstColumn:
				return "CellMoveFirstColumn";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank:
				return "CellMoveRightNonBlank";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomNonBlank:
				return "CellMoveBottomNonBlank";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomEdge:
				return "CellMoveBottomEdge";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopNonBlank:
				return "CellMoveTopNonBlank";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopEdge:
				return "CellMoveTopEdge";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstCell:
				return "CellMoveFirstCell";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellMoveEndSpreadsheet:
				return "CellMoveEndSpreadsheet";
			case Asc.c_oAscSpreadsheetShortcutTypes.PreviousWorksheet:
				return "PreviousWorksheet";
			case Asc.c_oAscSpreadsheetShortcutTypes.NextWorksheet:
				return "NextWorksheet";
			case Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn:
				return "ZoomIn";
			case Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut:
				return "ZoomOut";
			case Asc.c_oAscSpreadsheetShortcutTypes.NavigatePreviousControl:
				return "NavigatePreviousControl";
			case Asc.c_oAscSpreadsheetShortcutTypes.NavigateNextControl:
				return "NavigateNextControl";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectColumn:
				return "SelectColumn";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectRow:
				return "SelectRow";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellRight:
				return "SelectOneCellRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellLeft:
				return "SelectOneCellLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellUp:
				return "SelectOneCellUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellDown:
				return "SelectOneCellDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorBeginningRow:
				return "SelectCursorBeginningRow";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorEndRow:
				return "SelectCursorEndRow";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankRight:
				return "SelectNextNonblankRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankLeft:
				return "SelectNextNonblankLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankUp:
				return "SelectNextNonblankUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankDown:
				return "SelectNextNonblankDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningWorksheet:
				return "SelectBeginningWorksheet";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectLastUsedCell:
				return "SelectLastUsedCell";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankRight:
				return "SelectNearestNonblankRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNonblankLeft:
				return "SelectNonblankLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectFirstColumn:
				return "SelectFirstColumn";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankDown:
				return "SelectNearestNonblankDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankUp:
				return "SelectNearestNonblankUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectDownOneScreen:
				return "SelectDownOneScreen";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectUpOneScreen:
				return "SelectUpOneScreen";
			case Asc.c_oAscSpreadsheetShortcutTypes.EditUndo:
				return "EditUndo";
			case Asc.c_oAscSpreadsheetShortcutTypes.EditRedo:
				return "EditRedo";
			case Asc.c_oAscSpreadsheetShortcutTypes.Cut:
				return "Cut";
			case Asc.c_oAscSpreadsheetShortcutTypes.Copy:
				return "Copy";
			case Asc.c_oAscSpreadsheetShortcutTypes.Paste:
				return "Paste";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormula:
				return "PasteOnlyFormula";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNumberFormat:
				return "PasteFormulaNumberFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaAllFormatting:
				return "PasteFormulaAllFormatting";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNoBorders:
				return "PasteFormulaNoBorders";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaColumnWidth:
				return "PasteFormulaColumnWidth";
			case Asc.c_oAscSpreadsheetShortcutTypes.Transpose:
				return "Transpose";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyValue:
				return "PasteOnlyValue";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteValueNumberFormat:
				return "PasteValueNumberFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteValueAllFormatting:
				return "PasteValueAllFormatting";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormatting:
				return "PasteOnlyFormatting";
			case Asc.c_oAscSpreadsheetShortcutTypes.PasteLink:
				return "PasteLink";
			case Asc.c_oAscSpreadsheetShortcutTypes.InsertHyperlink:
				return "InsertHyperlink";
			case Asc.c_oAscSpreadsheetShortcutTypes.VisitHyperlink:
				return "VisitHyperlink";
			case Asc.c_oAscSpreadsheetShortcutTypes.Bold:
				return "Bold";
			case Asc.c_oAscSpreadsheetShortcutTypes.Italic:
				return "Italic";
			case Asc.c_oAscSpreadsheetShortcutTypes.Underline:
				return "Underline";
			case Asc.c_oAscSpreadsheetShortcutTypes.Strikeout:
				return "Strikeout";
			case Asc.c_oAscSpreadsheetShortcutTypes.EditOpenCellEditor:
				return "EditOpenCellEditor";
			case Asc.c_oAscSpreadsheetShortcutTypes.ToggleAutoFilter:
				return "ToggleAutoFilter";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenFilterWindow:
				return "OpenFilterWindow";
			case Asc.c_oAscSpreadsheetShortcutTypes.FormatAsTableTemplate:
				return "FormatAsTableTemplate";
			case Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveDown:
				return "CompleteCellEntryMoveDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveUp:
				return "CompleteCellEntryMoveUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveRight:
				return "CompleteCellEntryMoveRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveLeft:
				return "CompleteCellEntryMoveLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryStay:
				return "CompleteCellEntryStay";
			case Asc.c_oAscSpreadsheetShortcutTypes.FillSelectedCellRange:
				return "FillSelectedCellRange";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellStartNewLine:
				return "CellStartNewLine";
			case Asc.c_oAscSpreadsheetShortcutTypes.AddPlaceholderEquation:
				return "AddPlaceholderEquation";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellEntryCancel:
				return "CellEntryCancel";
			case Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharLeft:
				return "RemoveCharLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharRight:
				return "RemoveCharRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.ClearActiveCellContent:
				return "ClearActiveCellContent";
			case Asc.c_oAscSpreadsheetShortcutTypes.ClearSelectedCellsContent:
				return "ClearSelectedCellsContent";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow:
				return "OpenInsertCellsWindow";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow:
				return "OpenDeleteCellsWindow";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate:
				return "CellInsertDate";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime:
				return "CellInsertTime";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellAddSeparator:
				return "CellAddSeparator";
			case Asc.c_oAscSpreadsheetShortcutTypes.AutoFill:
				return "AutoFill";
			case Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordLeft:
				return "RemoveWordLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordRight:
				return "RemoveWordRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll:
				return "EditSelectAll";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterLeft:
				return "MoveCharacterLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterRight:
				return "MoveCharacterRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineUp:
				return "MoveCursorLineUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineDown:
				return "MoveCursorLineDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterRight:
				return "SelectCharacterRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterLeft:
				return "SelectCharacterLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveWordLeft:
				return "MoveWordLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveWordRight:
				return "MoveWordRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectWordLeft:
				return "SelectWordLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectWordRight:
				return "SelectWordRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningText:
				return "MoveBeginningText";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveEndText:
				return "MoveEndText";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningText:
				return "SelectBeginningText";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectEndText:
				return "SelectEndText";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningLine:
				return "MoveBeginningLine";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveEndLine:
				return "MoveEndLine";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningLine:
				return "SelectBeginningLine";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectEndLine:
				return "SelectEndLine";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectLineUp:
				return "SelectLineUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.SelectLineDown:
				return "SelectLineDown";
			case Asc.c_oAscSpreadsheetShortcutTypes.RefreshSelectedPivots:
				return "RefreshSelectedPivots";
			case Asc.c_oAscSpreadsheetShortcutTypes.RefreshAllPivots:
				return "RefreshAllPivots";
			case Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues:
				return "SlicerClearSelectedValues";
			case Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect:
				return "SlicerSwitchMultiSelect";
			case Asc.c_oAscSpreadsheetShortcutTypes.FormatTableAddSummaryRow:
				return "FormatTableAddSummaryRow";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertFunctionDialog:
				return "OpenInsertFunctionDialog";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction:
				return "CellInsertSumFunction";
			case Asc.c_oAscSpreadsheetShortcutTypes.RecalculateAll:
				return "RecalculateAll";
			case Asc.c_oAscSpreadsheetShortcutTypes.RecalculateActiveSheet:
				return "RecalculateActiveSheet";
			case Asc.c_oAscSpreadsheetShortcutTypes.DisplayFunctionsSheet:
				return "DisplayFunctionsSheet";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellEditorSwitchReference:
				return "CellEditorSwitchReference";
			case Asc.c_oAscSpreadsheetShortcutTypes.OpenNumberFormatDialog:
				return "OpenNumberFormatDialog";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellGeneralFormat:
				return "CellGeneralFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellCurrencyFormat:
				return "CellCurrencyFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellPercentFormat:
				return "CellPercentFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellExponentialFormat:
				return "CellExponentialFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellDateFormat:
				return "CellDateFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellTimeFormat:
				return "CellTimeFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.CellNumberFormat:
				return "CellNumberFormat";
			case Asc.c_oAscSpreadsheetShortcutTypes.EditShape:
				return "EditShape";
			case Asc.c_oAscSpreadsheetShortcutTypes.EditChart:
				return "EditChart";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepRight:
				return "MoveShapeLittleStepRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepLeft:
				return "MoveShapeLittleStepLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepUp:
				return "MoveShapeLittleStepUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepBottom:
				return "MoveShapeLittleStepBottom";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepLeft:
				return "MoveShapeBigStepLeft";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepRight:
				return "MoveShapeBigStepRight";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepUp:
				return "MoveShapeBigStepUp";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepBottom:
				return "MoveShapeBigStepBottom";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusNextObject:
				return "MoveFocusNextObject";
			case Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusPreviousObject:
				return "MoveFocusPreviousObject";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingAddTab:
				return "DrawingAddTab";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript:
				return "DrawingSubscript";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript:
				return "DrawingSuperscript";
			case Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize:
				return "IncreaseFontSize";
			case Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize:
				return "DecreaseFontSize";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara:
				return "DrawingCenterPara";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara:
				return "DrawingJustifyPara";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara:
				return "DrawingRightPara";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara:
				return "DrawingLeftPara";
			case Asc.c_oAscSpreadsheetShortcutTypes.EndParagraph:
				return "EndParagraph";
			case Asc.c_oAscSpreadsheetShortcutTypes.AddLineBreak:
				return "AddLineBreak";
			case Asc.c_oAscSpreadsheetShortcutTypes.RemoveGraphicalObject:
				return "RemoveGraphicalObject";
			case Asc.c_oAscSpreadsheetShortcutTypes.ExitAddingShapesMode:
				return "ExitAddingShapesMode";
			case Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker:
				return "SpeechWorker";
			case Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash:
				return "DrawingEnDash";

			default:
				return null;
		}
	}

	function getShortcutTypeFromString(str) {
		switch (str) {
			case "OpenFilePanel":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenFilePanel;
			case "OpenFindDialog":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenFindDialog;
			case "OpenFindReplaceMenu":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenFindReplaceMenu;
			case "OpenCommentsPanel":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentsPanel;
			case "OpenCommentField":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenCommentField;
			case "OpenChatPanel":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenChatPanel;
			case "Save":
				return Asc.c_oAscSpreadsheetShortcutTypes.Save;
			case "Print":
				return Asc.c_oAscSpreadsheetShortcutTypes.Print;
			case "DownloadAs":
				return Asc.c_oAscSpreadsheetShortcutTypes.DownloadAs;
			case "HelpMenu":
				return Asc.c_oAscSpreadsheetShortcutTypes.HelpMenu;
			case "OpenExistingFile":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenExistingFile;
			case "NextFileTab":
				return Asc.c_oAscSpreadsheetShortcutTypes.NextFileTab;
			case "PreviousFileTab":
				return Asc.c_oAscSpreadsheetShortcutTypes.PreviousFileTab;
			case "CloseFile":
				return Asc.c_oAscSpreadsheetShortcutTypes.CloseFile;
			case "ElementContextualMenu":
				return Asc.c_oAscSpreadsheetShortcutTypes.ElementContextualMenu;
			case "CloseMenuModal":
				return Asc.c_oAscSpreadsheetShortcutTypes.CloseMenuModal;
			case "ResetZoom":
				return Asc.c_oAscSpreadsheetShortcutTypes.ResetZoom;
			case "CellMoveUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveUp;
			case "CellMoveDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveDown;
			case "CellMoveLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeft;
			case "CellMoveRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRight;
			case "CellMoveActiveCellDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellDown;
			case "CellMoveActiveCellUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellUp;
			case "CellMoveActiveCellRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellRight;
			case "CellMoveActiveCellLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveActiveCellLeft;
			case "CellMoveLeftNonBlank":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveLeftNonBlank;
			case "CellMoveFirstColumn":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstColumn;
			case "CellMoveRightNonBlank":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveRightNonBlank;
			case "CellMoveBottomNonBlank":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomNonBlank;
			case "CellMoveBottomEdge":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveBottomEdge;
			case "CellMoveTopNonBlank":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopNonBlank;
			case "CellMoveTopEdge":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveTopEdge;
			case "CellMoveFirstCell":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveFirstCell;
			case "CellMoveEndSpreadsheet":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellMoveEndSpreadsheet;
			case "PreviousWorksheet":
				return Asc.c_oAscSpreadsheetShortcutTypes.PreviousWorksheet;
			case "NextWorksheet":
				return Asc.c_oAscSpreadsheetShortcutTypes.NextWorksheet;
			case "ZoomIn":
				return Asc.c_oAscSpreadsheetShortcutTypes.ZoomIn;
			case "ZoomOut":
				return Asc.c_oAscSpreadsheetShortcutTypes.ZoomOut;
			case "NavigatePreviousControl":
				return Asc.c_oAscSpreadsheetShortcutTypes.NavigatePreviousControl;
			case "NavigateNextControl":
				return Asc.c_oAscSpreadsheetShortcutTypes.NavigateNextControl;
			case "SelectColumn":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectColumn;
			case "SelectRow":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectRow;
			case "SelectOneCellRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellRight;
			case "SelectOneCellLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellLeft;
			case "SelectOneCellUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellUp;
			case "SelectOneCellDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectOneCellDown;
			case "SelectCursorBeginningRow":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorBeginningRow;
			case "SelectCursorEndRow":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectCursorEndRow;
			case "SelectNextNonblankRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankRight;
			case "SelectNextNonblankLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankLeft;
			case "SelectNextNonblankUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankUp;
			case "SelectNextNonblankDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNextNonblankDown;
			case "SelectBeginningWorksheet":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningWorksheet;
			case "SelectLastUsedCell":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectLastUsedCell;
			case "SelectNearestNonblankRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankRight;
			case "SelectNonblankLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNonblankLeft;
			case "SelectFirstColumn":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectFirstColumn;
			case "SelectNearestNonblankDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankDown;
			case "SelectNearestNonblankUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectNearestNonblankUp;
			case "SelectDownOneScreen":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectDownOneScreen;
			case "SelectUpOneScreen":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectUpOneScreen;
			case "EditUndo":
				return Asc.c_oAscSpreadsheetShortcutTypes.EditUndo;
			case "EditRedo":
				return Asc.c_oAscSpreadsheetShortcutTypes.EditRedo;
			case "Cut":
				return Asc.c_oAscSpreadsheetShortcutTypes.Cut;
			case "Copy":
				return Asc.c_oAscSpreadsheetShortcutTypes.Copy;
			case "Paste":
				return Asc.c_oAscSpreadsheetShortcutTypes.Paste;
			case "PasteOnlyFormula":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormula;
			case "PasteFormulaNumberFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNumberFormat;
			case "PasteFormulaAllFormatting":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaAllFormatting;
			case "PasteFormulaNoBorders":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaNoBorders;
			case "PasteFormulaColumnWidth":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteFormulaColumnWidth;
			case "Transpose":
				return Asc.c_oAscSpreadsheetShortcutTypes.Transpose;
			case "PasteOnlyValue":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyValue;
			case "PasteValueNumberFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteValueNumberFormat;
			case "PasteValueAllFormatting":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteValueAllFormatting;
			case "PasteOnlyFormatting":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteOnlyFormatting;
			case "PasteLink":
				return Asc.c_oAscSpreadsheetShortcutTypes.PasteLink;
			case "InsertHyperlink":
				return Asc.c_oAscSpreadsheetShortcutTypes.InsertHyperlink;
			case "VisitHyperlink":
				return Asc.c_oAscSpreadsheetShortcutTypes.VisitHyperlink;
			case "Bold":
				return Asc.c_oAscSpreadsheetShortcutTypes.Bold;
			case "Italic":
				return Asc.c_oAscSpreadsheetShortcutTypes.Italic;
			case "Underline":
				return Asc.c_oAscSpreadsheetShortcutTypes.Underline;
			case "Strikeout":
				return Asc.c_oAscSpreadsheetShortcutTypes.Strikeout;
			case "EditOpenCellEditor":
				return Asc.c_oAscSpreadsheetShortcutTypes.EditOpenCellEditor;
			case "ToggleAutoFilter":
				return Asc.c_oAscSpreadsheetShortcutTypes.ToggleAutoFilter;
			case "OpenFilterWindow":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenFilterWindow;
			case "FormatAsTableTemplate":
				return Asc.c_oAscSpreadsheetShortcutTypes.FormatAsTableTemplate;
			case "CompleteCellEntryMoveDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveDown;
			case "CompleteCellEntryMoveUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveUp;
			case "CompleteCellEntryMoveRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveRight;
			case "CompleteCellEntryMoveLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryMoveLeft;
			case "CompleteCellEntryStay":
				return Asc.c_oAscSpreadsheetShortcutTypes.CompleteCellEntryStay;
			case "FillSelectedCellRange":
				return Asc.c_oAscSpreadsheetShortcutTypes.FillSelectedCellRange;
			case "CellStartNewLine":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellStartNewLine;
			case "AddPlaceholderEquation":
				return Asc.c_oAscSpreadsheetShortcutTypes.AddPlaceholderEquation;
			case "CellEntryCancel":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellEntryCancel;
			case "RemoveCharLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharLeft;
			case "RemoveCharRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.RemoveCharRight;
			case "ClearActiveCellContent":
				return Asc.c_oAscSpreadsheetShortcutTypes.ClearActiveCellContent;
			case "ClearSelectedCellsContent":
				return Asc.c_oAscSpreadsheetShortcutTypes.ClearSelectedCellsContent;
			case "OpenInsertCellsWindow":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertCellsWindow;
			case "OpenDeleteCellsWindow":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenDeleteCellsWindow;
			case "CellInsertDate":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellInsertDate;
			case "CellInsertTime":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellInsertTime;
			case "CellAddSeparator":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellAddSeparator;
			case "AutoFill":
				return Asc.c_oAscSpreadsheetShortcutTypes.AutoFill;
			case "RemoveWordLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordLeft;
			case "RemoveWordRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.RemoveWordRight;
			case "EditSelectAll":
				return Asc.c_oAscSpreadsheetShortcutTypes.EditSelectAll;
			case "MoveCharacterLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterLeft;
			case "MoveCharacterRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveCharacterRight;
			case "MoveCursorLineUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineUp;
			case "MoveCursorLineDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveCursorLineDown;
			case "SelectCharacterRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterRight;
			case "SelectCharacterLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectCharacterLeft;
			case "MoveWordLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveWordLeft;
			case "MoveWordRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveWordRight;
			case "SelectWordLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectWordLeft;
			case "SelectWordRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectWordRight;
			case "MoveBeginningText":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningText;
			case "MoveEndText":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveEndText;
			case "SelectBeginningText":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningText;
			case "SelectEndText":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectEndText;
			case "MoveBeginningLine":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveBeginningLine;
			case "MoveEndLine":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveEndLine;
			case "SelectBeginningLine":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectBeginningLine;
			case "SelectEndLine":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectEndLine;
			case "SelectLineUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectLineUp;
			case "SelectLineDown":
				return Asc.c_oAscSpreadsheetShortcutTypes.SelectLineDown;
			case "RefreshSelectedPivots":
				return Asc.c_oAscSpreadsheetShortcutTypes.RefreshSelectedPivots;
			case "RefreshAllPivots":
				return Asc.c_oAscSpreadsheetShortcutTypes.RefreshAllPivots;
			case "SlicerClearSelectedValues":
				return Asc.c_oAscSpreadsheetShortcutTypes.SlicerClearSelectedValues;
			case "SlicerSwitchMultiSelect":
				return Asc.c_oAscSpreadsheetShortcutTypes.SlicerSwitchMultiSelect;
			case "FormatTableAddSummaryRow":
				return Asc.c_oAscSpreadsheetShortcutTypes.FormatTableAddSummaryRow;
			case "OpenInsertFunctionDialog":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenInsertFunctionDialog;
			case "CellInsertSumFunction":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellInsertSumFunction;
			case "RecalculateAll":
				return Asc.c_oAscSpreadsheetShortcutTypes.RecalculateAll;
			case "RecalculateActiveSheet":
				return Asc.c_oAscSpreadsheetShortcutTypes.RecalculateActiveSheet;
			case "DisplayFunctionsSheet":
				return Asc.c_oAscSpreadsheetShortcutTypes.DisplayFunctionsSheet;
			case "CellEditorSwitchReference":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellEditorSwitchReference;
			case "OpenNumberFormatDialog":
				return Asc.c_oAscSpreadsheetShortcutTypes.OpenNumberFormatDialog;
			case "CellGeneralFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellGeneralFormat;
			case "CellCurrencyFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellCurrencyFormat;
			case "CellPercentFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellPercentFormat;
			case "CellExponentialFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellExponentialFormat;
			case "CellDateFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellDateFormat;
			case "CellTimeFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellTimeFormat;
			case "CellNumberFormat":
				return Asc.c_oAscSpreadsheetShortcutTypes.CellNumberFormat;
			case "EditShape":
				return Asc.c_oAscSpreadsheetShortcutTypes.EditShape;
			case "EditChart":
				return Asc.c_oAscSpreadsheetShortcutTypes.EditChart;
			case "MoveShapeLittleStepRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepRight;
			case "MoveShapeLittleStepLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepLeft;
			case "MoveShapeLittleStepUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepUp;
			case "MoveShapeLittleStepBottom":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeLittleStepBottom;
			case "MoveShapeBigStepLeft":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepLeft;
			case "MoveShapeBigStepRight":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepRight;
			case "MoveShapeBigStepUp":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepUp;
			case "MoveShapeBigStepBottom":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveShapeBigStepBottom;
			case "MoveFocusNextObject":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusNextObject;
			case "MoveFocusPreviousObject":
				return Asc.c_oAscSpreadsheetShortcutTypes.MoveFocusPreviousObject;
			case "DrawingAddTab":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingAddTab;
			case "DrawingSubscript":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingSubscript;
			case "DrawingSuperscript":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingSuperscript;
			case "IncreaseFontSize":
				return Asc.c_oAscSpreadsheetShortcutTypes.IncreaseFontSize;
			case "DecreaseFontSize":
				return Asc.c_oAscSpreadsheetShortcutTypes.DecreaseFontSize;
			case "DrawingCenterPara":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingCenterPara;
			case "DrawingJustifyPara":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingJustifyPara;
			case "DrawingRightPara":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingRightPara;
			case "DrawingLeftPara":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingLeftPara;
			case "EndParagraph":
				return Asc.c_oAscSpreadsheetShortcutTypes.EndParagraph;
			case "AddLineBreak":
				return Asc.c_oAscSpreadsheetShortcutTypes.AddLineBreak;
			case "RemoveGraphicalObject":
				return Asc.c_oAscSpreadsheetShortcutTypes.RemoveGraphicalObject;
			case "ExitAddingShapesMode":
				return Asc.c_oAscSpreadsheetShortcutTypes.ExitAddingShapesMode;
			case "SpeechWorker":
				return Asc.c_oAscSpreadsheetShortcutTypes.SpeechWorker;
			case "DrawingEnDash":
				return Asc.c_oAscSpreadsheetShortcutTypes.DrawingEnDash;
			default:
				return null;
		}
	}


	window["Asc"]["c_oAscDefaultShortcuts"] = window["Asc"].c_oAscDefaultShortcuts = c_oAscDefaultShortcuts;
	window["Asc"]["c_oAscUnlockedShortcutActionTypes"] = window["Asc"].c_oAscUnlockedShortcutActionTypes = c_oAscUnlockedShortcutActionTypes;
	window["AscCommon"].getStringFromShortcutType = getStringFromShortcutType;
	window["AscCommon"].getShortcutTypeFromString = getShortcutTypeFromString;
})();
