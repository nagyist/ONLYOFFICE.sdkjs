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

AscDFH.changesFactory[AscDFH.historyitem_ParaRun_AddItem]               = CChangesRunAddItem;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RemoveItem]            = CChangesRunRemoveItem;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Bold]                  = CChangesRunBold;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Italic]                = CChangesRunItalic;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Strikeout]             = CChangesRunStrikeout;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Underline]             = CChangesRunUnderline;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_FontFamily]            = undefined; // obsolete
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_FontSize]              = CChangesRunFontSize;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Color]                 = CChangesRunColor;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_VertAlign]             = CChangesRunVertAlign;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_HighLight]             = CChangesRunHighLight;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_HighlightColor]        = CChangesRunHighlightColor;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RStyle]                = CChangesRunRStyle;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Spacing]               = CChangesRunSpacing;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_DStrikeout]            = CChangesRunDStrikeout;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Caps]                  = CChangesRunCaps;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_SmallCaps]             = CChangesRunSmallCaps;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Position]              = CChangesRunPosition;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Value]                 = undefined; // obsolete
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts]                = CChangesRunRFonts;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Lang]                  = CChangesRunLang;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_Ascii]          = CChangesRunRFontsAscii;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_HAnsi]          = CChangesRunRFontsHAnsi;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_CS]             = CChangesRunRFontsCS;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_EastAsia]       = CChangesRunRFontsEastAsia;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_Hint]           = CChangesRunRFontsHint;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Lang_Bidi]             = CChangesRunLangBidi;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Lang_EastAsia]         = CChangesRunLangEastAsia;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Lang_Val]              = CChangesRunLangVal;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_TextPr]                = CChangesRunTextPr;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Unifill]               = CChangesRunUnifill;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Shd]                   = CChangesRunShd;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_MathStyle]             = CChangesRunMathStyle;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_MathPrp]               = CChangesRunMathPrp;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_ReviewType]            = undefined; // obsolete
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_PrChange]              = CChangesRunPrChange;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_TextFill]              = CChangesRunTextFill;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_TextOutline]           = CChangesRunTextOutline;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_PrReviewInfo]          = CChangesRunPrReviewInfo;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_ContentReviewInfo]     = CChangesRunContentReviewInfo;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_OnStartSplit]          = CChangesRunOnStartSplit;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_OnEndSplit]            = CChangesRunOnEndSplit;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_MathAlnAt]             = CChangesRunMathAlnAt;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_MathForcedBreak]       = CChangesRunMathForcedBreak;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_Ascii_Theme]    = CChangesRunRFontsAsciiTheme;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_HAnsi_Theme]    = CChangesRunRFontsHAnsiTheme;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_CS_Theme]       = CChangesRunRFontsCSTheme;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RFonts_EastAsia_Theme] = CChangesRunRFontsEastAsiaTheme;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_BoldCS]                = CChangesRunBoldCS;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_ItalicCS]              = CChangesRunItalicCS;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_FontSizeCS]            = CChangesRunFontSizeCS;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_Ligatures]             = CChangesRunLigatures;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_CS]                    = CChangesRunCS;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_RTL]                   = CChangesRunRTL;
AscDFH.changesFactory[AscDFH.historyitem_ParaRun_MathMetaData]          = CChangesRunMathMetaData;

//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_AddItem]           = [
	AscDFH.historyitem_ParaRun_AddItem,
	AscDFH.historyitem_ParaRun_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RemoveItem]        = [
	AscDFH.historyitem_ParaRun_AddItem,
	AscDFH.historyitem_ParaRun_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Bold]              = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Bold];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Italic]            = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Italic];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Strikeout]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Strikeout];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Underline]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Underline];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_FontSize]          = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_FontSize];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Color]             = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Color];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_VertAlign]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_VertAlign];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_HighLight]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_HighLight];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_HighlightColor]    = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_HighlightColor];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RStyle]            = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RStyle];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Spacing]           = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Spacing];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_DStrikeout]        = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_DStrikeout];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Caps]              = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Caps];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_SmallCaps]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_SmallCaps];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Position]          = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Position];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts]            = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_RFonts,
	AscDFH.historyitem_ParaRun_RFonts_Ascii,
	AscDFH.historyitem_ParaRun_RFonts_HAnsi,
	AscDFH.historyitem_ParaRun_RFonts_CS,
	AscDFH.historyitem_ParaRun_RFonts_EastAsia,
	AscDFH.historyitem_ParaRun_RFonts_Hint,
	AscDFH.historyitem_ParaRun_RFonts_Ascii_Theme,
	AscDFH.historyitem_ParaRun_RFonts_HAnsi_Theme,
	AscDFH.historyitem_ParaRun_RFonts_CS_Theme,
	AscDFH.historyitem_ParaRun_RFonts_EastAsia_Theme
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Lang]              = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_Lang,
	AscDFH.historyitem_ParaRun_Lang_Bidi,
	AscDFH.historyitem_ParaRun_Lang_EastAsia,
	AscDFH.historyitem_ParaRun_Lang_Val
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_Ascii]      = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_Ascii];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_HAnsi]      = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_HAnsi];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_CS]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_CS];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_EastAsia]   = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_EastAsia];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_Hint]       = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_Hint];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Lang_Bidi]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Lang, AscDFH.historyitem_ParaRun_Lang_Bidi];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Lang_EastAsia]     = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Lang, AscDFH.historyitem_ParaRun_Lang_EastAsia];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Lang_Val]          = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Lang, AscDFH.historyitem_ParaRun_Lang_Val];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_TextPr]            = [
	AscDFH.historyitem_ParaRun_Bold,
	AscDFH.historyitem_ParaRun_Italic,
	AscDFH.historyitem_ParaRun_Strikeout,
	AscDFH.historyitem_ParaRun_Underline,
	AscDFH.historyitem_ParaRun_FontSize,
	AscDFH.historyitem_ParaRun_Color,
	AscDFH.historyitem_ParaRun_VertAlign,
	AscDFH.historyitem_ParaRun_HighLight,
	AscDFH.historyitem_ParaRun_HighlightColor,
	AscDFH.historyitem_ParaRun_RStyle,
	AscDFH.historyitem_ParaRun_Spacing,
	AscDFH.historyitem_ParaRun_DStrikeout,
	AscDFH.historyitem_ParaRun_Caps,
	AscDFH.historyitem_ParaRun_SmallCaps,
	AscDFH.historyitem_ParaRun_Position,
	AscDFH.historyitem_ParaRun_RFonts,
	AscDFH.historyitem_ParaRun_Lang,
	AscDFH.historyitem_ParaRun_RFonts_Ascii,
	AscDFH.historyitem_ParaRun_RFonts_HAnsi,
	AscDFH.historyitem_ParaRun_RFonts_CS,
	AscDFH.historyitem_ParaRun_RFonts_EastAsia,
	AscDFH.historyitem_ParaRun_RFonts_Hint,
	AscDFH.historyitem_ParaRun_Lang_Bidi,
	AscDFH.historyitem_ParaRun_Lang_EastAsia,
	AscDFH.historyitem_ParaRun_Lang_Val,
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_Unifill,
	AscDFH.historyitem_ParaRun_Shd,
	AscDFH.historyitem_ParaRun_PrChange,
	AscDFH.historyitem_ParaRun_TextFill,
	AscDFH.historyitem_ParaRun_TextOutline,
	AscDFH.historyitem_ParaRun_PrReviewInfo,
	AscDFH.historyitem_ParaRun_RFonts_Ascii_Theme,
	AscDFH.historyitem_ParaRun_RFonts_HAnsi_Theme,
	AscDFH.historyitem_ParaRun_RFonts_CS_Theme,
	AscDFH.historyitem_ParaRun_RFonts_EastAsia_Theme,
	AscDFH.historyitem_ParaRun_Ligatures,
	AscDFH.historyitem_ParaRun_CS,
	AscDFH.historyitem_ParaRun_RTL
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Unifill]           = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Unifill];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Shd]               = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_Shd];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_MathStyle]         = [
	AscDFH.historyitem_ParaRun_MathStyle,
	AscDFH.historyitem_ParaRun_MathPrp
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_MathPrp]           = [
	AscDFH.historyitem_ParaRun_MathStyle,
	AscDFH.historyitem_ParaRun_MathPrp,
	AscDFH.historyitem_ParaRun_MathAlnAt,
	AscDFH.historyitem_ParaRun_MathForcedBreak
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_ReviewType]        = [
	AscDFH.historyitem_ParaRun_ReviewType,
	AscDFH.historyitem_ParaRun_ContentReviewInfo
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_PrChange]          = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_PrChange,
	AscDFH.historyitem_ParaRun_PrReviewInfo
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_TextFill]          = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_TextFill];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_TextOutline]       = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_TextOutline];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_PrReviewInfo]      = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_PrChange,
	AscDFH.historyitem_ParaRun_PrReviewInfo
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_ContentReviewInfo] = [
	AscDFH.historyitem_ParaRun_ReviewType,
	AscDFH.historyitem_ParaRun_ContentReviewInfo
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_OnStartSplit]      = [];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_OnEndSplit]        = [];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_MathAlnAt]         = [
	AscDFH.historyitem_ParaRun_MathPrp,
	AscDFH.historyitem_ParaRun_MathAlnAt,
	AscDFH.historyitem_ParaRun_MathForcedBreak
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_MathForcedBreak]   = [
	AscDFH.historyitem_ParaRun_MathPrp,
	AscDFH.historyitem_ParaRun_MathAlnAt,
	AscDFH.historyitem_ParaRun_MathForcedBreak
];

AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_Ascii_Theme]      = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_Ascii_Theme];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_HAnsi_Theme]      = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_HAnsi_Theme];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_CS_Theme]         = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_CS_Theme];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RFonts_EastAsia_Theme]   = [AscDFH.historyitem_ParaRun_TextPr, AscDFH.historyitem_ParaRun_RFonts, AscDFH.historyitem_ParaRun_RFonts_EastAsia_Theme];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_Ligatures] = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_Ligatures
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_CS] = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_CS
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_RTL] = [
	AscDFH.historyitem_ParaRun_TextPr,
	AscDFH.historyitem_ParaRun_RTL
];
AscDFH.changesRelationMap[AscDFH.historyitem_ParaRun_MathMetaData] = [AscDFH.historyitem_ParaRun_MathMetaData];

/**
 * Общая функция для загрузки измнения настроек текста
 */
function private_ParaRunChangesLoadTextPr(color)
{
	this.Redo();
	if (this.Color && color)
		this.Class.setCollPrChangeColor(color);
}
/**
 * Общая функция объединения изменений, которые зависят только от себя и AscDFH.historyitem_ParaRun_TextPr
 * @param oChange
 * @returns {boolean}
 */
function private_ParaRunChangesOnMergeTextPr(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (oChange.Type === this.Type || oChange.Type === AscDFH.historyitem_ParaRun_TextPr)
		return false;

	return true;
}
/**
 * Общая функция объединения изменений, которые зависят от себя, AscDFH.historyitem_ParaRun_TextPr,
 * AscDFH.historyitem_ParaRun_RFonts
 * @param oChange
 * @returns {boolean}
 */
function private_ParaRunChangesOnMergeRFontsTextPr(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (oChange.Type === this.Type || oChange.Type === AscDFH.historyitem_ParaRun_TextPr || oChange.Type === AscDFH.historyitem_ParaRun_RFonts)
		return false;

	return true;
}
/**
 * Общая функция объединения изменений, которые зависят от себя, AscDFH.historyitem_ParaRun_TextPr,
 * AscDFH.historyitem_ParaRun_Lang
 * @param oChange
 * @returns {boolean}
 */
function private_ParaRunChangesOnMergeLangTextPr(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (oChange.Type === this.Type || oChange.Type === AscDFH.historyitem_ParaRun_TextPr || oChange.Type === AscDFH.historyitem_ParaRun_Lang)
		return false;

	return true;
}

//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesRunAddItem(Class, Pos, Items, Color)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);

	this.Color = true === Color ? true : false;
}
CChangesRunAddItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesRunAddItem.prototype.constructor = CChangesRunAddItem;
CChangesRunAddItem.prototype.Type = AscDFH.historyitem_ParaRun_AddItem;
CChangesRunAddItem.prototype.Undo = function()
{
	var oRun = this.Class;

	oRun.Content.splice(this.Pos, this.Items.length);

	oRun.RecalcInfo.Measure = true;
	oRun.OnContentChange();
	oRun.private_UpdateTrackRevisionOnChangeContent(false);
	oRun.private_UpdatePositionsOnRemove(this.Pos, this.Items.length);
};
CChangesRunAddItem.prototype.Redo = function()
{
	var oRun = this.Class;

	var Array_start = oRun.Content.slice(0, this.Pos);
	var Array_end   = oRun.Content.slice(this.Pos);

	oRun.Content = Array_start.concat(this.Items, Array_end);

	oRun.RecalcInfo.Measure = true;
	oRun.OnContentChange();
	oRun.private_UpdateTrackRevisionOnChangeContent(false);

	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		oRun.private_UpdatePositionsOnAdd(this.Pos + nIndex);
		if (this.Items[nIndex].SetParent)
			this.Items[nIndex].SetParent(oRun);
	}
};
CChangesRunAddItem.prototype.private_WriteItem = function(Writer, Item)
{
	Item.Write_ToBinary(Writer);
};
CChangesRunAddItem.prototype.private_ReadItem = function(Reader)
{
	return AscWord.ReadRunElementFromBinary(Reader);
};
CChangesRunAddItem.prototype.Load = function(Color)
{
	var oRun = this.Class;
	for (var Index = 0, Count = this.Items.length; Index < Count; Index++)
	{
		var Pos = oRun.m_oContentChanges.Check(AscCommon.contentchanges_Add, this.PosArray[Index]);
		var Element = this.Items[Index];
		if (null != Element)
		{
			if (true === this.Color && null !== Color)
			{
				oRun.CollaborativeMarks.Update_OnAdd(Pos);
				oRun.CollaborativeMarks.Add(Pos, Pos + 1, Color);
				AscCommon.CollaborativeEditing.Add_ChangedClass(oRun);
			}

			oRun.Content.splice(Pos, 0, Element);
			oRun.private_UpdatePositionsOnAdd(Pos);
			oRun.private_UpdateCompositeInputPositionsOnAdd(Pos);
			AscCommon.CollaborativeEditing.Update_DocumentPositionsOnAdd(oRun, Pos);

			if (Element.SetParent)
				Element.SetParent(oRun);
		}
	}

	oRun.RecalcInfo.Measure = true;
	oRun.OnContentChange();
	oRun.private_UpdateTrackRevisionOnChangeContent(false);
	oRun.private_UpdateDocumentOutline();
};
CChangesRunAddItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_ParaRun_AddItem === oChanges.Type || AscDFH.historyitem_ParaRun_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesRunAddItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesRunRemoveItem);
};
CChangesRunAddItem.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesRunRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, false);
}
CChangesRunRemoveItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesRunRemoveItem.prototype.constructor = CChangesRunRemoveItem;
CChangesRunRemoveItem.prototype.Type = AscDFH.historyitem_ParaRun_RemoveItem;
CChangesRunRemoveItem.prototype.Undo = function()
{
	var oRun = this.Class;

	var Array_start = oRun.Content.slice(0, this.Pos);
	var Array_end   = oRun.Content.slice(this.Pos);

	oRun.Content = Array_start.concat(this.Items, Array_end);

	oRun.RecalcInfo.Measure = true;
	oRun.OnContentChange();
	oRun.private_UpdateTrackRevisionOnChangeContent(false);

	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		oRun.private_UpdatePositionsOnAdd(this.Pos + nIndex);
		if (this.Items[nIndex].SetParent)
			this.Items[nIndex].SetParent(oRun);
	}
};
CChangesRunRemoveItem.prototype.Redo = function()
{
	var oRun = this.Class;
	oRun.Content.splice(this.Pos, this.Items.length);

	oRun.RecalcInfo.Measure = true;
	oRun.OnContentChange();
	oRun.private_UpdateTrackRevisionOnChangeContent(false);
	oRun.private_UpdatePositionsOnRemove(this.Pos, this.Items.length);
};
CChangesRunRemoveItem.prototype.private_WriteItem = function(Writer, Item)
{
	Item.Write_ToBinary(Writer);
};
CChangesRunRemoveItem.prototype.private_ReadItem = function(Reader)
{
	return AscWord.ReadRunElementFromBinary(Reader);
};
CChangesRunRemoveItem.prototype.Load = function()
{
	var oRun = this.Class;

	var nLastChangesPos = null;
	var nChangesCount   = 0;
	for (var Index = 0, Count = this.PosArray.length; Index < Count; Index++)
	{
		var nChangesPos = oRun.m_oContentChanges.Check(AscCommon.contentchanges_Remove, this.PosArray[Index]);

		// действие совпало, не делаем его
		if (false === nChangesPos)
			continue;

		// В большинстве случаев удаления идут подряд и не меняются, т.е. их позиции у другого клиента тоже идут подряд
		// Для улучшения производительности мы объединяем подряд идущие удаления
		if (null === nLastChangesPos)
		{
			nLastChangesPos = nChangesPos;
			nChangesCount   = 1;
		}
		else if (nLastChangesPos === nChangesPos)
		{
			nChangesCount++;
		}
		else
		{
			oRun.CollaborativeMarks.Update_OnRemove(nLastChangesPos, nChangesCount);
			oRun.Content.splice(nLastChangesPos, nChangesCount);
			oRun.private_UpdatePositionsOnRemove(nLastChangesPos, nChangesCount);
			oRun.private_UpdateCompositeInputPositionsOnRemove(nLastChangesPos, nChangesCount);
			AscCommon.CollaborativeEditing.Update_DocumentPositionsOnRemove(oRun, nLastChangesPos, nChangesCount);

			nLastChangesPos = nChangesPos;
			nChangesCount   = 1;
		}
	}

	if (nChangesCount)
	{
		oRun.CollaborativeMarks.Update_OnRemove(nLastChangesPos, nChangesCount);
		oRun.Content.splice(nLastChangesPos, nChangesCount);
		oRun.private_UpdatePositionsOnRemove(nLastChangesPos, nChangesCount);
		oRun.private_UpdateCompositeInputPositionsOnRemove(nLastChangesPos, nChangesCount);
		AscCommon.CollaborativeEditing.Update_DocumentPositionsOnRemove(oRun, nLastChangesPos, nChangesCount);
	}

	oRun.RecalcInfo.Measure = true;
	oRun.OnContentChange();
	oRun.private_UpdateTrackRevisionOnChangeContent(false);
	oRun.private_UpdateDocumentOutline();
};
CChangesRunRemoveItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_ParaRun_AddItem === oChanges.Type || AscDFH.historyitem_ParaRun_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesRunRemoveItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesRunAddItem);
};
CChangesRunRemoveItem.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunBold(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunBold.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunBold.prototype.constructor = CChangesRunBold;
CChangesRunBold.prototype.Type = AscDFH.historyitem_ParaRun_Bold;
CChangesRunBold.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Bold = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
	oRun.OnContentChange();
};
CChangesRunBold.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunBold.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunBold.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunItalic(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunItalic.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunItalic.prototype.constructor = CChangesRunItalic;
CChangesRunItalic.prototype.Type = AscDFH.historyitem_ParaRun_Italic;
CChangesRunItalic.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Italic = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunItalic.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunItalic.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunItalic.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunStrikeout(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunStrikeout.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunStrikeout.prototype.constructor = CChangesRunStrikeout;
CChangesRunStrikeout.prototype.Type = AscDFH.historyitem_ParaRun_Strikeout;
CChangesRunStrikeout.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Strikeout = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunStrikeout.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunStrikeout.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunStrikeout.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunUnderline(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunUnderline.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunUnderline.prototype.constructor = CChangesRunUnderline;
CChangesRunUnderline.prototype.Type = AscDFH.historyitem_ParaRun_Underline;
CChangesRunUnderline.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Underline = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunUnderline.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunUnderline.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunUnderline.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesRunFontSize(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesRunFontSize.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesRunFontSize.prototype.constructor = CChangesRunFontSize;
CChangesRunFontSize.prototype.Type = AscDFH.historyitem_ParaRun_FontSize;
CChangesRunFontSize.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.FontSize = Value;
	const oSmartArtShape = oRun.IsInsideSmartArtShape(true);
	if (oSmartArtShape)
	{
		oSmartArtShape.resetSmartArtMaxFontSize();
	}
	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunFontSize.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunFontSize.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunFontSize.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunColor(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunColor.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunColor.prototype.constructor = CChangesRunColor;
CChangesRunColor.prototype.Type = AscDFH.historyitem_ParaRun_Color;
CChangesRunColor.prototype.private_CreateObject = function()
{
	return new CDocumentColor(0, 0, 0, false);
};
CChangesRunColor.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Color = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunColor.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunColor.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunColor.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunVertAlign(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesRunVertAlign.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunVertAlign.prototype.constructor = CChangesRunVertAlign;
CChangesRunVertAlign.prototype.Type = AscDFH.historyitem_ParaRun_VertAlign;
CChangesRunVertAlign.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.VertAlign = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunVertAlign.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunVertAlign.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunVertAlign.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunHighLight(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunHighLight.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunHighLight.prototype.constructor = CChangesRunHighLight;
CChangesRunHighLight.prototype.Type = AscDFH.historyitem_ParaRun_HighLight;
CChangesRunHighLight.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is "none" New ?
	// 4-bit : Is undefined New ?
	// 5-bit : Is "none" New ?
	// Variable(?CDocumentColor) : New (если 2 и 3 биты нулевые)
	// Variable(?CDocumentColor) : Old (если 4 и 5 биты нулевые)

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;
	else if (highlight_None === this.New)
		nFlags |= 4;

	if (undefined === this.Old)
		nFlags |= 8;
	else if (highlight_None === this.Old)
		nFlags |= 16;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New && highlight_None !== this.New)
		this.New.Write_ToBinary(Writer);

	if (undefined !== this.Old && highlight_None !== this.Old)
		this.Old.Write_ToBinary(Writer);
};
CChangesRunHighLight.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is "none" New ?
	// 4-bit : Is undefined New ?
	// 5-bit : Is "none" New ?
	// Variable(?CDocumentColor) : New (если 2 и 3 биты нулевые)
	// Variable(?CDocumentColor) : Old (если 4 и 5 биты нулевые)

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
	{
		this.New = undefined;
	}
	else if (nFlags & 4)
	{
		this.New = highlight_None;
	}
	else
	{
		this.New = new CDocumentColor(0, 0, 0);
		this.New.Read_FromBinary(Reader);
	}

	if (nFlags & 8)
	{
		this.Old = undefined;
	}
	else if (nFlags & 16)
	{
		this.Old = highlight_None;
	}
	else
	{
		this.Old = new CDocumentColor(0, 0, 0);
		this.Old.Read_FromBinary(Reader);
	}
};
CChangesRunHighLight.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.HighLight = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunHighLight.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunHighLight.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunHighLight.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunHighlightColor(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunHighlightColor.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunHighlightColor.prototype.constructor = CChangesRunHighlightColor;
CChangesRunHighlightColor.prototype.Type = AscDFH.historyitem_ParaRun_HighlightColor;
CChangesRunHighlightColor.prototype.WriteToBinary = function(Writer)
{
	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		this.New.Write_ToBinary(Writer);

	if (undefined !== this.Old)
		this.Old.Write_ToBinary(Writer);
};
CChangesRunHighlightColor.prototype.ReadFromBinary = function(Reader)
{
	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
	{
		this.New = undefined;
	}
	else
	{
		this.New = new AscFormat.CUniColor();
		this.New.Read_FromBinary(Reader);
	}

	if (nFlags & 4)
	{
		this.Old = undefined;
	}
	else
	{
		this.Old = new AscFormat.CUniColor();
		this.Old.Read_FromBinary(Reader);
	}
};
CChangesRunHighlightColor.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.HighlightColor = Value;

	oRun.Recalc_CompiledPr(false);
};
CChangesRunHighlightColor.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunHighlightColor.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunHighlightColor.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesRunRStyle(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesRunRStyle.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesRunRStyle.prototype.constructor = CChangesRunRStyle;
CChangesRunRStyle.prototype.Type = AscDFH.historyitem_ParaRun_RStyle;
CChangesRunRStyle.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RStyle = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRStyle.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRStyle.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunRStyle.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesRunSpacing(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesRunSpacing.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesRunSpacing.prototype.constructor = CChangesRunSpacing;
CChangesRunSpacing.prototype.Type = AscDFH.historyitem_ParaRun_Spacing;
CChangesRunSpacing.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Spacing = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunSpacing.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunSpacing.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunSpacing.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunDStrikeout(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunDStrikeout.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunDStrikeout.prototype.constructor = CChangesRunDStrikeout;
CChangesRunDStrikeout.prototype.Type = AscDFH.historyitem_ParaRun_DStrikeout;
CChangesRunDStrikeout.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.DStrikeout = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunDStrikeout.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunDStrikeout.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunDStrikeout.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunCaps(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunCaps.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunCaps.prototype.constructor = CChangesRunCaps;
CChangesRunCaps.prototype.Type = AscDFH.historyitem_ParaRun_Caps;
CChangesRunCaps.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Caps = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunCaps.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunCaps.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunCaps.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunSmallCaps(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunSmallCaps.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunSmallCaps.prototype.constructor = CChangesRunSmallCaps;
CChangesRunSmallCaps.prototype.Type = AscDFH.historyitem_ParaRun_SmallCaps;
CChangesRunSmallCaps.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.SmallCaps = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunSmallCaps.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunSmallCaps.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunSmallCaps.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesRunPosition(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesRunPosition.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesRunPosition.prototype.constructor = CChangesRunPosition;
CChangesRunPosition.prototype.Type = AscDFH.historyitem_ParaRun_Position;
CChangesRunPosition.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Position = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunPosition.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunPosition.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunPosition.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunRFonts(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFonts.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunRFonts.prototype.constructor = CChangesRunRFonts;
CChangesRunRFonts.prototype.Type = AscDFH.historyitem_ParaRun_RFonts;
CChangesRunRFonts.prototype.private_CreateObject = function()
{
	return new CRFonts();
};
CChangesRunRFonts.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
CChangesRunRFonts.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RFonts = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFonts.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFonts.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_TextPr === oChange.Type)
		return false;

	if (!this.New)
		this.New = new CRFonts();

	switch (oChange.Type)
	{
		case AscDFH.historyitem_ParaRun_RFonts_Ascii:
		{
			this.New.Ascii = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_HAnsi:
		{
			this.New.HAnsi = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_CS:
		{
			this.New.CS = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_EastAsia:
		{
			this.New.EastAsia = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_Hint:
		{
			this.New.Hint = oChange.New;
			break;
		}
	}

	return true;
};
CChangesRunRFonts.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunLang(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunLang.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunLang.prototype.constructor = CChangesRunLang;
CChangesRunLang.prototype.Type = AscDFH.historyitem_ParaRun_Lang;
CChangesRunLang.prototype.private_CreateObject = function()
{
	return new CLang();
};
CChangesRunLang.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
CChangesRunLang.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Lang = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateSpellChecking();
	oRun.private_UpdateShapeText();
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunLang.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunLang.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_TextPr === oChange.Type)
		return false;

	if (!this.New)
		this.New = new CLang();

	switch (oChange.Type)
	{
		case AscDFH.historyitem_ParaRun_Lang_Bidi:
		{
			this.New.Bidi = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Lang_EastAsia:
		{
			this.New.EastAsia = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Lang_Val:
		{
			this.New.Val = oChange.New;
			break;
		}
	}

	return true;
};
CChangesRunLang.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsAscii(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsAscii.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsAscii.prototype.constructor = CChangesRunRFontsAscii;
CChangesRunRFontsAscii.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_Ascii;
CChangesRunRFontsAscii.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteString2(this.New.Name);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old.Name);
};
CChangesRunRFontsAscii.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
	{
		this.New = undefined;
	}
	else
	{
		this.New = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}

	if (nFlags & 4)
	{
		this.Old = undefined;
	}
	else
	{
		this.Old = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}
};
CChangesRunRFontsAscii.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RFonts.Ascii = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFontsAscii.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFontsAscii.prototype.Merge = private_ParaRunChangesOnMergeRFontsTextPr;
CChangesRunRFontsAscii.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsHAnsi(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsHAnsi.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsHAnsi.prototype.constructor = CChangesRunRFontsHAnsi;
CChangesRunRFontsHAnsi.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_HAnsi;
CChangesRunRFontsHAnsi.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteString2(this.New.Name);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old.Name);
};
CChangesRunRFontsHAnsi.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
	{
		this.New = undefined;
	}
	else
	{
		this.New = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}

	if (nFlags & 4)
	{
		this.Old = undefined;
	}
	else
	{
		this.Old = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}
};
CChangesRunRFontsHAnsi.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RFonts.HAnsi = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFontsHAnsi.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFontsHAnsi.prototype.Merge = private_ParaRunChangesOnMergeRFontsTextPr;
CChangesRunRFontsHAnsi.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsCS(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsCS.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsCS.prototype.constructor = CChangesRunRFontsCS;
CChangesRunRFontsCS.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_CS;
CChangesRunRFontsCS.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteString2(this.New.Name);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old.Name);
};
CChangesRunRFontsCS.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
	{
		this.New = undefined;
	}
	else
	{
		this.New = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}

	if (nFlags & 4)
	{
		this.Old = undefined;
	}
	else
	{
		this.Old = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}
};
CChangesRunRFontsCS.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RFonts.CS = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFontsCS.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFontsCS.prototype.Merge = private_ParaRunChangesOnMergeRFontsTextPr;
CChangesRunRFontsCS.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsEastAsia(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsEastAsia.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsEastAsia.prototype.constructor = CChangesRunRFontsEastAsia;
CChangesRunRFontsEastAsia.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_EastAsia;
CChangesRunRFontsEastAsia.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteString2(this.New.Name);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old.Name);
};
CChangesRunRFontsEastAsia.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
	{
		this.New = undefined;
	}
	else
	{
		this.New = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}

	if (nFlags & 4)
	{
		this.Old = undefined;
	}
	else
	{
		this.Old = {
			Name  : Reader.GetString2(),
			Index : -1
		};
	}
};
CChangesRunRFontsEastAsia.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RFonts.EastAsia = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFontsEastAsia.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFontsEastAsia.prototype.Merge = private_ParaRunChangesOnMergeRFontsTextPr;
CChangesRunRFontsEastAsia.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunRFontsHint(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsHint.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunRFontsHint.prototype.constructor = CChangesRunRFontsHint;
CChangesRunRFontsHint.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_Hint;
CChangesRunRFontsHint.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.RFonts.Hint = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFontsHint.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFontsHint.prototype.Merge = private_ParaRunChangesOnMergeRFontsTextPr;
CChangesRunRFontsHint.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunLangBidi(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesRunLangBidi.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunLangBidi.prototype.constructor = CChangesRunLangBidi;
CChangesRunLangBidi.prototype.Type = AscDFH.historyitem_ParaRun_Lang_Bidi;
CChangesRunLangBidi.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Lang.Bidi = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateSpellChecking();
	oRun.private_UpdateShapeText();
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunLangBidi.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunLangBidi.prototype.Merge = private_ParaRunChangesOnMergeLangTextPr;
CChangesRunLangBidi.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunLangEastAsia(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesRunLangEastAsia.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunLangEastAsia.prototype.constructor = CChangesRunLangEastAsia;
CChangesRunLangEastAsia.prototype.Type = AscDFH.historyitem_ParaRun_Lang_EastAsia;
CChangesRunLangEastAsia.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Lang.EastAsia = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateSpellChecking();
	oRun.private_UpdateShapeText();
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunLangEastAsia.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunLangEastAsia.prototype.Merge = private_ParaRunChangesOnMergeLangTextPr;
CChangesRunLangEastAsia.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunLangVal(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesRunLangVal.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunLangVal.prototype.constructor = CChangesRunLangVal;
CChangesRunLangVal.prototype.Type = AscDFH.historyitem_ParaRun_Lang_Val;
CChangesRunLangVal.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Lang.Val = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateSpellChecking();
	oRun.private_UpdateShapeText();
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunLangVal.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunLangVal.prototype.Merge = private_ParaRunChangesOnMergeLangTextPr;
CChangesRunLangVal.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunTextPr(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunTextPr.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunTextPr.prototype.constructor = CChangesRunTextPr;
CChangesRunTextPr.prototype.Type = AscDFH.historyitem_ParaRun_TextPr;
CChangesRunTextPr.prototype.private_CreateObject = function()
{
	return new CTextPr();
};
CChangesRunTextPr.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
CChangesRunTextPr.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunTextPr.prototype.Load = function(Color)
{
	this.Redo();

	var Unifill = this.Class.Pr.Unifill;
	if (AscCommon.CollaborativeEditing
		&& Unifill
		&& Unifill.fill
		&& Unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_BLIP
		&& typeof Unifill.fill.RasterImageId === "string"
		&& Unifill.fill.RasterImageId.length > 0)
	{
		AscCommon.CollaborativeEditing.Add_NewImage(Unifill.fill.RasterImageId);
	}

	if (this.Color && Color)
		this.Class.setCollPrChangeColor(Color);
};
CChangesRunTextPr.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type)
		return false;

	if (!this.New)
		this.New = new CTextPr();

	switch (oChange.Type)
	{
		case AscDFH.historyitem_ParaRun_Bold:
		{
			this.New.Bold = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Italic:
		{
			this.New.Italic = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Strikeout:
		{
			this.New.Strikeout = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Underline:
		{
			this.New.Underline = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_FontSize:
		{
			this.New.FontSize = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Color:
		{
			this.New.Color = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_VertAlign:
		{
			this.New.VertAlign = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_HighLight:
		{
			this.New.HighLight = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_HighlightColor:
		{
			this.New.HighlightColor = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RStyle:
		{
			this.New.RStyle = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Spacing:
		{
			this.New.Spacing = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_DStrikeout:
		{
			this.New.DStrikeout = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Caps:
		{
			this.New.Caps = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_SmallCaps:
		{
			this.New.SmallCaps = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Position:
		{
			this.New.Position = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts:
		{
			this.New.RFonts = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Lang:
		{
			this.New.Lang = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_Ascii:
		{
			if (!this.New.RFonts)
				this.New.RFonts = new CRFonts();

			this.New.RFonts.Ascii = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_HAnsi:
		{
			if (!this.New.RFonts)
				this.New.RFonts = new CRFonts();

			this.New.RFonts.HAnsi = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_CS:
		{
			if (!this.New.RFonts)
				this.New.RFonts = new CRFonts();

			this.New.RFonts.CS = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_EastAsia:
		{
			if (!this.New.RFonts)
				this.New.RFonts = new CRFonts();

			this.New.RFonts.EastAsia = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_RFonts_Hint:
		{
			if (!this.New.RFonts)
				this.New.RFonts = new CRFonts();

			this.New.RFonts.Hint = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Lang_Bidi:
		{
			if (!this.New.Lang)
				this.New.Lang = new CLang();

			this.New.Lang.Bidi = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Lang_EastAsia:
		{
			if (!this.New.Lang)
				this.New.Lang = new CLang();

			this.New.Lang.EastAsia = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Lang_Val:
		{
			if (!this.New.Lang)
				this.New.Lang = new CLang();

			this.New.Lang.Val = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Unifill:
		{
			this.New.Unifill = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Shd:
		{
			this.New.Shd = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_PrChange:
		{
			this.New.PrChange   = oChange.New.PrChange;
			this.New.ReviewInfo = oChange.New.ReviewInfo;
			break;
		}
		case AscDFH.historyitem_ParaRun_TextFill:
		{
			this.New.TextFill = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_TextOutline:
		{
			this.New.TextOutline = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_PrReviewInfo:
		{
			this.New.ReviewInfo = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_BoldCS:
		{
			this.New.BoldCS = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_ItalicCS:
		{
			this.New.ItalicCS = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_FontSizeCS:
		{
			this.New.FontSizeCS = oChange.New;
			break;
		}
		case AscDFH.historyitem_ParaRun_Ligatures:
		{
			this.New.Ligatures = oChange.New;
			break
		}
	}

	return true;
};
CChangesRunTextPr.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunUnifill(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunUnifill.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunUnifill.prototype.constructor = CChangesRunUnifill;
CChangesRunUnifill.prototype.Type = AscDFH.historyitem_ParaRun_Unifill;
CChangesRunUnifill.prototype.private_CreateObject = function()
{
	return new AscFormat.CUniFill();
};
CChangesRunUnifill.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Unifill = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunUnifill.prototype.Load = function(Color)
{
	this.Redo();

	var Unifill = this.Class.Pr.Unifill;
	if (AscCommon.CollaborativeEditing
		&& Unifill
		&& Unifill.fill
		&& Unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_BLIP
		&& typeof Unifill.fill.RasterImageId === "string"
		&& Unifill.fill.RasterImageId.length > 0)
	{
		AscCommon.CollaborativeEditing.Add_NewImage(Unifill.fill.RasterImageId);
	}

	if (this.Color && Color)
		this.Class.setCollPrChangeColor(Color);
};
CChangesRunUnifill.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunUnifill.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunShd(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunShd.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunShd.prototype.constructor = CChangesRunShd;
CChangesRunShd.prototype.Type = AscDFH.historyitem_ParaRun_Shd;
CChangesRunShd.prototype.private_CreateObject = function()
{
	return new CDocumentShd();
};
CChangesRunShd.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.Shd = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunShd.prototype.Load = function(Color)
{
	this.Redo();

	var Unifill = this.Class.Pr.Unifill;
	if (AscCommon.CollaborativeEditing
		&& Unifill
		&& Unifill.fill
		&& Unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_BLIP
		&& typeof Unifill.fill.RasterImageId === "string"
		&& Unifill.fill.RasterImageId.length > 0)
	{
		AscCommon.CollaborativeEditing.Add_NewImage(Unifill.fill.RasterImageId);
	}

	if (this.Color && Color)
		this.Class.setCollPrChangeColor(Color);
};
CChangesRunShd.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunShd.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunMathStyle(Class, Old, New)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
}
CChangesRunMathStyle.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunMathStyle.prototype.constructor = CChangesRunMathStyle;
CChangesRunMathStyle.prototype.Type = AscDFH.historyitem_ParaRun_MathStyle;
CChangesRunMathStyle.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.MathPrp.sty = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunMathStyle.prototype.Merge = function(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_MathPrp === oChange.Type)
		return false;

	return true;
};
CChangesRunMathStyle.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunMathPrp(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunMathPrp.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunMathPrp.prototype.constructor = CChangesRunMathPrp;
CChangesRunMathPrp.prototype.Type = AscDFH.historyitem_ParaRun_MathPrp;
CChangesRunMathPrp.prototype.private_CreateObject = function()
{
	return new CMPrp();
};
CChangesRunMathPrp.prototype.private_IsCreateEmptyObject = function()
{
	return true;
};
CChangesRunMathPrp.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.MathPrp = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunMathPrp.prototype.Merge = function(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (this.Type === oChange.Type)
		return false;

	if (!this.New)
		this.New = new CMPrp();

	if (AscDFH.historyitem_ParaRun_MathStyle === oChange.Type)
	{
		this.New.sty = oChange.New;
	}
	else if (AscDFH.historyitem_ParaRun_MathAlnAt === oChange.Type)
	{
		if (undefined !== this.New.brk)
			this.New.brk.Apply_AlnAt(oChange.New);
	}
	else if (AscDFH.historyitem_ParaRun_MathForcedBreak === oChange.Type)
	{
		if(oChange.bInsert)
			this.New.Insert_ForcedBreak(oChange.alnAt);
		else
			this.New.Delete_ForcedBreak();
	}


	return true;
};
CChangesRunMathPrp.prototype.CheckLock = private_ParagraphContentChangesCheckLock;

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunPrChange(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunPrChange.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunPrChange.prototype.constructor = CChangesRunPrChange;
CChangesRunPrChange.prototype.Type = AscDFH.historyitem_ParaRun_PrChange;
CChangesRunPrChange.prototype.WriteToBinary = function(Writer)
{
	// Long : Flags
	// 1-bit : is New.PrChange undefined ?
	// 2-bit : is New.ReviewInfo undefined ?
	// 3-bit : is Old.PrChange undefined ?
	// 4-bit : is Old.ReviewInfo undefined ?
	// Variable(CTextPr)     : New.PrChange   (1bit = 0)
	// Variable(AscWord.ReviewInfo) : New.ReviewInfo (2bit = 0)
	// Variable(CTextPr)     : Old.PrChange   (3bit = 0)
	// Variable(AscWord.ReviewInfo) : Old.ReviewInfo (4bit = 0)
	var nFlags = 0;
	if (undefined === this.New.PrChange)
		nFlags |= 1;

	if (undefined === this.New.ReviewInfo)
		nFlags |= 2;

	if (undefined === this.Old.PrChange)
		nFlags |= 4;

	if (undefined === this.Old.ReviewInfo)
		nFlags |= 8;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New.PrChange)
		this.New.PrChange.Write_ToBinary(Writer);

	if (undefined !== this.New.ReviewInfo)
		this.New.ReviewInfo.Write_ToBinary(Writer);

	if (undefined !== this.Old.PrChange)
		this.Old.PrChange.Write_ToBinary(Writer);

	if (undefined !== this.Old.ReviewInfo)
		this.Old.ReviewInfo.Write_ToBinary(Writer);
};
CChangesRunPrChange.prototype.ReadFromBinary = function(Reader)
{
	// Long : Flags
	// 1-bit : is New.PrChange undefined ?
	// 2-bit : is New.ReviewInfo undefined ?
	// 3-bit : is Old.PrChange undefined ?
	// 4-bit : is Old.ReviewInfo undefined ?
	// Variable(CTextPr)     : New.PrChange   (1bit = 0)
	// Variable(AscWord.ReviewInfo) : New.ReviewInfo (2bit = 0)
	// Variable(CTextPr)     : Old.PrChange   (3bit = 0)
	// Variable(AscWord.ReviewInfo) : Old.ReviewInfo (4bit = 0)
	var nFlags = Reader.GetLong();

	this.New = {
		PrChange   : undefined,
		ReviewInfo : undefined
	};

	this.Old = {
		PrChange   : undefined,
		ReviewInfo : undefined
	};

	if (nFlags & 1)
	{
		this.New.PrChange = undefined;
	}
	else
	{
		this.New.PrChange = new CTextPr();
		this.New.PrChange.Read_FromBinary(Reader);
	}

	if (nFlags & 2)
	{
		this.New.ReviewInfo = undefined;
	}
	else
	{
		this.New.ReviewInfo = new AscWord.ReviewInfo();
		this.New.ReviewInfo.Read_FromBinary(Reader);
	}

	if (nFlags & 4)
	{
		this.Old.PrChange = undefined;
	}
	else
	{
		this.Old.PrChange = new CTextPr();
		this.Old.PrChange.Read_FromBinary(Reader);
	}

	if (nFlags & 8)
	{
		this.Old.ReviewInfo = undefined;
	}
	else
	{
		this.Old.ReviewInfo = new AscWord.ReviewInfo();
		this.Old.ReviewInfo.Read_FromBinary(Reader);
	}
};
CChangesRunPrChange.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;

	oRun.Pr.PrChange   = Value.PrChange;
	oRun.Pr.ReviewInfo = Value.ReviewInfo;
	oRun.updateTrackRevisions();
};
CChangesRunPrChange.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_TextPr === oChange.Type)
		return false;

	if (AscDFH.historyitem_ParaRun_PrReviewInfo === oChange.Type)
		this.New.ReviewInfo = oChange.New;

	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunTextFill(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunTextFill.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunTextFill.prototype.constructor = CChangesRunTextFill;
CChangesRunTextFill.prototype.Type = AscDFH.historyitem_ParaRun_TextFill;
CChangesRunTextFill.prototype.private_CreateObject = function()
{
	return new AscFormat.CUniFill();
};
CChangesRunTextFill.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.TextFill = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunTextFill.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunTextFill.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunTextFill.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunTextOutline(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunTextOutline.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunTextOutline.prototype.constructor = CChangesRunTextOutline;
CChangesRunTextOutline.prototype.Type = AscDFH.historyitem_ParaRun_TextOutline;
CChangesRunTextOutline.prototype.private_CreateObject = function()
{
	return new AscFormat.CLn();
};
CChangesRunTextOutline.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.TextOutline = Value;

	oRun.Recalc_CompiledPr(false);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunTextOutline.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunTextOutline.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunTextOutline.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunPrReviewInfo(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunPrReviewInfo.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunPrReviewInfo.prototype.constructor = CChangesRunPrReviewInfo;
CChangesRunPrReviewInfo.prototype.Type = AscDFH.historyitem_ParaRun_PrReviewInfo;
CChangesRunPrReviewInfo.prototype.private_CreateObject = function()
{
	return new AscWord.ReviewInfo();
};
CChangesRunPrReviewInfo.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.ReviewInfo = Value;
};
CChangesRunPrReviewInfo.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_TextPr === oChange.Type || AscDFH.historyitem_ParaRun_PrChange === oChange.Type)
		return false;

	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseObjectProperty}
 */
function CChangesRunContentReviewInfo(Class, Old, New, Color)
{
	AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New, Color);
}
CChangesRunContentReviewInfo.prototype = Object.create(AscDFH.CChangesBaseObjectProperty.prototype);
CChangesRunContentReviewInfo.prototype.constructor = CChangesRunContentReviewInfo;
CChangesRunContentReviewInfo.prototype.Type = AscDFH.historyitem_ParaRun_ContentReviewInfo;
CChangesRunContentReviewInfo.prototype.private_CreateObject = function()
{
	return new AscWord.ReviewInfo();
};
CChangesRunContentReviewInfo.prototype.private_IsCreateEmptyObject = function()
{
	return false;
};
CChangesRunContentReviewInfo.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.ReviewInfo = Value;
};
CChangesRunContentReviewInfo.prototype.Merge = function(oChange)
{
	if (oChange.Class !== this.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_ReviewType === oChange.Type)
		return false;

	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesRunOnStartSplit(Class, Pos)
{
	AscDFH.CChangesBase.call(this, Class);
	this.Pos = Pos;
}
CChangesRunOnStartSplit.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesRunOnStartSplit.prototype.constructor = CChangesRunOnStartSplit;
CChangesRunOnStartSplit.prototype.Type = AscDFH.historyitem_ParaRun_OnStartSplit;
CChangesRunOnStartSplit.prototype.Undo = function()
{
};
CChangesRunOnStartSplit.prototype.Redo = function()
{
};
CChangesRunOnStartSplit.prototype.WriteToBinary = function(Writer)
{
	Writer.WriteLong(this.Pos);
};
CChangesRunOnStartSplit.prototype.ReadFromBinary = function(Reader)
{
	this.Pos = Reader.GetLong();
};
CChangesRunOnStartSplit.prototype.Load = function()
{
	if (AscCommon.CollaborativeEditing)
		AscCommon.CollaborativeEditing.OnStart_SplitRun(this.Class, this.Pos);
};
CChangesRunOnStartSplit.prototype.CreateReverseChange = function()
{
	return null;
};
CChangesRunOnStartSplit.prototype.Merge = function(oChange)
{
	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesRunOnEndSplit(Class, NewRun)
{
	AscDFH.CChangesBase.call(this, Class);
	this.NewRun = NewRun;
}
CChangesRunOnEndSplit.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesRunOnEndSplit.prototype.constructor = CChangesRunOnEndSplit;
CChangesRunOnEndSplit.prototype.Type = AscDFH.historyitem_ParaRun_OnEndSplit;
CChangesRunOnEndSplit.prototype.Undo = function()
{
};
CChangesRunOnEndSplit.prototype.Redo = function()
{
};
CChangesRunOnEndSplit.prototype.WriteToBinary = function(Writer)
{
	Writer.WriteString2(this.NewRun.Get_Id());
};
CChangesRunOnEndSplit.prototype.ReadFromBinary = function(Reader)
{
	var RunId = Reader.GetString2();
	this.NewRun = g_oTableId.Get_ById(RunId);
};
CChangesRunOnEndSplit.prototype.Load = function()
{
	if (AscCommon.CollaborativeEditing)
		AscCommon.CollaborativeEditing.OnEnd_SplitRun(this.NewRun);
};
CChangesRunOnEndSplit.prototype.CreateReverseChange = function()
{
	return null;
};
CChangesRunOnEndSplit.prototype.Merge = function(oChange)
{
	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunMathAlnAt(Class, Old, New)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesRunMathAlnAt.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunMathAlnAt.prototype.constructor = CChangesRunMathAlnAt;
CChangesRunMathAlnAt.prototype.Type = AscDFH.historyitem_ParaRun_MathAlnAt;
CChangesRunMathAlnAt.prototype.private_SetValue = function(Value)
{
	this.Class.MathPrp.Apply_AlnAt(Value);
};
CChangesRunMathAlnAt.prototype.WriteToBinary = function(Writer)
{
	// Long : Flags
	// 1-bit : New is undefined ?
	// 2-bit : Old is undefined ?
	// 1bit = 0 long : this.New
	// 2bit = 0 long : this.Old

	var nFlags = 0;
	if (undefined === this.New)
		nFlags |= 1;

	if (undefined === this.Old)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteLong(this.New);

	if (undefined !== this.Old)
		Writer.WriteLong(this.Old);
};
CChangesRunMathAlnAt.prototype.ReadFromBinary = function(Reader)
{
	// Long : Flags
	// 1-bit : New is undefined ?
	// 2-bit : Old is undefined ?
	// 1bit = 0 long : this.New
	// 2bit = 0 long : this.Old

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.New = undefined;
	else
		this.New = Reader.GetLong();

	if (nFlags & 2)
		this.Old = undefined;
	else
		this.Old = Reader.GetLong();
};
CChangesRunMathAlnAt.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_MathPrp === oChange.Type || AscDFH.historyitem_ParaRun_MathForcedBreak === oChange.Type)
		return false;

	return true;
};
CChangesRunMathAlnAt.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesRunMathForcedBreak(Class, bInsert, alnAt)
{
	AscDFH.CChangesBase.call(this, Class);

	this.bInsert = bInsert;
	this.alnAt   = alnAt;
}
CChangesRunMathForcedBreak.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesRunMathForcedBreak.prototype.constructor = CChangesRunMathForcedBreak;
CChangesRunMathForcedBreak.prototype.Type = AscDFH.historyitem_ParaRun_MathForcedBreak;
CChangesRunMathForcedBreak.prototype.Undo = function()
{
	var oRun = this.Class;

	if(this.bInsert)
		oRun.MathPrp.Delete_ForcedBreak();
	else
		oRun.MathPrp.Insert_ForcedBreak(this.alnAt);
};
CChangesRunMathForcedBreak.prototype.Redo = function()
{
	var oRun = this.Class;

	if(this.bInsert)
		oRun.MathPrp.Insert_ForcedBreak(this.alnAt);
	else
		oRun.MathPrp.Delete_ForcedBreak();
};
CChangesRunMathForcedBreak.prototype.WriteToBinary = function(Writer)
{
	// Long : Flags
	// 1-bit : bInsert ?
	// 2-bit : alnAt is undefined ?
	// 2bit = 0 long : this.alnAt

	var nFlags = 0;
	if (true === this.bInsert)
		nFlags |= 1;

	if (undefined === this.alnAt)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (undefined !== this.alnAt)
		Writer.WriteLong(this.alnAt);
};
CChangesRunMathForcedBreak.prototype.ReadFromBinary = function(Reader)
{
	// Long : Flags
	// 1-bit : bInsert ?
	// 2-bit : alnAt is undefined ?
	// 2bit = 0 long : this.alnAt

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.bInsert = true;
	else
		this.bInsert = false;

	if (nFlags & 2)
		this.alnAt = undefined;
	else
		this.alnAt = Reader.GetLong();
};
CChangesRunMathForcedBreak.prototype.CreateReverseChange = function()
{
	return new CChangesRunMathForcedBreak(this.Class, !this.bInsert, this.alnAt);
};
CChangesRunMathForcedBreak.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type || AscDFH.historyitem_ParaRun_MathPrp === oChange.Type)
		return false;

	if (AscDFH.historyitem_ParaRun_MathAlnAt === oChange.Type)
		this.alnAt = oChange.New;

	return true;
};
CChangesRunMathForcedBreak.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * Базовый класс для изменений ссылок на шрифт из темы
 * @constructor
 */
function CChangesRunRFontsThemeBase()
{}
CChangesRunRFontsThemeBase.prototype.WriteToBinary = function(oWriter)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	oWriter.WriteLong(nFlags);

	if (undefined !== this.New)
		oWriter.WriteString2(this.New);

	if (undefined !== this.Old)
		oWriter.WriteString2(this.Old);
};
CChangesRunRFontsThemeBase.prototype.ReadFromBinary = function(oReader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : Is undefined New ?
	// 3-bit : Is undefined Old ?
	// String : New
	// String : Old

	var nFlags = oReader.GetLong();
	this.Color = !!(nFlags & 1);
	this.New   = !!(nFlags & 2) ? undefined : oReader.GetString2();
	this.Old   = !!(nFlags & 4) ? undefined : oReader.GetString2();
};
CChangesRunRFontsThemeBase.prototype.private_SetRFontsValue = function(sValue)
{
};
CChangesRunRFontsThemeBase.prototype.private_SetValue = function(sValue)
{
	var oRun = this.Class;
	this.private_SetRFontsValue(sValue);
	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRFontsThemeBase.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRFontsThemeBase.prototype.Merge = private_ParaRunChangesOnMergeRFontsTextPr;
CChangesRunRFontsThemeBase.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsAsciiTheme(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsAsciiTheme.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsAsciiTheme.prototype.constructor = CChangesRunRFontsAsciiTheme;
AscCommon.ExtendPrototype(CChangesRunRFontsAsciiTheme, CChangesRunRFontsThemeBase);
CChangesRunRFontsAsciiTheme.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_Ascii_Theme;
CChangesRunRFontsAsciiTheme.prototype.private_SetRFontsValue = function(sValue)
{
	this.Class.Pr.RFonts.AsciiTheme = sValue;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsHAnsiTheme(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsHAnsiTheme.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsHAnsiTheme.prototype.constructor = CChangesRunRFontsHAnsiTheme;
AscCommon.ExtendPrototype(CChangesRunRFontsHAnsiTheme, CChangesRunRFontsThemeBase);
CChangesRunRFontsHAnsiTheme.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_HAnsi_Theme;
CChangesRunRFontsHAnsiTheme.prototype.private_SetRFontsValue = function(sValue)
{
	this.Class.Pr.RFonts.HAnsiTheme = sValue;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsCSTheme(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsCSTheme.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsCSTheme.prototype.constructor = CChangesRunRFontsCSTheme;
CChangesRunRFontsCSTheme.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_CS_Theme;
AscCommon.ExtendPrototype(CChangesRunRFontsCSTheme, CChangesRunRFontsThemeBase);
CChangesRunRFontsCSTheme.prototype.private_SetRFontsValue = function(sValue)
{
	this.Class.Pr.RFonts.CSTheme = sValue;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunRFontsEastAsiaTheme(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesRunRFontsEastAsiaTheme.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunRFontsEastAsiaTheme.prototype.constructor = CChangesRunRFontsEastAsiaTheme;
CChangesRunRFontsEastAsiaTheme.prototype.Type = AscDFH.historyitem_ParaRun_RFonts_EastAsia_Theme;
AscCommon.ExtendPrototype(CChangesRunRFontsEastAsiaTheme, CChangesRunRFontsThemeBase);
CChangesRunRFontsEastAsiaTheme.prototype.private_SetRFontsValue = function(sValue)
{
	this.Class.Pr.RFonts.EastAsiaTheme = sValue;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunBoldCS(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunBoldCS.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunBoldCS.prototype.constructor = CChangesRunBoldCS;
CChangesRunBoldCS.prototype.Type = AscDFH.historyitem_ParaRun_BoldCS;
CChangesRunBoldCS.prototype.private_SetValue = function(Value)
{
	let oRun = this.Class;
	oRun.Pr.BoldCS = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunBoldCS.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunBoldCS.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunBoldCS.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunItalicCS(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunItalicCS.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunItalicCS.prototype.constructor = CChangesRunItalicCS;
CChangesRunItalicCS.prototype.Type = AscDFH.historyitem_ParaRun_ItalicCS;
CChangesRunItalicCS.prototype.private_SetValue = function(Value)
{
	var oRun = this.Class;
	oRun.Pr.ItalicCS = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunItalicCS.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunItalicCS.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunItalicCS.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseDoubleProperty}
 */
function CChangesRunFontSizeCS(Class, Old, New, Color)
{
	AscDFH.CChangesBaseDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesRunFontSizeCS.prototype = Object.create(AscDFH.CChangesBaseDoubleProperty.prototype);
CChangesRunFontSizeCS.prototype.constructor = CChangesRunFontSizeCS;
CChangesRunFontSizeCS.prototype.Type = AscDFH.historyitem_ParaRun_FontSizeCS;
CChangesRunFontSizeCS.prototype.private_SetValue = function(Value)
{
	let oRun = this.Class;
	oRun.Pr.FontSizeCS = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunFontSizeCS.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunFontSizeCS.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunFontSizeCS.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesRunLigatures(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesRunLigatures.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesRunLigatures.prototype.constructor = CChangesRunLigatures;
CChangesRunLigatures.prototype.Type = AscDFH.historyitem_ParaRun_Ligatures;
CChangesRunLigatures.prototype.private_SetValue = function(Value)
{
	let oRun = this.Class;
	oRun.Pr.Ligatures = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateShapeText();
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunLigatures.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunLigatures.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunCS(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunCS.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunCS.prototype.constructor = CChangesRunCS;
CChangesRunCS.prototype.Type = AscDFH.historyitem_ParaRun_CS;
CChangesRunCS.prototype.private_SetValue = function(Value)
{
	let oRun = this.Class;
	oRun.Pr.CS = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunCS.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunCS.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunCS.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesRunRTL(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesRunRTL.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesRunRTL.prototype.constructor = CChangesRunRTL;
CChangesRunRTL.prototype.Type = AscDFH.historyitem_ParaRun_RTL;
CChangesRunRTL.prototype.private_SetValue = function(Value)
{
	let oRun = this.Class;
	oRun.Pr.RTL = Value;

	oRun.Recalc_CompiledPr(true);
	oRun.private_UpdateTrackRevisionOnChangeTextPr(false);
};
CChangesRunRTL.prototype.Load = private_ParaRunChangesLoadTextPr;
CChangesRunRTL.prototype.Merge = private_ParaRunChangesOnMergeTextPr;
CChangesRunRTL.prototype.CheckLock = private_ParagraphContentChangesCheckLock;

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesRunMathMetaData(Class, Old, New)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesRunMathMetaData.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesRunMathMetaData.prototype.constructor = CChangesRunMathMetaData;
CChangesRunMathMetaData.prototype.Type = AscDFH.historyitem_ParaRun_MathMetaData;
CChangesRunMathMetaData.prototype.Undo = function()
{
	var oRun = this.Class;
	oRun.math_autocorrection = this.Old;
};
CChangesRunMathMetaData.prototype.Redo = function()
{
	var oRun = this.Class;
	oRun.math_autocorrection = this.New;
};
CChangesRunMathMetaData.prototype.WriteToBinary = function(Writer)
{
	if (this.New)
	{
		Writer.WriteBool(false);
		this.New.Write_ToBinary(Writer);
	}
	else
	{
		Writer.WriteBool(true);
	}

	if (this.Old)
	{
		Writer.WriteBool(false);
		this.Old.Write_ToBinary(Writer);
	}
	else
	{
		Writer.WriteBool(true);
	}
};
CChangesRunMathMetaData.prototype.ReadFromBinary = function(Reader)
{
	if (!Reader.GetBool())
	{
		let oMetaData = new AscMath.MathMetaData();
		oMetaData.Read_FromBinary(Reader);
		this.New = oMetaData;
	}

	if (!Reader.GetBool())
	{
		let oOldMetaData = new AscMath.MathMetaData();
		oOldMetaData.Read_FromBinary(Reader);
		this.Old = oOldMetaData;
	}
};
CChangesRunMathMetaData.prototype.Load = function(Color)
{
	this.Redo();
};
CChangesRunMathMetaData.prototype.IsRelated = function(oChanges)
{
};
CChangesRunMathMetaData.prototype.CreateReverseChange = function()
{
	return new CChangesRunMathMetaData(this.Class, this.New, this.Old);
};
CChangesRunMathMetaData.prototype.CheckLock = private_ParagraphContentChangesCheckLock;
