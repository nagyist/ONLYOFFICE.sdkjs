/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

(function (window)
{
	/**
	 * @param editor
	 * @constructor
	 */
	function MacroRecorder(editor)
	{
		this.editor = editor;
		this.inProgress = false;
		this.paused = false;
		this.macroName = "";
		this.result = "";
	}
	
	MacroRecorder.prototype.start = function(macroName)
	{
		if (this.inProgress)
			return;
		
		this.macroName = macroName;
		this.result = "";
		this.paused = false;
		this.inProgress = true;
		
		this.editor.sendEvent("asc_onMacroRecordingStart");
	};
	MacroRecorder.prototype.stop = function()
	{
		if (!this.inProgress)
			return;
		
		this.inProgress = false;
		this.paused = false;

		let macroData = "";
		try
		{
			let data = this.editor.macros.GetData();
			if (data && "" !== data)
			{
				macroData = JSON.parse(this.editor.macros.GetData());
			}
			else
			{
				macroData = {
					macrosArray : [],
					current     : -1
				};
			}
		}
		catch (e)
		{
			return;
		}
		
		let name = this.macroName ? this.macroName : this.getNewName(macroData.macrosArray);
		let value = "(function()\n{\n" + this.result + "})();"
		macroData.macrosArray.push({
			guid : AscCommon.CreateUUID(true),
			name : name,
			autostart : false,
			value : value
		});
		
		this.editor.asc_setMacros(JSON.stringify(macroData));
		this.editor.sendEvent("asc_onMacroRecordingStop");
	};
	MacroRecorder.prototype.cancel = function()
	{
		if (!this.inProgress)
			return;
		
		this.inProgress = false;
		this.paused = false;
		this.editor.sendEvent("asc_onMacroRecordingStop");
	};
	MacroRecorder.prototype.pause = function()
	{
		if (!this.inProgress || this.paused)
			return;
		
		this.paused = true;
		this.editor.sendEvent("asc_onMacroRecordingPause");
	};
	MacroRecorder.prototype.resume = function()
	{
		if (!this.inProgress || !this.paused)
			return;
		
		this.paused = false;
		this.editor.sendEvent("asc_onMacroRecordingResume");
	};
	MacroRecorder.prototype.isInProgress = function()
	{
		return this.inProgress;
	};
	MacroRecorder.prototype.isPaused = function()
	{
		return this.paused;
	};
	MacroRecorder.prototype.onAction = function(type, additional)
	{
		if (!this.isInProgress() || this.isPaused())
			return;

		let actionsMacros = null;
		
		if (this.editor.editorId === AscCommon.c_oEditorId.Word)
			actionsMacros = WordActionsMacroList;
		else if (this.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			actionsMacros = CellActionsMacroList;
		else if (this.editor.editorId === AscCommon.c_oEditorId.Presentation)
			actionsMacros = PresentationActionMacroList;

		if (!actionsMacros)
			return;

		let actionMacroFunction = actionsMacros[type];
		if (actionMacroFunction)
			this.result += actionMacroFunction(additional, type);
	};
	MacroRecorder.prototype.getNewName = function(macros)
	{
		let maxId = 0;
		for (let i = 0, count = macros.length; i < count; ++i)
		{
			if (0 !== macros[i].name.indexOf("Macro "))
				continue;
			
			let curId = parseInt(macros[i].name.substr(6));
			if (isNaN(curId))
				continue;
			
			maxId = Math.max(curId, maxId);
		}
		
		return "Macro " + (maxId + 1);
	};

	const wordActions = {
		setTextBold				: function(){return "Api.GetDocument().GetRangeBySelect().SetBold(true);\n"},
		setTextItalic			: function(){return "Api.GetDocument().GetRangeBySelect().SetItalic(true);\n"},
		setTextUnderline		: function(){return "Api.GetDocument().GetRangeBySelect().SetUnderline(true);\n"},
		setTextStrikeout		: function(){return "Api.GetDocument().GetRangeBySelect().SetStrikeout(true);\n"},
		setTextFontName			: function(additional){ return (additional && additional.fontName) ? "Api.GetDocument().GetRangeBySelect().SetFontFamily(\"" + additional.fontName + "\");\n" : ""},
		setTextFontSize			: function(additional){ return (additional && additional.fontSize) ? "Api.GetDocument().GetRangeBySelect().SetFontSize(\"" + additional.fontSize + "\");\n" : ""},
		setTextHighlightColor	: function(additional){
			if (!(additional && additional.highlight))
				return "";
	
			let color = new CDocumentColor(highlight.r, highlight.g, highlight.b);
			let highlightColor = color.ToHighlightColor();
	
			if (highlightColor === "")
				highlightColor = 'none';
	
			return "Api.GetDocument().GetRangeBySelect().SetHighlight(\"" + highlightColor + "\");\n";
		},
		setTextHighlightNone	: function(){return "Api.GetDocument().GetRangeBySelect().SetHighlight(\"none\");\n"},
		setTextVertAlign		: function(additional, type)
		{
			if (!(additional && additional.baseline))
				return "";
	
			if (additional.baseline === true)
				return "Api.GetDocument().GetRangeBySelect().SetVertAlign(\"baseline\");\n";
			else if (AscDFH.historydescription_Document_SetTextVertAlignHotKey3 === type)
				return "Api.GetDocument().GetRangeBySelect().SetVertAlign(\"subscript\");\n";
			else if (AscDFH.historydescription_Document_SetTextVertAlignHotKey2 === type)
				return "Api.GetDocument().GetRangeBySelect().SetVertAlign(\"superscript\");\n";
		},
		setTextColor			: function(additional){ return (additional && additional.color) ? "Api.GetDocument().GetRangeBySelect().SetColor(\"" + additional.color.r + "," + additional.color.g + "," + additional.color.b  + "\");\n" : ""},
		setStyleHeading			: function(additional){ return (additional && additional.name) ? "Api.GetDocument().GetRangeBySelect().SetStyle(\"" + additional.name + "\");\n" : ""},
		clearFormat				: function(){return "Api.GetDocument().GetRangeBySelect().ClearFormating()\n"},
		cut						: function(){return "Api.GetDocument().GetRangeBySelect().Cut();\n"},
		changeTextCase			: function(additional){ return (additional && additional.changeType) ? "Api.GetDocument().GetRangeBySelect().SetTextCase(\"" + additional.changeType + "\");\n" : ""},
		incFontSize				: function(){ return "Api.GetDocument().GetRangeBySelect().Grow();\n"},
		addLetter				: function(additional)
		{
			if (!(additional && additional.codePoints))
				return "";
	
			let text = "";
			for (let i = 0; i < additional.codePoints.length; ++i)
				text += String.fromCodePoint(additional.codePoints[i]);
	
			return "Api.GetDocument().GetCurrentParagraph().AddText(\"" + text + "\");\n";
		}
	}

	const WordActionsMacroList = {
		[AscDFH.historydescription_Document_SetTextBold]					: wordActions.setTextBold,
		[AscDFH.historydescription_Document_SetTextBoldHotKey]				: wordActions.setTextBold,
		[AscDFH.historydescription_Document_SetTextItalic]					: wordActions.setTextItalic,
		[AscDFH.historydescription_Document_SetTextItalicHotKey]			: wordActions.setTextItalic,
		[AscDFH.historydescription_Document_SetTextUnderline]				: wordActions.setTextUnderline,
		[AscDFH.historydescription_Document_SetTextUnderlineHotKey]			: wordActions.setTextUnderline,
		[AscDFH.historydescription_Document_SetTextStrikeout]				: wordActions.setTextStrikeout,
		[AscDFH.historydescription_Document_SetTextStrikeoutHotKey]			: wordActions.setTextStrikeout,
		[AscDFH.historydescription_Document_SetTextFontName]				: wordActions.setTextFontName,
		[AscDFH.historydescription_Document_SetTextFontSize]				: wordActions.setTextFontSize,
		[AscDFH.historydescription_Document_SetTextHighlightColor]			: wordActions.setTextHighlightColor,
		[AscDFH.historydescription_Document_SetTextHighlightNone]			: wordActions.setTextHighlightNone,
		[AscDFH.historydescription_Document_SetTextVertAlignHotKey2]		: wordActions.setTextVertAlign,
		[AscDFH.historydescription_Document_SetTextVertAlignHotKey3]		: wordActions.setTextVertAlign,
		[AscDFH.historydescription_Document_SetTextColor]					: wordActions.setTextColor,
		[AscDFH.historydescription_Document_SetStyleHeading]				: wordActions.setStyleHeading,
		[AscDFH.historydescription_Document_Shortcut_ClearFormatting]		: wordActions.clearFormat,
		[AscDFH.historydescription_Document_ClearFormatting]				: wordActions.clearFormat,
		[AscDFH.historydescription_Cut]										: wordActions.cut,
		[AscDFH.historydescription_Document_ChangeTextCase]					: wordActions.changeTextCase,
		[AscDFH.historydescription_Document_AddLetter]						: wordActions.addLetter,
		//[AscDFH.historydescription_Document_IncFontSize]					: wordActions.incFontSize,
	};

	const cellActions = {
		setCellIncreaseFontSize	: function(){return "Api.GetSelection().FontIncrease();\n";},
		setCellDecreaseFontSize	: function(){return "Api.GetSelection().FontDecrease();\n";},
		setCellFontSize			: function(additional){ return (additional && additional.val) ? "Api.GetSelection().SetFontSize(\"" + additional.val + "\");\n" : "";},
		setCellFontName			: function(additional){ return (additional && additional.val) ? "Api.GetSelection().SetFontName(\"" + additional.val + "\");\n" : "";},
		setCellBold				: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().SetBold(" + additional.val + ");\n" : "";},
		setCellItalic			: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().SetItalic(" + additional.val + ");\n" : "";},
		setCellUnderline		: function(additional){
			if (!(additional && additional.val !== undefined))
				return "";

			let underlineType = null;

			switch (additional.val) {
				case Asc.EUnderline.underlineSingle:				underlineType = 'single';				break;
				case Asc.EUnderline.underlineSingleAccounting:		underlineType = 'singleAccounting';	break;
				case Asc.EUnderline.underlineDouble:				underlineType = 'double';				break;
				case Asc.EUnderline.underlineDoubleAccounting:		underlineType = 'doubleAccounting';	break;
				case Asc.EUnderline.underlineNone:
				default:											underlineType = 'none';					break;
			}

			return "Api.GetSelection().SetUnderline(\"" + underlineType + "\");\n";
		},
		setCellStrikeout		: function(additional){
			if (!(additional && additional.val !== undefined))
				return "";

			return "Api.GetSelection().SetStrikeout(" + (!!additional.val) + ");\n";
		},
		setCellSubscript		: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().GetCharacters().GetFont().SetSubscript(" + additional.val + ");\n" : "";},
		setCellSuperscript		: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().GetCharacters().GetFont().SetSuperscript(" + additional.val + ");\n" : "";},
		setCellReadingOrder		: function(additional){
			if (!(additional && additional.val !== undefined))
				return "";

			let direction = null;

			switch (additional.val) {
				case 0:		direction = 'context';	break;
				case 1:		direction = 'ltr';		break;
				case 2:		direction = 'rtl';		break;
				default:	return "";
			}

			return "Api.GetSelection().SetReadingOrder(\"" + direction + "\");\n";
		},
		setCellAlign			: function(additional){
			if ((additional && additional.val === undefined))
				return "";
	
			let align = null;
	
			switch (additional.val) {
				case AscCommon.align_Left:		align = 'left';		break;
				case AscCommon.align_Right:		align = 'right';	break;
				case AscCommon.align_Justify:	align = 'justify';	break;
				case AscCommon.align_Center:	align = 'center';	break;
				default:						return "";
			}
	
			return "Api.GetSelection().SetAlignHorizontal(\"" + align + "\");\n";
		},
		setCellVerticalAlign	: function(additional){
			if ((additional && additional.val === undefined))
				return "";
	
			let align = null;
	
			switch (additional.val) {
				case Asc.c_oAscVAlign.Center:	align = 'center';		break;
				case Asc.c_oAscVAlign.Bottom:	align = 'bottom';		break;
				case Asc.c_oAscVAlign.Top:		align = 'top';			break;
				case Asc.c_oAscVAlign.Dist:		align = 'distributed';	break;
				case Asc.c_oAscVAlign.Just:		align = 'justify';		break;
				default:						return "";
			}
	
			return "Api.GetSelection().SetAlignVertical(\"" + align + "\");\n";
		},
		setCellTextColor		: function(additional){
			if (!additional && !additional.val)
				return "";

			let color = "Api.CreateColorFromRGB(" + additional.val.getR() + ", " + additional.val.getG() + ", " + additional.val.getB() + ")";
			return "Api.GetSelection().SetFontColor(" + color + ");\n"
		},
		setCellBackgroundColor	: function(additional){
			if (!additional && !additional.val)
				return "";

			let color = "Api.CreateColorFromRGB(" + additional.val.getR() + ", " + additional.val.getG() + ", " + additional.val.getB() + ")";
			return "Api.GetSelection().SetBackgroundColor(" + color + ");\n"
		},
		setCellWrap				: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().SetWrap(" + additional.val + ");\n" : "";},
		//setCellShrinkToFit		: function(additional){ return (additional && additional.val !== undefined) ? "Api.GetSelection().SetShrinkToFit(" + additional.val + ");\n" : "";},
		setCellValue			: function(additional){ 
			if (!(additional && additional.val !== undefined))
				return "";
			
			let value = additional.val;
			if (typeof value === 'string') {
				value = '"' + value.replace(/"/g, '\\"') + '"';
			} else {
				value = value.toString();
			}
			
			return "Api.GetSelection().SetValue(" + value + ");\n";
		},
		setCellAngle			: function(additional){ 
			if (!(additional && additional.val !== undefined))
				return "";
			
			let angle = additional.val;
			
			switch (angle) {
				case -90:	return "Api.GetSelection().SetOrientation('xlDownward');\n";
				case 0:		return "Api.GetSelection().SetOrientation('xlHorizontal');\n";
				case 90:	return "Api.GetSelection().SetOrientation('xlUpward');\n";
				case 255:	return "Api.GetSelection().SetOrientation('xlVertical');\n";
			}
		},
		setCellChangeTextCase	: function(additional){ 
			if (!(additional && additional.val !== undefined))
				return "";
			
			return "Api.GetSelection().ChangeTextCase(" + additional.val + ");\n";
		},
		setCellChangeFontSize	: function(additional){ 
			if (!(additional && additional.val !== undefined))
				return "";
			
			return additional.val ? "Api.asc_increaseFontSize();\n" : "Api.asc_decreaseFontSize();\n";
		},
		setCellBorder			: function(additional){
			if (!(additional && additional.val !== undefined)) {
				return "";
			}
				
			let borderArray = additional.val;
			if (!Array.isArray(borderArray) || borderArray.length === 0) {
				return "";
			}
			
			let result = "";
			
			for (let i = 0; i < borderArray.length; i++) {
				let border = borderArray[i];
				if (border && border.style !== undefined) {
					
					let positionStr = null;
					switch (i) {
						case 0: positionStr = 'Top'; break;
						case 1: positionStr = 'Right'; break;
						case 2: positionStr = 'Bottom'; break;
						case 3: positionStr = 'Left'; break;
						case 4: positionStr = 'DiagonalDown'; break;
						case 5: positionStr = 'DiagonalUp'; break;
						case 6: positionStr = 'InsideVertical'; break;
						case 7: positionStr = 'InsideHorizontal'; break;
						default: continue;
					}
					
					let styleStr = null;
					switch (border.style) {
						case window['Asc'].c_oAscBorderStyles.None: styleStr = 'None'; break;
						case window['Asc'].c_oAscBorderStyles.Double: styleStr = 'Double'; break;
						case window['Asc'].c_oAscBorderStyles.Hair: styleStr = 'Hair'; break;
						case window['Asc'].c_oAscBorderStyles.DashDotDot: styleStr = 'DashDotDot'; break;
						case window['Asc'].c_oAscBorderStyles.DashDot: styleStr = 'DashDot'; break;
						case window['Asc'].c_oAscBorderStyles.Dotted: styleStr = 'Dotted'; break;
						case window['Asc'].c_oAscBorderStyles.Dashed: styleStr = 'Dashed'; break;
						case window['Asc'].c_oAscBorderStyles.Thin: styleStr = 'Thin'; break;
						case window['Asc'].c_oAscBorderStyles.MediumDashDotDot: styleStr = 'MediumDashDotDot'; break;
						case window['Asc'].c_oAscBorderStyles.SlantDashDot: styleStr = 'SlantDashDot'; break;
						case window['Asc'].c_oAscBorderStyles.MediumDashDot: styleStr = 'MediumDashDot'; break;
						case window['Asc'].c_oAscBorderStyles.MediumDashed: styleStr = 'MediumDashed'; break;
						case window['Asc'].c_oAscBorderStyles.Medium: styleStr = 'Medium'; break;
						case window['Asc'].c_oAscBorderStyles.Thick: styleStr = 'Thick'; break;
						default: continue;
					}
					
					let colorStr = "Api.CreateColorFromRGB(0, 0, 0)";
					if (border.color) {
						if (typeof border.color === 'string') {
							let hex = border.color.replace('#', '');
							if (hex.length === 3) {
								hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
							}
							let r = parseInt(hex.substr(0, 2), 16) || 0;
							let g = parseInt(hex.substr(2, 2), 16) || 0;
							let b = parseInt(hex.substr(4, 2), 16) || 0;
							colorStr = "Api.CreateColorFromRGB(" + r + ", " + g + ", " + b + ")";
						} else if (typeof border.color === 'object') {
							colorStr = "Api.CreateColorFromRGB(" + (border.color.r || 0) + ", " + (border.color.g || 0) + ", " + (border.color.b || 0) + ")";
						}
					}
					
					result += "Api.GetSelection().SetBorders(\"" + positionStr + "\", \"" + styleStr + "\", " + colorStr + ");\n";
				}
			}
			
			return result;
		},
		setCellHyperlinkAdd		: function(additional) {return (additional && additional.url) ? "" : ""},
		setCellHyperlinkModify	: function(additional) {return (additional && additional.url) ? "" : ""},
		setCellHyperlinkRemove	: function(additional) {return (additional && additional.url) ? "" : ""},
		cut						: function(){return "ApiApi.GetSelection().Cut();\n"},
		cellChangeValue			: function(additional){
			if (!additional || !additional.data)
				return "";
			return "Api.GetSelection().SetValue(\"" + additional.data + "\");\n";
		},
		
	};

	const CellActionsMacroList = {
		//[AscDFH.historydescription_Spreadsheet_SetCellIncreaseFontSize]	: cellActions.setCellIncreaseFontSize,
		//[AscDFH.historydescription_Spreadsheet_SetCellDecreaseFontSize]	: cellActions.setCellDecreaseFontSize,
		[AscDFH.historydescription_Spreadsheet_SetCellFontSize]				: cellActions.setCellFontSize,
		[AscDFH.historydescription_Spreadsheet_SetCellFontName]				: cellActions.setCellFontName,
		[AscDFH.historydescription_Spreadsheet_SetCellBold]					: cellActions.setCellBold,
		[AscDFH.historydescription_Spreadsheet_SetCellItalic]				: cellActions.setCellItalic,
		[AscDFH.historydescription_Spreadsheet_SetCellUnderline]			: cellActions.setCellUnderline,
		[AscDFH.historydescription_Spreadsheet_SetCellStrikeout]			: cellActions.setCellStrikeout,
		[AscDFH.historydescription_Spreadsheet_SetCellSubscript]			: cellActions.setCellSubscript,
		[AscDFH.historydescription_Spreadsheet_SetCellSuperscript]			: cellActions.setCellSuperscript,
		[AscDFH.historydescription_Spreadsheet_SetCellReadingOrder]			: cellActions.setCellReadingOrder,
		[AscDFH.historydescription_Spreadsheet_SetCellAlign]				: cellActions.setCellAlign,
		[AscDFH.historydescription_Spreadsheet_SetCellVertAlign]			: cellActions.setCellVerticalAlign,
		[AscDFH.historydescription_Spreadsheet_SetCellTextColor]			: cellActions.setCellTextColor,
		[AscDFH.historydescription_Spreadsheet_SetCellBackgroundColor]	    : cellActions.setCellBackgroundColor,
		[AscDFH.historydescription_Spreadsheet_SetCellWrap]				    : cellActions.setCellWrap,
		//[AscDFH.historydescription_Spreadsheet_SetCellShrinkToFit]			: cellActions.setCellShrinkToFit,
		[AscDFH.historydescription_Spreadsheet_SetCellBorder]				: cellActions.setCellBorder,
		[AscDFH.historydescription_Spreadsheet_SetCellValue]				: cellActions.setCellValue,
		[AscDFH.historydescription_Spreadsheet_SetCellAngle]				: cellActions.setCellAngle,
		//[AscDFH.historydescription_Spreadsheet_SetCellMerge]				: cellActions.setCellMerge,
		//[AscDFH.historydescription_Spreadsheet_SetCellStyle]				: cellActions.setCellStyle,
		[AscDFH.historydescription_Spreadsheet_SetCellChangeTextCase]		: cellActions.setCellChangeTextCase,
		[AscDFH.historydescription_Spreadsheet_SetCellChangeFontSize]		: cellActions.setCellChangeFontSize
		//[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkAdd]		: cellActions.setCellHyperlinkAdd,
		//[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkModify]	: cellActions.setCellHyperlinkModify,
		//[AscDFH.historydescription_Spreadsheet_SetCellHyperlinkRemove]	: cellActions.setCellHyperlinkRemove,
		//[AscDFH.historydescription_Cut]										: cellActions.cut,
	};



	const PresentationActionMacroList = {

	};

	//--------------------------------------------------------export----------------------------------------------------
	AscCommon.MacroRecorder = MacroRecorder;
	
	MacroRecorder.prototype["start"]        = MacroRecorder.prototype.start;
	MacroRecorder.prototype["stop"]         = MacroRecorder.prototype.stop;
	MacroRecorder.prototype["cancel"]       = MacroRecorder.prototype.cancel;
	MacroRecorder.prototype["pause"]        = MacroRecorder.prototype.pause;
	MacroRecorder.prototype["resume"]       = MacroRecorder.prototype.resume;
	MacroRecorder.prototype["isInProgress"] = MacroRecorder.prototype.isInProgress;
	MacroRecorder.prototype["isPaused"]     = MacroRecorder.prototype.isPaused;
	
})(window);
