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
	};
	MacroRecorder.prototype.stop = function()
	{
		this.inProgress = false;
		this.paused = false;

		let macroData = "";
		try
		{
			macroData = JSON.parse(this.editor.macros.GetData());
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
	};
	MacroRecorder.prototype.pause = function()
	{
		if (!this.inProgress || this.paused)
			return;
		
		this.paused = true;
	};
	MacroRecorder.prototype.resume = function()
	{
		if (!this.inProgress || !this.paused)
			return;
		
		this.paused = false;
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
		
		if (AscDFH.historydescription_Document_SetTextBold === type
			|| AscDFH.historydescription_Document_SetTextBoldHotKey === type)
		{
			this.result += "Api.GetDocument().GetRangeBySelect().SetBold(true);\n";
		}
		else if (AscDFH.historydescription_Document_SetTextItalic === type
			|| AscDFH.historydescription_Document_SetTextItalicHotKey === type)
		{
			this.result += "Api.GetDocument().GetRangeBySelect().SetItalic(true);\n";
		}
		else if (type === AscDFH.historydescription_Document_AddLetter && additional && additional.codePoints)
		{
			let text = "";
			for (let i = 0; i < additional.codePoints.length; ++i)
				text += String.fromCodePoint(additional.codePoints[i]);
			this.result += "Api.GetDocument().GetCurrentParagraph().AddText(\"" + text + "\");\n";
		}
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
	//--------------------------------------------------------export----------------------------------------------------
	AscCommon.MacroRecorder = MacroRecorder;
	
	MacroRecorder.prototype["start"]        = MacroRecorder.prototype.start;
	MacroRecorder.prototype["stop"]         = MacroRecorder.prototype.stop;
	MacroRecorder.prototype["pause"]        = MacroRecorder.prototype.pause;
	MacroRecorder.prototype["resume"]       = MacroRecorder.prototype.resume;
	MacroRecorder.prototype["isInProgress"] = MacroRecorder.prototype.isInProgress;
	MacroRecorder.prototype["isPaused"]     = MacroRecorder.prototype.isPaused;
	
})(window);
