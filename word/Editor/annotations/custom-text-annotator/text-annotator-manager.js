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

(function(window)
{
	/**
	 * @param editor
	 * @constructor
	 */
	function TextAnnotatorEventManager(editor)
	{
		this.editor = editor;
		
		this.logicDocument = null;
		this.textAnnotator = null;
		
		this.paragraphs = {};
	}
	TextAnnotatorEventManager.prototype.init = function()
	{
		if (this.logicDocument)
			return;
		
		this.logicDocument = this.editor.private_GetLogicDocument();
		this.textAnnotator = this.logicDocument.CustomTextAnnotator;
	};
	TextAnnotatorEventManager.prototype.send = function(obj)
	{
		this.init();
		
		let recalcId = this.logicDocument.GetRecalcId();
		obj["recalcId"] = recalcId;
		this.paragraphs[obj["paragraphId"]] = recalcId;
		
		let text = obj["text"];
		let len  = text.length;
		
		let _t = this;
		setTimeout(function(){
			let _start = Math.floor(Math.random() * (len - 1));
			let _len   = Math.min(Math.floor(Math.random() * 10), len - _start);
			_t.onResponse({
				"guid" : "guid-1",
				"type" : "highlightText",
				"paragraphId" : obj["paragraphId"],
				"recalcId" : recalcId,
				"ranges" : [{
					"start" : _start,
					"length" : _len,
					"id" : "1"
				}]
			});
			
			_start = Math.floor(Math.random() * (len - 1));
			_len   = Math.min(Math.floor(Math.random() * 10), len - _start);
			_t.onResponse({
				"guid" : "guid-2",
				"type" : "highlightText",
				"paragraphId" : obj["paragraphId"],
				"recalcId" : recalcId,
				"ranges" : [{
					"start" : _start,
					"length" : _len,
					"id" : "1"
				}]
			});
		}, 2000);
		//window.g_asc_plugins.onPluginEvent("onParagraphText", obj);
		
		// TODO: Чтобы не было моргания при быстром изменении параграфа, мы не должны чистить метки сразу при изменении
		//       Поэтому, до получения ответа мы оставляем метки в прежних местах. Далее либо обновляем их с ответом,
		//       либо на определенном таймере чистим их (если ответ не приходит)
	};
	TextAnnotatorEventManager.prototype.onResponse = function(obj)
	{
		this.init();
		
		if (!obj)
			return;
		
		switch (obj["type"])
		{
			case "highlightText":
			{
				let guid     = obj["guid"];
				let paraId   = obj["paragraphId"];
				let recalcId = obj["recalcId"];
				let ranges   = obj["ranges"];
				
				if (undefined === guid
					|| this.paragraphs[paraId] !== recalcId
					|| !ranges
					|| !Array.isArray(ranges))
					return;
				
				let _ranges = [];
				for (let i = 0; i < ranges.length; ++i)
				{
					let _r = this.parseHighlightTextRange(ranges[i]);
					if (_r)
						_ranges.push(_r);
				}
				
				this.logicDocument.CustomTextAnnotator.highlightTextResponse(guid, paraId, _ranges);
				break;
			}
		}
	};
	TextAnnotatorEventManager.prototype.parseHighlightTextRange = function(obj)
	{
		if (!obj || undefined === obj["start"] || undefined === obj["length"] || undefined === obj["id"])
			return null;
		
		return {
			start : obj["start"],
			length : obj["length"],
			id : obj["id"]
		};
	};
	//-------------------------------------------------------------export-----------------------------------------------
	AscCommon.TextAnnotatorEventManager = TextAnnotatorEventManager;
})(window);
