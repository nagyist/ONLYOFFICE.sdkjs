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
	const EVENT_TIMEOUT = 1000;
	
	/**
	 * @param editor
	 * @constructor
	 */
	function TextAnnotatorEventManager(editor)
	{
		this.editor = editor;
		
		this.logicDocument = null;
		this.textAnnotator = null;
		
		this.sendPara = {};
		this.waitPara = {};
		
		this.curParagraph = null;
		this.curRanges    = null;
	}
	TextAnnotatorEventManager.prototype.init = function()
	{
		if (this.logicDocument)
			return;
		
		this.logicDocument = this.editor.private_GetLogicDocument();
		this.textAnnotator = this.logicDocument.CustomTextAnnotator;
	};
	/**
	 * @param {AscWord.Paragraph} paragraph
	 */
	TextAnnotatorEventManager.prototype.onChangeParagraph = function(paragraph)
	{
		this.init();
		
		let paraId = paragraph.GetId();
		
		this.sendPara[paraId] = this.logicDocument.GetRecalcId();
		
		// Не посылаем сразу сообщение, чтобы не посылать их на каждое действие при быстром наборе
		if (this.waitPara[paraId])
			clearTimeout(this.waitPara[paraId]);
		
		let _t = this;
		this.waitPara[paraId] = setTimeout(function(){
			_t.send(paragraph);
			_t.waitPara[paraId] = null;
		}, EVENT_TIMEOUT);
	};
	TextAnnotatorEventManager.prototype.send = function(paragraph)
	{
		let obj = this.textAnnotator.getEventObject(paragraph);
		let paraId = paragraph.GetId();
		let recalcId = this.sendPara[paraId];
		obj["recalcId"] = recalcId;
		
		let text = obj["text"];
		let len  = text.length;
		
		//console.log(`Request ParaId=${paragraph.GetId()}; ParaText=${text}`);
		
		let _t = this;
		// setTimeout(function(){
		// 	let _start = Math.floor(Math.random() * (len - 1));
		// 	let _len   = Math.min(Math.floor(Math.random() * 10), len - _start);
		// 	_t.onResponse({
		// 		"guid" : "guid-1",
		// 		"type" : "highlightText",
		// 		"paragraphId" : paraId,
		// 		"recalcId" : recalcId,
		// 		"ranges" : [{
		// 			"start" : _start,
		// 			"length" : _len,
		// 			"id" : "1"
		// 		}]
		// 	});
		// }, 2000);
		// setTimeout(function() {
		// 	let _start = Math.floor(Math.random() * (len - 1));
		// 	let _len   = Math.min(Math.floor(Math.random() * 10), len - _start);
		// 	_t.onResponse({
		// 		"guid"        : "guid-2",
		// 		"type"        : "highlightText",
		// 		"paragraphId" : paraId,
		// 		"recalcId"    : recalcId,
		// 		"ranges"      : [{
		// 			"start"  : _start,
		// 			"length" : _len,
		// 			"id"     : "1"
		// 		}]
		// 	});
		// }, 3000);
		
		window.g_asc_plugins.onPluginEvent("onAnnotateText", obj);
		
		// TODO: Чтобы не было моргания при быстром изменении параграфа, мы не должны чистить метки сразу при изменении
		//       Поэтому, до получения ответа мы оставляем метки в прежних местах. Далее либо обновляем их с ответом,
		//       либо на определенном таймере чистим их (если ответ не приходит)
	};
	TextAnnotatorEventManager.prototype.onResponse = function(obj)
	{
		this.init();
		
		if (!obj)
			return;
		
		let name = obj["name"];
		if (name)
			obj["guid"] += "AnnotationName:" + name;
		
		switch (obj["type"])
		{
			case "highlightText":
			{
				let guid     = obj["guid"];
				let paraId   = obj["paragraphId"];
				let recalcId = obj["recalcId"];
				let ranges   = obj["ranges"];
				
				if (undefined === guid
					|| this.sendPara[paraId] !== recalcId
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
	TextAnnotatorEventManager.prototype.selectRange = function(obj)
	{
		if (!obj || !obj["paragraphId"] || !obj["guid"] || !obj["rangeId"])
			return;
		
		let handlerId = this.getHandlerId(obj);
		this.textAnnotator.getMarks().selectRange(obj["paragraphId"], handlerId, obj["rangeId"]);
	};
	TextAnnotatorEventManager.prototype.removeRange = function(obj)
	{
		if (!obj || !obj["paragraphId"] || !obj["guid"] || !obj["rangeId"])
			return;
		
		let handlerId = this.getHandlerId(obj);
		this.textAnnotator.getMarks().removeRange(obj["paragraphId"], handlerId, obj["rangeId"]);
	};
	TextAnnotatorEventManager.prototype.onCurrentRanges = function(paragraph, ranges)
	{
		let prevRanges = this.curRanges;
		let prevPara   = this.curParagraph;
		
		let currPara   = paragraph;
		let currRanges = ranges;
		
		let changePara = currPara !== prevPara;
		
		for (let handlerId in prevRanges)
		{
			let noHandler = !currRanges[handlerId];
			for (let rangeId in prevRanges[handlerId])
			{
				if (changePara || noHandler || !currRanges[handlerId][rangeId])
				{
					let obj = {
						"paragraphId" : prevPara.GetId(),
						"rangeId"     : rangeId
					};
					this.addNameFromHandlerId(handlerId, obj);
					window.g_asc_plugins.onPluginEvent("onBlurAnnotation", obj, this.getGuid(handlerId));
				}
			}
		}
		
		for (let handlerId in currRanges)
		{
			let noHandler = !prevRanges[handlerId];
			for (let rangeId in currRanges[handlerId])
			{
				if (changePara || noHandler || !prevRanges[handlerId][rangeId])
				{
					let obj = {
						"paragraphId" : currPara.GetId(),
						"rangeId"     : rangeId
					};
					this.addNameFromHandlerId(handlerId, obj);
					window.g_asc_plugins.onPluginEvent("onFocusAnnotation", obj, this.getGuid(handlerId));
				}
			}
		}
		
		this.curParagraph = currPara;
		this.curRanges    = currRanges;
	};
	TextAnnotatorEventManager.prototype.onClick = function(paragraph, ranges)
	{
		let paraId = paragraph.GetId();
		for (let handlerId in ranges)
		{
			let _ranges = [];
			let obj = {
				"paragraphId" : paraId,
				"ranges"      : _ranges
			};
			this.addNameFromHandlerId(handlerId, obj);
			for (let rangeId in ranges[handlerId])
			{
				_ranges.push(rangeId);
			}
			window.g_asc_plugins.onPluginEvent("onClickAnnotation", obj, this.getGuid(handlerId));
		}
	};
	TextAnnotatorEventManager.prototype.getHandlerId = function(obj)
	{
		return (obj["name"] ? obj["guid"] + "AnnotationName:" + obj["name"] : obj["guid"]);
	};
	TextAnnotatorEventManager.prototype.addNameFromHandlerId = function(handlerId, obj)
	{
		if (!obj)
			return;
		
		let pos = handlerId.indexOf("AnnotationName:");
		if (-1 !== pos)
			obj["name"] = handlerId.substr(pos + 15);
	};
	TextAnnotatorEventManager.prototype.getGuid = function(handlerId)
	{
		let pos = handlerId.indexOf("AnnotationName:");
		if (-1 !== pos)
			return handlerId.substr(0, pos);
		else
			return handlerId;
	};
	//-------------------------------------------------------------export-----------------------------------------------
	AscCommon.TextAnnotatorEventManager = TextAnnotatorEventManager;
})(window);
