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
	const MAX_ACTION_TIME = 20;
	
	/**
	 * CustomTextAnnotator manages external text annotation workflows for document paragraphs.
	 *
	 * This class handles the process of sending paragraph text to external applications/plugins
	 * for analysis and receives back highlight positions that can be rendered and made interactive.
	 * It maintains state for paragraphs at different stages of the annotation pipeline.
	 
	 * @param {AscWord.Document} logicDocument
	 * @constructor
	 */
	function CustomTextAnnotator(logicDocument)
	{
		this.logicDocument = logicDocument;
		
		this.waitingParagraphs  = {};
		this.paragraphs         = {};
		this.checkingParagraphs = {};
		
		this.textGetter = new ParagraphText();
		
		
		this.eventManager = this.logicDocument.GetApi().getTextAnnotatorEventManager();
	}
	
	CustomTextAnnotator.prototype.isActive = function()
	{
		return true;
	};
	CustomTextAnnotator.prototype.addParagraphToCheck = function(para)
	{
		this.checkingParagraphs[para.GetId()] = para;
	};
	CustomTextAnnotator.prototype.continueProcessing = function()
	{
		if (!this.isActive())
			return;
		
		let startTime = performance.now();
		while (true)
		{
			if (performance.now() - startTime > MAX_ACTION_TIME)
				break;
			
			let paragraph = this.popNextParagraph();
			if (!paragraph)
				break;
			
			this.handleParagraph(paragraph);
		}
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CustomTextAnnotator.prototype.popNextParagraph = function()
	{
		for (let paraId in this.checkingParagraphs)
		{
			let paragraph = this.checkingParagraphs[paraId];
			delete this.checkingParagraphs[paraId];
			
			if (!paragraph.IsUseInDocument())
				continue;
			
			return paragraph;
		}
		
		return null;
	};
	CustomTextAnnotator.prototype.handleParagraph = function(paragraph)
	{
		this.textGetter.check(paragraph);
		console.log(`ParaId=${paragraph.GetId()}; ParaText=${this.textGetter.text}`);
		
		this.eventManager.send({
			"paragraphId" : paragraph.GetId(),
			"text"        : this.textGetter.text
		});
	};
	CustomTextAnnotator.prototype.highlightTextResponse = function(handlerId, paraId, ranges)
	{
		let _ranges = [];
		ranges.forEach(r => _ranges.push([r.start, r.length, r.id]))
		console.log(`Response from handlerId=${handlerId} ParaId=${paraId}; Ranges=${_ranges}`);
	};
	
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
	AscCommon.TextAnnotatorEventManager = TextAnnotatorEventManager;
	/**
	 *
	 * @constructor
	 */
	function ParagraphText()
	{
		AscWord.DocumentVisitor.call(this);
		this.text = "";
	}
	ParagraphText.prototype = Object.create(AscWord.DocumentVisitor.prototype);
	ParagraphText.prototype.constructor = ParagraphText;
	ParagraphText.prototype.check = function(paragraph)
	{
		this.text = "";
		this.traverseParagraph(paragraph);
	};
	ParagraphText.prototype.run = function(run)
	{
		this.text += run.GetText();
		return true;
	};
	//-------------------------------------------------------------export-----------------------------------------------
	AscWord.CustomTextAnnotator = CustomTextAnnotator;
})(window);
