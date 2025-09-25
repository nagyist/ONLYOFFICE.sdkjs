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

(function()
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
	};
	
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
})();
