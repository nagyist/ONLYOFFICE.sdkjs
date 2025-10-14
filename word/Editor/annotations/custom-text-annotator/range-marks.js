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
	 * Class for handling a collection of custom annotation range marks
	 * @constructor
	 */
	function CustomMarks()
	{
		this.marks = {};
	}
	
	CustomMarks.prototype.addMark = function(mark)
	{
		let run = mark.getRun();
		if (!run)
			return;
		
		let runId = run.GetId();
		if (!this.marks[runId])
			this.marks[runId] = {};
		
		let handlerId = mark.getHandlerId();
		let markId    = mark.getMarkId();
		
		if (!this.marks[runId][handlerId])
			this.marks[runId][handlerId] = {};
		
		if (!this.marks[runId][handlerId][markId])
			this.marks[runId][handlerId][markId] = {start : null, end : null};
		
		if (markId.isStart())
			this.marks[runId][handlerId][markId].start = mark;
		else
			this.marks[runId][handlerId][markId].end = mark;
	};
	CustomMarks.prototype.clearMarks = function(handlerId)
	{
		for (let runId in this.marks)
		{
			delete this.marks[runId][handlerId];
		}
	};
	/**
	 * @param {AscWord.Paragraph} paragraph
	 * @param {AscWord.CParagraphContentPos} paraContentPos
	 * @returns {[]}
	 */
	CustomMarks.prototype.getStartedMarks = function(paragraph, paraContentPos)
	{
		let result = [];
		for (let runId in this.marks)
		{
			let run = AscCommon.g_oTableId.GetById(runId);
			if (!run || run.GetParagraph() !== paragraph)
				continue;
			
			for (let handlerId in this.marks[runId])
			{
				for (let markId in this.marks[runId][handlerId])
				{
					let startMark = this.marks[runId][handlerId][markId].start;
					let endMark   = this.marks[runId][handlerId][markId].end;
					if (!startMark || !endMark)
						continue;
					
					let startPos = startMark.getParaPos();
					let endPos   = endMark.getParaPos();
					if (!startPos || !endPos)
						continue;
					
					if (paraContentPos.Compare(startPos) >= 0 && paraContentPos.Compare(endPos) <= 0)
						result.push([handlerId, markId]);
				}
			}
		}
		return result;
	};
	CustomMarks.prototype.onAddToRun = function(runId, pos)
	{
		this.forEachInRun(runId, function(mark){
			mark.onAdd(pos);
		});
	};
	CustomMarks.prototype.onRemoveFromRun = function(runId, pos, count)
	{
		this.forEachInRun(runId, function(mark){
			mark.onRemove(pos, count);
		});
	};
	CustomMarks.prototype.onSplitRun = function(runId, pos, nextRunId)
	{
		let nextRun = AscCommon.g_oTableId.GetById(nextRunId);
		this.forEachInRun(runId, function(mark){
			mark.onSplit(pos, nextRun);
		});
	};
	CustomMarks.prototype.forEachInRun = function(runId, f)
	{
		if (!this.marks[runId])
			return;
		
		for (let handlerId in this.marks[runId])
		{
			for (let markId in this.marks[runId][handlerId])
			{
				let start = this.marks[runId][handlerId][markId].start;
				let end = this.marks[runId][handlerId][markId].end;
				if (start)
					f.call(start);
				
				if (end)
					f.call(end);
			}
		}
	};
	
	/**
	 * @param run {AscWord.Run}
	 * @param pos {number}
	 * @param handlerId {string}
	 * @param markId {string}
	 * @constructor
	 */
	function CustomMark(run, pos, handlerId, markId)
	{
		this.run       = run;
		this.pos       = 0;
		this.handlerId = null;
		this.markId    = null;
	}
	CustomMark.prototype.getMarkId = function()
	{
		return this.markId;
	};
	CustomMark.prototype.getHandlerId = function()
	{
		return this.handlerId;
	};
	CustomMark.prototype.isStart = function()
	{
		return true;
	};
	CustomMark.prototype.onAdd = function(pos)
	{
		if (this.pos >= pos)
			++this.pos;
	};
	CustomMark.prototype.onRemove = function(pos, count)
	{
		if (this.pos > pos + count)
			this.pos -= count;
		else if (this.pos > pos)
			this.pos = Math.max(0, pos);
	};
	CustomMark.prototype.onSplit = function(pos, nextRun)
	{
		if (this.pos < pos)
			return;
		
		this.pos -= pos;
		this.run = nextRun;
	};
	CustomMark.prototype.getPos = function()
	{
		return this.pos;
	};
	CustomMark.prototype.getRun = function()
	{
		return this.run;
	};
	CustomMark.prototype.getParagraph = function()
	{
		return this.run ? this.run.GetParagraph() : null;
	};
	CustomMark.prototype.getParaPos = function()
	{
		let paragraph = this.getParagraph();
		let paraPos   = paragraph ? paragraph.GetPosByElement(this.run) : null;
		if (!paraPos)
			return new AscWord.CParagraphContentPos();
		
		paraPos.Update(this.pos, paraPos.GetDepth() + 1);
		return paraPos;
	};
	
	/**
	 * Метка начала промежутка
	 * @constructor
	 */
	function CustomMarkStart()
	{
		CustomMark.call(this, arguments);
	}
	CustomMarkStart.prototype = Object.create(CustomMark.prototype);
	CustomMarkStart.prototype.constructor = CustomMarkStart;
	CustomMarkStart.prototype.isStart = function()
	{
		return true;
	};
	
	/**
	 * Метка окончания промежутка
	 * @constructor
	 */
	/**
	 * Метка начала промежутка
	 * @constructor
	 */
	function CustomMarkEnd()
	{
		CustomMark.call(this, arguments);
	}
	CustomMarkEnd.prototype = Object.create(CustomMark.prototype);
	CustomMarkEnd.prototype.constructor = CustomMarkEnd;
	CustomMarkEnd.prototype.isStart = function()
	{
		return false;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CustomMarks       = CustomMarks;
	AscWord.CustomMarkStart   = CustomMarkStart;
	AscWord.CustomMarkEnd     = CustomMarkEnd;
	
})(window);
