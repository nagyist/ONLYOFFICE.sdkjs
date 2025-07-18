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
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

// Import
	var g_oTableId = AscCommon.g_oTableId;
	var History = AscCommon.History;

	AscDFH.changesFactory[AscDFH.historyitem_Comment_Position] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Comment_Change] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_Comment_TypeInfo] = AscDFH.CChangesDrawingsLong;


	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_Comment_Position] = AscFormat.CDrawingBaseCoordsWritable;
	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_Comment_Change] = CCommentData;


	AscDFH.drawingsChangesMap[AscDFH.historyitem_Comment_Position] = function (oClass, value) {
		oClass.x = value.a;
		oClass.y = value.b;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Comment_Change] = function (oClass, value) {
		oClass.Data = value;
		if (value) {
			editor.sync_ChangeCommentData(oClass.Id, value);
		}
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Comment_TypeInfo] = function (oClass, value) {
		oClass.m_oTypeInfo = value;
	};

	function ParaComment(Start, Id) {
		this.Id = AscCommon.g_oIdCounter.Get_NewId();

		this.Paragraph = null;

		this.Start = Start;
		this.CommentId = Id;

		this.Type = para_Comment;

		this.StartLine = 0;
		this.StartRange = 0;

		this.Lines = [];
		this.LinesLength = 0;
	}

	ParaComment.prototype.Get_Id = function () {
		return this.Id;
	};

	ParaComment.prototype.SetCommentId = function (NewCommentId) {
	};

	ParaComment.prototype.Is_Empty = function () {
		return true;
	};

	ParaComment.prototype.Is_CheckingNearestPos = function () {
		return false;
	};

	ParaComment.prototype.Get_CompiledTextPr = function () {
		return null;
	};

	ParaComment.prototype.Clear_TextPr = function () {

	};

	ParaComment.prototype.Remove = function () {
		return false;
	};

	ParaComment.prototype.Get_DrawingObjectRun = function (Id) {
		return null;
	};

	ParaComment.prototype.GetRunByElement = function (oRunElement) {
		return null;
	};

	ParaComment.prototype.Get_DrawingObjectContentPos = function (Id, ContentPos, Depth) {
		return false;
	};

	ParaComment.prototype.Get_Layout = function (DrawingLayout, UseContentPos, ContentPos, Depth) {
	};

	ParaComment.prototype.GetNextRunElements = function (RunElements, UseContentPos, Depth) {
	};

	ParaComment.prototype.GetPrevRunElements = function (RunElements, UseContentPos, Depth) {
	};

	ParaComment.prototype.CollectDocumentStatistics = function (ParaStats) {
	};

	ParaComment.prototype.Create_FontMap = function (Map) {
	};

	ParaComment.prototype.Get_AllFontNames = function (AllFonts) {
	};

	ParaComment.prototype.GetSelectedText = function (bAll, bClearText) {
		return "";
	};

	ParaComment.prototype.GetSelectDirection = function () {
		return 1;
	};

	ParaComment.prototype.Clear_TextFormatting = function (DefHyper) {
	};

	ParaComment.prototype.CanAddDropCap = function () {
		return null;
	};

	ParaComment.prototype.CheckSelectionForDropCap = function (isUsePos, oEndPos, nDepth) {
		return true;
	};

	ParaComment.prototype.Get_TextForDropCap = function (DropCapText, UseContentPos, ContentPos, Depth) {
	};

	ParaComment.prototype.Get_StartTabsCount = function (TabsCounter) {
		return true;
	};

	ParaComment.prototype.Remove_StartTabs = function (TabsCounter) {
		return true;
	};

	ParaComment.prototype.Copy = function (Selected) {
		return new ParaComment(this.Start, this.CommentId);
	};

	ParaComment.prototype.CopyContent = function (Selected) {
		return [];
	};

	ParaComment.prototype.Split = function () {
		return new ParaRun();
	};

	ParaComment.prototype.Apply_TextPr = function () {
	};

	ParaComment.prototype.CheckRevisionsChanges = function (Checker, ContentPos, Depth) {
	};

	ParaComment.prototype.Get_ParaPosByContentPos = function (ContentPos, Depth) {
		return new CParaPos(this.StartRange, this.StartLine, 0, 0);
	};
//-----------------------------------------------------------------------------------
// Функции пересчета
//-----------------------------------------------------------------------------------

	ParaComment.prototype.Recalculate_Reset = function (StartRange, StartLine) {
		this.StartLine = StartLine;
		this.StartRange = StartRange;
	};

	ParaComment.prototype.Recalculate_Range = function (PRS, ParaPr) {
	};

	ParaComment.prototype.Recalculate_Set_RangeEndPos = function (PRS, PRP, Depth) {
	};

	ParaComment.prototype.Recalculate_LineMetrics = function (PRS, ParaPr, _CurLine, _CurRange) {
	};

	ParaComment.prototype.Recalculate_Range_Width = function (PRSC, _CurLine, _CurRange) {
	};

	ParaComment.prototype.Recalculate_Range_Spaces = function (PRSA, CurLine, CurRange, CurPage) {
	};

	ParaComment.prototype.Recalculate_PageEndInfo = function (PRSI, _CurLine, _CurRange) {
	};

	ParaComment.prototype.RecalculateEndInfo = function (PRSI) {

	};

	ParaComment.prototype.SaveRecalculateObject = function (Copy) {
	};

	ParaComment.prototype.LoadRecalculateObject = function (RecalcObj, Parent) {
	};

	ParaComment.prototype.PrepareRecalculateObject = function () {
	};

	ParaComment.prototype.IsEmptyRange = function (_CurLine, _CurRange) {
		return true;
	};

	ParaComment.prototype.Check_Range_OnlyMath = function (Checker, CurRange, CurLine) {
	};

	ParaComment.prototype.Check_MathPara = function (Checker) {
	};

	ParaComment.prototype.Check_PageBreak = function () {
		return false;
	};

	ParaComment.prototype.CheckSplitPageOnPageBreak = function (oPBChecker) {
		return false;
	};

	ParaComment.prototype.RecalculateMinMaxContentWidth = function () {

	};

	ParaComment.prototype.Get_Range_VisibleWidth = function (RangeW, _CurLine, _CurRange) {
	};

	ParaComment.prototype.Shift_Range = function (Dx, Dy, _CurLine, _CurRange, _CurPage) {
	};
//-----------------------------------------------------------------------------------
// Функции отрисовки
//-----------------------------------------------------------------------------------
	ParaComment.prototype.Draw_HighLights = function (PDSH) {
	};

	ParaComment.prototype.Draw_Elements = function (PDSE) {
	};

	ParaComment.prototype.Draw_Lines = function (PDSL) {
	};
//-----------------------------------------------------------------------------------
// Функции для работы с курсором
//-----------------------------------------------------------------------------------
	ParaComment.prototype.IsCursorPlaceable = function () {
		return false;
	};

	ParaComment.prototype.Cursor_Is_Start = function () {
		return true;
	};

	ParaComment.prototype.Cursor_Is_NeededCorrectPos = function () {
		return true;
	};

	ParaComment.prototype.Cursor_Is_End = function () {
		return true;
	};

	ParaComment.prototype.MoveCursorToStartPos = function () {
	};

	ParaComment.prototype.MoveCursorToEndPos = function (SelectFromEnd) {
	};

	ParaComment.prototype.getParagraphContentPosByXY = function (state) {
		return false;
	};

	ParaComment.prototype.Get_ParaContentPos = function (bSelection, bStart, ContentPos, bUseCorrection) {
	};

	ParaComment.prototype.Set_ParaContentPos = function (ContentPos, Depth) {
	};

	ParaComment.prototype.Get_PosByElement = function (Class, ContentPos, Depth, UseRange, Range, Line) {
		if (this === Class) return true;

		return false;
	};

	ParaComment.prototype.Get_ElementByPos = function (ContentPos, Depth) {
		return this;
	};

	ParaComment.prototype.Get_ClassesByPos = function (Classes, ContentPos, Depth) {
		Classes.push(this);
	};

	ParaComment.prototype.GetPosByDrawing = function (Id, ContentPos, Depth) {
		return false;
	};

	ParaComment.prototype.Get_RunElementByPos = function (ContentPos, Depth) {
		return null;
	};

	ParaComment.prototype.Get_LastRunInRange = function (_CurLine, _CurRange) {
		return null;
	};

	ParaComment.prototype.Get_LeftPos = function (SearchPos, ContentPos, Depth, UseContentPos) {
	};

	ParaComment.prototype.Get_RightPos = function (SearchPos, ContentPos, Depth, UseContentPos, StepEnd) {
	};

	ParaComment.prototype.Get_WordStartPos = function (SearchPos, ContentPos, Depth, UseContentPos) {
	};

	ParaComment.prototype.Get_WordEndPos = function (SearchPos, ContentPos, Depth, UseContentPos, StepEnd) {
	};

	ParaComment.prototype.Get_EndRangePos = function (_CurLine, _CurRange, SearchPos, Depth) {
		return false;
	};

	ParaComment.prototype.Get_StartRangePos = function (_CurLine, _CurRange, SearchPos, Depth) {
		return false;
	};

	ParaComment.prototype.Get_StartRangePos2 = function (_CurLine, _CurRange, ContentPos, Depth) {
	};

	ParaComment.prototype.Get_EndRangePos2 = function (_CurLine, _CurRange, ContentPos, Depth) {
	};

	ParaComment.prototype.Get_StartPos = function (ContentPos, Depth) {
	};

	ParaComment.prototype.Get_EndPos = function (BehindEnd, ContentPos, Depth) {
	};
//-----------------------------------------------------------------------------------
// Функции для работы с селектом
//-----------------------------------------------------------------------------------
	ParaComment.prototype.Set_SelectionContentPos = function (StartContentPos, EndContentPos, Depth, StartFlag, EndFlag) {
	};

	ParaComment.prototype.RemoveSelection = function () {
	};

	ParaComment.prototype.SelectAll = function (Direction) {
	};

	ParaComment.prototype.drawSelectionInRange = function (line, range, drawState) {
	};

	ParaComment.prototype.IsSelectionEmpty = function (CheckEnd) {
		return true;
	};

	ParaComment.prototype.Selection_CheckParaEnd = function () {
		return false;
	};

	ParaComment.prototype.IsSelectedAll = function (Props) {
		return true;
	};

	ParaComment.prototype.SkipAnchorsAtSelectionStart = function (nDirection) {
		return true;
	};

	ParaComment.prototype.Selection_CheckParaContentPos = function (ContentPos) {
		return true;
	};


	ParaComment.prototype.Refresh_RecalcData = function () {
	};

	ParaComment.prototype.Write_ToBinary2 = function (Writer) {
	};

	ParaComment.prototype.Read_FromBinary2 = function (Reader) {
	};
	ParaComment.prototype.SetParagraph = function (Paragraph) {
		this.Paragraph = Paragraph;
	};
	ParaComment.prototype.GetCurrentParaPos = function () {
		return new CParaPos(this.StartRange, this.StartLine, 0, 0);
	};
	ParaComment.prototype.Get_TextPr = function (ContentPos, Depth) {
		return new CTextPr();
	};
//----------------------------------------------------------------------------------------------------------------------
// Разное
//----------------------------------------------------------------------------------------------------------------------
	ParaComment.prototype.SetReviewType = function (ReviewType, RemovePrChange) {
	};
	ParaComment.prototype.SetReviewTypeWithInfo = function (ReviewType, ReviewInfo) {
	};
	ParaComment.prototype.CheckRevisionsChanges = function (Checker, ContentPos, Depth) {
	};
	ParaComment.prototype.AcceptRevisionChanges = function (Type, bAll) {
	};
	ParaComment.prototype.RejectRevisionChanges = function (Type, bAll) {
	};

	function CWriteCommentData() {
		this.Data = null; // CCommentData

		this.WriteAuthorId = 0;
		this.WriteCommentId = 0;
		this.WriteParentAuthorId = 0;
		this.WriteParentCommentId = 0;
		this.WriteTime = "";
		this.WriteText = "";

		this.AdditionalData = "";
		this.timeZoneBias = null;

		this.x = 0;
		this.y = 0;
	}

	CWriteCommentData.prototype.Calculate = function () {
		this.WriteTime = new Date(this.Data.m_sTime - 0).toISOString().slice(0, 19) + 'Z';
		this.timeZoneBias = this.Data.m_nTimeZoneBias;

		this.CalculateAdditionalData();
	};

	CWriteCommentData.prototype.Calculate2 = function () {
		var dateMs = AscCommon.getTimeISO8601(this.WriteTime);
		if (!isNaN(dateMs)) {
			this.WriteTime = dateMs + "";
		}
		else {
			this.WriteTime = "1";
		}
	};

	CWriteCommentData.prototype.CalculateAdditionalData = function () {
		if (!this.Data) {
			this.AdditionalData = "";
		}
		else {
			let sUserId = this.Data.m_sUserId;
			let sUserName = this.Data.m_sUserName;
			if (typeof sUserId === "string" && sUserId.length > 0 && typeof sUserName === "string" && sUserName.length > 0) {
				this.AdditionalData = "teamlab_data:";
				this.AdditionalData += ("0;" + sUserId.length + ";" + sUserId + ";");
				this.AdditionalData += ("1;" + sUserName.length + ";" + sUserName + ";");
				this.AdditionalData += ("2;1;" + (this.Data.m_bSolved ? "1;" : "0;"));
				if (this.Data.m_sOOTime) {
					var WriteOOTime = new Date(this.Data.m_sOOTime - 0).toISOString().slice(0, 19) + 'Z';
					this.AdditionalData += ("3;" + WriteOOTime.length + ";" + WriteOOTime + ";");
				}
				if (this.Data.m_sGuid) {
					this.AdditionalData += "4;" + this.Data.m_sGuid.length + ";" + this.Data.m_sGuid + ";";
				}
				if (this.Data.m_sUserData) {
					this.AdditionalData += "5;" + this.Data.m_sUserData.length + ";" + this.Data.m_sUserData + ";";
				}
			}
			else {
				this.AdditionalData = "";
			}
		}
	};

	CWriteCommentData.prototype.ReadNextInteger = function (_parsed) {
		var _len = _parsed.data.length;
		var _found = -1;

		var _Found = ";".charCodeAt(0);
		for (var i = _parsed.pos; i < _len; i++) {
			if (_Found == _parsed.data.charCodeAt(i)) {
				_found = i;
				break;
			}
		}

		if (-1 == _found) return -1;

		var _ret = parseInt(_parsed.data.substr(_parsed.pos, _found - _parsed.pos));
		if (isNaN(_ret)) return -1;

		_parsed.pos = _found + 1;
		return _ret;
	};

	CWriteCommentData.prototype.ParceAdditionalData = function (_comment_data) {
		if (this.AdditionalData.indexOf("teamlab_data:") != 0) return;

		var _parsed = {data: this.AdditionalData, pos: "teamlab_data:".length};

		while (true) {
			var _attr = this.ReadNextInteger(_parsed);
			if (-1 == _attr) break;

			var _len = this.ReadNextInteger(_parsed);
			if (-1 == _len) break;

			var _value = _parsed.data.substr(_parsed.pos, _len);
			_parsed.pos += (_len + 1);

			if (0 == _attr) _comment_data.m_sUserId = _value; else if (1 == _attr) _comment_data.m_sUserName = _value; else if (2 == _attr) _comment_data.m_bSolved = ("1" == _value) ? true : false; else if (3 == _attr) {
				var dateMs = AscCommon.getTimeISO8601(_value);
				if (!isNaN(dateMs)) _comment_data.m_sOOTime = dateMs + "";
			}
			else if (4 == _attr) _comment_data.m_sGuid = _value; else if (5 == _attr) _comment_data.m_sUserData = _value;
		}
	};

	function CCommentAuthor() {
		AscFormat.CBaseNoIdObject.call(this);
		this.Name = "";
		this.Id = 0;
		this.LastId = 0;
		this.Initials = "";
	}

	AscFormat.InitClass(CCommentAuthor, AscFormat.CBaseNoIdObject, 0);
	CCommentAuthor.prototype.Calculate = function () {
		var arr = this.Name.split(" ");
		this.Initials = "";
		for (var i = 0; i < arr.length; i++) {
			if (arr[i].length > 0) this.Initials += (arr[i].substring(0, 1));
		}
	};


	function CCommentData() {
		this.m_sText = "";
		this.m_sTime = "";
		this.m_sOOTime = "";
		this.m_sUserId = "";
		this.m_sUserName = "";
		this.m_sGuid = "";
		this.m_sQuoteText = null;
		this.m_bSolved = false;
		this.m_nTimeZoneBias = null;
		this.m_aReplies = [];
	}

	CCommentData.prototype.createDuplicate = function (bNewGuid) {
		var ret = new CCommentData();
		ret.m_sText = this.m_sText;
		ret.m_sTime = this.m_sTime;
		ret.m_sOOTime = this.m_sOOTime;
		ret.m_sUserId = this.m_sUserId;
		ret.m_sUserName = this.m_sUserName;
		ret.m_sGuid = bNewGuid ? AscCommon.CreateGUID() : this.m_sGuid;
		ret.m_sQuoteText = this.m_sQuoteText;
		ret.m_bSolved = this.m_bSolved;
		ret.m_nTimeZoneBias = this.m_nTimeZoneBias;
		for (var i = 0; i < this.m_aReplies.length; ++i) {
			ret.m_aReplies.push(this.m_aReplies[i].createDuplicate(bNewGuid));
		}
		return ret;
	};

	CCommentData.prototype.Add_Reply = function (CommentData) {
		this.m_aReplies.push(CommentData);
	};

	CCommentData.prototype.Set_Text = function (Text) {
		this.SetText(Text);
	};

	CCommentData.prototype.SetText = function (Text) {
		this.m_sText = Text;
	};

	CCommentData.prototype.Get_Text = function () {
		return this.m_sText;
	};

	CCommentData.prototype.Get_QuoteText = function () {
		return this.GetQuoteText();
	};

	CCommentData.prototype.GetQuoteText = function () {
		return this.m_sQuoteText;
	};


	CCommentData.prototype.Set_QuoteText = function (Quote) {
		this.SetQuoteText(Quote);
	};


	CCommentData.prototype.SetQuoteText = function (Quote) {
		this.m_sQuoteText = Quote;
	};

	CCommentData.prototype.Get_Solved = function () {
		return this.GetSolved();
	};
	CCommentData.prototype.IsSolved = CCommentData.prototype.Get_Solved;

	CCommentData.prototype.Set_Solved = function (Solved) {
		this.SetSolved(Solved);
	};

	CCommentData.prototype.GetSolved = function () {
		return this.m_bSolved;
	};

	CCommentData.prototype.SetSolved = function (isSolved) {
		this.m_bSolved = isSolved;
	};

	CCommentData.prototype.Set_Name = function (Name) {
		this.SetUserName(Name);
	};

	CCommentData.prototype.Get_Name = function () {
		return this.GetUserName();
	};

	CCommentData.prototype.SetUserName = function (Name) {
		this.m_sUserName = Name;
	};

	CCommentData.prototype.GetUserName = function () {
		return this.m_sUserName;
	};

	CCommentData.prototype.Set_Guid = function (Guid) {
		this.m_sGuid = Guid;
	};

	CCommentData.prototype.Get_Guid = function () {
		return this.m_sGuid;
	};

	CCommentData.prototype.Set_TimeZoneBias = function (timeZoneBias) {
		this.m_nTimeZoneBias = timeZoneBias;
	};

	CCommentData.prototype.Get_TimeZoneBias = function () {
		return this.m_nTimeZoneBias;
	};

	CCommentData.prototype.Get_RepliesCount = function () {
		return this.m_aReplies.length;
	};

	CCommentData.prototype.Get_Reply = function (Index) {
		return this.GetReply(Index);
	};


	CCommentData.prototype.GetReply = function (Index) {
		if (Index < 0 || Index >= this.m_aReplies.length) return null;

		return this.m_aReplies[Index];
	};

	CCommentData.prototype.GetDateTime = function () {
		var nTime = parseInt(this.m_sTime);
		if (isNaN(nTime)) nTime = 0;

		return nTime;
	};

	CCommentData.prototype.Read_FromAscCommentData = function (AscCommentData) {
		this.m_sText = AscCommentData.asc_getText();
		this.m_sTime = AscCommentData.asc_getTime();
		this.m_sOOTime = AscCommentData.asc_getOnlyOfficeTime();
		this.m_sUserId = AscCommentData.asc_getUserId();
		this.m_sQuoteText = AscCommentData.asc_getQuoteText();
		this.m_bSolved = AscCommentData.asc_getSolved();
		this.m_sUserName = AscCommentData.asc_getUserName();
		this.m_sGuid = AscCommentData.asc_getGuid();
		this.m_nTimeZoneBias = AscCommentData.asc_getTimeZoneBias();

		var RepliesCount = AscCommentData.asc_getRepliesCount();
		for (var Index = 0; Index < RepliesCount; Index++) {
			var Reply = new CCommentData();
			Reply.Read_FromAscCommentData(AscCommentData.asc_getReply(Index));
			this.m_aReplies.push(Reply);
		}
	};

	CCommentData.prototype.ConvertToSimpleObject = function()
	{
		var obj = {};

		obj["Text"]      = this.m_sText;
		obj["Time"]      = this.m_sTime;
		obj["UserName"]  = this.m_sUserName;
		obj["QuoteText"] = this.m_sQuoteText;
		obj["Solved"]    = this.m_bSolved;
		obj["UserData"]  = this.m_sUserData;
		obj["Replies"]   = [];

		for (var nIndex = 0, nCount = this.m_aReplies.length; nIndex < nCount; ++nIndex)
		{
			obj["Replies"].push(this.m_aReplies[nIndex].ConvertToSimpleObject());
		}

		return obj;
	};

	CCommentData.prototype.ReadFromSimpleObject = function (oData) {
		if (!oData) return;

		if (oData["Text"]) this.m_sText = oData["Text"];

		if (oData["Time"]) this.m_sTime = oData["Time"];

		if (oData["UserName"]) this.m_sUserName = oData["UserName"];

		if (oData["UserId"]) this.m_sUserId = oData["UserId"];

		if (oData["Solved"]) this.m_bSolved = oData["Solved"];

		if (oData["UserData"]) this.m_sUserData = oData["UserData"];

		if (oData["Replies"] && oData["Replies"].length) {
			for (var nIndex = 0, nCount = oData["Replies"].length; nIndex < nCount; ++nIndex) {
				var oCD = new CCommentData();
				oCD.ReadFromSimpleObject(oData["Replies"][nIndex]);
				this.m_aReplies.push(oCD);
			}
		}
	};

	CCommentData.prototype.Write_ToBinary2 = function (Writer) {
		// String            : m_sText
		// String            : m_sTime
		// String            : m_sOOTime
		// String            : m_sUserId
		// String            : m_sUserName
		// String            : m_sGuid
		// Bool              : Null ли TimeZoneBias
		// Long              : TimeZoneBias
		// Bool              : Null ли QuoteText
		// String            : (Если предыдущий параметр false) QuoteText
		// Bool              : Solved
		// Long              : Количество отетов
		// Array of Variable : Ответы

		var Count = this.m_aReplies.length;
		Writer.WriteString2(this.m_sText);
		Writer.WriteString2(this.m_sTime);
		Writer.WriteString2(this.m_sOOTime);
		Writer.WriteString2(this.m_sUserId);
		Writer.WriteString2(this.m_sUserName);
		Writer.WriteString2(this.m_sGuid);

		if (null === this.m_nTimeZoneBias) Writer.WriteBool(true); else {
			Writer.WriteBool(false);
			Writer.WriteLong(this.m_nTimeZoneBias);
		}
		if (null === this.m_sQuoteText) Writer.WriteBool(true); else {
			Writer.WriteBool(false);
			Writer.WriteString2(this.m_sQuoteText);
		}
		Writer.WriteBool(this.m_bSolved);
		Writer.WriteLong(Count);

		for (var Index = 0; Index < Count; Index++) {
			this.m_aReplies[Index].Write_ToBinary2(Writer);
		}
	};

	CCommentData.prototype.Read_FromBinary2 = function (Reader) {
		// String            : m_sText
		// String            : m_sTime
		// String            : m_sOOTime
		// String            : m_sUserId
		// String            : m_sGuid
		// Bool              : Null ли TimeZoneBias
		// Long              : TimeZoneBias
		// Bool              : Null ли QuoteText
		// String            : (Если предыдущий параметр false) QuoteText
		// Bool              : Solved
		// Long              : Количество отетов
		// Array of Variable : Ответы

		this.m_sText = Reader.GetString2();
		this.m_sTime = Reader.GetString2();
		this.m_sOOTime = Reader.GetString2();
		this.m_sUserId = Reader.GetString2();
		this.m_sUserName = Reader.GetString2();
		this.m_sGuid = Reader.GetString2();

		if (true != Reader.GetBool()) this.m_nTimeZoneBias = Reader.GetLong(); else this.m_nTimeZoneBias = null;
		var bNullQuote = Reader.GetBool();
		if (true != bNullQuote) this.m_sQuoteText = Reader.GetString2(); else this.m_sQuoteText = null;

		this.m_bSolved = Reader.GetBool();

		var Count = Reader.GetLong();
		this.m_aReplies.length = 0;
		for (var Index = 0; Index < Count; Index++) {
			var oReply = new CCommentData();
			oReply.Read_FromBinary2(Reader);
			this.m_aReplies.push(oReply);
		}
	};

	CCommentData.prototype.Write_ToBinary = function (Writer) {
		this.Write_ToBinary2(Writer);
	};

	CCommentData.prototype.Read_FromBinary = function (Reader) {
		this.Read_FromBinary2(Reader);
	};

	CCommentData.prototype.HasUserData = function (sUserId) {
		if (this.m_sUserId === sUserId) {
			return true;
		}
		return this.HasUserReplies(sUserId);
	};

	CCommentData.prototype.HasUserReplies = function (sUserId) {
		for (var nReply = 0; nReply < this.m_aReplies.length; ++nReply) {
			if (this.m_aReplies[nReply].HasUserData(sUserId)) {
				return true;
			}
		}
		return false;
	};
	CCommentData.prototype.IsUserComment = function (sUserId) {
		if (this.m_sUserId === sUserId) {
			return true;
		}
		return false;
	};

	CCommentData.prototype.RemoveUserReplies = function (sUserId) {
		for (var nReply = this.m_aReplies.length - 1; nReply > -1; --nReply) {
			if (this.m_aReplies[nReply].m_sUserId === sUserId) {
				this.m_aReplies.splice(nReply, 1);
			}
		}
	};


	var comment_type_Common = 1; // Комментарий к обычному тексу
	var comment_type_HdrFtr = 2; // Комментарий к колонтитулу

	function CComment(Parent, Data) {
		this.Id = AscCommon.g_oIdCounter.Get_NewId();

		this.Parent = Parent;
		this.Data = Data;

		this.x = null;
		this.y = null;
		this.selected = false;
		this.m_oTypeInfo = {
			Type: comment_type_Common, Data: null
		};

		this.m_oStartInfo = {
			X: 0, Y: 0, H: 0, PageNum: 0, ParaId: null
		};

		this.m_oEndInfo = {
			X: 0, Y: 0, H: 0, PageNum: 0, ParaId: null
		};

		this.Lock = new AscCommon.CLock(); // Зажат ли комментарий другим пользователем
		if (false === AscCommon.g_oIdCounter.m_bLoad) {
			this.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeMine, false);
			AscCommon.CollaborativeEditing.Add_Unlock2(this);
		}

		// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
		g_oTableId.Add(this, this.Id);
	}


	CComment.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_Comment;
	};
	CComment.prototype.GetId = function () {
		return this.Get_Id();
	};
	CComment.prototype.createDuplicate = function (Parent, bNewGuid) {
		var oData = this.Data ? this.Data.createDuplicate(bNewGuid) : null;
		var ret = new CComment(Parent, oData);
		ret.setPosition(this.x, this.y);
		return ret;
	};

	CComment.prototype.removeUserReplies = function (sUserId) {
		if (this.Data) {
			var oDataCopy = this.Data.createDuplicate();
			oDataCopy.RemoveUserReplies(sUserId);
			if (this.Data.Get_RepliesCount() !== oDataCopy.Get_RepliesCount()) {
				this.Set_Data(oDataCopy);
				editor.sync_ChangeCommentData(this.Get_Id(), this.Data);
			}
		}
	};

	CComment.prototype.hasUserReplies = function (sUserId) {
		if (!this.Data) {
			return false;
		}
		return this.Data.HasUserReplies(sUserId);
	};

	CComment.prototype.isMineComment = function () {
		var oDocInfo = editor && editor.DocInfo;
		if (oDocInfo) {
			return this.isUserComment(oDocInfo.get_UserId());
		}
		return false;
	};

	CComment.prototype.IsSolved = function () {
		return this.Data.Get_Solved();
	};

	CComment.prototype.isUserComment = function (sUserId) {
		if (!this.Data) {
			return false;
		}
		return this.Data.IsUserComment(sUserId);
	};

	CComment.prototype.hasUserData = function (sUserId) {
		if (!this.Data) {
			return false;
		}
		return this.Data.HasUserData(sUserId);
	};

	CComment.prototype.canBeDeleted = function () {
		var sUserName = this.GetUserName();
		if (AscCommon.UserInfoParser.canViewComment(sUserName) && AscCommon.UserInfoParser.canDeleteComment(sUserName)) {
			return true;
		}
		return false;
	};

	CComment.prototype.hit = function (x, y) {
		if (AscCommon.UserInfoParser.canViewComment(this.GetUserName()) === false) {
			return false;
		}
		var Flags = 0;
		if (this.selected) {
			Flags |= 1;
		}
		if (this.Data.m_aReplies.length > 0) {
			Flags |= 2;
		}
		var dd = editor.WordControl.m_oDrawingDocument;
		return x > this.x && x < this.x + dd.GetCommentWidth(Flags) && y > this.y && y < this.y + dd.GetCommentHeight(Flags);
	};

	CComment.prototype.setPosition = function (x, y) {
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Comment_Position, new AscFormat.CDrawingBaseCoordsWritable(this.x, this.y), new AscFormat.CDrawingBaseCoordsWritable(x, y)));
		this.x = x;
		this.y = y;
	};

	CComment.prototype.getPosition = function () {
		return { x: this.x, y: this.y };
	};

	CComment.prototype.draw = function (graphics) {
		var Flags = 0;
		if (this.selected) {
			Flags |= 1;
		}
		if (this.Data.m_aReplies.length > 0) {
			Flags |= 2;
		}
		var dd = editor.WordControl.m_oDrawingDocument;
		var w = dd.GetCommentWidth();
		var h = dd.GetCommentHeight();
		graphics.DrawPresentationComment(Flags, this.x, this.y, w, h);

		var oLock = this.Lock;
		if (oLock && AscCommon.c_oAscLockTypes.kLockTypeNone !== oLock.Get_Type()) {
			var bCoMarksDraw = true;
			var oApi = editor || Asc['editor'];
			if (oApi) {
				bCoMarksDraw = (!AscCommon.CollaborativeEditing.Is_Fast() || AscCommon.c_oAscLockTypes.kLockTypeMine !== oLock.Get_Type());
			}
			if (bCoMarksDraw) {
				graphics.DrawLockObjectRect(oLock.Get_Type(), this.x, this.y, w, h);
				return true;
			}
		}
	};

	CComment.prototype.Set_StartInfo = function (PageNum, X, Y, H, ParaId) {
		this.m_oStartInfo.X = X;
		this.m_oStartInfo.Y = Y;
		this.m_oStartInfo.H = H;
		this.m_oStartInfo.ParaId = ParaId;

		// Если у нас комментарий в колонтитуле, то номер страницы обновляется при нажатии на комментарий
		if (comment_type_Common === this.m_oTypeInfo.Type) this.m_oStartInfo.PageNum = PageNum;
	};

	CComment.prototype.Set_EndInfo = function (PageNum, X, Y, H, ParaId) {
		this.m_oEndInfo.X = X;
		this.m_oEndInfo.Y = Y;
		this.m_oEndInfo.H = H;
		this.m_oEndInfo.ParaId = ParaId;

		if (comment_type_Common === this.m_oTypeInfo.Type) this.m_oEndInfo.PageNum = PageNum;
	};

	CComment.prototype.Check_ByXY = function (PageNum, X, Y, Type) {
		if (this.m_oTypeInfo.Type != Type) return false;

		if (comment_type_Common === Type) {
			if (PageNum < this.m_oStartInfo.PageNum || PageNum > this.m_oEndInfo.PageNum) return false;

			if (PageNum === this.m_oStartInfo.PageNum && (Y < this.m_oStartInfo.Y || (Y < (this.m_oStartInfo.Y + this.m_oStartInfo.H) && X < this.m_oStartInfo.X))) return false;

			if (PageNum === this.m_oEndInfo.PageNum && (Y > this.m_oEndInfo.Y + this.m_oEndInfo.H || (Y > this.m_oEndInfo.Y && X > this.m_oEndInfo.X))) return false;
		}
		else if (comment_type_HdrFtr === Type) {
			var HdrFtr = this.m_oTypeInfo.Data;

			if (null === HdrFtr || false === HdrFtr.Check_Page(PageNum)) return false;

			if (Y < this.m_oStartInfo.Y || (Y < (this.m_oStartInfo.Y + this.m_oStartInfo.H) && X < this.m_oStartInfo.X)) return false;

			if (Y > this.m_oEndInfo.Y + this.m_oEndInfo.H || (Y > this.m_oEndInfo.Y && X > this.m_oEndInfo.X)) return false;

			this.m_oStartInfo.PageNum = PageNum;
			this.m_oEndInfo.PageNum = PageNum;
		}

		return true;
	};

	CComment.prototype.Set_Data = function (Data) {
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Comment_Change, this.Data, Data));
		this.Data = Data;
	};

	CComment.prototype.Get_Data = function () {
		return this.GetData();
	};

	CComment.prototype.GetData = function () {
		return this.Data;
	};

	CComment.prototype.RemoveMarks = function () {
		var Para_start = g_oTableId.Get_ById(this.m_oStartInfo.ParaId);
		var Para_end = g_oTableId.Get_ById(this.m_oEndInfo.ParaId);

		if (Para_start === Para_end) {
			if (null != Para_start) Para_start.RemoveCommentMarks(this.Id);
		}
		else {
			if (null != Para_start) Para_start.RemoveCommentMarks(this.Id);

			if (null != Para_end) Para_end.RemoveCommentMarks(this.Id);
		}
	};

	CComment.prototype.Set_TypeInfo = function (Type, Data) {
		var New = {
			Type: Type, Data: Data
		};

		History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_Comment_TypeInfo, this.m_oTypeInfo, New));

		this.m_oTypeInfo = New;
	};

	CComment.prototype.Get_TypeInfo = function () {
		return this.m_oTypeInfo;
	};


	CComment.prototype.Refresh_RecalcData = function (Data) {
		if (this.slideComments) {
			this.slideComments.Refresh_RecalcData();
		}
	};

	CComment.prototype.recalculate = function () {
	};
	//-----------------------------------------------------------------------------------
	// Функции для работы с совместным редактированием
	//-----------------------------------------------------------------------------------
	CComment.prototype.Get_Id = function () {
		return this.Id;
	};

	CComment.prototype.Write_ToBinary2 = function (Writer) {
		Writer.WriteLong(AscDFH.historyitem_type_Comment);

		// String   : Id
		// Variable : Data
		// Long     : m_oTypeInfo.Type
		//          : m_oTypeInfo.Data
		//    Если comment_type_HdrFtr
		//    String : Id колонтитула

		Writer.WriteString2(this.Id);
		AscFormat.writeObject(Writer, this.Parent);
		this.Data.Write_ToBinary2(Writer);
		Writer.WriteLong(this.m_oTypeInfo.Type);

		if (comment_type_HdrFtr === this.m_oTypeInfo.Type) Writer.WriteString2(this.m_oTypeInfo.Data.Get_Id());
	};

	CComment.prototype.Read_FromBinary2 = function (Reader) {
		// String   : Id
		// Variable : Data
		// Long     : m_oTypeInfo.Type
		//          : m_oTypeInfo.Data
		//    Если comment_type_HdrFtr
		//    String : Id колонтитула

		this.Id = Reader.GetString2();
		this.Parent = AscFormat.readObject(Reader);
		this.Data = new CCommentData();
		this.Data.Read_FromBinary2(Reader);
		this.m_oTypeInfo.Type = Reader.GetLong();
		if (comment_type_HdrFtr === this.m_oTypeInfo.Type) this.m_oTypeInfo.Data = g_oTableId.Get_ById(Reader.GetString2());
	};

	CComment.prototype.Check_MergeData = function () {
		// Проверяем, не удалили ли мы параграф, к которому был сделан данный комментарий
		// Делаем это в самом конце, а не сразу, чтобы заполнились данные о начальном и
		// конечном параграфах.

		var bUse = true;

		if (null != this.m_oStartInfo.ParaId) {
			var Para_start = g_oTableId.Get_ById(this.m_oStartInfo.ParaId);

			if (true != Para_start.IsUseInDocument()) bUse = false;
		}

		if (true === bUse && null != this.m_oEndInfo.ParaId) {
			var Para_end = g_oTableId.Get_ById(this.m_oEndInfo.ParaId);

			if (true != Para_end.IsUseInDocument()) bUse = false;
		}

		if (false === bUse) editor.WordControl.m_oLogicDocument.RemoveComment(this.Id, true);
	};

	CComment.prototype.GetUserName = function () {
		if (this.Data) {
			return this.Data.Get_Name();
		}
		return "";
	};

//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};

	window['AscCommon'].comment_type_Common = comment_type_Common;
	window['AscCommon'].comment_type_HdrFtr = comment_type_HdrFtr;

	window['AscCommon'].CCommentData = CCommentData;
	window['AscCommon'].CComment = CComment;
	window['AscCommon'].ParaComment = ParaComment;
	window['AscCommon'].CCommentAuthor = CCommentAuthor;
	window['AscCommon'].CWriteCommentData = CWriteCommentData;

})(window);
