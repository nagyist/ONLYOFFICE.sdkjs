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
var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
function CParagraphContentBase()
{
	this.Type      = para_Unknown;
	this.Paragraph = null;
	this.Parent    = null;

	this.StartLine  = -1;
	this.StartRange = -1;

	this.Lines       = [];
	this.LinesLength = 0;
}
CParagraphContentBase.prototype.GetType = function()
{
	return this.Type;
};
CParagraphContentBase.prototype.Get_Type = function()
{
	return this.Type;
};
CParagraphContentBase.prototype.GetLogicDocument = function()
{
	return this.Paragraph ? this.Paragraph.GetLogicDocument() : null;
};
CParagraphContentBase.prototype.GetLinesCount = function()
{
	return 0;
};
CParagraphContentBase.prototype.CanSplit = function()
{
	return false;
};
CParagraphContentBase.prototype.IsParagraphContentElement = function()
{
	return true;
};
CParagraphContentBase.prototype.IsStopCursorOnEntryExit = function()
{
	return false;
};
CParagraphContentBase.prototype.PreDelete = function()
{
};
CParagraphContentBase.prototype.GetCurrentPermRanges = function(permRanges, isCurrent)
{
};
/**
 * Выствялем параграф, в котром лежит данный элемент
 * @param {Paragraph} oParagraph
 */
CParagraphContentBase.prototype.SetParagraph = function(oParagraph)
{
	this.Paragraph = oParagraph;
};
/**
 * ВЫставляем родительский класс
 * @param oParent
 */
CParagraphContentBase.prototype.SetParent = function(oParent)
{
	this.Parent = oParent;
};
/**
 * Получаем параграф, в котором лежит данный элемент
 * @returns {null|Paragraph}
 */
CParagraphContentBase.prototype.GetParagraph = function()
{
	return this.Paragraph;
};
CParagraphContentBase.prototype.IsThisElementCurrent = function()
{
	return false;
};
CParagraphContentBase.prototype.IsRun = function()
{
	return false;
};
CParagraphContentBase.prototype.IsMath = function()
{
	return false;
};
CParagraphContentBase.prototype.Is_Empty = function()
{
	return true;
};
CParagraphContentBase.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CParagraphContentBase.prototype.Is_CheckingNearestPos = function()
{
	return false;
};
CParagraphContentBase.prototype.Get_CompiledTextPr = function()
{
	return null;
};
CParagraphContentBase.prototype.Clear_TextPr = function()
{

};
CParagraphContentBase.prototype.Remove = function()
{
	return false;
};
CParagraphContentBase.prototype.Get_DrawingObjectRun = function(Id)
{
	return null;
};
CParagraphContentBase.prototype.Get_DrawingObjectContentPos = function(Id, ContentPos, Depth)
{
	return false;
};
CParagraphContentBase.prototype.GetRunByElement = function(oRunElement)
{
	return null;
};
CParagraphContentBase.prototype.Get_Layout = function(DrawingLayout, UseContentPos, ContentPos, Depth)
{
};
/**
 * Ищем список элементов, идущих после заданной позициц
 * @param oRunElements {CParagraphRunElements}
 * @param isUseContentPos {boolean}
 * @param nDepth {number}
 */
CParagraphContentBase.prototype.GetNextRunElements = function(oRunElements, isUseContentPos, nDepth)
{
};
/**
 * Ищем список элементов рана, предществующих заданной позиции
 * @param oRunElements {CParagraphRunElements}
 * @param isUseContentPos {boolean}
 * @param nDepth {number}
 */
CParagraphContentBase.prototype.GetPrevRunElements = function(oRunElements, isUseContentPos, nDepth)
{
};
CParagraphContentBase.prototype.CollectDocumentStatistics = function(ParaStats)
{
};
CParagraphContentBase.prototype.Create_FontMap = function(Map)
{
};
CParagraphContentBase.prototype.Get_AllFontNames = function(AllFonts)
{
};
CParagraphContentBase.prototype.GetSelectedText = function(bAll, bClearText, oPr)
{
	return "";
};
CParagraphContentBase.prototype.GetSelectDirection = function()
{
	return 1;
};
CParagraphContentBase.prototype.Clear_TextFormatting = function( DefHyper )
{
};
/**
 * Проверяем можно ли добавлять буквицу
 * @returns {null | boolean}
 */
CParagraphContentBase.prototype.CanAddDropCap = function()
{
	return null;
};
/**
 * Проверяем можно ли использовать селект для добавления буквицы
 * @param isUsePos {boolean}
 * @param oEndPos {AscWord.CParagraphContentPos}
 * @param nDepth {number}
 * @returns {boolean}
 */
CParagraphContentBase.prototype.CheckSelectionForDropCap = function(isUsePos, oEndPos, nDepth)
{
	return true;
};
CParagraphContentBase.prototype.Get_TextForDropCap = function(DropCapText, UseContentPos, ContentPos, Depth)
{
};
CParagraphContentBase.prototype.Get_StartTabsCount = function(TabsCounter)
{
	return true;
};
CParagraphContentBase.prototype.Remove_StartTabs = function(TabsCounter)
{
	return true;
};
CParagraphContentBase.prototype.Copy = function(Selected, oPr, isCopyReviewPr)
{
	return new this.constructor();
};
CParagraphContentBase.prototype.GetSelectedContent = function(oSelectedContent)
{
	return this.Copy();
};
CParagraphContentBase.prototype.CopyContent = function(Selected)
{
	return [];
};
CParagraphContentBase.prototype.Split = function()
{
	return new ParaRun();
};
CParagraphContentBase.prototype.SplitNoDuplicate = function(oContentPos, nDepth, oNewParagraph)
{

};
CParagraphContentBase.prototype.Get_Text = function(Text)
{
};
CParagraphContentBase.prototype.Apply_TextPr = function(oTextPr, isIncFontSize, isApplyToAll)
{
};
CParagraphContentBase.prototype.Get_ParaPosByContentPos = function(ContentPos, Depth)
{
	return new CParaPos(this.StartRange, this.StartLine, 0, 0);
};
CParagraphContentBase.prototype.UpdateBookmarks = function(oManager)
{
};
/**
 * @param oSpellCheckerEngine {AscWord.CParagraphSpellCheckerCollector}
 * @param nDepth {number}
 */
CParagraphContentBase.prototype.CheckSpelling = function(oSpellCheckerEngine, nDepth)
{
};
CParagraphContentBase.prototype.GetParent = function()
{
	if (this.Parent)
		return this.Parent;

	if (!this.Paragraph)
		return null;

	var oContentPos = this.Paragraph.Get_PosByElement(this);
	if (!oContentPos || oContentPos.GetDepth() < 0)
		return null;

	oContentPos.DecreaseDepth(1);
	return this.Paragraph.Get_ElementByPos(oContentPos);
};
CParagraphContentBase.prototype.GetPosInParent = function(_oParent)
{
	var oParent = (_oParent? _oParent : this.GetParent());
	if (!oParent || !oParent.Content)
		return -1;

	for (var nPos = 0, nCount = oParent.Content.length; nPos < nCount; ++nPos)
	{
		if (this === oParent.Content[nPos])
			return nPos;
	}

	return -1;
};
CParagraphContentBase.prototype.RemoveThisFromParent = function(updatePosition)
{
	let parent      = this.GetParent();
	let posInParent = this.GetPosInParent(parent);
	
	if (parent && -1 !== posInParent)
		parent.RemoveFromContent(posInParent, 1);
	
	if (false !== updatePosition)
	{
		if (posInParent < parent.GetElementsCount() && parent.GetElement(posInParent).IsCursorPlaceable())
		{
			parent.GetElement(posInParent).MoveCursorToStartPos();
			parent.GetElement(posInParent).SetThisElementCurrentInParagraph();
		}
		else if (posInParent > 0 && parent.GetElement(posInParent - 1).IsCursorPlaceable())
		{
			parent.GetElement(posInParent - 1).MoveCursorToStartPos();
			parent.GetElement(posInParent - 1).SetThisElementCurrentInParagraph();
		}
	}
	
	parent.CorrectContent();
};
CParagraphContentBase.prototype.RemoveTabsForTOC = function(isTab)
{
	return isTab;
};
/**
 * Ищем сложное поле заданного типа
 * @param nType
 * @returns {?CComplexField}
 */
CParagraphContentBase.prototype.GetComplexField = function(nType)
{
	return null;
};
/**
 * Ищем все сложные поля заданного типа
 * @param nType
 * @param arrComplexFields
 */
CParagraphContentBase.prototype.GetComplexFieldsArray = function(nType, arrComplexFields)
{
};
//----------------------------------------------------------------------------------------------------------------------
// Функции пересчета
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentBase.prototype.Recalculate_Reset = function(StartRange, StartLine)
{
	this.StartLine   = StartLine;
	this.StartRange  = StartRange;
};
CParagraphContentBase.prototype.Recalculate_Range = function(PRS, ParaPr)
{
};
CParagraphContentBase.prototype.Recalculate_Set_RangeEndPos = function(PRS, PRP, Depth)
{
};
CParagraphContentBase.prototype.Recalculate_SetRangeBounds = function(_CurLine, _CurRange, oStartPos, oEndPos, nDepth)
{
};
CParagraphContentBase.prototype.GetContentWidthInRange = function(oStartPos, oEndPos, nDepth)
{
	return 0;
};
CParagraphContentBase.prototype.Recalculate_LineMetrics = function(PRS, ParaPr, _CurLine, _CurRange)
{
};
CParagraphContentBase.prototype.Recalculate_Range_Width = function(PRSC, _CurLine, _CurRange)
{
};
CParagraphContentBase.prototype.Recalculate_Range_Spaces = function(PRSA, CurLine, CurRange, CurPage)
{
};
CParagraphContentBase.prototype.Recalculate_PageEndInfo = function(PRSI, _CurLine, _CurRange)
{
};
CParagraphContentBase.prototype.RecalculateEndInfo = function(oPRSI)
{

};
CParagraphContentBase.prototype.SaveRecalculateObject = function(Copy)
{
	var RecalcObj = new CRunRecalculateObject(this.StartLine, this.StartRange);
	return RecalcObj;
};
CParagraphContentBase.prototype.LoadRecalculateObject = function(RecalcObj, Parent)
{
	this.StartLine  = RecalcObj.StartLine;
	this.StartRange = RecalcObj.StartRange;
};
CParagraphContentBase.prototype.PrepareRecalculateObject = function()
{
};
/**
 * Пустой ли заданный отрезок
 * @param nCurLine {number}
 * @param nCurRange {number}
 * @returns {boolean}
 */
CParagraphContentBase.prototype.IsEmptyRange = function(nCurLine, nCurRange)
{
	return true;
};
CParagraphContentBase.prototype.Check_Range_OnlyMath = function(Checker, CurRange, CurLine)
{
};
/**
 * Проверяем является ли элемент в заданной позиции неинлайновой формулой
 * @param {number} nMathPos
 * @return {boolean}
 */
CParagraphContentBase.prototype.CheckMathPara = function(nMathPos)
{
	return false;
};
CParagraphContentBase.prototype.ProcessNotInlineObjectCheck = function(oChecker)
{
};
CParagraphContentBase.prototype.CheckNotInlineObject = function(nMathPos, nDirection)
{
	return false;
};
CParagraphContentBase.prototype.Check_PageBreak = function()
{
	return false;
};
/**
 * Проверяем нужно ли разрывать страницу после заданного PageBreak элемента
 * @param oPBChecker {CParagraphCheckSplitPageOnPageBreak}
 * @returns {boolean}
 */
CParagraphContentBase.prototype.CheckSplitPageOnPageBreak = function(oPBChecker)
{
	return false;
};
CParagraphContentBase.prototype.recalculateCursorPosition = function(positionCalculator, isCurrent)
{
};
CParagraphContentBase.prototype.RecalculateMinMaxContentWidth = function(MinMax)
{
};
CParagraphContentBase.prototype.Get_Range_VisibleWidth = function(RangeW, _CurLine, _CurRange)
{
};
CParagraphContentBase.prototype.Shift_Range = function(Dx, Dy, _CurLine, _CurRange, _CurPage)
{
};
//----------------------------------------------------------------------------------------------------------------------
// Функции отрисовки
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentBase.prototype.Draw_HighLights = function(PDSH)
{
};
CParagraphContentBase.prototype.Draw_Elements = function(PDSE)
{
};
CParagraphContentBase.prototype.Draw_Lines = function(PDSL)
{
};
CParagraphContentBase.prototype.SkipDraw = function(PDS)
{
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с курсором
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentBase.prototype.IsCursorPlaceable = function()
{
	return false;
};
CParagraphContentBase.prototype.Cursor_Is_Start = function()
{
	return true;
};
CParagraphContentBase.prototype.Cursor_Is_NeededCorrectPos = function()
{
	return true;
};
CParagraphContentBase.prototype.Cursor_Is_End = function()
{
	return true;
};
CParagraphContentBase.prototype.IsStartPos = function(contentPos, depth)
{
	return true;
};
CParagraphContentBase.prototype.IsEndPos = function(contentPos, depth)
{
	return true;
};
/**
 * TODO: Надо объединить эту функцию с  IsCursorPlaceable, поскольку они по смыслу одинаковые
 * и сделать тут по умолчанию false
 */
CParagraphContentBase.prototype.CanPlaceCursorInside = function()
{
	return true;
};
CParagraphContentBase.prototype.MoveCursorToStartPos = function()
{
};
CParagraphContentBase.prototype.MoveCursorToEndPos = function(SelectFromEnd)
{
};
CParagraphContentBase.prototype.getParagraphContentPosByXY = function(searchState)
{
	return false;
};
CParagraphContentBase.prototype.Get_ParaContentPos = function(bSelection, bStart, ContentPos, bUseCorrection)
{
};
CParagraphContentBase.prototype.Set_ParaContentPos = function(ContentPos, Depth)
{
};
CParagraphContentBase.prototype.Get_PosByElement = function(Class, ContentPos, Depth, UseRange, Range, Line)
{
	if (this === Class)
		return true;

	return false;
};
CParagraphContentBase.prototype.Get_ElementByPos = function(ContentPos, Depth)
{
	return this;
};
CParagraphContentBase.prototype.Get_ClassesByPos = function(Classes, ContentPos, Depth)
{
	Classes.push(this);
};
CParagraphContentBase.prototype.GetPosByDrawing = function(Id, ContentPos, Depth)
{
	return false;
};
CParagraphContentBase.prototype.Get_RunElementByPos = function(ContentPos, Depth)
{
	return null;
};
CParagraphContentBase.prototype.Get_LastRunInRange = function(_CurLine, _CurRange)
{
	return null;
};
CParagraphContentBase.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
};
CParagraphContentBase.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
};
CParagraphContentBase.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
};
CParagraphContentBase.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
};
CParagraphContentBase.prototype.Get_EndRangePos = function(_CurLine, _CurRange, SearchPos, Depth)
{
	return false;
};
CParagraphContentBase.prototype.Get_StartRangePos = function(_CurLine, _CurRange, SearchPos, Depth)
{
	return false;
};
CParagraphContentBase.prototype.Get_StartRangePos2 = function(_CurLine, _CurRange, ContentPos, Depth)
{
};
CParagraphContentBase.prototype.Get_EndRangePos2 = function(_CurLine, _CurRange, ContentPos, Depth)
{
};
CParagraphContentBase.prototype.Get_StartPos = function(ContentPos, Depth)
{
};
CParagraphContentBase.prototype.Get_EndPos = function(BehindEnd, ContentPos, Depth)
{
};
CParagraphContentBase.prototype.MoveCursorOutsideElement = function(isBefore)
{
	var oParent = this.GetParent();
	if (!oParent)
		return;

	var nPosInParent = this.GetPosInParent(oParent);

	if (isBefore)
	{
		if (nPosInParent <= 0)
		{
			if (this.SetThisElementCurrent)
				this.SetThisElementCurrent();

			this.MoveCursorToStartPos();
		}
		else
		{
			var oElement = oParent.GetElement(nPosInParent - 1);
			if (oElement.IsCursorPlaceable())
			{
				if (oElement.SetThisElementCurrent)
					oElement.SetThisElementCurrent();

				oElement.MoveCursorToEndPos();
			}
			else
			{
				if (this.SetThisElementCurrent)
					this.SetThisElementCurrent();

				this.MoveCursorToStartPos();
			}
		}
	}
	else
	{
		if (nPosInParent >= oParent.GetElementsCount() - 1)
		{
			if (this.SetThisElementCurrent)
				this.SetThisElementCurrent();

			this.MoveCursorToEndPos();
		}
		else
		{
			var oElement = oParent.GetElement(nPosInParent + 1);
			if (oElement.IsCursorPlaceable())
			{
				if (oElement.SetThisElementCurrent)
					oElement.SetThisElementCurrent();

				oElement.MoveCursorToStartPos();
			}
			else
			{
				if (this.SetThisElementCurrent)
					this.SetThisElementCurrent();

				this.MoveCursorToEndPos();
			}
		}
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с селектом
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentBase.prototype.Set_SelectionContentPos = function(StartContentPos, EndContentPos, Depth, StartFlag, EndFlag)
{
};
CParagraphContentBase.prototype.RemoveSelection = function()
{
};
CParagraphContentBase.prototype.SelectAll = function(Direction)
{
};
CParagraphContentBase.prototype.drawSelectionInRange = function(line, range, drawState)
{
};
CParagraphContentBase.prototype.IsSelectionEmpty = function(CheckEnd)
{
	return true;
};
CParagraphContentBase.prototype.Selection_CheckParaEnd = function()
{
	return false;
};
CParagraphContentBase.prototype.IsSelectedAll = function(Props)
{
	return true;
};
CParagraphContentBase.prototype.IsSelectedFromStart = function()
{
	return true;
};
CParagraphContentBase.prototype.IsSelectedToEnd = function()
{
	return true;
};
/**
 * Функция коррекции селекта, чтобы убрать из селекта плавающие объекты, идущие в начале
 * @param nDirection {number} - направление селекта
 * @returns {boolean}
 */
CParagraphContentBase.prototype.SkipAnchorsAtSelectionStart = function(nDirection)
{
	return true;
};
CParagraphContentBase.prototype.Selection_CheckParaContentPos = function(ContentPos)
{
	return true;
};
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentBase.prototype.GetCurrentParaPos = function(align)
{
	return new CParaPos(this.StartRange, this.StartLine, 0, 0);
};
CParagraphContentBase.prototype.Get_TextPr = function(ContentPos, Depth)
{
	return new CTextPr();
};
CParagraphContentBase.prototype.Get_FirstTextPr = function(bByPos)
{
	return new CTextPr();
};
CParagraphContentBase.prototype.SetReviewType = function(ReviewType, RemovePrChange)
{
};
CParagraphContentBase.prototype.SetReviewTypeWithInfo = function(ReviewType, ReviewInfo)
{
};
CParagraphContentBase.prototype.CheckRevisionsChanges = function(Checker, ContentPos, Depth)
{
};
CParagraphContentBase.prototype.AcceptRevisionChanges = function(Type, bAll)
{
};
CParagraphContentBase.prototype.RejectRevisionChanges = function(Type, bAll)
{
};
CParagraphContentBase.prototype.GetTextPr = function(ContentPos, Depth)
{
	return this.Get_TextPr(ContentPos, Depth);
};
CParagraphContentBase.prototype.ApplyTextPr = function(oTextPr, isIncFontSize, isApplyToAll)
{
	return this.Apply_TextPr(oTextPr, isIncFontSize, isApplyToAll);
};
/**
 * Функция для поиска внутри элементов параграфа
 * @param {AscCommonWord.CParagraphSearch} oParaSearch
 */
CParagraphContentBase.prototype.Search = function(oParaSearch)
{
};
CParagraphContentBase.prototype.AddSearchResult = function(oSearchResult, isStart, oContentPos, nDepth)
{
};
CParagraphContentBase.prototype.ClearSearchResults = function()
{
};
CParagraphContentBase.prototype.RemoveSearchResult = function(oSearchResult)
{
};
CParagraphContentBase.prototype.GetSearchElementId = function(bNext, bUseContentPos, ContentPos, Depth)
{
	return null;
};
CParagraphContentBase.prototype.Check_NearestPos = function(ParaNearPos, Depth)
{
};
CParagraphContentBase.prototype.RestartSpellCheck = function()
{
};
CParagraphContentBase.prototype.GetDirectTextPr = function()
{
	return null;
};
CParagraphContentBase.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	return arrFields ? arrFields : [];
};
CParagraphContentBase.prototype.GetAllSeqFieldsByType = function(sType, aFields)
{
};
/**
 * Проверяем можно ли добавлять комментарий по заданому селекту
 * @returns {boolean}
 */
CParagraphContentBase.prototype.CanAddComment = function()
{
	return true;
};
/**
 * Получаем позицию заданного элемента в документе
 * @param {?Array} arrPos
 * @returns {Array}
 */
CParagraphContentBase.prototype.GetDocumentPositionFromObject = function(arrPos)
{
	if (!arrPos)
		arrPos = [];

	var oParagraph = this.GetParagraph();
	if (oParagraph)
	{
		if (arrPos.length > 0)
		{
			var oParaContentPos = oParagraph.Get_PosByElement(this);
			if (oParaContentPos)
			{
				var nDepth = oParaContentPos.GetDepth();
				while (nDepth > 0)
				{
					var Pos = oParaContentPos.Get(nDepth);
					oParaContentPos.SetDepth(nDepth - 1);
					var Class = oParagraph.Get_ElementByPos(oParaContentPos);
					nDepth--;

					arrPos.splice(0, 0, {Class : Class, Position : Pos});
				}
				arrPos.splice(0, 0, {Class : this.Paragraph, Position : oParaContentPos.Get(0)});
			}

			this.Paragraph.GetDocumentPositionFromObject(arrPos);
		}
		else
		{
			this.Paragraph.GetDocumentPositionFromObject(arrPos);

			var oParaContentPos = this.Paragraph.Get_PosByElement(this);
			if (oParaContentPos)
			{
				arrPos.push({Class : this.Paragraph, Position : oParaContentPos.Get(0)});

				var nDepth    = oParaContentPos.GetDepth();
				var nCurDepth = 1;
				while (nCurDepth <= nDepth)
				{
					var Pos = oParaContentPos.Get(nCurDepth);
					oParaContentPos.SetDepth(nCurDepth - 1);
					var Class = this.Paragraph.Get_ElementByPos(oParaContentPos);
					++nCurDepth;

					arrPos.push({Class : Class, Position : Pos});
				}
			}
		}
	}

	return arrPos;
};
/**
 * Получаем массив всех конент контролов, внутри которых лежит данный класс
 * @returns {Array}
 */
CParagraphContentBase.prototype.GetParentContentControls = function()
{
	var oDocPos = this.GetDocumentPositionFromObject();
	oDocPos.push({Class : this, Pos : 0});

	var arrContentControls = [];
	for (var nIndex = 0, nCount = oDocPos.length; nIndex < nCount; ++nIndex)
	{
		if (oDocPos[nIndex].Class instanceof CInlineLevelSdt)
			arrContentControls.push(oDocPos[nIndex].Class);
		else if (oDocPos[nIndex].Class instanceof CDocumentContent && oDocPos[nIndex].Class.Parent instanceof CBlockLevelSdt)
			arrContentControls.push(oDocPos[nIndex].Class.Parent);
	}

	return arrContentControls;
};
/**
 * Проверяем есть ли выделение внутри объекта
 * @returns {boolean}
 */
CParagraphContentBase.prototype.IsSelectionUse = function()
{
	return false;
};
/**
 * Начинается ли элемент с новой строки
 * @returns {boolean}
 */
CParagraphContentBase.prototype.IsStartFromNewLine = function()
{
	return false;
};
/**
 * Удаляем из параграфа заданный элемент, если он тут есть
 * @param element
 */
CParagraphContentBase.prototype.RemoveElement = function(element)
{
};
/**
 * Пробегаемся по все ранам с заданной функцией
 * @param fCheck - функция проверки содержимого рана
 * @param {AscWord.CParagraphContentPos} oStartPos
 * @param {AscWord.CParagraphContentPos} oEndPos
 * @param {number} nDepth
 * @param {?AscWord.CParagraphContentPos} oCurrentPos
 * @param {boolean} [isForward=false]
 * @returns {boolean}
 */
CParagraphContentBase.prototype.CheckRunContent = function(fCheck, oStartPos, oEndPos, nDepth, oCurrentPos, isForward)
{
	return false;
};
/**
 * Собираем сложные поля параграфа
 * @param {AscWord.ParagraphComplexFieldStack} oComplexFields
 */
CParagraphContentBase.prototype.ProcessComplexFields = function(oComplexFields)
{
};
/**
 * Собираем информацию о выделенной части документа
 * @param oInfo {CSelectedElementsInfo}
 * @param oContentPos
 * @param nDepth
 */
CParagraphContentBase.prototype.GetSelectedElementsInfo = function(oInfo, oContentPos, nDepth)
{
};
/**
 * Проверяем является ли данный элемент цельным, т.е. его нальзя разбить на части и
 * @returns {boolean}
 */
CParagraphContentBase.prototype.IsSolid = function()
{
	return true;
};
/**
 * Корректируем позицию внутри контента, для элементов с содержимым
 */
CParagraphContentBase.prototype.CorrectContentPos = function()
{
};
/**
 * Получаем самый первый ран в параграфе
 * @returns {?ParaRun}
 */
CParagraphContentBase.prototype.GetFirstRun = function()
{
	return null;
};
/**
 * Делаем данный элемент, состоящим из одного рана и возвращаем его, либо возвращаем null
 * @returns {?ParaRun}
 */
CParagraphContentBase.prototype.MakeSingleRunElement = function()
{
	return null;
};
/**
 * Очищаем полностью содержимое данного элемента
 * !!!ПУСТОЙ РАН ТУТ НЕ ДОБАВЛЯЕТСЯ!!!
 */
CParagraphContentBase.prototype.ClearContent = function()
{
};
/**
 * Получаем позиции до и после первого элемента у рана заданного типа
 * @param nType {number}
 * @param oStartPos {AscWord.CParagraphContentPos}
 * @param oEndPos {AscWord.CParagraphContentPos}
 * @param nDepth {number} глубина данного элемента
 * @returns {boolean}
 */
CParagraphContentBase.prototype.GetFirstRunElementPos = function(nType, oStartPos, oEndPos, nDepth)
{
	return false;
};
/**
 * @param isRecalculated
 */
CParagraphContentBase.prototype.SetIsRecalculated = function(isRecalculated)
{
};
/**
 * Устанавливаем текущие позиции на текущий элемент
 */
CParagraphContentBase.prototype.SetThisElementCurrentInParagraph = function()
{
	var oParagraph = this.GetParagraph();
	if (!this.IsCursorPlaceable() || !oParagraph)
		return;

	let contentPos = this.Paragraph.Get_PosByElement(this);
	if (!contentPos)
		return;
	
	// Дополним полученную позицию текущей в текущем элементе
	this.Get_ParaContentPos(false, false, contentPos, false);

	this.Paragraph.Set_ParaContentPos(contentPos, true, -1, -1, false);
};
CParagraphContentBase.prototype.createDuplicateForSmartArt = function(oPr)
{
	return this.Copy(false, oPr, false);
};
/**
 * Подсчитываем на сколько элементов разбивается данный элемент с заданным сепаратором
 * @param oEngine {CTextToTableEngine}
 */
CParagraphContentBase.prototype.CalculateTextToTable = function(oEngine){};

CParagraphContentBase.prototype.GetAllPermRangeMarks = function(marks)
{
	return [];
};
CParagraphContentBase.prototype.IsUseInDocument = function()
{
	return !!(this.Paragraph
		&& this.Paragraph.IsUseInDocument()
		&& this.IsUseInParagraph());
};
CParagraphContentBase.prototype.IsUseInParagraph = function()
{
	return (this.Paragraph && !!this.Paragraph.Get_PosByElement(this));
};

/**
 * Это базовый класс для элементов содержимого(контент) параграфа, у которых есть свое содержимое.
 * @constructor
 * @extends {CParagraphContentBase}
 */
function CParagraphContentWithContentBase()
{
	CParagraphContentBase.call(this);
    
    // Массив Lines разделен на три части
    // 1. Состоит из одного элемента, означающего количество строк
    // 2. Количество элементов указывается в первой части, каждый элемент означает относительный сдвиг начала информации 
    //    о строке в 3 части (поэтому первый элемент всегда равен 0).    
    // 3. Сама информация о начале и конце отрезка в строке. Каждый отрезок представлен парой StartPos, EndPos.
    //
    // Пример. 2 строки, в первой строке 3 отрезка, во второй строке 1 отрезок
    // this.Lines = [2, 0, 6, 0, 15, 15, 17, 17, 20, 20, 25];

	this.Lines = [0];

    this.StartLine   = -1;
    this.StartRange  = -1;
}

CParagraphContentWithContentBase.prototype = Object.create(CParagraphContentBase.prototype);
CParagraphContentWithContentBase.prototype.constructor = CParagraphContentWithContentBase;

CParagraphContentWithContentBase.prototype.Recalculate_Reset = function(StartRange, StartLine)
{
    this.StartLine   = StartLine;
    this.StartRange  = StartRange;

    this.protected_ClearLines();
};

CParagraphContentWithContentBase.prototype.protected_ClearLines = function()
{
	this.Lines = [0];
};

CParagraphContentWithContentBase.prototype.protected_GetRangeOffset = function(LineIndex, RangeIndex)
{
    return (1 + this.Lines[0] + this.Lines[1 + LineIndex] + RangeIndex * 2);
};

CParagraphContentWithContentBase.prototype.protected_GetRangeStartPos = function(LineIndex, RangeIndex)
{
    return this.Lines[this.protected_GetRangeOffset(LineIndex, RangeIndex)];
};

CParagraphContentWithContentBase.prototype.protected_GetRangeEndPos = function(LineIndex, RangeIndex)
{
    return this.Lines[this.protected_GetRangeOffset(LineIndex, RangeIndex) + 1];
};

CParagraphContentWithContentBase.prototype.protected_GetLinesCount = function()
{
    return this.Lines[0];
};

CParagraphContentWithContentBase.prototype.protected_GetRangesCount = function(LineIndex)
{
    if (LineIndex === this.Lines[0] - 1)
        return (this.Lines.length - this.Lines[1 + LineIndex] - (this.Lines[0] + 1)) / 2;
    else
        return (this.Lines[1 + LineIndex + 1] - this.Lines[1 + LineIndex]) / 2;
};

CParagraphContentWithContentBase.prototype.getRangePos = function(line, range)
{
	let _line  = line - this.StartLine;
	let _range = _line ? range : range - this.StartRange;
	
	return [
		this.protected_GetRangeStartPos(_line, _range),
		this.protected_GetRangeEndPos(_line, _range),
	];
};
CParagraphContentWithContentBase.prototype.GetLinesCount = function()
{
	return this.protected_GetLinesCount();
};

// Здесь предполагается, что строки с номерами меньше, чем LineIndex заданы, а также заданы и отрезки в строке 
// LineIndex, с номерами меньшими, чем RangeIndex. В данной функции удаляются все записи, которые идут после LineIndex,
// RangeIndex. Т.е. удаляются все строки, с номерами больше, чем LineIndex, и в строке LineIndex удаляются все отрезки 
// с номерами больше, чем RangeIndex. Возвращается позиция предпоследнего отрезка, либо 0.
CParagraphContentWithContentBase.prototype.protected_AddRange = function(LineIndex, RangeIndex)
{
    // Удаляем лишние записи о строках и отрезках
    if (this.Lines[0] >= LineIndex + 1)
    {
        var RangeOffset = this.protected_GetRangeOffset(LineIndex, 0) + RangeIndex * 2;
        this.Lines.splice(RangeOffset, this.Lines.length - RangeOffset);

        if (this.Lines[0] !== LineIndex + 1 && 0 === RangeIndex)
            this.Lines.splice(LineIndex + 1, this.Lines[0] - LineIndex);
        else if (this.Lines[0] !== LineIndex + 1 && 0 !== RangeIndex)
        {
            this.Lines.splice(LineIndex + 2, this.Lines[0] - LineIndex - 1);
            this.Lines[0] = LineIndex + 1;
        }
    }

    if (0 === RangeIndex)
    {
        if (this.Lines[0] !== LineIndex + 1)
        {
            // Добавляем информацию о новой строке, сначала ее относительный сдвиг, потом меняем само количество строк
            var OffsetValue = this.Lines.length - LineIndex - 1;
            this.Lines.splice(LineIndex + 1, 0, OffsetValue);
            this.Lines[0] = LineIndex + 1;
        }
    }
    
    var RangeOffset = 1 + this.Lines[0] + this.Lines[LineIndex + 1] + RangeIndex * 2; // this.protected_GetRangeOffset(LineIndex, RangeIndex);
    
    // Резервируем место для StartPos и EndPos заданного отрезка
    this.Lines[RangeOffset + 0] = 0;
    this.Lines[RangeOffset + 1] = 0;    
    
    if (0 !== LineIndex || 0 !== RangeIndex)
        return this.Lines[RangeOffset - 1];
    else
        return 0;
};

// Заполняем добавленный отрезок значениями
CParagraphContentWithContentBase.prototype.protected_FillRange = function(LineIndex, RangeIndex, StartPos, EndPos)
{
    var RangeOffset = this.protected_GetRangeOffset(LineIndex, RangeIndex);
    this.Lines[RangeOffset + 0] = StartPos;
    this.Lines[RangeOffset + 1] = EndPos;
};
CParagraphContentWithContentBase.prototype.protected_FillRangeEndPos = function(LineIndex, RangeIndex, EndPos)
{
    var RangeOffset = this.protected_GetRangeOffset(LineIndex, RangeIndex);
    this.Lines[RangeOffset + 1] = EndPos;
};
CParagraphContentWithContentBase.prototype.private_UpdateSpellChecking = function()
{
	if (this.Paragraph)
	{
		this.Paragraph.SpellChecker.ClearCollector();
		this.Paragraph.RecalcInfo.NeedSpellCheck();
	}
};
CParagraphContentWithContentBase.prototype.private_UpdateShapeText = function()
{
	if (this.Paragraph)
		this.Paragraph.RecalcInfo.NeedShapeText();
};
CParagraphContentWithContentBase.prototype.SelectThisElement = function(nDirection, isUseInnerSelection)
{
	if (!this.Paragraph)
		return false;

	var ContentPos = this.Paragraph.Get_PosByElement(this);
	if (!ContentPos)
		return false;

	var StartPos = ContentPos.Copy();
	var EndPos   = ContentPos.Copy();

	if (isUseInnerSelection)
	{
		this.Get_ParaContentPos(true, true, StartPos, false);
		this.Get_ParaContentPos(true, false, EndPos, false);
	}
	else
	{
		this.Get_StartPos(StartPos, StartPos.GetDepth() + 1);
		this.Get_EndPos(true, EndPos, EndPos.GetDepth() + 1);
	}

	if (nDirection < 0)
	{
		let Temp = StartPos;
		StartPos = EndPos;
		EndPos   = Temp;
	}

	this.Paragraph.Selection.Use   = true;
	this.Paragraph.Selection.Start = false;
	this.Paragraph.Set_ParaContentPos(StartPos, true, -1, -1);
	this.Paragraph.Set_SelectionContentPos(StartPos, EndPos, false);
	this.Paragraph.Document_SetThisElementCurrent(false);

	return true;
};
CParagraphContentWithContentBase.prototype.SetThisElementCurrent = function()
{
	let paragraph = this.GetParagraph();
	if (!paragraph)
		return;

	var ContentPos = paragraph.Get_PosByElement(this);
	if (!ContentPos)
		return;

	var StartPos = ContentPos.Copy();
	this.Get_StartPos(StartPos, StartPos.GetDepth() + 1);

	paragraph.Set_ParaContentPos(StartPos, true, -1, -1, false);
	paragraph.Document_SetThisElementCurrent(false);
};
CParagraphContentWithContentBase.prototype.IsThisElementCurrent = function()
{
	if (!this.Paragraph)
		return false;

	let oParaPos = this.Paragraph.GetPosByElement(this);
	if (!oParaPos)
		return false;

	if (!this.Paragraph.IsSelectionUse())
	{
		let oCurPos = this.Paragraph.Get_ParaContentPos(false, false, false);

		if (!oParaPos.IsPartOf(oCurPos))
			return false;
	}
	else
	{
		let oStartPos = this.Paragraph.Get_ParaContentPos(true, true, false);
		let oEndPos   = this.Paragraph.Get_ParaContentPos(true, false, false);

		if (!oParaPos.IsPartOf(oStartPos) || !oParaPos.IsPartOf(oEndPos))
			return false;
	}

	return this.Paragraph.IsThisElementCurrent();
};
CParagraphContentWithContentBase.prototype.GetStartPosInParagraph = function()
{
	if (!this.Paragraph)
		return null;

	let oContentPos = this.Paragraph.GetPosByElement(this);
	if (!oContentPos)
		return null;

	let oResultPos = oContentPos.Copy();
	this.Get_StartPos(oResultPos, oResultPos.GetDepth() + 1);
	return oResultPos;
};
CParagraphContentWithContentBase.prototype.GetEndPosInParagraph = function()
{
	if (!this.Paragraph)
		return null;

	let oContentPos = this.Paragraph.GetPosByElement(this);
	if (!oContentPos)
		return null;

	let oResultPos = oContentPos.Copy();
	this.Get_EndPos(true, oResultPos, oResultPos.GetDepth() + 1);
	return oResultPos;
};
CParagraphContentWithContentBase.prototype.protected_GetPrevRangeEndPos = function(LineIndex, RangeIndex)
{
    var RangeCount  = this.protected_GetRangesCount(LineIndex - 1);
    var RangeOffset = this.protected_GetRangeOffset(LineIndex - 1, RangeCount - 1);

    return LineIndex == 0 && RangeIndex == 0 ? 0 : this.Lines[RangeOffset + 1];
};
CParagraphContentWithContentBase.prototype.updateTrackRevisions = function()
{
	if (this.Paragraph)
		this.Paragraph.updateTrackRevisions();
};
CParagraphContentWithContentBase.prototype.CanSplit = function()
{
	return true;
};
CParagraphContentWithContentBase.prototype.PreDelete = function(isDeep)
{
};
CParagraphContentWithContentBase.prototype.private_UpdateDocumentOutline = function()
{
	if (this.Paragraph)
		this.Paragraph.UpdateDocumentOutline();
};
CParagraphContentWithContentBase.prototype.IsSolid = function()
{
	return false;
};
CParagraphContentWithContentBase.prototype.ProcessNotInlineObjectCheck = function(oChecker)
{
	oChecker.Result = false;
	oChecker.Found  = true;
};
CParagraphContentWithContentBase.prototype.OnContentChange = function()
{
	let oParent = this.GetParent();
	if (oParent)
	{
		oParent.OnContentChange();
	}
	else
	{
		let oParagraph = this.GetParagraph();
		if (oParagraph)
			oParagraph.OnContentChange();
	}
};
CParagraphContentWithContentBase.prototype.OnTextPrChange = function()
{
	let oParent = this.GetParent();
	if (oParent && oParent.OnTextPrChange)
	{
		oParent.OnTextPrChange();
	}
	else
	{
		let oParagraph = this.GetParagraph();
		if (oParagraph)
			oParagraph.OnTextPrChange();
	}
};

/**
 * Это базовый класс для элементов параграфа, которые сами по себе могут содержать элементы параграфа.
 * @constructor
 * @extends {CParagraphContentWithContentBase}
 */
function CParagraphContentWithParagraphLikeContent()
{
	CParagraphContentWithContentBase.call(this);

    this.Type              = undefined;
    this.Paragraph         = null;                  // Ссылка на родительский класс параграф.
    this.m_oContentChanges = new AscCommon.CContentChanges(); // Список изменений(добавление/удаление элементов)
    this.Content           = [];                    // Содержимое данного элемента.

    this.State             = new CParaRunState();   // Состояние курсора/селекта.
    this.Selection         = this.State.Selection;  // Для более быстрого и более простого обращения к селекту.

    this.NearPosArray      = [];
    this.SearchMarks       = [];
}

CParagraphContentWithParagraphLikeContent.prototype = Object.create(CParagraphContentWithContentBase.prototype);
CParagraphContentWithParagraphLikeContent.prototype.constructor = CParagraphContentWithParagraphLikeContent;

CParagraphContentWithParagraphLikeContent.prototype.Get_Type = function()
{
    return this.Type;
};
CParagraphContentWithParagraphLikeContent.prototype.Copy = function(Selected, oPr)
{
	var NewElement = new this.constructor();

	var StartPos = 0;
	var EndPos   = this.Content.length - 1;

	if (true === Selected && true === this.State.Selection.Use)
	{
		StartPos = this.State.Selection.StartPos;
		EndPos   = this.State.Selection.EndPos;

		if (StartPos > EndPos)
		{
			StartPos = this.State.Selection.EndPos;
			EndPos   = this.State.Selection.StartPos;
		}
	}

	let newElementPos = 0;
	for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
	{
		let newItems = this.Content[CurPos].Copy(Selected && (StartPos === CurPos || EndPos === CurPos), oPr);
		if (Array.isArray(newItems))
		{
			for (let newIndex = 0, newCount = newItems.length; newIndex < newCount; ++newIndex)
			{
				NewElement.AddToContent(newElementPos++, newItems[newIndex]);
			}
		}
		else if (newItems)
		{
			NewElement.AddToContent(newElementPos++, newItems);
		}
	}

	return NewElement;
};
CParagraphContentWithParagraphLikeContent.prototype.IsPlaceHolder = function()
{
	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.GetSelectedContent = function(oSelectedContent)
{
	var oNewElement = new this.constructor();

	var nStartPos = this.State.Selection.StartPos;
	var nEndPos   = this.State.Selection.EndPos;

	if (nStartPos > nEndPos)
	{
		nStartPos = this.State.Selection.EndPos;
		nEndPos   = this.State.Selection.StartPos;
	}

	var nItemPos = 0;
	for (var nPos = nStartPos, nItemPos = 0; nPos <= nEndPos; ++nPos)
	{
		var oNewItem = this.Content[nPos].GetSelectedContent(oSelectedContent);
		if (oNewItem)
		{
			oNewElement.AddToContent(nItemPos, oNewItem);
			nItemPos++;
		}
	}

	if (0 === nItemPos)
		return null;

	return oNewElement;
};
CParagraphContentWithParagraphLikeContent.prototype.CopyContent = function(Selected)
{
    var CopyContent = [];

    var StartPos = 0;
    var EndPos = this.Content.length - 1;

    if (true === Selected && true === this.State.Selection.Use)
    {
        StartPos = this.State.Selection.StartPos;
        EndPos   = this.State.Selection.EndPos;
        if (StartPos > EndPos)
        {
            StartPos = this.State.Selection.EndPos;
            EndPos   = this.State.Selection.StartPos;
        }
    }

    for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
    {
        var Item = this.Content[CurPos];

        if ((StartPos === CurPos || EndPos === CurPos) && true !== Item.IsSelectedAll())
        {
            var Content = Item.CopyContent(Selected);
            for (var ContentPos = 0, ContentLen = Content.length; ContentPos < ContentLen; ContentPos++)
            {
                CopyContent.push(Content[ContentPos]);
            }
        }
        else
        {
            CopyContent.push(Item.Copy(false, {CopyReviewPr : true}));
        }
    }

    return CopyContent;
};
CParagraphContentWithParagraphLikeContent.prototype.Clear_ContentChanges = function()
{
    this.m_oContentChanges.Clear();
};
CParagraphContentWithParagraphLikeContent.prototype.Add_ContentChanges = function(Changes)
{
    this.m_oContentChanges.Add( Changes );
};
CParagraphContentWithParagraphLikeContent.prototype.Refresh_ContentChanges = function()
{
    this.m_oContentChanges.Refresh();
};
CParagraphContentWithParagraphLikeContent.prototype.Recalc_RunsCompiledPr = function()
{
    var Count = this.Content.length;
    for (var Pos = 0; Pos < Count; Pos++)
    {
        var Item = this.Content[Pos];

        if (Item.Recalc_RunsCompiledPr)
            Item.Recalc_RunsCompiledPr();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.GetAllDrawingObjects = function(arrDrawingObjects)
{
	if (!arrDrawingObjects)
		arrDrawingObjects = [];

	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		var oItem = this.Content[nPos];

		if (oItem.GetAllDrawingObjects)
			oItem.GetAllDrawingObjects(arrDrawingObjects);
	}

	return arrDrawingObjects;
};
CParagraphContentWithParagraphLikeContent.prototype.SetParagraph = function(Paragraph)
{
	this.Paragraph = Paragraph;

	var ContentLen = this.Content.length;
	for (var CurPos = 0; CurPos < ContentLen; CurPos++)
	{
		this.Content[CurPos].SetParagraph(Paragraph);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.SetParent = function(oParent)
{
	this.Parent = oParent;

	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		if (this.Content[nPos].SetParent)
			this.Content[nPos].SetParent(this);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.SetCurrentPos = function(nPos)
{
	this.State.ContentPos = Math.max(0, Math.min(this.Content.length - 1, nPos));
};
CParagraphContentWithParagraphLikeContent.prototype.Is_Empty = function(oPr)
{
    for (var Index = 0, ContentLen = this.Content.length; Index < ContentLen; Index++)
    {
        if (false === this.Content[Index].Is_Empty(oPr))
            return false;
    }

    return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Is_CheckingNearestPos = function()
{
    return (this.NearPosArray.length > 0);
};
CParagraphContentWithParagraphLikeContent.prototype.IsStartFromNewLine = function()
{
    if (this.Content.length < 0)
        return false;

    return this.Content[0].IsStartFromNewLine();
};
CParagraphContentWithParagraphLikeContent.prototype.GetSelectedElementsInfo = function(oInfo, oContentPos, nDepth)
{
	if (oContentPos)
	{
		var nPos = oContentPos.Get(nDepth);
		if (this.Content[nPos].GetSelectedElementsInfo)
			this.Content[nPos].GetSelectedElementsInfo(oInfo, oContentPos, nDepth + 1);
	}
	else
	{
		if (true === this.Selection.Use && (oInfo.IsCheckAllSelection() || this.Selection.StartPos === this.Selection.EndPos))
		{
			var nStartPos = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
			var nEndPos   = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;

			for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
			{
				this.Content[nPos].GetSelectedElementsInfo(oInfo);
			}
		}
		else if (false === this.Selection.Use)
		{
			this.Content[this.State.ContentPos].GetSelectedElementsInfo(oInfo);
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetSelectedText = function(bAll, bClearText, oPr)
{
    var Str = "";
    for (var Pos = 0, Count = this.Content.length; Pos < Count; Pos++)
    {
        var _Str = this.Content[Pos].GetSelectedText(bAll, bClearText, oPr);

        if (null === _Str)
            return null;

        Str += _Str;
    }

    return Str;
};
CParagraphContentWithParagraphLikeContent.prototype.GetSelectDirection = function()
{
    if (true !== this.Selection.Use)
        return 0;

    if (this.Selection.StartPos < this.Selection.EndPos)
        return 1;
    else if (this.Selection.StartPos > this.Selection.EndPos)
        return -1;

    return this.Content[this.Selection.StartPos].GetSelectDirection();
};
CParagraphContentWithParagraphLikeContent.prototype.Get_TextPr = function(_ContentPos, Depth)
{
    if ( undefined === _ContentPos )
        return this.Content[0].Get_TextPr();
    else
        return this.Content[_ContentPos.Get(Depth)].Get_TextPr(_ContentPos, Depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.Get_FirstTextPr = function(bByPos)
{
	var oElement = null;
	if (this.Content.length > 0)
	{
		if (true === bByPos)
		{
			if (true === this.Selection.Use)
			{
				if (this.Selection.StartPos > this.Selection.EndPos)
					oElement = this.Content[this.Selection.EndPos];
				else
					oElement = this.Content[this.Selection.StartPos];
			}
			else
			{
				oElement = this.Content[this.State.ContentPos];
			}
		}
		else
		{
			for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
			{
				if (this.Content[nPos].IsCursorPlaceable())
				{
					oElement = this.Content[nPos];
					break;
				}
			}
		}
	}

	if (null !== oElement && undefined !== oElement)
	{
		if (para_Run === oElement.Type)
			return oElement.Get_TextPr();
		else
			return oElement.Get_FirstTextPr();
	}

	return new CTextPr();
};
CParagraphContentWithParagraphLikeContent.prototype.Get_CompiledTextPr = function(Copy)
{
    var TextPr = null;

    if (true === this.State.Selection)
    {
        var StartPos = this.State.Selection.StartPos;
        var EndPos   = this.State.Selection.EndPos;

        if (StartPos > EndPos)
        {
            StartPos = this.State.Selection.EndPos;
            EndPos   = this.State.Selection.StartPos;
        }

        TextPr = this.Content[StartPos].Get_CompiledTextPr(Copy);

        while (null === TextPr && StartPos < EndPos)
        {
            StartPos++;
            TextPr = this.Content[StartPos].Get_CompiledTextPr(Copy);
        }

        for (var CurPos = StartPos + 1; CurPos <= EndPos; CurPos++)
        {
            var CurTextPr = this.Content[CurPos].Get_CompiledPr(false);

            if (null !== CurTextPr)
                TextPr = TextPr.Compare(CurTextPr);
        }
    }
    else
    {
        var CurPos = this.State.ContentPos;

        if (CurPos >= 0 && CurPos < this.Content.length)
            TextPr = this.Content[CurPos].Get_CompiledTextPr(Copy);
    }

    return TextPr;
};
CParagraphContentWithParagraphLikeContent.prototype.Check_Content = function()
{
    // Данная функция запускается при чтении файла. Заглушка, на случай, когда в данном классе ничего не будет
    if (this.Content.length <= 0)
        this.Add_ToContent(0, new ParaRun(), false);
};
CParagraphContentWithParagraphLikeContent.prototype.Add_ToContent = function(Pos, Item, UpdatePosition)
{
    this.Content.splice(Pos, 0, Item);
    this.updateTrackRevisions();
	this.private_UpdateDocumentOutline();
    this.private_CheckUpdateBookmarks([Item]);
	this.private_UpdateSelectionPosOnAdd(Pos, 1);

    if (false !== UpdatePosition)
    {
		// Также передвинем всем метки переносов страниц и строк
        var LinesCount = this.protected_GetLinesCount();
        for (var CurLine = 0; CurLine < LinesCount; CurLine++)
        {
            var RangesCount = this.protected_GetRangesCount(CurLine);
            for (var CurRange = 0; CurRange < RangesCount; CurRange++)
            {
                var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
                var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

                if (StartPos > Pos)
                    StartPos++;

                if (EndPos > Pos)
                    EndPos++;

                this.protected_FillRange(CurLine, CurRange, StartPos, EndPos);
            }

            // Особый случай, когда мы добавляем элемент в самый последний ран
            if (Pos === this.Content.length - 1 && LinesCount - 1 === CurLine)
            {
                this.protected_FillRangeEndPos(CurLine, RangesCount - 1, this.protected_GetRangeEndPos(CurLine, RangesCount - 1) + 1);
            }
        }
    }

    // Обновляем позиции в NearestPos
    var NearPosLen = this.NearPosArray.length;
    for (var Index = 0; Index < NearPosLen; Index++)
    {
        var HyperNearPos = this.NearPosArray[Index];
        var ContentPos = HyperNearPos.NearPos.ContentPos;
        var Depth      = HyperNearPos.Depth;

        if (ContentPos.Data[Depth] >= Pos)
            ContentPos.Data[Depth]++;
    }

    // Обновляем позиции в поиске
    var SearchMarksCount = this.SearchMarks.length;
    for (var Index = 0; Index < SearchMarksCount; Index++)
    {
        var Mark       = this.SearchMarks[Index];
        var ContentPos = (true === Mark.Start ? Mark.SearchResult.StartPos : Mark.SearchResult.EndPos);
        var Depth      = Mark.Depth;

        if (ContentPos.Data[Depth] >= Pos)
            ContentPos.Data[Depth]++;
    }

	if (Item.SetParent)
		Item.SetParent(this);

    if (Item.SetParagraph)
    	Item.SetParagraph(this.GetParagraph());

	this.OnContentChange();
};
CParagraphContentWithParagraphLikeContent.prototype.ConcatContent = function (Items)
{
	let Pos = this.GetElementsCount();
	for (let i = 0; i < Items.length; ++i) {
		this.Add_ToContent(Pos + i, Items[i]);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Remove_FromContent = function(Pos, Count, UpdatePosition)
{
	if (Count <= 0)
		return;

	for (var nIndex = Pos; nIndex < Pos + Count; ++nIndex)
	{
		this.Content[nIndex].PreDelete();
	}

	var DeletedItems = this.Content.slice(Pos, Pos + Count);
	this.Content.splice(Pos, Count);
    this.updateTrackRevisions();
	this.private_UpdateDocumentOutline();
	this.private_CheckUpdateBookmarks(DeletedItems);
	this.private_UpdateSelectionPosOnRemove(Pos, Count);

    if (false !== UpdatePosition)
    {
        // Также передвинем всем метки переносов страниц и строк
        var LinesCount = this.protected_GetLinesCount();
        for (var CurLine = 0; CurLine < LinesCount; CurLine++)
        {
            var RangesCount = this.protected_GetRangesCount(CurLine);
            for (var CurRange = 0; CurRange < RangesCount; CurRange++)
            {
                var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
                var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

                if (StartPos > Pos + Count)
                    StartPos -= Count;
                else if (StartPos > Pos)
                    StartPos = Math.max(0, Pos);

                if (EndPos >= Pos + Count)
                    EndPos -= Count;
                else if (EndPos >= Pos)
                    EndPos = Math.max(0, Pos);

                this.protected_FillRange(CurLine, CurRange, StartPos, EndPos);
            }
        }
    }

    // Обновляем позиции в NearestPos
    var NearPosLen = this.NearPosArray.length;
    for (var Index = 0; Index < NearPosLen; Index++)
    {
        var HyperNearPos = this.NearPosArray[Index];
        var ContentPos = HyperNearPos.NearPos.ContentPos;
        var Depth      = HyperNearPos.Depth;

        if (ContentPos.Data[Depth] > Pos + Count)
            ContentPos.Data[Depth] -= Count;
        else if (ContentPos.Data[Depth] > Pos)
            ContentPos.Data[Depth] = Math.max(0, Pos);
    }

    // Обновляем позиции в поиске
    var SearchMarksCount = this.SearchMarks.length;
    for (var Index = 0; Index < SearchMarksCount; Index++)
    {
        var Mark       = this.SearchMarks[Index];
        var ContentPos = (true === Mark.Start ? Mark.SearchResult.StartPos : Mark.SearchResult.EndPos);
        var Depth      = Mark.Depth;

        if (ContentPos.Data[Depth] > Pos + Count)
            ContentPos.Data[Depth] -= Count;
        else if (ContentPos.Data[Depth] > Pos)
            ContentPos.Data[Depth] = Math.max(0, Pos);
    }

	this.OnContentChange();
};
CParagraphContentWithParagraphLikeContent.prototype.Remove = function(Direction, bOnAddText)
{
	var Selection = this.State.Selection;

	if (true === Selection.Use)
	{
		var StartPos = Selection.StartPos;
		var EndPos   = Selection.EndPos;

		if (StartPos > EndPos)
		{
			StartPos = Selection.EndPos;
			EndPos   = Selection.StartPos;
		}

		var oTextPr = this.IsSelectedAll() ? this.GetDirectTextPr() : null;

		if (StartPos === EndPos)
		{
			if (this.Content[StartPos].IsSolid())
			{
				this.RemoveFromContent(StartPos, 1, true);
			}
			else
			{
				this.Content[StartPos].Remove(Direction, bOnAddText);
				
				let isTextDrag = this.Paragraph && this.Paragraph.LogicDocument ? this.Paragraph.LogicDocument.DragAndDropAction : false;
				if (StartPos !== this.Content.length - 1 && true === this.Content[StartPos].Is_Empty() && (!bOnAddText || isTextDrag))
				{
					this.RemoveFromContent(StartPos, 1, true);
					this.State.ContentPos = StartPos;
					this.Content[StartPos].MoveCursorToStartPos();
				}
				else
				{
					this.State.ContentPos = StartPos;
				}
			}
		}
		else
		{
			if (this.Content[EndPos].IsSolid())
			{
				this.RemoveFromContent(EndPos, 1, true);
			}
			else
			{
				this.Content[EndPos].Remove(Direction, bOnAddText);
				
				let isTextDrag = this.Paragraph && this.Paragraph.LogicDocument ? this.Paragraph.LogicDocument.DragAndDropAction : false;
				if (EndPos !== this.Content.length - 1
					&& true === this.Content[EndPos].Is_Empty()
					&& !(this.Content[StartPos] instanceof AscWord.CInlineLevelSdt)
					&& (!bOnAddText || isTextDrag))
				{
					this.Remove_FromContent(EndPos, 1, true);
				}
			}

			if (this.Paragraph && this.Paragraph.LogicDocument && true === this.Paragraph.LogicDocument.IsTrackRevisions())
			{
				for (var nCurPos = EndPos - 1; nCurPos > StartPos; --nCurPos)
				{
					if (para_Run === this.Content[nCurPos].Type)
					{
						if (para_Run == this.Content[nCurPos].Type && this.Content[nCurPos].CanDeleteInReviewMode())
							this.RemoveFromContent(nCurPos, 1);
						else
							this.Content[nCurPos].SetReviewType(reviewtype_Remove, true);
					}
					else
					{
						this.Content[nCurPos].Remove(Direction, bOnAddText);
						if (this.Content[nCurPos].IsEmpty())
							this.RemoveFromContent(nCurPos, 1);
					}
				}
			}
			else
			{
				for (var CurPos = EndPos - 1; CurPos > StartPos; CurPos--)
				{
					this.Remove_FromContent(CurPos, 1, true);
				}
			}

			if (this.Content[StartPos].IsSolid())
			{
				this.RemoveFromContent(StartPos, 1, true);
			}
			else
			{
				this.Content[StartPos].Remove(Direction, bOnAddText);

				if (true === this.Content[StartPos].Is_Empty())
					this.Remove_FromContent(StartPos, 1, true);
			}
		}
		this.RemoveSelection();

		if (this.Content.length <= 0)
		{
			this.AddToContent(0, new ParaRun(this.GetParagraph(), false));
			this.State.ContentPos = 0;

			if (oTextPr)
				this.Content[0].SetPr(oTextPr);
		}
		else
		{
			this.State.ContentPos = StartPos;
		}
	}
	else
	{
		var ContentPos = this.State.ContentPos;

		if ((true === this.Cursor_Is_Start() || true === this.Cursor_Is_End())
			&& !this.IsEmpty()
			&& (!this.CanPlaceCursorInside()
				|| !(this instanceof CInlineLevelSdt)
				|| (!this.IsComplexForm() && !this.IsTextForm() && !this.IsComboBox())))
		{
			this.SelectAll();
			this.SelectThisElement(1);
		}
		else
		{
			while (false === this.Content[ContentPos].Remove(Direction, bOnAddText))
			{
				if (Direction < 0)
					ContentPos--;
				else
					ContentPos++;

				if (ContentPos < 0 || ContentPos >= this.Content.length)
					break;

				if (Direction < 0)
					this.Content[ContentPos].MoveCursorToEndPos(false);
				else
					this.Content[ContentPos].MoveCursorToStartPos();
				
				// Если после перемещения в следующий элемент появился селект, то мы останавливаем удаление,
				// чтобы пользователь видел, что он удаляет
				if (this.Content[ContentPos].IsSelectionUse())
					return true;
			}

			if (ContentPos < 0 || ContentPos >= this.Content.length)
				return false;
			else
			{
				if (ContentPos !== this.Content.length - 1 && true === this.Content[ContentPos].Is_Empty() && true !== bOnAddText)
					this.Remove_FromContent(ContentPos, 1, true);

				this.State.ContentPos = ContentPos;
			}
		}
	}

	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.GetCurrentParaPos = function(align)
{
    var CurPos = this.State.ContentPos;

    if (CurPos >= 0 && CurPos < this.Content.length)
        return this.Content[CurPos].GetCurrentParaPos(align);

    return new CParaPos(this.StartRange, this.StartLine, 0, 0);
};
CParagraphContentWithParagraphLikeContent.prototype.Apply_TextPr = function(TextPr, IncFontSize, ApplyToAll)
{
    if ( true === ApplyToAll )
    {
        var ContentLen = this.Content.length;
        for ( var CurPos = 0; CurPos < ContentLen; CurPos++ )
        {
            this.Content[CurPos].Apply_TextPr( TextPr, IncFontSize, true );
        }
    }
    else
    {
        var Selection = this.State.Selection;

        if ( true === Selection.Use )
        {
            var StartPos = Selection.StartPos;
            var EndPos   = Selection.EndPos;

            if ( StartPos === EndPos )
            {
                var NewElements = this.Content[EndPos].Apply_TextPr( TextPr, IncFontSize, false );

                if ( para_Run === this.Content[EndPos].Type )
                {
                    var CenterRunPos = this.private_ReplaceRun( EndPos, NewElements );

                    if ( StartPos === this.State.ContentPos )
                        this.State.ContentPos = CenterRunPos;

                    // Подправим метки селекта
                    Selection.StartPos = CenterRunPos;
                    Selection.EndPos   = CenterRunPos;
                }
            }
            else
            {
                var Direction = 1;
                if ( StartPos > EndPos )
                {
                    var Temp = StartPos;
                    StartPos = EndPos;
                    EndPos = Temp;

                    Direction = -1;
                }

                for ( var CurPos = StartPos + 1; CurPos < EndPos; CurPos++ )
                {
                    this.Content[CurPos].Apply_TextPr( TextPr, IncFontSize, false );
                }


                var NewElements = this.Content[EndPos].Apply_TextPr( TextPr, IncFontSize, false );
                if ( para_Run === this.Content[EndPos].Type )
                    this.private_ReplaceRun( EndPos, NewElements );

                var NewElements = this.Content[StartPos].Apply_TextPr( TextPr, IncFontSize, false );
                if ( para_Run === this.Content[StartPos].Type )
                    this.private_ReplaceRun( StartPos, NewElements );

                // Подправим селект. Заметим, что метки выделения изменяются внутри функции Add_ToContent
                // за счет того, что EndPos - StartPos > 1.
                if ( Selection.StartPos < Selection.EndPos && true === this.Content[Selection.StartPos].IsSelectionEmpty() )
                    Selection.StartPos++;
                else if ( Selection.EndPos < Selection.StartPos && true === this.Content[Selection.EndPos].IsSelectionEmpty() )
                    Selection.EndPos++;

                if ( Selection.StartPos < Selection.EndPos && true === this.Content[Selection.EndPos].IsSelectionEmpty() )
                    Selection.EndPos--;
                else if ( Selection.EndPos < Selection.StartPos && true === this.Content[Selection.StartPos].IsSelectionEmpty() )
                    Selection.StartPos--;
            }
        }
        else
        {
            var Pos = this.State.ContentPos;
            var Element = this.Content[Pos];
            var NewElements = Element.Apply_TextPr( TextPr, IncFontSize, false );

            if ( para_Run === Element.Type )
            {
                var CenterRunPos = this.private_ReplaceRun( Pos, NewElements );
                this.State.ContentPos = CenterRunPos;
            }
        }
    }
};
CParagraphContentWithParagraphLikeContent.prototype.private_ReplaceRun = function(Pos, NewRuns)
{
    // По логике, можно удалить Run, стоящий в позиции Pos и добавить все раны, которые не null в массиве NewRuns.
    // Но, согласно работе ParaRun.Apply_TextPr, в массиве всегда идет ровно 3 рана (возможно null). Второй ран
    // всегда не null. Первый не null ран и есть ран, идущий в позиции Pos.

    var LRun = NewRuns[0];
    var CRun = NewRuns[1];
    var RRun = NewRuns[2];

    // CRun - всегда не null
    var CenterRunPos = Pos;

    if (null !== LRun)
    {
        this.Add_ToContent(Pos + 1, CRun, true);
        CenterRunPos = Pos + 1;
    }
    else
    {
        // Если LRun - null, значит CRun - это и есть тот ран который стоит уже в позиции Pos
    }

    if (null !== RRun)
        this.Add_ToContent(CenterRunPos + 1, RRun, true);

    return CenterRunPos;
};
CParagraphContentWithParagraphLikeContent.prototype.Clear_TextPr = function()
{
    var Count = this.Content.length;
    for ( var Index = 0; Index < Count; Index++ )
    {
        var Item = this.Content[Index];
        Item.Clear_TextPr();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Check_NearestPos = function(ParaNearPos, Depth)
{
    var HyperNearPos = new CParagraphElementNearPos();
    HyperNearPos.NearPos = ParaNearPos.NearPos;
    HyperNearPos.Depth   = Depth;

    this.NearPosArray.push(HyperNearPos);
    ParaNearPos.Classes.push(this);

    var CurPos = ParaNearPos.NearPos.ContentPos.Get(Depth);
    this.Content[CurPos].Check_NearestPos(ParaNearPos, Depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.Get_DrawingObjectRun = function(Id)
{
    var Run = null;

    var ContentLen = this.Content.length;
    for ( var CurPos = 0; CurPos < ContentLen; CurPos++ )
    {
        var Element = this.Content[CurPos];
        Run = Element.Get_DrawingObjectRun( Id );
        if (null !== Run)
            return Run;
    }

    return Run;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_DrawingObjectContentPos = function(Id, ContentPos, Depth)
{
    for (var Index = 0, ContentLen = this.Content.length; Index < ContentLen; Index++)
    {
        var Element = this.Content[Index];

        if (true === Element.Get_DrawingObjectContentPos(Id, ContentPos, Depth + 1))
        {
            ContentPos.Update2(Index, Depth);
            return true;
        }
    }

    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.GetRunByElement = function(oRunElement)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		var oResult = this.Content[nPos].GetRunByElement(oRunElement);
		if (oResult)
			return oResult;
	}

	return null;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_Layout = function(DrawingLayout, UseContentPos, ContentPos, Depth)
{
    var CurLine  = DrawingLayout.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? DrawingLayout.Range - this.StartRange : DrawingLayout.Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var CurContentPos = ( true === UseContentPos ? ContentPos.Get(Depth) : -1 );

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Get_Layout(DrawingLayout, ( CurPos === CurContentPos ? true : false ), ContentPos, Depth + 1 );

        if (true === DrawingLayout.Layout)
            return;
    }
};
CParagraphContentWithParagraphLikeContent.prototype.GetNextRunElements = function(oRunElements, isUseContentPos, nDepth)
{
	if (oRunElements.IsEnoughElements())
		return;

	var nCurPos     = true === isUseContentPos ? oRunElements.ContentPos.Get(nDepth) : 0;
	var nContentLen = this.Content.length;
	
	if (nCurPos >= nContentLen)
		return;

	oRunElements.UpdatePos(nCurPos, nDepth);
	this.Content[nCurPos].GetNextRunElements(oRunElements, isUseContentPos, nDepth + 1);

	nCurPos++;

	while (nCurPos < nContentLen)
	{
		if (oRunElements.IsEnoughElements())
			return;

		oRunElements.UpdatePos(nCurPos, nDepth);
		this.Content[nCurPos].GetNextRunElements(oRunElements, false, nDepth + 1);

		nCurPos++;
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetPrevRunElements = function(oRunElements, isUseContentPos, nDepth)
{
	if (oRunElements.IsEnoughElements())
		return;

	var nCurPos = true === isUseContentPos ? oRunElements.ContentPos.Get(nDepth) : this.Content.length - 1;
	
	if (nCurPos < 0)
		return;

	oRunElements.UpdatePos(nCurPos, nDepth);
	this.Content[nCurPos].GetPrevRunElements(oRunElements, isUseContentPos, nDepth + 1);

	nCurPos--;

	while (nCurPos >= 0)
	{
		if (oRunElements.IsEnoughElements())
			return;

		oRunElements.UpdatePos(nCurPos, nDepth);
		this.Content[nCurPos].GetPrevRunElements(oRunElements, false, nDepth + 1);

		nCurPos--;
	}
};
CParagraphContentWithParagraphLikeContent.prototype.CollectDocumentStatistics = function(ParaStats)
{
	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
		this.Content[Index].CollectDocumentStatistics(ParaStats);
};
CParagraphContentWithParagraphLikeContent.prototype.Create_FontMap = function(Map)
{
    var Count = this.Content.length;
    for (var Index = 0; Index < Count; Index++)
        this.Content[Index].Create_FontMap( Map );
};
CParagraphContentWithParagraphLikeContent.prototype.Get_AllFontNames = function(AllFonts)
{
    var Count = this.Content.length;
    for (var Index = 0; Index < Count; Index++)
        this.Content[Index].Get_AllFontNames( AllFonts );
};
CParagraphContentWithParagraphLikeContent.prototype.Clear_TextFormatting = function()
{
    for (var Pos = 0, Count = this.Content.length; Pos < Count; Pos++)
    {
        var Item = this.Content[Pos];
        Item.Clear_TextFormatting();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.CanAddDropCap = function()
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		var bResult = this.Content[nPos].CanAddDropCap();
		if (null !== bResult)
			return bResult;
	}

	return null;
};
CParagraphContentWithParagraphLikeContent.prototype.CheckSelectionForDropCap = function(isUsePos, oEndPos, nDepth)
{
	var nEndPos = isUsePos ? oEndPos.Get(nDepth) : this.Content.length - 1;
	for (var nPos = 0; nPos <= nEndPos; ++nPos)
	{
		if (!this.Content[nPos].CheckSelectionForDropCap(nPos === nEndPos && isUsePos, oEndPos, nDepth + 1))
			return false;
	}

	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_TextForDropCap = function(DropCapText, UseContentPos, ContentPos, Depth)
{
    var EndPos = ( true === UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1 );

    for ( var Pos = 0; Pos <= EndPos; Pos++ )
    {
        this.Content[Pos].Get_TextForDropCap( DropCapText, (true === UseContentPos && Pos === EndPos ? true : false), ContentPos, Depth + 1 );

        if ( true === DropCapText.Mixed && ( true === DropCapText.Check || DropCapText.Runs.length > 0 ) )
            return;
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Get_StartTabsCount = function(TabsCounter)
{
    var ContentLen = this.Content.length;
    for ( var Pos = 0; Pos < ContentLen; Pos++ )
    {
        var Element = this.Content[Pos];
        if ( false === Element.Get_StartTabsCount( TabsCounter ) )
            return false;
    }

    return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Remove_StartTabs = function(TabsCounter)
{
    var ContentLen = this.Content.length;
    for ( var Pos = 0; Pos < ContentLen; Pos++ )
    {
        var Element = this.Content[Pos];
        if ( false === Element.Remove_StartTabs( TabsCounter ) )
            return false;
    }

    return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Document_UpdateInterfaceState = function()
{
    if ( true === this.Selection.Use )
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;
        if (StartPos > EndPos)
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
        {
            var Element = this.Content[CurPos];

            if (true !== Element.IsSelectionEmpty() && Element.Document_UpdateInterfaceState)
                Element.Document_UpdateInterfaceState();
        }
    }
    else
    {
        var Element = this.Content[this.State.ContentPos];
        if (Element.Document_UpdateInterfaceState)
            Element.Document_UpdateInterfaceState();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Split = function(ContentPos, Depth)
{
    var Element = new this.constructor();

    var CurPos = ContentPos.Get(Depth);

    var TextPr = this.Get_TextPr(ContentPos, Depth);

    // Разделяем текущий элемент (возвращается правая, отделившаяся часть, если она null, тогда заменяем
    // ее на пустой ран с заданными настройками).
    var NewElement = this.Content[CurPos].Split( ContentPos, Depth + 1 );

    if ( null === NewElement )
    {
        NewElement = new ParaRun();
        NewElement.Set_Pr( TextPr.Copy() );
    }

    // Теперь делим на три части:
    // 1. До элемента с номером CurPos включительно (оставляем эту часть в исходном параграфе)
    // 2. После элемента с номером CurPos (добавляем эту часть в новый параграф)
    // 3. Новый элемент, полученный после разделения элемента с номером CurPos, который мы
    //    добавляем в начало нового параграфа.

    var NewContent = this.Content.slice( CurPos + 1 );
    this.Remove_FromContent( CurPos + 1, this.Content.length - CurPos - 1, false );

    // Добавляем в новую гиперссылку Right элемент и NewContent
    var Count = NewContent.length;
    for ( var Pos = 0; Pos < Count; Pos++ )
        Element.Add_ToContent( Pos, NewContent[Pos], false );

    Element.Add_ToContent( 0, NewElement, false );

    return Element;
};
CParagraphContentWithParagraphLikeContent.prototype.SplitNoDuplicate = function(oContentPos, nDepth, oNewParagraph)
{
	if (this.IsSolid())
		return;

	var nCurPos = oContentPos.Get(nDepth);

	this.Content[nCurPos].SplitNoDuplicate(oContentPos, nDepth + 1, oNewParagraph);

	var arrNewContent = this.Content.slice(nCurPos + 1);
	this.RemoveFromContent(nCurPos + 1, this.Content.length - nCurPos - 1, false);

	var nNewPos = oNewParagraph.Content.length;
	for (var nPos = 0, nCount = arrNewContent.length; nPos < nCount; ++nPos)
		oNewParagraph.AddToContent(nNewPos + nPos, arrNewContent[nPos], false);
};
CParagraphContentWithParagraphLikeContent.prototype.Get_Text = function(Text)
{
    var ContentLen = this.Content.length;
    for ( var CurPos = 0; CurPos < ContentLen; CurPos++ )
    {
        this.Content[CurPos].Get_Text( Text );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.GetAllPermRangeMarks = function(marks)
{
	if (!marks)
		marks = [];
	
	for (let i = 0, count = this.Content.length; i < count; ++i)
	{
		this.Content[i].GetAllPermRangeMarks(marks);
	}
	
	return marks;
};
CParagraphContentWithParagraphLikeContent.prototype.GetAllParagraphs = function(Props, ParaArray)
{
    var ContentLen = this.Content.length;
    for (var CurPos = 0; CurPos < ContentLen; CurPos++)
    {
        if (this.Content[CurPos].GetAllParagraphs)
            this.Content[CurPos].GetAllParagraphs(Props, ParaArray);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.GetAllTables = function(oProps, arrTables)
{
	if (!arrTables)
		arrTables = [];

	for (var nCurPos = 0, nLen = this.Content.length; nCurPos < nLen; ++nCurPos)
	{
		if (this.Content[nCurPos].GetAllTables)
			this.Content[nCurPos].GetAllTables(oProps, arrTables);
	}

	return arrTables;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_ClassesByPos = function(Classes, ContentPos, Depth)
{
    Classes.push(this);
    var CurPos = ContentPos.Get(Depth);
    if (0 <= CurPos && CurPos <= this.Content.length - 1)
        this.Content[CurPos].Get_ClassesByPos(Classes, ContentPos, Depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.GetContentLength = function()
{
    return this.Content.length;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_Parent = function()
{
    if (!this.Paragraph)
        return null;

    var ContentPos = this.Paragraph.Get_PosByElement(this);
    if (!ContentPos || ContentPos.GetDepth() < 0)
        return null;

    ContentPos.DecreaseDepth(1);
    return this.Paragraph.Get_ElementByPos(ContentPos);
};
CParagraphContentWithParagraphLikeContent.prototype.Get_PosInParent = function(Parent)
{
    var _Parent = (_Parent? Parent : this.Get_Parent());
    if (!_Parent)
        return -1;

    for (var Pos = 0, Count = _Parent.Content.length; Pos < Count; ++Pos)
    {
        if (this === _Parent.Content[Pos])
            return Pos;
    }

    return -1;
};
CParagraphContentWithParagraphLikeContent.prototype.Correct_Content = function()
{
	if (this.Paragraph && !this.Paragraph.CanCorrectContent())
		return;

    if (this.Content.length <= 0)
        this.Add_ToContent(0, new ParaRun(this.GetParagraph(), false));
};
CParagraphContentWithParagraphLikeContent.prototype.CorrectContent = function()
{
	if (this.Paragraph && !this.Paragraph.CanCorrectContent())
		return;

	this.Correct_Content();
};
CParagraphContentWithParagraphLikeContent.prototype.UpdateBookmarks = function(oManager)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].UpdateBookmarks(oManager);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.RemoveTabsForTOC = function(_isTab)
{
	var isTab = _isTab;
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		if (this.Content[nIndex].RemoveTabsForTOC(isTab))
			isTab = true;
	}

	return isTab;
};
CParagraphContentWithParagraphLikeContent.prototype.RemoveAll = function()
{
	this.Remove_FromContent(0, this.Content.length);
};
/**
 * Обновляем позиции курсора и селекта во время добавления элементов
 * @param nPosition {number}
 * @param [nCount=1] {number}
 */
CParagraphContentWithParagraphLikeContent.prototype.private_UpdateSelectionPosOnAdd = function(nPosition, nCount)
{
	if (this.Content.length <= 0)
	{
		this.State.ContentPos   = 0;
		this.Selection.StartPos = 0;
		this.Selection.EndPos   = 0;
		return;
	}

	if (undefined === nCount || null === nCount)
		nCount = 1;

	if (this.State.ContentPos >= nPosition)
		this.State.ContentPos += nCount;

	if (this.Selection.StartPos >= nPosition)
		this.Selection.StartPos += nCount;

	if (this.Selection.EndPos >= nPosition)
		this.Selection.EndPos += nCount;

	this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, this.Selection.StartPos));
	this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, this.Selection.EndPos));
	this.State.ContentPos   = Math.max(0, Math.min(this.Content.length - 1, this.State.ContentPos));
};
/**
 * Обновляем позиции курсора и селекта во время удаления элементов
 * @param nPosition {number}
 * @param nCount {number}
 */
CParagraphContentWithParagraphLikeContent.prototype.private_UpdateSelectionPosOnRemove = function(nPosition, nCount)
{
	if (this.State.ContentPos >= nPosition + nCount)
	{
		this.State.ContentPos -= nCount;
	}
	else if (this.State.ContentPos >= nPosition)
	{
		if (nPosition < this.Content.length)
			this.State.ContentPos = nPosition;
		else if (nPosition > 0)
			this.State.ContentPos = nPosition - 1;
		else
			this.State.ContentPos = 0;
	}

	if (this.Selection.StartPos <= this.Selection.EndPos)
	{
		if (this.Selection.StartPos >= nPosition + nCount)
			this.Selection.StartPos -= nCount;
		else if (this.Selection.StartPos >= nPosition)
			this.Selection.StartPos = nPosition;

		if (this.Selection.EndPos >= nPosition + nCount)
			this.Selection.EndPos -= nCount;
		else if (this.Selection.EndPos >= nPosition)
			this.Selection.StartPos = nPosition - 1;

		if (this.Selection.StartPos > this.Selection.EndPos)
		{
			this.Selection.Use = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;
		}
	}
	else
	{
		if (this.Selection.EndPos >= nPosition + nCount)
			this.Selection.EndPos -= nCount;
		else if (this.Selection.EndPos >= nPosition)
			this.Selection.EndPos = nPosition;

		if (this.Selection.StartPos >= nPosition + nCount)
			this.Selection.StartPos -= nCount;
		else if (this.Selection.StartPos >= nPosition)
			this.Selection.StartPos = nPosition - 1;

		if (this.Selection.EndPos > this.Selection.StartPos)
		{
			this.Selection.Use = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;
		}
	}

	this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, this.Selection.StartPos));
	this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, this.Selection.EndPos));
	this.State.ContentPos   = Math.max(0, Math.min(this.Content.length - 1, this.State.ContentPos));
};
CParagraphContentWithParagraphLikeContent.prototype.AddToContent = function(nPos, oItem, isUpdatePositions)
{
	return this.Add_ToContent(nPos, oItem, isUpdatePositions);
};
CParagraphContentWithParagraphLikeContent.prototype.AddToContentToEnd = function(oItem, isUpdatePositions)
{
	return this.Add_ToContent(this.GetElementsCount(), oItem, isUpdatePositions);
};
CParagraphContentWithParagraphLikeContent.prototype.RemoveFromContent = function(nPos, nCount, isUpdatePositions)
{
	return this.Remove_FromContent(nPos, nCount, isUpdatePositions);
};
CParagraphContentWithParagraphLikeContent.prototype.GetComplexField = function(nType)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		var oResult = this.Content[nIndex].GetComplexField(nType);
		if (oResult)
			return oResult;
	}
	return null;
};
CParagraphContentWithParagraphLikeContent.prototype.GetComplexFieldsArray = function(nType, arrComplexFields)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetComplexFieldsArray(nType, arrComplexFields);
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Функции пересчета
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_Range = function(PRS, ParaPr, Depth)
{
    if ( this.Paragraph !== PRS.Paragraph )
    {
        this.Paragraph = PRS.Paragraph;
        this.private_UpdateSpellChecking();
    }

    var CurLine  = PRS.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PRS.Range - this.StartRange : PRS.Range );

    // Добавляем информацию о новом отрезке
    var RangeStartPos = this.protected_AddRange(CurLine, CurRange);
    var RangeEndPos   = 0;

    var ContentLen = this.Content.length;
    var Pos = RangeStartPos;
    for ( ; Pos < ContentLen; Pos++ )
    {
        var Item = this.Content[Pos];

		if (para_Math === Item.Type)
		{
			Item.Set_Inline(!this.CheckMathPara(Pos));
		}

        if ( ( 0 === Pos && 0 === CurLine && 0 === CurRange ) || Pos !== RangeStartPos )
        {
            Item.Recalculate_Reset( PRS.Range, PRS.Line );
        }

        PRS.Update_CurPos( Pos, Depth );
        Item.Recalculate_Range( PRS, ParaPr, Depth + 1 );

        if ( true === PRS.NewRange )
        {
            RangeEndPos = Pos;
            break;
        }
    }

    if ( Pos >= ContentLen )
    {
        RangeEndPos = Pos - 1;
    }

    this.protected_FillRange(CurLine, CurRange, RangeStartPos, RangeEndPos);
};
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_Set_RangeEndPos = function(PRS, PRP, Depth)
{
    var CurLine  = PRS.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PRS.Range - this.StartRange : PRS.Range );
    var CurPos   = PRP.Get(Depth);

    this.protected_FillRangeEndPos(CurLine, CurRange, CurPos);

    this.Content[CurPos].Recalculate_Set_RangeEndPos( PRS, PRP, Depth + 1 );
};
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_SetRangeBounds = function(_CurLine, _CurRange, oStartPos, oEndPos, nDepth)
{
	let isStartPos = oStartPos && nDepth <= oStartPos.GetDepth();
	let isEndPos   = oEndPos && nDepth <= oEndPos.GetDepth();

	let nStartPos = isStartPos ?  oStartPos.Get(nDepth) : 0;
	let nEndPos   = isEndPos ? oEndPos.Get(nDepth) : this.Content.length - 1;

	var CurLine  = _CurLine - this.StartLine;
	var CurRange = 0 === CurLine ? _CurRange - this.StartRange : _CurRange;

	if (isStartPos)
	{
		this.protected_FillRangeEndPos(CurLine, CurRange, nEndPos);
	}
	else
	{
		this.protected_AddRange(CurLine, CurRange);
		this.protected_FillRange(CurLine, CurRange, nStartPos, nEndPos);
	}

	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		let oItem = this.Content[nPos];
		if (nPos !== nStartPos)
			oItem.Recalculate_Reset(_CurRange, _CurLine);

		oItem.Recalculate_SetRangeBounds(_CurLine, _CurRange, nPos === nStartPos ? oStartPos : null, nPos === nEndPos ? oEndPos : null, nDepth + 1);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetContentWidthInRange = function(oStartPos, oEndPos, nDepth)
{
	let nWidth = 0;

	let nStartPos = oStartPos && nDepth <= oStartPos.GetDepth() ? oStartPos.Get(nDepth) : 0;
	let nEndPos   = oEndPos && nDepth <= oEndPos.GetDepth() ? oEndPos.Get(nDepth) : this.Content.length - 1;

	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		nWidth += this.Content[nPos].GetContentWidthInRange(nPos === nStartPos ? oStartPos : null, nPos === nEndPos ? oEndPos : null, nDepth + 1);
	}

	return nWidth;
};
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_LineMetrics = function(PRS, ParaPr, _CurLine, _CurRange)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = (0 === CurLine ? _CurRange - this.StartRange : _CurRange);

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
    {
        this.Content[CurPos].Recalculate_LineMetrics(PRS, ParaPr, _CurLine, _CurRange);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_Range_Width = function(PRSC, _CurLine, _CurRange)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Recalculate_Range_Width( PRSC, _CurLine, _CurRange );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_Range_Spaces = function(PRSA, _CurLine, _CurRange, _CurPage)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Recalculate_Range_Spaces( PRSA, _CurLine, _CurRange, _CurPage );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Recalculate_PageEndInfo = function(PRSI, _CurLine, _CurRange)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Recalculate_PageEndInfo( PRSI, _CurLine, _CurRange );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.RecalculateEndInfo = function(oPRSI)
{
	for (var nCurPos = 0, nCount = this.Content.length; nCurPos < nCount; ++nCurPos)
	{
		this.Content[nCurPos].RecalculateEndInfo(oPRSI);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.SaveRecalculateObject = function(Copy)
{
	var RecalcObj = new CRunRecalculateObject(this.StartLine, this.StartRange);
	RecalcObj.Save_Lines(this, Copy);
	RecalcObj.Save_Content(this, Copy);
	return RecalcObj;
};
CParagraphContentWithParagraphLikeContent.prototype.LoadRecalculateObject = function(RecalcObj)
{
    RecalcObj.Load_Lines( this );
    RecalcObj.Load_Content( this );
};
CParagraphContentWithParagraphLikeContent.prototype.PrepareRecalculateObject = function()
{
	this.protected_ClearLines();

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].PrepareRecalculateObject();
	}
};
CParagraphContentWithParagraphLikeContent.prototype.IsEmptyRange = function(_CurLine, _CurRange)
{
	var CurLine  = _CurLine - this.StartLine;
	var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

	var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
	var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

	for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
	{
		if (false === this.Content[CurPos].IsEmptyRange(_CurLine, _CurRange))
			return false;
	}

	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Check_Range_OnlyMath = function(Checker, _CurRange, _CurLine)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Check_Range_OnlyMath(Checker, _CurRange, _CurLine);

        if (false === Checker.Result)
            break;
    }
};
/**
 * Проверяем является ли элемент в заданной позиции неинлайновой формулой
 * @param {number} nMathPos
 * @return {boolean}
 */
CParagraphContentWithParagraphLikeContent.prototype.CheckMathPara = function(nMathPos)
{
	if (!this.Content[nMathPos] || para_Math !== this.Content[nMathPos].Type)
		return false;

	return this.CheckNotInlineObject(nMathPos);
};
CParagraphContentWithParagraphLikeContent.prototype.CheckNotInlineObject = function(nMathPos, nDirection)
{
	var oParent = this.GetParent();

	var oChecker = new CParagraphMathParaChecker();
	if (undefined === nDirection || -1 === nDirection)
	{
		oChecker.SetDirection(-1);
		for (var nCurPos = nMathPos - 1; nCurPos >= 0; --nCurPos)
		{
			this.Content[nCurPos].ProcessNotInlineObjectCheck(oChecker);
			if (oChecker.IsStop())
				break;
		}

		if (!oChecker.GetResult())
			return false;


		if (!oChecker.IsStop() && oParent && !oParent.CheckNotInlineObject(this.GetPosInParent(oParent), -1))
			return false
	}

	if (undefined === nDirection || 1 === nDirection)
	{
		oChecker.SetDirection(1);
		for (var nCurPos = nMathPos + 1, nCount = this.Content.length; nCurPos < nCount; ++nCurPos)
		{
			this.Content[nCurPos].ProcessNotInlineObjectCheck(oChecker);
			if (oChecker.IsStop())
				break;
		}

		if (!oChecker.GetResult())
			return false;

		if (!oChecker.IsStop() && oParent && !oParent.CheckNotInlineObject(this.GetPosInParent(oParent), 1))
			return false;
	}

	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Check_PageBreak = function()
{
    var Count = this.Content.length;
    for (var Pos = 0; Pos < Count; Pos++)
    {
        if (true === this.Content[Pos].Check_PageBreak())
            return true;
    }

    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.CheckSplitPageOnPageBreak = function(oPBChecker)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		if (this.Content[nPos].CheckSplitPageOnPageBreak(oPBChecker))
			return true;
	}

	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_ParaPosByContentPos = function(ContentPos, Depth)
{
    var Pos = ContentPos.Get(Depth);

    return this.Content[Pos].Get_ParaPosByContentPos( ContentPos, Depth + 1 );
};
CParagraphContentWithParagraphLikeContent.prototype.recalculateCursorPosition = function(positionCalculator, isCurrent)
{
	let rangePos = this.getRangePos(positionCalculator.line, positionCalculator.range);
	let startPos = rangePos[0];
	let endPos   = rangePos[1];
	
	for (let pos = startPos; pos <= endPos; ++pos)
	{
		let item = this.Content[pos];
		item.recalculateCursorPosition(positionCalculator, isCurrent && pos === this.State.ContentPos);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Refresh_RecalcData = function(Data)
{
    if (undefined !== this.Paragraph && null !== this.Paragraph)
        this.Paragraph.Refresh_RecalcData2(0);
};
CParagraphContentWithParagraphLikeContent.prototype.RecalculateMinMaxContentWidth = function(MinMax)
{
	var Count = this.Content.length;
	for (var Pos = 0; Pos < Count; Pos++)
	{
		this.Content[Pos].RecalculateMinMaxContentWidth(MinMax);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Get_Range_VisibleWidth = function(RangeW, _CurLine, _CurRange)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Get_Range_VisibleWidth(RangeW, _CurLine, _CurRange);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Shift_Range = function(Dx, Dy, _CurLine, _CurRange, _CurPage)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Shift_Range(Dx, Dy, _CurLine, _CurRange, _CurPage);
    }
};
//----------------------------------------------------------------------------------------------------------------------
// Функции отрисовки
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.Draw_HighLights = function(PDSH)
{
    var CurLine  = PDSH.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PDSH.Range - this.StartRange : PDSH.Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Draw_HighLights( PDSH );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Draw_Elements = function(PDSE)
{
	let textAlpha;
	let placeholderAlpha = this.IsPlaceHolder() && this.IsForm && this.IsForm();
	if (placeholderAlpha)
	{
		textAlpha = PDSE.Graphics.getTextGlobalAlpha();
		PDSE.Graphics.setTextGlobalAlpha(0.5);
	}

	var CurLine  = PDSE.Line - this.StartLine;
	var CurRange = (0 === CurLine ? PDSE.Range - this.StartRange : PDSE.Range);

	var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
	var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

	for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
	{
		this.Content[CurPos].Draw_Elements(PDSE);
	}

	if (placeholderAlpha)
		PDSE.Graphics.setTextGlobalAlpha(textAlpha);
};
CParagraphContentWithParagraphLikeContent.prototype.Draw_Lines = function(PDSL)
{
	var CurLine  = PDSL.Line - this.StartLine;
	var CurRange = ( 0 === CurLine ? PDSL.Range - this.StartRange : PDSL.Range );

	var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
	var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

	var nCurDepth = PDSL.CurDepth;
	for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
	{
		PDSL.CurPos.Update(CurPos, nCurDepth);
		PDSL.CurDepth = nCurDepth + 1;

		this.Content[CurPos].Draw_Lines(PDSL);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.SkipDraw = function(PDS)
{
	var CurLine  = PDS.Line - this.StartLine;
	var CurRange = (0 === CurLine ? PDS.Range - this.StartRange : PDS.Range);

	var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
	var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

	for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
	{
		this.Content[CurPos].SkipDraw(PDS);
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с курсором
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.IsCursorPlaceable = function()
{
    return true;
};
CParagraphContentWithParagraphLikeContent.prototype.CanPlaceCursorInside = function()
{
	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.IsCursorAtBegin = function()
{
	if (this.IsPlaceHolder())
		return true;

	return this.Cursor_Is_Start();
};
CParagraphContentWithParagraphLikeContent.prototype.IsCursorAtEnd = function()
{
	if (this.IsPlaceHolder())
		return true;

	return this.Cursor_Is_End();
};
CParagraphContentWithParagraphLikeContent.prototype.Cursor_Is_Start = function()
{
    var CurPos = 0;
    while ( CurPos < this.State.ContentPos && CurPos < this.Content.length - 1 )
    {
        if ( true === this.Content[CurPos].Is_Empty() )
            CurPos++;
        else
            return false;
    }

    return this.Content[CurPos].Cursor_Is_Start();
};
CParagraphContentWithParagraphLikeContent.prototype.Cursor_Is_NeededCorrectPos = function()
{
    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Cursor_Is_End = function()
{
    var CurPos = this.Content.length - 1;
    while ( CurPos > this.State.ContentPos && CurPos > 0 )
    {
        if ( true === this.Content[CurPos].Is_Empty() )
            CurPos--;
        else
            return false;
    }

    return this.Content[CurPos].Cursor_Is_End();
};
CParagraphContentWithParagraphLikeContent.prototype.IsStartPos = function(contentPos, depth)
{
	if (depth >= contentPos.Depth)
		return true;
	
	let pos = contentPos.Get(depth);
	if (!this.Content[pos])
		return false;
	
	if (pos !== 0)
		return false;
	
	return this.Content[pos].IsStartPos(contentPos, depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.IsEndPos = function(contentPos, depth)
{
	if (depth >= contentPos.Depth)
		return true;
	
	let pos = contentPos.Get(depth);
	if (!this.Content[pos])
		return true;
	
	if (pos !== this.Content.length - 1)
		return false;
	
	return this.Content[pos].IsEndPos(contentPos, depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.MoveCursorToStartPos = function()
{
    this.State.ContentPos = 0;

    if ( this.Content.length > 0 )
    {
        this.Content[0].MoveCursorToStartPos();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.MoveCursorToEndPos = function(SelectFromEnd)
{
    var ContentLen = this.Content.length;

    if ( ContentLen > 0 )
    {
        this.State.ContentPos = ContentLen - 1;
        this.Content[ContentLen - 1].MoveCursorToEndPos( SelectFromEnd );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.getParagraphContentPosByXY = function(searchState)
{
	let rangePos = this.getRangePos(searchState.line, searchState.range);
	let startPos = rangePos[0];
	let endPos   = rangePos[1];
	for (let pos = startPos; pos <= endPos; ++pos)
	{
		this.Content[pos].getParagraphContentPosByXY(searchState);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Get_ParaContentPos = function(bSelection, bStart, ContentPos, bUseCorrection)
{
	var Pos = ( true === bSelection ? ( true === bStart ? this.State.Selection.StartPos : this.State.Selection.EndPos ) : this.State.ContentPos );
	ContentPos.Add(Pos);

	if (Pos < 0 || Pos >= this.Content.length)
		return;

	this.Content[Pos].Get_ParaContentPos(bSelection, bStart, ContentPos, bUseCorrection);
};
CParagraphContentWithParagraphLikeContent.prototype.Set_ParaContentPos = function(ContentPos, Depth)
{
    var Pos = ContentPos.Get(Depth);

    if ( Pos >= this.Content.length )
        Pos = this.Content.length - 1;

    if ( Pos < 0 )
        Pos = 0;

    this.State.ContentPos = Pos;

    this.Content[Pos].Set_ParaContentPos( ContentPos, Depth + 1 );
};
CParagraphContentWithParagraphLikeContent.prototype.Get_PosByElement = function(Class, ContentPos, Depth, UseRange, Range, Line)
{
    if ( this === Class )
        return true;

    if (this.Content.length <= 0)
    	return false;

    var StartPos = 0;
    var EndPos   = this.Content.length - 1;

    if ( true === UseRange )
    {
        var CurLine  = Line - this.StartLine;
        var CurRange = ( 0 === CurLine ? Range - this.StartRange : Range );

        if (CurLine >= 0 && CurLine < this.protected_GetLinesCount() && CurRange >= 0 && CurRange < this.protected_GetRangesCount(CurLine))
        {
            StartPos = Math.min(this.Content.length - 1, Math.max(0, this.protected_GetRangeStartPos(CurLine, CurRange)));
            EndPos   = Math.min(this.Content.length - 1, Math.max(0, this.protected_GetRangeEndPos(CurLine, CurRange)));
        }
    }

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        var Element = this.Content[CurPos];

        ContentPos.Update( CurPos, Depth );

        if ( true === Element.Get_PosByElement(Class, ContentPos, Depth + 1, true, CurRange, CurLine) )
            return true;
    }

    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_ElementByPos = function(ContentPos, Depth)
{
    if (Depth >= ContentPos.Depth)
        return this;

    var CurPos = ContentPos.Get(Depth);
    if (!this.Content[CurPos])
    	return null;

    return this.Content[CurPos].Get_ElementByPos(ContentPos, Depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.ConvertParaContentPosToRangePos = function(oContentPos, nDepth)
{
	var nRangePos = 0;

	var nCurPos = oContentPos ? Math.max(0, Math.min(this.Content.length - 1, oContentPos.Get(nDepth))) : this.Content.length - 1;
	for (var nPos = 0; nPos < nCurPos; ++nPos)
	{
		if (this.Content[nPos] instanceof ParaRun)
			nRangePos++;

		nRangePos += this.Content[nPos].ConvertParaContentPosToRangePos(null);
	}

	if (this.Content[nCurPos])
	{
		if (this.Content[nPos] instanceof ParaRun)
			nRangePos++;

		nRangePos += this.Content[nCurPos].ConvertParaContentPosToRangePos(oContentPos, nDepth + 1);
	}
		
	return nRangePos;
};
CParagraphContentWithParagraphLikeContent.prototype.GetPosByDrawing = function(Id, ContentPos, Depth)
{
    var Count = this.Content.length;
    for ( var CurPos = 0; CurPos < Count; CurPos++ )
    {
        var Element = this.Content[CurPos];

        ContentPos.Update( CurPos, Depth );

        if ( true === Element.GetPosByDrawing(Id, ContentPos, Depth + 1) )
            return true;
    }

    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_RunElementByPos = function(ContentPos, Depth)
{
    if ( undefined !== ContentPos )
    {
        var Pos = ContentPos.Get(Depth);

        return this.Content[Pos].Get_RunElementByPos( ContentPos, Depth + 1 );
    }
    else
    {
        var Count = this.Content.length;
        if ( Count <= 0 )
            return null;

        var Pos = 0;
        var Element = this.Content[Pos];

        while ( null === Element && Pos < Count - 1 )
            Element = this.Content[++Pos];

        return Element;
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Get_LastRunInRange = function(_CurLine, _CurRange)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    if (CurLine < this.protected_GetLinesCount() && CurRange < this.protected_GetRangesCount(CurLine))
    {
        var LastItem = this.Content[this.protected_GetRangeEndPos(CurLine, CurRange)];
        if ( undefined !== LastItem )
            return LastItem.Get_LastRunInRange(_CurLine, _CurRange);
    }

    return null;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
	if (this.Content.length <= 0)
		return false;

	var CurPos = ( true === UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1 );

	SearchPos.Pos.Update(CurPos, Depth);
	this.Content[CurPos].Get_LeftPos(SearchPos, ContentPos, Depth + 1, UseContentPos);

	if (true === SearchPos.Found)
		return true;

	CurPos--;

	if (CurPos >= 0 && this.Content[CurPos + 1].IsStopCursorOnEntryExit())
	{
		SearchPos.Pos.Update(CurPos, Depth);
		this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
		SearchPos.Found = true;
		return true;
	}

	while (CurPos >= 0)
	{
		SearchPos.Pos.Update(CurPos, Depth);
		this.Content[CurPos].Get_LeftPos(SearchPos, ContentPos, Depth + 1, false);

		if (true === SearchPos.Found)
			return true;

		CurPos--;
	}

	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
	if (this.Content.length <= 0)
		return false;

	var CurPos = ( true === UseContentPos ? ContentPos.Get(Depth) : 0 );

	SearchPos.Pos.Update(CurPos, Depth);
	this.Content[CurPos].Get_RightPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepEnd);

	if (true === SearchPos.Found)
		return true;

	CurPos++;

	var Count = this.Content.length;
	if (CurPos < Count && this.Content[CurPos - 1].IsStopCursorOnEntryExit())
	{
		SearchPos.Pos.Update(CurPos, Depth);
		this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
		SearchPos.Found = true;
		return true;
	}

	while (CurPos < this.Content.length)
	{
		SearchPos.Pos.Update(CurPos, Depth);
		this.Content[CurPos].Get_RightPos(SearchPos, ContentPos, Depth + 1, false, StepEnd);

		if (true === SearchPos.Found)
			return true;

		CurPos++;
	}

	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
	var CurPos = ( true === UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1 );

	this.Content[CurPos].Get_WordStartPos(SearchPos, ContentPos, Depth + 1, UseContentPos);

	if (true === SearchPos.UpdatePos)
		SearchPos.Pos.Update2(CurPos, Depth);

	if (true === SearchPos.Found)
		return;

	CurPos--;

	if (SearchPos.Shift && CurPos >= 0 && this.Content[CurPos].IsStopCursorOnEntryExit())
	{
		SearchPos.Found = true;
		return;
	}

	if (CurPos >= 0 && this.Content[CurPos + 1].IsStopCursorOnEntryExit())
	{
		this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update2(CurPos, Depth);
		SearchPos.Found = true;
		return;
	}

	while (CurPos >= 0)
	{
		var OldUpdatePos = SearchPos.UpdatePos;

		this.Content[CurPos].Get_WordStartPos(SearchPos, ContentPos, Depth + 1, false);

		if (true === SearchPos.UpdatePos)
			SearchPos.Pos.Update2(CurPos, Depth);
		else
			SearchPos.UpdatePos = OldUpdatePos;

		if (true === SearchPos.Found)
			return;

		CurPos--;

		if (SearchPos.Shift && CurPos >= 0 && this.Content[CurPos].IsStopCursorOnEntryExit())
		{
			SearchPos.Found = true;
			return;
		}

		if (CurPos >= 0 && this.Content[CurPos + 1].IsStopCursorOnEntryExit())
		{
			this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
			SearchPos.Pos.Update2(CurPos, Depth);
			SearchPos.Found = true;
			return;
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
	var CurPos = ( true === UseContentPos ? ContentPos.Get(Depth) : 0 );

	this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepEnd);

	if (true === SearchPos.UpdatePos)
		SearchPos.Pos.Update(CurPos, Depth);

	if (true === SearchPos.Found)
		return;

	CurPos++;

	var Count = this.Content.length;

	if (SearchPos.Shift && CurPos < Count && this.Content[CurPos].IsStopCursorOnEntryExit())
	{
		SearchPos.Found = true;
		return;
	}

	if (CurPos < Count && this.Content[CurPos - 1].IsStopCursorOnEntryExit())
	{
		this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(CurPos, Depth);
		SearchPos.Found     = true;
		SearchPos.UpdatePos = true;
		return;
	}

	while (CurPos < Count)
	{
		var OldUpdatePos = SearchPos.UpdatePos;

		this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, false, StepEnd);

		if (true === SearchPos.UpdatePos)
			SearchPos.Pos.Update(CurPos, Depth);
		else
			SearchPos.UpdatePos = OldUpdatePos;

		if (true === SearchPos.Found)
			return;

		CurPos++;

		if (SearchPos.Shift && CurPos < Count && this.Content[CurPos].IsStopCursorOnEntryExit())
		{
			SearchPos.Found = true;
			return;
		}

		if (CurPos < Count && this.Content[CurPos - 1].IsStopCursorOnEntryExit())
		{
			this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
			SearchPos.Pos.Update(CurPos, Depth);
			SearchPos.Found     = true;
			SearchPos.UpdatePos = true;
			return;
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Get_EndRangePos = function(nCurLine, nCurRange, oSearchPos, nDepth)
{
	var _nCurLine  = nCurLine - this.StartLine;
	var _nCurRange = (0 === _nCurLine ? nCurRange - this.StartRange : nCurRange);

	var nStartPos = Math.max(0, Math.min(this.Content.length - 1, this.protected_GetRangeStartPos(_nCurLine, _nCurRange)));
	var nEndPos   = Math.min(this.Content.length - 1, Math.max(0, this.protected_GetRangeEndPos(_nCurLine, _nCurRange)));

	var bResult = false;
	for (var nPos = nEndPos; nPos >= nStartPos; --nPos)
	{
		if (this.Content[nPos].Get_EndRangePos(nCurLine, nCurRange, oSearchPos, nDepth + 1))
		{
			oSearchPos.Pos.Update(nPos, nDepth);
			bResult = true;
			break;
		}
	}

	return bResult;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_StartRangePos = function(nCurLine, nCurRange, oSearchPos, nDepth)
{
	var _nCurLine  = nCurLine - this.StartLine;
	var _nCurRange = ( 0 === _nCurLine ? nCurRange - this.StartRange : nCurRange );

	var nStartPos = Math.max(0, Math.min(this.Content.length - 1, this.protected_GetRangeStartPos(_nCurLine, _nCurRange)));
	var nEndPos   = Math.min(this.Content.length - 1, Math.max(0, this.protected_GetRangeEndPos(_nCurLine, _nCurRange)));

	var bResult = false;
	for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		if (this.Content[nPos].Get_StartRangePos(nCurLine, nCurRange, oSearchPos, nDepth + 1))
		{
			oSearchPos.Pos.Update(nPos, nDepth);
			bResult = true;
			break;
		}
	}

	return bResult;
};
CParagraphContentWithParagraphLikeContent.prototype.Get_StartRangePos2 = function(_CurLine, _CurRange, ContentPos, Depth)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var Pos = this.protected_GetRangeStartPos(CurLine, CurRange);

    ContentPos.Update( Pos, Depth );

    this.Content[Pos].Get_StartRangePos2( _CurLine, _CurRange, ContentPos, Depth + 1 );
};
CParagraphContentWithParagraphLikeContent.prototype.Get_EndRangePos2 = function(_CurLine, _CurRange, ContentPos, Depth)
{
	var CurLine  = _CurLine - this.StartLine;
	var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

	var Pos = this.protected_GetRangeEndPos(CurLine, CurRange);
	ContentPos.Update(Pos, Depth);
	this.Content[Pos].Get_EndRangePos2(_CurLine, _CurRange, ContentPos, Depth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.Get_StartPos = function(ContentPos, Depth)
{
    if ( this.Content.length > 0 )
    {
        ContentPos.Update( 0, Depth );

        this.Content[0].Get_StartPos( ContentPos, Depth + 1 );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.Get_EndPos = function(BehindEnd, ContentPos, Depth)
{
    var ContentLen = this.Content.length;
    if ( ContentLen > 0 )
    {
        ContentPos.Update( ContentLen - 1, Depth );

        this.Content[ContentLen - 1].Get_EndPos( BehindEnd, ContentPos, Depth + 1 );
    }
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с селектом
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.Set_SelectionContentPos = function(StartContentPos, EndContentPos, Depth, StartFlag, EndFlag)
{
	if (this.Content.length <= 0)
		return;
	
	if (!this.CanPlaceCursorInside())
	{
		if (this.Paragraph && this.Paragraph.GetSelectDirection() > 0)
			this.SelectAll(1);
		else
			this.SelectAll(-1);
		return;
	}

    var Selection = this.Selection;

    var OldStartPos = Selection.StartPos;
    var OldEndPos   = Selection.EndPos;

    if ( OldStartPos > OldEndPos )
    {
        OldStartPos = Selection.EndPos;
        OldEndPos   = Selection.StartPos;
    }

    var StartPos = 0;
    switch (StartFlag)
    {
        case  1: StartPos = 0; break;
        case -1: StartPos = this.Content.length - 1; break;
        case  0: StartPos = StartContentPos.Get(Depth); break;
    }

    var EndPos = 0;
    switch (EndFlag)
    {
        case  1: EndPos = 0; break;
        case -1: EndPos = this.Content.length - 1; break;
        case  0: EndPos = EndContentPos.Get(Depth); break;
    }

    // Удалим отметки о старом селекте
    if ( OldStartPos < StartPos && OldStartPos < EndPos )
    {
        var TempBegin = Math.max(0, OldStartPos);
        var TempEnd   = Math.min(this.Content.length - 1, Math.min(StartPos, EndPos) - 1);
        for (var CurPos = TempBegin; CurPos <= TempEnd; ++CurPos)
        {
            this.Content[CurPos].RemoveSelection();
        }
    }

    if ( OldEndPos > StartPos && OldEndPos > EndPos )
    {
        var TempBegin = Math.max(0, Math.max(StartPos, EndPos) + 1);
        var TempEnd   = Math.min(OldEndPos, this.Content.length - 1);
        for (var CurPos = TempBegin; CurPos <= TempEnd; ++CurPos)
        {
            this.Content[CurPos].RemoveSelection();
        }
    }

    // Выставим метки нового селекта

    Selection.Use      = true;
    Selection.StartPos = StartPos;
    Selection.EndPos   = EndPos;

    if ( StartPos != EndPos )
    {
        this.Content[StartPos].Set_SelectionContentPos( StartContentPos, null, Depth + 1, StartFlag, StartPos > EndPos ? 1 : -1 );
        this.Content[EndPos].Set_SelectionContentPos( null, EndContentPos, Depth + 1, StartPos > EndPos ? -1 : 1, EndFlag );

        var _StartPos = StartPos;
        var _EndPos   = EndPos;
        var Direction = 1;

        if ( _StartPos > _EndPos )
        {
            _StartPos = EndPos;
            _EndPos   = StartPos;
            Direction = -1;
        }

        for ( var CurPos = _StartPos + 1; CurPos < _EndPos; CurPos++ )
        {
            this.Content[CurPos].SelectAll( Direction );
        }
    }
    else
    {
        this.Content[StartPos].Set_SelectionContentPos( StartContentPos, EndContentPos, Depth + 1, StartFlag, EndFlag );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.SetContentSelection = function(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag)
{
	if (this.Content.length <= 0)
		return;

    if ((0 === StartFlag && (!StartDocPos[Depth] || this !== StartDocPos[Depth].Class)) || (0 === EndFlag && (!EndDocPos[Depth] || this !== EndDocPos[Depth].Class)))
        return;

    var StartPos = 0, EndPos = 0;
    switch (StartFlag)
    {
        case 0 : StartPos = StartDocPos[Depth].Position; break;
        case 1 : StartPos = 0; break;
        case -1: StartPos = this.Content.length - 1; break;
    }

    switch (EndFlag)
    {
        case 0 : EndPos = EndDocPos[Depth].Position; break;
        case 1 : EndPos = 0; break;
        case -1: EndPos = this.Content.length - 1; break;
    }

    var _StartDocPos = StartDocPos, _StartFlag = StartFlag;
    if (null !== StartDocPos && true === StartDocPos[Depth].Deleted)
    {
        if (StartPos < this.Content.length)
        {
            _StartDocPos = null;
            _StartFlag = 1;
        }
        else if (StartPos > 0)
        {
            StartPos--;
            _StartDocPos = null;
            _StartFlag = -1;
        }
        else
        {
            // Такого не должно быть
            return;
        }
    }

    var _EndDocPos = EndDocPos, _EndFlag = EndFlag;
    if (null !== EndDocPos && true === EndDocPos[Depth].Deleted)
    {
        if (EndPos < this.Content.length)
        {
            _EndDocPos = null;
            _EndFlag = 1;
        }
        else if (EndPos > 0)
        {
            EndPos--;
            _EndDocPos = null;
            _EndFlag = -1;
        }
        else
        {
            // Такого не должно быть
            return;
        }
    }

    this.Selection.Use      = true;
    this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, StartPos));
    this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, EndPos));

    if (StartPos !== EndPos)
    {
        if (this.Content[StartPos] && this.Content[StartPos].SetContentSelection)
            this.Content[StartPos].SetContentSelection(_StartDocPos, null, Depth + 1, _StartFlag, StartPos > EndPos ? 1 : -1);

        if (this.Content[EndPos] && this.Content[EndPos].SetContentSelection)
            this.Content[EndPos].SetContentSelection(null, _EndDocPos, Depth + 1, StartPos > EndPos ? -1 : 1, _EndFlag);

        var _StartPos = StartPos;
        var _EndPos = EndPos;
        var Direction = 1;

        if (_StartPos > _EndPos)
        {
            _StartPos = EndPos;
            _EndPos = StartPos;
            Direction = -1;
        }

        for (var CurPos = _StartPos + 1; CurPos < _EndPos; CurPos++)
        {
            this.Content[CurPos].SelectAll(Direction);
        }
    }
    else
    {
        if (this.Content[StartPos] && this.Content[StartPos].SetContentSelection)
            this.Content[StartPos].SetContentSelection(_StartDocPos, _EndDocPos, Depth + 1, _StartFlag, _EndFlag);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.SetContentPosition = function(DocPos, Depth, Flag)
{
	if (this.Content.length <= 0)
		return;

    if (0 === Flag && (!DocPos[Depth] || this !== DocPos[Depth].Class))
        return;

    var Pos = 0;
    switch (Flag)
    {
        case 0 : Pos = DocPos[Depth].Position; break;
        case 1 : Pos = 0; break;
        case -1: Pos = this.Content.length - 1; break;
    }

    var _DocPos = DocPos, _Flag = Flag;
    if (null !== DocPos && true === DocPos[Depth].Deleted)
    {
        if (Pos < this.Content.length)
        {
            _DocPos = null;
            _Flag = 1;
        }
        else if (Pos > 0)
        {
            Pos--;
            _DocPos = null;
            _Flag = -1;
        }
        else
        {
            // Такого не должно быть
            return;
        }
    }

    this.State.ContentPos = Math.max(0, Math.min(this.Content.length - 1, Pos));

    // TODO: Как только в CMathContent CurPos перейдет на стандартное this.State.ContentPos убрать эту проверку
    if (this.CurPos)
        this.CurPos = this.State.ContentPos;

    if (this.Content[Pos] && this.Content[Pos].SetContentPosition)
        this.Content[Pos].SetContentPosition(_DocPos, Depth + 1, _Flag);
    else
        this.Content[Pos].MoveCursorToStartPos();
};
CParagraphContentWithParagraphLikeContent.prototype.RemoveSelection = function()
{
    var Selection = this.Selection;

    if ( true === Selection.Use )
    {
        var StartPos = Selection.StartPos;
        var EndPos   = Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = Selection.EndPos;
            EndPos   = Selection.StartPos;
        }

        StartPos = Math.max( 0, StartPos );
        EndPos   = Math.min( this.Content.length - 1, EndPos );

        for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
        {
            this.Content[CurPos].RemoveSelection();
        }
    }

    Selection.Use      = false;
    Selection.StartPos = 0;
    Selection.EndPos   = 0;
};
CParagraphContentWithParagraphLikeContent.prototype.SelectAll = function(Direction)
{
    var ContentLen = this.Content.length;

    var Selection = this.Selection;

    Selection.Use = true;

    if ( -1 === Direction )
    {
        Selection.StartPos = ContentLen - 1;
        Selection.EndPos   = 0;
    }
    else
    {
        Selection.StartPos = 0;
        Selection.EndPos   = ContentLen - 1;
    }

    for ( var CurPos = 0; CurPos < ContentLen; CurPos++ )
    {
        this.Content[CurPos].SelectAll( Direction );
    }
};
CParagraphContentWithParagraphLikeContent.prototype.drawSelectionInRange = function(line, range, drawState)
{
	let rangeInfo  = this.getRangePos(line, range);
	let rangeStart = rangeInfo[0];
	let rangeEnd   = rangeInfo[1];
	for (let pos = rangeStart; pos <= rangeEnd; ++pos)
	{
		this.Content[pos].drawSelectionInRange(line, range, drawState);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.IsSelectionEmpty = function(CheckEnd)
{
	if (this.Content.length <= 0)
		return true;

    var StartPos = this.State.Selection.StartPos;
    var EndPos   = this.State.Selection.EndPos;

    if ( StartPos > EndPos )
    {
        StartPos = this.State.Selection.EndPos;
        EndPos   = this.State.Selection.StartPos;
    }

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        if ( false === this.Content[CurPos].IsSelectionEmpty(CheckEnd) )
            return false;
    }

    return true;
};
CParagraphContentWithParagraphLikeContent.prototype.Selection_CheckParaEnd = function()
{
    // Сюда не должен попадать ParaEnd
    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.Selection_CheckParaContentPos = function(ContentPos, Depth, bStart, bEnd)
{
    var CurPos = ContentPos.Get(Depth);

    if (this.Selection.StartPos <= CurPos && CurPos <= this.Selection.EndPos)
        return this.Content[CurPos].Selection_CheckParaContentPos(ContentPos, Depth + 1, bStart && this.Selection.StartPos === CurPos, bEnd && CurPos === this.Selection.EndPos);
    else if (this.Selection.EndPos <= CurPos && CurPos <= this.Selection.StartPos)
        return this.Content[CurPos].Selection_CheckParaContentPos(ContentPos, Depth + 1, bStart && this.Selection.EndPos === CurPos, bEnd && CurPos === this.Selection.StartPos);

    return false;
};
CParagraphContentWithParagraphLikeContent.prototype.IsSelectedAll = function(Props)
{
    var Selection = this.State.Selection;

    if ( false === Selection.Use && true !== this.Is_Empty( Props ) )
        return false;

    var StartPos = Selection.StartPos;
    var EndPos   = Selection.EndPos;

    if ( EndPos < StartPos )
    {
        StartPos = Selection.EndPos;
        EndPos   = Selection.StartPos;
    }

    for ( var Pos = 0; Pos <= StartPos; Pos++ )
    {
        if ( false === this.Content[Pos].IsSelectedAll( Props ) )
            return false;
    }

    var Count = this.Content.length;
    for ( var Pos = EndPos; Pos < Count; Pos++ )
    {
        if ( false === this.Content[Pos].IsSelectedAll( Props ) )
            return false;
    }

    return true;
};
CParagraphContentWithParagraphLikeContent.prototype.IsSelectedFromStart = function()
{
	if (!this.Selection.Use && !this.IsEmpty())
		return false;

	var nStartPos = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
	return this.Content[nStartPos].IsSelectedFromStart();
};
CParagraphContentWithParagraphLikeContent.prototype.IsSelectedToEnd = function()
{
	if (!this.Selection.Use && !this.IsEmpty())
		return false;

	var nEndPos = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
	return this.Content[nEndPos].IsSelectedToEnd();
};

CParagraphContentWithParagraphLikeContent.prototype.SkipAnchorsAtSelectionStart = function(nDirection)
{
	if (false === this.Selection.Use || true === this.IsEmpty({SkipAnchor : true}))
		return true;

	var oSelection = this.State.Selection;
	var nStartPos  = Math.min(oSelection.StartPos, oSelection.EndPos);
	var nEndPos    = Math.max(oSelection.StartPos, oSelection.EndPos);

	for (var nPos = 0; nPos < nStartPos; ++nPos)
	{
		if (true !== this.Content[nPos].IsEmpty({SkipAnchor : true}))
			return false;
	}

	for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		if (true === this.Content[nPos].SkipAnchorsAtSelectionStart(nDirection))
		{
			if (1 === nDirection)
				this.Selection.StartPos = nPos + 1;
			else
				this.Selection.EndPos = nPos + 1;

			this.Content[nPos].RemoveSelection();
		}
		else
		{
			return false;
		}
	}

	if (nEndPos < this.Content.length - 1)
		return false;

	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.IsSelectionUse = function()
{
    return this.State.Selection.Use;
};
//----------------------------------------------------------------------------------------------------------------------
// SpellCheck
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.RestartSpellCheck = function()
{
    for (let nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
    {
        this.Content[nIndex].RestartSpellCheck();
    }
};
/**
 * @param oSpellCheckerEngine {AscWord.CParagraphSpellCheckerCollector}
 * @param nDepth {number}
 */
CParagraphContentWithParagraphLikeContent.prototype.CheckSpelling = function(oSpellCheckerEngine, nDepth)
{
	if (oSpellCheckerEngine.IsExceedLimit())
		return;

	var nStartPos = 0;
	if (oSpellCheckerEngine.IsFindStart())
		nStartPos = oSpellCheckerEngine.GetPos(nDepth);

    for (var nPos = nStartPos, nCount = this.Content.length; nPos < nCount; ++nPos)
    {
    	var oItem = this.Content[nPos];

    	oSpellCheckerEngine.UpdatePos(nPos, nDepth);
        oItem.CheckSpelling(oSpellCheckerEngine, nDepth + 1);

        if (oSpellCheckerEngine.IsExceedLimit())
        	return;
    }
};
//----------------------------------------------------------------------------------------------------------------------
// Search and Replace
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.Search = function(oParaSearch)
{
	this.SearchMarks = [];
	
	for (var nPos = 0, nContentLen = this.Content.length; nPos < nContentLen; ++nPos)
	{
		this.Content[nPos].Search(oParaSearch);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.AddSearchResult = function(oSearchResult, isStart, oContentPos, nDepth)
{
	oSearchResult.RegisterClass(isStart, this);
	this.SearchMarks.push(new AscCommonWord.CParagraphSearchMark(oSearchResult, isStart, nDepth));
	this.Content[oContentPos.Get(nDepth)].AddSearchResult(oSearchResult, isStart, oContentPos, nDepth + 1);
};
CParagraphContentWithParagraphLikeContent.prototype.ClearSearchResults = function()
{
	this.SearchMarks = [];
};
CParagraphContentWithParagraphLikeContent.prototype.RemoveSearchResult = function(oSearchResult)
{
	for (var nIndex = 0, nMarksCount = this.SearchMarks.length; nIndex < nMarksCount; ++nIndex)
	{
		var oMark = this.SearchMarks[nIndex];
		if (oSearchResult === oMark.SearchResult)
		{
			this.SearchMarks.splice(nIndex, 1);
			nIndex--;
			nMarksCount--;
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetSearchElementId = function(bNext, bUseContentPos, ContentPos, Depth)
{
    // Определим позицию, начиная с которой мы будем искать ближайший найденный элемент
    var StartPos = 0;

    if ( true === bUseContentPos )
    {
        StartPos = ContentPos.Get( Depth );
    }
    else
    {
        if ( true === bNext )
        {
            StartPos = 0;
        }
        else
        {
            StartPos = this.Content.length - 1;
        }
    }

    // Производим поиск ближайшего элемента
    if ( true === bNext )
    {
        var ContentLen = this.Content.length;

        for ( var CurPos = StartPos; CurPos < ContentLen; CurPos++ )
        {
            var ElementId = this.Content[CurPos].GetSearchElementId( true, bUseContentPos && CurPos === StartPos ? true : false, ContentPos, Depth + 1 );
            if ( null !== ElementId )
                return ElementId;
        }
    }
    else
    {
        var ContentLen = this.Content.length;

        for ( var CurPos = StartPos; CurPos >= 0; CurPos-- )
        {
            var ElementId = this.Content[CurPos].GetSearchElementId( false, bUseContentPos && CurPos === StartPos ? true : false, ContentPos, Depth + 1 );
            if ( null !== ElementId )
                return ElementId;
        }
    }

    return null;
};
//----------------------------------------------------------------------------------------------------------------------
// Разное
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.SetReviewType = function(ReviewType, RemovePrChange)
{
    for (var Index = 0, Count = this.Content.length; Index < Count; Index++)
    {
        var Element = this.Content[Index];
        if (para_Run === Element.Type)
        {
            Element.SetReviewType(ReviewType);

            if (true === RemovePrChange)
                Element.RemovePrChange();
        }
        else if (Element.SetReviewType)
            Element.SetReviewType(ReviewType);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.SetReviewTypeWithInfo = function(ReviewType, ReviewInfo)
{
    for (var Index = 0, Count = this.Content.length; Index < Count; Index++)
    {
        var Element = this.Content[Index];
        if (Element && Element.SetReviewTypeWithInfo)
            Element.SetReviewTypeWithInfo(ReviewType, ReviewInfo);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.CheckRevisionsChanges = function(Checker, ContentPos, Depth)
{
    for (var CurPos = 0, Count = this.Content.length; CurPos < Count; CurPos++)
    {
        ContentPos.Update(CurPos, Depth);
        this.Content[CurPos].CheckRevisionsChanges(Checker, ContentPos, Depth + 1);
    }
};
CParagraphContentWithParagraphLikeContent.prototype.AcceptRevisionChanges = function(Type, bAll)
{
    if (true === this.Selection.Use || true === bAll)
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;
        if (StartPos > EndPos)
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        if (true === bAll)
        {
            StartPos = 0;
            EndPos   = this.Content.length - 1;
        }

        // Начинаем с конца, потому что при выполнении данной функции, количество элементов может изменяться
        if (this.Content[EndPos].AcceptRevisionChanges)
            this.Content[EndPos].AcceptRevisionChanges(Type, bAll);

        if (StartPos < EndPos)
        {
            for (var CurPos = EndPos - 1; CurPos > StartPos; CurPos--)
            {
                var Element = this.Content[CurPos];
                var ReviewType = Element.GetReviewType ? Element.GetReviewType() : reviewtype_Common;

                var isGoInside = false;
                if (reviewtype_Add === ReviewType)
                {
                    if (undefined === Type || c_oAscRevisionsChangeType.TextAdd === Type)
                        Element.SetReviewType(reviewtype_Common);

                    isGoInside = true;
                }
                else if (reviewtype_Remove === ReviewType)
                {
                    if (undefined === Type || c_oAscRevisionsChangeType.TextRem === Type)
                        this.Remove_FromContent(CurPos, 1, true);
                }
                else if (reviewtype_Common === ReviewType)
                {
                    isGoInside = true;
                }

                if (true === isGoInside && Element.AcceptRevisionChanges)
                    Element.AcceptRevisionChanges(Type, true);
            }

            if (this.Content[StartPos].AcceptRevisionChanges)
                this.Content[StartPos].AcceptRevisionChanges(Type, bAll);
        }

        this.Correct_Content();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.RejectRevisionChanges = function(Type, bAll)
{
    if (true === this.Selection.Use || true === bAll)
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;
        if (StartPos > EndPos)
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        if (true === bAll)
        {
            StartPos = 0;
            EndPos   = this.Content.length - 1;
        }

        // Начинаем с конца, потому что при выполнении данной функции, количество элементов может изменяться
        if (this.Content[EndPos].RejectRevisionChanges)
            this.Content[EndPos].RejectRevisionChanges(Type, bAll);

        if (StartPos < EndPos)
        {
            for (var CurPos = EndPos - 1; CurPos > StartPos; CurPos--)
            {
                var Element = this.Content[CurPos];
                var ReviewType = Element.GetReviewType ? Element.GetReviewType() : reviewtype_Common;

                var isGoInside = false;
                if (reviewtype_Remove === ReviewType)
                {
                    if (undefined === Type || c_oAscRevisionsChangeType.TextRem === Type)
                        Element.SetReviewType(reviewtype_Common);

                    isGoInside = true;
                }
                else if (reviewtype_Add === ReviewType)
                {
                    if (undefined === Type || c_oAscRevisionsChangeType.TextAdd === Type)
                        this.Remove_FromContent(CurPos, 1, true);
                }
                else if (reviewtype_Common === ReviewType)
                {
                    isGoInside = true;
                }

                if (true === isGoInside && Element.RejectRevisionChanges)
                    Element.RejectRevisionChanges(Type, true);
            }

            if (this.Content[StartPos].RejectRevisionChanges)
                this.Content[StartPos].RejectRevisionChanges(Type, bAll);
        }

        this.Correct_Content();
    }
};
CParagraphContentWithParagraphLikeContent.prototype.private_CheckUpdateBookmarks = function(Items)
{
	if (!Items)
		return;

	for (var nIndex = 0, nCount = Items.length; nIndex < nCount; ++nIndex)
	{
		var oItem = Items[nIndex];
		if (oItem && para_Bookmark === oItem.Type)
		{
			var oLogicDocument = this.Paragraph && this.Paragraph.LogicDocument ? this.Paragraph.LogicDocument : editor.WordControl.m_oLogicDocument;
			oLogicDocument.GetBookmarksManager().SetNeedUpdate(true);
			return;
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetFootnotesList = function(oEngine)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		if (this.Content[nIndex].GetFootnotesList)
			this.Content[nIndex].GetFootnotesList(oEngine);

		if (oEngine.IsRangeFull())
			return;
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GotoFootnoteRef = function(isNext, isCurrent, isStepOver, isStepFootnote, isStepEndnote)
{
	var nPos = 0;

	if (true === isCurrent)
	{
		if (true === this.Selection.Use)
			nPos = Math.min(this.Selection.StartPos, this.Selection.EndPos);
		else
			nPos = this.State.ContentPos;
	}
	else
	{
		if (true === isNext)
			nPos = 0;
		else
			nPos = this.Content.length - 1;
	}

	if (true === isNext)
	{
		for (var nIndex = nPos, nCount = this.Content.length - 1; nIndex < nCount; ++nIndex)
		{
			var nRes = this.Content[nIndex].GotoFootnoteRef ? this.Content[nIndex].GotoFootnoteRef(true, true === isCurrent && nPos === nIndex, isStepOver, isStepFootnote, isStepEndnote) : 0;

			if (nRes > 0)
				isStepOver = true;
			else  if (-1 === nRes)
				return true;
		}
	}
	else
	{
		for (var nIndex = nPos; nIndex >= 0; --nIndex)
		{
			var nRes = this.Content[nIndex].GotoFootnoteRef ? this.Content[nIndex].GotoFootnoteRef(true, true === isCurrent && nPos === nIndex, isStepOver, isStepFootnote, isStepEndnote) : 0;

			if (nRes > 0)
				isStepOver = true;
			else  if (-1 === nRes)
				return true;
		}
	}

	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.GetFootnoteRefsInRange = function(arrFootnotes, _CurLine, _CurRange)
{
	var CurLine = _CurLine - this.StartLine;
	var CurRange = (0 === CurLine ? _CurRange - this.StartRange : _CurRange);

	var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
	var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

	for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
	{
		if (this.Content[CurPos].GetFootnoteRefsInRange)
			this.Content[CurPos].GetFootnoteRefsInRange(arrFootnotes, _CurLine, _CurRange);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetAllContentControls = function(arrContentControls)
{
	if (!arrContentControls)
		arrContentControls = [];

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		if (this.Content[nIndex].GetAllContentControls)
			this.Content[nIndex].GetAllContentControls(arrContentControls);
	}

	return arrContentControls;
};
CParagraphContentWithParagraphLikeContent.prototype.GetSelectedContentControls = function(arrContentControls)
{
	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (StartPos > EndPos)
		{
			StartPos = this.Selection.EndPos;
			EndPos   = this.Selection.StartPos;
		}

		for (var Index = StartPos; Index <= EndPos; ++Index)
		{
			if (this.Content[Index].GetSelectedContentControls)
				this.Content[Index].GetSelectedContentControls(arrContentControls);
		}
	}
	else
	{
		if (this.Content[this.State.ContentPos].GetSelectedContentControls)
			this.Content[this.State.ContentPos].GetSelectedContentControls(arrContentControls);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.CreateRunWithText = function(sValue)
{
	var oRun = new ParaRun();
	oRun.AddText(sValue);
	oRun.Set_Pr(this.Get_FirstTextPr());
	return oRun;
};
CParagraphContentWithParagraphLikeContent.prototype.ReplaceAllWithText = function(sValue)
{
	var oRun = this.CreateRunWithText(sValue);
	oRun.Apply_TextPr(this.Get_TextPr(), undefined, true);
	this.Remove_FromContent(0, this.Content.length);
	this.Add_ToContent(0, oRun);
	this.MoveCursorToStartPos();
};
CParagraphContentWithParagraphLikeContent.prototype.FindNextFillingForm = function(isNext, isCurrent, isStart)
{
	var nCurPos = this.Selection.Use === true ? this.Selection.EndPos : this.State.ContentPos;

	var nStartPos = 0, nEndPos = 0;
	if (isCurrent)
	{
		if (isStart)
		{
			nStartPos = nCurPos;
			nEndPos   = isNext ? this.Content.length - 1 : 0;
		}
		else
		{
			nStartPos = isNext ? 0 : this.Content.length - 1;
			nEndPos   = nCurPos;
		}
	}
	else
	{
		if (isNext)
		{
			nStartPos = 0;
			nEndPos   = this.Content.length - 1;
		}
		else
		{
			nStartPos = this.Content.length - 1;
			nEndPos   = 0;
		}
	}

	if (isNext)
	{
		for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
		{
			if (this.Content[nIndex].FindNextFillingForm)
			{
				var oRes = this.Content[nIndex].FindNextFillingForm(true, isCurrent && nIndex === nCurPos, isStart);
				if (oRes)
					return oRes;
			}
		}
	}
	else
	{
		for (var nIndex = nStartPos; nIndex >= nEndPos; --nIndex)
		{
			if (this.Content[nIndex].FindNextFillingForm)
			{
				var oRes = this.Content[nIndex].FindNextFillingForm(false, isCurrent && nIndex === nCurPos, isStart);
				if (oRes)
					return oRes;
			}
		}
	}

	return null;
};
CParagraphContentWithParagraphLikeContent.prototype.IsEmpty = function(oPr)
{
	return this.Is_Empty(oPr);
};
CParagraphContentWithParagraphLikeContent.prototype.AddContentControl = function()
{
	if (true === this.IsSelectionUse())
	{
		if (this.Selection.StartPos === this.Selection.EndPos && para_Run !== this.Content[this.Selection.StartPos].Type)
		{
			if (this.Content[this.Selection.StartPos].AddContentControl)
				return this.Content[this.Selection.StartPos].AddContentControl();

			return null;
		}
		else
		{
			var nStartPos = this.Selection.StartPos;
			var nEndPos   = this.Selection.EndPos;
			if (nEndPos < nStartPos)
			{
				nStartPos = this.Selection.EndPos;
				nEndPos   = this.Selection.StartPos;
			}

			for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
			{
				if (para_Run !== this.Content[nIndex].Type)
				{
					// TODO: Вывести сообщение, что в данном месте нельзя добавить Plain text content control
					return null;
				}
			}

			var oContentControl = new CInlineLevelSdt();
			oContentControl.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
			oContentControl.SetDefaultTextPr(this.GetDirectTextPr());

			var oNewRun = this.Content[nEndPos].Split_Run(Math.max(this.Content[nEndPos].Selection.StartPos, this.Content[nEndPos].Selection.EndPos));
			this.Add_ToContent(nEndPos + 1, oNewRun);

			oNewRun = this.Content[nStartPos].Split_Run(Math.min(this.Content[nStartPos].Selection.StartPos, this.Content[nStartPos].Selection.EndPos));
			this.Add_ToContent(nStartPos + 1, oNewRun);

			oContentControl.ReplacePlaceHolderWithContent();
			for (var nIndex = nEndPos + 1; nIndex >= nStartPos + 1; --nIndex)
			{
				oContentControl.Add_ToContent(0, this.Content[nIndex]);
				this.Remove_FromContent(nIndex, 1);
			}
			if (oContentControl.IsEmpty())
				oContentControl.ReplaceContentWithPlaceHolder();

			this.Add_ToContent(nStartPos + 1, oContentControl);
			this.Selection.StartPos = nStartPos + 1;
			this.Selection.EndPos   = nStartPos + 1;
			oContentControl.SelectAll(1);

			return oContentControl;
		}
	}
	else
	{
		var oContentControl = new CInlineLevelSdt();
		oContentControl.SetDefaultTextPr(this.GetDirectTextPr());
		oContentControl.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
		oContentControl.ReplaceContentWithPlaceHolder(false);
		this.Add(oContentControl);
		return oContentControl;
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetElement = function(nPos)
{
	if (nPos < 0 || nPos >= this.Content.length)
		return null;

	return this.Content[nPos];
};
CParagraphContentWithParagraphLikeContent.prototype.GetElementsCount = function()
{
	return this.Content.length;
};
CParagraphContentWithParagraphLikeContent.prototype.PreDelete = function()
{
	if (this.Paragraph && this.Paragraph.isPreventedPreDelete())
		return;
	
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		if (this.Content[nIndex] && this.Content[nIndex].PreDelete)
			this.Content[nIndex].PreDelete(true);
	}

	this.RemoveSelection();
};
CParagraphContentWithParagraphLikeContent.prototype.GetCurrentPermRanges = function(permRanges, isCurrent)
{
	let endPos = isCurrent ? Math.min(this.State.ContentPos, this.Content.length - 1) : this.Content.length - 1;
	for (let pos = 0; pos <= endPos; ++pos)
	{
		this.Content[pos].GetCurrentPermRanges(permRanges, isCurrent && pos === endPos);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetCurrentComplexFields = function(arrComplexFields, isCurrent, isFieldPos)
{
	var nEndPos = isCurrent ? this.State.ContentPos : this.Content.length - 1;
	for (var nIndex = 0; nIndex <= nEndPos; ++nIndex)
	{
		if (this.Content[nIndex] && this.Content[nIndex].GetCurrentComplexFields)
			this.Content[nIndex].GetCurrentComplexFields(arrComplexFields, isCurrent && nIndex === nEndPos, isFieldPos);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetDirectTextPr = function()
{
	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;

		if (StartPos > EndPos)
		{
			StartPos = this.Selection.EndPos;
			EndPos   = this.Selection.StartPos;
		}

		while (true === this.Content[StartPos].IsSelectionEmpty() && StartPos < EndPos)
			StartPos++;

		return this.Content[StartPos].GetDirectTextPr();
	}
	else
	{
		return this.Content[this.State.ContentPos].GetDirectTextPr();
	}
};
CParagraphContentWithParagraphLikeContent.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	if (!arrFields)
		arrFields = [];

	var nStartPos = isUseSelection ?
		(this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos)
		: 0;

	var nEndPos = isUseSelection ?
		(this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos)
		: this.Content.length - 1;

	for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
	{
		this.Content[nIndex].GetAllFields(isUseSelection, arrFields);
	}

	return arrFields;
};
CParagraphContentWithParagraphLikeContent.prototype.CanAddComment = function()
{
	if (!this.Selection.Use)
		return true;

	var nStartPos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
	var nEndPos   = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;

	for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		if (this.Content[nPos].CanAddComment && !this.Content[nPos].CanAddComment())
			return false;
	}

	return true;
};
CParagraphContentWithParagraphLikeContent.prototype.RemoveElement = function(element)
{
	for (let i = 0, count = this.Content.length; i < count; ++i)
	{
		let item = this.Content[i];
		if (item === element)
		{
			this.RemoveFromContent(i, 1);
			return true;
		}
		else if (item.RemoveElement(element))
		{
			return true;
		}
	}
	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.CheckRunContent = function(fCheck, oStartPos, oEndPos, nDepth, oCurrentPos, isForward)
{
	if (undefined === isForward)
		isForward = true;
	
	let nStartPos = oStartPos && oStartPos.GetDepth() >= nDepth ? oStartPos.Get(nDepth) : 0;
	let nEndPos   = oEndPos && oEndPos.GetDepth() >= nDepth ? oEndPos.Get(nDepth) : this.Content.length - 1;
	
	if (isForward)
	{
		for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
		{
			let _s = oStartPos && nPos === nStartPos ? oStartPos : null;
			let _e = oEndPos && nPos === nEndPos ? oEndPos : null;
			
			if (oCurrentPos)
				oCurrentPos.Update(nPos, nDepth);
			
			if (this.Content[nPos].CheckRunContent(fCheck, _s, _e, nDepth + 1, oCurrentPos, isForward))
				return true;
		}
	}
	else
	{
		for (var nPos = nEndPos; nPos >= nStartPos; --nPos)
		{
			let _s = oStartPos && nPos === nStartPos ? oStartPos : null;
			let _e = oEndPos && nPos === nEndPos ? oEndPos : null;
			
			if (oCurrentPos)
				oCurrentPos.Update(nPos, nDepth);
			
			if (this.Content[nPos].CheckRunContent(fCheck, _s, _e, nDepth + 1, oCurrentPos, isForward))
				return true;
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.ProcessComplexFields = function(oComplexFields)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		this.Content[nPos].ProcessComplexFields(oComplexFields);
	}
};
CParagraphContentWithParagraphLikeContent.prototype.CorrectContentPos = function()
{
	if (this.IsSelectionUse())
		return;

	var nCount = this.Content.length;
	var nCurPos = Math.min(Math.max(0, this.State.ContentPos), nCount - 1);

	// Ищем элемент, в котором может стоять курсор
	while (nCurPos > 0 && !this.Content[nCurPos].IsCursorPlaceable())
	{
		nCurPos--;
		this.Content[nCurPos].MoveCursorToEndPos();
	}

	while (nCurPos < nCount && !this.Content[nCurPos].IsCursorPlaceable())
	{
		nCurPos++;
		this.Content[nCurPos].MoveCursorToStartPos(false);
	}

	// Если курсор находится в начале или конце гиперссылки, тогда выводим его из гиперссылки
	while (nCurPos > 0 && para_Run !== this.Content[nCurPos].Type && para_Math !== this.Content[nCurPos].Type && para_Field !== this.Content[nCurPos].Type && para_InlineLevelSdt !== this.Content[nCurPos].Type && true === this.Content[nCurPos].Cursor_Is_Start())
	{
		if (!this.Content[nCurPos - 1].IsCursorPlaceable())
			break;

		nCurPos--;
		this.Content[nCurPos].MoveCursorToEndPos();
	}

	while (nCurPos < nCount && para_Run !== this.Content[nCurPos].Type && para_Math !== this.Content[nCurPos].Type && para_Field !== this.Content[nCurPos].Type && para_InlineLevelSdt !== this.Content[nCurPos].Type && true === this.Content[nCurPos].Cursor_Is_End())
	{
		if (!this.Content[nCurPos + 1].IsCursorPlaceable())
			break;

		nCurPos++;
		this.Content[nCurPos].MoveCursorToStartPos(false);
	}

	this.State.ContentPos = nCurPos;

	this.Content[this.State.ContentPos].CorrectContentPos();
};
CParagraphContentWithParagraphLikeContent.prototype.GetFirstRun = function()
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		var oRun = this.Content[nIndex].GetFirstRun();
		if (oRun)
			return oRun;
	}

	return null;
};
CParagraphContentWithParagraphLikeContent.prototype.GetFirstRunNonEmpty = function()
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		var oRun = this.Content[nIndex].GetFirstRun();
		if (oRun&& oRun.GetElementsCount() > 0)
			return oRun;
	}

	return null;
};
CParagraphContentWithParagraphLikeContent.prototype.MakeSingleRunElement = function(isClearRun)
{
	if (this.Content.length !== 1 || para_Run !== this.Content[0].Type)
	{
		var oRun = new ParaRun(this.GetParagraph(), false);

		if (true !== isClearRun)
		{
			// У нас при открытии ран делится маскимально по 255 элементов внутри каждого рана, поэтому
			// в формах, где должен быть только 1 ран после открытия их может быть несколько. Объединяем здесь
			// все раны в один общий ран, чтобы исправить данную ситуцию

			var oParagraph = this.GetParagraph();
			var oCurrentRun = null;
			if (oParagraph)
			{
				var oCurPos = oParagraph.Get_ParaContentPos(false, false, false);
				oCurPos.DecreaseDepth(1);
				oCurrentRun = oParagraph.GetClassByPos(oCurPos);
				if (!oCurrentRun || !(oCurrentRun instanceof ParaRun))
					oCurrentRun = null;
			}

			var nNewCurPos = 0;
			var isFirst    = true;
			this.CheckRunContent(function(_oRun)
			{
				if (_oRun === oCurrentRun)
					nNewCurPos = _oRun.State.ContentPos + oRun.Content.length;

				var arrContentToInsert = [];
				for (var nPos = 0, nCount = _oRun.Content.length; nPos < nCount; ++nPos)
				{
					arrContentToInsert.push(_oRun.Content[nPos].Copy());
				}

				oRun.ConcatToContent(arrContentToInsert);

				if (isFirst && arrContentToInsert.length > 0)
				{
					oRun.SetPr(_oRun.GetDirectTextPr().Copy());
					isFirst = false;
				}
			});

			oRun.State.ContentPos = nNewCurPos;
		}

		if (this.Content.length > 0)
			this.RemoveFromContent(0, this.Content.length, true);

		this.AddToContent(0, oRun, true);
	}

	var oRun = this.Content[0];

	if (false !== isClearRun)
		oRun.ClearContent();

	return oRun;
};
CParagraphContentWithParagraphLikeContent.prototype.ClearContent = function()
{
	if (this.Content.length <= 0)
		return;

	this.RemoveFromContent(0, this.Content.length, true);
};
CParagraphContentWithParagraphLikeContent.prototype.GetFirstRunElementPos = function(nType, oStartPos, oEndPos, nDepth)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		oStartPos.Update(nPos, nDepth);
		oEndPos.Update(nPos, nDepth);

		if (this.Content[nPos].GetFirstRunElementPos(nType, oStartPos, oEndPos, nDepth + 1))
			return true;
	}

	return false;
};
CParagraphContentWithParagraphLikeContent.prototype.SetIsRecalculated = function(isRecalculated)
{
	if (!isRecalculated && this.GetParagraph())
		this.GetParagraph().SetIsRecalculated(false);
};
CParagraphContentWithParagraphLikeContent.prototype.CalculateTextToTable = function(oEngine)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		if (this.Content[nIndex].IsSolid())
			continue;

		this.Content[nIndex].CalculateTextToTable(oEngine);
	}
};

//----------------------------------------------------------------------------------------------------------------------
// Функции, которые должны быть реализованы в классах наследниках
//----------------------------------------------------------------------------------------------------------------------
CParagraphContentWithParagraphLikeContent.prototype.Add = function(Item)
{
	if (undefined !== Item.Parent)
		Item.Parent = this;

	switch (Item.Type)
	{
		case para_Run:
		case para_Hyperlink:
		case para_InlineLevelSdt:
		case para_Field:
		{
			var TextPr = this.Get_FirstTextPr();
			Item.SelectAll();
			Item.Apply_TextPr(TextPr);
			Item.RemoveSelection();

			var CurPos = this.State.ContentPos;
			var CurItem = this.Content[CurPos];
			if (para_Run === CurItem.Type)
			{
				var NewRun = CurItem.Split2(CurItem.State.ContentPos);
				this.Add_ToContent(CurPos + 1, Item);
				this.Add_ToContent(CurPos + 2, NewRun);

				this.State.ContentPos = CurPos + 2;
				this.Content[this.State.ContentPos].MoveCursorToStartPos();
			}
			else
			{
				CurItem.Add(Item);
			}

			break;
		}
		case para_Math :
		{
			var ContentPos = new AscWord.CParagraphContentPos();
			this.Get_ParaContentPos(false, false, ContentPos);
			var CurPos = ContentPos.Get(0);

			// Ран формула делит на части, а в остальные элементы добавляется целиком
			if (para_Run === this.Content[CurPos].Type)
			{
				// Разделяем текущий элемент (возвращается правая часть)
				var NewElement = this.Content[CurPos].Split(ContentPos, 1);

				if (null !== NewElement)
					this.Add_ToContent(CurPos + 1, NewElement, true);
				
				let paraMath = null;
				if (Item instanceof ParaMath)
				{
					paraMath = Item;
				}
				else if (Item instanceof AscCommonWord.MathMenu)
				{
					let textPr = Item.GetTextPr();
					paraMath = new ParaMath();
					paraMath.Root.Load_FromMenu(Item.Menu, this.GetParagraph(), textPr.Copy(), Item.GetText());
					paraMath.Root.Correct_Content(true);
					paraMath.ApplyTextPr(textPr.Copy(), undefined, true);
				}
				
				if (paraMath)
				{
					this.AddToContent(CurPos + 1, paraMath, true);
					this.State.ContentPos = CurPos + 1;
					this.Content[this.State.ContentPos].MoveCursorToEndPos(false);
				}
			}
			else
			{
				this.Content[CurPos].Add(Item);
			}

			break;
		}
		default:
		{
			this.Content[this.State.ContentPos].Add(Item);
			break;
		}
	}
};
CParagraphContentWithParagraphLikeContent.prototype.Undo = function(Data){};
CParagraphContentWithParagraphLikeContent.prototype.Redo = function(Data){};
CParagraphContentWithParagraphLikeContent.prototype.Save_Changes = function(Data, Writer){};
CParagraphContentWithParagraphLikeContent.prototype.Load_Changes = function(Reader){};
CParagraphContentWithParagraphLikeContent.prototype.Write_ToBinary2 = function(Writer){};
CParagraphContentWithParagraphLikeContent.prototype.Read_FromBinary2 = function(Reader){};

// TODO: Сделать и перенести в коммоны для изменений
/**
 * Универсальный метод для проверки лока для простых изменений внутри параграфа
 */
function private_ParagraphContentChangesCheckLock(lockData)
{
	let obj = this.Class;
	if (!this.IsContentChange() && lockData && lockData.isFillingForm())
		return lockData.lock();
	
	let isForm = false;
	let isCC   = false;
	while (obj)
	{
		if (obj.Lock)
			obj.Lock.Check(obj.Get_Id());
		
		isForm = isForm || (obj instanceof AscWord.CInlineLevelSdt && obj.IsForm());
		isCC   = isCC || obj instanceof AscWord.CInlineLevelSdt;
		
		if (!(obj instanceof AscWord.Paragraph) && obj.GetParent)
			obj = obj.GetParent()
		else
			obj = null;
	}
	
	if (this.IsContentChange())
	{
		if (isForm && lockData && !lockData.isSkipFormCheck())
			lockData.lock();

		if (!isCC && lockData && lockData.isFillingForm())
			lockData.lock();
	}
}
//--------------------------------------------------------export----------------------------------------------------
AscWord.ParagraphContentBase                     = CParagraphContentBase;
AscWord.ParagraphContentWithParagraphLikeContent = CParagraphContentWithParagraphLikeContent;
