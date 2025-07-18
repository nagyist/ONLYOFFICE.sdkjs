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

// Import
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var c_oAscSectionBreakType    = Asc.c_oAscSectionBreakType;

// TODO: В колонтитулах быстрые пересчеты отключены. Надо реализовать.

/**
 * Здесь мы пытаемся быстро пересчитать текущий параграф. Если быстрый пересчет срабатывает, тогда возвращаются
 * страницы, которые нужно перерисовать, в противном случае возвращается пустой массив.
 * @returns {*}
 */
Paragraph.prototype.Recalculate_FastWholeParagraph = function()
{
    if (this.Pages.length <= 0 || undefined === this.Parent)
        return [];

    if (true === this.Parent.IsHdrFtr(false))
        return [];

    //Не запускаемм быстрый пересчет, когда параграф находится в автофигуре с выставленным флагом подбора размера по размеру контента,
    // т. к. при расчете контента потребуется пересчет в новых размерах.
    if(true === this.Parent.Check_AutoFit())
    {
        return [];
    }

    // TODO: Отключаем это ускорение в таблицах, т.к. в таблицах и так есть свое ускорение. Но можно и это ускорение
    // подключить, для этого надо проверять изменились ли MinMax ширины и набираем ли мы в строке заголовков.
	var oCell = this.IsTableCellContent(true);
    if (oCell && oCell.GetTable())
	{
		if (tbllayout_AutoFit === oCell.GetTable().Get_CompiledPr(false).TablePr.TableLayout || oCell.IsInHeader(true))
			return [];
	}

    // Если изменения происходят в специальном пустом параграфе-конце секции, тогда запускаем обычный пересчет
    if (this.bFromDocument && this.LogicDocument && (!this.LogicDocument.Pages[this.Get_StartPage_Absolute()] || true === this.LogicDocument.Pages[this.Get_StartPage_Absolute()].Check_EndSectionPara(this)))
        return [];

    // Если параграф - рамка с автошириной, надо пересчитывать по обычному
    if (1 === this.Lines.length && true !== this.Is_Inline())
        return [];

	this.SetIsRecalculated(true);

    // Здесь мы отдельно обрабатываем случаи быстрого пересчета параграфов, которые были разбиты на 1-2
    // страницы. Если параграф был разбит более чем на 2 страницы, то такое ускорение уже не имеет смысла.
	// В обеих ситуациях нужно проверить, что EndInfo остался прежним, иначе требуется пересчет
    if (1 === this.Pages.length)
    {
        // Если параграф был разбит на 1 страницу изначально, тогда мы проверяем, чтобы он после пересчета
        // был также разбит на 1 страницу, кроме этого проверяем изменились ли границы параграфа, и проверяем
        // последнюю строку на pageBreak/columnBreak, а во время пересчета смотрим изменяeтся ли положение
        // flow-объектов, привязанных к данному параграфу, кроме того, если по какой-то причине пересчет возвращает
        // не recalcresult_NextElement, тогда тоже отменяем быстрый пересчет
		
		

		var oEndInfo             = this.GetEndInfo().Copy();
        var OldBounds            = this.Pages[0].Bounds;
        var isPageBreakLastLine1 = this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakPage;
        var isPageBreakLastLine2 = this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakRealPage;
        var FastRecalcResult     = this.Recalculate_Page(0, true, true);

		if (!this.GetEndInfo().IsEqual(oEndInfo))
			return [];

        if (FastRecalcResult & recalcresult_NextElement
            && 1 === this.Pages.length
            && true === this.Pages[0].Bounds.Compare(OldBounds)
            && this.Lines.length > 0
            && isPageBreakLastLine1 === (this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakPage)
            && isPageBreakLastLine2 === (this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakRealPage))
        {
            //console.log("Recalc Fast WholeParagraph 1 page");
            return [this.Get_AbsolutePage(0)];
        }
    }
    else if (2 === this.Pages.length)
    {
		var oEndInfo = this.GetEndInfo().Copy();

        // Если параграф был разбит на 2 страницы изначально, тогда мы проверяем, чтобы он после пересчета
        // был также разбит на 2 страницы, кроме этого проверяем изменились ли границы параграфа на каждой странице,
        // а во время пересчета смотрим изменяeтся ли положение flow-объектов, привязанных к данному параграфу.
        // Кроме того, если по какой-то причине пересчет возвращает не recalcresult_NextPage на первой странице, или не
        // recalcresult_NextElement, тогда тоже отменяем быстрый пересчет.
        var OldBounds_0 = this.Pages[0].Bounds;
        var OldBounds_1 = this.Pages[1].Bounds;

        var isPageBreakLastLine1 = this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakPage;
        var isPageBreakLastLine2 = this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakRealPage;

        // Чтобы защититься от неправильной работы, связанной с переносом параграфа на новую страницу,
        // будем следить за тем, начинался ли изначально параграф с новой страницы, и начинается ли он с
        // новой страницы сейчас.
        var OldStartFromNewPage = this.Pages[0].StartLine < 0 ? true : false;

        // Чтобы защититься от неправильной работой с висячими строками, будем следить за количеством строк
        // если оно меньше либо равно 2 на какой-либо странице до/после пересчета.
        var OldLinesCount_0 = this.Pages[0].EndLine - this.Pages[0].StartLine + 1;
        var OldLinesCount_1 = this.Pages[1].EndLine - this.Pages[1].StartLine + 1;

		var FastRecalcResult = this.Recalculate_Page(0, true, true);

		if (!(FastRecalcResult & recalcresult_NextPage))
			return [];

        FastRecalcResult = this.Recalculate_Page(1, true, true);
        if (!(FastRecalcResult & recalcresult_NextElement))
            return [];

        // Сравниваем количество страниц (хотя оно должно быть 2 к данному моменту) и границы каждой страницы
        if (2 !== this.Pages.length || true !== this.Pages[0].Bounds.Compare(OldBounds_0) || true !== this.Pages[1].Bounds.Compare(OldBounds_1))
            return [];

        // Сравниваем наличие pageBreak/columnBreak в последней строке
        if (this.Lines.length <= 0
            || isPageBreakLastLine1 !== (this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakPage)
            || isPageBreakLastLine2 !== (this.Lines[this.Lines.length - 1].Info & paralineinfo_BreakRealPage))
            return [];

        // Проверяем пустую первую страницу
        var StartFromNewPage = this.Pages[0].StartLine < 0 ? true : false;
        if (StartFromNewPage !== OldStartFromNewPage)
            return [];

        // Если параграф начался с новой страницы, тогда у него не надо проверять висячие строки
        if (true !== StartFromNewPage)
        {
            var LinesCount_0 = this.Pages[0].EndLine - this.Pages[0].StartLine + 1;
            var LinesCount_1 = this.Pages[1].EndLine - this.Pages[1].StartLine + 1;

            if ((OldLinesCount_0 <= 2 || LinesCount_0 <= 2) && OldLinesCount_0 !== LinesCount_0)
                return [];

            if ((OldLinesCount_1 <= 2 || LinesCount_1 <= 2) && OldLinesCount_1 !== LinesCount_1)
                return [];
        }

		if (!this.GetEndInfo().IsEqual(oEndInfo))
			return [];

        //console.log("Recalc Fast WholeParagraph 2 pages");

        // Если параграф начинается с новой страницы, тогда не надо перерисовывать первую страницу, т.к. она
        // изначально была пустая, и сейчас пустая.

        if (true === StartFromNewPage)
        {
            return [this.Get_AbsolutePage(1)];
        }
        else
        {
            var PageAbs0 = this.Get_AbsolutePage(0);
            var PageAbs1 = this.Get_AbsolutePage(1);
            if (PageAbs0 !== PageAbs1)
                return [PageAbs0, PageAbs1];
            else
                return [PageAbs0];
        }
    }

    return [];
};
/**
 * Ивент, если удалось быстро пересчитать параграф
 */
Paragraph.prototype.OnFastRecalculate = function()
{
	let topDocument = this.GetTopDocumentContent();
	if (topDocument && (topDocument instanceof AscWord.FootEndnote))
		topDocument.OnFastRecalculate();
};
/**
 * Пытаемся быстро рассчитать отрезок, в котором произошли изменения, и если ничего не съехало, тогда
 * перерисовываем страницу, в противном случаем запускаем обычный пересчет.
 * @param {CParaPos} oParaPos
 * @returns {*} -1 если быстрый пересчет не получился, либо номер страницы, которую нужно перерисовать
 */
Paragraph.prototype.RecalculateFastRunRange = function(oParaPos)
{
	if (this.Pages.length <= 0)
        return -1;

    if (true === this.Parent.IsHdrFtr(false))
        return -1;

    if (!oParaPos)
        return -1;

    var Line  = oParaPos.Line;
    var Range = oParaPos.Range;

    // Такое возможно, если у нас шел долгий пересчет (например, из-за изменений второго пользователя) и в это же время
	// запустился быстрый (ввод символа). Долгий пересчет успел сбросить рассчет данного параграфа, но не пересчитал параграф
	// до конца, а в это время у данного параграфа запросился быстрый пересчет.
    if (this.Lines.length <= oParaPos.Line)
    	return -1;

    // TODO: Отключаем это ускорение в таблицах, т.к. в таблицах и так есть свое ускорение. Но можно и это ускорение
    // подключить, для этого надо проверять изменились ли MinMax ширины и набираем ли мы в строке заголовков.
    if ( undefined === this.Parent || true === this.IsTableCellContent() )
        return -1;

    //Не запускаемм быстрый пересчет, когда параграф находится в автофигуре с выставленным флагом подбора размера по размеру контента,
    // т. к. при расчете контента потребуется пересчет в новых размерах.
    if(true === this.Parent.Check_AutoFit())
    {
        return -1;
    }

    // Если мы находимся в строке, которая была полностью перенесена из-за обтекания,  и мы добавляем пробел, или
    // удаляем символ, тогда нам запускать обычный пересчет, т.к. первое слово может начать убираться в промежутках
    // обтекания, которых у нас нет в отрезках строки
    if ( true === this.Lines[Line].RangeY )
    {
        // TODO: Сделать проверку на добавление пробела и удаление
        return -1;
    }

    // Если у нас есть PageBreak в строке, запускаем обычный пересчет, либо если это пустой параграф.
    if (this.Lines[Line].Info & paralineinfo_BreakPage || (this.Lines[Line].Info & paralineinfo_Empty &&  this.Lines[Line].Info & paralineinfo_End))
        return  -1;

    if ( 0 === Line && 0 === Range && undefined !== this.Get_SectionPr() )
    {
        return -1;
    }

    // Если наш параграф является рамкой с авто шириной, тогда пересчитываем по обычному
    // TODO: Улучишить данную проверку
    if ( 1 === this.Lines.length && true !== this.Is_Inline() )
        return -1;


	// Мы должны пересчитать как минимум 3 отрезка: текущий, предыдущий и следующий, потому что при удалении элемента
    // или добавлении пробела первое слово в данном отрезке может убраться в предыдущем отрезке, и кроме того при
    // удалении возможен вариант, когда мы неправильно определили отрезок (т.е. более ранний взяли). Но возможен
    // вариант, при котором предыдущий или/и следующий отрезки - пустые, т.е. там нет ни одного текстового элемента
    // тогда мы начинаем проверять с отрезка, в котором есть хоть что-то.

	var PrevLine  = Line;
	var PrevRange = Range;

	while (PrevLine >= 0)
	{
		PrevRange--;

		if (PrevRange < 0)
		{
			PrevLine--;

			if (PrevLine < 0)
				break;

			PrevRange = this.Lines[PrevLine].Ranges.length - 1;
		}

		if (!this.IsEmptyRange(PrevLine, PrevRange))
			break;
	}

	if (PrevLine < 0)
	{
		PrevLine  = Line;
		PrevRange = Range;
	}

	var NextLine  = Line;
	var NextRange = Range;

	var LinesCount = this.Lines.length;

	while (NextLine <= LinesCount - 1)
	{
		NextRange++;

		if (NextRange > this.Lines[NextLine].Ranges.length - 1)
		{
			NextLine++;

			if (NextLine > LinesCount - 1)
				break;

			NextRange = 0;
		}

		if (!this.IsEmptyRange(NextLine, NextRange))
			break;
	}

	if (NextLine > LinesCount - 1)
	{
		NextLine  = Line;
		NextRange = Range;
	}

	// Важно, что здесь мы формируем текст без учета переносов (т.е. не Temporary), но при этом сам флаг
	// Temporary c элементов текста не снимаем, т.к. он рассчитывается уже в процессе Recalculate_Range,
	// а для всех строк, не учавствующих в быстром пересчете, мы должны все сохранить как есть
	this.ShapeText();
	this.ShapeTextInRange(this.Get_StartRangePos2(Line, Range), this.Get_EndRangePos2(Line, Range));
	
	this.HyphenateText();

	// Если у нас отрезок, в котором произошли изменения является отрезком с нумерацией, тогда надо запустить
	// обычный пересчет.
	if (null !== this.Numbering.Item && (PrevLine < this.Numbering.Line || (PrevLine === this.Numbering.Line && PrevRange <= this.Numbering.Range)))
	{
		// TODO: Сделать проверку на само изменение, переместилась ли нумерация
		var CompiledParaPr = this.Get_CompiledPr2(false).ParaPr;
		if (this.Numbering.Type === para_Numbering)
		{
			var NumPr = CompiledParaPr.NumPr;
			if (undefined !== NumPr && undefined !== NumPr.NumId && 0 !== NumPr.NumId && "0" !== NumPr.NumId)
			{
				return -1;
			}
		}
		else
		{
			var Bullet = this.Numbering.Bullet;
			if (Bullet && null !== Bullet.m_oTextPr && null !== Bullet.m_nNum && null != Bullet.m_sString && Bullet.m_sString.length !== 0)
			{
				return -1;
			}
		}
	}

	// Если мы дошли до данного места, значит быстрый пересчет отрезка разрешен
    var CurLine  = PrevLine;
    var CurRange = PrevRange;

    // TODO: Для включения данной проверки нужно пробегаться по строке и пересчитывать Line.Info
    // var arrLinesMetrics = [];
    // for (var nLineIndex = 0; nLineIndex <= NextLine - CurLine; ++nLineIndex)
	// {
	// 	arrLinesMetrics.push(this.Lines[CurLine + nLineIndex].Metrics.Copy());
	// }

	var Result;
    while ( ( CurLine < NextLine ) || ( CurLine === NextLine && CurRange <= NextRange ) )
    {
        var TempResult = this.recalculateRangeFast(CurRange, CurLine);
        if ( -1 === TempResult )
			return -1;

        if ( CurLine === Line && CurRange === Range )
            Result = TempResult;

        CurRange++;

        if ( CurRange > this.Lines[CurLine].Ranges.length - 1 )
        {
            CurLine++;
            CurRange = 0;
        }
    }

    // var oParaPr = this.Get_CompiledPr2(false).ParaPr;
	// var oPRS    = this.m_oPRSW;
	// for (var nLineIndex = PrevLine; nLineIndex <= NextLine; ++nLineIndex)
	// {
	// 	oPRS.Reset_Line();
	// 	this.Lines[nLineIndex].Metrics.Reset();
	// 	this.private_RecalculateLineMetrics(nLineIndex, oParaPos.Page, oPRS, oParaPr);
	//
	// 	if (!this.Lines[nLineIndex].Metrics.IsEqual(arrLinesMetrics[nLineIndex - PrevLine]))
	// 		return null;
	// }

    // Во время пересчета сбрасываем привязку курсора к строке.
    this.CurPos.Line  = -1;
    this.CurPos.Range = -1;

    this.RequestSpellCheck();
	this.SetIsRecalculated(true);

	//console.log("Recalc Fast Range");

    return this.Get_AbsolutePage(Result);
};

/**
 * Функция для пересчета страницы параграфа.
 * @param {number} CurPage - Номер страницы, которую нужно пересчитать (относительный номер страницы для параграфа).
 * Предыдущая страница ДОЛЖНА быть пересчитана, если задано не нулевое значение.
 * @param {boolean} isStart - Устаревший параметр, добавлен для совместимости
 * @param {boolean} isFast - быстрый ли пересчет
 * @returns {*} Возвращается результат пересчета
 */
Paragraph.prototype.Recalculate_Page = function(CurPage, isStart, isFast)
{
	if (0 === CurPage)
	{
		this.CalculatedFrame = null;
		this.ShapeText();
		this.HyphenateText();
	}

    this.Clear_NearestPosArray();

    // Во время пересчета сбрасываем привязку курсора к строке.
    this.CurPos.Line  = -1;
    this.CurPos.Range = -1;
	this.SetIsRecalculated(true);

    this.FontMap.NeedRecalc = true;

    this.RequestSpellCheck();
    this.RecalculateEndInfo(isFast, true);

	var RecalcResult = this.private_RecalculatePage( CurPage, isFast );

    this.private_CheckColumnBreak(CurPage);

    this.Parent.RecalcInfo.Reset_WidowControl();

    if (RecalcResult & recalcresult_NextElement && window['AscCommon'].g_specialPasteHelper && window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph === this.GetId())
		window['AscCommon'].g_specialPasteHelper.SpecialPasteButtonById_Show();

    if (RecalcResult & recalcresult_NextElement)
		this.UpdateLineNumbersInfo();

    return RecalcResult;
};

/**
 * Функция для пересчета страницы параграфа, так чтобы на данной странице ничего не было. Применяется, когда из-за
 * пересчета плавающей автофигуры нужно сразу перейти на следующую страницу, пропустив несколько колонок.
 * @param {number} PageIndex - Номер страницы, пересчет которой мы пропускаем. (предыдущая страница ДОЛЖНА быть
 * пересчитана, если это не нулевое значение)
 */
Paragraph.prototype.Recalculate_SkipPage = function(PageIndex)
{
    if (0 === PageIndex)
    {
        this.StartFromNewPage();
    }
    else
    {
        var PrevPage = this.Pages[PageIndex - 1];

        var EndLine       = Math.max(PrevPage.StartLine, PrevPage.EndLine); // На случай, если предыдущая страница тоже пустая
        var NewPage       = new CParaPage(PrevPage.X, PrevPage.Y, PrevPage.XLimit, PrevPage.YLimit, EndLine);
        NewPage.StartLine = EndLine;
        NewPage.EndLine   = EndLine - 1;
        NewPage.TextPr    = PrevPage.TextPr;

        this.Pages[PageIndex] = NewPage;
    }
};

/**
 * Функция для сохранения объекта пересчета.
 * @returns {*} Возвращается объект (CParagraphRecalculateObject) с информацией о текущем пересчете параграфа
 */
Paragraph.prototype.SaveRecalculateObject = function()
{
	var RecalcObj = new CParagraphRecalculateObject();
	RecalcObj.Save(this);
	return RecalcObj;
};

/**
 * Загрузка сохраненного раннее пересчета.
 * @param RecalcObj (CParagraphRecalculateObject)
 */
Paragraph.prototype.LoadRecalculateObject = function(RecalcObj)
{
    RecalcObj.Load(this);
};

/**
 * Очистка рассчетных классов параграфа.
 */
Paragraph.prototype.PrepareRecalculateObject = function()
{
	this.Pages = [];
	this.Lines = [];

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].PrepareRecalculateObject();
	}
};

/**
 * Пересчитываем первую страницу параграфа так, чтобы он начинался с новой страницы.
 */
Paragraph.prototype.StartFromNewPage = function()
{
    this.Pages.length = 1;
    this.Pages[0] = new CParaPage(this.X, this.Y, this.XLimit, this.YLimit, 0);

    // Добавляем разрыв страницы
    this.Pages[0].Set_EndLine(-1);
    this.Lines[-1] = new CParaLine();
};
/**
 * Быстрый пересчет заданного отрезка
 * @param {number} iRange - номер отрезка
 * @param {number} iLine - номер строки
 * @returns {number} -1 - если требуется полный пересчет, либо номер текущей страницы
 */
Paragraph.prototype.recalculateRangeFast = function(iRange, iLine)
{
	let wrapState = AscWord.ParagraphStatePool.getWrapState();
	wrapState.SetFast(true);
	let result = this.private_RecalculateFastRange(wrapState, iRange, iLine);
	AscWord.ParagraphStatePool.release(wrapState);
	return result;
};
Paragraph.prototype.private_RecalculateFastRange       = function(PRS, CurRange, CurLine)
{
    var XStart, YStart, XLimit, YLimit;

    // Определим номер страницы
    var CurPage  = 0;
    var PagesLen = this.Pages.length;
    for ( var TempPage = 0; TempPage < PagesLen; TempPage++ )
    {
        var __Page = this.Pages[TempPage];
        if ( CurLine <= __Page.EndLine && CurLine >= __Page.FirstLine )
        {
            CurPage = TempPage;
            break;
        }
    }

    if ( -1 === CurPage )
        return -1;

    var ParaPr = this.Get_CompiledPr2(false).ParaPr;

    if ( 0 === CurPage )//|| ( undefined != this.Get_FramePr() && this.Parent instanceof CDocument ) )
    {
        XStart = this.X;
        YStart = this.Y;
        XLimit = this.XLimit;
        YLimit = this.YLimit;
    }
    else
    {
        var PageStart = this.Parent.Get_PageContentStartPos2(this.PageNum, this.ColumnNum, CurPage, this.Index);

        XStart = PageStart.X;
        YStart = PageStart.Y;
        XLimit = PageStart.XLimit;
        YLimit = PageStart.YLimit;
    }

    PRS.XStart = XStart;
    PRS.YStart = YStart;
    PRS.XLimit = XLimit;// - ParaPr.Ind.Right;
    PRS.YLimit = YLimit;

    // Обнуляем параметры PRS для строки и отрезка
    PRS.Reset_Line();

    PRS.Page  = CurPage;
    PRS.Line  = CurLine;
    PRS.Range = CurRange;

    PRS.Ranges      = this.Lines[CurLine].Ranges;
    PRS.RangesCount = this.Lines[CurLine].Ranges.length - 1;

    PRS.Paragraph = this;

    var Line  = this.Lines[CurLine];
    var Range = Line.Ranges[CurRange];

    let nStartPos = Range.StartPos;
    let nEndPos   = Range.EndPos;

    // Обновляем состояние пересчета
	PRS.resetRange(Range);

	let arrSavedLines = [];
	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		arrSavedLines.push(this.Content[nPos].SaveRecalculateObject(true));
	}

	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		let oItem = this.Content[nPos];
		if (para_Math === oItem.Type)
		{
			// TODO: Надо бы перенести эту проверку на изменение контента параграфа
			oItem.Set_Inline(true !== this.CheckMathPara(nPos));
			PRS.bFastRecalculate = true; // чтобы не обновить случайно StartLine (Recalculate_Reset)
		}

		PRS.Update_CurPos(nPos, 0);

		oItem.Recalculate_Range(PRS, ParaPr, 1);

		if (PRS.NewRange && PRS.MoveToLBP && PRS.LongWord)
		{
			if (PRS.LineBreakPos.Get(0) !== nEndPos)
				return -1;

			break;
		}
		else if ((true === PRS.NewRange && nPos !== nEndPos) || (nPos === nEndPos && true !== PRS.NewRange))
		{
			return -1;
		}
		else if (nPos === nEndPos && true === PRS.NewRange && true === PRS.MoveToLBP)
		{
			var BreakPos = PRS.LineBreakPos.Get(0);
			if (BreakPos !== nPos)
				return -1;
			else
				oItem.Recalculate_Set_RangeEndPos(PRS, PRS.LineBreakPos, 1);
		}
	}

	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		let oLines = arrSavedLines[nPos - nStartPos];
		let oItem  = this.Content[nPos];

		if (!oLines.Compare(CurLine, CurRange, oItem))
			return -1;

		oItem.LoadRecalculateObject(arrSavedLines[nPos - nStartPos], this);
	}

    // TODO: Здесь пересчеты идут целиком для строки, а не для конкретного отрезка.
    if (!(this.private_RecalculateLineAlign(CurLine, CurPage, PRS, ParaPr, true) & recalcresult_NextElement))
        return -1;

    return CurPage;
};
Paragraph.prototype.private_RecalculatePage = function(CurPage, isFast)
{
	let wrapState = AscWord.ParagraphStatePool.getWrapState();
	wrapState.SetFast(isFast);
	let result = this.private_RecalculatePageInternal(wrapState, CurPage, true);
	AscWord.ParagraphStatePool.release(wrapState);
	return result;
};
Paragraph.prototype.private_RecalculatePageInternal = function(PRS, CurPage, bFirstRecalculate)
{
	PRS.Reset_Page(this, CurPage);

    var Pr     = this.Get_CompiledPr();
    var ParaPr = Pr.ParaPr;

    var CurLine = (CurPage > 0 ? this.Pages[CurPage - 1].EndLine + 1 : 0);
    //-------------------------------------------------------------------------------------------------------------
    // Обрабатываем настройку "не отрывать от следующего"
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculatePageKeepNext(CurPage, PRS, ParaPr))
        return PRS.RecalcResult;

    //-------------------------------------------------------------------------------------------------------------
    // Получаем начальные координаты параграфа
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculatePageXY(CurLine, CurPage, PRS, ParaPr);

    //-------------------------------------------------------------------------------------------------------------
    // Делаем проверки, не нужно ли сразу перенести параграф на новую страницу
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculatePageBreak(CurLine, CurPage, PRS,ParaPr))
	{
		this.Recalculate_PageEndInfo(null, CurPage);
		return PRS.RecalcResult;
	}

    // Изначально обнуляем промежутки обтекания и наличие переноса строки
    PRS.Reset_Ranges();

    if (false !== bFirstRecalculate)
    {
        PRS.ResetMathRecalcInfo();
        PRS.Reset_MathRecalcInfo();
		PRS.SaveFootnotesInfo();
    }
	else
	{
		PRS.LoadFootnotesInfo();
	}

    var RecalcResult;
    while (true)
    {
        PRS.Line = CurLine;
        PRS.RecalcResult = recalcresult_NextLine;
        let complexFieldState = PRS.ComplexFields.getState();

        this.private_RecalculateLine(CurLine, CurPage, PRS, ParaPr);

        RecalcResult = PRS.RecalcResult;

        if (RecalcResult & recalcresult_NextLine)
		{
			// В эту ветку мы попадаем, если строка пересчиталась в нормальном режиме и можно переходить к следующей.
			CurLine++;

			PRS.Reset_Ranges();
			PRS.Reset_RunRecalcInfo();
			PRS.Reset_MathRecalcInfo();
		}
		else if (RecalcResult & recalcresult_ParaMath)
		{
			// В эту ветку попадаем, если нужно заново пересчитать неинлайновую формулу с начала
			CurLine = PRS.resetToMathFirstLine();
			PRS.Reset_RunRecalcInfo();
		}
        else if (RecalcResult & recalcresult_CurLine)
        {
            // В эту ветку мы попадаем, если нам необходимо заново пересчитать данную строку. Такое случается
            // когда у нас появляются плавающие объекты, относительно которых необходимо произвести обтекание.
            // В данном случае мы ничего не делаем, т.к. номер строки не меняется, а новые отрезки обтекания
            // были заполнены при последнем неудачном рассчете.

            PRS.Restore_RunRecalcInfo();
			PRS.ComplexFields.setState(complexFieldState);
        }
        else if (RecalcResult & recalcresult_NextElement || RecalcResult & recalcresult_NextPage)
        {
            // В эту ветку мы попадаем, если мы достигли конца страницы или конца параграфа. Просто выходим
            // из цикла.
            break;
        }
        else if (RecalcResult & recalcresult_CurPagePara)
        {
            // В эту ветку мы попадаем, если в параграфе встретилась картинка, которая находится ниже данного
            // параграфа, и можно пересчитать заново данный параграф.
            RecalcResult = this.private_RecalculatePageInternal(PRS, CurPage, false);
            break;
        }
        else //if (RecalcResult & recalcresult_CurPage || RecalcResult & recalcresult_PrevPage)
        {
            // В эту ветку мы попадаем, если в нашем параграфе встретилось, что-то из-за чего надо пересчитывать
            // эту страницу или предыдущую страницу. Поэтому далее можно ничего не делать, а сообщать верхнему
            // классу об этом.
            return RecalcResult;
        }
    }

    //-------------------------------------------------------------------------------------------------------------
    // Получаем некоторую информацию для следующей страницы (например незакрытые комментарии)
    //-------------------------------------------------------------------------------------------------------------
    this.Recalculate_PageEndInfo(PRS, CurPage);

    return RecalcResult;
};

Paragraph.prototype.private_RecalculatePageKeepNext = function(CurPage, PRS, paraPr)
{
	if (paraPr.PageBreakBefore)
		return true;
	
	let recalcResult = this.RecalculateKeepNext(CurPage);
	if (recalcresult_NextElement !== recalcResult)
	{
		PRS.RecalcResult = recalcResult;
		return false;
	}
	
	return true;
};

Paragraph.prototype.private_RecalculatePageXY          = function(CurLine, CurPage, PRS, ParaPr)
{
    // Если это первая страница параграфа (CurPage = 0), тогда мы должны использовать координаты, которые нам
    // были заданы сверху, а если не первая, тогда координаты мы должны запросить у родительского класса.
    // TODO: Тут отдельно обрабатывается случай, когда рамка переносится на новую страницу, т.е. страница начинается
    //       сразу с рамки. Надо бы не разбивать в данной ситуации рамку на страницы, а просто новую страницу начать
    //       с нее на уровне DocumentContent.

    var XStart, YStart, XLimit, YLimit, oFramePr;
	if (0 === CurPage || ((oFramePr = this.Get_FramePr()) && !oFramePr.IsInline() && this.LogicDocument === this.Parent))
    {
        XStart = this.X;
        YStart = this.Y;
        XLimit = this.XLimit;
        YLimit = this.YLimit;
    }
    else
    {
        var PageStart = this.Parent.Get_PageContentStartPos2(this.PageNum, this.ColumnNum, CurPage, this.Index);

        XStart = PageStart.X;
        YStart = PageStart.Y;
        XLimit = PageStart.XLimit;
        YLimit = PageStart.YLimit;
    }

    PRS.XStart = XStart;
    PRS.YStart = YStart;
    PRS.XLimit = XLimit;// - ParaPr.Ind.Right;
    PRS.YLimit = YLimit;
    PRS.Y      = YStart;

    this.Pages.length   = CurPage + 1;
    this.Pages[CurPage] = new CParaPage(XStart, YStart, XLimit, YLimit, CurLine);
};

Paragraph.prototype.private_RecalculatePageBreak       = function(CurLine, CurPage, PRS, ParaPr)
{
	// Для пустых параграфов с разрывом секции не делаем переноса страницы
	if (undefined !== this.Get_SectionPr() && true === this.IsEmpty())
		return true;

	var isParentDocument = this.Parent instanceof CDocument;
	var isParentBlockSdt = this.Parent instanceof CDocumentContent
		&& this.Parent.Parent instanceof CBlockLevelSdt
		&& PRS.GetTopDocument() instanceof CDocument
		&& !PRS.IsInTable();

    if (isParentDocument || isParentBlockSdt)
    {
        // Начинаем параграф с новой страницы
        var PageRelative = this.private_GetRelativePageIndex(CurPage) - this.Get_StartPage_Relative();
        if (0 === PageRelative && true === ParaPr.PageBreakBefore)
        {
            // Если это первый элемент документа или секции, тогда не надо начинать его с новой страницы.
            // Кроме случая, когда у нас разрыв секции на текущей странице. Также не добавляем разрыв страницы для
            // особого пустого параграфа с разрывом секции.

            var bNeedPageBreak = true;

			var Prev = this.Get_DocumentPrev();
			if (!Prev && isParentBlockSdt)
			{
				var oSdt = this.Parent.Parent;
				while (oSdt instanceof CBlockLevelSdt)
				{
					Prev = oSdt.Get_DocumentPrev();
					if (Prev)
						break;

					if (oSdt.Parent instanceof CDocumentContent && oSdt.Parent.Parent instanceof CBlockLevelSdt)
						oSdt = oSdt.Parent.Parent;
					else
						oSdt = null;
				}
			}

			while (Prev && (Prev instanceof CBlockLevelSdt))
				Prev = Prev.GetLastElement();

            if ((true === this.IsEmpty() && undefined !== this.Get_SectionPr()) || null === Prev)
            {
                bNeedPageBreak = false;
            }
            else if (this.Parent === this.LogicDocument && type_Paragraph === Prev.GetType() && undefined !== Prev.Get_SectionPr())
            {
                var PrevSectPr = Prev.Get_SectionPr();
                var CurSectPr = this.LogicDocument.SectionsInfo.Get_SectPr(this.Index).SectPr;
                if (c_oAscSectionBreakType.Continuous !== CurSectPr.Get_Type() || true !== CurSectPr.Compare_PageSize(PrevSectPr))
                    bNeedPageBreak = false;
            }

            if (true === bNeedPageBreak)
            {
                // Добавляем разрыв страницы
                this.Pages[CurPage].Set_EndLine(CurLine - 1);

                if (0 === CurLine)
                    this.Lines[-1] = new CParaLine();

                PRS.RecalcResult = recalcresult_NextPage | recalcresultflags_Column;
                return false;
            }
        }
        else if (isParentDocument && true === this.Parent.RecalcInfo.Check_KeepNext(this) && 0 === CurPage && null != this.Get_DocumentPrev())
        {
            this.Pages[CurPage].Set_EndLine( CurLine - 1 );
            if ( 0 === CurLine )
                this.Lines[-1] = new CParaLine( 0 );

            PRS.RecalcResult = recalcresult_NextPage;
            return false;
        }
        else if (true === this.Is_Inline()) // Случай Flow разбирается в Document.js
        {
            // Проверяем PageBreak и ColumnBreak в предыдущей строке
            var isPageBreakOnPrevLine   = false;
            var isColumnBreakOnPrevLine = false;

            var PrevElement = this.Get_DocumentPrev();
            if (!PrevElement && isParentBlockSdt)
			{
				var oSdt = this.Parent.Parent;
				while (oSdt instanceof CBlockLevelSdt)
				{
					PrevElement = oSdt.Get_DocumentPrev();
					if (PrevElement)
						break;

					if (oSdt.Parent instanceof CDocumentContent && oSdt.Parent.Parent instanceof CBlockLevelSdt)
						oSdt = oSdt.Parent.Parent;
					else
						oSdt = null;
				}
			}

			while (PrevElement && (PrevElement instanceof CBlockLevelSdt))
				PrevElement = PrevElement.GetLastElement();

			var oFootnotes = this.LogicDocument ? this.LogicDocument.Footnotes : null;
			if (null !== PrevElement && type_Paragraph === PrevElement.Get_Type() && true === PrevElement.Is_Empty() && undefined !== PrevElement.Get_SectionPr())
			{
				var PrevSectPr = PrevElement.Get_SectionPr();
				var CurSectPr  = this.LogicDocument.SectionsInfo.Get_SectPr(this.Index).SectPr;
				if (c_oAscSectionBreakType.Continuous === CurSectPr.Get_Type() && true === CurSectPr.Compare_PageSize(PrevSectPr) && oFootnotes && oFootnotes.IsEmptyPage(PrevElement.GetAbsolutePage(PrevElement.GetPagesCount() - 1)))
					PrevElement = PrevElement.Get_DocumentPrev();
			}

			if (0 !== CurPage && true !== this.Check_EmptyPages(CurPage - 1))
			{
				var EndLine = this.Pages[CurPage - 1].EndLine;
				if (-1 !== EndLine && this.Lines[EndLine].Info & paralineinfo_BreakRealPage)
					isPageBreakOnPrevLine = true;
			}
			else if (null !== PrevElement && type_Paragraph === PrevElement.Get_Type())
			{
				var bNeedPageBreak = true;
				if (type_Paragraph === PrevElement.GetType() && undefined !== PrevElement.Get_SectionPr())
				{
					var PrevSectPr = PrevElement.Get_SectionPr();
					var CurSectPr  = this.LogicDocument.SectionsInfo.Get_SectPr(this.Index).SectPr;
					if (c_oAscSectionBreakType.Continuous !== CurSectPr.Get_Type() || true !== CurSectPr.Compare_PageSize(PrevSectPr) || (oFootnotes && !oFootnotes.IsEmptyPage(PrevElement.GetAbsolutePage(PrevElement.GetPagesCount() - 1))))
						bNeedPageBreak = false;
				}

				if (true === bNeedPageBreak)
				{
					var EndLine = PrevElement.Pages[PrevElement.Pages.length - 1].EndLine;
					if (-1 !== EndLine && PrevElement.Lines[EndLine].Info & paralineinfo_BreakRealPage)
						isPageBreakOnPrevLine = true;
				}
			}

            // ColumnBreak для случая CurPage > 0 не разбираем здесь, т.к. он срабатывает автоматически
            if (0 === CurPage && null !== PrevElement && type_Paragraph === PrevElement.Get_Type())
            {
                var EndLine = PrevElement.Pages[PrevElement.Pages.length - 1].EndLine;
                if (-1 !== EndLine && !(PrevElement.Lines[EndLine].Info & paralineinfo_BreakRealPage) && PrevElement.Lines[EndLine].Info & paralineinfo_BreakPage)
                    isColumnBreakOnPrevLine = true;
            }

            // Здесь используем GetAbsoluteColumn, а не у текущего класса, т.к. в данную ветку мы попадаем только, если
			// верхний DocContent - документ и если не в таблице
            if ((true === isPageBreakOnPrevLine && (0 !== this.GetAbsoluteColumn(CurPage) || (0 === CurPage && null !== PrevElement)))
                || (true === isColumnBreakOnPrevLine && 0 === CurPage))
            {
                this.Pages[CurPage].Set_EndLine(CurLine - 1);

                if (0 === CurLine)
                    this.Lines[-1] = new CParaLine();

                PRS.RecalcResult = recalcresult_NextPage | recalcresultflags_Column;
                return false;
            }
        }
    }

    return true;
};

Paragraph.prototype.private_RecalculateLine            = function(CurLine, CurPage, PRS, ParaPr)
{
    // При пересчете любой строки обновляем эти поля
    this.ParaEnd.Line  = -1;
    this.ParaEnd.Range = -1;

    //-------------------------------------------------------------------------------------------------------------
    // 1. Добавляем новую строку в параграф
    //-------------------------------------------------------------------------------------------------------------
    this.Lines.length   = CurLine + 1;
    this.Lines[CurLine] = new CParaLine();
	this.Lines[CurLine].CF = PRS.ComplexFields.getState();

    //-------------------------------------------------------------------------------------------------------------
    // 2. Проверяем, является ли данная строка висячей
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineWidow(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 3. Заполняем строку отрезками обтекания
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculateLineFillRanges(CurLine, CurPage, PRS, ParaPr);

    //-------------------------------------------------------------------------------------------------------------
    // 4. Пересчитываем отрезки данной строки
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineRanges(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 5. Заполняем информацию о строке
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculateLineInfo(CurLine, CurPage, PRS, ParaPr);

    //-------------------------------------------------------------------------------------------------------------
    // 6. Пересчитываем метрики данной строки
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculateLineMetrics(CurLine, CurPage, PRS, ParaPr);

    //-------------------------------------------------------------------------------------------------------------
    // 7. Рассчитываем высоту строки, а также положение верхней и нижней границ
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculateLinePosition(CurLine, CurPage, PRS, ParaPr);

    //-------------------------------------------------------------------------------------------------------------
    // 8. Проверяем достигла ли данная строка конца страницы
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineBottomBound(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 9. Проверяем обтекание данной строки относительно плавающих объектов
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineCheckRanges(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 10. Выставляем вертикальное смещение данной строки
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculateLineBaseLine(CurLine, CurPage, PRS, ParaPr);

    //-------------------------------------------------------------------------------------------------------------
    // 11. Проверяем не съехала ли вся строка из-за обтекания
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineCheckRangeY(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 12. Пересчитываем сдвиги элементов внутри параграфа и видимые ширины пробелов, в зависимости от align.
    //-------------------------------------------------------------------------------------------------------------
    if (!(this.private_RecalculateLineAlign(CurLine, CurPage, PRS, ParaPr, false) & recalcresult_NextElement))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 13. Последние проверки
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineEnd(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 14. Проверяем Последние проверки
    //-------------------------------------------------------------------------------------------------------------
    if (false === this.private_RecalculateLineCheckFootnotes(CurLine, CurPage, PRS, ParaPr))
        return;

    //-------------------------------------------------------------------------------------------------------------
    // 15. Регистрируем концевые сноски на странице
    //-------------------------------------------------------------------------------------------------------------
    this.private_RecalculateLineCheckEndnotes(CurLine, CurPage, PRS, ParaPr);
};

Paragraph.prototype.private_RecalculateLineWidow       = function(CurLine, CurPage, PRS, ParaPr)
{
    // Висячие строки обрабатываются только внутри основного документа
    if ( this.Parent instanceof CDocument && true === this.Parent.RecalcInfo.Check_WidowControl(this, CurLine) )
    {
        this.Parent.RecalcInfo.Need_ResetWidowControl();

        this.Pages[CurPage].Set_EndLine(CurLine - 1);
        if (0 === CurLine)
        {
            this.Lines[-1] = new CParaLine(0);
        }

        PRS.RecalcResult = recalcresult_NextPage | recalcresultflags_Column;
        return false;
    }

    return true;
};

Paragraph.prototype.private_RecalculateLineFillRanges  = function(CurLine, CurPage, PRS, paraPr)
{
    this.Lines[CurLine].Info = 0;

    // Параметры Ranges и RangesCount не обнуляются здесь, они задаются выше
    var Ranges      = PRS.Ranges;
    var RangesCount = PRS.RangesCount;

    // Обнуляем параметры PRS для строки
    PRS.Reset_Line();

    // Проверим, нужно ли в данной строке учитывать FirstLine (т.к. не всегда это первая строка должна быть)
    var UseFirstLine = true;
    for ( var TempCurLine = CurLine - 1; TempCurLine >= 0; TempCurLine-- )
    {
        var TempInfo = this.Lines[TempCurLine].Info;
        if (!(TempInfo & paralineinfo_BreakPage) || !(TempInfo & paralineinfo_Empty))
        {
            UseFirstLine = false;
            break;
        }
    }

    // Проверим неинлайн формулу в первой строке
    if (0 === CurLine && true === UseFirstLine)
    {
        var CurPos = 0;
        var Count = this.Content.length;
        while (CurPos < Count)
        {
            if (true === this.CheckMathPara(CurPos))
            {
                UseFirstLine = false;
                break;
            }
            else if (true !== this.Content[CurPos].Is_Empty())
            {
                break;
            }

            CurPos++;
        }
    }

    PRS.UseFirstLine = UseFirstLine;
	
	// Заполняем строку отрезками обтекания. Выставляем начальные сдвиги для отрезков. Начало промежутка = конец вырезаемого промежутка
	this.Lines[CurLine].Reset();
	if (paraPr.Bidi)
	{
		let xStart = PRS.XStart + paraPr.Ind.Right;
		
		let x0 = RangesCount > 0 ? Ranges[Ranges.length - 1].X1 : xStart;
		let x1 = (UseFirstLine ? PRS.XLimit - paraPr.Ind.Left - paraPr.Ind.FirstLine : PRS.XLimit - paraPr.Ind.Left);
		this.Lines[CurLine].addRange(x0, x1);
		for (let rangeIndex = Ranges.length - 1; rangeIndex >= 0; --rangeIndex)
		{
			x0 = rangeIndex > 0 ? Ranges[rangeIndex - 1].X1 : xStart;
			x1 = Ranges[rangeIndex].X0;
			this.Lines[CurLine].addRange(x0, x1);
		}

	}
	else
	{
		let xLimit = PRS.XLimit - paraPr.Ind.Right;
		
		let x0 = (UseFirstLine ? PRS.XStart + paraPr.Ind.Left + paraPr.Ind.FirstLine : PRS.XStart + paraPr.Ind.Left);
		let x1 = RangesCount > 0 ? Ranges[0].X0 : xLimit;
		this.Lines[CurLine].addRange(x0, x1);
		for (let rangeIndex = 1, rangeCount = Ranges.length; rangeIndex <= rangeCount; ++rangeIndex)
		{
			x0 = Ranges[rangeIndex - 1].X1
			x1 = rangeIndex === RangesCount ? xLimit : Ranges[rangeIndex].X0;
			this.Lines[CurLine].addRange(x0, x1);
		}
	}

    if (true === PRS.RangeY)
    {
        PRS.RangeY = false;
        this.Lines[CurLine].Info |= paralineinfo_RangeY;
    }
};

Paragraph.prototype.private_RecalculateLineRanges      = function(CurLine, CurPage, PRS, ParaPr)
{
    var RangesCount = PRS.RangesCount;
    var CurRange = 0;
    while ( CurRange <= RangesCount )
    {
        PRS.Range = CurRange;
        this.private_RecalculateRange(CurRange, CurLine, CurPage, RangesCount, PRS, ParaPr);

        if (PRS.isForceLineBreak())
        {
            // Поскольку мы выходим досрочно из цикла, нам надо удалить лишние отрезки обтекания
            this.Lines[CurLine].Ranges.length = CurRange + 1;
            break;
        }

        if ( -1 === this.ParaEnd.Line && true === PRS.End )
        {
            this.ParaEnd.Line  = CurLine;
            this.ParaEnd.Range = CurRange;
        }

        // Такое может случиться, если мы насильно переносим автофигуру на следующую страницу
        if (PRS.RecalcResult & recalcresult_NextPage || PRS.RecalcResult & recalcresult_ParaMath || PRS.RecalcResult & recalcresult_CurLine || PRS.RecalcResult & recalcresult_CurPagePara)
            return false;

        CurRange++;
    }

    return true;
};

Paragraph.prototype.private_RecalculateLineInfo        = function(CurLine, CurPage, PRS, ParaPr)
{
    if (true === PRS.BreakPageLine)
        this.Lines[CurLine].Info |= paralineinfo_BreakPage;

    if (true === PRS.BreakRealPageLine)
        this.Lines[CurLine].Info |= paralineinfo_BreakRealPage;

    if (true === PRS.EmptyLine)
        this.Lines[CurLine].Info |= paralineinfo_Empty;

    if (true === PRS.End)
        this.Lines[CurLine].Info |= paralineinfo_End;

    if (true === PRS.BadLeftTab)
        this.Lines[CurLine].Info |= paralineinfo_BadLeftTab;

    if (PRS.GetFootnoteReferencesCount(null, true) > 0 || PRS.GetEndnoteReferenceCount() > 0)
    	this.Lines[CurLine].Info |= paralineinfo_Notes;

    if (true === PRS.TextOnLine)
    	this.Lines[CurLine].Info |= paralineinfo_TextOnLine;

    if (true === PRS.BreakLine)
    	this.Lines[CurLine].Info |= paralineinfo_BreakLine;

	if (PRS.LongWord)
		this.Lines[CurLine].Info |= paralineinfo_LongWord;
	
	if (PRS.LastHyphenItem)
		this.Lines[CurLine].Info |= paralineinfo_AutoHyphen;
};

Paragraph.prototype.private_RecalculateLineMetrics     = function(CurLine, CurPage, PRS, ParaPr)
{
    var Line = this.Lines[CurLine];
    var RangesCount = Line.Ranges.length;

    for (var CurRange = 0; CurRange < RangesCount; CurRange++)
    {
        var Range = Line.Ranges[CurRange];

        var StartPos = Range.StartPos;
        var EndPos   = Range.EndPos;

        for (var Pos = StartPos; Pos <= EndPos; Pos++)
        {
            this.Content[Pos].Recalculate_LineMetrics(PRS, ParaPr, CurLine, CurRange);
        }
    }

    // Строка пустая, у нее надо выставить ненулевую высоту. Делаем как Word, выставляем высоту по размеру
    // текста, на котором закончилась данная строка.
    if ( true === PRS.EmptyLine || (PRS.LineAscent < 0.001 && PRS.LineDescent < 0.001) || (true === PRS.End && true !== PRS.TextOnLine))
	{
		var LastItem = (true === PRS.End ? this.Content[this.Content.length - 1] : this.Content[this.Lines[CurLine].Ranges[this.Lines[CurLine].Ranges.length - 1].EndPos]);

		if (true === PRS.End)
		{
			// TODO: Как только переделаем para_End переделать тут
			let oTextPr  = this.GetParaEndCompiledPr();
			let oMetrics = oTextPr.GetTextMetrics(oTextPr.CS || oTextPr.RTL ? AscWord.fontslot_CS : AscWord.fontslot_ASCII, this.GetTheme());

			let EndTextDescent = oMetrics.Descent;
			let EndTextAscent  = oMetrics.Ascent + oMetrics.LineGap;
			let EndTextAscent2 = oMetrics.Ascent;

			PRS.LineTextAscent  = EndTextAscent;
			PRS.LineTextAscent2 = EndTextAscent2;
			PRS.LineTextDescent = EndTextDescent;

			if (PRS.LineAscent < EndTextAscent)
				PRS.LineAscent = EndTextAscent;

			if (PRS.LineDescent < EndTextDescent)
				PRS.LineDescent = EndTextDescent;
		}
		else if (undefined !== LastItem)
		{
			let lastRun = LastItem.Get_LastRunInRange(PRS.Line, PRS.Range);
			if (lastRun && lastRun instanceof AscWord.CRun)
			{
				let metrics = lastRun.getTextMetrics();
				
				let textDescent = metrics.Descent;
				let textAscent  = metrics.Ascent + metrics.LineGap;
				let textAscent2 = metrics.Ascent;
				
				if (PRS.LineTextAscent < textAscent)
					PRS.LineTextAscent = textAscent;

				if (PRS.LineTextAscent2 < textAscent2)
					PRS.LineTextAscent2 = textAscent2;

				if (PRS.LineTextDescent < textDescent)
					PRS.LineTextDescent = textDescent;

				if (PRS.LineAscent < textAscent)
					PRS.LineAscent = textAscent;

				if (PRS.LineDescent < textDescent)
					PRS.LineDescent = textDescent;
			}
		}
	}

    // Рассчитаем метрики строки
    this.Lines[CurLine].Metrics.Update( PRS.LineTextAscent, PRS.LineTextAscent2, PRS.LineTextDescent, PRS.LineAscent, PRS.LineDescent, ParaPr );

	if (true === PRS.End && true !== PRS.EmptyLine && true !== PRS.TextOnLine && Math.abs(this.Lines[CurLine].Metrics.Descent - this.Lines[CurLine].Metrics.TextDescent) < 0.001)
		this.Lines[CurLine].Metrics.Descent = 0;
};

Paragraph.prototype.private_RecalculateLinePosition    = function(CurLine, CurPage, PRS, ParaPr)
{
	// Важно: Значение Border.Space учитывается всегда, даже когда Border.Value = none, а
	//        вот Border.Size зависит уже от Border.Value

    var BaseLineOffset = 0;
    if (CurLine === this.Pages[CurPage].FirstLine)
    {
    	var oForm = null;
    	if (this.IsInFixedForm() && (oForm = this.GetInnerForm()) && oForm.IsMultiLineForm() && oForm.IsTextForm())
		{
			var oRun = oForm.GetElement(0);
			if (oRun && oRun instanceof ParaRun)
			{
				// Адобовский вариант отступа первой строки для многострочных форм
				var oTextPr = oRun.Get_CompiledTextPr(false);
				g_oTextMeasurer.SetTextPr(oTextPr, this.GetTheme());
				g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, 1);
				var oLimits = g_oTextMeasurer.GetLimitsY();
				var nBBoxH  = oLimits.max - oLimits.min + 2 * 25.4 / 72;

				if (this.Lines[CurLine].Metrics.Ascent < nBBoxH)
					this.Lines[CurLine].Metrics.Ascent = nBBoxH;
			}
		}

    	BaseLineOffset = this.Lines[CurLine].Metrics.Ascent;

        if (this.Check_FirstPage(CurPage, true))
		{
			// Добавляем расстояние до параграфа (Pr.Spacing.Before)
			if (this.private_CheckNeedBeforeSpacing(CurPage, PRS.Parent, PRS.GetPageAbs(), ParaPr))
				BaseLineOffset += ParaPr.Spacing.Before;

			// Добавляем толщину границы параграфа (если граница задана)
			if ((true === ParaPr.Brd.First || 1 === CurPage))
			{
				BaseLineOffset += ParaPr.Brd.Top.Space;
				if (border_Single === ParaPr.Brd.Top.Value)
					BaseLineOffset += ParaPr.Brd.Top.Size;
			}
			else if (false === ParaPr.Brd.First)
			{
				BaseLineOffset += ParaPr.Brd.Between.Space;
				if (border_Single === ParaPr.Brd.Between.Value)
					BaseLineOffset += ParaPr.Brd.Between.Size;
			}
		}

        PRS.BaseLineOffset = BaseLineOffset;
    }
    else
	{
		if (this.Lines[CurLine].Info & paralineinfo_RangeY)
			PRS.BaseLineOffset = this.Lines[CurLine].Metrics.Ascent;
		else
			BaseLineOffset = PRS.BaseLineOffset;
	}

    var Top, Bottom;
    var Top2, Bottom2; // верх и низ без Pr.Spacing

    var PrevBottom = this.Pages[CurPage].Bounds.Bottom;

    if (this.Lines[CurLine].Info & paralineinfo_RangeY)
    {
        Top  = PRS.Y;
        Top2 = PRS.Y;

        if (CurLine === this.Pages[CurPage].FirstLine && this.Check_FirstPage(CurPage, true))
		{
			if (this.private_CheckNeedBeforeSpacing(CurPage, PRS.Parent, PRS.GetPageAbs(), ParaPr))
			{
				Top2    = Top + ParaPr.Spacing.Before;
				Bottom2 = Top + ParaPr.Spacing.Before + this.Lines[0].Metrics.Ascent + this.Lines[0].Metrics.Descent;

				if (true === ParaPr.Brd.First)
				{
					Top2 += ParaPr.Brd.Top.Space;
					Bottom2 += ParaPr.Brd.Top.Space;
					if (border_Single === ParaPr.Brd.Top.Value)
					{
						Top2 += ParaPr.Brd.Top.Size;
						Bottom2 += ParaPr.Brd.Top.Size;
					}
				}
				else if (false === ParaPr.Brd.First)
				{
					Top2 += ParaPr.Brd.Between.Space;
					Bottom2 += ParaPr.Brd.Between.Space;

					if (border_Single === ParaPr.Brd.Between.Value)
					{
						Top2 += ParaPr.Brd.Between.Size;
						Bottom2 += ParaPr.Brd.Between.Size;
					}
				}
			}
			else
			{
				// Параграф начинается с новой страницы
				Bottom2 = Top + this.Lines[0].Metrics.Ascent + this.Lines[0].Metrics.Descent;

				Top2 += ParaPr.Brd.Top.Space;
				Bottom2 += ParaPr.Brd.Top.Space;

				if (border_Single === ParaPr.Brd.Top.Value)
				{
					Top2 += ParaPr.Brd.Top.Size;
					Bottom2 += ParaPr.Brd.Top.Size;
				}
			}
		}
        else
        {
            Bottom2 = Top + this.Lines[CurLine].Metrics.Ascent + this.Lines[CurLine].Metrics.Descent;
        }
    }
    else
	{
		if (CurLine !== this.Pages[CurPage].FirstLine || !this.Check_FirstPage(CurPage, true))
		{
			if (CurLine !== this.Pages[CurPage].FirstLine)
			{
				Top     = PRS.Y + BaseLineOffset + this.Lines[CurLine - 1].Metrics.Descent + this.Lines[CurLine - 1].Metrics.LineGap;
				Top2    = Top;
				Bottom2 = Top + this.Lines[CurLine].Metrics.Ascent + this.Lines[CurLine].Metrics.Descent;
			}
			else
			{
				Top     = this.Pages[CurPage].Y;
				Top2    = Top;
				Bottom2 = Top + this.Lines[CurLine].Metrics.Ascent + this.Lines[CurLine].Metrics.Descent;
			}
		}
		else
		{
			Top  = PRS.Y;
			Top2 = PRS.Y;

			if (this.private_CheckNeedBeforeSpacing(CurPage, PRS.Parent, PRS.GetPageAbs(), ParaPr))
			{
				Top2    = Top + ParaPr.Spacing.Before;
				Bottom2 = Top + ParaPr.Spacing.Before + this.Lines[CurLine].Metrics.Ascent + this.Lines[CurLine].Metrics.Descent;

				if (true === ParaPr.Brd.First)
				{
					Top2 += ParaPr.Brd.Top.Space;
					Bottom2 += ParaPr.Brd.Top.Space;

					if (border_Single === ParaPr.Brd.Top.Value)
					{
						Top2 += ParaPr.Brd.Top.Size;
						Bottom2 += ParaPr.Brd.Top.Size;
					}
				}
				else if (false === ParaPr.Brd.First)
				{
					Top2 += ParaPr.Brd.Between.Space;
					Bottom2 += ParaPr.Brd.Between.Space;

					if (border_Single === ParaPr.Brd.Between.Value)
					{
						Top2 += ParaPr.Brd.Between.Size;
						Bottom2 += ParaPr.Brd.Between.Size;
					}
				}
			}
			else
			{
				// Параграф начинается с новой страницы
				Bottom2 = Top + this.Lines[CurLine].Metrics.Ascent + this.Lines[CurLine].Metrics.Descent;

				Top2 += ParaPr.Brd.Top.Space;
				Bottom2 += ParaPr.Brd.Top.Space;

				if (border_Single === ParaPr.Brd.Top.Value)
				{
					Top2 += ParaPr.Brd.Top.Size;
					Bottom2 += ParaPr.Brd.Top.Size;
				}
			}
		}
	}

    Bottom  = Bottom2;
    Bottom += this.Lines[CurLine].Metrics.LineGap;
	
	if (this.Lines[CurLine].Metrics.LineGap < 0)
		Bottom2 += this.Lines[CurLine].Metrics.LineGap;

    // Если данная строка последняя, тогда подкорректируем нижнюю границу
    if ( true === PRS.End )
	{
		Bottom += ParaPr.Spacing.After;

		// Если нижняя граница Between, тогда она учитывается в следующем параграфе
		if (true === ParaPr.Brd.Last)
		{
			Bottom += ParaPr.Brd.Bottom.Space;

			if (border_Single === ParaPr.Brd.Bottom.Value)
				Bottom += ParaPr.Brd.Bottom.Size;
		}
		else
		{
			Bottom += ParaPr.Brd.Between.Space;
		}

		// TODO: Здесь нужно сделать корректировку YLimit с учетом сносок. Надо разобраться почему вообще здесь
		// используется this.YLimit вместо Page.YLimit

		if (!this.Parent.IsCalculatingContinuousSectionBottomLine() && false === this.IsTableCellContent() && Bottom > this.YLimit && Bottom - this.YLimit <= ParaPr.Spacing.After)
			Bottom = this.YLimit;
	}

	this.Lines[CurLine].Top    = Top    - this.Pages[CurPage].Y;
	this.Lines[CurLine].Bottom = Bottom - this.Pages[CurPage].Y;

	// В MSWord версиях 14 и ниже пустая строка с переносом колонки не имеет высоты
	// Заметим, что границы строки мы оставляем корректными
	if (PRS.getCompatibilityMode() <= AscCommon.document_compatibility_mode_Word14
		&& this.Lines[CurLine].Info & paralineinfo_BreakPage
		&& this.Lines[CurLine].Info & paralineinfo_Empty
		&& !(this.Lines[CurLine].Info & paralineinfo_BreakRealPage))
	{
		Bottom  = Top;
		Top2    = Top;
		Bottom2 = Top;
	}

	// Верхнюю границу мы сохраняем только для первой строки данной страницы
    if (CurLine === this.Pages[CurPage].FirstLine && !(this.Lines[CurLine].Info & paralineinfo_RangeY))
        this.Pages[CurPage].Bounds.Top = Top;

    this.Pages[CurPage].Bounds.Bottom = Bottom;

    PRS.LineTop        = AscCommon.CorrectMMToTwips(Top);
    PRS.LineBottom     = AscCommon.CorrectMMToTwips(Bottom);
    PRS.LineTop2       = AscCommon.CorrectMMToTwips(Top2);
    PRS.LineBottom2    = AscCommon.CorrectMMToTwips(Bottom2);
    PRS.LinePrevBottom = AscCommon.CorrectMMToTwips(PrevBottom);
};

Paragraph.prototype.private_RecalculateLineBottomBound = function(CurLine, CurPage, PRS, ParaPr)
{
    var Top     = PRS.LineTop;
    var Bottom2 = PRS.LineBottom2;

    // В ячейке перенос страницы происходит по нижней границе, т.е. с учетом Spacing.After и границы
    if ( true === this.IsTableCellContent() )
        Bottom2 = PRS.LineBottom;

    // Переносим строку по PageBreak. Если в строке ничего нет кроме PageBreak, и это не конец параграфа, тогда нам не надо проверять высоту строки и обтекание.
    var LineInfo = this.Lines[CurLine].Info;
    var BreakPageLineEmpty = (LineInfo & paralineinfo_BreakPage && LineInfo & paralineinfo_Empty && !(LineInfo & paralineinfo_End) ? true : false);
    PRS.BreakPageLineEmpty = BreakPageLineEmpty;

    var RealCurPage = this.private_GetRelativePageIndex(CurPage) - this.Get_StartPage_Relative();

    var YLimit = PRS.YLimit;
    var oTopDocument = PRS.TopDocument;
	var bNoFootnotes = true;
	if (oTopDocument instanceof CDocument)
	{
		// bNoFootnotes - означает есть или нет сноска на данной колонке
		var nHeight =  oTopDocument.Footnotes.GetHeight(PRS.PageAbs, PRS.ColumnAbs);
		if (nHeight > 0.001)
		{
			bNoFootnotes = false;

			// В таблицах граница разруливается по своему
			if (true !== PRS.IsInTable())
				YLimit -= nHeight;
		}
	}
    else if (oTopDocument instanceof CFootEndnote)
	{
		// bNoFootnotes - означает, первая или нет данная сноска в колонке. Если она не первая,
		// тогда если у нее не убирается первая строка первого параграфа, все равно надо делать перенос
		var oController = oTopDocument.GetParent();
		if (oController instanceof CEndnotesController || !oController.IsEmptyPageColumn(PRS.PageAbs, PRS.ColumnAbs, oTopDocument.GetSectionIndex()))
			bNoFootnotes = false;
	}

    // Сначала проверяем не нужно ли сделать перенос страницы в данном месте
    // Перенос не делаем, если это первая строка на новой странице
    if (true === this.UseLimit()
		&& (Top > YLimit || Bottom2 > YLimit)
		&& (CurLine != this.Pages[CurPage].FirstLine
		|| false === bNoFootnotes
		|| (0 === RealCurPage && ((null != this.Get_DocumentPrev() && !this.Parent.IsElementStartOnNewPage(this.GetIndex()))
		|| (true === this.IsTableCellContent() && true !== this.Parent.IsTableFirstRowOnNewPage())
		|| (true === this.Parent.IsBlockLevelSdtContent() && true !== this.Parent.IsBlockLevelSdtFirstOnNewPage()))))
		&& false === BreakPageLineEmpty)
    {
		this.private_RecalculateMoveLineToNextPage(CurLine, CurPage, PRS, ParaPr);
		return false;
    }

    return true;
};

Paragraph.prototype.private_RecalculateLineCheckRanges = function(CurLine, CurPage, PRS, ParaPr)
{
    var Top     = PRS.LineTop;
    var Bottom  = PRS.LineBottom;
    var Top2    = PRS.LineTop2;
    var Bottom2 = PRS.LineBottom2;
	
	var Left  = this.Pages[CurPage].X;
	var Right = this.Pages[CurPage].XLimit;
	if (!PRS.MathNotInline)
	{
		if (ParaPr.Bidi)
		{
			Left += ParaPr.Ind.Right;
			Right -= PRS.UseFirstLine ? ParaPr.Ind.Left + ParaPr.Ind.FirstLine : ParaPr.Ind.Left;
		}
		else
		{
			Right -= ParaPr.Ind.Right;
			Left  += PRS.UseFirstLine ? ParaPr.Ind.Left + ParaPr.Ind.FirstLine : ParaPr.Ind.Left;
		}
	}

	var PageFields = null;
    if (this.bFromDocument && PRS.GetTopDocument() === this.LogicDocument && !PRS.IsInTable())
	{
		// Заглушка для случая, когда параграф лежит в CBlockLevelSdt
		PageFields = this.LogicDocument.Get_ColumnFields(PRS.GetTopIndex(), this.Get_AbsoluteColumn(CurPage), this.GetAbsolutePage(CurPage));
	}
	else
	{
		PageFields = this.Parent.Get_ColumnFields ? this.Parent.Get_ColumnFields(this.Get_Index(), this.Get_AbsoluteColumn(CurPage), this.GetAbsolutePage(CurPage)) : this.Parent.Get_PageFields(this.private_GetRelativePageIndex(CurPage), this.Parent.IsHdrFtr());
	}

    var Ranges = PRS.Ranges;
    var Ranges2;

    for (var nIndex = 0, nCount = Ranges.length; nIndex < nCount; ++nIndex)
	{
		Ranges[nIndex].Y1 = AscCommon.CorrectMMToTwips(Ranges[nIndex].Y1);
	}

	if (PRS.getCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15)
	{
		Bottom = Bottom2;
		Top2   = Top;
	}

    if ( true === this.Use_Wrap() )
        Ranges2 = this.Parent.CheckRange(Left, Top, Right, Bottom, Top2, Bottom2, PageFields.X, PageFields.XLimit, this.private_GetRelativePageIndex(CurPage), true, PRS.MathNotInline);
    else
        Ranges2 = [];

    // Проверяем совпали ли промежутки. Если совпали, тогда данная строчка рассчитана верно, и мы переходим к
    // следующей, если нет, тогда заново рассчитываем данную строчку, но с новыми промежутками.
    // Заметим, что тут возможен случай, когда Ranges2 меньше, чем Ranges, такое может случится при повторном
    // обсчете строки. (После первого расчета мы выяснили что Ranges < Ranges2, при повторном обсчете строки, т.к.
    // она стала меньше, то у нее и рассчитанная высота могла уменьшиться, а значит Ranges2 могло оказаться
    // меньше чем Ranges). В таком случае не надо делать повторный пересчет, иначе будет зависание.
    if (-1 === FlowObjects_CompareRanges(Ranges, Ranges2) && true === FlowObjects_CheckInjection(Ranges, Ranges2) && false === PRS.BreakPageLineEmpty)
    {
        // Выставляем новые отрезки обтекания и сообщаем, что надо заново персчитать данную строку
        PRS.Ranges       = Ranges2;
        PRS.RangesCount  = Ranges2.length;
        PRS.RecalcResult = recalcresult_CurLine;

        if (this.Lines[CurLine].Info & paralineinfo_RangeY)
            PRS.RangeY = true;

        return false;
    }

    return true;
};

Paragraph.prototype.private_RecalculateLineBaseLine    = function(CurLine, CurPage, PRS, ParaPr)
{
    if (this.Lines[CurLine].Info & paralineinfo_RangeY)
    {
        this.Lines[CurLine].Y = PRS.Y - this.Pages[CurPage].Y;
    }
    else
    {
        if (CurLine > 0)
        {
            // Первая линия на странице не должна двигаться
            if (CurLine != this.Pages[CurPage].FirstLine && ( true === PRS.End || true !== PRS.EmptyLine || PRS.RangesCount <= 0 || true === PRS.NewPage  ))
                PRS.Y += this.Lines[CurLine - 1].Metrics.Descent + this.Lines[CurLine - 1].Metrics.LineGap + this.Lines[CurLine].Metrics.Ascent;

            this.Lines[CurLine].Y = PRS.Y - this.Pages[CurPage].Y;
        }
        else
            this.Lines[0].Y = 0;
    }

    this.Lines[CurLine].Y += PRS.BaseLineOffset;
    if (this.Lines[CurLine].Metrics.LineGap < 0)
        this.Lines[CurLine].Y += this.Lines[CurLine].Metrics.LineGap;
};

Paragraph.prototype.private_RecalculateLineCheckRangeY = function(CurLine, CurPage, PRS, ParaPr)
{
    // Такое случается, когда у нас после пересчета Flow картинки, место к которому она была привязана перешло на
    // следующую страницу.
    if (PRS.RecalcResult & recalcresult_NextPage)
        return false;

    // Если строка пустая в следствии того, что у нас было обтекание, тогда мы не добавляем новую строку,
    // а просто текущую смещаем ниже.

	if (true === PRS.EmptyLine && true === PRS.bMathRangeY) // нужный PRS.Y выставляется в ParaMath
	{
		PRS.bMathRangeY = false;

		// Отмечаем, что данная строка переносится по Y из-за обтекания
		PRS.RangeY = true;

		// Пересчитываем заново данную строку
		PRS.Reset_Ranges();
		PRS.RecalcResult = recalcresult_CurLine;

		return false;
	}
    else if (true !== PRS.End && true === PRS.EmptyLine && PRS.RangesCount > 0)
    {
        // Найдем верхнюю точку объектов обтекания (т.е. так чтобы при новом обсчете не учитывался только
        // этот объект, заканчивающийся выше всех)

        var Ranges = PRS.Ranges;

        var RangesMaxY = Ranges[0].Y1;
        for (var Index = 1; Index < Ranges.length; Index++)
        {
            if (RangesMaxY > Ranges[Index].Y1)
                RangesMaxY = Ranges[Index].Y1;
        }

        if (Math.abs(RangesMaxY - PRS.Y) < 0.001)
            PRS.Y = RangesMaxY + 1; // смещаемся по 1мм
        else
            PRS.Y = RangesMaxY + AscCommon.TwipsToMM(1) + 0.001; // Добавляем 0.001, чтобы избавиться от погрешности

        // Отмечаем, что данная строка переносится по Y из-за обтекания
        PRS.RangeY = true;

        // Пересчитываем заново данную строку
        PRS.Reset_Ranges();
        PRS.RecalcResult = recalcresult_CurLine;

        return false;
    }

    return true;
};

Paragraph.prototype.private_RecalculateLineEnd         = function(CurLine, CurPage, PRS, ParaPr)
{
    if ( true === PRS.NewPage )
    {
        // Если это последний элемент параграфа, тогда нам не надо переносить текущий параграф
        // на новую страницу. Нам надо выставить границы так, чтобы следующий параграф начинался
        // с новой страницы.
        this.Pages[CurPage].Set_EndLine( CurLine );
        PRS.RecalcResult = recalcresult_NextPage;
        return false;
    }

    if (true !== PRS.End)
    {
        if ( true === PRS.ForceNewPage )
        {
            this.Pages[CurPage].Set_EndLine( CurLine - 1 );

            if ( 0 === CurLine )
                this.Lines[-1] = new CParaLine();

            PRS.RecalcResult = recalcresult_NextPage;
            return false;
        }
    }
    else
    {
        // В последней строке могут быть заполнены не все отрезки обтекания. Удаляем лишние.
        if (PRS.Range < PRS.RangesCount)
            this.Lines[CurLine].Ranges.length = PRS.Range + 1;

        // Проверим висячую строку
        if (true === ParaPr.WidowControl && CurLine === this.Pages[CurPage].StartLine && CurLine >= 1 && false === this.private_CheckSkipKeepLinesAndWidowControl(CurPage))
        {
            // Проверим не встречается ли в предыдущей строке BreakPage, если да, тогда не учитываем WidowControl
            var BreakPagePrevLine = (this.Lines[CurLine - 1].Info & paralineinfo_BreakPage) | 0;

            if (this.Parent instanceof CDocument
                && true === this.Parent.RecalcInfo.Can_RecalcWidowControl()
                && 0 === BreakPagePrevLine
                && (1 === CurPage && null != this.Get_DocumentPrev())
                && this.Lines[CurLine - 1].Ranges.length <= 1)
            {
                var bBreakPageFromStart = false;
                for (var Index = 0, Count = this.Pages[CurPage - 1].Drawings.length; Index < Count; Index++)
                {
                    var Drawing = this.Pages[CurPage - 1].Drawings[Index];
                    var DrawingLine = Drawing.LineNum;

                    if (DrawingLine >= CurLine - 1)
                    {
                        bBreakPageFromStart = true;
                        break;
                    }
                }

                // Если в строках, которые мы переносим есть картинки, либо, если у нас в параграфе 3 строки,
                // тогда сразу начинаем параграф с новой строки
                if (true === bBreakPageFromStart || CurLine <= 2)
                {
                    CurLine = 0;
                    // Вызываем данную функцию для удаления картинок с предыдущей страницы
                    this.Recalculate_Drawing_AddPageBreak(0, 0, true);
                }
                else
                    CurLine = CurLine - 1;

                this.Parent.RecalcInfo.Set_WidowControl(this, CurLine);
                PRS.RecalcResult = recalcresult_PrevPage | recalcresultflags_Column;
                return false;
            }
        }

        // Если у нас нумерация относится к знаку конца параграфа, тогда в такой
        // ситуации не рисуем нумерацию у такого параграфа.
        if (para_End === this.Numbering.Item.Type && this.Lines[CurLine].Info & paralineinfo_BreakPage)
        {
            this.Numbering.Item  = null;
            this.Numbering.Run   = null;
            this.Numbering.Line  = -1;
            this.Numbering.Range = -1;
        }

        this.Pages[CurPage].Set_EndLine( CurLine );
        PRS.RecalcResult = recalcresult_NextElement;
    }

    return true;
};

Paragraph.prototype.private_RecalculateLineAlign       = function(CurLine, CurPage, PRS, ParaPr, Fast)
{
    // Здесь мы пересчитываем ширину пробелов (и в особенных случаях дополнительное
    // расстояние между символами) с учетом прилегания параграфа.
    // 1. Если align = left, тогда внутри каждого промежутка текста выравниваем его
    //    к левой границе промежутка.
    // 2. Если align = right, тогда внутри каждого промежутка текста выравниваем его
    //    к правой границе промежутка.
    // 3. Если align = center, тогда внутри каждого промежутка текста выравниваем его
    //    по центру промежутка.
    // 4. Если align = justify, тогда
    //    4.1 Если внутри промежутка ровно 1 слово.
    //        4.1.1 Если промежуток в строке 1 и слово занимает почти всю строку,
    //              добавляем в слове к каждой букве дополнительное расстояние между
    //              символами, чтобы ширина слова совпала с шириной строки.
    //        4.1.2 Если промежуток первый, тогда слово приставляем к левой границе
    //              промежутка
    //        4.1.3 Если промежуток последний, тогда приставляем слово к правой
    //              границе промежутка
    //        4.1.4 Если промежуток ни первый, ни последний, тогда ставим слово по
    //              середине промежутка
    //    4.2 Если слов больше 1, тогда, исходя из количества пробелов между словами в
    //        промежутке, увеличиваем их на столько, чтобы правая граница последнего
    //        слова совпала с правой границей промежутка
    var PRSW = PRS;
    var PRSC = PRS.getCounterState();
    var PRSA = PRS.getAlignState();
    PRSA.Paragraph    = this;
    PRSA.LastW        = 0;
    PRSA.RecalcFast   = Fast;
    PRSA.RecalcResult = recalcresult_NextElement;
    PRSA.PageY        = this.Pages[CurPage].Bounds.Top;
    PRSA.PageX        = this.Pages[CurPage].Bounds.Left;

    var Line        = this.Lines[CurLine];
    var RangesCount = Line.Ranges.length;

    var isDoNotExpandShiftReturn = this.LogicDocument ? this.LogicDocument.IsDoNotExpandShiftReturn() : false;

    for (var CurRange = 0; CurRange < RangesCount; CurRange++)
    {
        var Range = Line.Ranges[CurRange];

        var StartPos = Range.StartPos;
        var EndPos   = Range.EndPos;

        PRSC.Reset( this, Range );

		PRSC.Range.W      = 0;
		PRSC.Range.WEnd   = 0;
		PRSC.Range.WBreak = 0;
        if ( true === this.Numbering.checkRange(CurRange, CurLine) )
            PRSC.Range.W += this.Numbering.WidthVisible;

        for ( var Pos = StartPos; Pos <= EndPos; Pos++ )
        {
            var Item = this.Content[Pos];
            Item.Recalculate_Range_Width( PRSC, CurLine, CurRange );
        }
		

        var JustifyWord  = 0;
        var JustifySpace = 0;
        var RangeWidth   = Range.XEnd - Range.X;

        var X = 0;
		
		let rtlShift = PRSC.SpaceLen + Range.WBreak + PRSC.Range.WEnd;
		let bRtlAlign = ParaPr.Bidi;
		let jc = ParaPr.Jc;
		
		if(!this.bFromDocument && bRtlAlign)
		{
			if(jc === AscCommon.align_Left)
				jc = AscCommon.align_Right;
			else if(jc === AscCommon.align_Right)
				jc = AscCommon.align_Left;
		}

        // Если данный отрезок содержит только формулу, тогда прилегание данного отрезка определяется формулой
        var ParaMath = this.Check_Range_OnlyMath(CurRange, CurLine);
        if (null !== ParaMath)
        {
            var Math_X      = ( 1 === RangesCount ? this.Pages[CurPage].X      + ParaPr.Ind.Left  : Range.X );
            var Math_XLimit = ( 1 === RangesCount ? this.Pages[CurPage].XLimit - ParaPr.Ind.Right : Range.XEnd );

            X = ParaMath.Get_AlignToLine(CurLine, CurRange, PRS.Page, Math_X, Math_XLimit);
        }
        else
        {
			if (this.Lines[CurLine].Info & paralineinfo_BadLeftTab)
			{
				if (bRtlAlign)
					X = Range.X + RangeWidth - Range.W - rtlShift;
				else
					X = Range.X;
				
				JustifyWord  = 0;
				JustifySpace = 0;
			}
			else
			{
				// RangeWidth - ширина всего пространства в данном отрезке, а Range.W - ширина занимаемого пространства
				switch (jc)
				{
					case AscCommon.align_Left :
					{
						if (bRtlAlign)
						{
							X = Range.X + RangeWidth - Range.W - rtlShift;
							if (this.IsUseXLimit())
								X = Math.max(X, Range.X - rtlShift);
						}
						else
						{
							X = Range.X;
						}
						break;
					}
					case AscCommon.align_Right:
					{
						if (bRtlAlign)
						{
							X = Range.X - rtlShift;
							if (this.IsUseXLimit())
								X = Math.max(X, Range.X - rtlShift);
						}
						else
						{
							X = Range.X + RangeWidth - Range.W;
							if (this.IsUseXLimit())
								X = Math.max(X, Range.X);
						}
						
						break;
					}
					case AscCommon.align_Center:
					{
						if (bRtlAlign)
						{
							X = Range.X + (RangeWidth - Range.W) / 2 - rtlShift;
							if (this.IsUseXLimit())
								X = Math.max(X, Range.X - rtlShift);
						}
						else
						{
							X = Range.X + (RangeWidth - Range.W) / 2;
							if (this.IsUseXLimit())
								X = Math.max(X, Range.X);
						}
						break;
					}
					case AscCommon.align_Justify:
					{
						if (Range.WEnd > AscWord.EPSILON || (Range.WBreak > AscWord.EPSILON && isDoNotExpandShiftReturn))
						{
							if (bRtlAlign)
								X = Range.X + RangeWidth - Range.W - rtlShift;
							else
								X = Range.X;
							
							JustifyWord  = 0;
							JustifySpace = 0;
						}
                        else if (1 === PRSC.Words || PRSC.Spaces <= 0)
                        {
							// Проверяем по количеству пробелов, т.к., например, в китайском языке пробелов нет, но
							// каждый иероглиф как отдельное слово идет.
                            if (1 === RangesCount && !(Line.Info & paralineinfo_End) && !bRtlAlign)
                            {
								X = Range.X;
                                // Либо слово целиком занимает строку, либо не целиком, но разница очень мала
								// либо это набор китайских иероглифов (PRSC.Words > 1)
                                if ((RangeWidth - Range.W <= 0.05 * RangeWidth || PRSC.Words > 1) && PRSC.Letters > 1)
                                    JustifyWord = (RangeWidth - Range.W) / (PRSC.Letters - 1);
                            }
                            else if (0 === CurRange || Line.Info & paralineinfo_End)
                            {
                                // TODO: Здесь нужно улучшить проверку, т.к. отключено выравнивание по центру для всей
                                //       последней строки, а нужно отключить для последнего отрезка, в котором идет
                                //       конец параграфа.
								if (bRtlAlign)
									X = Range.X + RangeWidth - Range.W - rtlShift;
								else
									X = Range.X;
							}
							else if (CurRange === RangesCount - 1)
							{
								if (bRtlAlign)
									X = Range.X - rtlShift;
								else
									X = Range.X + RangeWidth - Range.W;
							}
							else
							{
								if (bRtlAlign)
									X = Range.X + (RangeWidth - Range.W) / 2 - rtlShift;
								else
									X = Range.X + (RangeWidth - Range.W) / 2;
							}
						}
						else
						{
							// TODO: Переделать проверку последнего отрезка в последней строке (нужно выставлять флаг когда пришел PRS.End в отрезке)
							
							// Последний промежуток последней строки не надо растягивать по ширине.
							if (PRSC.Spaces > 0 && (!(Line.Info & paralineinfo_End) || CurRange !== Line.Ranges.length - 1))
							{
								if (bRtlAlign)
									X = Range.X - rtlShift;
								else
									X = Range.X;
								
								JustifySpace = (RangeWidth - Range.W) / PRSC.Spaces;
							}
							else
							{
								if (bRtlAlign)
									X = Range.X + RangeWidth - Range.W - rtlShift;
								else
									X = Range.X;
								
								JustifySpace = 0;
							}
						}
						break;
					}
					default:
					{
						if (bRtlAlign)
							X = Range.X + RangeWidth - Range.W - rtlShift;
						else
							X = Range.X;
						break;
					}
				}
			}
		}
		
        Range.Spaces = PRSC.Spaces + PRSC.SpacesSkip;

        PRSA.X    = X;
        PRSA.Y    = this.Pages[CurPage].Y + this.Lines[CurLine].Y;
        PRSA.XEnd = Range.XEnd;
        PRSA.JustifyWord   = JustifyWord;
        PRSA.JustifySpace  = JustifySpace;
        PRSA.SpacesCounter = PRSC.Spaces;
        PRSA.SpacesSkip    = PRSC.SpacesSkip;
        PRSA.LettersSkip   = PRSC.LettersSkip;
        PRSA.RecalcResult  = recalcresult_NextElement;

        var _LineMetrics = this.Lines[CurLine].Metrics;
        PRSA.Y0 = (this.Pages[CurPage].Y + this.Lines[CurLine].Y - _LineMetrics.Ascent);
        PRSA.Y1 = (this.Pages[CurPage].Y + this.Lines[CurLine].Y + _LineMetrics.Descent);
        if (_LineMetrics.LineGap < 0)
            PRSA.Y1 += _LineMetrics.LineGap;

        this.Lines[CurLine].Ranges[CurRange].XVisible = X;

        if ( 0 === CurRange )
            this.Lines[CurLine].X = X - PRSW.XStart;

        if ( true === this.Numbering.checkRange(CurRange, CurLine) )
            PRSA.X += this.Numbering.WidthVisible;

        for ( var Pos = StartPos; Pos <= EndPos; Pos++ )
        {
            var Item = this.Content[Pos];
            Item.Recalculate_Range_Spaces(PRSA, CurLine, CurRange, CurPage);

            if (!(PRSA.RecalcResult & recalcresult_NextElement))
            {
                PRSW.RecalcResult = PRSA.RecalcResult;
                return PRSA.RecalcResult;
            }
        }
		
		Range.XEndVisible = PRSA.X;
    }

    return PRSA.RecalcResult;
};

Paragraph.prototype.private_RecalculateLineCheckFootnotes = function(CurLine, CurPage, PRS, ParaPr)
{
    if (!((PRS.RecalcResult & recalcresult_NextElement) || (PRS.RecalcResult & recalcresult_NextLine)))
        return false;

	if (PRS.Fast)
		return true;

	var oTopDocument  = PRS.TopDocument;
	var arrFootnotes  = [];
	var oLineBreakPos = this.GetLineEndPos(CurLine);
	for (var nIndex = 0, nCount = PRS.Footnotes.length; nIndex < nCount; ++nIndex)
	{
		var oFootnote = PRS.Footnotes[nIndex].FootnoteReference.GetFootnote();
		var oPos      = PRS.Footnotes[nIndex].Pos;

		// Проверим позицию
		if (oLineBreakPos.Compare(oPos) <= 0)
			continue;

		arrFootnotes.push(oFootnote);
	}

	if (oTopDocument instanceof CDocument)
	{
		if (!oTopDocument.Footnotes.RecalculateFootnotes(PRS.PageAbs, PRS.ColumnAbs, this.Pages[CurPage].Y + this.Lines[CurLine].Bottom, arrFootnotes))
		{
			this.private_RecalculateMoveLineToNextPage(CurLine, CurPage, PRS, ParaPr);
			return false;
		}
	}

	return true;
};

Paragraph.prototype.private_RecalculateLineCheckEndnotes = function(CurLine, CurPage, PRS, ParaPr)
{
    if (!((PRS.RecalcResult & recalcresult_NextElement) || (PRS.RecalcResult & recalcresult_NextLine)) || PRS.Fast)
        return;

    var oTopDocument  = PRS.TopDocument;
    var arrEndnotes   = [];
    var oLineBreakPos = this.GetLineEndPos(CurLine);
    for (var nIndex = 0, nCount = PRS.Endnotes.length; nIndex < nCount; ++nIndex)
    {
        var oEndnote = PRS.Endnotes[nIndex].EndnoteReference.GetFootnote();
        var oPos     = PRS.Endnotes[nIndex].Pos;

        // Проверим позицию
        if (oLineBreakPos.Compare(oPos) <= 0)
            continue;

        arrEndnotes.push(oEndnote);
    }

    if (oTopDocument instanceof CDocument)
        oTopDocument.GetEndnotesController().RegisterEndnotes(PRS.PageAbs, arrEndnotes);
};

Paragraph.prototype.private_RecalculateRange           = function(CurRange, CurLine, CurPage, RangesCount, PRS, paraPr)
{
    // Найдем начальную позицию данного отрезка
    var StartPos = 0;
    if ( 0 === CurLine && 0 === CurRange )
        StartPos = 0;
    else if ( CurRange > 0 )
        StartPos = this.Lines[CurLine].Ranges[CurRange - 1].EndPos;
    else
        StartPos = this.Lines[CurLine - 1].Ranges[ this.Lines[CurLine - 1].Ranges.length - 1 ].EndPos;

    var Line = this.Lines[CurLine];
    var Range = Line.Ranges[CurRange];
	
	this.Lines[CurLine].setRangeStartPos(CurRange, StartPos);
	
	// Correct first line indentation if previous ranges were empty
	if (PRS.UseFirstLine && 0 !== CurRange && PRS.EmptyLine)
	{
		let shift = 0;
		if (PRS.getCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15)
			shift = Math.max(paraPr.Ind.FirstLine, 0);
		else
			shift = paraPr.Ind.FirstLine < -AscWord.EPSILON ? paraPr.Ind.Left + paraPr.Ind.FirstLine : paraPr.Ind.FirstLine;
			
		if (paraPr.Bidi)
			Range.XEnd -= shift;
		else
			Range.X += shift;
	}
	
	PRS.resetRange(Range);

    var ContentLen = this.Content.length;

    var Pos = StartPos;
    for ( ;Pos < ContentLen; Pos++ )
    {
        var Item = this.Content[Pos];

        if ( para_Math === Item.Type )
        {
            var NotInlineMath = this.CheckMathPara(Pos);
            if (true === NotInlineMath && true !== PRS.EmptyLine)
            {
                PRS.ForceNewLine = true;
                PRS.NewRange = true;
                Pos--;
                break;
            }
            // TODO: Надо бы перенести эту проверку на изменение контента параграфа
            Item.Set_Inline(true !== NotInlineMath);
        }

        if ( ( 0 === Pos && 0 === CurLine && 0 === CurRange ) || Pos !== StartPos )
        {
            Item.Recalculate_Reset(CurRange, CurLine);
        }

        PRS.Update_CurPos( Pos, 0 );
        Item.Recalculate_Range( PRS, paraPr, 1 );

        if ( true === PRS.NewRange )
        {
            break;
        }
    }

    if ( Pos >= ContentLen )
        Pos = ContentLen - 1;

    if (PRS.RecalcResult & recalcresult_NextLine)
    {
        // У нас отрезок пересчитался нормально и тут возможны 2 варианта :
        // 1. Отрезок закончился в данной позиции
        // 2. Не все убралось в заданный отрезок и перенос нужно поставить в позиции PRS.LineBreakPos

        if ( true === PRS.MoveToLBP )
        {
            // Отмечаем, что в заданной позиции заканчивается отрезок
            this.private_RecalculateRangeEndPos( PRS, PRS.LineBreakPos, 0 );
        }
        else
			this.Lines[CurLine].setRangeEndPos(CurRange, Pos);
    }
};

Paragraph.prototype.private_RecalculateRangeEndPos     = function(PRS, PRP, Depth)
{
    var CurLine  = PRS.Line;
    var CurRange = PRS.Range;
    var CurPos   = PRP.Get(Depth);

    this.Content[CurPos].Recalculate_Set_RangeEndPos(PRS, PRP, Depth + 1);
	this.Lines[CurLine].setRangeEndPos(CurRange, CurPos);
};

Paragraph.prototype.private_RecalculateGetTabPos = function(PRS, X, ParaPr, CurPage, NumTab)
{
	let contentFrame = this.Parent.Get_PageContentStartPos2(this.PageNum, this.ColumnNum, CurPage, this.Index);
	
	let startX = contentFrame.X;
	let endX   = contentFrame.XLimit;
	
	let paraFrame = this.Get_FramePr();
	if (paraFrame)
	{
		startX = 0;
		endX   = paraFrame.W;
	}
	
	if (PRS.RangesCount > 0 && Math.abs(PRS.Ranges[0].X0 - contentFrame.X) < 0.001)
		startX = PRS.Ranges[0].X1;
	
	if (this.isRtlDirection())
	{
		let pageRel = this.private_GetRelativePageIndex(CurPage);
		let pageLimits = this.Parent.Get_PageLimits(pageRel);
		
		let range = this.Lines[PRS.Line].Ranges[PRS.Range];
		X = X - range.X + pageLimits.XLimit - range.XEnd;
		startX = pageLimits.XLimit - endX;
	}
	
	// Если у данного параграфа есть табы, тогда ищем среди них
	// Добавим в качестве таба левую границу
	let tabs = [];
	let addLefInd = true;
	let paraTabs  = ParaPr.Tabs;
	for (let tabIndex = 0, tabCount = paraTabs.GetCount(); tabIndex < tabCount; ++tabIndex)
	{
		let tab    = paraTabs.Get(tabIndex);
		let tabPos = tab.Pos + startX;
		
		if (addLefInd && tabPos > startX + ParaPr.Ind.Left)
		{
			tabs.push(new CParaTab(tab_Left, ParaPr.Ind.Left));
			addLefInd = false;
		}
		
		if (tab_Clear !== tab.Value)
			tabs.push(tab);
	}
	
	if (addLefInd)
		tabs.push(new CParaTab(tab_Left, ParaPr.Ind.Left));
	
	let customTab = null;
	for (let tabIndex = 0, tabCount = tabs.length; tabIndex < tabCount; ++tabIndex)
	{
		let tab = tabs[tabIndex];
		
		// TODO: Пока здесь сделаем поправку на погрешность. Когда мы сделаем так, чтобы все наши значения хранились
		//       в тех же единицах, что и в формате Docx, тогда и здесь можно будет вернуть обычное сравнение (см. баг 22586)
		//       Разница с NumTab возникла из-за бага 22586, везде нестрогое оставлять нельзя из-за бага 32051.
		
		let twX      = AscCommon.MMToTwips(X);
		let twTabPos = AscCommon.MMToTwips(tab.Pos + startX);
		
		if ((true === NumTab && twX <= twTabPos) || (true !== NumTab && twX < twTabPos))
		{
			customTab = tab;
			break;
		}
	}

    var isTabToRightEdge = false;

    var NewX = 0;

    // Если табов нет, либо их позиции левее текущей позиции ставим таб по умолчанию
    var DefTab = ParaPr.DefaultTab != null ? ParaPr.DefaultTab : AscCommonWord.Default_Tab_Stop;
	if (customTab)
	{
		NewX = customTab.Pos + startX;
	}
    else
    {
        if ( X < startX + ParaPr.Ind.Left )
        {
            NewX = startX + ParaPr.Ind.Left;
        }
        else if (DefTab < 0.001)
        {
            NewX = X;
        }
        else
        {
            NewX = startX;
            while ( X >= NewX - 0.001 )
                NewX += DefTab;
        }

		// Так работает Word: если таб начался в допустимом отрезке, а заканчивается вне его,
		// то мы ограничиваем его правым полем документа, но только если правый отступ параграфа
		// неположителен (<= 0). (смотри bug 32345)
        var twX      = AscCommon.MMToTwips(X);
        var twEndPos = AscCommon.MMToTwips(endX);
        var twNewX   = AscCommon.MMToTwips(NewX);

        if (twX < twEndPos && twNewX >= twEndPos && AscCommon.MMToTwips(ParaPr.Ind.Right) <= 0)
		{
			NewX = endX;
			isTabToRightEdge = true;
		}
    }

	return {
		TabWidth     : NewX - X,
		TabValue     : customTab ? customTab.Value : tab_Left,
		DefaultTab   : !customTab,
		TabLeader    : customTab ? customTab.Leader : Asc.c_oAscTabLeader.None,
		TabRightEdge : isTabToRightEdge,
		PageX        : startX,
		PageXLimit   : endX
	};
};

Paragraph.prototype.private_CheckSkipKeepLinesAndWidowControl = function(CurPage)
{
    var bSkipWidowAndKeepLines = false;
    if (this.ColumnsCount > 1)
    {
        var bWrapDrawing = false;
        for (var TempPage = 0; TempPage <= CurPage; ++TempPage)
        {
            for (var DrawingIndex = 0, DrawingsCount = this.Pages[TempPage].Drawings.length; DrawingIndex < DrawingsCount; ++DrawingIndex)
            {
                if (this.Pages[TempPage].Drawings[DrawingIndex].Use_TextWrap())
                {
                    bWrapDrawing = true;
                    break;
                }
            }

            if (bWrapDrawing)
                break;
        }

        bSkipWidowAndKeepLines = bWrapDrawing;
    }

    return bSkipWidowAndKeepLines;
};

Paragraph.prototype.private_CheckColumnBreak = function(CurPage)
{
    if (this.IsEmptyPage(CurPage))
        return;

    var Page = this.Pages[CurPage];
    var Line = this.Lines[Page.EndLine];

    if (!Line)
        return;

    if (Line.Info & paralineinfo_BreakPage && !(Line.Info & paralineinfo_BreakRealPage))
    {
        if (this.bFromDocument && this.LogicDocument)
            this.LogicDocument.OnColumnBreak_WhileRecalculate();
    }
};

Paragraph.prototype.private_RecalculateMoveLineToNextPage = function(CurLine, CurPage, PRS, ParaPr)
{
	// TODO: Неразрывные абзацы и висячие строки внутри колонок вместе с плавающими объектами пока не обсчитываем
	var bSkipWidowAndKeepLines = this.private_CheckSkipKeepLinesAndWidowControl(CurPage);

	// Проверим висячую строку
	if (this.Parent instanceof CDocument
		&& false === bSkipWidowAndKeepLines
		&& true === this.Parent.RecalcInfo.Can_RecalcWidowControl()
		&& true === ParaPr.WidowControl
		&& CurLine - this.Pages[CurPage].StartLine <= 1
		&& CurLine >= 1 && true != PRS.BreakPageLine
		&& ( 0 === CurPage && null != this.Get_DocumentPrev() ) )
	{
		// Вызываем данную функцию для удаления картинок с предыдущей страницы
		this.Recalculate_Drawing_AddPageBreak(0, 0, true);

		// TODO: Здесь перенос нужно делать сразу же, если в строке не было объектов с обтеканием
		this.Parent.RecalcInfo.Set_WidowControl(this, CurLine - 1);
		PRS.RecalcResult = recalcresult_CurPage | recalcresultflags_Column;
		return false;
	}
	else
	{
		// Учитываем неразрывные абзацы:
		//   1. В Word2010 (версия <= 14) просто проверяем, если параграф разбивается на 2 страницы, тогда
		//      переносим его с новой страницы. Также не учитываем неразрывные параграфы внутри таблиц.
		//   2. В Word2016 (версия >= 15) в добавок к предыдущему ориентируемся на колонки: пытаемся текущую
		//      страницу параграфа разместить в какой либо колонке (пересчитывая их по очереди), если параграф
		//      все равно не рассчитан до конца, тогда размещаем его в первой колонке и делаем перенос на следующую
		//      страницу.
		if (true === ParaPr.KeepLines && false === bSkipWidowAndKeepLines)
		{
			let compatibilityMode = PRS.getCompatibilityMode();
			if (compatibilityMode <= AscCommon.document_compatibility_mode_Word14)
			{
				if (null != this.Get_DocumentPrev() && !this.IsTableCellContent() && 0 === CurPage)
				{
					CurLine = 0;
					PRS.RunRecalcInfoBreak = null;
				}
			}
			else if (compatibilityMode >= AscCommon.document_compatibility_mode_Word15)
			{
				// TODO: Разобраться с 2016 вордом
				if (null != this.Get_DocumentPrev() && 0 === CurPage)
				{
					CurLine = 0;
					PRS.RunRecalcInfoBreak = null;
				}
			}
		}

		// Восстанавливаем позицию нижней границы предыдущей страницы
		this.Pages[CurPage].Bounds.Bottom = PRS.LinePrevBottom;
		this.Pages[CurPage].Set_EndLine( CurLine - 1 );

		if ( 0 === CurLine )
			this.Lines[-1] = new CParaLine(0);

		// Добавляем разрыв страницы
		PRS.RecalcResult = recalcresult_NextPage;
		return false;
	}
};

Paragraph.prototype.private_CheckNeedBeforeSpacing = function(CurPage, Parent, PageAbs, ParaPr)
{
	if (CurPage <= 0)
	{
		let oPrevElement = this.GetPrevDocumentElement();
		while (oPrevElement && !oPrevElement.IsInline())
		{
			oPrevElement = oPrevElement.GetPrevDocumentElement();
		}

		return (!oPrevElement
			|| oPrevElement.GetAbsolutePage(oPrevElement.GetPagesCount() - 1) >= PageAbs
			|| !oPrevElement.IsParagraph()
			|| oPrevElement.Get_SectionPr());
	}

	if (!this.Check_FirstPage(CurPage))
	{
		// Если на предыдущих страницах были только разрывы страниц и колонок, тогда добавляем расстояние
		if (this.Check_FirstPage(CurPage, true))
			return true;
		else
			return false;
	}

	if (this.LogicDocument
		&& this.LogicDocument.GetCompatibilityMode
		&& this.LogicDocument.GetCompatibilityMode() <= AscCommon.document_compatibility_mode_Word14
		&& true === ParaPr.PageBreakBefore)
		return true;
	
	let topDocument = this.GetTopDocumentContent();
	if (!(topDocument instanceof AscWord.Document) || this.IsTableCellContent())
	{
		if (Parent instanceof AscFormat.CDrawingDocContent && 0 !== CurPage)
			return false;
		
		return true;
	}

	// Если сюда дошли, значит мы либо на верхнем уровне, либо в блочном контроле, который лежит на верхнем уровне.
	// Дальше все зависит от того на какой мы странице. Если на первой странице данной секции,
	// тогда добавляем расстояние, а если нет - нет. Но подсчет первой страницы здесь не совпадает с тем, как она
	// считается для нумерации. Если разрыв секции идет на текущей странице, то первой считается сразу данная страница.

	var LogicDocument = topDocument;
	var SectionIndex  = LogicDocument.GetSectionIndexByElementIndex(this.Get_Index());
	var FirstElement  = LogicDocument.GetFirstElementInSection(SectionIndex);

	if (0 !== SectionIndex && (!FirstElement || FirstElement.Get_AbsolutePage(0) === PageAbs))
		return true;

	return false;
};

Paragraph.prototype.ShapeText = function()
{
	if (!this.RecalcInfo.ShapeText)
		return;

	// TODO: Код для теста скорости функции ShapeText
	// let nRecalcId = this.LogicDocument ? this.LogicDocument.GetRecalcId() : -1;
	// if (this.ShapeId === nRecalcId)
	// 	return;
	//
	// this.ShapeId = nRecalcId;

	// TODO: Сейчас мы шейпим текст целиком во всем параграфе. Для ускорения нужно отслеживать позиции, в которых
	//       произошли изменения (далее влево и вправо найти позиции пробела/таба или другого разделителя слова)
	//       и шейпить текст только в заданном промежутке

	AscWord.ParagraphTextShaper.Shape(this);
	this.RecalcInfo.ShapeText = false;
};
Paragraph.prototype.HyphenateText = function()
{
	if (!this.RecalcInfo.HyphenateText || !this.isAutoHyphenation())
		return;
	
	AscWord.TextHyphenator.hyphenate(this);
};

Paragraph.prototype.ShapeTextInRange = function(oStartPos, oEndPos)
{
	AscWord.ParagraphTextShaper.ShapeRange(this, oStartPos, oEndPos, true);
};
Paragraph.prototype.GetLigatureEndPos = function(oStartPos)
{
	let oLigature = this.GetNextRunElement(oStartPos);
	if (!oLigature || !oLigature.IsText())
		return oStartPos;

	if (!oLigature.IsLigature())
	{
		let oResultPos = oStartPos.Copy();
		oResultPos.Update(oStartPos.GetPos(oStartPos.GetDepth()) + 1, oStartPos.GetDepth());
		return oStartPos;
	}

	let oCurrentPos = oStartPos;
	let oSearchPos = new CParagraphSearchPos();
	this.Get_RightPos(oSearchPos, oCurrentPos, false);

	while (oSearchPos.IsFound())
	{
		oCurrentPos = oSearchPos.GetPos().Copy();

		let oNext = this.GetNextRunElement(oCurrentPos);
		let oPrev = this.GetPrevRunElement(oCurrentPos);

		if (!oPrev
			|| !oNext
			|| !oPrev.IsText()
			|| !oNext.IsText()
			|| !oNext.IsLigatureContinue())
			break;

		oSearchPos.Reset();
		this.Get_RightPos(oSearchPos, oCurrentPos, false);

		if (!oSearchPos.IsFound())
			break;
	}

	return oCurrentPos;
};
Paragraph.prototype.CollectRunItemsInRange = function(oStartPos, oEndPos)
{
	let arrPositions = [];
	let arrItems     = [];

	this.CheckRunContent(function(oRun, nStartPos, nEndPos, oCurrentPos)
	{
		for (let nPos = nStartPos; nPos < nEndPos; ++nPos)
		{
			let oParaPos = oCurrentPos.Copy();
			oParaPos.Add(nPos);

			arrItems.push(oRun.GetElement(nPos));
			arrPositions.push(oParaPos);
		}
	}, oStartPos, oEndPos, true);

	return {
		Positions : arrPositions,
		Items     : arrItems
	};
};
Paragraph.prototype.FindLineBreakInLongWord = function(nWidth, oLineStartPos, oCurPos)
{
	// TODO: Когда будут прокидываться типы HB_GLYPH_FLAG_UNSAFE_TO_BREAK, HB_GLYPH_FLAG_UNSAFE_TO_CONCAT
	//       переделать здесь поиск начальной точки для формирования текста

	let oInfo = this.CollectRunItemsInRange(oLineStartPos, oCurPos);

	let arrPositions = oInfo.Positions;
	let arrItems     = oInfo.Items;

	// По логике первая позиция ДОЛЖНА совпадать с oLineStartPos, поэтому
	// мы не отдаем разрыв в первой позиции, чтобы как минимум 1 символ был на строке
	if (arrPositions.length <= 1)
		return oCurPos;

	let oBreakPosition = oCurPos;
	let nLastPos       = arrPositions.length - 1;
	while (nLastPos > 0)
	{
		// TODO: Возможно здесь проверку стоит изменить (или дополнить) на проверку может ли символ находится
		//       в начале строки и может ли предыдущий находится в конце строки
		while (arrItems[nLastPos].IsCombiningMark() && nLastPos > 0)
		{
			nLastPos--;
		}

		if (0 === nLastPos)
			return oBreakPosition;

		oBreakPosition = arrPositions[nLastPos];

		this.ShapeTextInRange(oLineStartPos, arrPositions[nLastPos]);

		let nTempWidth = 0;
		for (let nPos = 0; nPos < nLastPos; ++nPos)
		{
			nTempWidth += arrItems[nPos].GetWidth();
		}

		if (nTempWidth < nWidth)
			return arrPositions[nLastPos];

		nLastPos--;
	}

	this.ShapeTextInRange(oLineStartPos, arrPositions[1]);
	return arrPositions[1];
};
Paragraph.prototype.Recalculate_SetRangeBounds = function(CurLine, CurRange, oStartPos, oEndPos)
{
	let nStartPos = oStartPos.Get(0);
	let nEndPos   = oEndPos.Get(0);

	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		let oItem = this.Content[nPos];
		if (nPos !== nStartPos)
			oItem.Recalculate_Reset(CurRange, CurLine);

		oItem.Recalculate_SetRangeBounds(CurLine, CurRange, nPos === nStartPos ? oStartPos : null, nPos === nEndPos ? oEndPos : null, 1);
	}
};
Paragraph.prototype.GetContentWidthInRange = function(oStartPos, oEndPos)
{
	let nWidth = 0;

	let nStartPos = oStartPos && 0 <= oStartPos.GetDepth()? oStartPos.Get(0) : 0;
	let nEndPos   = oEndPos && 0 <= oEndPos.GetDepth() ? oEndPos.Get(0) : this.Content.length - 1;

	for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		nWidth += this.Content[nPos].GetContentWidthInRange(nPos === nStartPos ? oStartPos : null, nPos === nEndPos ? oEndPos : null, 1);
	}

	return nWidth;
};


var ERecalcPageType =
{
    START   : 0x00, // начать заново пересчет, с начала страницы
    ELEMENT : 0x01, // начать заново пересчет, начиная с заданного элемента
    Y       : 0x02  // начать заново пересчет, начиная с заданной позиции по вертикали
};

function CRecalcPageType()
{
    this.Type    = ERecalcPageType.START;
    this.Element = null;
    this.Y       = -1;
}

CRecalcPageType.prototype.Reset = function()
{
    this.Type    = ERecalcPageType.START;
    this.Element = null;
    this.Y       = -1;
};
CRecalcPageType.prototype.Set_Element = function(Element)
{
    this.Type    = ERecalcPageType.Element;
    this.Element = Element;
};
CRecalcPageType.prototype.Set_Y = function(Y)
{
    this.Type = ERecalcPageType.Y;
    this.Y    = Y;
};

var paralineinfo_BreakPage     = 0x0001; // В строке есть PageBreak или ColumnBreak
var paralineinfo_Empty         = 0x0002; // Строка пустая
var paralineinfo_End           = 0x0004; // Последняя строка параграфа
var paralineinfo_RangeY        = 0x0008; // Строка начинается после какого-либо объекта с обтеканием
var paralineinfo_BreakRealPage = 0x0010; // В строке есть PageBreak
var paralineinfo_BadLeftTab    = 0x0020; // В строке есть левый таб, который правее правой границы
var paralineinfo_Notes         = 0x0040; // В строке есть сноски
var paralineinfo_TextOnLine    = 0x0080; // Есть ли в строке текст
var paralineinfo_BreakLine     = 0x0100; // Строка закончилась переносом строки
var paralineinfo_LongWord      = 0x0200; // В строке длинное слово, которое не убралось
let paralineinfo_AutoHyphen    = 0x0400; // Строка закончилась автопереносом

function CParaLine()
{
    this.Y       = 0; // Позиция BaseLine
    this.Top     = 0;
    this.Bottom  = 0;
    this.Metrics = new CParaLineMetrics();
    this.Ranges  = []; // Массив CParaLineRanges
    this.Info    = 0;  // Побитовая информация о строке:
                       // 1 бит : есть ли PageBreak в строке
                       // 2 бит : пустая ли строка (без учета PageBreak)
                       // 3 бит : последняя ли это строка (т.е. строка с ParaEnd)
                       // 4 бит : строка переносится по Y по обтекаемому объекту
	this.CF      = [];
}
CParaLine.prototype =
{
    Shift : function(Dx, Dy)
    {
        // По Y мы ничего не переносим, т.к. все значени по Y у строки относительно начала страницы данного параграфа
        for (var CurRange = 0, RangesCount = this.Ranges.length; CurRange < RangesCount; CurRange++)
        {
            this.Ranges[CurRange].Shift(Dx, Dy);
        }
    },

    Get_StartPos : function()
    {
        if (this.Ranges.length <= 0)
            return 0;

        return this.Ranges[0].StartPos;
    },

    Get_EndPos : function()
    {
        if (this.Ranges.length <= 0)
            return 0;

        return this.Ranges[this.Ranges.length - 1].EndPos;
    },

    Copy : function()
    {
        var NewLine = new CParaLine();

        NewLine.Y      = this.Y;
        NewLine.Top    = this.Top;
        NewLine.Bottom = this.Bottom;

        NewLine.Metrics.Ascent      = this.Ascent;
        NewLine.Metrics.Descent     = this.Descent;
        NewLine.Metrics.TextAscent  = this.TextAscent;
        NewLine.Metrics.TextAscent2 = this.TextAscent2;
        NewLine.Metrics.TextDescent = this.TextDescent;
        NewLine.Metrics.LineGap     = this.LineGap;

        for (var CurRange = 0, RangesCount = this.Ranges.length; CurRange < RangesCount; CurRange++)
        {
            NewLine.Ranges[CurRange] = this.Ranges[CurRange].Copy();
        }

        NewLine.Info = this.Info;

        return NewLine;
    },

    Reset : function()
    {
        //this.Y        = 0;
        this.Top      = 0;
        this.Bottom   = 0;
        this.Metrics  = new CParaLineMetrics();
        this.Ranges   = [];
        this.Info     = 0;
    }
};
CParaLine.prototype.addRange = function(x, xEnd)
{
	this.Ranges.push(new AscWord.CParaLineRange(x, xEnd));
};
CParaLine.prototype.setRangeStartPos = function(rangeIndex, startPos)
{
	this.Ranges[rangeIndex].StartPos = startPos;
};
CParaLine.prototype.setRangeEndPos = function(rangeIndex, endPos)
{
	this.Ranges[rangeIndex].EndPos = endPos;
};

function CParaLineMetrics()
{
    this.Ascent      = 0; // Высота над BaseLine
    this.Descent     = 0; // Высота после BaseLine
    this.TextAscent  = 0; // Высота текста над BaseLine
    this.TextAscent2 = 0; // Высота текста над BaseLine
    this.TextDescent = 0; // Высота текста после BaseLine
    this.LineGap     = 0; // Дополнительное расстояние между строками
}

CParaLineMetrics.prototype =
{
    Update : function(TextAscent, TextAscent2, TextDescent, Ascent, Descent, ParaPr)
    {
        if ( TextAscent > this.TextAscent )
            this.TextAscent = TextAscent;

        if ( TextAscent2 > this.TextAscent2 )
            this.TextAscent2 = TextAscent2;

        if ( TextDescent > this.TextDescent )
            this.TextDescent = TextDescent;

        if ( Ascent > this.Ascent )
            this.Ascent = Ascent;

        if ( Descent > this.Descent )
            this.Descent = Descent;

        if ( this.Ascent < this.TextAscent )
            this.Ascent = this.TextAscent;

        if ( this.Descent < this.TextDescent )
            this.Descent = this.TextDescent;

        this.LineGap = this.Recalculate_LineGap( ParaPr, this.TextAscent, this.TextDescent );

        if (Asc.linerule_AtLeast === ParaPr.Spacing.LineRule && (this.Ascent + this.Descent + this.LineGap) > (this.TextAscent + this.TextDescent))
        {
            // В такой ситуации Word располагает текст внизу строки
            this.Ascent  = this.Ascent + this.LineGap;
            this.LineGap = 0;
        }
    },

    Recalculate_LineGap : function(ParaPr, TextAscent, TextDescent)
    {
        var LineGap = 0;
        switch ( ParaPr.Spacing.LineRule )
        {
            case Asc.linerule_Auto:
            {
                LineGap = ( TextAscent + TextDescent ) * ( ParaPr.Spacing.Line - 1 );
                break;
            }
            case Asc.linerule_Exact:
            {
                var ExactValue = Math.max( 25.4 / 72, ParaPr.Spacing.Line );
                LineGap = ExactValue - ( TextAscent + TextDescent );

                var Gap = this.Ascent + this.Descent - ExactValue;

                if ( Gap > 0 )
                {
                    var DescentDiff = this.Descent - this.TextDescent;

                    if ( DescentDiff > 0 )
                    {
                        if ( DescentDiff < Gap )
                        {
                            this.Descent = this.TextDescent;
                            Gap -= DescentDiff;
                        }
                        else
                        {
                            this.Descent -= Gap;
                            Gap = 0;
                        }
                    }

                    var AscentDiff = this.Ascent - this.TextAscent;

                    if ( AscentDiff > 0 )
                    {
                        if ( AscentDiff < Gap )
                        {
                            this.Ascent = this.TextAscent;
                            Gap -= AscentDiff;
                        }
                        else
                        {
                            this.Ascent -= Gap;
                            Gap = 0;
                        }
                    }

                    if ( Gap > 0 )
                    {
                        // Уменьшаем пропорционально TextAscent и TextDescent
                        var OldTA = this.TextAscent;
                        var OldTD = this.TextDescent;

                        var Sum = OldTA + OldTD;

                        this.Ascent  = OldTA * (Sum - Gap) / Sum;
                        this.Descent = OldTD * (Sum - Gap) / Sum;
                    }
                }
                else
                {
                    this.Ascent -= Gap; // все в Ascent
                }

                LineGap = 0;


                break;
            }
            case Asc.linerule_AtLeast:
            {
                var TargetLineGap = ParaPr.Spacing.Line;
                var TextLineGap   = TextAscent + TextDescent;
                var RealLineGap   = this.Ascent + this.Descent;

                // Специальный случай, когда в строке нет никакого текста
                if (Math.abs(TextLineGap) < 0.001 || RealLineGap >= TargetLineGap)
                    LineGap = 0;
                else
                    LineGap = TargetLineGap - RealLineGap;

                break;
            }

        }
        return LineGap;
    }
};
CParaLineMetrics.prototype.Copy = function()
{
	var oMetrics = new CParaLineMetrics();

	oMetrics.Ascent      = this.Ascent;
	oMetrics.Descent     = this.Descent;
	oMetrics.TextAscent  = this.TextAscent;
	oMetrics.TextAscent2 = this.TextAscent2;
	oMetrics.TextDescent = this.TextDescent;
	oMetrics.LineGap     = this.LineGap;

	return oMetrics;
};
CParaLineMetrics.prototype.IsEqual = function(oMetrics)
{
	return (Math.abs(oMetrics.Ascent - this.Ascent) < 0.001
		&& Math.abs(oMetrics.Descent - this.Descent) < 0.001
		&& Math.abs(oMetrics.TextAscent - this.TextAscent) < 0.001
		&& Math.abs(oMetrics.TextAscent2 - this.TextAscent2) < 0.001
		&& Math.abs(oMetrics.TextDescent - this.TextDescent) < 0.001
		&& Math.abs(oMetrics.LineGap - this.LineGap) < 0.001);
};
CParaLineMetrics.prototype.Reset = function()
{
	this.Ascent      = 0;
	this.Descent     = 0;
	this.TextAscent  = 0;
	this.TextAscent2 = 0;
	this.TextDescent = 0;
	this.LineGap     = 0;
};

function CParaLineRange(X, XEnd)
{
	this.X           = X;    // Начальная позиция отрезка без учета прилегания содержимого
	this.XVisible    = 0;    // Начальная позиция отрезка с учетом прилегания содержимого
	this.XEnd        = XEnd; // Предельное значение по X для данного отрезка
	this.XEndVisible = X;    // Где фактически заканчивается содержимое в данном отрезке
	this.StartPos    = 0;    // Позиция в контенте параграфа, с которой начинается данный отрезок
	this.EndPos      = 0;    // Позиция в контенте параграфа, на которой заканчиваетсяданный отрезок
	this.W           = 0;
	this.Spaces      = 0;    // Количество пробелов в отрезке, без учета пробелов в конце отрезка
	this.WEnd        = 0;    // Если есть знак конца параграфа в данном отрезке, то это его ширина
	this.WBreak      = 0;    // Если в конце отрезка есть разрыв строки/колонки/страницы
}

CParaLineRange.prototype =
{
    Shift : function(Dx, Dy)
    {
		this.X           += Dx;
		this.XEnd        += Dx;
		this.XVisible    += Dx;
		this.XEndVisible += Dx;
		
		if (this.XEndOrigin)
			this.XEndOrigin += Dx;
    },

    Copy : function()
    {
        var NewRange = new CParaLineRange();

        NewRange.X           = this.X;
        NewRange.XVisible    = this.XVisible;
        NewRange.XEnd        = this.XEnd;
		NewRange.XEndVisible = this.XEndVisible;
        NewRange.StartPos    = this.StartPos;
        NewRange.EndPos      = this.EndPos;
        NewRange.W           = this.W;
        NewRange.Spaces      = this.Spaces;
		
		if (this.XEndOrigin)
			NewRange.XEndOrigin = this.XEndOrigin;

        return NewRange;
    }
};
CParaLineRange.prototype.CorrectX = function(nX)
{
	let x = nX;

	if (x > this.XEnd)
		x = this.XEnd;

	if (x < this.X)
		x = this.X;

	return x;
};
CParaLineRange.prototype.IsZeroRange = function()
{
	return ((this.XEnd - this.X) < 0.001);
};
CParaLineRange.prototype.getXEndOrigin = function()
{
	return (undefined !== this.XEndOrigin ? this.XEndOrigin : this.XEnd);
};
AscWord.CParaLineRange = CParaLineRange;

function CParaPage(X, Y, XLimit, YLimit, FirstLine)
{
    this.X         = X;
    this.Y         = Y;
    this.XLimit    = XLimit;
    this.YLimit    = YLimit;
    this.FirstLine = FirstLine;
    this.Bounds    = new CDocumentBounds( X, Y, XLimit, Y );
    this.StartLine = FirstLine; // Номер строки, с которой начинается данная страница
    this.EndLine   = FirstLine; // Номер последней строки на данной странице
    this.TextPr    = null;      // Расситанные текстовые настройки для начала страницы

    this.Drawings  = [];
    this.EndInfo   = new CParagraphPageEndInfo();
}

CParaPage.prototype =
{
    Reset : function(X, Y, XLimit, YLimit, FirstLine)
    {
        this.X         = X;
        this.Y         = Y;
        this.XLimit    = XLimit;
        this.YLimit    = YLimit;
        this.FirstLine = FirstLine;
        this.Bounds    = new CDocumentBounds( X, Y, XLimit, Y );
        this.StartLine = FirstLine;
        this.Drawings  = [];
    },

    Shift : function(Dx, Dy)
    {
        this.X      += Dx;
        this.Y      += Dy;
        this.XLimit += Dx;
        this.YLimit += Dy;
        this.Bounds.Shift( Dx, Dy );
    },

    Set_EndLine : function(EndLine)
    {
        this.EndLine = EndLine;
    },

    Add_Drawing : function(Item)
    {
        this.Drawings.push(Item);
    },

    Copy : function()
    {
        var NewPage = new CParaPage();

        NewPage.X             = this.X;
        NewPage.Y             = this.Y;
        NewPage.XLimit        = this.XLimit;
        NewPage.YLimit        = this.YLimit;
        NewPage.FirstLine     = this.FirstLine;

        NewPage.Bounds.Left   = this.Bounds.Left;
        NewPage.Bounds.Right  = this.Bounds.Right;
        NewPage.Bounds.Top    = this.Bounds.Top;
        NewPage.Bounds.Bottom = this.Bounds.Bottom;

        NewPage.StartLine     = this.StartLine;
        NewPage.EndLine       = this.EndLine;

        var Count = this.Drawings.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            NewPage.Drawings.push( this.Drawings[Index] );
        }

        NewPage.EndInfo = this.EndInfo.Copy();

        return NewPage;
    }
};

function CParagraphRecalculateTabInfo()
{
    this.TabPos =  0;
    this.X      =  0;
    this.Value  = -1;
    this.Item   = null;
}

CParagraphRecalculateTabInfo.prototype =
{
    Reset : function()
    {
        this.TabPos =  0;
        this.X      =  0;
        this.Value  = -1;
        this.Item   = null;
    }
};

function ParagraphRecalculateStateBase()
{
	this.locked = false;
}
ParagraphRecalculateStateBase.prototype.isLocked = function()
{
	return this.locked;
};
ParagraphRecalculateStateBase.prototype.lock = function()
{
	this.locked = true;
};
ParagraphRecalculateStateBase.prototype.unlock = function()
{
	this.locked = false;
};
window['AscWord'].ParagraphRecalculateStateBase = ParagraphRecalculateStateBase;

function ParagraphStatePool()
{
	this.wrap    = [];
	this.endInfo = [];
	this.draw    = [];
}
ParagraphStatePool.prototype.getInstance = function(pool, className)
{
	let instance = null;
	for (let i = 0, n = pool.length; i < n; ++i)
	{
		if (!pool[i].isLocked())
		{
			instance = pool[i];
			break;
		}
	}
	
	if (!instance)
	{
		instance = new className();
		pool.push(instance);
	}
	
	instance.lock();
	return instance;
};
ParagraphStatePool.prototype.release = function(instance)
{
	instance.unlock();
};
ParagraphStatePool.prototype.getWrapState = function()
{
	return this.getInstance(this.wrap, CParagraphRecalculateStateWrap);
};
ParagraphStatePool.prototype.getEndInfoState = function()
{
	return this.getInstance(this.endInfo, CParagraphRecalculateStateInfo);
};
ParagraphStatePool.prototype.getDrawState = function()
{
	return this.getInstance(this.draw, AscWord.ParagraphDrawState);
};
window['AscWord'].ParagraphStatePool = new ParagraphStatePool();

function CParagraphRecalculateStateWrap()
{
	ParagraphRecalculateStateBase.call(this);
	
    // Общие параметры, которые заполняются 1 раз на пересчет всей страницы
    this.Paragraph       = null;
    this.Parent          = null;
    this.TopDocument     = null;
    this.TopIndex        = -1;   // Номер элемента контейнера (содержащего данный параграф), либо номер данного параграфа в самом верхнем документе
    this.PageAbs         = 0;
    this.ColumnAbs       = 0;
	this.InTable         = false;
    this.SectPr          = null; // настройки секции, к которой относится данный параграф
	this.CondensedSpaces = false;
	this.BalanceSBDB     = false; // BalanceSingleByteDoubleByteWidth
	this.autoHyphenation = false;

	this.Fast            = false; // Быстрый ли пересчет
	
	this.alignState   = new CParagraphRecalculateStateAlign(this);
	this.counterState = new CParagraphRecalculateStateCounter(this);

    //
    this.Page            = 0;
    this.Line            = 0;
    this.Range           = 0;

    this.Ranges          = [];
    this.RangesCount     = 0;
	
	this.LineY = [];

    this.FirstItemOnLine = true;
	this.PrevItemFirst   = false;
    this.EmptyLine       = true;
    this.StartWord       = false;
    this.Word            = false;
    this.AddNumbering    = true;
    this.TextOnLine      = false;
    this.RangeSpaces     = [];

    this.BreakPageLine      = false; // Разрыв страницы (параграфа) в данной строке
    this.UseFirstLine       = false;
    this.BreakPageLineEmpty = false;
    this.BreakRealPageLine  = false; // Разрыв страницы документа (не только параграфа) в данной строке
    this.BadLeftTab         = false; // Левый таб правее правой границы
	this.BreakLine          = false; // Строка закончилась принудительным разрывом
	this.LongWord           = false;

	this.ComplexFields = new AscWord.ParagraphComplexFieldStack();

	this.WordLen         = 0;
    this.SpaceLen        = 0;
    this.SpacesCount     = 0;
    this.LastTab         = new CParagraphRecalculateTabInfo();

    this.LineTextAscent  = 0;
    this.LineTextDescent = 0;
    this.LineTextAscent2 = 0;
    this.LineAscent      = 0;
    this.LineDescent     = 0;

    this.LineTop        = 0;
    this.LineBottom     = 0;
    this.LineTop2       = 0;
    this.LineBottom2    = 0;
    this.LinePrevBottom = 0;

    this.XRange = 0; // Начальное положение по горизонтали для данного отрезка
    this.X      = 0; // Текущее положение по горизонтали
    this.XEnd   = 0; // Предельное значение по горизонтали для текущего отрезка

    this.Y      = 0; // Текущее положение по вертикали

    this.XStart = 0; // Начальное значение для X на данной страницы
    this.YStart = 0; // Начальное значение для Y на данной страницы
    this.XLimit = 0; // Предельное значение для X на данной страницы
    this.YLimit = 0; // Предельное значение для Y на данной страницы

    this.NewPage  = false; // Переходим на новую страницу
    this.NewRange = false; // Переходим к новому отрезку
    this.End      = false;
    this.RangeY   = false; // Текущая строка переносится по Y из-за обтекания

    this.CurPos       = new AscWord.CParagraphContentPos();

    this.NumberingPos = new AscWord.CParagraphContentPos(); // Позиция элемента вместе с которым идет нумерация
	
	this.MoveToLBP      = false; // Делаем ли разрыв в позиции this.LineBreakPos
	this.UpdateLBP      = true;  // Флаг для первичного обновления позиции переноса в отрезке
	this.LineBreakFirst = true;  // Последняя позиция для переноса - это первый элемент в отрезке
	
	// Последняя позиция в которой можно будет добавить разрыв отрезка или строки, если что-то не умещается (например,
	// если у нас не убирается слово, то разрыв ставим перед ним)
	this.LineBreakPos   = new AscWord.CParagraphContentPos();
	
	this.LastItem        = null; // Последний непробельный элемент
	this.LastItemRun     = null; // Run, в котором лежит последний элемент LastItem
	this.LastHyphenItem  = null;
	this.lastAutoHyphen  = null; // Последний элемент с переносом, который убирался в отрезке вместо с дефисом
	this.autoHyphenLimit = 0;
	this.hyphenationZone = 0;

    this.RunRecalcInfoLast  = null; // RecalcInfo последнего рана
    this.RunRecalcInfoBreak = null; // RecalcInfo рана, на котором произошел разрыв отрезка/строки

    this.BaseLineOffset = 0;

    this.RecalcResult = 0x00;//recalcresult_NextElement;
	
	// Управляющий объект для пересчета неинлайновой формулы
	this.MathRecalcInfo = {
		Line : 0,    // Номер строки, с которой начинается формула на текущей странице
		Math : null  // Сам объект формулы
	};
	
    this.Footnotes                  = [];
	this.FootnotesRecalculateObject = null;

	this.Endnotes = [];

    // for ParaMath
    this.bMath_OneLine       = false;
    this.bMathWordLarge      = false;
    this.bEndRunToContent    = false;
    this.PosEndRun           = new AscWord.CParagraphContentPos();

    // параметры, необходимые для расчета разбиения по операторам
    // у "крайних" в строке операторов/мат объектов сооответствующий Gap равен нулю
    this.OperGapRight        = 0;
    this.OperGapLeft         = 0;
    this.bPriorityOper       = true;  // есть ли в контенте операторы с высоким приоритетом разбиения
    this.WrapIndent          = 0;     // WrapIndent нужен для сравнения с длиной слова (когда слово разбивается по Compare Oper): ширина первой строки формулы не должна быть меньше WrapIndent
    this.bContainCompareOper = true;  // содержаться ли в текущем контенте операторы с высоким приоритетом
    this.MathFirstItem       = true;  // параметр необходим для принудительного переноса
    this.bFirstLine          = false;

    this.bNoOneBreakOperator = true;  // прежде чем обновлять позицию в контент Run, учтем были ли до этого break-операторы (проверки на Word == false не достаточно, т.к. формула мб инлайновая и тогда не нужно обновлять позицию)
    this.bForcedBreak        = false;
    this.bInsideOper         = false; // учитываем есть ли разбивка внутри мат объекта, чтобы случайно не вставить в конец пред оператора (при Brk_Before == false)
    this.bOnlyForcedBreak    = false; // учитывается, если возможна разбивка только по операторам выше уровням => в этом случае можно сделать принудительный разрыв во внутреннем контенте
    this.bBreakBox           = false;

    //-----------------------------//
    this.bFastRecalculate    = false;
    this.bBreakPosInLWord    = true; // обновляем LineBreakPos (Set_LineBreakPos) для WordLarge. Не обновляем для инлайновой формулы, перед формулой есть еще текст, чтобы не перебить LineBreakPos и выставить по тем меткам, которые были до формулы разбиение
    this.bContinueRecalc     = false;
    this.bMathRangeY         = false; // используется для переноса формулы под картинку
    this.MathNotInline       = null;
}
CParagraphRecalculateStateWrap.prototype = Object.create(ParagraphRecalculateStateBase.prototype);
CParagraphRecalculateStateWrap.prototype.constructor = CParagraphRecalculateStateWrap;

CParagraphRecalculateStateWrap.prototype.getAlignState = function()
{
	return this.alignState;
};
CParagraphRecalculateStateWrap.prototype.getCounterState = function()
{
	return this.counterState;
};
CParagraphRecalculateStateWrap.prototype.Reset_Page = function(Paragraph, CurPage)
{
	this.Paragraph   = Paragraph;
	this.Parent      = Paragraph.Parent;
	this.TopDocument = Paragraph.Parent.GetTopDocumentContent();
	this.PageAbs     = Paragraph.Get_AbsolutePage(CurPage);
	this.ColumnAbs   = Paragraph.Get_AbsoluteColumn(CurPage);
	this.InTable     = Paragraph.IsTableCellContent();
	this.SectPr      = null;
	this.TopIndex    = -1;
	
	this.CondensedSpaces = Paragraph.IsCondensedSpaces();
	this.BalanceSBDB     = Paragraph.IsBalanceSingleByteDoubleByteWidth();
	
	let settings = this.getDocumentSettings();
	this.autoHyphenation = settings.isAutoHyphenation();
	this.autoHyphenLimit = settings.getConsecutiveHyphenLimit();
	this.hyphenationZone = AscCommon.TwipsToMM(settings.getHyphenationZone());
	
	if (settings.getCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15)
		this.hyphenationZone = AscCommon.TwipsToMM(AscWord.DEFAULT_HYPHENATION_ZONE);
	
	this.Page               = CurPage;
	this.RunRecalcInfoLast  = (0 === CurPage ? null : Paragraph.Pages[CurPage - 1].EndInfo.RunRecalcInfo);
	this.RunRecalcInfoBreak = this.RunRecalcInfoLast;
	
	this.ComplexFields.resetPage(Paragraph, CurPage);
	this.alignState.ComplexFields.resetPage(Paragraph, CurPage);
	this.counterState.ComplexFields.resetPage(Paragraph, CurPage);
};
CParagraphRecalculateStateWrap.prototype.Reset_Line = function()
{
	this.RecalcResult = recalcresult_NextLine;
	
	this.EmptyLine         = true;
	this.BreakPageLine     = false;
	this.BreakLine         = false;
	this.LongWord          = false;
	this.End               = false;
	this.UseFirstLine      = false;
	this.BreakRealPageLine = false;
	this.BadLeftTab        = false
	this.TextOnLine        = false;
	
	this.LineTextAscent  = 0;
	this.LineTextAscent2 = 0;
	this.LineTextDescent = 0;
	this.LineAscent      = 0;
	this.LineDescent     = 0;
	
	this.NewPage      = false;
	this.ForceNewPage = false;
	this.ForceNewLine = false;
	
	this.bMath_OneLine    = false;
	this.bMathWordLarge   = false;
	this.bEndRunToContent = false;
	this.PosEndRun        = new AscWord.CParagraphContentPos();
	this.Footnotes        = [];
	this.Endnotes         = [];
	
	this.OperGapRight        = 0;
	this.OperGapLeft         = 0;
	this.WrapIndent          = 0;
	this.MathFirstItem       = true;
	this.bContainCompareOper = true;
	this.bInsideOper         = false;
	this.bOnlyForcedBreak    = false;
	this.bBreakBox           = false;
	this.bNoOneBreakOperator = true;
	this.bFastRecalculate    = false;
	this.bForcedBreak        = false;
	this.bBreakPosInLWord    = true;
	
	this.MathNotInline = null;
	
	this.LineY.length = this.Line + 1;
	if (this.Line >= 0)
		this.LineY[this.Line] = this.Y;
};
CParagraphRecalculateStateWrap.prototype.resetRange = function(range)
{
	this.LastTab.Reset();
	
	this.BreakLine       = false;
	this.SpaceLen        = 0;
	this.WordLen         = 0;
	this.SpacesCount     = 0;
	this.Word            = false;
	this.FirstItemOnLine = true;
	this.StartWord       = false;
	this.NewRange        = false;
	this.X               = range.X;
	this.XEnd            = range.XEnd;
	this.XRange          = range.X;
	this.RangeSpaces     = [];
	
	this.MoveToLBP      = false;
	this.LineBreakPos   = new AscWord.CParagraphContentPos();
	this.LineBreakFirst = true;
	this.LastItem       = null;
	this.LastItemRun    = null;
	this.UpdateLBP      = true;
	this.LastHyphenItem = null;
	this.lastAutoHyphen = null;
	
	// for ParaMath
	this.bMath_OneLine    = false;
	this.bMathWordLarge   = false;
	this.bEndRunToContent = false;
	this.PosEndRun        = new AscWord.CParagraphContentPos();
	
	this.OperGapRight        = 0;
	this.OperGapLeft         = 0;
	this.WrapIndent          = 0;
	this.bContainCompareOper = true;
	this.bInsideOper         = false;
	this.bOnlyForcedBreak    = false;
	this.bBreakBox           = false;
	this.bNoOneBreakOperator = true;
	this.bForcedBreak        = false;
	this.bFastRecalculate    = false;
	this.bBreakPosInLWord    = true;
};
CParagraphRecalculateStateWrap.prototype.Set_LineBreakPos = function(PosObj, isFirstItemOnLine)
{
	this.LineBreakPos.Set(this.CurPos);
	this.LineBreakPos.Add(PosObj);
	this.LineBreakFirst = isFirstItemOnLine;
	this.ResetLastAutoHyphen();
};
CParagraphRecalculateStateWrap.prototype.getDocumentSettings = function()
{
	let logicDocument = this.Paragraph.GetLogicDocument();
	if (logicDocument && logicDocument.IsDocumentEditor())
		return logicDocument.getDocumentSettings();
	
	return AscWord.DEFAULT_DOCUMENT_SETTINGS;
};
CParagraphRecalculateStateWrap.prototype.getCompatibilityMode = function()
{
	return this.getDocumentSettings().getCompatibilityMode();
};
CParagraphRecalculateStateWrap.prototype.getXLimit = function()
{
	// TODO: Когда перенесем весь расчет в данный класс (из Run.Recalculate_Range), то
	//       при изменении XEnd сразу расчитывать это значение и заменить вызов на простой this.XEnd
	return this.Paragraph.IsUseXLimit() ? this.XEnd : MEASUREMENT_MAX_MM_VALUE * 10;
};
CParagraphRecalculateStateWrap.prototype.ResetLastAutoHyphen = function()
{
	if (!this.LastHyphenItem)
		return;
	
	this.LastHyphenItem.SetTemporaryHyphenAfter(false);
	this.LastHyphenItem = null;
};
CParagraphRecalculateStateWrap.prototype.checkLastAutoHyphen = function()
{
	if (!this.isAutoHyphenation())
		return;
	
	this.ResetLastAutoHyphen();
	let lastItem = this.LastItem;
	if (!lastItem || lastItem !== this.lastAutoHyphen)
		return;
	
	if (this.isExceedConsecutiveAutoHyphenLimit())
		return;
	
	this.LastHyphenItem = lastItem;
	lastItem.SetTemporaryHyphenAfter(true);
};
CParagraphRecalculateStateWrap.prototype.Set_NumberingPos = function(PosObj, Item)
{
	this.NumberingPos.Set(this.CurPos);
	this.NumberingPos.Add(PosObj);
	
	this.Paragraph.Numbering.Pos  = this.NumberingPos;
	this.Paragraph.Numbering.Item = Item;
};
CParagraphRecalculateStateWrap.prototype.Update_CurPos = function(PosObj, Depth)
{
	this.CurPos.Update(PosObj, Depth);
};
CParagraphRecalculateStateWrap.prototype.Reset_Ranges = function()
{
	this.Ranges      = [];
	this.RangesCount = 0;
};
CParagraphRecalculateStateWrap.prototype.Reset_RunRecalcInfo = function()
{
	this.RunRecalcInfoBreak = this.RunRecalcInfoLast;
};
CParagraphRecalculateStateWrap.prototype.Reset_MathRecalcInfo = function()
{
	this.bContinueRecalc = false;
};
CParagraphRecalculateStateWrap.prototype.Restore_RunRecalcInfo = function()
{
	this.RunRecalcInfoLast = this.RunRecalcInfoBreak;
};
CParagraphRecalculateStateWrap.prototype.Recalculate_Numbering = function(Item, Run, ParaPr, _X)
{
	var CurPage = this.Page, CurLine = this.Line, CurRange = this.Range;
	var Para    = this.Paragraph;
	var X       = _X, LineAscent = this.LineAscent;
	
	// Если нужно добавить нумерацию и на текущем элементе ее можно добавить, тогда добавляем её
	var NumberingItem = Para.Numbering;
	var NumberingType = Para.Numbering.Type;
	
	if (para_Numbering === NumberingType)
	{
		var oReviewInfo = this.Paragraph.GetReviewInfo();
		var nReviewType = this.Paragraph.GetReviewType();
		
		var isHavePrChange = this.Paragraph.HavePrChange();
		var oPrevNumPr     = this.Paragraph.GetPrChangeNumPr();
		
		var NumPr = ParaPr.NumPr;
		
		if (!NumPr || !NumPr.IsValid())
			NumPr = undefined;
		
		if (!oPrevNumPr || !oPrevNumPr.IsValid())
		{
			oPrevNumPr = undefined;
		}
		else
		{
			oPrevNumPr = oPrevNumPr.Copy();
			if (undefined === oPrevNumPr.Lvl)
				oPrevNumPr.Lvl = 0;
		}
		
		var isHaveNumbering = false;
		if ((undefined === Para.Get_SectionPr() || true !== Para.IsEmpty()) && (NumPr || oPrevNumPr))
		{
			isHaveNumbering = true;
		}
		
		if (!isHaveNumbering || (!NumPr && !oPrevNumPr) || (!NumPr && reviewtype_Add === nReviewType))
		{
			// Так мы обнуляем все рассчитанные ширины данного элемента
			NumberingItem.Measure(g_oTextMeasurer, undefined);
		}
		else
		{
			var oSavedNumberingValues = this.Paragraph.GetSavedNumberingValues();
			var arrSavedNumInfo       = oSavedNumberingValues ? oSavedNumberingValues.NumInfo : null;
			var arrSavedPrevNumInfo   = oSavedNumberingValues ? oSavedNumberingValues.PrevNumInfo : null;
			
			var oNumbering = Para.Parent.GetNumbering();
			
			var oNumLvl  = null;
			let nNumSuff = Asc.c_oAscNumberingSuff.None;
			
			if (NumPr)
			{
				oNumLvl  = oNumbering.GetNum(NumPr.NumId).GetLvl(NumPr.Lvl);
				nNumSuff = oNumLvl.GetSuff();
			}
			else if (oPrevNumPr)
			{
				// MSWord uses tab instead of suff from PrevNum (74525)
				oNumLvl  = oNumbering.GetNum(oPrevNumPr.NumId).GetLvl(oPrevNumPr.Lvl);
				nNumSuff = Asc.c_oAscNumberingSuff.Tab;
			}
			
			var oNumTextPr = Para.GetNumberingTextPr();
			var nNumJc     = oNumLvl.GetJc();
			
			// Здесь измеряется только ширина символов нумерации, без суффикса
			if ((!isHavePrChange && NumPr) || (oPrevNumPr && NumPr && oPrevNumPr.NumId === NumPr.NumId && oPrevNumPr.Lvl === NumPr.Lvl))
			{
				var arrNumInfo = arrSavedNumInfo ? arrSavedNumInfo : Para.Parent.CalculateNumberingValues(Para, NumPr, true);
				var nLvl       = NumPr.Lvl;
				
				var arrRelatedLvls = oNumLvl.GetRelatedLvlList();
				var isEqual        = true;
				for (var nLvlIndex = 0, nLvlsCount = arrRelatedLvls.length; nLvlIndex < nLvlsCount; ++nLvlIndex)
				{
					var nTempLvl = arrRelatedLvls[nLvlIndex];
					if (arrNumInfo[0][nTempLvl] !== arrNumInfo[1][nTempLvl])
					{
						isEqual = false;
						break;
					}
				}
				
				if (!isEqual)
				{
					if (reviewtype_Common === nReviewType)
					{
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, arrNumInfo[1], NumPr);
					}
					else
					{
						if (reviewtype_Remove === nReviewType && oReviewInfo.GetPrevAdded())
						{
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, undefined, undefined);
						}
						else if (reviewtype_Remove === nReviewType)
						{
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo[1], NumPr);
						}
						else
						{
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, undefined, undefined);
						}
					}
				}
				else
				{
					if (reviewtype_Remove === nReviewType)
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo[1], NumPr);
					else
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr);
				}
			}
			else if (oPrevNumPr && !NumPr)
			{
				var arrNumInfo2 = arrSavedPrevNumInfo ? arrSavedPrevNumInfo : Para.Parent.CalculateNumberingValues(Para, oPrevNumPr, true);
				NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo2[1], oPrevNumPr);
			}
			else if (isHavePrChange && !oPrevNumPr && NumPr)
			{
				if (reviewtype_Remove === nReviewType)
				{
					NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, undefined, undefined);
				}
				else
				{
					var arrNumInfo = arrSavedNumInfo ? arrSavedNumInfo : Para.Parent.CalculateNumberingValues(Para, NumPr, true);
					NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, undefined, undefined);
				}
			}
			else if (oPrevNumPr && NumPr)
			{
				var arrNumInfo  = arrSavedNumInfo ? arrSavedNumInfo : Para.Parent.CalculateNumberingValues(Para, NumPr, true);
				var arrNumInfo2 = arrSavedPrevNumInfo ? arrSavedPrevNumInfo : Para.Parent.CalculateNumberingValues(Para, oPrevNumPr, true);
				
				var isEqual = false;
				if (arrNumInfo[0][NumPr.Lvl] === arrNumInfo[1][oPrevNumPr.Lvl])
				{
					var oSourceNumLvl = oNumbering.GetNum(oPrevNumPr.NumId).GetLvl(oPrevNumPr.Lvl);
					var oFinalNumLvl  = oNumbering.GetNum(NumPr.NumId).GetLvl(NumPr.Lvl);
					
					isEqual = oSourceNumLvl.IsSimilar(oFinalNumLvl);
					if (isEqual)
					{
						var arrRelatedLvls = oSourceNumLvl.GetRelatedLvlList();
						for (var nLvlIndex = 0, nLvlsCount = arrRelatedLvls.length; nLvlIndex < nLvlsCount; ++nLvlIndex)
						{
							var nTempLvl = arrRelatedLvls[nLvlIndex];
							if (arrNumInfo[0][nTempLvl] !== arrNumInfo[1][nTempLvl])
							{
								isEqual = false;
								break;
							}
						}
					}
				}
				
				if (isEqual)
				{
					NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr);
				}
				else
				{
					if (reviewtype_Remove === nReviewType)
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo2[1], oPrevNumPr);
					else if (reviewtype_Add === nReviewType)
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, undefined, undefined);
					else
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, arrNumInfo2[1], oPrevNumPr);
				}
			}
			else
			{
				// Такого быть не должно
			}
			
			// При рассчете высоты строки, если у нас параграф со списком, то размер символа
			// в списке влияет только на высоту строки над Baseline, но не влияет на высоту строки
			// ниже baseline.
			if (LineAscent < NumberingItem.Height)
				LineAscent = NumberingItem.Height;
			
			switch (nNumJc)
			{
				case AscCommon.align_Right:
				{
					NumberingItem.WidthVisible = 0;
					break;
				}
				case AscCommon.align_Center:
				{
					NumberingItem.WidthVisible = NumberingItem.WidthNum / 2;
					break;
				}
				case AscCommon.align_Left:
				default:
				{
					NumberingItem.WidthVisible = NumberingItem.WidthNum;
					break;
				}
			}
			
			X += NumberingItem.WidthVisible;
			
			if (oNumLvl.IsLegacy())
			{
				var nLegacySpace  = AscCommon.TwipsToMM(oNumLvl.GetLegacySpace());
				var nLegacyIndent = AscCommon.TwipsToMM(oNumLvl.GetLegacyIndent());
				var nNumWidth     = NumberingItem.WidthNum;
				
				NumberingItem.WidthSuff = Math.max(nNumWidth, nLegacyIndent, nNumWidth + nLegacySpace) - nNumWidth;
			}
			else
			{
				switch (nNumSuff)
				{
					case Asc.c_oAscNumberingSuff.None:
					{
						// Ничего не делаем
						break;
					}
					case Asc.c_oAscNumberingSuff.Space:
					{
						var OldTextPr = g_oTextMeasurer.GetTextPr();
						
						
						var Theme = Para.Get_Theme();
						g_oTextMeasurer.SetTextPr(oNumTextPr, Theme);
						g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);
						NumberingItem.WidthSuff = g_oTextMeasurer.Measure(" ").Width;
						g_oTextMeasurer.SetTextPr(OldTextPr, Theme);
						break;
					}
					case Asc.c_oAscNumberingSuff.Tab:
					{
						NumberingItem.WidthSuff = Para.private_RecalculateGetTabPos(this, X, ParaPr, CurPage, true).TabWidth;
						break;
					}
				}
			}
			
			NumberingItem.Width = NumberingItem.WidthNum;
			NumberingItem.WidthVisible += NumberingItem.WidthSuff;
			
			X += NumberingItem.WidthSuff;
		}
	}
	else if (para_PresentationNumbering === NumberingType)
	{
		var Level  = Para.PresentationPr.Level;
		var Bullet = Para.PresentationPr.Bullet;
		
		var BulletNum = Para.GetBulletNum();
		if (BulletNum === null)
		{
			BulletNum = 1;
		}
		// Найдем настройки для первого текстового элемента
		var FirstTextPr = Para.Get_FirstTextPr2();
		
		
		if (Bullet.IsAlpha())
		{
			if (BulletNum > 780)
			{
				BulletNum = (BulletNum % 780);
			}
		}
		if (BulletNum > 32767)
		{
			BulletNum = (BulletNum % 32767);
		}
		
		
		NumberingItem.Bullet    = Bullet;
		NumberingItem.BulletNum = BulletNum;
		NumberingItem.Measure(g_oTextMeasurer, FirstTextPr, Para.Get_Theme(), Para.Get_ColorMap());
		
		
		if (!Bullet.IsNone())
		{
			if (ParaPr.Ind.FirstLine < 0)
				NumberingItem.WidthVisible = Math.max(NumberingItem.Width, Para.Pages[CurPage].X + ParaPr.Ind.Left + ParaPr.Ind.FirstLine - X, Para.Pages[CurPage].X + ParaPr.Ind.Left - X);
			else
				NumberingItem.WidthVisible = Math.max(Para.Pages[CurPage].X + ParaPr.Ind.Left + NumberingItem.Width - X, Para.Pages[CurPage].X + ParaPr.Ind.Left + ParaPr.Ind.FirstLine - X, Para.Pages[CurPage].X + ParaPr.Ind.Left - X);
		}
		
		X += NumberingItem.WidthVisible;
	}
	
	// Заполним обратные данные в элементе нумерации
	NumberingItem.Item       = Item;
	NumberingItem.Run        = Run;
	NumberingItem.Line       = CurLine;
	NumberingItem.Range      = CurRange;
	NumberingItem.LineAscent = LineAscent;
	NumberingItem.Page       = CurPage;
	
	return X;
};
CParagraphRecalculateStateWrap.prototype.IsFast = function()
{
	return this.Fast;
};
CParagraphRecalculateStateWrap.prototype.AddFootnoteReference = function(oFootnoteReference, oPos)
{
	// Ссылки могут добавляться несколько раз, если строка разбита на несколько отрезков
	for (var nIndex = 0, nCount = this.Footnotes.length; nIndex < nCount; ++nIndex)
	{
		if (this.Footnotes[nIndex].FootnoteReference === oFootnoteReference)
			return;
	}

	this.Footnotes.push({FootnoteReference : oFootnoteReference, Pos : oPos});
};
CParagraphRecalculateStateWrap.prototype.GetFootnoteReferencesCount = function(oFootnoteReference, isAllowCustom)
{
	var _isAllowCustom = (true === isAllowCustom ? true : false);

	// Если данную ссылку мы добавляли уже в строке, тогда ищем сколько было элементов до нее, если не добавляли,
	// тогда возвращаем просто количество ссылок. Ссылки с флагом CustomMarkFollows не учитываются

	var nRefsCount = 0;
	for (var nIndex = 0, nCount = this.Footnotes.length; nIndex < nCount; ++nIndex)
	{
		if (this.Footnotes[nIndex].FootnoteReference === oFootnoteReference)
			return nRefsCount;

		if (true === _isAllowCustom || true !== this.Footnotes[nIndex].FootnoteReference.IsCustomMarkFollows())
			nRefsCount++;
	}

	return nRefsCount;
};
CParagraphRecalculateStateWrap.prototype.AddEndnoteReference = function(oEndnoteReference, oPos)
{
	for (var nIndex = 0, nCount = this.Endnotes.length; nIndex < nCount; ++nIndex)
	{
		if (this.Endnotes[nIndex].EndnoteReference === oEndnoteReference)
			return;
	}

	this.Endnotes.push({EndnoteReference : oEndnoteReference, Pos : oPos});
};
CParagraphRecalculateStateWrap.prototype.GetEndnoteReferenceNumber = function(oEndnoteReference)
{
	if (this.Endnotes.length <= 0 || this.Endnotes[0].EndnoteReference === oEndnoteReference)
		return -1;

	var nRefsCount = 0;
	for (var nIndex = 0, nCount = this.Endnotes.length; nIndex < nCount; ++nIndex)
	{
		if (this.Endnotes[nIndex].EndnoteReference === oEndnoteReference)
			return (this.Endnotes[0].EndnoteReference.Number + nRefsCount);

		if (true !== this.Endnotes[nIndex].EndnoteReference.IsCustomMarkFollows())
			nRefsCount++;
	}

	return (this.Endnotes[0].EndnoteReference.Number + nRefsCount);
};
CParagraphRecalculateStateWrap.prototype.GetEndnoteReferenceCount = function()
{
	return this.Endnotes.length;
};
CParagraphRecalculateStateWrap.prototype.SetFast = function(bValue)
{
	this.Fast = bValue;
};
CParagraphRecalculateStateWrap.prototype.IsFastRecalculate = function()
{
	return this.Fast;
};
CParagraphRecalculateStateWrap.prototype.isFastRecalculation = function()
{
	return this.Fast;
};
CParagraphRecalculateStateWrap.prototype.GetPageAbs = function()
{
	return this.PageAbs;
};
CParagraphRecalculateStateWrap.prototype.GetColumnAbs = function()
{
	return this.ColumnAbs;
};
CParagraphRecalculateStateWrap.prototype.GetCurrentContentPos = function(nPos)
{
	var oContentPos = this.CurPos.Copy();
	oContentPos.Set(this.CurPos);
	oContentPos.Add(nPos);
	return oContentPos;
};
CParagraphRecalculateStateWrap.prototype.SaveFootnotesInfo = function()
{
	var oTopDocument = this.TopDocument;
	if (oTopDocument instanceof CDocument)
		this.FootnotesRecalculateObject = oTopDocument.Footnotes.SaveRecalculateObject(this.PageAbs, this.ColumnAbs);
};
CParagraphRecalculateStateWrap.prototype.LoadFootnotesInfo = function()
{
	var oTopDocument = this.TopDocument;
	if (oTopDocument instanceof CDocument && this.FootnotesRecalculateObject)
		oTopDocument.Footnotes.LoadRecalculateObject(this.PageAbs, this.ColumnAbs, this.FootnotesRecalculateObject);
};
CParagraphRecalculateStateWrap.prototype.IsInTable = function()
{
	return this.InTable;
};
CParagraphRecalculateStateWrap.prototype.GetSectPr = function()
{
	if (null === this.SectPr && this.Paragraph)
		this.SectPr = this.Paragraph.Get_SectPr();

	return this.SectPr;
};
CParagraphRecalculateStateWrap.prototype.GetTopDocument = function()
{
	return this.TopDocument;
};
CParagraphRecalculateStateWrap.prototype.GetTopIndex = function()
{
	if (-1 === this.TopIndex)
	{
		var arrPos = this.Paragraph.GetDocumentPositionFromObject();
		if (arrPos.length > 0)
			this.TopIndex = arrPos[0].Position;
	}

	return this.TopIndex;
};
CParagraphRecalculateStateWrap.prototype.ResetMathRecalcInfo = function()
{
	this.MathRecalcInfo.Line = 0;
	this.MathRecalcInfo.Math = null;
};
CParagraphRecalculateStateWrap.prototype.SetMathRecalcInfo = function(math)
{
	this.MathRecalcInfo.Line = this.Line;
	this.MathRecalcInfo.Math = math;
};
CParagraphRecalculateStateWrap.prototype.resetToMathFirstLine = function()
{
	this.Line = this.MathRecalcInfo.Line;
	this.Y    = this.LineY[this.Line];
	return this.Line;
};
CParagraphRecalculateStateWrap.prototype.GetMathRecalcInfoObject = function()
{
	return this.MathRecalcInfo.Math;
};
CParagraphRecalculateStateWrap.prototype.SetMathRecalcInfoObject = function(oMath)
{
	this.MathRecalcInfo.Math = oMath;
};
CParagraphRecalculateStateWrap.prototype.IsCondensedSpaces = function()
{
	return this.CondensedSpaces;
};
CParagraphRecalculateStateWrap.prototype.IsBalanceSingleByteDoubleByteWidth = function(oRun, nPos)
{
	if (this.BalanceSBDB)
	{
		let oParaPos = this.Paragraph.GetPosByElement(oRun);
		if (!oParaPos)
			return true;

		oParaPos.Add(nPos);

		let oRunElements = new CParagraphRunElements(oParaPos, 1, null);
		this.Paragraph.GetPrevRunElements(oRunElements);
		let arrElements = oRunElements.GetElements();
		if (arrElements.length <= 0)
			return true;

		let oItem = arrElements[0];
		if (!oItem || para_Text !== oItem.Type || AscCommon.isEastAsianScript(oItem.Value))
			return true;

		oParaPos.Update(nPos + 1, oParaPos.GetDepth());

		oRunElements = new CParagraphRunElements(oParaPos, 1, null);
		this.Paragraph.GetNextRunElements(oRunElements);
		arrElements = oRunElements.GetElements();
		if (arrElements.length <= 0)
			return true;

		oItem = arrElements[0];
		return (!oItem || para_Text !== oItem.Type || AscCommon.isEastAsianScript(oItem.Value));
	}

	return false;
};
CParagraphRecalculateStateWrap.prototype.AddCondensedSpaceToRange = function(oSpace)
{
	this.RangeSpaces.push(oSpace);
	oSpace.ResetCondensedWidth();
};
/**
 * Проверяем убирается ли в заданном отрезке заданная ширина содержимого
 * @param x {number} - текущая позиция
 * @param width {number} - ширина проверяемого промежутка
 * @returns {boolean}
 */
CParagraphRecalculateStateWrap.prototype.isFitOnLine = function(x, width)
{
	let xLimit = this.getXLimit();
	if (x + width <= xLimit)
		return true;
	
	return this.tryCondenseSpaces(x, xLimit, width);
};
/**
 * Пытаемся ужать пробелы по
 * @param x {number} - текущая позиция
 * @param xLimit {number} - предельная позиция
 * @param width {number} - ширина проверяемого промежутка
 * @returns {boolean}
 */
CParagraphRecalculateStateWrap.prototype.tryCondenseSpaces = function(x, xLimit, width)
{
	if (!this.CondensedSpaces)
		return false;
	
	var nKoef = 1 - 0.25 * (Math.min(12.5, width) / 12.5);

	var nSumSpaces = 0;
	for (var nIndex = 0, nCount = this.RangeSpaces.length; nIndex < nCount; ++nIndex)
	{
		nSumSpaces += this.RangeSpaces[nIndex].WidthOrigin / AscWord.TEXTWIDTH_DIVIDER;
	}

	var nSpace = nSumSpaces * (1 - nKoef);
	if (x - nSpace + width < xLimit)
	{
		for (var nIndex = 0, nCount = this.RangeSpaces.length; nIndex < nCount; ++nIndex)
		{
			this.RangeSpaces[nIndex].SetCondensedWidth(nKoef);
		}

		return true;
	}
	else
	{
		for (var nIndex = 0, nCount = this.RangeSpaces.length; nIndex < nCount; ++nIndex)
		{
			this.RangeSpaces[nIndex].ResetCondensedWidth();
		}
	}

	return false;
};
CParagraphRecalculateStateWrap.prototype.CheckUpdateLBP = function(nInRunPos)
{
	 if (this.UpdateLBP)
	 {
		 this.UpdateLBP = false;
		 this.LineBreakPos.Set(this.CurPos);
		 this.LineBreakPos.Add(nInRunPos);
	 }
};
CParagraphRecalculateStateWrap.prototype.IsNeedShapeFirstWord = function(nCurLine)
{
	let arrLines = this.Paragraph.Lines;

	return (0 !== nCurLine
		&& arrLines.length > nCurLine
		&& arrLines[nCurLine - 1].Info & paralineinfo_LongWord);
};
CParagraphRecalculateStateWrap.prototype.IsLastElementInWord = function(oRun, nPos)
{
	let oItem = oRun.GetElement(nPos)
	if (!oItem)
		return false;

	if (oItem.IsSpaceAfter())
		return true;

	let oParent      = oRun.GetParent();
	let nInParentPos = oRun.GetPosInParent(oParent);
	if (!oParent || -1 === nInParentPos)
		return false;

	let oNextItem  = oRun.GetElement(nPos + 1);
	let nParentLen = oParent.GetElementsCount();
	while (!oNextItem && nInParentPos < nParentLen - 1)
	{
		oRun = oParent.GetElement(++nInParentPos);
		if (!oRun || !(oRun instanceof ParaRun))
			return true;

		oNextItem = oRun.GetElement(0);
	}

	return (!oNextItem || !oNextItem.IsText());
};
CParagraphRecalculateStateWrap.prototype.isAutoHyphenation = function()
{
	return this.autoHyphenation;
};
CParagraphRecalculateStateWrap.prototype.getAutoHyphenLimit = function()
{
	return this.autoHyphenLimit;
};
CParagraphRecalculateStateWrap.prototype.getHyphenationZone = function()
{
	return this.hyphenationZone;
};
CParagraphRecalculateStateWrap.prototype.onEndRecalculateLineRange = function()
{
	// Сюда заходим, если закончили пересчиытывать отрезок насильно, а не из-за того, что какой-то элемент не убрался
	// (перенос строк или конец параграфа)
	this.ResetLastAutoHyphen();
};
/**
 * Получам ширину дефиса, если на данном элементе можно разбить слово
 * @returns {number}
 */
CParagraphRecalculateStateWrap.prototype.getAutoHyphenWidth = function(item, run)
{
	if (!this.isAutoHyphenation() || !item || !item.IsText() || !item.isHyphenAfter())
		return 0;
	
	let textPr = run.Get_CompiledPr(false);
	let fontInfo = textPr.GetFontInfo(AscWord.fontslot_ASCII);
	return AscFonts.GetGraphemeWidth(AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(0x002D, fontInfo.Name, fontInfo.Style)) * textPr.FontSize;
};
/**
 * Проверяем нужно ли сделать обязательный перенос строки после расчета одного диапазона
 * @returns {boolean}
 */
CParagraphRecalculateStateWrap.prototype.isForceLineBreak = function()
{
	return (this.ForceNewPage
		|| this.NewPage
		|| this.ForceNewLine
		|| this.LastHyphenItem);
};
CParagraphRecalculateStateWrap.prototype.isExceedConsecutiveAutoHyphenLimit = function()
{
	let limit = this.getAutoHyphenLimit();
	if (!limit)
		return false;
	
	let lines   = this.Paragraph.Lines;
	let curLine = this.Line - 1;
	
	while (curLine >= 0 && lines[curLine].Info & paralineinfo_AutoHyphen)
		--curLine;
	
	++curLine;
	
	return this.Line - curLine >= limit;
};
CParagraphRecalculateStateWrap.prototype.canPlaceAutoHyphenAfter = function(runItem)
{
	return (this.isAutoHyphenation()
		&& !this.isExceedConsecutiveAutoHyphenLimit()
		&& runItem.isHyphenAfter());
};
CParagraphRecalculateStateWrap.prototype.checkHyphenationZone = function(x)
{
	// Делаем как в MSWord (проверено в 2019 версии):
	// отмеряем сколько уже занято на текущей строке от начала строки, добавляем это значение к левому полю документа
	// и вычитаем из позиции правого поля параграфа. Если полученное значение больше hyphenationZone, значит можно
	// делать перенос.
	// Схема немного странная, т.к. мы считаем расстояние от левой границы параграфа, а добавляем его к левому полю,
	// поэтому при смещении параграфа целиком влево или вправо (одинаковом изменении левого и правого отступов)
	// разбиение может происходить по-разному, хотя ширина параграфа не меняется
	
	let paraPr = this.Paragraph.Get_CompiledPr2(false).ParaPr;
	
	let shift = paraPr.Ind.Left;
	if (this.UseFirstLine)
		shift += paraPr.Ind.FirstLine;
	
	return x - shift < this.XLimit - this.getHyphenationZone();
};
AscWord.ParagraphRecalculationWrapState = CParagraphRecalculateStateWrap;

function CParagraphRecalculateStateCounter(wrapState)
{
	this.wrapState   = wrapState;
    this.Paragraph   = undefined;
    this.Range       = undefined;
    this.Word        = false;
    this.SpaceLen    = 0;
    this.SpacesCount = 0;

    this.Words       = 0;
    this.Spaces      = 0;
    this.Letters     = 0;
    this.SpacesSkip  = 0;
    this.LettersSkip = 0;

    this.ComplexFields = new AscWord.ParagraphComplexFieldStack();
}
CParagraphRecalculateStateCounter.prototype.Reset = function(Paragraph, Range)
{
	this.Paragraph   = Paragraph;
	this.Range       = Range;
	this.Word        = false;
	this.SpaceLen    = 0;
	this.SpacesCount = 0;

	this.Words       = 0;
	this.Spaces      = 0;
	this.Letters     = 0;
	this.SpacesSkip  = 0;
	this.LettersSkip = 0;
};
CParagraphRecalculateStateCounter.prototype.isFastRecalculation = function()
{
	return this.wrapState.isFastRecalculation();
};

function CParagraphRecalculateStateAlign(wrapState)
{
	this.wrapState     = wrapState;
    this.X             = 0; // Текущая позиция по горизонтали
    this.Y             = 0; // Текущая позиция по вертикали
    this.XEnd          = 0; // Предельная позиция по горизонтали
    this.JustifyWord   = 0; // Добавочная ширина символов
    this.JustifySpace  = 0; // Добавочная ширина пробелов
    this.SpacesCounter = 0; // Счетчик пробелов с добавочной шириной (чтобы пробелы в конце строки не трогать)
    this.SpacesSkip    = 0; // Количество пробелов, которые мы пропускаем в начале строки
    this.LettersSkip   = 0; // Количество букв, которые мы пропускаем (из-за таба)
    this.LastW         = 0; // Ширина последнего элемента (необходимо для позиционирования картинки)
    this.Paragraph     = undefined;
    this.RecalcResult  = 0x00;//recalcresult_NextElement;

    this.Y0            = 0; // Верхняя граница строки
    this.Y1            = 0; // Нижняя граница строки

    this.CurPage       = 0;
    this.PageY         = 0;
    this.PageX         = 0;

    this.RecalcFast    = false; // Если пересчет быстрый, тогда все "плавающие" объекты мы не трогаем
    this.RecalcFast2   = false; // Второй вариант быстрого пересчета

	this.ComplexFields = new AscWord.ParagraphComplexFieldStack();
}
CParagraphRecalculateStateAlign.prototype.IsFastRangeRecalc = function()
{
	return this.RecalcFast;
};
CParagraphRecalculateStateAlign.prototype.getLogicDocument = function()
{
	return this.wrapState.Paragraph.GetLogicDocument();
};
CParagraphRecalculateStateAlign.prototype.getDocumentSettings = function()
{
	let logicDocument = this.Paragraph.GetLogicDocument();
	if (logicDocument && logicDocument.IsDocumentEditor())
		return logicDocument.getDocumentSettings();
	
	return AscWord.DEFAULT_DOCUMENT_SETTINGS;
};
CParagraphRecalculateStateAlign.prototype.getCompatibilityMode = function()
{
	return this.getDocumentSettings().getCompatibilityMode();
};
CParagraphRecalculateStateAlign.prototype.getLineTop = function()
{
	let p = this.Paragraph;
	return p.Pages[this.wrapState.Page].Y + p.Lines[this.wrapState.Line].Top;
};
CParagraphRecalculateStateAlign.prototype.getLineBottom = function()
{
	let p = this.Paragraph;
	return p.Pages[this.wrapState.Page].Y + p.Lines[this.wrapState.Line].Bottom;
};


function CParagraphRecalculateStateInfo()
{
	ParagraphRecalculateStateBase.call(this);
	this.fast          = false;
    this.Comments      = [];
    this.ComplexFields = [];
	this.PermRanges    = [];
}
CParagraphRecalculateStateInfo.prototype = Object.create(ParagraphRecalculateStateBase.prototype);
CParagraphRecalculateStateInfo.prototype.constructor = CParagraphRecalculateStateInfo;
CParagraphRecalculateStateInfo.prototype.setFast = function(isFast)
{
	this.fast = isFast;
};
CParagraphRecalculateStateInfo.prototype.isFastRecalculation = function()
{
	return this.fast;
};
CParagraphRecalculateStateInfo.prototype.Reset = function(prevInfo)
{
	this.Comments      = [];
	this.ComplexFields = [];
	this.PermRanges    = [];
	
	if (!prevInfo)
		return;
	
	if (prevInfo.Comments)
		this.Comments = prevInfo.Comments.slice();

	if (prevInfo.ComplexFields)
	{
		for (let index = 0, count = prevInfo.ComplexFields.length; index < count; ++index)
		{
			this.ComplexFields[index] = prevInfo.ComplexFields[index].Copy();
		}
	}
	
	if (prevInfo.PermRanges)
		this.PermRanges = prevInfo.PermRanges.slice();
};
CParagraphRecalculateStateInfo.prototype.AddComment = function(Id)
{
	this.Comments.push(Id);
};
CParagraphRecalculateStateInfo.prototype.RemoveComment = function(Id)
{
	var CommentsLen = this.Comments.length;
	for (var CurPos = 0; CurPos < CommentsLen; CurPos++)
	{
		if (this.Comments[CurPos] === Id)
		{
			this.Comments.splice(CurPos, 1);
			break;
		}
	}
};
CParagraphRecalculateStateInfo.prototype.addPermRange = function(rangeId)
{
	this.PermRanges.push(rangeId);
};
CParagraphRecalculateStateInfo.prototype.removePermRange = function(rangeId)
{
	let pos = this.PermRanges.indexOf(rangeId);
	if (-1 === pos)
		return;
	
	if (this.PermRanges.length - 1 === pos)
		--this.PermRanges.length;
	else
		this.PermRanges.splice(pos, 1);
};
CParagraphRecalculateStateInfo.prototype.processFieldChar = function(oFieldChar)
{
	if (!oFieldChar || !oFieldChar.IsUse())
		return;

	var oComplexField = oFieldChar.GetComplexField();

	if (oFieldChar.IsBegin())
	{
		this.ComplexFields.push(new CComplexFieldStatePos(oComplexField, true));
	}
	else if (oFieldChar.IsSeparate())
	{
		for (var nIndex = 0, nCount = this.ComplexFields.length; nIndex < nCount; ++nIndex)
		{
			if (oComplexField === this.ComplexFields[nIndex].ComplexField)
			{
				this.ComplexFields[nIndex].SetFieldCode(false);
				break;
			}
		}
	}
	else if (oFieldChar.IsEnd())
	{
		for (var nIndex = 0, nCount = this.ComplexFields.length; nIndex < nCount; ++nIndex)
		{
			if (oComplexField === this.ComplexFields[nIndex].ComplexField)
			{
				this.ComplexFields.splice(nIndex, 1);
				break;
			}
		}
	}
};
CParagraphRecalculateStateInfo.prototype.isComplexField = function()
{
	return (this.ComplexFields.length > 0 ? true : false);
};
CParagraphRecalculateStateInfo.prototype.isComplexFieldCode = function()
{
	if (!this.isComplexField())
		return false;

	for (var nIndex = 0, nCount = this.ComplexFields.length; nIndex < nCount; ++nIndex)
	{
		if (this.ComplexFields[nIndex].IsFieldCode())
			return true;
	}

	return false;
};
CParagraphRecalculateStateInfo.prototype.isHiddenComplexFieldPart = function()
{
	for (let fieldIndex = 0, fieldCount = this.ComplexFields.length; fieldIndex < fieldCount; ++ fieldIndex)
	{
		let isFieldCode = this.ComplexFields[fieldIndex].IsFieldCode();
		let isShowCode  = this.ComplexFields[fieldIndex].IsShowFieldCode();
		if (isFieldCode !== isShowCode)
			return true;
	}
	
	return false;
};
CParagraphRecalculateStateInfo.prototype.processFieldCharAndCollectComplexField = function(oChar)
{
	if (oChar.IsBegin())
	{
		var oComplexField = oChar.GetComplexField();
		if (!oComplexField)
		{
			oChar.SetUse(false);
		}
		else
		{
			oChar.SetUse(true);
			oComplexField.SetBeginChar(oChar);
			this.ComplexFields.push(new CComplexFieldStatePos(oComplexField, true));
		}
	}
	else if (oChar.IsEnd())
	{
		if (this.ComplexFields.length > 0)
		{
			oChar.SetUse(true);
			var oComplexField = this.ComplexFields[this.ComplexFields.length - 1].ComplexField;
			oComplexField.SetEndChar(oChar);
			this.ComplexFields.splice(this.ComplexFields.length - 1, 1);

			if (this.ComplexFields.length > 0 && this.ComplexFields[this.ComplexFields.length - 1].IsFieldCode())
				this.ComplexFields[this.ComplexFields.length - 1].ComplexField.SetInstructionCF(oComplexField);
		}
		else
		{
			oChar.SetUse(false);
		}
	}
	else if (oChar.IsSeparate())
	{
		if (this.ComplexFields.length > 0)
		{
			oChar.SetUse(true);
			var oComplexField = this.ComplexFields[this.ComplexFields.length - 1].ComplexField;
			oComplexField.SetSeparateChar(oChar);
			this.ComplexFields[this.ComplexFields.length - 1].SetFieldCode(false);
		}
		else
		{
			oChar.SetUse(false);
		}
	}
};
CParagraphRecalculateStateInfo.prototype.processInstruction = function(oInstruction)
{
	if (this.ComplexFields.length <= 0)
		return;

	var oComplexField = this.ComplexFields[this.ComplexFields.length - 1].ComplexField;
	if (oComplexField && null === oComplexField.GetSeparateChar())
		oComplexField.SetInstruction(oInstruction);
};

const g_PRSI = new CParagraphRecalculateStateInfo();

function CParagraphRecalculateObject()
{
    this.X      = 0;
    this.Y      = 0;
    this.XLimit = 0;
    this.YLimit = 0;

    this.Pages   = [];
    this.Lines   = [];
    this.Content = [];
}

CParagraphRecalculateObject.prototype =
{
    Save : function(Para)
    {
        this.X      = Para.X;
        this.Y      = Para.Y;
        this.XLimit = Para.XLimit;
        this.YLimit = Para.YLimit;

        this.Pages  = Para.Pages;
        this.Lines  = Para.Lines;

        var Content = Para.Content;
        var Count = Content.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            this.Content[Index] = Content[Index].SaveRecalculateObject();
        }
    },

    Load : function(Para)
    {
        Para.X      = this.X;
        Para.Y      = this.Y;
        Para.XLimit = this.XLimit;
        Para.YLimit = this.YLimit;

        Para.Pages = this.Pages;
        Para.Lines = this.Lines;

        var Count = Para.Content.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            Para.Content[Index].LoadRecalculateObject(this.Content[Index], Para);
        }
    },

    Get_DrawingFlowPos : function(FlowPos)
    {
        var Count = this.Content.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            this.Content[Index].Get_DrawingFlowPos( FlowPos );
        }
    }
};
