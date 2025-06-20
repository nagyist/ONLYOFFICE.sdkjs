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
var align_Right = AscCommon.align_Right;
var align_Left = AscCommon.align_Left;
var align_Center = AscCommon.align_Center;
var align_Justify = AscCommon.align_Justify;
var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
var g_oTableId = AscCommon.g_oTableId;
var History = AscCommon.History;

var g_dMathArgSizeKoeff_1 = 0.76;
var g_dMathArgSizeKoeff_2 = 0.6498; // 0.76 * 0.855

function CMathMenuSettings(oMathPr)
{
    if(oMathPr)
    {
        this.BrkBin = oMathPr.brkBin === BREAK_AFTER ? c_oAscMathInterfaceSettingsBrkBin.BreakAfter : c_oAscMathInterfaceSettingsBrkBin.BreakBefore;

        switch(oMathPr.defJc)
        {
            case align_Justify:
            {
                this.Justification  = c_oAscMathInterfaceSettingsAlign.Justify;
                break;
            }
            case align_Center:
            {
                this.Justification  = c_oAscMathInterfaceSettingsAlign.Center;
                break;
            }
            case align_Left:
            {
                this.Justification  = c_oAscMathInterfaceSettingsAlign.Left;
                break;
            }
            case align_Right :
            {
                this.Justification  = c_oAscMathInterfaceSettingsAlign.Right;
                break;
            }
            default:
            {
                this.Justification  = align_Justify;
                break;
            }
        }

        this.UseSettings    = oMathPr.dispDef;

        this.IntLim         = oMathPr.intLim === NARY_SubSup ? Asc.c_oAscMathInterfaceNaryLimitLocation.SubSup : Asc.c_oAscMathInterfaceNaryLimitLocation.UndOvr;
        this.NaryLim        = oMathPr.naryLim === NARY_SubSup ? Asc.c_oAscMathInterfaceNaryLimitLocation.SubSup : Asc.c_oAscMathInterfaceNaryLimitLocation.UndOvr;

        this.LeftMargin     = oMathPr.lMargin/10;
        this.RightMargin    = oMathPr.rMargin/10;
        this.WrapIndent     = oMathPr.wrapIndent/10;
        this.WrapRight      = oMathPr.wrapRight;
        this.SmallFraction  = oMathPr.smallFrac;
    }
    else
    {
        this.BrkBin          = undefined;
        this.Justification   = undefined;
        this.UseSettings     = undefined;
        this.IntLim          = undefined;
        this.NaryLim         = undefined;
        this.LeftMargin      = undefined;
        this.RightMargin     = undefined;
        this.WrapIndent      = undefined;
        this.WrapRight       = undefined;
        this.SmallFraction   = undefined;
    }
}
CMathMenuSettings.prototype.get_BreakBin = function(){ return this.BrkBin;};
CMathMenuSettings.prototype.put_BreakBin = function(BreakBin){ this.BrkBin = BreakBin;};
CMathMenuSettings.prototype.get_UseSettings = function(){ return this.UseSettings;};
CMathMenuSettings.prototype.put_UseSettings = function(UseSettings){ this.UseSettings = UseSettings;};
CMathMenuSettings.prototype.get_Justification = function(){ return this.Justification;};
CMathMenuSettings.prototype.put_Justification = function(Align){ this.Justification = Align;};
CMathMenuSettings.prototype.get_IntLim = function(){ return this.IntLim;};
CMathMenuSettings.prototype.put_IntLim = function(IntLim){ this.IntLim = IntLim;};
CMathMenuSettings.prototype.get_NaryLim = function(){ return this.NaryLim;};
CMathMenuSettings.prototype.put_NaryLim = function(NaryLim){ this.NaryLim = NaryLim;};
CMathMenuSettings.prototype.get_LeftMargin = function(){ return this.LeftMargin;};
CMathMenuSettings.prototype.put_LeftMargin = function(lMargin){ this.LeftMargin = lMargin;};
CMathMenuSettings.prototype.get_RightMargin = function(){ return this.RightMargin;};
CMathMenuSettings.prototype.put_RightMargin = function(rMargin){ this.RightMargin = rMargin;};
CMathMenuSettings.prototype.get_WrapIndent = function(){return this.WrapIndent;};
CMathMenuSettings.prototype.put_WrapIndent = function(WrapIndent){this.WrapIndent = WrapIndent;};
CMathMenuSettings.prototype.get_WrapRight = function(){ return this.WrapRight;};
CMathMenuSettings.prototype.put_WrapRight = function(WrapRight){ this.WrapRight = WrapRight;};
CMathMenuSettings.prototype.get_SmallFraction = function() {return this.SmallFraction;};
CMathMenuSettings.prototype.put_SmallFraction = function(SmallFrac) {this.SmallFraction = SmallFrac;};

window["CMathMenuSettings"]                         = CMathMenuSettings;
CMathMenuSettings.prototype["get_BreakBin"]        = CMathMenuSettings.prototype.get_BreakBin;
CMathMenuSettings.prototype["put_BreakBin"]        = CMathMenuSettings.prototype.put_BreakBin;
CMathMenuSettings.prototype["get_UseSettings"]     = CMathMenuSettings.prototype.get_UseSettings;
CMathMenuSettings.prototype["put_UseSettings"]     = CMathMenuSettings.prototype.put_UseSettings;
CMathMenuSettings.prototype["get_Justification"]   = CMathMenuSettings.prototype.get_Justification;
CMathMenuSettings.prototype["put_Justification"]   = CMathMenuSettings.prototype.put_Justification;
CMathMenuSettings.prototype["get_IntLim"]          = CMathMenuSettings.prototype.get_IntLim;
CMathMenuSettings.prototype["put_IntLim"]          = CMathMenuSettings.prototype.put_IntLim;
CMathMenuSettings.prototype["get_NaryLim"]         = CMathMenuSettings.prototype.get_NaryLim;
CMathMenuSettings.prototype["put_NaryLim"]         = CMathMenuSettings.prototype.put_NaryLim;
CMathMenuSettings.prototype["get_LeftMargin"]      = CMathMenuSettings.prototype.get_LeftMargin;
CMathMenuSettings.prototype["put_LeftMargin"]      = CMathMenuSettings.prototype.put_LeftMargin;
CMathMenuSettings.prototype["get_RightMargin"]     = CMathMenuSettings.prototype.get_RightMargin;
CMathMenuSettings.prototype["put_RightMargin"]     = CMathMenuSettings.prototype.put_RightMargin;
CMathMenuSettings.prototype["get_WrapIndent"]      = CMathMenuSettings.prototype.get_WrapIndent;
CMathMenuSettings.prototype["put_WrapIndent"]      = CMathMenuSettings.prototype.put_WrapIndent;
CMathMenuSettings.prototype["get_WrapRight"]       = CMathMenuSettings.prototype.get_WrapRight;
CMathMenuSettings.prototype["put_WrapRight"]       = CMathMenuSettings.prototype.put_WrapRight;
CMathMenuSettings.prototype["get_SmallFraction"]   = CMathMenuSettings.prototype.get_SmallFraction;
CMathMenuSettings.prototype["put_SmallFraction"]   = CMathMenuSettings.prototype.put_SmallFraction;


function Get_WordDocumentDefaultMathSettings()
{
    if (!editor || !editor.WordControl.m_oLogicDocument || !editor.WordControl.m_oLogicDocument.Settings)
        return new AscWord.MathSettings();

    return editor.WordControl.m_oLogicDocument.Settings.MathSettings;
}

function MathMenu(type, textPr)
{
	this.Type   = para_Math;
	this.Menu   = type == undefined ? c_oAscMathType.Default_Text : type;
	this.Text   = null;
	this.TextPr = textPr ? textPr : new CTextPr();
}
MathMenu.prototype.Get_Type = function()
{
	return this.Type;
};
MathMenu.prototype.GetType = function()
{
	return this.Type;
};
MathMenu.prototype.SetText = function(sText)
{
	this.Text = sText;
};
MathMenu.prototype.GetText = function()
{
	return this.Text;
};
MathMenu.prototype.GetTextPr = function()
{
	return this.TextPr;
};

function CMathLineState()
{
    this.StyleLine        = MATH_LINE_WRAP;
    this.Width            = 0;
    this.SpaceAlign       = 0;
    this.MaxWidth         = 0;
    this.WrapState        = ALIGN_EMPTY;
}

function CMathLineInfo()
{
    this.StyleLine        = MATH_LINE_WRAP;
    this.Measure          = 0;
    this.SpaceAlign       = 0;
    this.bWordLarge       = false;
}
CMathLineInfo.prototype.Get_GeneralWidth = function()
{
    return this.Measure + this.SpaceAlign;
};

function CParaMathLineParameters(FirstLineNumber)
{
    this.FirstLineNumber  = FirstLineNumber;
    this.WrapState        = ALIGN_EMPTY;

    this.LineParameters    = [];

    this.MaxW             = 0;

    this.bMathWordLarge   = false;
    this.NeedUpdateWrap   = true;
}
CParaMathLineParameters.prototype.UpdateWidth = function(Line, W)
{
    var bUpdMaxWidth = false;
    var NumLine = this.private_GetNumberLine(Line);

    if(Math.abs(this.LineParameters[NumLine].Measure - W) > 0.00001)
    {
        var Max = this.MaxW;
        var CountLines = this.LineParameters.length;

        this.LineParameters[NumLine].Measure = W;

        this.MaxW = this.LineParameters[0].Get_GeneralWidth();

        for(var Pos = 1; Pos < CountLines; Pos++)
        {
            var GeneralWidth = this.LineParameters[Pos].Get_GeneralWidth();

            if(this.MaxW < GeneralWidth)
                this.MaxW = GeneralWidth;
        }

        bUpdMaxWidth = Math.abs(Max - this.MaxW) > 0.0001;
    }

    return bUpdMaxWidth;
};
CParaMathLineParameters.prototype.Set_WordLarge = function(Line, bWordLarge)
{
    var NumLine = this.private_GetNumberLine(Line);

    if(this.WrapState !== ALIGN_EMPTY)
        bWordLarge = false;

    if(bWordLarge)
    {
        this.bMathWordLarge = true;
        this.LineParameters[NumLine].bWordLarge = true;
    }
    else
    {
        this.bMathWordLarge = false;

        var CountLines = this.LineParameters.length;

        for(var Pos = 0; Pos < CountLines; Pos++)
        {
            if(this.LineParameters[Pos].bWordLarge == true)
            {
                this.bMathWordLarge = true;
                break;
            }
        }

        this.LineParameters[NumLine].bWordLarge = false;
    }
};
CParaMathLineParameters.prototype.Add_Line = function(Line, bStartLine, bStartPage, AlignAt)
{
    var bInsideBounds = true;

    var NumLine = this.private_GetNumberLine(Line);

    if(NumLine >= this.LineParameters.length)
    {
        this.LineParameters[NumLine] = new CMathLineInfo();

        if(bStartLine)
        {
            this.LineParameters[NumLine].StyleLine = MATH_LINE_START;
        }
        else if(AlignAt !== undefined && AlignAt !== null)
        {
            this.LineParameters[NumLine].StyleLine  = MATH_LINE_ALiGN_AT;
            this.LineParameters[NumLine].SpaceAlign = AlignAt;
        }
        else
        {
            var MathSettings = Get_WordDocumentDefaultMathSettings();

            this.LineParameters[NumLine].StyleLine  = MATH_LINE_WRAP;
            this.LineParameters[NumLine].SpaceAlign = bStartPage == true ? MathSettings.Get_WrapIndent(this.WrapState) : 0;
        }


        bInsideBounds = false;
    }

    return bInsideBounds;
};
CParaMathLineParameters.prototype.Is_Large = function()
{
    return this.bMathWordLarge;
};
CParaMathLineParameters.prototype.Get_MaxWidth = function()
{
    return this.MaxW;
};
CParaMathLineParameters.prototype.private_GetNumberLine = function(NumLine)
{
    return NumLine - this.FirstLineNumber;
};
CParaMathLineParameters.prototype.Get_LineState = function(Line)
{
    var NumLine = this.private_GetNumberLine(Line);

    var LineInfo  = this.LineParameters[NumLine];

    var LineState = new CMathLineState();
    LineState.StyleLine  = LineInfo.StyleLine;
    LineState.Width      = LineInfo.Measure;
    LineState.SpaceAlign = this.Get_SpaceAlign(Line);
    LineState.MaxWidth   = this.MaxW;
    LineState.WrapState  = this.WrapState;

    return LineState;
};
CParaMathLineParameters.prototype.Get_SpaceAlign = function(Line)
{
    var NumLine = this.private_GetNumberLine(Line);

    var SpaceAlign = 0;

    if(this.bMathWordLarge == false && (this.WrapState == ALIGN_MARGIN_WRAP || this.WrapState == ALIGN_WRAP))
        SpaceAlign = this.LineParameters[NumLine].SpaceAlign;

    return SpaceAlign;
};


function CMathPageInfo()
{
    this.WPages            = [];    // widths on page
    this.StartLine         = -1;
    this.StartPage         = -1;
    this.CurPage           = -1;
    this.RelativePage      = -1;
}
CMathPageInfo.prototype.Reset = function()
{
    this.StartLine         = -1;
    this.StartPage         = -1;
    this.CurPage           = -1;
    this.RelativePage      = -1;
    this.WPages.length     =  0;
};
CMathPageInfo.prototype.Reset_Page = function(_Page)
{
    if(this.StartPage >= 0) // если нет, то только начали расчет формулы
    {
        var Page = _Page - this.StartPage;

        if(Page < this.WPages.length) // если нет, то только начали расчет страницы
        {
            // уберем из массива информацию о страницах, начиная с текущей
            // не делаем Reset для текущей страницы, т.к. это приведет к тому, что выставятся только параметры по умолчанию
            // а проверка на стартовую позицию рассчитана именно на длину массива this.WPages
            this.WPages.length = Page;
        }
    }
};
CMathPageInfo.prototype.Set_StartPos = function(Page, StartLine)
{
    this.StartPage   = Page;
    this.StartLine   = StartLine;
};
CMathPageInfo.prototype.Update_RelativePage = function(RelativePage)
{
    this.RelativePage = RelativePage;
};
CMathPageInfo.prototype.Update_CurrentPage = function(Page, ParaLine)
{
    this.CurPage = Page - this.StartPage;

    var Lng = this.WPages.length;
    if(this.CurPage >= Lng)
    {
        var FirstLineOnPage = ParaLine - this.StartLine;
        this.WPages[this.CurPage] = new CParaMathLineParameters(FirstLineOnPage);
    }
};
CMathPageInfo.prototype.Update_CurrentWrap = function(DispDef, bInline)
{
	if (this.WPages[this.CurPage].NeedUpdateWrap)
    {
        var WrapState;

        if(DispDef == false || bInline == true)
            WrapState = ALIGN_EMPTY;
        else if(this.CurPage == 0)
            WrapState = ALIGN_MARGIN_WRAP;
        else
            WrapState = ALIGN_MARGIN;

        this.WPages[this.CurPage].WrapState = WrapState;
        this.WPages[this.CurPage].NeedUpdateWrap = false;
    }
};
CMathPageInfo.prototype.Set_NeedUpdateWrap = function()
{
    this.WPages[this.CurPage].NeedUpdateWrap = true;
};
CMathPageInfo.prototype.Set_CurrentWrapState = function(WrapState)
{
    this.WPages[this.CurPage].WrapState = WrapState;
};
CMathPageInfo.prototype.Set_NextWrapState = function()
{
    var InfoPage = this.WPages[this.CurPage];

    if(InfoPage.WrapState !== ALIGN_EMPTY)
        InfoPage.WrapState++;
};
CMathPageInfo.prototype.Set_StateWordLarge = function(_Line, bWordLarge)
{
    this.WPages[this.CurPage].Set_WordLarge(_Line - this.StartLine, bWordLarge);
};
CMathPageInfo.prototype.Get_CurrentWrapState = function()
{
    return this.WPages[this.CurPage].WrapState;
};
CMathPageInfo.prototype.Get_CurrentStateWordLarge = function()
{
    return this.WPages[this.CurPage].Is_Large();
};
CMathPageInfo.prototype.UpdateWidth = function(_Line, Width)
{
    return this.WPages[this.CurPage].UpdateWidth(_Line - this.StartLine, Width);
};
CMathPageInfo.prototype.Get_MaxWidthOnCurrentPage = function()
{
    return this.WPages[this.CurPage].Get_MaxWidth();
};
CMathPageInfo.prototype.Get_FirstLineOnPage = function(_Page)
{
    var FirstLineOnPage = this.StartLine;
    var Page = _Page - this.StartPage;

    if(Page >= 0 && Page < this.WPages.length)
        FirstLineOnPage = this.StartLine + this.WPages[Page].FirstLineNumber;

    return FirstLineOnPage;
};
CMathPageInfo.prototype.Is_ResetNextPage = function(_Page)
{
    var bReset = true;

    if(this.CurPage == -1)
    {
        bReset = false;
    }
    else
    {
        var Page = _Page - this.StartPage;
        bReset = this.CurPage < Page;
    }

    return bReset;
};
CMathPageInfo.prototype.Is_ResetRelativePage = function(_RelativePage)
{
    return this.CurPage == -1 ? false : _RelativePage !== this.RelativePage;
};
CMathPageInfo.prototype.Is_FirstLineOnPage = function(_Line, _Page)
{
    var bFirstLine = true;

    if(this.StartPage >= 0) // если нет, то только начали расчет формулы
    {
        var Page = _Page - this.StartPage;

        if(Page < this.WPages.length) // если нет, то только начали расчет страницы
            bFirstLine = _Line == this.Get_FirstLineOnPage(_Page);
    }

    return bFirstLine;
};
CMathPageInfo.prototype.Get_LineState = function(_Line, _Page)
{
    var Page = _Page - this.StartPage;

    return this.WPages[Page].Get_LineState(_Line - this.StartLine);
};
// создаем новый объект с параметрами, необходимые для рассчета линий в ф-ии ParaMath.private_UpdateXLimits
// именно в этой функции нужно получить смещения линий (Wrap, Align)
// это происходит до принятия ширины контента для текущей линии (функция Update_CurrentWidth) и до получения флага WordLarge (функция Set_StateWordLarge)
CMathPageInfo.prototype.Add_Line = function(_Line, _Page, AlignAt)
{
    var Page       = _Page - this.StartPage,
        Line       = _Line - this.StartLine,
        bStartLine = _Line == this.StartLine,
        bStartPage = _Page == this.StartPage;

    this.WPages[Page].Add_Line(Line, bStartLine, bStartPage, AlignAt);
};
CMathPageInfo.prototype.Get_SpaceAlign = function(_Line, _Page)
{
    var Page       = _Page - this.StartPage,
        Line       = _Line - this.StartLine;

    return this.WPages[Page].Get_SpaceAlign(Line);
};



/**
 *
 * @constructor
 * @extends {CParagraphContentWithContentBase}
 */
function ParaMath()
{
	CParagraphContentWithContentBase.call(this);

    this.Id                 = AscCommon.g_oIdCounter.Get_NewId();
    this.Type               = para_Math;

    this.Jc                 = undefined;

    this.Root               = new CMathContent();
    this.Root.bRoot         = true;
    this.Root.ParentElement = this;

    this.X                  = 0;
    this.Y                  = 0;

    this.FirstPage          = -1;
    this.PageInfo           = new CMathPageInfo();

    this.ParaMathRPI        = new CMathRecalculateInfo();

    this.bSelectionUse      = false;
    this.Paragraph          = null;
    this.bFastRecalculate   = true;

    this.NearPosArray       = [];

    this.Width              = 0;
    this.WidthVisible       = 0;
    this.Height             = 0;
    this.Ascent             = 0;
    this.Descent            = 0;

    this.DispositionOpers   = [];

    this.DefaultTextPr      = new CTextPr();

    this.DefaultTextPr.FontFamily = {Name  : "Cambria Math", Index : -1 };
    this.DefaultTextPr.RFonts.SetAll("Cambria Math", -1);

    // Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	AscCommon.g_oTableId.Add( this, this.Id );
}

ParaMath.prototype = Object.create(CParagraphContentWithContentBase.prototype);
ParaMath.prototype.constructor = ParaMath;
ParaMath.prototype.Get_Type = function()
{
    return this.Type;
};

ParaMath.prototype.Get_Id = function()
{
    return this.Id;
};
ParaMath.prototype.IsMath = function()
{
	return true;
};
ParaMath.prototype.Copy = function(Selected, oPr)
{
    var NewMath = new ParaMath();
    NewMath.Root.bRoot = true;

    if(Selected)
    {
        var result = this.GetSelectContent();
        result.Content.CopyTo(NewMath.Root, Selected, oPr);
    }
    else
    {
        this.Root.CopyTo(NewMath.Root, Selected, oPr);
    }

    NewMath.Root.Correct_Content(true);

    return NewMath;
};
ParaMath.prototype.CopyContent = function(Selected)
{
    return [this.Copy(Selected)];
};
ParaMath.prototype.SetParagraph = function(Paragraph)
{
	this.Paragraph = Paragraph;
	this.Root.SetParagraph(Paragraph);
	this.Root.Set_ParaMath(this, null);
};
ParaMath.prototype.GetParagraph = function()
{
	return this.Paragraph;
};

ParaMath.prototype.Get_Text = function(Text)
{
	if (true === Text.BreakOnNonText)
		Text.Text = null;
    else if (true === Text.Math)
        Text.Text += this.GetText();
};

ParaMath.prototype.Is_Empty = function(oPr)
{
	if (this.Root.Content.length <= 0)
		return true;

	var isSkipPlcHldr = oPr && undefined !== oPr.SkipPlcHldr ? oPr.SkipPlcHldr : true;

	for (var nIndex = 0, nCount = this.Root.Content.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.Root.Content[nIndex];
		if (para_Math_Run !== oItem.Type || !oItem.Is_Empty({SkipPlcHldr : isSkipPlcHldr}))
			return false;
	}

	return true;
};

ParaMath.prototype.Is_CheckingNearestPos = function()
{
    return this.Root.Is_CheckingNearestPos();
};

ParaMath.prototype.IsStartFromNewLine = function()
{
    return false;
};

ParaMath.prototype.Get_TextPr = function(_ContentPos, Depth)
{
    var TextPr = new CTextPr();

    var mTextPr = this.Root.Get_TextPr(_ContentPos, Depth);
    TextPr.Merge( mTextPr );

    return TextPr;
};

ParaMath.prototype.Get_CompiledTextPr = function(isCopy)
{
	let oContent = this.GetSelectContent();

	let oTextPr = oContent.Content.Get_CompiledTextPr(isCopy);
	if (!oTextPr)
		return this.GetCompiledDefaultTextPr();

	return oTextPr;
};
ParaMath.prototype.GetCompiledDefaultTextPr = function()
{
	let oTextPr = new CTextPr();
	oTextPr.InitDefault();
	oTextPr.Merge(this.DefaultTextPr);
	return oTextPr;
};
ParaMath.prototype.GetDirectTextPr = function()
{
	let mathContent = this.GetSelectContent().Content;
	return mathContent.GetDirectTextPr();
};
/**
 * Добавляем элемент в текущую позицию (с учетом возможной глубины)
 * @param Item
 */
ParaMath.prototype.Add = function(Item)
{
	let logicDocument = this.GetLogicDocument();

	var Type = Item.Type;
	var oSelectedContent = this.GetSelectContent();

	var oContent = oSelectedContent.Content;

	var StartPos = oSelectedContent.Start;
	var Run = oContent.Content[StartPos];

	// Мы вставляем только в Run
	if (para_Math_Run !== Run.Type)
		return;

	var NewElement = null;
	if (para_Text === Type)
	{
		// заглушка для текстовых настроек плейсхолдера

		if(oContent.bRoot == false && Run.IsPlaceholder())
		{
			AscCommon.executeNoRevisions(function()
			{
				var CtrRunPr = oContent.Get_ParentCtrRunPr(false); // ctrPrp (не копия)
				Run.Apply_TextPr(CtrRunPr, undefined, true);
			}, logicDocument, this);
		}

		if(Item.Value == 38)
		{
			NewElement = new CMathAmp();
			Run.Add(NewElement, true);
		}
		else
		{
			NewElement = new CMathText(false);
			NewElement.add(Item.Value);
			Run.Add(NewElement, true);
		}
	}
	else if (para_Space === Type)
	{
		NewElement = new CMathText(false);
		NewElement.add(32);
		Run.Add(NewElement, true);
	}
	else if (Item instanceof AscWord.ParaMath)
	{
		let contentPos = new AscWord.CParagraphContentPos();
		
		if (this.bSelectionUse)
			this.Get_ParaContentPos(true, true, contentPos);
		else
			this.Get_ParaContentPos(false, false, contentPos);
		
		let rightRun = Run.Split2(Run.State.ContentPos);
		
		let pos = StartPos + 1;
		for (let i = 0; i < Item.Root.Content.length; ++i, ++pos)
		{
			oContent.Internal_Content_Add(pos, Item.Root.Content[i], false);
		}
		
		oContent.Internal_Content_Add(pos, rightRun, false);
		oContent.CurPos = pos;
		rightRun.MoveCursorToStartPos();
		
		oContent.Correct_ContentCurPos();
	}
	else if (para_Math === Type)
	{
		var ContentPos = new AscWord.CParagraphContentPos();

		if(this.bSelectionUse == true)
			this.Get_ParaContentPos(true, true, ContentPos);
		else
			this.Get_ParaContentPos(false, false, ContentPos);
		
		var TextPr = this.Root.GetMathTextPrForMenu(ContentPos, 0);
		if (Item.TextPr)
			TextPr.Merge(Item.TextPr);
		
		// Нам нужно разделить данный Run на 2 части
		var RightRun = Run.Split2(Run.State.ContentPos);

		oContent.Internal_Content_Add(StartPos + 1, RightRun, false);
		// Выставляем позицию в начало этого рана
		oContent.CurPos = StartPos + 1;
		RightRun.MoveCursorToStartPos();

		oContent.Load_FromMenu(Item.Menu, this.Paragraph, TextPr, Item.GetText());
		oContent.Correct_ContentCurPos();
	}
	
	if ((para_Text === Type || para_Space === Type) && null !== NewElement)
		this.ProcessAutoCorrect(oContent, NewElement);

	// Корректируем данный контент
	oContent.Correct_Content(true);
};
ParaMath.prototype.ProcessAutoCorrect = function(content, triggerElement)
{
	if (!content)
	{
		let selectContent = this.GetSelectContent();
		if (!selectContent)
			return;
		
		content = selectContent.Content;
	}
	
	if (!triggerElement)
	{
		triggerElement = new CMathText();
		triggerElement.add(32);
	}
	
	this.bFastRecalculate = content.bOneLine == false; // многострочный контент => можно осуществлять быстрый пересчет
	content.Process_AutoCorrect(triggerElement);
};
/**
 * Добавляем элемент в конец корневого контента
 * @param oElement
 */
ParaMath.prototype.Push = function(oElement)
{
	this.Root.AddToContent(this.Root.GetElementsCount(), oElement);
};
/**
 * Добавляем все элементы заданного ParaMath в конец текущего
 * @param oMath {ParaMath}
 */
ParaMath.prototype.Concat = function(oMath)
{
	let nCount = oMath.Root.GetElementsCount();
	for (let nIndex = 0, nCount = oMath.Root.GetElementsCount(); nIndex < nCount; ++nIndex)
	{
		this.Push(oMath.Root.GetElement(nIndex));
	}
	oMath.Root.RemoveFromContent(0, nCount);
};

ParaMath.prototype.Get_AlignToLine = function(_CurLine, _CurRange, _Page, _X, _XLimit)
{
    // отступ первой строки не учитывается для неинлайновых формул
    var X = _X;

    var MathSettings = Get_WordDocumentDefaultMathSettings();

    // выставим сначала Position до пересчета выравнивания для формулы
    // для расчета смещений относительно операторов

    var PosInfo = new CMathPosInfo();

    PosInfo.CurLine  = _CurLine;
    PosInfo.CurRange = _CurRange;

    if(true == this.NeedDispOperators(_CurLine))
    {
        this.DispositionOpers.length = 0;
        PosInfo.DispositionOpers = this.DispositionOpers;
    }

    var pos   = new CMathPosition();
    this.Root.setPosition(pos, PosInfo);

    var XStart, XEnd;

    if(this.ParaMathRPI.bInline == false)
    {
        XStart = this.ParaMathRPI.XStart;
        XEnd   = this.ParaMathRPI.XEnd;
    }
    else
    {
        XStart = _X;
        XEnd   = _XLimit;
    }

    var Page = this.Paragraph == null ? 0 : this.Paragraph.Get_AbsolutePage(_Page);
    var LineState = this.PageInfo.Get_LineState(_CurLine, Page);
    var StyleLine = LineState.StyleLine,
        WidthLine = LineState.Width,
        MaxWidth  = LineState.MaxWidth,
        WrapLine  = LineState.SpaceAlign,
        WrapState = LineState.WrapState;

    XStart += MathSettings.Get_LeftMargin(WrapState);
    XEnd   -= MathSettings.Get_RightMargin(WrapState);

    var Jc = this.Get_Align();

    if(StyleLine == MATH_LINE_START || StyleLine == MATH_LINE_ALiGN_AT) // первая строка первой страницы, если строка разбивается на несколько отрезков, то это уже будет inline-формула => ф-ия Get_AlignToLine не будет вызвана
    {                                                                   // либо строка выровнена относительно первой строки
        switch(Jc)
        {
            case align_Left:    X = XStart + WrapLine; break;
            case align_Right:   X = Math.max(XEnd - WidthLine  + WrapLine, XStart); break;
            case align_Center:  X = Math.max(XStart + (XEnd - XStart - WidthLine)/2  + WrapLine, XStart); break;
            case align_Justify:
            {
                X = Math.max(XStart + (XEnd - XStart - MaxWidth)/2 + WrapLine , XStart);
                break;
            }
        }
    }
    else
    {
        if(true == MathSettings.Get_WrapRight()) // флаг свидетельствует о том, что строки кроме первой и строк, выровненных относительно первой, нужно размещать окол правой границы
        {
            X = Math.max(XEnd - WidthLine  + WrapLine, XStart);
        }
        else if(Jc == align_Justify)
        {
            X = XEnd - XStart > MaxWidth ? XStart + (XEnd - XStart - MaxWidth)/2 + WrapLine : XStart;
        }
        else
        {
            X = XEnd - XStart > MaxWidth ? XStart + WrapLine : XStart;
        }
    }


    return X;
};

ParaMath.prototype.Remove = function(Direction, bOnAddText)
{
    var TrackRevisions = null;
    if (this.Paragraph && this.Paragraph.LogicDocument)
        TrackRevisions = this.Paragraph.LogicDocument.IsTrackRevisions();

    var oSelectedContent = this.GetSelectContent();

    var nStartPos = oSelectedContent.Start;
    var nEndPos = oSelectedContent.End;
    var oContent = oSelectedContent.Content;

    if (nStartPos === nEndPos)
    {
        var oElement = oContent.getElem(nStartPos);
        var ElementReviewType = oElement.GetReviewType();

        // Если данный элемент - ран, удаляем внутри рана, если нет, тогда удаляем целиком элемент
        if (para_Math_Run === oElement.Type)
        {
            if ((true === oElement.IsPlaceholder() && !bOnAddText) && oElement.Parent.bRoot == true)
            {
                this.Root.Remove_FromContent(0, 1);
                return true;
            }
            else if ((true === oElement.IsPlaceholder() && !bOnAddText) || (false === oElement.IsPlaceholder() && false === oElement.Remove(Direction) && true !== this.bSelectionUse))
            {
                if ((Direction > 0 && oContent.Content.length - 1 === nStartPos) || (Direction < 0 && 0 === nStartPos))
                {
                    // Проверяем находимся ли мы на верхнем уровне
                    if (oContent.bRoot)
                        return false;

                    // Значит мы в каком-то элементе, тогда надо выделить данный элемент
                    oContent.ParentElement.Select_WholeElement();

                    // выставляем выделение для родительского элемента
                    // TODO нужно понять почему mathContent.Select_WholeElement() не выделяет родительский элемент
                    this.Parent.SelectFotMath();

                    return true;
                }

                if (Direction > 0)
                {
                    var oNextElement = oContent.getElem(nStartPos + 1);
                    if (para_Math_Run === oNextElement.Type)
                    {
                        // Здесь мы не проверяем результат Remove, потому что ран не должен быть пустым после
                        // Correct_Content
                        oNextElement.MoveCursorToStartPos();
                        oNextElement.Remove(1);

                        if (oNextElement.Is_Empty())
                        {
                            oContent.Correct_Content();
                            oContent.Correct_ContentPos(1);
                        }

                        this.RemoveSelection();
                    }
                    else
                    {
						this.SelectThisElement(1);
						this.RemoveSelection();
                        oContent.SelectElementByPos(nStartPos + 1);
                    }
                }
                else
                {
                    var oPrevElement = oContent.getElem(nStartPos - 1);
                    if (para_Math_Run === oPrevElement.Type)
                    {
                        // Здесь мы не проверяем результат Remove, потому что ран не должен быть пустым после
                        // Correct_Content
                        oPrevElement.MoveCursorToEndPos();
                        oPrevElement.Remove(-1);

                        if (oPrevElement.Is_Empty())
                        {
                            oContent.Correct_Content();
                            oContent.Correct_ContentPos(-1);
                        }

                        this.RemoveSelection();
                    }
                    else
                    {
						this.SelectThisElement(1);
						this.RemoveSelection();
						oContent.SelectElementByPos(nStartPos - 1);
                    }
                }
            }
            else
            {
                if (oElement.Is_Empty())
                {
                    oContent.CurPos = nStartPos;
                    oContent.Correct_Content();
                    oContent.Correct_ContentPos(-1); // -1, потому что нам надо встать перед элементом, а не после
                }

                this.RemoveSelection();
            }

            return true;
        }
        else
        {
            this.RemoveSelection();
            if (true === TrackRevisions)
            {
                if (reviewtype_Common === ElementReviewType)
                {
                    if (para_Math_Run === oElement.Type !== oElement.Type)
                        oElement.RejectRevisionChanges(c_oAscRevisionsChangeType.TextAdd, true);

                    oElement.SetReviewType(reviewtype_Remove);
                }
                else if (reviewtype_Add === ElementReviewType)
                {
                    oContent.Remove_FromContent(nStartPos, 1);
                    if (para_Math_Run === oContent.Content[nStartPos].Type)
                        oContent.Content[nStartPos].MoveCursorToStartPos();
                }
            }
            else
            {
                oContent.Remove_FromContent(nStartPos, 1);
                if (para_Math_Run === oContent.Content[nStartPos].Type)
                    oContent.Content[nStartPos].MoveCursorToStartPos();
            }
            oContent.CurPos = nStartPos;
            oContent.Correct_Content();
            oContent.Correct_ContentPos(-1); // -1, потому что нам надо встать перед элементом, а не после
        }
    }
    else
    {
        if (nStartPos > nEndPos)
        {
            var nTemp = nEndPos;
            nEndPos = nStartPos;
            nStartPos = nTemp;
        }

        // Проверяем начальный и конечный элементы
        var oStartElement = oContent.getElem(nStartPos);
        var oEndElement = oContent.getElem(nEndPos);

        if (true === TrackRevisions)
        {
            for (var CurPos = nEndPos; CurPos >= nStartPos; --CurPos)
            {
                var Element = oContent.getElem(CurPos);
                var ElementReviewType = Element.GetReviewType();

                if (para_Math_Run === Element.Type && (CurPos === nEndPos || CurPos === nStartPos))
                {
                    // Удаление разруливается внутри рана
                    Element.Remove(Direction);
                }
                else
                {
                    if (reviewtype_Common === ElementReviewType)
                    {
                        if (para_Math_Run === Element.Type !== Element.Type)
                            Element.RejectRevisionChanges(c_oAscRevisionsChangeType.TextAdd, true);

                        Element.SetReviewType(reviewtype_Remove);
                    }
                    else if (reviewtype_Add === ElementReviewType)
                    {
                        oContent.Remove_FromContent(CurPos, 1);
                        nEndPos--;
                    }
                }
            }

            this.RemoveSelection();
            if (Direction < 0)
            {
                oContent.CurPos = nStartPos;
            }
            else
            {
                oContent.CurPos = nEndPos;
                if (para_Math_Run === oContent.Content[nEndPos].Type)
                    oContent.Content[nEndPos].MoveCursorToStartPos();
            }
        }
        else
        {
            // Если последний элемент - ран, удаляем внутри, если нет, тогда удаляем целиком элемент
            if (para_Math_Run === oEndElement.Type)
                oEndElement.Remove(Direction);
            else
                oContent.Remove_FromContent(nEndPos, 1);

            // Удаляем все промежуточные элементы
            oContent.Remove_FromContent(nStartPos + 1, nEndPos - nStartPos - 1);

            // Если первый элемент - ран, удаляем внутри рана, если нет, тогда удаляем целиком элемент
            if (para_Math_Run === oStartElement.Type)
                oStartElement.Remove(Direction);
            else
                oContent.Remove_FromContent(nStartPos, 1);

            this.RemoveSelection();
            oContent.CurPos = nStartPos;
        }
        oContent.Correct_Content();
        oContent.Correct_ContentPos(Direction);

		// если в контенте остался только плейсхолдер, то нам нужно выделить его
		if (oContent.IsPlaceholder())
		{
			oContent.SelectThisElement(1);
			oContent.SelectAll(1);
		}
    }
};

ParaMath.prototype.GetSelectContent = function(isAll)
{
    return this.Root.GetSelectContent(isAll);
};

ParaMath.prototype.GetCurrentParaPos = function(align)
{
    return this.Root.GetCurrentParaPos();
};

ParaMath.prototype.Apply_TextPr = function(TextPr, IncFontSize, ApplyToAll)
{
    if(ApplyToAll == true) // для ситуации, когда ApplyToAll = true, в Root формулы при этом позиции селекта не проставлены
    {
        this.Root.Apply_TextPr(TextPr, IncFontSize, true);
    }
    else
    {
        var content = this.GetSelectContent().Content;

        var NewTextPr = new CTextPr();
        var bSetInRoot = false;


        if(IncFontSize == undefined)
        {
            if(TextPr.Underline !== undefined)
            {

                NewTextPr.Underline   = TextPr.Underline;
                bSetInRoot            = true;
            }

            if(TextPr.FontSize !== undefined && content.IsNormalTextInRuns() == false)
            {
                NewTextPr.FontSize    = TextPr.FontSize;
                bSetInRoot            = true;

            }

            content.Apply_TextPr(TextPr, IncFontSize, ApplyToAll);

            if(bSetInRoot)
                this.Root.Apply_TextPr(NewTextPr, IncFontSize, true);
        }
        else
        {

            if(content.IsNormalTextInRuns() == false)
                this.Root.Apply_TextPr(TextPr, IncFontSize, true);
            else
                content.Apply_TextPr(TextPr, IncFontSize, ApplyToAll);
        }
    }
};

ParaMath.prototype.Clear_TextPr = function()
{

};

ParaMath.prototype.Check_NearestPos = function(ParaNearPos, Depth)
{
    this.Root.Check_NearestPos(ParaNearPos, Depth);
};

ParaMath.prototype.Get_DrawingObjectRun = function(Id)
{
    return null;
};

ParaMath.prototype.GetRunByElement = function(oRunElement)
{
	return null;
};

ParaMath.prototype.Get_DrawingObjectContentPos = function(Id, ContentPos, Depth)
{
    return false;
};

ParaMath.prototype.Get_Layout = function(DrawingLayout, UseContentPos, ContentPos, Depth)
{
    if (true === UseContentPos)
        DrawingLayout.Layout = true;
    else
        DrawingLayout.X += this.Width;
};

ParaMath.prototype.GetNextRunElements = function(oRunElements, isUseContentPos, nDepth)
{
	if (oRunElements.IsBreakOnMath())
		return oRunElements.Stop();
	
	if (oRunElements.IsSkipMath())
		return;
	
	this.Root.GetNextRunElements(oRunElements, isUseContentPos, nDepth);
};
ParaMath.prototype.GetPrevRunElements = function(oRunElements, isUseContentPos, nDepth)
{
	if (oRunElements.IsBreakOnMath())
		return oRunElements.Stop();
	
	if (oRunElements.IsSkipMath())
		return;

	this.Root.GetPrevRunElements(oRunElements, isUseContentPos, nDepth);
};

ParaMath.prototype.CollectDocumentStatistics = function(ParaStats)
{
    // TODO: ParaMath.CollectDocumentStatistics
};

ParaMath.prototype.Create_FontMap = function(Map)
{
    // Styles.js
    // Document_CreateFontMap

    this.Root.Create_FontMap(Map);
};

ParaMath.prototype.Get_AllFontNames = function(AllFonts)
{
    // выставить для всех шрифтов, к-ые используются в AllFonts true
    AllFonts["Cambria Math"] = true;

    this.Root.Get_AllFontNames(AllFonts);
};

ParaMath.prototype.GetSelectedElementsInfo = function(Info, ContentPos, Depth)
{
    Info.SetMath(this);
};

ParaMath.prototype.GetSelectedText = function(bAll, bClearText, oPr)
{
	if (true === bAll || true === this.IsSelectionUse())
	{
		if (true === bClearText)
			return null;

		if (oPr && false === oPr.Math)
			return "";

		var res             = "";
		var selectedContent = this.GetSelectContent(bAll);
		if (selectedContent && selectedContent.Content && selectedContent.Content.GetTextContent)
		{
			var textContent = selectedContent.Content.GetTextContent(!bAll);
			if (textContent && textContent.str)
			{
				res = textContent.str;
			}
		}
		return res;
	}
	return "";
};

ParaMath.prototype.GetText = function(isLaTeX)
{
    let oMathText = this.GetTextOfElement(isLaTeX);
    return oMathText.GetText();
};

ParaMath.prototype.GetTextOfElement = function (isLaTeX, isDefaultText)
{
    return this.Root.GetTextOfElement(isLaTeX, isDefaultText);
};

ParaMath.prototype.GetSelectDirection = function()
{
    return this.Root.GetSelectDirection();
};

ParaMath.prototype.Clear_TextFormatting = function( DefHyper )
{
};

ParaMath.prototype.CanAddDropCap = function()
{
    return false;
};

ParaMath.prototype.Get_TextForDropCap = function(DropCapText, UseContentPos, ContentPos, Depth)
{
    if ( true === DropCapText.Check )
        DropCapText.Mixed = true;
};

ParaMath.prototype.Get_StartTabsCount = function(TabsCounter)
{
    return false;
};

ParaMath.prototype.Remove_StartTabs = function(TabsCounter)
{
    return false;
};

ParaMath.prototype.Add_ToContent = function(Pos, Item, UpdatePosition)
{

};

ParaMath.prototype.Get_MenuProps = function()
{
    return this.Root.Get_MenuProps();
};
ParaMath.prototype.Set_MenuProps = function(Props)
{
    if(Props != undefined)
        this.Root.Set_MenuProps(Props);
};

ParaMath.prototype.CheckRunContent = function(fCheck, oStartPos, oEndPos, nDepth, oCurrentPos, isForward)
{
	return this.Root.CheckRunContent(fCheck, oStartPos, oEndPos, nDepth, oCurrentPos, isForward);
};

//-----------------------------------------------------------------------------------
// Функции пересчета
//-----------------------------------------------------------------------------------
ParaMath.prototype.Recalculate_Range = function(PRS, ParaPr, Depth)
{
    // Paragraph_Recalculate.js
    // private_RecalculateFastRange

    // Document.js
    // Recalculate
    // SimpleChanges

    // Paragraph.js
    // CRunRecalculateObject.Compare

	if (PRS.bFastRecalculate == true && this.bFastRecalculate == false)
		return;

	if (this.Paragraph !== PRS.Paragraph)
	{
		this.Paragraph = PRS.Paragraph;
		this.private_UpdateSpellChecking();
	}

    var Para         = PRS.Paragraph;
    var ParaLine     = PRS.Line;
    var ParaRange    = PRS.Range;
    var Page         = this.Paragraph == null ? 0 : this.Paragraph.Get_AbsolutePage(PRS.Page);
    var RelativePage = PRS.Page;

    var bStartRange  = this.Root.IsStartRange(ParaLine, ParaRange);
    var isInline     = this.IsInline();

	// первый пересчет
	var PrevLineObject = PRS.GetMathRecalcInfoObject();
	var bCurrentObj    = PrevLineObject == null || PrevLineObject == this;
	var PrevObject     = bCurrentObj ? null : PrevLineObject;

    var bStartRecalculate = PrevLineObject == null && true == bStartRange && PRS.bFastRecalculate == false;
    var bContinueRecalc   = !isInline && PRS.bContinueRecalc === true;

    var LDRecalcInfo  = this.Paragraph.Parent.RecalcInfo;

	if (bStartRecalculate == true && bContinueRecalc == false) // первый пересчет
	{
		// информация о пересчете
		var RPI = new CRPI();
		RPI.MergeMathInfo(this.ParaMathRPI);
		this.Root.PreRecalc(null, this, new CMathArgSize(), RPI);

		this.PageInfo.Reset();
		this.PageInfo.Set_StartPos(Page, ParaLine);
		this.PageInfo.Update_RelativePage(RelativePage);
		this.ParaMathRPI.Reset(PRS, ParaPr);
	}
    else
	{
		// true == this.PageInfo.Is_ResetNextPage(Page)
		/// при переходе на следующую страницу выставляем стартовые параметры для отрезка, в к-ом пересчитываем
		// может произойти в одной из 2-х ситуаций:
		// 1. первый раз пересчитываем формулу => для PageInfo ширины и др . параметры еще не рассчитали
		// 2. произошли изменения на пред страницах, их пересчитали, перешли к следующей => для PageInfo нужно выставить дефолтные настройки для параметров и обнулить массив ширин
		// параметры для ParaMathRPI выставляем дефолтные в любом из этих двух случаев

		// false == this.PageInfo.Is_ResetNextPage(Page) && true == this.PageInfo.Is_FirstLineOnPage(Line, Page)
		// т.е. рассчитываем текущую страницу с первой строки
		// может произойти, если вновь стали (PrevLineObject !== null) пересчитывать формулу на данной странице (из-за того что изменилась макс ширина и нужно заново пересчитать формулу на странице и т.п.)
		// или же произошли какие-то изменения на странице и вызвался пересчет для этой страницы (PrevLineObject == null) и отсутствует быстрый пересчет (PRS.bFastRecalculate == false)

		var bResetNextPage = true == this.PageInfo.Is_ResetNextPage(Page);
		var bResetPageInfo = PrevLineObject == null && bContinueRecalc == false && PRS.bFastRecalculate == false && true == this.PageInfo.Is_FirstLineOnPage(ParaLine, Page);

		if (bResetNextPage == true || bResetPageInfo == true)
		{
			this.ParaMathRPI.Reset(PRS, ParaPr);
			this.PageInfo.Reset_Page(Page);
		}

		if (true == this.PageInfo.Is_ResetRelativePage(PRS.Page))
		{
			this.ParaMathRPI.Reset(PRS, ParaPr);
			this.PageInfo.Update_RelativePage(RelativePage);
		}
	}

    PRS.MathNotInline = !isInline; // если неинлайновая формула, то рассчитываем Ranges по максимальному измерению
    PRS.bPriorityOper = !isInline;

	if (!isInline)
	{
		// чтобы на проверке bResetPageInfo не перебить параметры
		PRS.SetMathRecalcInfoObject(this);
	}


	// такая сиуация возможна, если разместили формулу под картинкой и нужно заново пересчитать формулу
	// для этого меняем PRS.Y и ждем пока не произойдет private_RecalculateLineCheckRangeY
	// т.к. в противном случае мы можем изменить Ranges на те, что находятся PRS.Y, который был до того как его изменили в данном блоке (возникает из-за того, что ф-ия private_RecalculateLineCheckRanges нах-ся выше ф-ии private_RecalculateLineCheckRangeY)
	if (this.ParaMathRPI.ShiftY - PRS.Y > 0 || PRS.bMathRangeY == true)
	{
		// выполняем перенос здесь для того, чтобы перенести под нужную картинку
		// к примеру, у нас есть нескоолько картинок с выравниванием по контору и картинка с выравниванием top_bottom(она расположена ниже всех остальных картинок) =>
		// нужно сделать перенос под картинку top_bottom и пересчитывать заново формулу
		//
		//this.Set_EmptyRange(PRS);
		PRS.bMathRangeY  = true;
		PRS.ForceNewPage = true;
		this.private_UpdateRangeY(PRS, this.ParaMathRPI.ShiftY);
	}
	else
	{
		// такая ситуация возможна  когда пришел пересчет заново и кол-во отрезков выравнивания 0 (либо меньше, чем нужно)
		// при этом если это первый Range данной формулы, то пришел еще Reset, то есть пересчитать придется, при этом не меняем max ширину, т.к. если мы уже пересчитали с учетом Range, она не должна поменяться
		this.Root.SetParagraph(Para);
		this.Root.Set_ParaMath(this, null);

		this.PageInfo.Update_CurrentPage(Page, ParaLine);

		var bRecalcNormal = true;

		if (bCurrentObj == true && !isInline && PRS.bFastRecalculate == false) // меняем отрезки обтекания только для случая, когда пересчитываем текущий объект (а не вновь возвращаемся к последнему пересчету)
		{
			var UpdWrap = this.private_UpdateWrapSettings(PRS, ParaPr);

			if (UpdWrap == MATH_UPDWRAP_NEWRANGE)
			{
				// Уберем из массива информацию о рассчитанных ширинах, чтобы не учлась рассчитанная ранее максимальная ширина (в связи с тем, что отрезок, в к-ом нужно расположить, изменился по ширине)
				// выставляем EmptyLine = false, т.к. нужно сделать заново пересчет в новом отрезке (а не перенести формулу под картинку)
				// т.к. инициируем пересчет заново, то в проверку на ParaNewLine : if (true === NotInlineMath && true !== PRS.EmptyLine) не зайдем, т.к. NewRange = true
				this.private_SetRestartRecalcInfo(PRS);
				this.PageInfo.Reset_Page(Page);
				this.ParaMathRPI.bInternalRanges = true;
				// не выставляем EmtyLine = false, т.к. так и так выйдем из пересчета данной строки при расчете Ranges, до пересчета картинок не дойдем, поэтому PRS.EmptyLine = false  выставлять не нужно
				//PRS.EmptyLine = false;
			}
			else if (UpdWrap == MATH_UPDWRAP_UNDERFLOW)
			{
				// пересчитаем PRS.Y на след пересчете при this.ParaMathRPI.ShiftY - PRS.Y > 0
				// в блоке if(this.ParaMathRPI.ShiftY - PRS.Y > 0)
			}

			bRecalcNormal = UpdWrap == MATH_UPDWRAP_NOCHANGES; // пересчитываем всю строку заново, т.к. может получиться так, что добавилась еще одна картинка (+ к уже существующим) и нужно заново выбрать Range, в котором необходимо разместить формулу
		}

		if (bRecalcNormal == true) // пересчет в штатном режиме
		{
			var MathSettings = Get_WordDocumentDefaultMathSettings();

			var DispDef = MathSettings.Get_DispDef(),
				bInline = this.Is_Inline(); // учитываем, если формула внутристроковая или же разбивается плавающим объектом (в этом случае тоже нужно рассчитывать как инлайновую)

			//здесь обновляем WrapState, исходя из этого параметра будем считать WrapIndent
			this.PageInfo.Update_CurrentWrap(DispDef, bInline);


			// формулы не инлайновая, есть Ranges пересчитываем формулу в макс Range => private_RecalculateRangeInsideInterval
			if (this.ParaMathRPI.IntervalState !== MATH_INTERVAL_EMPTY && this.ParaMathRPI.bInternalRanges == true/*this.ParaMathRPI.bStartRanges == true*/) // картинки в другом параграфе и формула пересчитывается с учетом Ranges
			{
				// X и XEnd не перебиваем выше, т.к. они понадобятся для учета попадания в Range в ф-ии  private_RecalculateRangeInsideInterval
				this.private_RecalculateRangeInsideInterval(PRS, ParaPr, Depth);
			}
			else
			{
				this.private_RecalculateRangeWrap(PRS, ParaPr, Depth);
			}

			this.ParaMathRPI.ClearRecalculate();
		}
	}

    // обновляем LDRecalcInfo здесь, т.к. формула - многострочный мат объект, нельзя каждый раз изменять LDRecalcInfo => иначе для четных/нечетных сток будет чередование ParaMath = this/null
    if (PRS.RecalcResult & recalcresult_ParaMath && (true === LDRecalcInfo.Can_RecalcObject() || true === LDRecalcInfo.Check_ParaMath(this)))
    {
        LDRecalcInfo.Set_ParaMath(this);
    }
    else if(true === LDRecalcInfo.Check_ParaMath(this))
    {
        LDRecalcInfo.Reset();
    }
    else
    {

    }

	if (bCurrentObj == true)
	{
		// этот параметр необходим для пересчета нескольких неинлайновых (и инлайновых формул) внутри одного параграфа,
		// чтобы не перетирать параметры внутри пересчитанных формул (и соответственно избежать зацикливания)
		PRS.bContinueRecalc = true;

		if (PRS.NewRange == false)
		{
			// на случай когда у нас несколько неинлайновых формул в одном параграфе
			PRS.SetMathRecalcInfoObject(null);
			
			// TODO: Вообще инлайновая формула должна вести себя как буква на строке, т.е. идти как часть слова, но
			//       пока будем считать, что на ней всегда заканчивается слово
			
			PRS.X += PRS.WordLen + PRS.SpaceLen;
			
			PRS.Word            = false;
			PRS.FirstItemOnLine = false;
			PRS.EmptyLine       = false;
			PRS.TextOnLine      = true;
			PRS.SpaceLen        = 0;
			PRS.WordLen         = 0;
		}
	}
	else
	{
		PRS.SetMathRecalcInfoObject(PrevObject); // возвращаем формулу, которая инициировала пересчет (если это была текущая формула, то null)
	}
};
ParaMath.prototype.private_UpdateWrapSettings = function(PRS)
{
    // запомним PRS.Ranges.Y для смещения, чтобы выставить потом смещение, т.к. возможен случай, что картинка, под которой нужно расположить формулу, будет не первой, которая встретиться, пр первом пересчете, или же будет отсутствовать в текущем пересчете
    // (т.к. надо расположить под картинков), отсюда проще запомнить смещение, чем гонять пересчет до конкретной строки, чтобы private_RecalculateLineCheckRangeY вернула нужное значение

    /// значение this.ParaMathRPI.bInternalRanges может изменить значение после того как будет вызвана данная функция

    var UpdateWrap = MATH_UPDWRAP_NOCHANGES;

    var LngR       = PRS.Ranges.length,
        Ranges     = PRS.Ranges;


    if(LngR > 0)
    {
        this.ParaMathRPI.IntervalState = MATH_INTERVAL_ON_SIDE;

        var  RY_NotWrap   = null;

        for(var Pos = 0; Pos < LngR; Pos++)
        {
            var WrapType = Ranges[Pos].typeLeft;

            if(WrapType !== WRAPPING_TYPE_SQUARE && WrapType !== WRAPPING_TYPE_THROUGH && WrapType !== WRAPPING_TYPE_TIGHT)
            {
                // выберем картинку с max RangeY c учетом данного условия, под которой попробуем расположить формулу
                if(RY_NotWrap == null || RY_NotWrap < Ranges[Pos].Y1)
                {
                    RY_NotWrap = Ranges[Pos].Y1;
                }

                this.ParaMathRPI.IntervalState = MATH_INTERVAL_EMPTY;
            }
        }

        if(this.ParaMathRPI.IntervalState == MATH_INTERVAL_ON_SIDE) // WrapType == WRAPPING_TYPE_SQUARE || WrapType == WRAPPING_TYPE_THROUGH || WrapType == WRAPPING_TYPE_TIGHT
        {
            // вычтем здесь Ind.Left для корректного сравнения (т.к.стартовый отрезок = граница Range + Ind.Left ), а также если XStart окажется левой границей (прибаится лишний Ind.Left)
            var XRange = this.ParaMathRPI.XRange - this.ParaMathRPI.IndLeft,
                XLimit = this.ParaMathRPI.XLimit;

            // рассчитываем XStart, XEnd
            var XStart = XRange,
                XEnd   = Ranges[0].X0;

            for(var Pos = 0; Pos < LngR - 1; Pos++)
            {
                if(XEnd - XStart < Ranges[Pos+1].X0 - Ranges[Pos].X1)
                {
                    XStart = Ranges[Pos].X1;
                    XEnd   = Ranges[Pos+1].X0;
                }
            }

            if(XEnd - XStart < XLimit - Ranges[LngR - 1].X1)
            {
                XStart = Ranges[LngR - 1].X1;
                XEnd   = XLimit;
            }

            // учтем Ind.Left
            // если впоследствии Ind.Left в word не будет учитываться, то нужно пересмотреть схему => в каких случаях и с какими параметрами рассчитыывать в Range

            XStart += this.ParaMathRPI.IndLeft;


            // в конце сравним с текущим отрезком, т.к. может произойти например след ситуация :
            // 2 плавающих объекта находятся в различных строках +> PRS.Ranges.length <=1
            // при этом формула должна расположится в макс по ширине из отрезков, образованными обоими плавающими мат объектами

            // учтем предыдущие отрезки:
            if(this.ParaMathRPI.XStart > XStart)
            {
                XStart = this.ParaMathRPI.XStart;
            }

            if(this.ParaMathRPI.XEnd < XEnd)
            {
                XEnd = this.ParaMathRPI.XEnd;
            }

            // рассчитываем RangeY

            var RangeY = Ranges[0].Y1;

            for(var Pos = 1; Pos < Ranges.length; Pos++)
            {
                if(Ranges[Pos].Y1 < RangeY)
                    RangeY = Ranges[Pos].Y1;
            }

            if(this.ParaMathRPI.RangeY == null || RangeY < this.ParaMathRPI.RangeY)
            {
                this.ParaMathRPI.RangeY = RangeY;
            }


            var DiffXStart = Math.abs(this.ParaMathRPI.XStart - XStart),
                DiffXEnd   = Math.abs(this.ParaMathRPI.XEnd - XEnd);

            if(DiffXStart > 0.001 || DiffXEnd > 0.001)
            {
                this.ParaMathRPI.XStart     = XStart;
                this.ParaMathRPI.XEnd       = XEnd;

                UpdateWrap = MATH_UPDWRAP_NEWRANGE;
            }
        }
        else
        {
            // если появился плавающий объект, относительно которого нельзя разместить формулу (в одном из Range, образованным плавающими объектами), то, соответсвенно, формула должна располагаться под плавающим объектом

            this.private_SetShiftY(PRS, RY_NotWrap);
            UpdateWrap = MATH_UPDWRAP_UNDERFLOW;
        }

    }

    return UpdateWrap;
};
ParaMath.prototype.private_RecalculateRangeInsideInterval = function(PRS, ParaPr, Depth)
{
    // var bInsideRange = PRS.X - 0.001 < this.ParaMathRPI.XStart && this.ParaMathRPI.XEnd < PRS.XEnd + 0.001;
    // наложим менее строгие условия попадания в отрезок
    var bNotInsideRange = this.ParaMathRPI.XStart > PRS.XEnd || this.ParaMathRPI.XEnd < PRS.X;
    var bNextRangeSide   = this.ParaMathRPI.IntervalState == MATH_INTERVAL_ON_SIDE && bNotInsideRange == true; // пересчитываем только в том отрезке, в котором находится формула

    // Номер  Range не влияет на UpdateWrapSettings, т.к. картинки могут располагаться одна под другой, и в одной ситуации это будет 0-ой Range,  в другой 1-ый

    if(bNextRangeSide) // при пересчете формулы между картинками/сбоку от картинки рассчитываем формулу в самом большом Range, остальные делаем пустыми
    {
        // переход к следующему Range
        this.Set_EmptyRange(PRS);
    }
    else
    {
        PRS.X = this.ParaMathRPI.XStart;
        PRS.XEnd = this.ParaMathRPI.XEnd;

        this.private_UpdateXLimits(PRS);

        this.private_RecalculateRoot(PRS, ParaPr, Depth);


        if(PRS.bMathWordLarge == true)
        {
            this.private_SetShiftY(PRS, this.ParaMathRPI.RangeY);
        }
        else if(PRS.NewRange == false && PRS.EmptyLine == true) // формула пересчиталась корректно, располагаем в данной строке => не разбивается на слова, выставим EmptyLine = false, чтобы не перенесли под картинку
        {
            PRS.EmptyLine = false;
        }

        PRS.SetMathRecalcInfoObject(this);

        if(PRS.NewRange == true && !(PRS.RecalcResult & recalcresult_ParaMath))
            PRS.ForceNewLine = true;
    }
};
ParaMath.prototype.private_RecalculateRangeWrap = function(PRS, ParaPr, Depth)
{
	var isInline = this.IsInline();

	// попадем сюда только, когда  либо нет плавающих объектов, привязанных к другому параграфу, нежели формула
	// либо когда не получилось расположить формулу в Range и формула пересчитывается обычным образом

	var PrevLineObject = PRS.GetMathRecalcInfoObject();
	if (PrevLineObject == null || PrevLineObject == this)
	{
		PRS.RecalcResult = recalcresult_NextLine;

		PRS.ResetMathRecalcInfo();

		// выставляем только для инлайновых формул => может случится так, что в одном параграфе окажутся несколько формул и для того, чтобы при первом пересчете пересчитались настройки нужно возвращать null
		// при последующих пересчетах PRS.MathRecalcInfo.Math будет выставлен null на ResetMathRecalcInfo в ф-ии private_RecalculatePage
		if (!isInline)
			PRS.SetMathRecalcInfoObject(this);
	}

	// здесь перебивается для неинлайновых формул и отступ первой строки и тот случай, когда формула не пересекает область расположения картинки (FlowBounds), но тем неменее пришли
	if (!isInline)
	{
		PRS.X    = this.ParaMathRPI.XStart;
		PRS.XEnd = this.ParaMathRPI.XEnd;
	}

	this.private_UpdateXLimits(PRS);

	var bStartRange      = this.Root.IsStartRange(PRS.Line, PRS.Range);
	var bNotBrPosInLWord = isInline && bStartRange == true && PRS.Ranges.length > 0 && PRS.Word == true;
	PRS.bBreakPosInLWord = bNotBrPosInLWord == false;  //не обновляем для инлайновой формулы, когда WordLarge (слово вышло за границы Range), перед формулой есть еще текст, чтобы не перебить LineBreakPos и выставить разбиение по тем меткам, которые были до пересчета формулы

	var bEmptyLine = PRS.EmptyLine;

	this.private_RecalculateRoot(PRS, ParaPr, Depth);

	var WrapState = this.PageInfo.Get_CurrentWrapState();

	this.PageInfo.Set_StateWordLarge(PRS.Line, PRS.bMathWordLarge);

	if (PRS.bMathWordLarge == true)
	{
		if (WrapState !== ALIGN_EMPTY)
		{
			this.private_SetRestartRecalcInfo(PRS);
			this.PageInfo.Set_NextWrapState();
		}
		else if (isInline && PRS.Ranges.length > 0)
		{
			if (PRS.bBreakPosInLWord == true) // когда для инлайновой формулы WordLarge (слово вышло за границы Range), перед формулой есть еще текст, чтобы не перебить LineBreakPos и выставить разбиение по тем меткам, которые были до пересчета формулы
			{
				PRS.EmptyLine = bEmptyLine; // вернем пред знач-е
				this.Root.Math_Set_EmptyRange(PRS.Line, PRS.Range);
				PRS.bMathWordLarge = false;
				PRS.NewRange       = true;
				PRS.MoveToLBP      = false;
			}
			else
			{
				//не обновляем для инлайновой формулы, когда WordLarge, перед формулой есть еще текст, чтобы не перебить LineBreakPos и выставить по тем меткам, которые были до формулы разбиение
				PRS.MoveToLBP = true;
			}
		}
	}
};
ParaMath.prototype.private_RecalculateRoot = function(PRS, ParaPr, Depth)
{
    var Para      = PRS.Paragraph;
    var ParaLine  = PRS.Line;
    var ParaRange = PRS.Range;

    // заглушка для пересчета Gaps элементов в текущей строке
    // если быстрый пересчет проверим нужно ли пересчитывать gaps у элементов текущей строки
    if(PRS.bFastRecalculate == true)
    {
        this.Root.Math_UpdateGaps(ParaLine, ParaRange);
    }

    this.Root.Recalculate_Range(PRS, ParaPr, Depth);

    if(PRS.NewRange == false)
    {
        // обнуляем GapRight для операторов
        PRS.OperGapRight       = 0;

        var WidthLine = PRS.X - PRS.XRange + PRS.SpaceLen + PRS.WordLen;

        var bFirstItem =  PRS.FirstItemOnLine == true && true === Para.IsSingleRangeOnLine(ParaLine, ParaRange);
        if(bFirstItem && PRS.X + PRS.SpaceLen + PRS.WordLen > PRS.XEnd)
        {
            PRS.bMathWordLarge = true;
        }

        this.UpdateWidthLine(PRS, WidthLine);
    }
};
ParaMath.prototype.private_SetRestartRecalcInfo = function(PRS)
{
	let absPage   = PRS.Paragraph.GetAbsolutePage(PRS.Page);
	let firstLine = this.PageInfo.Get_FirstLineOnPage(absPage);
	PRS.SetMathRecalcInfo(firstLine, this);
	PRS.RecalcResult = recalcresult_ParaMath;
	PRS.NewRange     = true;
};
ParaMath.prototype.Set_EmptyRange = function(PRS)
{
    // не выставляем PRS.EmptyLine = false, чтобы корректно не произошел перенос на след строку для ParaNewLine (PRS.EmptyLine == false && bInline == false)

    this.Root.Math_Set_EmptyRange(PRS.Line, PRS.Range);

    PRS.RecalcResult = recalcresult_NextLine;
    PRS.SetMathRecalcInfoObject(this);

    PRS.NewRange = true;
};
ParaMath.prototype.private_UpdateRangeY = function(PRS, RY)
{
	if (Math.abs(RY - PRS.Y) < 0.001)
		PRS.Y = RY + 1; // смещаемся по 1мм
	else
		PRS.Y = RY + AscCommon.TwipsToMM(1) + 0.001; // Добавляем 0.001, чтобы избавиться от погрешности

	PRS.NewRange = true;
};
ParaMath.prototype.private_SetShiftY = function(PRS, RY)
{
    this.PageInfo.Set_NeedUpdateWrap();
    this.ParaMathRPI.UpdateShiftY(RY);
    this.ParaMathRPI.Reset_WrapSettings();
    this.private_SetRestartRecalcInfo(PRS);
};
ParaMath.prototype.private_UpdateXLimits = function(PRS)
{
    var MathSettings = Get_WordDocumentDefaultMathSettings();
    var WrapState = this.PageInfo.Get_CurrentWrapState();

    var Page = this.Paragraph == null ? 0 : this.Paragraph.Get_AbsolutePage(PRS.Page);

    PRS.X    += MathSettings.Get_LeftMargin(WrapState);
    PRS.XEnd -= MathSettings.Get_RightMargin(WrapState);


    PRS.WrapIndent = MathSettings.Get_WrapIndent(WrapState);

    PRS.bFirstLine = this.Root.IsStartLine(PRS.Line);

    var AlignAt;
    var Jc = this.Get_Align();

    var alignBrk = this.Root.Get_AlignBrk(PRS.Line, this.Is_BrkBinBefore());

    if(alignBrk !== null)
    {
        if(alignBrk == 0 || Jc == align_Right || Jc == align_Center)
        {
            AlignAt = 0;
        }
        else
        {
            var PosAln  = alignBrk < this.DispositionOpers.length ?  alignBrk - 1 : this.DispositionOpers.length - 1;
            AlignAt = this.DispositionOpers[PosAln];
        }
    }

    this.PageInfo.Add_Line(PRS.Line, Page, AlignAt);

    PRS.X += this.PageInfo.Get_SpaceAlign(PRS.Line, Page);

    PRS.XRange = PRS.X;
};
ParaMath.prototype.Save_MathInfo = function(Copy)
{
    var RecalculateObject = new CMathRecalculateObject();

    var StructRecalc =
    {
        bFastRecalculate:   this.bFastRecalculate,
        PageInfo:           this.PageInfo,
        bInline:            this.ParaMathRPI.bInline,
        Align:              this.Get_Align(),
        bEmptyFirstRange:   this.Root.IsEmptyRange(this.Root.StartLine, this.Root.StartRange)
    };

    RecalculateObject.Fill(StructRecalc);

    return RecalculateObject;
};
ParaMath.prototype.Load_MathInfo = function(RecalculateObject)
{
    RecalculateObject.Load_MathInfo(this.PageInfo);
};
ParaMath.prototype.CompareMathInfo = function(RecalculateObject)
{
    return RecalculateObject.Compare(this.PageInfo);
};
ParaMath.prototype.Recalculate_Reset = function(CurRange, CurLine, RecalcResult)
{
    this.Root.Recalculate_Reset(CurRange, CurLine); // обновим StartLine и StartRange только для Root (в CParagraphContentWithContentBase), для внутренних элементов обновится на Recalculate_Range
};
ParaMath.prototype.Recalculate_Set_RangeEndPos = function(PRS, PRP, Depth)
{
    this.Root.Recalculate_Set_RangeEndPos(PRS, PRP, Depth);
};
ParaMath.prototype.Recalculate_LineMetrics = function(PRS, ParaPr, _CurLine, _CurRange)
{
    var ContentMetrics = new CMathBoundsMeasures();

    // обновляем LineAscent и LineDescent для пересчета инлайновой формулы с картинкой
    // если в формуле находится картинка, то может так получится, что в отрезках обтекания не будет ни одного элемента => PRS.Ascent и PRS.Descent равны 0
    // далее при вычилении отрезков (PRS.Ranges) для следующей строки  учитываются PRS.Ascent и PRS.Descent предыдщей строки, а они будут равны 0 , соответственно получим те же самые отрезки обтекания, что и в предыдущей строке
    // произойдет зацикливание

    this.Root.Recalculate_LineMetrics(PRS, ParaPr, _CurLine, _CurRange, ContentMetrics);

    var RootAscent  = this.Root.GetAscent(_CurLine, _CurRange),
        RootDescent = this.Root.GetDescent(_CurLine, _CurRange);


    if(PRS.LineAscent < RootAscent)
        PRS.LineAscent = RootAscent;

    if(PRS.LineDescent < RootDescent)
        PRS.LineDescent = RootDescent;

    this.Root.Math_UpdateLineMetrics(PRS, ParaPr);
};
ParaMath.prototype.Recalculate_Range_Width = function(PRSC, _CurLine, _CurRange)
{
    var SpaceLen   = PRSC.SpaceLen;

    var bBrkBefore = this.Is_BrkBinBefore();

    var bGapLeft  = bBrkBefore == true,
        bGapRight = bBrkBefore == false;

    this.Root.UpdateOperators(_CurLine, _CurRange, bGapLeft, bGapRight);

    this.Root.Recalculate_Range_Width(PRSC, _CurLine, _CurRange);

    PRSC.Range.W        += PRSC.SpaceLen - SpaceLen;
    PRSC.Range.SpaceLen = SpaceLen;

    PRSC.Words++;
};
ParaMath.prototype.UpdateWidthLine = function(PRS, Width)
{
	var PrevRecalcObject = PRS.GetMathRecalcInfoObject();

	if (PrevRecalcObject == null || PrevRecalcObject == this)
	{
		var W = Width - PRS.OperGapRight - PRS.OperGapLeft;
		this.PageInfo.UpdateWidth(PRS.Line, W);
	}
};
ParaMath.prototype.Recalculate_Range_Spaces = function(PRSA, _CurLine, _CurRange, _CurPage)
{
    // до пересчета Bounds для текущей строки ранее должны быть вызваны Recalculate_Range_Width (для ширины), Recalculate_LineMetrics(для высоты и аскента)

    // для инлайновой формулы не вызывается ф-ия setPosition, поэтому необходимо вызвать здесь
    // для неилайновой setPosition вызывается на Get_AlignToLine
    var PosInfo = new CMathPosInfo();

    PosInfo.CurLine  = _CurLine;
    PosInfo.CurRange = _CurRange;

    var pos   = new CMathPosition();
    this.Root.setPosition(pos, PosInfo);

    // страиницу для смещния параграфа относительно документа добавим на Get_Bounds, т.к. если формула находится в автофигуре, то для нее не прийдет Recalculate_Range_Spaces при перемещении автофигуры а другую страницу
    this.Root.UpdateBoundsPosInfo(PRSA, _CurLine, _CurRange, _CurPage);
    this.Root.Recalculate_Range_Spaces(PRSA, _CurLine, _CurRange, _CurPage);
};
ParaMath.prototype.Recalculate_PageEndInfo = function(PRSI, _CurLine, _CurRange)
{

};
ParaMath.prototype.RecalculateEndInfo = function(oPRSI)
{
};
ParaMath.prototype.SaveRecalculateObject = function(Copy)
{
	var RecalcObj = this.Root.SaveRecalculateObject(Copy);
	RecalcObj.Save_MathInfo(this, Copy);

	return RecalcObj;
};
ParaMath.prototype.LoadRecalculateObject = function(RecalcObj)
{
    RecalcObj.Load_MathInfo(this);
    this.Root.LoadRecalculateObject(RecalcObj);
};
ParaMath.prototype.PrepareRecalculateObject = function()
{
	this.Root.PrepareRecalculateObject();
};
ParaMath.prototype.IsEmptyRange = function(nCurLine, nCurRange)
{
	return this.Root.IsEmptyRange(nCurLine, nCurRange);
};
ParaMath.prototype.Check_Range_OnlyMath = function(Checker, CurRange, CurLine)
{
    if (null !== Checker.Math)
    {
        Checker.Math   = null;
        Checker.Result = false;
    }
    else
        Checker.Math = this;
};

ParaMath.prototype.Check_MathPara = function(Checker)
{
    Checker.Found  = true;
    Checker.Result = false;
};

ParaMath.prototype.Check_PageBreak = function()
{
    return false;
};

ParaMath.prototype.CheckSplitPageOnPageBreak = function(oPBChecker)
{
	return true;
};
ParaMath.prototype.Get_ParaPosByContentPos = function(ContentPos, Depth)
{
    return this.Root.Get_ParaPosByContentPos(ContentPos, Depth);
};
ParaMath.prototype.recalculateCursorPosition = function(positionCalculator, isCurrent)
{
	positionCalculator.handleParaMath(this, isCurrent);
};
ParaMath.prototype.Refresh_RecalcData = function(Data)
{
    this.Paragraph.Refresh_RecalcData2(0);
};

ParaMath.prototype.Refresh_RecalcData2 = function(Data)
{
    this.Paragraph.Refresh_RecalcData2(0);
};

ParaMath.prototype.RecalculateMinMaxContentWidth = function(MinMax)
{
    var RPI = new CRPI();
    RPI.MergeMathInfo(this.ParaMathRPI);

    this.Root.PreRecalc(null, this, new CMathArgSize(), RPI);
	this.Root.recalculateAllSize(AscCommon.g_oTextMeasurer);
    this.Root.RecalculateMinMaxContentWidth(MinMax);
};

ParaMath.prototype.Get_Range_VisibleWidth = function(RangeW, _CurLine, _CurRange)
{
    this.Root.Get_Range_VisibleWidth(RangeW, _CurLine, _CurRange);
};
ParaMath.prototype.Is_BrkBinBefore = function()
{
    var MathSettings = Get_WordDocumentDefaultMathSettings();

    return this.Is_Inline() ? false : MathSettings.Get_BrkBin() == BREAK_BEFORE;
};
ParaMath.prototype.Shift_Range = function(Dx, Dy, _CurLine, _CurRange, _CurPage)
{
    this.Root.Shift_Range(Dx, Dy, _CurLine, _CurRange, _CurPage);
};
//-----------------------------------------------------------------------------------
// Функция для работы с формулой
// в тч с  дефолтными текстовыми настройками и argSize
//-----------------------------------------------------------------------------------
ParaMath.prototype.Set_Inline = function(value)
{
    if(value !== this.ParaMathRPI.bInline)
    {
        this.ParaMathRPI.bChangeInline = true;
        this.ParaMathRPI.bInline = value;
        this.bFastRecalculate = false;          // после смены инлайновости, требуется полностью пересчитать формулу
    }
};
ParaMath.prototype.Get_Inline = function()
{
    return this.ParaMathRPI.bInline;
};
ParaMath.prototype.Is_Inline = function()
{
	return this.IsInline();
};
/**
 * Формула внутристрочная или нет
 * @returns {boolean}
 */
ParaMath.prototype.IsInline = function()
{
	return (this.ParaMathRPI.bInline === true);
};
ParaMath.prototype.ConvertToInlineMode = function()
{
	let oParagraph = this.GetParagraph();
	if (!oParagraph)
		return false;

	if (this.IsInlineMode())
		return true;
	
	let oParent = this.GetParent();
	let oInlineLevel, nPosInParent;
	if (oParent instanceof AscWord.CInlineLevelSdt && oParent.IsContentControlEquation())
	{
		oInlineLevel = oParent;
		oParent      = oParent.GetParent();
		nPosInParent = oInlineLevel.GetPosInParent(oParent);
	}
	else
	{
		nPosInParent = this.GetPosInParent(oParent);
	}
	
	if (!oParent || -1 === nPosInParent)
		return false;

	let oContentPos = this.GetStartPosInParagraph();
	let oRunElementsBefore = new CParagraphRunElements(oContentPos, 1, null, false);
	oRunElementsBefore.SetSaveContentPositions(true);
	oParagraph.GetPrevRunElements(oRunElementsBefore);
	let arrElements = oRunElementsBefore.GetElements();
	if (arrElements.length > 0 && arrElements[0].IsBreak())
		oParagraph.RemoveRunElement(oRunElementsBefore.GetContentPositions()[0]);

	oContentPos = this.GetEndPosInParagraph();
	let oRunElementsAfter = new CParagraphRunElements(oContentPos, 1, null, false);
	oRunElementsAfter.SetSaveContentPositions(true);
	oParagraph.GetNextRunElements(oRunElementsAfter);
	arrElements = oRunElementsAfter.GetElements();
	if (arrElements.length > 0 && arrElements[0].IsBreak())
		oParagraph.RemoveRunElement(oRunElementsAfter.GetContentPositions()[0]);

	let oAfterItem = oParagraph.GetNextRunElement(this.GetEndPosInParagraph());
	if (!oAfterItem || !oAfterItem.IsSpace())
	{
		let oRun = new ParaRun(oParagraph, false);
		oRun.Add(new AscWord.CRunSpace());
		oParent.AddToContent(nPosInParent + 1, oRun);
	}

	let oBeforeItem = oParagraph.GetPrevRunElement(this.GetStartPosInParagraph());
	if (oBeforeItem && oBeforeItem.IsText())
	{
		let oRun = new ParaRun(oParagraph, false);
		oRun.Add(new AscWord.CRunSpace());
		oParent.AddToContent(nPosInParent, oRun);
	}
	this.Set_Inline(true);

	return true;
};
ParaMath.prototype.ConvertToDisplayMode = function()
{
	let oParagraph = this.GetParagraph();
	if (!oParagraph)
		return false;

	if (!this.IsInlineMode())
		return true;
	
	let oParent = this.GetParent();
	let oInlineLevel, nPosInParent;
	if (oParent instanceof AscWord.CInlineLevelSdt && oParent.IsContentControlEquation())
	{
		oInlineLevel = oParent;
		oParent      = oParent.GetParent();
		nPosInParent = oInlineLevel.GetPosInParent(oParent);
	}
	else
	{
		nPosInParent = this.GetPosInParent(oParent);
	}
	
	if (!oParent || -1 === nPosInParent)
		return false;
	
	let oContentPos = this.GetStartPosInParagraph();

    let oRunElementsBefore = new CParagraphRunElements(oContentPos, 1, null, false);
	oRunElementsBefore.SetSaveContentPositions(true);
	oParagraph.GetPrevRunElements(oRunElementsBefore);

    let arrElements = oRunElementsBefore.GetElements();
	if (arrElements.length > 0 && arrElements[0].IsSpace())
		oParagraph.RemoveRunElement(oRunElementsBefore.GetContentPositions()[0]);

	oContentPos = this.GetEndPosInParagraph();
	let oRunElementsAfter = new CParagraphRunElements(oContentPos, 1, null, false);
	oRunElementsAfter.SetSaveContentPositions(true);
	oParagraph.GetNextRunElements(oRunElementsAfter);
	arrElements = oRunElementsAfter.GetElements();
	if (arrElements.length > 0 && arrElements[0].IsSpace())
		oParagraph.RemoveRunElement(oRunElementsAfter.GetContentPositions()[0]);

	let oAfterItem = oParagraph.GetNextRunElement(this.GetEndPosInParagraph());
	if (oAfterItem && !oAfterItem.IsParaEnd())
	{
		let oRun = new ParaRun(oParagraph, false);
		oRun.Add(new AscWord.CRunBreak(AscWord.break_Line));
		oParent.AddToContent(nPosInParent + 1, oRun);
	}

	let oBeforeItem = oParagraph.GetPrevRunElement(this.GetStartPosInParagraph());
	if (oBeforeItem || oParagraph.HaveNumbering())
	{
		let oRun = new ParaRun(oParagraph, false);
		oRun.Add(new AscWord.CRunBreak(AscWord.break_Line));
		oParent.AddToContent(nPosInParent, oRun);
	}
	this.Set_Inline(false);

	return true;
};
ParaMath.prototype.IsInlineMode = function()
{
	let paragraph = this.GetParagraph();
	if (!paragraph)
		return false;
	
	let paraPos = paragraph.GetPosByElement(this);
	if (!paraPos)
		return false;
	
	let startPos = this.GetStartPosInParagraph();
	let endPos   = this.GetEndPosInParagraph();
	
	let prev = paragraph.GetPrevRunElement(startPos);
	let next = paragraph.GetNextRunElement(endPos);
	
	return ((prev && !prev.IsBreak())
		|| (!prev && paragraph.HaveNumbering())
		|| (next && !next.IsBreak() && !next.IsParaEnd()));
};
ParaMath.prototype.NeedDispOperators = function(Line)
{
    return false === this.Is_Inline() &&  true == this.Root.IsStartLine(Line);
};
ParaMath.prototype.Get_Align = function()
{
	var Jc;
	if (this.ParaMathRPI.bInline)
	{
		var ParaPr = this.Paragraph.Get_CompiledPr2(false).ParaPr;
		Jc         = ParaPr.Jc;
	}
	else if (undefined !== this.Jc)
	{
		Jc = this.Jc;
	}
	else
	{
		var MathSettings = Get_WordDocumentDefaultMathSettings();
		Jc = MathSettings.Get_DefJc();
	}

	return Jc;
};
ParaMath.prototype.GetAlign = function()
{
	return this.Get_Align();
};
ParaMath.prototype.Set_Align = function(Align)
{
	if (this.Jc !== Align)
	{
		History.Add(new CChangesMathParaJc(this, this.Jc, Align));
		this.raw_SetAlign(Align);
	}
};
ParaMath.prototype.raw_SetAlign = function(Align)
{
    this.Jc = Align;
};
ParaMath.prototype.SetRecalcCtrPrp = function(Class)
{
    if(this.Root.Content.length > 0 && this.ParaMathRPI.bRecalcCtrPrp == false)
    {
        this.ParaMathRPI.bRecalcCtrPrp = this.Root.Content[0] == Class;
    }
};
ParaMath.prototype.MathToImageConverter = function(bCopy, _canvasInput, _widthPx, _heightPx, raster_koef)
{
    var bTurnOnId = false;
    if (false === g_oTableId.m_bTurnOff)
    {
        g_oTableId.m_bTurnOff = true;
        bTurnOnId = true;
    }

	History.TurnOff();

    var oldDefTabStop = AscCommonWord.Default_Tab_Stop;
	AscCommonWord.Default_Tab_Stop = 1;

    var oApi = Asc['editor'] || editor;
    if(!oApi || !oApi.textArtPreviewManager){

        History.TurnOn();
        if (true === bTurnOnId)
            g_oTableId.m_bTurnOff = false;
        return null;
    }
    var oShape = oApi.textArtPreviewManager.getShape();
    //var hdr = new CHeaderFooter(editor.WordControl.m_oLogicDocument.HdrFtr, editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, AscCommon.hdrftr_Header);
    var _dc = oShape.getDocContent();

    var par = _dc.Content[0];

    if (bCopy)
        par.Internal_Content_Add(0, this.Copy(false), false);
    else
        par.Internal_Content_Add(0, this, false);

   // _dc.Internal_Content_Add(0, par, false);

    par.Set_Align(align_Left);
    par.Set_Tabs(new CParaTabs());

    var _ind = new CParaInd();
    _ind.FirstLine = 0;
    _ind.Left = 0;
    _ind.Right = 0;
    par.Set_Ind(_ind, false);

    var _sp = new CParaSpacing();
    _sp.Line              = 1;
    _sp.LineRule          = Asc.linerule_Auto;
    _sp.Before            = 0;
    _sp.BeforeAutoSpacing = false;
    _sp.After             = 0;
    _sp.AfterAutoSpacing  = false;
    par.Set_Spacing(_sp, false);

    // берем Default настройки, т.к. функция MathToImageConverter вызывается для новой формулы (созданной при копировании)
    var XLimit = 210, YLimit = 10000;

    _dc.Reset(0, 0, XLimit, YLimit);
    _dc.Recalculate_Page(0, true);

    _dc.Reset(0, 0, par.Lines[0].Ranges[0].W + 0.001, 10000);
    _dc.Recalculate_Page(0, true);

	AscCommonWord.Default_Tab_Stop = oldDefTabStop;

    if (true === bTurnOnId)
        g_oTableId.m_bTurnOff = false;

    History.TurnOn();

    AscCommon.IsShapeToImageConverter = true;

    var dKoef = AscCommon.g_dKoef_mm_to_pix;

    var JointSize = this.Get_JointSize();

    var W = JointSize.W,
        H = JointSize.H;

    var w_mm = W;
    var h_mm = H;
    var w_px = (w_mm * dKoef) >> 0;
    var h_px = (h_mm * dKoef) >> 0;

    if (undefined !== raster_koef)
    {
        w_px *= raster_koef;
        h_px *= raster_koef;

        if (undefined !== _widthPx)
            _widthPx = (_widthPx * raster_koef) >> 0;
        if (undefined !== _heightPx)
            _heightPx = (_heightPx * raster_koef) >> 0;
    }

    var _canvas = null;
    if (window["NATIVE_EDITOR_ENJINE"] === true && window["IS_NATIVE_EDITOR"] !== true)
    {
        var _width  = undefined == _widthPx  ? w_px : _widthPx;
        var _height = undefined == _heightPx ? h_px : _heightPx;
        _canvas = new CNativeGraphics();
        _canvas.width  = _width;
        _canvas.height = _height;
        _canvas.create(window["native"], _width, _height, _width / dKoef, _height / dKoef);
        _canvas.CoordTransformOffset(_widthPx  !== undefined ? (_widthPx  - w_px) / 2 : 0,
                                     _heightPx !== undefined ? (_heightPx - h_px) / 2 : 0);
        _canvas.transform(1, 0, 0, 1, 0, 0);
        var bNeedSetParaMarks = false;
        if (AscCommon.isRealObject(editor) && editor.ShowParaMarks) 
        {
          bNeedSetParaMarks = true;
          editor.ShowParaMarks = false;
        }
        par.Draw(0, _canvas);
        if (bNeedSetParaMarks) 
        {
          editor.ShowParaMarks = true;
        }
    }
    else
    {
        var _canvas = (_canvasInput === undefined) ? document.createElement('canvas') : _canvasInput;

        _canvas.width   = (undefined == _widthPx) ? w_px : _widthPx;
        _canvas.height  = (undefined == _heightPx) ? h_px : _heightPx;

        var _ctx = _canvas.getContext('2d');

        var g = new AscCommon.CGraphics();
        g.init(_ctx, w_px, h_px, w_mm, h_mm);
        g.m_oFontManager = AscCommon.g_fontManager;

        g.m_oCoordTransform.tx = 0;
        g.m_oCoordTransform.ty = 0;

        if (_widthPx !== undefined && _heightPx !== undefined)
        {
            g.m_oCoordTransform.tx = (_widthPx - w_px) / 2;
            g.m_oCoordTransform.ty = (_heightPx - h_px) / 2;
        }

        g.transform(1,0,0,1,0,0);

        var bNeedSetParaMarks = false;
        if(AscCommon.isRealObject(editor) && editor.ShowParaMarks)
        {
            bNeedSetParaMarks = true;
            editor.ShowParaMarks = false;
        }

        par.Draw(0, g);

        if(bNeedSetParaMarks)
        {
            editor.ShowParaMarks = true;
        }
    }

    AscCommon.IsShapeToImageConverter = false;

    if (undefined === _canvasInput)
    {
        var _ret = { ImageNative: _canvas, ImageUrl: "", w_px: _canvas.width, h_px: _canvas.height, w_mm: w_mm, h_mm: h_mm};
        try
        {
            _ret.ImageUrl = _canvas.toDataURL("image/png");
        }
        catch (err)
        {
            _ret.ImageUrl = "";
        }
        if (_canvas.isNativeGraphics === true)
            _canvas.Destroy();
        return _ret;
    }
    return null;
};
ParaMath.prototype.Get_FirstTextPr = function()
{
    return this.Root.Get_FirstTextPr();
};
ParaMath.prototype.GetFirstRPrp = function()
{
    return this.Root.getFirstRPrp();
};
ParaMath.prototype.GetShiftCenter = function(oMeasure, font)
{
    oMeasure.SetFont(font);
    var metrics = oMeasure.Measure2Code(0x2217); // "+"

    return 0.6*metrics.Height;
};
ParaMath.prototype.GetPlh = function(oMeasure, font)
{
    oMeasure.SetFont(font);

    return oMeasure.Measure2Code(0x2B1A).Height;
};
ParaMath.prototype.Get_Default_TPrp = function()
{
    return this.DefaultTextPr;
};
//-----------------------------------------------------------------------------------
// Функции отрисовки
//-----------------------------------------------------------------------------------
ParaMath.prototype.Draw_HighLights = function(drawState)
{
	drawState.handleParaMath(this);
};
ParaMath.prototype.Draw_Elements = function(PDSE)
{
    /*PDSE.Graphics.p_color(255,0,0, 255);
     PDSE.Graphics.drawHorLine(0, PDSE.Y - this.Ascent, PDSE.X - 30, PDSE.X + this.Width + 30 , 1);*/

    var X = PDSE.X;

    this.Root.Draw_Elements(PDSE);

    PDSE.X = X + this.Root.GetWidth(PDSE.Line, PDSE.Range);
		if (PDSE.Graphics.m_bIsTextDrawer) {
			PDSE.Graphics.CheckSpaceDraw(null, true);
		}
    /*PDSE.Graphics.p_color(255,0,0, 255);
     PDSE.Graphics.drawHorLine(0, PDSE.Y - this.Ascent + this.Height, PDSE.X - 30, PDSE.X + this.Width + 30 , 1);*/
};
ParaMath.prototype.GetLinePosition = function(Line, Range)
{
    return this.Root.GetPos(Line, Range);
};
ParaMath.prototype.Draw_Lines = function(PDSL)
{
    if(false == this.Root.IsEmptyRange(PDSL.Line, PDSL.Range))
    {
        // Underline всей формулы
        var FirstRPrp = this.GetFirstRPrp();

        var Para = PDSL.Paragraph;

        var aUnderline  = PDSL.Underline;

        var X          = PDSL.X;
        var Y          = PDSL.Baseline;
        var UndOff     = PDSL.UnderlineOffset;
        var UnderlineY = Y + UndOff;
        var LineW      = (FirstRPrp.FontSize / 18) * g_dKoef_pt_to_mm;


        var BgColor = PDSL.BgColor;
        if ( undefined !== FirstRPrp.Shd && Asc.c_oAscShdNil !== FirstRPrp.Shd.Value )
            BgColor = FirstRPrp.Shd.Get_Color( Para );
        var AutoColor = ( undefined != BgColor && false === BgColor.Check_BlackAutoColor() ? new CDocumentColor( 255, 255, 255, false ) : new CDocumentColor( 0, 0, 0, false ) );
        var CurColor, RGBA, Theme = this.Paragraph.Get_Theme(), ColorMap = this.Paragraph.Get_ColorMap();

        // Выставляем цвет обводки
        if ( true === PDSL.VisitedHyperlink && ( undefined === FirstRPrp.Color && undefined === FirstRPrp.Unifill ) )
            CurColor = new CDocumentColor( 128, 0, 151 );
        else if ( true === FirstRPrp.Color.Auto && !FirstRPrp.Unifill)
            CurColor = new CDocumentColor( AutoColor.r, AutoColor.g, AutoColor.b );
        else
        {
            if(FirstRPrp.Unifill)
            {
                FirstRPrp.Unifill.check(Theme, ColorMap);
                RGBA = FirstRPrp.Unifill.getRGBAColor();
                CurColor = new CDocumentColor( RGBA.R, RGBA.G, RGBA.B );
            }
            else
            {
                CurColor = new CDocumentColor( FirstRPrp.Color.r, FirstRPrp.Color.g, FirstRPrp.Color.b );
            }
        }

        var Bound = this.Root.Get_LineBound(PDSL.Line, PDSL.Range),
            Width = Bound.W;

        if ( true === FirstRPrp.Underline ) {
					aUnderline.set(UnderlineY, LineW);
	        aUnderline.Add(X, X + Width, CurColor);
        }


        this.Root.Draw_Lines(PDSL);

        PDSL.X = Bound.X + Width;
    }
};

//-----------------------------------------------------------------------------------
// Функции для работы с курсором
//-----------------------------------------------------------------------------------
ParaMath.prototype.IsCursorPlaceable = function()
{
    return true;
};
ParaMath.prototype.IsCursorAtEnd = function()
{
	return this.Cursor_Is_End();
};
ParaMath.prototype.IsCursorAtBegin = function()
{
	return this.Cursor_Is_Start();
};
ParaMath.prototype.Cursor_Is_Start = function()
{
    // TODO: ParaMath.Cursor_Is_Start

    return this.Root.Cursor_Is_Start();
};

ParaMath.prototype.Cursor_Is_NeededCorrectPos = function()
{
    return false;
};

ParaMath.prototype.Cursor_Is_End = function()
{
    // TODO: ParaMath.Cursor_Is_End

    return this.Root.Cursor_Is_End();
};

ParaMath.prototype.MoveCursorToStartPos = function()
{
	this.Root.MoveCursorToStartPos();
};

ParaMath.prototype.MoveCursorToEndPos = function(SelectFromEnd)
{
	this.Root.MoveCursorToEndPos(SelectFromEnd);
};

ParaMath.prototype.getParagraphContentPosByXY = function(searchState)
{
	searchState.handleParaMath(this);
};

ParaMath.prototype.Get_ParaContentPos = function(bSelection, bStart, ContentPos, bUseCorrection)
{
    this.Root.Get_ParaContentPos(bSelection, bStart, ContentPos, bUseCorrection);
};

ParaMath.prototype.Set_ParaContentPos = function(ContentPos, Depth) // выставить логическую позицию в контенте
{
    this.Root.Set_ParaContentPos(ContentPos, Depth);
};

ParaMath.prototype.Get_PosByElement = function(Class, ContentPos, Depth, UseRange, Range, Line)
{
    if ( this === Class )
        return true;

    return this.Root.Get_PosByElement(Class, ContentPos, Depth, UseRange, Range, Line);
};

ParaMath.prototype.Get_ElementByPos = function(ContentPos, Depth)
{
    return this.Root.Get_ElementByPos(ContentPos, Depth);
};

ParaMath.prototype.Get_ClassesByPos = function(Classes, ContentPos, Depth)
{
    Classes.push(this);

    this.Root.Get_ClassesByPos(Classes, ContentPos, Depth);
};

ParaMath.prototype.GetPosByDrawing = function(Id, ContentPos, Depth)
{
    return false;
};

ParaMath.prototype.Get_RunElementByPos = function(ContentPos, Depth)
{
    return null;
};

ParaMath.prototype.Get_LastRunInRange = function(_CurLine, _CurRange)
{
    return this.Root.Get_LastRunInRange(_CurLine, _CurRange);
};

ParaMath.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
    return this.Root.Get_LeftPos(SearchPos, ContentPos, Depth, UseContentPos, false);
};

ParaMath.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
    return this.Root.Get_RightPos(SearchPos, ContentPos, Depth, UseContentPos, StepEnd, false);
};

ParaMath.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
    this.Root.Get_WordStartPos(SearchPos, ContentPos, Depth, UseContentPos, false);
};

ParaMath.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
    this.Root.Get_WordEndPos(SearchPos, ContentPos, Depth, UseContentPos, StepEnd, false);
};

ParaMath.prototype.Get_EndRangePos = function(_CurLine, _CurRange, SearchPos, Depth)
{
    return this.Root.Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth);
};

ParaMath.prototype.Get_StartRangePos = function(_CurLine, _CurRange, SearchPos, Depth)
{
    return this.Root.Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth);
};

ParaMath.prototype.Get_StartRangePos2 = function(_CurLine, _CurRange, ContentPos, Depth)
{
    return this.Root.Get_StartRangePos2(_CurLine, _CurRange, ContentPos, Depth);
};

ParaMath.prototype.Get_EndRangePos2 = function(_CurLine, _CurRange, ContentPos, Depth)
{
	return this.Root.Get_EndRangePos2(_CurLine, _CurRange, ContentPos, Depth);
};

ParaMath.prototype.Get_StartPos = function(ContentPos, Depth)
{
    this.Root.Get_StartPos(ContentPos, Depth);
};

ParaMath.prototype.Get_EndPos = function(BehindEnd, ContentPos, Depth)
{
    this.Root.Get_EndPos(BehindEnd, ContentPos, Depth);
};
//-----------------------------------------------------------------------------------
// Функции для работы с селектом
//-----------------------------------------------------------------------------------
ParaMath.prototype.Set_SelectionContentPos = function(StartContentPos, EndContentPos, Depth, StartFlag, EndFlag)
{
    this.Root.Set_SelectionContentPos(StartContentPos, EndContentPos, Depth, StartFlag, EndFlag);
    this.bSelectionUse = true;
};

ParaMath.prototype.IsSelectionUse = function()
{
    return this.bSelectionUse;
};

ParaMath.prototype.RemoveSelection = function()
{
    this.bSelectionUse = false;
    this.Root.RemoveSelection();
};

ParaMath.prototype.SelectAll = function(Direction)
{
    this.bSelectionUse = true;
    this.Root.SelectAll(Direction);
};

ParaMath.prototype.drawSelectionInRange = function(line, range, drawState)
{
    this.Root.drawSelectionInRange(line, range, drawState);
};

ParaMath.prototype.IsSelectionEmpty = function(CheckEnd)
{
    return this.Root.IsSelectionEmpty();
};

ParaMath.prototype.Selection_IsPlaceholder = function()
{
    var bPlaceholder = false;
    var result = this.GetSelectContent(),
        SelectContent = result.Content;
    var start = result.Start,
        end = result.End;

    if(start == end)
    {
        bPlaceholder = SelectContent.IsPlaceholder();
    }

    return bPlaceholder;
};
ParaMath.prototype.IsMathContentPlaceholder = function()
{
	return this.Selection_IsPlaceholder();
};
ParaMath.prototype.SelectAllInCurrentMathContent = function()
{
	let content = this.GetSelectContent().Content;
	if (content)
	{
		content.SelectThisElement(1);
		content.SelectAll(1);
	}
};

ParaMath.prototype.Selection_CheckParaEnd = function()
{
    return false;
};

ParaMath.prototype.Selection_CheckParaContentPos = function(ContentPos, Depth, bStart, bEnd)
{
    return this.Root.Selection_CheckParaContentPos(ContentPos, Depth, bStart, bEnd);
};

ParaMath.prototype.IsSelectedAll = function(Props)
{
	return this.Root.IsSelectedAll(Props);
};

ParaMath.prototype.SkipAnchorsAtSelectionStart = function(nDirection)
{
    return false;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции совместного редактирования
//----------------------------------------------------------------------------------------------------------------------
ParaMath.prototype.Write_ToBinary2 = function(Writer)
{
    Writer.WriteLong( AscDFH.historyitem_type_Math );

    // String : this.Id
    // Long   : this.Type
    // String : Root.Id

    Writer.WriteString2(this.Id);
    Writer.WriteLong(this.Type);
    Writer.WriteString2(this.Root.Id);
};

ParaMath.prototype.Read_FromBinary2 = function(Reader)
{
    // String : this.Id
    // Long   : this.Type
    // String : Root.Id

    this.Id   = Reader.GetString2();
    this.Type = Reader.GetLong();
    this.Root = g_oTableId.Get_ById(Reader.GetString2());
    this.Root.bRoot = true;
    this.Root.ParentElement = this;
};

ParaMath.prototype.Get_ContentSelection = function()
{
    var Bounds = null;
    var oContent = this.GetSelectContent().Content;

    if (true == oContent.Can_GetSelection())
    {
        if(oContent.bOneLine)
        {
            Bounds = [];
            Bounds[0] = [];

            var ContentBounds = oContent.Get_Bounds();
            var oBound = ContentBounds[0][0];

            Bounds[0][0] =
            {
                X:      oBound.X,
                Y:      oBound.Y,
                W:      oBound.W,
                H:      oBound.H,
                Page:   this.Paragraph.Get_AbsolutePage(oBound.Page)
            };
        }
        else
        {
            Bounds = this.private_GetBounds(oContent);
        }
    }

    return Bounds;
};

ParaMath.prototype.Recalc_RunsCompiledPr = function()
{
    this.Root.Recalc_RunsCompiledPr();
};

/**
 * Проверяем находимся ли мы во внутреннем (не самом верхнем) контенте формулы.
 */
ParaMath.prototype.Is_InInnerContent = function()
{
    var oContent = this.GetSelectContent().Content;

    if (oContent.bRoot)
        return false;

    return true;
};

/**
 * Обработка нажатия Enter внутри формулы
 */
ParaMath.prototype.Handle_AddNewLine = function()
{
    var ContentPos = new AscWord.CParagraphContentPos();

    var CurrContent = this.GetSelectContent().Content;

    if (true === CurrContent.bRoot)
        return false;

    CurrContent.Get_ParaContentPos(this.bSelectionUse, true, ContentPos);

    var NeedRecalculate = false;

    if(MATH_EQ_ARRAY === CurrContent.ParentElement.kind && this.Cursor_Is_End())
    {
        var oEqArray = CurrContent.ParentElement;
        oEqArray.Add_Row(oEqArray.CurPos + 1);
        oEqArray.CurPos++;

        NeedRecalculate = true;
    }
    //todo duplication EQ_ARRAY while creating new line 
    // else if(MATH_EQ_ARRAY === CurrContent.ParentElement.kind)
    // {

    // }
    else if(MATH_MATRIX !== CurrContent.ParentElement.kind)
    {
        var ctrPrp = CurrContent.Parent.CtrPrp.Copy();
        var props = {row: 2, ctrPrp: ctrPrp};
        var EqArray = new CEqArray(props);

        var FirstContent  = EqArray.getElementMathContent(0);
        var SecondContent = EqArray.getElementMathContent(1);

        CurrContent.SplitContent(SecondContent, ContentPos, 0);
        CurrContent.CopyTo(FirstContent, false);

        // вставим пустой Run в Content, чтобы не упала ф-ия Remove_FromContent
        // первый элемент всегда Run
        var Run = CurrContent.getElem(0);
        Run.Remove_FromContent(0, Run.Content.length, true);

        CurrContent.Remove_FromContent(1, CurrContent.Content.length);
        CurrContent.Add_ToContent(1, EqArray);
        CurrContent.Correct_Content(true);

        var CurrentContent = new AscWord.CParagraphContentPos();
        this.Get_ParaContentPos(false, false, CurrentContent);

        var RightContentPos = new CParagraphSearchPos();
        this.Get_RightPos(RightContentPos, CurrentContent, 0, true);
        this.Set_ParaContentPos(RightContentPos.Pos, 0);

        EqArray.CurPos = 1;
        SecondContent.MoveCursorToStartPos();

        NeedRecalculate = true;
    }
    //console.log(CurrContent.ParentElement)
    return NeedRecalculate;
};

/**
 * Разделение формулы на 2 части в заданной позиции. В текущем объекте остается левая часть формулы.
 * @param ContentPos Позиция
 * @param Depth
 * @returns Возвращается правая часть формулы.
 */
ParaMath.prototype.Split = function (ContentPos, Depth)
{
	if (this.Cursor_Is_End())
		return new ParaRun(this.Paragraph, false);

    var NewParaMath = new ParaMath();
    NewParaMath.Jc = this.Jc;

    this.Root.SplitContent(NewParaMath.Root, ContentPos, Depth);

    return NewParaMath;
};
ParaMath.prototype.SplitNoDuplicate = function(oContentPos, nDepth, oNewParagraph)
{
	var oNewElement = this.Split(oContentPos, nDepth);
	if (!oNewElement)
		return;

	oNewParagraph.AddToContent(oNewParagraph.Content.length, oNewElement, false);
};

/**
 * Пытаемся выполнить автозамену в формуле.
 * @returns {boolean} Выполнилась ли автозамена.
 */
ParaMath.prototype.Make_AutoCorrect = function()
{
    return false;
};

/**
 * Получаем рект формулы
 */
ParaMath.prototype.Get_Bounds = function()
{
    if (undefined === this.Paragraph || null === this.Paragraph)
    {
        return [ [{X : 0, Y : 0, W : 0, H : 0, Page : 0}] ];
    }
    else
    {
        return this.private_GetBounds(this.Root);
    }
};
ParaMath.prototype.GetBounds = function()
{
	return this.Get_Bounds();
};
ParaMath.prototype.Get_JointSize = function()
{
    var W = 0, H = 0;

    var _bounds = this.Get_Bounds();

    for(var Line = 0; Line < _bounds.length; Line++)
    {
        if(_bounds[Line] === undefined)
            break;

        for(var Range = 0; Range < _bounds[Line].length; Range++)
        {
            W += _bounds[Line][Range].W;
            H += _bounds[Line][Range].H;
        }
    }

    return {W: W, H: H};
};
ParaMath.prototype.private_GetBounds = function(Content)
{
    var Bounds = [ [{X : 0, Y : 0, W : 0, H : 0, Page : 0}] ];
    var ContentBounds = Content.Get_Bounds();

    for(var Line = 0; Line < ContentBounds.length; Line++)
    {
        Bounds[Line] = [];
        var CurLine = Line + Content.StartLine;
        var HLine = this.Paragraph.Lines[CurLine].Bottom - this.Paragraph.Lines[CurLine].Top;
        var Height = HLine;
        var Y;

        for(var Range = 0; Range < ContentBounds[Line].length; Range++)
        {
            var oBound   = ContentBounds[Line][Range],
                ParaPage = oBound.Page,
                YLine    = this.Paragraph.Pages[ParaPage].Y + this.Paragraph.Lines[CurLine].Top;

            Y = YLine;

            if(Content.bRoot == false)
            {
                if(HLine < oBound.H) // сделано для случаев, когда межстрочный интервал небольшой
                {
                    Height = HLine;
                    Y = YLine;
                }
                else
                {
                    Height = oBound.H;
                    Y  = oBound.Y;
                }
            }

            Bounds[Line][Range] =
            {
                X:      oBound.X,
                Y:      Y,
                W:      oBound.W,
                H:      Height,
                Page:   this.Paragraph.Get_AbsolutePage(oBound.Page)
            };
        }
    }

    return Bounds;
};

ParaMath.prototype.getPropsForWrite = function()
{
    return {Jc : this.Jc};
};

/**
 * Обновляем состояние интерфейса.
 */
ParaMath.prototype.Document_UpdateInterfaceState = function()
{
    var MathProps = this.Get_MenuProps();
    editor.sync_MathPropCallback(MathProps);
};

/**
 * Проверяем используется ли заданный MathContent на текущем уровне формулы
 * @param MathContent
 */
ParaMath.prototype.Is_ContentUse  = function(MathContent)
{
    if (this.Root === MathContent)
        return true;

    return false;
};

/*
 * Выполняем коректировку формулы после конвертирования ее из старой формулы (Equation 2-3).
 */
ParaMath.prototype.Correct_AfterConvertFromEquation = function()
{
    this.ParaMathRPI.bCorrect_ConvertFontSize = true;
};

ParaMath.prototype.CheckRevisionsChanges = function(Checker, ContentPos, Depth)
{
    return this.Root.CheckRevisionsChanges(Checker, ContentPos, Depth);
};
ParaMath.prototype.AcceptRevisionChanges = function(Type, bAll)
{
    return this.Root.AcceptRevisionChanges(Type, bAll);
};
ParaMath.prototype.RejectRevisionChanges = function(Type, bAll)
{
    return this.Root.RejectRevisionChanges(Type, bAll);
};
ParaMath.prototype.SetReviewType = function(ReviewType, RemovePrChange)
{
    return this.Root.SetReviewType(ReviewType, RemovePrChange);
};
ParaMath.prototype.HandleTab = function(isForward)
{
    if(this.ParaMathRPI.bInline == false)
    {
        var CountOperators = this.DispositionOpers.length;
        var bBrkBefore = this.Is_BrkBinBefore();
        this.Root.Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
    }
};
ParaMath.prototype.SetContentSelection = function(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag)
{
    return this.Root.SetContentSelection(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag);
};
ParaMath.prototype.SetContentPosition = function(DocPos, Depth, Flag)
{
    return this.Root.SetContentPosition(DocPos, Depth, Flag);
};
ParaMath.prototype.IsStopCursorOnEntryExit = function()
{
	return true;
};
ParaMath.prototype.RemoveTabsForTOC = function(isTab)
{
	return isTab;
};
/**
 * Проверяем, что данный класс лежит в плейсхолдере специального контейнера для формул
 * @return {boolean}
 */
ParaMath.prototype.IsParentEquationPlaceholder = function()
{
	var arrContentControl = this.GetParentContentControls();
	for (var nIndex = 0, nCount = arrContentControl.length; nIndex < nCount; ++nIndex)
	{
		if (arrContentControl[nIndex].IsContentControlEquation() && arrContentControl[nIndex].IsPlaceHolder())
			return true;
	}

	return false;
};
ParaMath.prototype.CalculateTextToTable = function(oEngine)
{
	this.Root.CalculateTextToTable(oEngine);
};
ParaMath.prototype.ConvertFromLaTeX = function()
{
	let oLaTeX = this.GetTextOfElement(true, true);
	this.Root.Remove_Content(0, this.Root.Content.length);
	this.Root.CurPos = 0;
	AscMath.ConvertLaTeXToTokensList(oLaTeX, this.Root);
	this.Root.Correct_Content(true);
    this.Root.CurPos++;
};
ParaMath.prototype.ConvertToLaTeX = function()
{
	let oLaTeXContent = this.GetTextOfElement(true);
	this.Root.Remove_Content(0,this.Root.Content.length);
	this.Root.AddDataFromFlatMathTextAndStyles(oLaTeXContent.Flat());
};
ParaMath.prototype.ConvertFromUnicodeMath = function()
{
	let oUnicode = this.GetTextOfElement(false);
	this.Root.Remove_Content(0, this.Root.Content.length);
	this.Root.CurPos = 0;
	AscMath.CUnicodeConverter(oUnicode, this.Root);
	this.Root.Correct_Content(true);
	this.Root.CurPos++;
};
ParaMath.prototype.ConvertToUnicodeMath = function()
{
	let oUnicodeContent = this.GetTextOfElement(false);
	this.Root.Remove_Content(0,this.Root.Content.length);
	this.Root.AddDataFromFlatMathTextAndStyles(oUnicodeContent.Flat());
	this.Paragraph.updateTrackRevisions();
};
ParaMath.prototype.ConvertView = function(isToLinear, nInputType)
{
	AscCommon.executeNoRevisions(this._convertView, this.GetLogicDocument(), this, arguments);
};
ParaMath.prototype._convertView = function(isToLinear, nInputType)
{
	if (undefined === nInputType)
	{
		let oApi = Asc.editor || editor;
		nInputType = oApi ? oApi.getMathInputType() : Asc.c_oAscMathInputType.Unicode;
	}

	if (this.IsEmpty())
		return;

	if (isToLinear)
	{
		if (Asc.c_oAscMathInputType.Unicode === nInputType)
			this.ConvertToUnicodeMath();
		else if (Asc.c_oAscMathInputType.LaTeX === nInputType)
			this.ConvertToLaTeX();
	}
	else
	{
		if (Asc.c_oAscMathInputType.Unicode === nInputType)
		{
			this.ConvertFromUnicodeMath();
		}
		else if (Asc.c_oAscMathInputType.LaTeX === nInputType)
		{
			this.ConvertFromLaTeX();
		}
	}
};
ParaMath.prototype.SplitSelectedContent = function()
{
    var oSelection = this.GetSelectContent();
    var oContent = oSelection.Content;
    oContent.SplitSelectedContent();
};
ParaMath.prototype.ConvertViewBySelection = function(isToLinear, nInputType)
{
	AscCommon.executeNoRevisions(this._convertViewBySelection, this.GetLogicDocument(), this, arguments);
};
ParaMath.prototype._convertViewBySelection = function(isToLinear, nInputType)
{
    this.SplitSelectedContent();

    var oSelection = this.GetSelectContent();

    oSelection.Content.ConvertContentView(
        oSelection.Start,
        oSelection.End,
        nInputType,
        isToLinear
    );
};
ParaMath.prototype.CheckSpelling = function(oCollector, nDepth)
{
	if (oCollector.IsExceedLimit())
		return;

	oCollector.FlushWord();
};
//----------------------------------------------------------------------------------------------------------------------
// Search
//----------------------------------------------------------------------------------------------------------------------
ParaMath.prototype.Search = function(oParaSearch)
{
	this.Root.Search(oParaSearch);
};
ParaMath.prototype.AddSearchResult = function(SearchResult, Start, ContentPos, Depth)
{
	this.Root.AddSearchResult(SearchResult, Start, ContentPos, Depth);
};
ParaMath.prototype.ClearSearchResults = function()
{
	this.Root.ClearSearchResults();
};
ParaMath.prototype.RemoveSearchResult = function(oSearchResult)
{
	this.Root.RemoveSearchResult(oSearchResult);
};
ParaMath.prototype.GetSearchElementId = function(bNext, bUseContentPos, ContentPos, Depth)
{
	return this.Root.GetSearchElementId(bNext, bUseContentPos, ContentPos, Depth);
};
//----------------------------------------------------------------------------------------------------------------------
ParaMath.prototype.IsContentControlEquation = function()
{
	let parent = this.GetParent();
	return (parent
		&& parent instanceof AscWord.CInlineLevelSdt
		&& parent.IsContentControlEquation()
		&& parent.IsPlaceHolder());
};
ParaMath.prototype.ProcessingOldEquationConvert = function()
{
	this.Root.ProcessingOldEquationConvert();
};
ParaMath.fromLatex = function(latex, textPr)
{
	let paraMath = new ParaMath();

	latex = latex.replaceAll('&amp;', '&');
	latex = latex.replaceAll('&lt;', '<');
	latex = latex.replaceAll('&gt;', '>');
	
	let run = new AscWord.Run(null, true);
	run.AddText(latex);
	
	paraMath.Root.addElementToContent(run);
	paraMath.Root.Correct_Content();
	paraMath.SetParagraph(null);
	
	paraMath.ConvertView(false, Asc.c_oAscMathInputType.LaTeX);
	
	if (textPr)
	{
		textPr.RFonts.SetAll("Cambria Math");
		paraMath.ApplyTextPr(textPr, undefined, true);
	}
	
	return paraMath;
};
ParaMath.fromMathML = function(xml, textPr)
{
	let paraMath = new ParaMath();

	this.mathMLData = {
		'mo-linebreak-todo': [],
		'mo-id': {}
	};
	
	let reader = new StaxParser(xml);
	
	if (!reader.ReadNextNode() || "math" !== reader.GetNameNoNS())
	{
		paraMath.Root.Correct_Content();
		return paraMath;
	}
	
	paraMath.Root.fromMathML(reader);
	paraMath.Root.Correct_Content();
	paraMath.SetParagraph(null);
	
	if (textPr)
	{
		textPr.RFonts.SetAll("Cambria Math");
		paraMath.ApplyTextPr(textPr, undefined, true);
	}
	
	return paraMath;
};
ParaMath.readMathMLNode = function(reader)
{
	let elements = [];
	let name = reader.GetNameNoNS();
	switch (name)
	{
		case 'ci':
		case 'csymbol':
		case 'mi':
		case 'mo':
		case 'mn':
		case 'mspace':
		case 'mtext':
			function decodeHexEntities(str) {
				var result = '';
				for (var i = 0; i < str.length; ) {
					if (
						str.charAt(i) === '&' &&
						str.charAt(i + 1) === '#' &&
						str.charAt(i + 2).toLowerCase() === 'x'
					) {
						var j = i + 3;
						var hex = '';
						while (j < str.length && str.charAt(j) !== ';') {
							hex += str.charAt(j);
							j++;
						}
						if (str.charAt(j) === ';') {
							var code = parseInt(hex, 16);
							if (!isNaN(code)) {
								result += String.fromCharCode(code);
								i = j + 1;
								continue;
							}
						}
					}

					result += str.charAt(i);
					i++;
				}

				return result;
			}
			
			function proceedAttributes(name, text, attributes)
			{
				const GetMathFontChar = AscMath.GetMathFontChar;
				switch (name)
				{
					case 'mi':
					{
						let type = -1;
						switch (attributes['mathvariant'])
						{
							case 'bold':					type = 0; break;
							case 'italic':					type = 1; break;
							case 'bold-italic':				type = 2; break;
							case 'double-struck':			type = 12; break;
							case 'bold-fraktur':			type = 10; break;
							case 'script':					type = 7; break;
							case 'bold-script':				type = 8; break;
							case 'fraktur':					type = 9; break;
							case 'sans-serif':				type = 3; break;
							case 'bold-sans-serif':			type = 4; break;
							case 'sans-serif-italic':		type = 5; break;
							case 'sans-serif-bold-italic':	type = 6; break;
							case 'monospace':				type = 11; break;
							default:						type = -1; break;
						}
						
						// single character must be italic
						if (text.length === 1 && type === -1)
							type = 1;
						
						let convertedText = "";
						for (let oIter = text.getUnicodeIterator(); oIter.check(); oIter.next())
						{
							let currentChar = String.fromCodePoint(oIter.value());
							
							if (GetMathFontChar[currentChar] && GetMathFontChar[currentChar][type])
								convertedText += GetMathFontChar[currentChar][type];
							else
								convertedText += currentChar;
						}
						text = convertedText;
						break;
					}
					case 'mo':
					{
						if (attributes['id'])
							this.mathMLData['mo-id'][attributes['id']] = elements[0];
						
						if (name === 'mo' && attributes['linebreak'])
						{
							this.mathMLData['mo-linebreak-todo'].push({
								element: elements[0],
								type: attributes['linebreak'],
								indentalign: attributes['indentalign'] === "id",
								indenttarget: attributes['indenttarget']
							});
						}
						
						// indent
						
						break;
					}
					case 'mspace':
					{
						// for now skip spaces
						text = " ";
						break;
					}
				}
				return text;
			}
			let attributes = reader.GetAttributes();

			let text = reader.GetText();
			text = text.trim();
			text = text.replaceAll(String.fromCharCode(8290), ""); // invisible *
			text = text.replaceAll(String.fromCharCode(8292), ""); // invisible +
			text = text.replaceAll("\n", "");
			text = text.replaceAll("\r", "");
			text = decodeHexEntities(text);
			
			text = proceedAttributes(name, text, attributes);
			if (text)
			{
				elements.push(new AscWord.Run(null, true));
				elements[0].AddText(text);
			}
			
			
			break;
		case 'merror':
		case 'menclose':
			elements.push(new AscMath.BorderBox.fromMathML(reader));
			break;
		case 'mphantom':
			elements.push(new AscMath.Phantom.fromMathML(reader));
			break;
		case 'msub':
			elements.push(new AscMath.Degree.fromMathML(reader, DEGREE_SUBSCRIPT));
			break;
		case 'msup':
			elements.push(new AscMath.Degree.fromMathML(reader, DEGREE_SUPERSCRIPT));
			break;
		case 'msubsup':
		case 'munderover':
			elements.push(new AscMath.DegreeSubSup.fromMathML(reader, DEGREE_SubSup));
			break;
		case 'mfrac':
			elements.push(AscMath.Fraction.fromMathML(reader));
			break;
		case 'msqrt':
			elements.push(AscMath.Radical.fromMathML(reader, false));
			break;
		case 'mroot':
			elements.push(AscMath.Radical.fromMathML(reader, true));
			break;
		case 'mpadded':
		case 'mstyle':
		case 'mrow':
		case 'semantics':
			elements = AscWord.ParaMath.readMathMLMRow(reader);
			break;
		case 'munder':
		{
			let attributes = reader.GetAttributes();

			if (attributes['accentunder'] && attributes['accentunder'] === 'false')
				elements.push(AscMath.GroupCharacter.fromMathML(reader, VJUST_BOT));
			else
				elements.push(AscMath.Accent.fromMathML(reader, VJUST_TOP));

			break;
		}
		case 'mover':
		{
			let attributes = reader.GetAttributes();

			if (attributes['accent'] && attributes['accent'] === 'false')
				elements.push(AscMath.GroupCharacter.fromMathML(reader, VJUST_TOP));
			else
				elements.push(AscMath.Accent.fromMathML(reader, VJUST_BOT));
			break;
		}
		case 'mtable':
			elements.push(AscMath.Matrix.fromMathML(reader));
			break;
		// case 'mstack':
		// 	elements.push(AscMath.EqArray.fromMathML(reader, true));
		// 	break;
		case 'mtr':
			elements.push(AscMath.Matrix.fromMathML_mtr(reader));
			break;
		case 'mtd':
			elements.push(AscMath.Matrix.fromMathML_mtd(reader))
			break;
		case 'mfenced':
			elements.push(AscMath.Delimiter.fromMathML(reader));
			break;
		// case 'apply':
		// 	elements = AscWord.ParaMath.proceedApply(reader);
		// 	break;
	}
	
	return elements;
};
ParaMath.proceedApply = function (reader)
{
	let operands = [];
	let depth = reader.GetDepth();
	while (reader.ReadNextSiblingNode(depth))
	{
		let node = AscWord.ParaMath.readMathMLNode(reader);

		if (node.length > 1)
		{
			let box = new CBox({});
			let content = box.getBase();

			for (let i = 0; i < node.length; i++)
			{
				content.addElementToContent(node[i]);
			}

			operands.push(box)
		}
		else
			operands.push(node[0]);
	}

	let operator = operands.shift();
	let strOperator = (operator instanceof CBox) ? "" : operator.GetTextOfElement().GetText();
	switch (strOperator)
	{
		case 'plus':
		case 'minus':
		case 'times':
		{
			let mathContent = [];
			for (let i = 0; i < operands.length; i++)
			{
				mathContent.push(operands[i]);
				if (i !== operands.length - 1)
				{
					let run = new AscWord.Run(null, true);

					switch (strOperator)
					{
						case 'plus': run.AddText('+'); break;
						case 'minus': run.AddText('-'); break;
						case 'times': run.AddText('*'); break;
					}

					mathContent.push(run);
				}

			}

			return mathContent;
			break;
		}
	}
}
ParaMath.readMathMLMRow = function(reader)
{
	let result = []
	let depth = reader.GetDepth();
	while (reader.ReadNextSiblingNode(depth))
	{
		result = result.concat(AscWord.ParaMath.readMathMLNode(reader));
	}
	return result;
};
ParaMath.checkAfterAddContent = function()
{
};
ParaMath.checkLinebreak = function (element)
{
	for (let i = 0; i < this.mathMLData['mo-linebreak-todo'].length; i++)
	{
		let current = this.mathMLData['mo-linebreak-todo'][i];
		let el = current.element;
		if (element === el)
		{
			this.mathMLData['mo-linebreak-todo'].splice(i, 1);
			let element =  this.mathMLData['mo-id'][current.indenttarget];
			return element.private_GetPosInParent();
			// todo -> to operators
		}
	}
	return null;
};

ParaMath.readMathMLContent = function(reader)
{
	let mathContent = new CMathContent();
	let elements = AscWord.ParaMath.readMathMLNode(reader);
	
	for (let i = 0; i < elements.length; ++i)
	{
		let current = elements[i];
		mathContent.addElementToContent(current);

		let breakPos = this.checkLinebreak(current)
		if (breakPos !== null)
		{
			current.Set_MathForcedBreak(true, undefined);
		}
	}
	
	mathContent.Correct_Content(true);
	this.checkAfterAddContent();
	return mathContent;
};
ParaMath.readMathMLContentOnLevel = function (reader)
{
	let content = [];
	let mathContent = new CMathContent();
	let depth = reader.GetDepth();

	while (reader.ReadNextSiblingNode(depth))
	{
		content = content.concat(AscWord.ParaMath.readMathMLNode(reader));
	}

	for (let i = 0; i < content.length; i++)
	{
		mathContent.addElementToContent(content[i]);
	}

	mathContent.Correct_Content(true);
	return mathContent;
}

function MatGetKoeffArgSize(FontSize, ArgSize)
{
    var FontKoef = 1;

    if(ArgSize == -1)
    {
        FontKoef = g_dMathArgSizeKoeff_1;
    }
    else if(ArgSize == -2)
    {
        FontKoef = g_dMathArgSizeKoeff_2;
    }

    if (1 !== FontKoef )
    {
        FontKoef = (((FontSize * FontKoef * 2 + 0.5) | 0) / 2) / FontSize;
    }

    return FontKoef;
};

function CMathRecalculateInfo()
{
    this.bInline                       = false;
    this.bChangeInline                 = true;
    this.bRecalcCtrPrp                 = false; // необходимо для пересчета CtrPrp (когда изменились текстовые настройки у первого элемнента, ctrPrp нужно пересчитать заново для всей формулы)
    this.bCorrect_ConvertFontSize      = false;

    this.IntervalState                  = MATH_INTERVAL_EMPTY;
    this.XStart                         = 0;
    this.XEnd                           = 0;
    this.XRange                         = 0;
    this.XLimit                         = 0;
    this.IndLeft                        = 0;
    this.bInternalRanges                = false;

    this.RangeY                         = null; // max среди нижних границ плавающих объектов
    this.ShiftY                         = 0;
}
CMathRecalculateInfo.prototype.Reset = function(PRS, ParaPr)
{
    this.XRange                 = PRS.XStart + ParaPr.Ind.Left;
    this.XLimit                 = PRS.XLimit;
    this.IndLeft                = ParaPr.Ind.Left;
    this.ShiftY                 = 0;

    this.Reset_WrapSettings();
};
CMathRecalculateInfo.prototype.Reset_WrapSettings = function()
{
    this.RangeY                 = null;

    this.bInternalRanges        = false;
    this.IntervalState          = MATH_INTERVAL_EMPTY;

    this.XStart                 = this.XRange;
    this.XEnd                   = this.XLimit;
};
CMathRecalculateInfo.prototype.UpdateShiftY = function(RY)
{
    this.ShiftY = RY;
};
CMathRecalculateInfo.prototype.ClearRecalculate = function()
{
    this.bChangeInline                  = false;
    this.bRecalcCtrPrp                  = false;
    this.bCorrect_ConvertFontSize       = false;
};

function CMathRecalculateObject()
{
    this.WrapState        = ALIGN_EMPTY;
    this.MaxW             = 0;              // для рассчета выравнивания формулы нужно учитывать изменилась ли максимальная ширина или нет
    this.bWordLarge       = false;          // если формула выходит за границы докумена, то не нужно учитывать выравнивание, а значит можно сделать быстрый пересчет
    this.bFastRecalculate = true;           /*если добавляем буквы во внутренний контент, который не бьется на строки, то отменяем быстрый пересчет,
                                            т.к. высота контента может поменяться (она рассчитывается точно исходя из размеров внутр элементов)*/

    this.bInline          = false;
    this.Align            = align_Justify;
    this.bEmptyFirstRange = false;
}
CMathRecalculateObject.prototype.Fill = function(StructRecalc)
{
    this.bFastRecalculate = StructRecalc.bFastRecalculate;
    this.bInline          = StructRecalc.bInline;
    this.Align            = StructRecalc.Align;

    var PageInfo          = StructRecalc.PageInfo;

    this.WrapState        = PageInfo.Get_CurrentWrapState();
    this.MaxW             = PageInfo.Get_MaxWidthOnCurrentPage();
    this.bWordLarge       = PageInfo.Get_CurrentStateWordLarge();

    this.bEmptyFirstRange = StructRecalc.bEmptyFirstRange;
};
CMathRecalculateObject.prototype.Load_MathInfo = function(PageInfo)
{
    PageInfo.Set_CurrentWrapState(this.WrapState);

    // текущая MaxW и MaxW в PageInfo это не одно и то же
    //PageInfo.SetCurrentMaxWidth(this.MaxW);
};
CMathRecalculateObject.prototype.Compare = function(PageInfo)
{
    var result = true;

    if(this.bFastRecalculate == false)
        result = false;

    if(this.WrapState !== PageInfo.Get_CurrentWrapState())
        result = false;

    if(this.bEmptyFirstRange !== PageInfo.bEmptyFirstRange)
        result = false;

    var DiffMaxW = this.MaxW - PageInfo.Get_MaxWidthOnCurrentPage();

    if(DiffMaxW < 0)
        DiffMaxW = -DiffMaxW;

    var LargeComposition = this.bWordLarge == true && true == PageInfo.Get_CurrentStateWordLarge();

    if(LargeComposition == false && this.bInline == false && this.Align == align_Justify && DiffMaxW > 0.001)
        result = false;

    return result;
};

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].MathMenu = MathMenu;
window['AscCommonWord'].ParaMath = ParaMath;
window['AscWord'].ParaMath = ParaMath;
