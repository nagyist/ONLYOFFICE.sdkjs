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

window['AscWord'] = window['AscWord'] || {};


var align_Left = AscCommon.align_Left;
var align_Right = AscCommon.align_Right;
var History = AscCommon.History;

var linerule_Auto = Asc.linerule_Auto;
var c_oAscShdClear = Asc.c_oAscShdClear;
var c_oAscShdNil = Asc.c_oAscShdNil;
var c_oAscDropCap = Asc.c_oAscDropCap;

var EvenAndOddHeaders = false;

var Default_Tab_Stop = 12.5;

var Default_Heading_Font = "Arial";
var Default_Font         = "Arial";

var highlight_None = -1;

var smallcaps_Koef = 0.8; // Коэффициент изменения размера шрифта для малых прописных букв

var smallcaps_and_script_koef = AscCommon.vaKSize * smallcaps_Koef; // суммарный коэффициент, когда текст одновременно и в индексе, и написан малыми прописными

var g_dKoef_pt_to_mm = 25.4 / 72;
var g_dKoef_pc_to_mm = g_dKoef_pt_to_mm / 12;
var g_dKoef_in_to_mm = 25.4;
var g_dKoef_twips_to_mm = g_dKoef_pt_to_mm / 20;
var g_dKoef_mm_to_pt = 1 / g_dKoef_pt_to_mm;
var g_dKoef_mm_to_twips = 1 / g_dKoef_twips_to_mm;
var g_dKoef_mm_to_emu = 36000;
var g_dKoef_emu_to_mm = 1 / 36000;
var g_dKoef_emu_to_twips = g_dKoef_emu_to_mm * g_dKoef_mm_to_twips;
var g_dKoef_pt_to_twips = 20;
var g_dKoef_twips_to_emu = 1 / g_dKoef_emu_to_twips;
var g_dKoef_twips_to_pt = 1 / g_dKoef_pt_to_twips;
var g_dKoef_em_to_mm = 4.21752;

var tblwidth_Auto = 0x00;
var tblwidth_Mm   = 0x01;
var tblwidth_Nil  = 0x02;
var tblwidth_Pct  = 0x03;

var tbllayout_Fixed   = 0x00;
var tbllayout_AutoFit = 0x01;

var vertalignjc_Top    = 0x00;
var vertalignjc_Center = 0x01;
var vertalignjc_Bottom = 0x02;

var vmerge_Restart  = 0x01;
var vmerge_Continue = 0x02;

var spacing_Auto = -1;

var styletype_Paragraph = 0x01;
var styletype_Numbering = 0x02;
var styletype_Table     = 0x03;
var styletype_Character = 0x04;

var textdirection_LRTB  = 0x00;
var textdirection_TBRL  = 0x01;
var textdirection_BTLR  = 0x02;
var textdirection_LRTBV = 0x03;
var textdirection_TBRLV = 0x04;
var textdirection_TBLRV = 0x05;

function private_GetWordLogicDocument()
{
	if(editor)
	{
		let oLogicDocument = null;
		oLogicDocument = editor.private_GetLogicDocument && editor.private_GetLogicDocument();
		if(oLogicDocument && oLogicDocument.IsDocumentEditor())
		{
			return oLogicDocument;
		}
	}
	return null;
}

function IsEqualStyleObjects(Object1, Object2)
{
    if (undefined === Object1 && undefined === Object2)
        return true;

    if (undefined === Object1 || undefined === Object2)
        return false;

    return Object1.Is_Equal(Object2);
}

function IsEqualNullableFloatNumbers(nNum1, nNum2)
{
	if (undefined === nNum1 && undefined === nNum2)
		return true;

	if (undefined === nNum1 || undefined === nNum2)
		return false;

	return Math.abs(nNum1 - nNum2) < 0.001;
}

function CheckUndefinedToNull(isConvert, Value)
{
	return (true === isConvert && undefined === Value ? null : Value);
}

function CTableStylePr()
{
    this.TextPr      = new CTextPr();
    this.ParaPr      = new CParaPr();
    this.TablePr     = new CTablePr();
    this.TableRowPr  = new CTableRowPr();
    this.TableCellPr = new CTableCellPr();
}

CTableStylePr.prototype =
{
    Merge : function(TableStylePr)
    {
        this.TextPr.Merge( TableStylePr.TextPr );
        this.ParaPr.Merge( TableStylePr.ParaPr );
        this.TablePr.Merge( TableStylePr.TablePr );
        this.TableRowPr.Merge( TableStylePr.TableRowPr );
        this.TableCellPr.Merge( TableStylePr.TableCellPr );
    },

    Copy : function(oPr)
    {
        var TableStylePr = new CTableStylePr();
        TableStylePr.TextPr      = this.TextPr.Copy(undefined, oPr);
        TableStylePr.ParaPr      = this.ParaPr.Copy(undefined, oPr);
        TableStylePr.TablePr     = this.TablePr.Copy();
        TableStylePr.TableRowPr  = this.TableRowPr.Copy();
        TableStylePr.TableCellPr = this.TableCellPr.Copy();
        return TableStylePr;
    },

    Is_Equal : function(TableStylePr)
    {
        if (true !== this.TextPr.Is_Equal(TableStylePr.TextPr)
            || true !== this.ParaPr.Is_Equal(TableStylePr.ParaPr)
            || true !== this.TablePr.Is_Equal(TableStylePr.TablePr)
            || true !== this.TableRowPr.Is_Equal(TableStylePr.TableRowPr)
            || true !== this.TableCellPr.Is_Equal(TableStylePr.TableCellPr))
            return false;

        return true;
    },

    Check_PresentationPr : function(Theme)
    {
        this.TextPr.Check_PresentationPr();
        this.TableCellPr.Check_PresentationPr(Theme);
    },

    Set_FromObject : function(Obj)
    {
        if ( undefined != Obj.TextPr )
            this.TextPr.Set_FromObject( Obj.TextPr );

        if ( undefined != Obj.ParaPr )
            this.ParaPr.Set_FromObject( Obj.ParaPr );

        if ( undefined != Obj.TablePr )
            this.TablePr.Set_FromObject( Obj.TablePr );

        if ( undefined != Obj.TableRowPr )
            this.TableRowPr.Set_FromObject( Obj.TableRowPr );

        if ( undefined != Obj.TableCellPr )
            this.TableCellPr.Set_FromObject( Obj.TableCellPr );
    },

    Write_ToBinary : function(Writer)
    {
        this.TextPr.Write_ToBinary( Writer );
        this.ParaPr.Write_ToBinary( Writer );
        this.TablePr.Write_ToBinary( Writer );
        this.TableRowPr.Write_ToBinary( Writer );
        this.TableCellPr.Write_ToBinary( Writer );
    },

    Read_FromBinary : function(Reader)
    {
        this.TextPr.Read_FromBinary( Reader );
        this.ParaPr.Read_FromBinary( Reader );
        this.TablePr.Read_FromBinary( Reader );
        this.TableRowPr.Read_FromBinary( Reader );
        this.TableCellPr.Read_FromBinary( Reader );
    },

	InitDefault : function()
    {
        this.TextPr.InitDefault();
        this.ParaPr.InitDefault();
        this.TablePr.InitDefault();
        this.TableRowPr.InitDefault();
        this.TableCellPr.InitDefault();
    }
};

function CStyle(Name, BasedOnId, NextId, type, bNoCreateTablePr)
{
    this.Id = AscCommon.g_oIdCounter.Get_NewId();
	
	this.Parent = null;

	this.StyleId = null;

    this.Name    = Name;
    this.BasedOn = BasedOnId;
    this.Next    = NextId;
    this.Type    = (null != type ? type : styletype_Paragraph);
    this.Link    = null;
    this.Custom  = false;

    this.qFormat        = null;// false
    this.uiPriority     = null;// -1
    this.hidden         = null;// false
    this.semiHidden     = null;// false
    this.unhideWhenUsed = null;// false

    this.TextPr = new CTextPr();
    this.ParaPr = new CParaPr();

    this.TablePr     = new CTablePr();
    this.TableRowPr  = new CTableRowPr();
    this.TableCellPr = new CTableCellPr();
	
	if (true !== bNoCreateTablePr)
		this.InitConditionalTableStyles();

    // Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
    AscCommon.g_oTableId.Add( this, this.Id );
}

CStyle.prototype =
{
//-----------------------------------------------------------------------------------
// Работаем с Id данного объекта
//-----------------------------------------------------------------------------------
    GetId : function()
    {
        return this.Id;
    },

    Get_Id : function()
    {
        return this.GetId();
    },
//-----------------------------------------------------------------------------------
// Базовые функции для работы со стилем
//-----------------------------------------------------------------------------------
    Copy : function()
    {
        var Style = new CStyle();

        Style.Name           = this.Name;
        Style.BasedOn        = this.BasedOn;
        Style.Next           = this.Next;
        Style.Type           = this.Type;
        Style.qFormat        = this.qFormat       ;
        Style.uiPriority     = this.uiPriority    ;
        Style.hidden         = this.hidden        ;
        Style.semiHidden     = this.semiHidden    ;
        Style.unhideWhenUsed = this.unhideWhenUsed;

        Style.TextPr      = this.TextPr.Copy();
        Style.ParaPr      = this.ParaPr.Copy();
        Style.TablePr     = this.TablePr.Copy();
        Style.TableRowPr  = this.TableRowPr.Copy();
        Style.TableCellPr = this.TableCellPr.Copy();

        if (undefined !== this.TableBand1Horz)
        {
            Style.TableBand1Horz = this.TableBand1Horz.Copy();
            Style.TableBand1Vert = this.TableBand1Vert.Copy();
            Style.TableBand2Horz = this.TableBand2Horz.Copy();
            Style.TableBand2Vert = this.TableBand2Vert.Copy();
            Style.TableFirstCol  = this.TableFirstCol.Copy();
            Style.TableFirstRow  = this.TableFirstRow.Copy();
            Style.TableLastCol   = this.TableLastCol.Copy();
            Style.TableLastRow   = this.TableLastRow.Copy();
            Style.TableTLCell    = this.TableTLCell.Copy();
            Style.TableTRCell    = this.TableTRCell.Copy();
            Style.TableBLCell    = this.TableBLCell.Copy();
            Style.TableBRCell    = this.TableBRCell.Copy();
            Style.TableWholeTable= this.TableWholeTable.Copy();
        }

        return Style;
    },

	Set : function(style)
	{
		if (!style || !(style instanceof CStyle))
			return;

		this.Set_Name(style.Name);
		this.Set_BasedOn(style.BasedOn);
		this.Set_Next(style.Next);
		this.Set_Type(style.Type);
		this.Set_QFormat(style.qFormat);
		this.Set_UiPriority(style.uiPriority);
		this.Set_Hidden(style.hidden);
		this.Set_SemiHidden(style.semiHidden);
		this.Set_UnhideWhenUsed(style.unhideWhenUsed);
		this.Set_TextPr(style.TextPr.Copy());
		this.Set_ParaPr(style.ParaPr.Copy());

		this.Set_TablePr(style.TablePr.Copy());
		this.Set_TableRowPr(style.TableRowPr.Copy());
		this.Set_TableCellPr(style.TableCellPr.Copy());

		if (style.TableBand1Horz)
		{
			this.Set_TableBand1Horz(style.TableBand1Horz.Copy());
			this.Set_TableBand1Vert(style.TableBand1Vert.Copy());
			this.Set_TableBand2Horz(style.TableBand2Horz.Copy());
			this.Set_TableBand2Vert(style.TableBand2Vert.Copy());
			this.Set_TableFirstCol(style.TableFirstCol.Copy());
			this.Set_TableFirstRow(style.TableFirstRow.Copy());
			this.Set_TableLastCol(style.TableLastCol.Copy());
			this.Set_TableLastRow(style.TableLastRow.Copy());
			this.Set_TableTLCell(style.TableTLCell.Copy());
			this.Set_TableTRCell(style.TableTRCell.Copy());
			this.Set_TableBLCell(style.TableBLCell.Copy());
			this.Set_TableBRCell(style.TableBRCell.Copy());
			this.Set_TableWholeTable(style.TableWholeTable.Copy());
		}
	},

	RemapIdReferences : function(OldId, NewId)
	{
		if (OldId === this.BasedOn)
			this.Set_BasedOn(NewId);

		if (OldId === this.Link)
			this.Set_Link(NewId);

		if (OldId === this.Next)
			this.Set_Next(NewId);
	},

	Set_TextPr : function(Value)
	{
		var Old = this.TextPr;
		var New = new CTextPr();
		New.Set_FromObject(Value);

		this.TextPr = New;

		History.Add(new CChangesStyleTextPr(this, Old, New));
	},

	Set_ParaPr : function(Value, isHandleNumbering)
	{
		var Old = this.ParaPr;
		var New = new CParaPr();
		New.Set_FromObject(Value);
		
		if (isHandleNumbering && Value.NumPr instanceof AscWord.NumPr && Value.NumPr.IsValid())
		{
			var oLogicDocument = private_GetWordLogicDocument();
			if (oLogicDocument)
			{
				var oNumbering = oLogicDocument.GetNumbering();
				var oNum = oNumbering.GetNum(Value.NumPr.NumId);
				if (oNum)
				{
					var oNumLvl = oNum.GetLvl(Value.NumPr.Lvl).Copy();
					oNumLvl.SetPStyle(this.GetId());
					oNum.SetLvl(oNumLvl, Value.NumPr.Lvl);

					New.NumPr = new AscWord.NumPr(Value.NumPr.NumId, Value.NumPr.Lvl);
				}
			}
		}

		if (Old.IsEqual(New))
			return;
		
		let change = new CChangesStyleParaPr(this, Old, New);
		AscCommon.History.Add(change);
		change.Redo();
	},

	Set_TablePr : function(Value)
	{
		var Old = this.TablePr;
		var New = new CTablePr();
		New.Set_FromObject(Value);

		this.TablePr = New;

		History.Add(new CChangesStyleTablePr(this, Old, New));
	},

	Set_TableRowPr : function(Value)
	{
		var Old = this.TableRowPr;
		var New = new CTableRowPr();
		New.Set_FromObject(Value);

		this.TableRowPr = New;

		History.Add(new CChangesStyleTableRowPr(this, Old, New));
	},

	Set_TableCellPr : function(Value)
	{
		var Old = this.TableCellPr;
		var New = new CTableCellPr();
		New.Set_FromObject(Value);

		this.TableCellPr = New;

		History.Add(new CChangesStyleTableCellPr(this, Old, New));
	},

	Set_TableBand1Horz : function(Value)
	{
		var Old = this.TableBand1Horz;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableBand1Horz = New;

		History.Add(new CChangesStyleTableBand1Horz(this, Old, New));
	},

	Set_TableBand1Vert : function(Value)
	{
		var Old = this.TableBand1Vert;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableBand1Vert = New;

		History.Add(new CChangesStyleTableBand1Vert(this, Old, New));
	},

    Set_TableBand2Horz : function(Value)
    {
        var Old = this.TableBand2Horz;
        var New = new CTableStylePr();
        New.Set_FromObject(Value);

        this.TableBand2Horz = New;

        History.Add(new CChangesStyleTableBand2Horz(this, Old, New));
    },

	Set_TableBand2Vert : function(Value)
	{
		var Old = this.TableBand2Vert;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableBand2Vert = New;

		History.Add(new CChangesStyleTableBand2Vert(this, Old, New));
	},

	Set_TableFirstCol : function(Value)
	{
		var Old = this.TableFirstCol;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableFirstCol = New;

		History.Add(new CChangesStyleTableFirstCol(this, Old, New));
	},

	Set_TableFirstRow : function(Value)
	{
		var Old = this.TableFirstRow;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableFirstRow = New;

		History.Add(new CChangesStyleTableFirstRow(this, Old, New));
	},

	Set_TableLastCol : function(Value)
	{
		var Old = this.TableLastCol;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableLastCol = New;

		History.Add(new CChangesStyleTableLastCol(this, Old, New));
	},

	Set_TableLastRow : function(Value)
	{
		var Old = this.TableLastRow;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableLastRow = New;

		History.Add(new CChangesStyleTableLastRow(this, Old, New));
	},

	Set_TableTLCell : function(Value)
	{
		var Old = this.TableTLCell;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableTLCell = New;

		History.Add(new CChangesStyleTableTLCell(this, Old, New));
	},

	Set_TableTRCell : function(Value)
	{
		var Old = this.TableTRCell;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableTRCell = New;

		History.Add(new CChangesStyleTableTRCell(this, Old, New));
	},

	Set_TableBLCell : function(Value)
	{
		var Old = this.TableBLCell;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableBLCell = New;

		History.Add(new CChangesStyleTableBLCell(this, Old, New));
	},

	Set_TableBRCell : function(Value)
	{
		var Old = this.TableBRCell;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableBRCell = New;

		History.Add(new CChangesStyleTableBRCell(this, Old, New));
	},

	Set_TableWholeTable : function(Value)
	{
		var Old = this.TableWholeTable;
		var New = new CTableStylePr();
		New.Set_FromObject(Value);

		this.TableWholeTable = New;

		History.Add(new CChangesStyleTableWholeTable(this, Old, New));
	},

	Set_Name : function(Value)
	{
		History.Add(new CChangesStyleName(this, this.Name, Value));
		this.Name = Value;
	},

    Get_Name : function(isSimplify)
    {
        return isSimplify ? this.Name.toLowerCase().replace(/\s/g,"") : this.Name;
    },

	Set_BasedOn : function(Value)
	{
		History.Add(new CChangesStyleBasedOn(this, this.BasedOn, Value));
		this.BasedOn = Value;
	},

    Get_BasedOn : function()
    {
        return this.BasedOn;
    },

	Set_Next : function(Value)
	{
		History.Add(new CChangesStyleNext(this, this.Next, Value));
		this.Next = Value;
	},

    Get_Next : function()
    {
        return this.Next;
    },

	Set_Link : function(Value)
	{
		History.Add(new CChangesStyleLink(this, this.Link, Value));
		this.Link = Value;
	},

	Get_Link : function()
	{
		return this.Link;
	},

	Set_Type : function(Value)
	{
		History.Add(new CChangesStyleType(this, this.Type, Value));
		this.Type = Value;
	},

    Get_Type : function()
    {
        return this.Type;
    },

	Set_QFormat : function(Value)
	{
		History.Add(new CChangesStyleQFormat(this, this.qFormat, Value));
		this.qFormat = Value;
	},

	Set_UiPriority : function(Value)
	{
		History.Add(new CChangesStyleUiPriority(this, this.uiPriority, Value));
		this.uiPriority = Value;
	},

	Set_Hidden : function(Value)
	{
		History.Add(new CChangesStyleHidden(this, this.hidden, Value));
		this.hidden = Value;
	},

	Set_SemiHidden : function(Value)
	{
		History.Add(new CChangesStyleSemiHidden(this, this.semiHidden, Value));
		this.semiHidden = Value;
	},

	Set_UnhideWhenUsed : function(Value)
	{
		History.Add(new CChangesStyleUnhideWhenUsed(this, this.unhideWhenUsed, Value));
		this.unhideWhenUsed = Value;
	},

    Clear : function(Name, BasedOnId, NextId, Type)
    {
        this.Set_Name(Name);
        this.Set_BasedOn(BasedOnId);
        this.Set_Next(NextId);
        this.Set_Type(Type);

        if (undefined != this.Link && null != this.Link)
            this.Set_Link(null);

        if (undefined != this.qFormat && null != this.qFormat)
            this.Set_QFormat(null);

        if (undefined != this.uiPriority && null != this.uiPriority)
            this.Set_UiPriority(null);

        if (undefined != this.hidden && null != this.hidden)
            this.Set_Hidden(null);

        if (undefined != this.semiHidden && null != this.semiHidden)
            this.Set_SemiHidden(null);

        if (undefined != this.unhideWhenUsed && null != this.unhideWhenUsed)
            this.Set_UnhideWhenUsed(null);

        this.Set_TextPr(new CTextPr());
        this.Set_ParaPr(new CParaPr());
        this.Set_TablePr(new CTablePr());
        this.Set_TableRowPr(new CTableRowPr());
        this.Set_TableCellPr(new CTableCellPr());

        if (undefined != this.TableBand1Horz ) this.Set_TableBand1Horz (new CTableStylePr());
        if (undefined != this.TableBand1Vert ) this.Set_TableBand1Vert (new CTableStylePr());
        if (undefined != this.TableBand2Horz ) this.Set_TableBand2Horz (new CTableStylePr());
        if (undefined != this.TableBand2Vert ) this.Set_TableBand2Vert (new CTableStylePr());
        if (undefined != this.TableFirstCol  ) this.Set_TableFirstCol  (new CTableStylePr());
        if (undefined != this.TableFirstRow  ) this.Set_TableFirstRow  (new CTableStylePr());
        if (undefined != this.TableLastCol   ) this.Set_TableLastCol   (new CTableStylePr());
        if (undefined != this.TableLastRow   ) this.Set_TableLastRow   (new CTableStylePr());
        if (undefined != this.TableTLCell    ) this.Set_TableTLCell    (new CTableStylePr());
        if (undefined != this.TableTRCell    ) this.Set_TableTRCell    (new CTableStylePr());
        if (undefined != this.TableBLCell    ) this.Set_TableBLCell    (new CTableStylePr());
        if (undefined != this.TableBRCell    ) this.Set_TableBRCell    (new CTableStylePr());
        if (undefined != this.TableWholeTable) this.Set_TableWholeTable(new CTableStylePr());
    },
//-----------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------
	Document_Get_AllFontNames : function(allFonts)
	{
		if (this.TextPr)
			this.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableBand1Horz && this.TableBand1Horz.TextPr)
			this.TableBand1Horz.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableBand1Vert && this.TableBand1Vert.TextPr)
			this.TableBand1Vert.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableBand2Horz && this.TableBand2Horz.TextPr)
			this.TableBand2Horz.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableBand2Vert && this.TableBand2Vert.TextPr)
			this.TableBand2Vert.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableFirstCol && this.TableFirstCol.TextPr)
			this.TableFirstCol.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableFirstRow && this.TableFirstRow.TextPr)
			this.TableFirstRow.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableLastCol && this.TableLastCol.TextPr)
			this.TableLastCol.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableLastRow && this.TableLastRow.TextPr)
			this.TableLastRow.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableTLCell && this.TableTLCell.TextPr)
			this.TableTLCell.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableTRCell && this.TableTRCell.TextPr)
			this.TableTRCell.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableBLCell && this.TableBLCell.TextPr)
			this.TableBLCell.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableBRCell && this.TableBRCell.TextPr)
			this.TableBRCell.TextPr.Document_Get_AllFontNames(allFonts);
		
		if (this.TableWholeTable && this.TableWholeTable.TextPr)
			this.TableWholeTable.TextPr.Document_Get_AllFontNames(allFonts);
	},

	private_CreateDefaultUnifillColor : function()
	{
		var Unifill        = new AscFormat.CUniFill();
		Unifill.fill       = new AscFormat.CSolidFill();
		Unifill.fill.color = AscFormat.builder_CreateSchemeColor('tx1');
		return Unifill;
	},

    Create_NormalTable : function()
    {
        var TablePr =
        {
            TableInd : 0 ,
            TableCellMar :
            {
                Top :
                {
                    W    : 0,
                    Type : tblwidth_Mm
                },

                Left :
                {
                    W    : 1.9, // 5.4pt
                    Type : tblwidth_Mm
                },

                Bottom :
                {
                    W    : 0,
                    Type : tblwidth_Mm
                },

                Right :
                {
                    W    : 1.9, // 5.4pt
                    Type : tblwidth_Mm
                }
            }
        };

        this.Set_UiPriority( 99 );
        this.Set_SemiHidden( true );
        this.Set_UnhideWhenUsed( true );
        this.Set_TablePr( TablePr );
    },

    Create_TableGrid: function () {
        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            }
        };

        var TablePr =
        {
            TableInd: 0,
            TableBorders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single
                },

                InsideH:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single
                },

                InsideV:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single
                }
            },

            TableCellMar:
            {
                Top:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Left:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                },

                Bottom:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Right:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                }
            }
        };

        this.Set_UiPriority(59);
        this.Set_TablePr(TablePr);
        this.Set_ParaPr(ParaPr);
    },

    Create_Table_LightShading : function()
    {
        this.uiPriority = 60;

        var ParaPr =
        {
            Spacing :
            {
                After    : 0,
                Line     : 1,
                LineRule : linerule_Auto
            }
        };

        var TablePr =
        {
            TableStyleColBandSize : 1,
            TableStyleRowBandSize : 1,
            TableInd : 0,

            TableBorders :
            {
                Top :
                {
                    Color : { r : 0, g : 0, b : 0 },
                    Space : 0,
                    Size  : 18 / 8 * g_dKoef_pt_to_mm,
                    Value : border_Single
                },

                Bottom :
                {
                    Color : { r : 0, g : 0, b : 0 },
                    Space : 0,
                    Size  : 18 / 8 * g_dKoef_pt_to_mm,
                    Value : border_Single
                }
            },

            TableCellMar :
            {
                Top :
                {
                    W    : 0,
                    Type : tblwidth_Mm
                },
                Left :
                {
                    W    : 5.75 * g_dKoef_pt_to_mm, // 0.08 inch
                    Type : tblwidth_Mm
                },

                Bottom :
                {
                    W    : 0,
                    Type : tblwidth_Mm
                },

                Right :
                {
                    W    : 5.75 * g_dKoef_pt_to_mm, // 0.08 inch
                    Type : tblwidth_Mm
                }
            }
        };

        var TableFirstRow =
        {
            TextPr :
            {
                Bold : true,
                Color :
                {
                    r : 255,
                    g : 255,
                    b : 255
                }
            },

            ParaPr :
            {
                Spacing :
                {
                    After    : 0,
                    Before   : 0,
                    Line     : 1,
                    LineRule : linerule_Auto
                }
            },

            TableCellPr :
            {
                TableCellBorders :
                {
                    Bottom :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    },

                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    },

                    Top :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0x4F,
                        g : 0x81,
                        b : 0xBD
                    }
                }
            }
        };

        var TableLastRow =
        {
            TextPr :
            {
                Color :
                {
                    r : 0,
                    g : 0,
                    b : 0
                }
            },

            TableCellPr :
            {
                TableCellBorders :
                {
                    Bottom :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    },

                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    },

                    Top :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 12 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0xFF,
                        g : 0xFF,
                        b : 0xFF
                    }
                }
            }
        };

        var TableFirstCol =
        {
            TextPr :
            {
                Bold  : true,
                Color :
                {
                    r : 255,
                    g : 255,
                    b : 255
                }
            },

            ParaPr :
            {
                Spacing :
                {
                    After    : 0,
                    Before   : 0,
                    Line     : 1,
                    LineRule : linerule_Auto
                }
            },

            TableCellPr :
            {
                TableCellBorder :
                {
                    Bottom :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    },

                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    },

                    Top :
                    {
                        Value : border_None
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0x4F,
                        g : 0x81,
                        b : 0xBD
                    }
                }
            }

        };

        var TableLastCol =
        {
            TextPr :
            {
                Bold  : true,
                Color :
                {
                    r : 255,
                    g : 255,
                    b : 255
                }
            },

            TableCellPr :
            {
                TableCellBorders :
                {
                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0x4F,
                        g : 0x81,
                        b : 0xBD
                    }
                }
            }
        };

        var TableBand1Vert =
        {
            TableCellPr :
            {
                TableCellBorders :
                {
                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0xD8,
                        g : 0xD8,
                        b : 0xD8
                    }
                }
            }
        };

        var TableBand1Horz =
        {
            TableCellPr :
            {
                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0xD8,
                        g : 0xD8,
                        b : 0xD8
                    }
                }
            }
        };

        var TableTRCell =
        {
            TableCellPr :
            {
                TableCellBorders:
                {
                    Bottom :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    },

                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    },

                    Top :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 0,
                        g : 255,
                        b : 0
                    }
                }
            }
        };

        var TableTLCell =
        {
            TextPr :
            {
                Color :
                {
                    r : 255,
                    g : 255,
                    b : 255
                }
            },

            TableCellPr :
            {
                TableCellBorders:
                {
                    Bottom :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    },

                    Left :
                    {
                        Value : border_None
                    },

                    Right :
                    {
                        Value : border_None
                    },

                    Top :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 18 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color :
                    {
                        r : 255,
                        g : 0,
                        b : 0
                    }
                }
            }
        };

        this.Set_UiPriority( 60 );
        this.Set_ParaPr( ParaPr );
        this.Set_TablePr( TablePr );
        this.Set_TableFirstRow( TableFirstRow );
        this.Set_TableLastRow( TableLastRow );
        this.Set_TableFirstCol( TableFirstCol );
        this.Set_TableLastCol( TableLastCol );
        this.Set_TableBand1Horz( TableBand1Horz );
        this.Set_TableBand1Vert( TableBand1Vert );
        this.Set_TableTRCell( TableTRCell );
        this.Set_TableTLCell( TableTLCell );
    },

    Create_Table_ColorfulListAccent6 : function()
    {
        this.uiPriority = 72;

        var ParaPr =
        {
            Spacing :
            {
                After    : 0,
                Line     : 1,
                LineRule : linerule_Auto
            }
        };

        var TextPr =
        {
            Color : { r : 0, g : 0, b : 0 }
        };

        var TablePr =
        {
            TableStyleColBandSize : 1,
            TableStyleRowBandSize : 1,
            TableInd              : 0,

            TableCellMar :
            {
                TableCellMar :
                {
                    Top :
                    {
                        W    : 0,
                        Type : tblwidth_Mm
                    },
                    Left :
                    {
                        W    : 5.75 * g_dKoef_pt_to_mm, // 0.08 inch
                        Type : tblwidth_Mm
                    },

                    Bottom :
                    {
                        W    : 0,
                        Type : tblwidth_Mm
                    },

                    Right :
                    {
                        W    : 5.75 * g_dKoef_pt_to_mm, // 0.08 inch
                        Type : tblwidth_Mm
                    }
                }
            }
        };

        var TableCellPr =
        {
            Shd :
            {
                Value : c_oAscShdClear,
                Color : { r : 0xFE, g : 0xF4, b : 0xEC }
            }
        };

        var TableFirstRow =
        {
            TextPr :
            {
                Bold  : true,
                Color : { r : 0xFF, g : 0xFF, b : 0xFF }
            },

            TableCellPr :
            {
                TableCellBorders :
                {
                    Bottom :
                    {
                        Color : { r : 0xFF, g : 0xFF, b : 0xFF },
                        Space : 0,
                        Size  : 12 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
					},
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color : { r : 0x34, g : 0x8D, b : 0xA5 }
                }
            }
        };

        var TableLastRow =
        {
            TextPr :
            {
                Bold  : true,
                Color : { r : 0x34, g : 0x8D, b : 0xA5 }
            },

            TableCellPr :
            {
                TableCellBorders :
                {
                    Top :
                    {
                        Color : { r : 0, g : 0, b : 0 },
                        Space : 0,
                        Size  : 12 / 8 * g_dKoef_pt_to_mm,
                        Value : border_Single
                    }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color : { r : 0xFF, g : 0xFF, b : 0xFF }
                }
            }
        };

        var TableFirstCol =
        {
            TextPr :
            {
                Bold : true
            }
        };

        var TableLastCol =
        {
            TextPr :
            {
                Bold : true
            }
        };

        var TableBand1Vert =
        {
            TableCellPr :
            {
                TableCellBorders :
                {
                    Top     : { Value : border_None },
                    Left    : { Value : border_None },
                    Bottom  : { Value : border_None },
                    Right   : { Value : border_None },
                    InsideH : { Value : border_None },
                    InsideV : { Value : border_None }
                },

                Shd :
                {
                    Value : c_oAscShdClear,
                    Color : { r : 0xFD, g : 0xE4, b : 0xD0 }
                }
            }
        };

        var TableBand1Horz =
        {
            TableCellPr :
            {
				TableCellBorders :
                {
                    Top     : { Value : border_None },
                    Left    : { Value : border_None },
                    Bottom  : { Value : border_None },
                    Right   : { Value : border_None },
                    InsideH : { Value : border_None },
                    InsideV : { Value : border_None }
                },
                Shd :
                {
                    Value : c_oAscShdClear,
                    Color : { r : 0xFD, g : 0xE9, b : 0xD9 }
                }
            }
        };

        this.Set_UiPriority( 72 );
        this.Set_ParaPr( ParaPr );
        this.Set_TextPr( TextPr );
        this.Set_TablePr( TablePr );
        this.Set_TableCellPr( TableCellPr );
        this.Set_TableFirstRow( TableFirstRow );
        this.Set_TableLastRow( TableLastRow );
        this.Set_TableFirstCol( TableFirstCol );
        this.Set_TableLastCol( TableLastCol );
        this.Set_TableBand1Horz( TableBand1Horz );
        this.Set_TableBand1Vert( TableBand1Vert );
    },

    Create_Table_Lined : function(oUnifill1, oUnifill2)
    {
        var TextColor1 = new CDocumentColor(0xF2, 0xF2, 0xF2, false);
        var TextFont1  = { Name : "Arial", Index : -1 };
        var TextSize1  = 11;

        var CellShd1   = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oUnifill1,
			ThemeFill  : oUnifill1
		});
		var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oUnifill2,
			ThemeFill  : oUnifill2
		});

        var TableStylePrBoundary =
        {
            TextPr :
            {
                RFonts   : { Ascii : TextFont1, HAnsi : TextFont1 },
                Color    : TextColor1,
                FontSize : TextSize1
            },

            TableCellPr :
            {
                Shd : CellShd1
            }
        };

        var ParaPr =
        {
            Spacing :
            {
                After    : 0,
                Line     : 1,
                LineRule : linerule_Auto
            },

            Borders :
            {
                Top :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Left :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Bottom :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Right :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Between :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                }
            }
        };

        var TextPr =
        {
            Color : { r : 0x40, g : 0x40, b : 0x40 }
        };

        var TablePr =
        {
            TableStyleColBandSize : 1,
            TableStyleRowBandSize : 1,
            TableInd : 0,

            // TableCellMar :
            // {
            //     Top    : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
            //     Left   : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm),
            //     Bottom : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
            //     Right  : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm)
            // }
        };

        var TableStylePrBand1 =
        {
            TextPr :
            {
                RFonts   : { Ascii : TextFont1, HAnsi : TextFont1 },
                Color    : { r : 0x40, g : 0x40, b : 0x40 },
                FontSize : TextSize1
            }
        };

        var TableStylePrBand2 =
        {
            TextPr :
            {
                RFonts   : { Ascii : TextFont1, HAnsi : TextFont1 },
                Color    : { r : 0x40, g : 0x40, b : 0x40 },
                FontSize : TextSize1
            },

            TableCellPr :
            {
                Shd : CellShd2

            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TextPr(TextPr);
        this.Set_TablePr(TablePr);

        this.Set_TableFirstRow(TableStylePrBoundary);
        this.Set_TableLastRow(TableStylePrBoundary);
        this.Set_TableFirstCol(TableStylePrBoundary);
        this.Set_TableLastCol(TableStylePrBoundary);

        this.Set_TableBand1Horz(TableStylePrBand1);
        this.Set_TableBand1Vert(TableStylePrBand1);

        this.Set_TableBand2Horz(TableStylePrBand2);
        this.Set_TableBand2Vert(TableStylePrBand2);
	},

    Create_TableGrid_Light: function (oBorderUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            }
        };

        var TablePr =
        {
            TableInd: 0,

            TableBorders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: CellShd1.Unifill
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: CellShd1.Unifill
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: CellShd1.Unifill
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: CellShd1.Unifill
                },

                InsideH:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: CellShd1.Unifill
                },

                InsideV:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: CellShd1.Unifill
                }
            },

            TableCellMar:
            {
                Top:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Left:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                },

                Bottom:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Right:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                }
            }
        };

        this.Set_UiPriority(59);
        this.Set_TablePr(TablePr);
        this.Set_ParaPr(ParaPr);
    },

    Create_Table_Plain_1: function (oBorderUnifill, oBandUnifill)
	{
        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            }
        };
        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11,
            Bold: true
        }
        var TablePr =
        {
            TableInd: 0,

            TableBorders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                InsideH:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                InsideV:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                }
            },

            TableCellMar:
            {
                Top:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Left:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                },

                Bottom:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Right:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                }
            }
        };
        var TableBand1 =
        {
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        var TableText =
        {
            TextPr: TableTextPr
        };

        this.Set_UiPriority(59);
        this.Set_TablePr(TablePr);
        this.Set_ParaPr(ParaPr);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
        this.Set_TableFirstRow(TableText);
        this.Set_TableLastRow(TableText);
        this.Set_TableFirstCol(TableText);
        this.Set_TableLastCol(TableText);
    },

    Create_Table_Plain_2: function (oBorderUnifill)
	{
        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: oBorderUnifill
        };

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11,
            Bold: true
        }
        var TablePr =
        {
            TableInd: 0,

            TableBorders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_None,
                    Unifill: oBorderUnifill
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_Single,
                    Unifill: oBorderUnifill
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0 },
                    Space: 0,
                    Size: 0.5 * g_dKoef_pt_to_mm,
                    Value: border_None,
                    Unifill: oBorderUnifill
                }
            },

            TableCellMar:
            {
                Top:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Left:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                },

                Bottom:
                {
                    W: 0,
                    Type: tblwidth_Mm
                },

                Right:
                {
                    W: 1.9,
                    Type: tblwidth_Mm
                }
            }
        };

        var TableBand1Horz =
        {
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1,
                    Bottom: TableBorder1
                }
            }
        };

        var TableBand1Vert =
        {
            TableCellPr:
            {
                TableCellBorders:
                {
                    Left: TableBorder1,
                    Right: TableBorder1
                }
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr
        }

        var TableFirstRow =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1,
                    Bottom: TableBorder1
                }
            }
        }
        var TableLastRow =
        {
            TextPr: TableTextPr,
            TableCellBorders:
            {
                Top: TableBorder1
            }
        }
        this.Set_UiPriority(59);
        this.Set_TablePr(TablePr);
        this.Set_ParaPr(ParaPr);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
        this.Set_TableBand2Vert(TableBand1Vert);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
    },

    Create_Table_Plain_3: function (oBorderUnifill, oBandUnifill)
	{
        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true,
            Caps: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0
        };
        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_None,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder1,
                    Left: TableBorder2
                }
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
            }
        };

        var TableFirstCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Left: TableBorder2,
                    Right: TableBorder1,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            }
        };

        var TableLastCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
            }
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_Plain_4: function (oBandUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0
        };

        var TableStyle =
        {
            TextPr: TableTextPr1,
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd1
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableStyle);
        this.Set_TableLastRow(TableStyle);
        this.Set_TableFirstCol(TableStyle);
        this.Set_TableLastCol(TableStyle);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_Plain_5: function (oTableCellUnifill, oBandUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Italic: true,
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0
        };
        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_None,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Bottom: TableBorder1,
                    Left: TableBorder2
                },
                Shd: CellShd1
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1,
                    Right: TableBorder2,
                    Left: TableBorder2
                },
                Shd: CellShd1
            }
        };

        var TableFirstCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder1
                },
                Shd: CellShd1
            },
            ParaPr:
            {
                Jc: align_Right
            }
        };

        var TableLastCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Left: TableBorder1
                },
                Shd: CellShd1
            }
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_Grid_1: function (oFirstRowBottomBorderUnifill, oBorderUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oFirstRowBottomBorderUnifill,
			ThemeFill  : oFirstRowBottomBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 1.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd2.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Top: TableBorder2,
                Left: TableBorder2,
                Bottom: TableBorder2,
                Right: TableBorder2,
                InsideH: TableBorder2,
                InsideV: TableBorder2
            }
        };

        var TableLastRow =
        {
            //TODO: Реализовать двойную линию
            TextPr: TableTextPr1
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Bottom: TableBorder1
                }
            }
        };


        var TableStyle =
        {
            TextPr: TableTextPr1
        };

        var TableBand1Horz =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder2,
                    Left: TableBorder2,
                    Bottom: TableBorder2,
                    Right: TableBorder2,
                    InsideH: TableBorder2,
                    InsideV: TableBorder2
                }
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableStyle);
        this.Set_TableLastCol(TableStyle);
        this.Set_TableBand1Horz(TableBand1Horz);
    },

    Create_Table_Grid_2: function (oBorderUnifill, oBandUniFill, oTableCellUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUniFill,
			ThemeFill  : oBandUniFill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };
        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 1.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };
        var TableBorder3 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_None,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,

        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Bottom: TableBorder1,
                InsideH: TableBorder1,
                InsideV: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder3,
                    Top: TableBorder3,
                    Bottom: TableBorder2,
                    Left: TableBorder3
                },
                Shd: CellShd3
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder3,
                    Top: TableBorder1,
                    Bottom: TableBorder3,
                    Left: TableBorder3
                },
                Shd: CellShd3
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr1
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_Grid_3: function (oBorderUnifill, oBandUnifill, oTableCellUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TextColumnPr =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Italic: true
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_None,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Bottom: TableBorder1,
                InsideH: TableBorder1,
                InsideV: TableBorder1
            }
        };

        var TableRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder2,
                    Left: TableBorder2
                },
                Shd: CellShd3
            }
        };


        var TableFirstCol =
        {
            TextPr: TextColumnPr,
            TableCellPr:
            {
                Shd:
                {
                    Color:
                    {
                        r: 255,
                        g: 255,
                        b: 255
                    }
                },
                TableCellBorders:
                {
                    Left: TableBorder2,
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            },
            ParaPr:
            {
                Jc: align_Right
            }
        };

        var TableLastCol =
        {
            TextPr: TextColumnPr,
            TableCellPr:
            {
                Shd:
                {
                    Color:
                    {
                        r: 255,
                        g: 255,
                        b: 255
                    }
                },
                TableCellBorders:
                {
                    Left: TableBorder2,
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            }
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableRow);
        this.Set_TableLastRow(TableRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_Grid_4: function (oHeaderUnifill, oBandUnifill, oBorderUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oHeaderUnifill,
			ThemeFill  : oHeaderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var TableTextPr2 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 255, g: 255, b: 255 },
            FontSize: 11,
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd3.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder3 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5001 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Left: TableBorder1,
                Right: TableBorder1,
                Top: TableBorder1,
                Bottom: TableBorder1,
                InsideH: TableBorder1,
                InsideV: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr2,
            TableCellPr:
            {
                Shd: CellShd1,
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Left: TableBorder2,
                    Bottom: TableBorder2,
                    Top: TableBorder2
                }
            }
        };

        var TableLastRow =
        {
             //TODO: Реализовать двойную линию
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder3
                }
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr1
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(59);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_Grid_5: function (oHeaderUnifill, oTableCellUnifill, oBorderUnifill, oBandUnifill)
	{
		var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oHeaderUnifill,
			ThemeFill  : oHeaderUnifill
		});

		var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

		var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

		var CellShd4 = new CDocumentShd();
		CellShd4.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };
        var TableTextPr1 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 255, g: 255, b: 255 },
            Bold: true,
            FontSize: 11
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd3.Unifill
        };

        var TablePr =
        {
            TextPr: TableTextPr,
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,
            Shd: CellShd2,
            TableBorders:
            {
                Left: TableBorder1,
                Right: TableBorder1,
                Top: TableBorder1,
                Bottom: TableBorder1,
                InsideH: TableBorder1,
                InsideV: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd: CellShd1,
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd: CellShd1
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd: CellShd1,
                TableCellBorders:
                {
                    Top: TableBorder1
                }
            }
        };

        var TableBand1Horz =
        {
            TableCellPr:
            {
                Shd: CellShd4
            }
        };

        var TableBand1Vert =
        {
            TableCellPr:
            {
                Shd: CellShd4
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
    },

    Create_Table_Grid_6: function (oBorderUnifill, oBandUnifill, oTextUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTextUnifill,
			ThemeFill  : oTextUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11,
            Unifill: CellShd3.Unifill
        };

        var TableTextPr1 =
        {
            Unifill: CellShd3.Unifill,
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 1.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Bottom: TableBorder1,
                InsideH: TableBorder1,
                InsideV: TableBorder1,
                Top: TableBorder1,
                Left: TableBorder1,
                Right: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Bottom: TableBorder2
                }
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1
        };

        var TableCol =
        {
            TextPr: TableTextPr1,
        };

        var TableBand1Horz =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };
        var TableBand1Vert =
        {
            TableCellPr:
            {
                Shd: CellShd2
            }
        };
        var TableBand2Horz =
        {
            TextPr: TableTextPr
        };
        var TableWholeTable =
        {
            TextPr: TableTextPr
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand2Horz(TableBand2Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
        this.Set_TableWholeTable(TableWholeTable);
    },

    Create_Table_Grid_7: function (oBorderUnifill, oBandUnifill, oTableCellUnifill, oTextUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

        var CellShd4 = new CDocumentShd();
		CellShd4.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTextUnifill,
			ThemeFill  : oTextUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Unifill: CellShd4.Unifill,
            Bold: true,
            FontSize: 11
        };

        var TableTextPr1 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Unifill: CellShd4.Unifill,
            Italic: true,
            FontSize: 11
        };
        var TableTextPr2 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Unifill: CellShd4.Unifill,
            FontSize: 11
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Value: border_None
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Bottom: TableBorder1,
                InsideH: TableBorder1,
                InsideV: TableBorder1,
                Right: TableBorder1
            },
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr,

            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder1,
                    Left: TableBorder2
                },
                Shd: CellShd3
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1,
                    Bottom: TableBorder2,
                    Right: TableBorder2,
                    Left: TableBorder2
                },
                Shd: CellShd3
            }
        };

        var TableFirstCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd:
                {
                    Color:
                    {
                        r: 255,
                        g: 255,
                        b: 255
                    }
                },
                TableCellBorders:
                {
                    Left: TableBorder2,
                    Right: TableBorder1,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            },
            ParaPr:
            {
                Jc: align_Right
            }
        };

        var TableLastCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd:
                {
                    Color:
                    {
                        r: 255,
                        g: 255,
                        b: 255
                    }
                },
                TableCellBorders:
                {
                    Left: TableBorder1,
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            }
        };

        var TableBand1Horz =
        {
            TextPr: TableTextPr2,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        var TableBand2Horz =
        {
            TextPr: TableTextPr2
        };

        var TableBand1Vert =
        {
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand2Horz(TableBand2Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
    },

    Create_Table_List_1: function (oBorderUnifill, oBandUnifill, oTableCellUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_None,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,

            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder1,
                    Left: TableBorder2
                }
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder1,
                    Bottom: TableBorder2,
                    Left: TableBorder2
                }
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr1
        };

        var TableBand1 =
        {
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_List_2: function (oUnifill, oBandUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oUnifill,
			ThemeFill  : oUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11,
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_None,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,

        };

        var TablePr =
        {
            TextPr: TableTextPr,
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                InsideH: TableBorder1,
                Bottom: TableBorder1,
                Top: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,

            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder1,
                    Bottom: TableBorder1,
                    Left: TableBorder2
                }
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Bottom: TableBorder1,
                    Top: TableBorder1,
                    Left: TableBorder2
                }
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr1

        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_List_3: function (oHeaderUnifill, oBorderUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oHeaderUnifill,
			ThemeFill  : oHeaderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var TextFirstRowPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 255, g: 255, b: 255 },
            FontSize: 11,
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Left: TableBorder1,
                Right: TableBorder1,
                Top: TableBorder1,
                Bottom: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TextFirstRowPr,
            TableCellPr:
            {
                Shd: CellShd1
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1
        };

        var TableCol =
        {
            TextPr: TableTextPr1
        };

        var TableBand1Vert =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Left: TableBorder1,
                    Right: TableBorder1
                }
            }
        };

        var TableBand1Horz =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1,
                    Bottom: TableBorder1
                }
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
    },

    Create_Table_List_4: function (oHeaderUnifill, oBandUnifill, oBorderUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oHeaderUnifill,
			ThemeFill  : oHeaderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11
        };

        var TableTextPr1 =
        {
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            Bold: true
        };

        var TextFirstRowPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 255, g: 255, b: 255 },
            FontSize: 11,
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd3.Unifill
        };

        var TablePr =
        {
            TextPr: TableTextPr,
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Left: TableBorder2,
                Right: TableBorder2,
                Top: TableBorder2,
                Bottom: TableBorder2,
                InsideH: TableBorder2
            }
        };

        var TableFirstRow =
        {
            TextPr: TextFirstRowPr,
            TableCellPr:
            {
                Shd: CellShd1
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1
        };

        var TableCol =
        {
            TextPr: TableTextPr1
        };

        var TableBand1 =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1);
        this.Set_TableBand1Vert(TableBand1);
    },

    Create_Table_List_5: function (oTableCellUnifill, oBorderUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableCellUnifill,
			ThemeFill  : oTableCellUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Unifill: CellShd2.Unifill,
            FontSize: 11
        };

        var TableTextPr1 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            FontSize: 11,
            Bold: true,
            Unifill: CellShd2.Unifill
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Unifill: CellShd1.Unifill,
                    Space: 0,
                    Size: 0,
                    Value: border_Single
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Unifill: CellShd1.Unifill,
                    Space: 0,
                    Size: 0,
                    Value: border_Single
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_Single
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 4 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 1.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd2.Unifill
        };
        var TableBorder3 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd2.Unifill
        };

        var TablePr =
        {
            TextPr: TableTextPr,
            Shd: CellShd1,
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Left: TableBorder1,
                Right: TableBorder1,
                Top: TableBorder1,
                Bottom: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd: CellShd1,
                TableCellBorders:
                {
                    Bottom: TableBorder2,
                    Top: TableBorder1
                }
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1
        };

        var TableFirstCol =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder3,
                    Left: TableBorder1
                }
            }
        };
        var TableLastCol =
        {
            TableCellPr:
            {
                TableCellBorders:
                {
                    Left: TableBorder3,
                    Right: TableBorder1
                }
            }
        };

        var TableBand1Horz =
        {
            TableCellPr:
            {
                Shd: CellShd1,
                TableCellBorders:
                {
                    Top: TableBorder3,
                    Bottom: TableBorder3
                }
            }
        };

        var TableBand2Horz =
        {
            TableCellPr:
            {
                Shd: CellShd1,
                TableCellBorders:
                {
                    Top: TableBorder3,
                    Bottom: TableBorder3
                }
            }
        };

        var TableBand1Vert =
        {
            TableCellPr:
            {
                Shd: CellShd1,
                TableCellBorders:
                {
                    Left: TableBorder3,
                    Right: TableBorder3
                }
            }
        };

        var TableBand2Vert =
        {
            Shd: CellShd1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Left: TableBorder3,
                    Right: TableBorder3
                }
            }
        };
        var TableWholeTable =
        {
            TextPr: TableTextPr
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand2Horz(TableBand2Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
        this.Set_TableBand2Vert(TableBand2Vert);
        this.Set_TableWholeTable(TableWholeTable);
    },

    Create_Table_List_6: function (oBorderUnifill, oBandUnifill, oTextUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTextUnifill,
			ThemeFill  : oTextUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Color: { r: 0x40, g: 0x40, b: 0x40 },
            FontSize: 11,
            Unifill: CellShd3.Unifill
        };

        var TableTextPr1 =
        {
            Unifill: CellShd3.Unifill,
            Bold: true
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Bottom: TableBorder1,
                Top: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr1,

            TableCellPr:
            {
                TableCellBorders:
                {
                    Bottom: TableBorder1
                }
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1
                }
            }
        };

        var TableCol =
        {
            TextPr: TableTextPr1
        };

        var TableBand1Horz =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        var TableBand1Vert =
        {
            TableCellPr:
            {
                Shd: CellShd2
            }
        };

        var TableBand2Horz =
        {
            TextPr: TableTextPr
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableCol);
        this.Set_TableLastCol(TableCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand2Horz(TableBand2Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
    },

    Create_Table_List_7: function (oBorderUnifill, oBandUnifill, oTableTextUnifill, oTextUnifill)
	{
        var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderUnifill,
			ThemeFill  : oBorderUnifill
		});

        var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBandUnifill,
			ThemeFill  : oBandUnifill
		});

        var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTableTextUnifill,
			ThemeFill  : oTableTextUnifill
		});

        var CellShd4 = new CDocumentShd();
		CellShd4.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oTextUnifill,
			ThemeFill  : oTextUnifill
		});

        var TableTextPr =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Unifill: CellShd4.Unifill,
            Italic: true,
            FontSize: 11
        };
        var TableTextPr1 =
        {
            RFonts: { Ascii: { Name: "Arial", Index: -1 }, HAnsi: { Name: "Arial", Index: -1 } },
            Unifill: CellShd4.Unifill,
            FontSize: 11
        };

        var ParaPr =
        {
            Spacing:
            {
                After: 0,
                Line: 1,
                LineRule: linerule_Auto
            },

            Borders:
            {
                Top:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Left:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Bottom:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Right:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                },

                Between:
                {
                    Color: { r: 0, g: 0, b: 0, Auto: true },
                    Space: 0,
                    Size: 0,
                    Value: border_None
                }
            }
        };

        var TableBorder1 =
        {
            Color: { r: 0, g: 0, b: 0 },
            Value: border_Single,
            Size: 0.5 * g_dKoef_pt_to_mm,
            Space: 0,
            Unifill: CellShd1.Unifill
        };

        var TableBorder2 =
        {
            Value: border_None
        };

        var TablePr =
        {
            TableStyleColBandSize: 1,
            TableStyleRowBandSize: 1,
            TableInd: 0,

            TableBorders:
            {
                Right: TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder1,
                    Left: TableBorder2
                },
                Shd: CellShd3
            }
        };

        var TableLastRow =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                TableCellBorders:
                {
                    Top: TableBorder1,
                    Bottom: TableBorder2,
                    Right: TableBorder2,
                    Left: TableBorder2
                },
                Shd: CellShd3
            }
        };

        var TableFirstCol =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd:
                {
                    Color:
                    {
                        r: 255,
                        g: 255,
                        b: 255
                    }
                },
                TableCellBorders:
                {
                    Left: TableBorder2,
                    Right: TableBorder1,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            },
            ParaPr:
            {
                Jc: align_Right
            }
        };

        var TableLastCol =
        {
            TextPr: TableTextPr,
            TableCellPr:
            {
                Shd:
                {
                    Color:
                    {
                        r: 255,
                        g: 255,
                        b: 255
                    }
                },
                TableCellBorders:
                {
                    Left: TableBorder1,
                    Right: TableBorder2,
                    Top: TableBorder2,
                    Bottom: TableBorder2
                }
            }
        };

        var TableBand1Horz =
        {
            TextPr: TableTextPr1,
            TableCellPr:
            {
                Shd: CellShd2
            }
        };
        var TableBand1Vert =
        {
            TableCellPr:
            {
                Shd: CellShd2
            }
        };
        var TableBand2Horz =
        {
            TextPr: TableTextPr1
        };

        var TableWholeTable =
        {
            TextPr: TableTextPr1
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand2Horz(TableBand2Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
        this.Set_TableWholeTable(TableWholeTable);
    },

    Create_Table_BorderedAndLined : function(oBorderFillUnifill, oHorBandUniFill, oVertBandUnifill,  oBorderUnifill)
    {
		var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oBorderFillUnifill,
			ThemeFill  : oBorderFillUnifill
		});

		var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oHorBandUniFill,
			ThemeFill  : oHorBandUniFill
		});

		var CellShd3 = new CDocumentShd();
		CellShd3.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : oVertBandUnifill,
			ThemeFill  : oVertBandUnifill
		});

        var TextPr1 =
        {
            RFonts   : { Ascii : { Name : "Arial", Index : -1 }, HAnsi : { Name : "Arial", Index : -1 } },
            Color    : { r : 0xF2, g : 0xF2, b : 0xF2 },
            FontSize : 11
        };

        var TextPr2 =
        {
            RFonts   : { Ascii : { Name : "Arial", Index : -1 }, HAnsi : { Name : "Arial", Index : -1 } },
            Color    : { r : 0x40, g : 0x40, b : 0x40 },
            FontSize : 11
        };

        var TableCellPr1 =
        {
			Shd : CellShd1
        };

        var TableCellPrVert =
        {
            Shd : CellShd2
        };

        var TableCellPrHorz =
        {
        	Shd : CellShd2
        };

        var ParaPr =
        {
            Spacing :
            {
                After    : 0,
                Line     : 1,
                LineRule : linerule_Auto
            },

            Borders :
            {
                Top :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Left :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Bottom :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Right :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Between :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                }
            }
        };

        var TextPr =
        {
            Color : { r : 0x40, g : 0x40, b : 0x40 }
        };

        var TableBorder1 =
        {
			Color : { r : 0, g : 0, b : 0 },
			Value : border_Single,
            Size  : 0.5 * g_dKoef_pt_to_mm,
            Space : 0,
			Unifill: oBorderUnifill
        };

        var TablePr =
        {
            TableStyleColBandSize : 1,
            TableStyleRowBandSize : 1,
            TableInd : 0,

            // TableCellMar :
            // {
            //     Top    : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
            //     Left   : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm),
            //     Bottom : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
            //     Right  : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm)
            // },

            TableBorders :
            {
                Top     : TableBorder1,
                Left    : TableBorder1,
                Bottom  : TableBorder1,
                Right   : TableBorder1,
                InsideH : TableBorder1,
                InsideV : TableBorder1
            }
        };

        var TableFirstRow =
        {
            TextPr      : TextPr1,
            TableCellPr : TableCellPr1
        };

        var TableLastRow =
        {
            TextPr      : TextPr1,
            TableCellPr : TableCellPr1
        };

        var TableFirstCol =
        {
            TextPr      : TextPr1,
            TableCellPr : TableCellPr1
        };

        var TableLastCol =
        {
            TextPr      : TextPr1,
            TableCellPr : TableCellPr1
        };

        var TableBand1Vert =
        {
            TextPr : TextPr2
        };

        var TableBand2Vert =
        {
            TextPr      : TextPr2,
            TableCellPr : TableCellPrVert
        };

        var TableBand1Horz =
        {
            TextPr : TextPr2
        };

        var TableBand2Horz =
        {
            TextPr      : TextPr2,
            TableCellPr : TableCellPrHorz
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TextPr(TextPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1Horz);
        this.Set_TableBand1Vert(TableBand1Vert);
        this.Set_TableBand2Horz(TableBand2Horz);
        this.Set_TableBand2Vert(TableBand2Vert);

	},

	Create_Grid_Table_Light : function(unifill1, unifill2)
    {
		var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : unifill1,
			ThemeFill  : unifill1
		});

		var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : unifill2,
			ThemeFill  : unifill2
		});

        var TableTextPr =
        {
            RFonts   : {Ascii : {Name : "Arial", Index : -1}, HAnsi : {Name : "Arial", Index : -1}},
            Color    : { r : 0x40, g : 0x40, b : 0x40 },
            FontSize : 11
        };

        var ParaPr =
        {
            Spacing :
            {
                After    : 0,
                Line     : 1,
                LineRule : linerule_Auto
            },

            Borders :
            {
                Top :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Left :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Bottom :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Right :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Between :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                }
            }
        };

        var TextPr =
        {
            Color : { r : 0x40, g : 0x40, b : 0x40 }
        };

        var TableBorder1 =
        {
			Color : { r : 0, g : 0, b : 0 },
            Value : border_Single,
            Size  : 0.5 * g_dKoef_pt_to_mm,
            Space : 0,
			Unifill: CellShd2.Unifill
        };

        var TableBorder2 =
        {
			Color : { r : 0, g : 0, b : 0 },
            Value : border_Single,
            Size  : 1.5 * g_dKoef_pt_to_mm,
            Space : 0,
			Unifill: CellShd1.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize : 1,
            TableStyleRowBandSize : 1,
            TableInd : 0,

            TableBorders :
            {
                Top     : TableBorder1,
                Left    : TableBorder1,
                Bottom  : TableBorder1,
                Right   : TableBorder1,
                InsideH : TableBorder1,
                InsideV : TableBorder1
            },

            // TableCellMar :
            // {
            //     Top    : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
            //     Left   : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm),
            //     Bottom : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
            //     Right  : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm)
            // }
        };

        var TableFirstRow =
        {
            TextPr : TableTextPr,

            TableCellPr :
            {
                TableCellBorders :
                {
                    Bottom : TableBorder2
                }
            }
        };

        var TableLastRow =
        {
            TextPr      : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Top : TableBorder2
                }
            }
        };

        var TableFirstCol =
        {
            TextPr      : TableTextPr,
            TableCellPr :
            {

            }
        };

        var TableLastCol =
        {
            TextPr      : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Left : TableBorder2
                }
            }
        };

        var TableBand1Horz =
        {
            TextPr : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Top     : TableBorder1,
                    Left    : TableBorder1,
                    Bottom  : TableBorder1,
                    Right   : TableBorder1,
                    InsideH : TableBorder1,
                    InsideV : TableBorder1
                }
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1Horz);

	},

	Create_Grid_Table : function(unifill1, unifill2)
    {
		var TextColor1 = new CDocumentColor(0xF2, 0xF2, 0xF2, false);
		var TextFont1  = { Name : "Arial", Index : -1 };
		var TextSize1  = 11;

		var CellShd1 = new CDocumentShd();
		CellShd1.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : unifill1,
			ThemeFill  : unifill1
		});

		var CellShd2 = new CDocumentShd();
		CellShd2.Set_FromObject({
			Value      : Asc.c_oAscShd.Clear,
			ThemeColor : unifill2,
			ThemeFill  : unifill2
		});

        var TableTextPr =
        {
            RFonts   : {Ascii : {Name : "Arial", Index : -1}, HAnsi : {Name : "Arial", Index : -1}},
            Color    : { r : 0x40, g : 0x40, b : 0x40 },
            FontSize : 11
        };

        var ParaPr =
        {
            Spacing :
            {
                After    : 0,
                Line     : 1,
                LineRule : linerule_Auto
            },

            Borders :
            {
                Top :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Left :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Bottom :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Right :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                },

                Between :
                {
                    Color : { r : 0, g : 0, b : 0, Auto : true },
                    Space : 0,
                    Size  : 0,
                    Value : border_None
                }
            }
        };

        var TextPr =
        {
            Color : { r : 0x40, g : 0x40, b : 0x40 }
        };

        var TableBorder1 =
        {
			Color : { r : 0, g : 0, b : 0 },
            Value : border_Single,
            Size  : 0.5 * g_dKoef_pt_to_mm,
            Space : 0,
			Unifill: CellShd2.Unifill
        };

        var TableBorder2 =
        {
			Color : { r : 0, g : 0, b : 0 },
            Value : border_Single,
            Size  : 1.5 * g_dKoef_pt_to_mm,
            Space : 0,
			Unifill: CellShd1.Unifill
        };

        var TablePr =
        {
            TableStyleColBandSize : 1,
            TableStyleRowBandSize : 1,
            TableInd : 0,

            TableBorders :
            {
                // Top     : TableBorder1,
                // Left    : TableBorder1,
                // Bottom  : TableBorder1,
                // Right   : TableBorder1,
                // InsideH : TableBorder1,
                // InsideV : TableBorder1
            },

            TableCellMar :
            {
                Top    : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
                Left   : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm),
                Bottom : new CTableMeasurement(tblwidth_Mm, 4.8 * g_dKoef_pt_to_mm),
                Right  : new CTableMeasurement(tblwidth_Mm, 8.5 * g_dKoef_pt_to_mm)
            }
        };

        var TableFirstRow =
        {
            TextPr : TableTextPr,

            TableCellPr :
            {
                TableCellBorders :
                {
                    Bottom : TableBorder2
                }
            }
        };

        var TableLastRow =
        {
            TextPr      : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Top : TableBorder2
                }
            }
        };

        var TableFirstCol =
        {
            TextPr      : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Right : TableBorder2
                }
            }
        };

        var TableLastCol =
        {
            TextPr      : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Left : TableBorder2
                }
            }
        };

        var TableBand1Horz =
        {
            TextPr : TableTextPr,
            TableCellPr :
            {
                TableCellBorders :
                {
                    Top     : TableBorder1,
                    Left    : TableBorder1,
                    Bottom  : TableBorder1,
                    Right   : TableBorder1,
                    InsideH : TableBorder1,
                    InsideV : TableBorder1
                }
            }
        };

        this.Set_UiPriority(99);
        this.Set_ParaPr(ParaPr);
        this.Set_TablePr(TablePr);
        this.Set_TableFirstRow(TableFirstRow);
        this.Set_TableLastRow(TableLastRow);
        this.Set_TableFirstCol(TableFirstCol);
        this.Set_TableLastCol(TableLastCol);
        this.Set_TableBand1Horz(TableBand1Horz);

    },

    isEqual: function(cStyles, isSimplifyName)
    {
        var result = false;
        if(this.BasedOn == cStyles.BasedOn && this.GetName(isSimplifyName) == cStyles.GetName(isSimplifyName) && this.Next == cStyles.Next && this.Type == cStyles.Type && this.hidden == cStyles.hidden)
        {
            if(this.qFormat == cStyles.qFormat && this.semiHidden == cStyles.semiHidden && this.uiPriority == cStyles.uiPriority && this.unhideWhenUsed == cStyles.unhideWhenUsed)
            {
                var isEqualParaPr = this.ParaPr.isEqual(this.ParaPr, cStyles.ParaPr);
                var isEqualTextPr = this.TextPr.isEqual(this.TextPr, cStyles.TextPr);
                if(isEqualParaPr && isEqualTextPr)
                    result = true;
            }
        }
        return result;
    },

    Is_Equal : function(oStyle)
    {
        if (oStyle.Name != this.Name
            || this.BasedOn !== oStyle.BasedOn
            || this.Next !== oStyle.Next
            || this.Type !== oStyle.Type
            || this.Link !== oStyle.Link
            || this.StyleId !== oStyle.StyleId
            || this.qFormat !== oStyle.qFormat
            || this.uiPriority !== oStyle.uiPriority
            || this.hidden !== oStyle.hidden
            || this.semiHidden !== oStyle.semiHidden
            || this.unhideWhenUsed !== oStyle.unhideWhenUsed
            || true !== this.TextPr.Is_Equal(oStyle.TextPr)
            || true !== this.ParaPr.Is_Equal(oStyle.ParaPr)
            || (styletype_Table === this.Type
                && (true !== this.TablePr.Is_Equal(oStyle.TablePr)
                    || true !== this.TableRowPr.Is_Equal(oStyle.TableRowPr)
                    || true !== this.TableCellPr.Is_Equal(oStyle.TableCellPr)
                    || true !== IsEqualStyleObjects(this.TableBand1Horz , oStyle.TableBand1Horz )
                    || true !== IsEqualStyleObjects(this.TableBand1Vert , oStyle.TableBand1Vert )
                    || true !== IsEqualStyleObjects(this.TableBand2Horz , oStyle.TableBand2Horz )
                    || true !== IsEqualStyleObjects(this.TableBand2Vert , oStyle.TableBand2Vert )
                    || true !== IsEqualStyleObjects(this.TableFirstCol  , oStyle.TableFirstCol  )
                    || true !== IsEqualStyleObjects(this.TableFirstRow  , oStyle.TableFirstRow  )
                    || true !== IsEqualStyleObjects(this.TableLastCol   , oStyle.TableLastCol   )
                    || true !== IsEqualStyleObjects(this.TableLastRow   , oStyle.TableLastRow   )
                    || true !== IsEqualStyleObjects(this.TableTLCell    , oStyle.TableTLCell    )
                    || true !== IsEqualStyleObjects(this.TableTRCell    , oStyle.TableTRCell    )
                    || true !== IsEqualStyleObjects(this.TableBLCell    , oStyle.TableBLCell    )
                    || true !== IsEqualStyleObjects(this.TableBRCell    , oStyle.TableBRCell    )
                    || true !== IsEqualStyleObjects(this.TableWholeTable, oStyle.TableWholeTable)
                    )
                )
            )
            return false;

        return true;
    },
//-----------------------------------------------------------------------------------
// Undo/Redo функции
//-----------------------------------------------------------------------------------
	GetSelectionState : function()
    {
    },

	SetSelectionState : function(State, StateIndex)
    {
    },

    Get_ParentObject_or_DocumentPos : function()
    {
        return { Type : AscDFH.historyitem_recalctype_Inline, Data : 0 };
    },

    Refresh_RecalcData : function(Data)
    {
        var Type = Data.Type;

        var bNeedRecalc = false;

        switch ( Type )
        {
            case AscDFH.historyitem_Style_TextPr          :
            case AscDFH.historyitem_Style_ParaPr          :
            case AscDFH.historyitem_Style_TablePr         :
            case AscDFH.historyitem_Style_TableRowPr      :
            case AscDFH.historyitem_Style_TableCellPr     :
            case AscDFH.historyitem_Style_TableBand1Horz  :
            case AscDFH.historyitem_Style_TableBand1Vert  :
            case AscDFH.historyitem_Style_TableBand2Horz  :
            case AscDFH.historyitem_Style_TableBand2Vert  :
            case AscDFH.historyitem_Style_TableFirstCol   :
            case AscDFH.historyitem_Style_TableFirstRow   :
            case AscDFH.historyitem_Style_TableLastCol    :
            case AscDFH.historyitem_Style_TableLastRow    :
            case AscDFH.historyitem_Style_TableTLCell     :
            case AscDFH.historyitem_Style_TableTRCell     :
            case AscDFH.historyitem_Style_TableBLCell     :
            case AscDFH.historyitem_Style_TableBRCell     :
            case AscDFH.historyitem_Style_TableWholeTable :
            case AscDFH.historyitem_Style_Name            :
            case AscDFH.historyitem_Style_BasedOn         :
            case AscDFH.historyitem_Style_Next            :
            case AscDFH.historyitem_Style_Type            :
            case AscDFH.historyitem_Style_QFormat         :
            case AscDFH.historyitem_Style_UiPriority      :
            case AscDFH.historyitem_Style_Hidden          :
            case AscDFH.historyitem_Style_SemiHidden      :
            case AscDFH.historyitem_Style_UnhideWhenUsed  :
            case AscDFH.historyitem_Style_Link            :
            {
                bNeedRecalc = true;
                break;
            }
        }

        if ( true === bNeedRecalc )
        {
            // Сообщаем родительскому классу, что изменения произошли в элементе с номером this.Index и на странице this.PageNum
            return this.Refresh_RecalcData2();
        }
    },

	Refresh_RecalcData2 : function()
	{
		var oHistory = History;
		if (!oHistory)
			return;

		if (!oHistory.AddChangedStyleToRecalculateData(this.Get_Id(), this))
			return;
	},

	RecalculateRelatedParagraphs : function()
	{
		var oHistory = History;
		if (!oHistory)
			return;

		var LogicDocument = private_GetWordLogicDocument();
		if(!LogicDocument)
			return;
		var Styles        = LogicDocument.Get_Styles();

		var AllParagraphs = [];

		if (this.Id != Styles.Default.Paragraph)
		{
			var AllStylesId = Styles.private_GetAllBasedStylesId(this.Id);
			AllParagraphs   = oHistory.GetAllParagraphsForRecalcData({Style : true, StylesId : AllStylesId});
			LogicDocument.Add_ChangedStyle(AllStylesId);
		}
		else
		{
			AllParagraphs = oHistory.GetAllParagraphsForRecalcData({All : true});
			LogicDocument.Add_ChangedStyle([this.Id]);
		}

		var Count = AllParagraphs.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var Para = AllParagraphs[Index];
			Para.Refresh_RecalcData({Type : AscDFH.historyitem_Paragraph_PStyle});
		}
	},
//-----------------------------------------------------------------------------------
// Функции для совместного редактирования
//-----------------------------------------------------------------------------------
    Write_ToBinary2 : function(Writer)
    {
        Writer.WriteLong(AscDFH.historyitem_type_Style);

        // String   : Id
        // Bool(und) -> Bool(null) -> String : Name
        // Bool(und) -> Bool(null) -> String : BasedOn
        // Bool(und) -> Bool(null) -> String : Next
        // Long : Type
        // Bool(und) -> Bool(null) -> Long : uiPriority
        // Bool(und) -> Bool(null) -> Bool : qFormat
        // Bool(und) -> Bool(null) -> Bool : hidden
        // Bool(und) -> Bool(null) -> Bool : semiHidden
        // Bool(und) -> Bool(null) -> Bool : unhideWhenUsed
        // Variable(CTextPr)      : TextPr
        // Variable(CParaPr)      : ParaPr
        // Variable(CTablePr)     : TablePr
        // Variable(CTableRowPr)  : TableRowPr
        // Variable(CTableCellPr) : TableCellPr
		// Bool : ConditionTableStyles?
        // Variable(CTableStylePr) : TableBand1Horz
        // Variable(CTableStylePr) : TableBand1Vert
        // Variable(CTableStylePr) : TableBand2Horz
        // Variable(CTableStylePr) : TableBand2Vert
        // Variable(CTableStylePr) : TableFirstCol
        // Variable(CTableStylePr) : TableFirstRow
        // Variable(CTableStylePr) : TableLastCol
        // Variable(CTableStylePr) : TableLastRow
        // Variable(CTableStylePr) : TableTLCell
        // Variable(CTableStylePr) : TableTRCell
        // Variable(CTableStylePr) : TableBLCell
        // Variable(CTableStylePr) : TableBRCell
        // Variable(CTableStylePr) : TableWholeTable

        Writer.WriteString2(this.Id);

        this.private_WriteUndefinedNullString(Writer, this.Name);
        this.private_WriteUndefinedNullString(Writer, this.BasedOn);
        this.private_WriteUndefinedNullString(Writer, this.Next);

        Writer.WriteLong(this.Type);

        this.private_WriterUndefinedNullLong(Writer, this.uiPriority);
        this.private_WriterUndefinedNullBool(Writer, this.qFormat);
        this.private_WriterUndefinedNullBool(Writer, this.hidden);
        this.private_WriterUndefinedNullBool(Writer, this.semiHidden);
        this.private_WriterUndefinedNullBool(Writer, this.unhideWhenUsed);

        this.TextPr.Write_ToBinary(Writer);
        this.ParaPr.Write_ToBinary(Writer);

        this.TablePr.Write_ToBinary(Writer);
        this.TableRowPr.Write_ToBinary(Writer);
        this.TableCellPr.Write_ToBinary(Writer);
		
		let conditionalTableStyles = !!this.TableBand1Horz;

		Writer.WriteBool(conditionalTableStyles);
		if (conditionalTableStyles)
		{
			this.TableBand1Horz.Write_ToBinary(Writer);
			this.TableBand1Vert.Write_ToBinary(Writer);
			this.TableBand2Horz.Write_ToBinary(Writer);
			this.TableBand2Vert.Write_ToBinary(Writer);
			this.TableFirstCol.Write_ToBinary(Writer);
			this.TableFirstRow.Write_ToBinary(Writer);
			this.TableLastCol.Write_ToBinary(Writer);
			this.TableLastRow.Write_ToBinary(Writer);
			this.TableTLCell.Write_ToBinary(Writer);
			this.TableTRCell.Write_ToBinary(Writer);
			this.TableBLCell.Write_ToBinary(Writer);
			this.TableBRCell.Write_ToBinary(Writer);
			this.TableWholeTable.Write_ToBinary(Writer);
		}
    },

    private_WriteUndefinedNullString : function(Writer, Value)
    {
        if (undefined === Value)
        {
            Writer.WriteBool(true);
        }
        else
        {
            Writer.WriteBool(false);
            if (null === Value)
            {
                Writer.WriteBool(true);
            }
            else
            {
                Writer.WriteBool(false);
                Writer.WriteString2(Value);
            }
        }
    },

    private_WriterUndefinedNullLong : function(Writer, Value)
    {
        if (undefined === Value)
        {
            Writer.WriteBool(true);
        }
        else
        {
            Writer.WriteBool(false);
            if (null === Value)
            {
                Writer.WriteBool(true);
            }
            else
            {
                Writer.WriteBool(false);
                Writer.WriteLong(Value);
            }
        }
    },

    private_WriterUndefinedNullBool : function(Writer, Value)
    {
        if (undefined === Value)
        {
            Writer.WriteBool(true);
        }
        else
        {
            Writer.WriteBool(false);
            if (null === Value)
            {
                Writer.WriteBool(true);
            }
            else
            {
                Writer.WriteBool(false);
                Writer.WriteBool(Value);
            }
        }
    },

    private_ReadUndefinedNullString : function(Reader)
    {
        if (true === Reader.GetBool())
            return undefined;
        else if (true === Reader.GetBool())
            return null;
        else
            return Reader.GetString2();
    },

    private_ReadUndefinedNullLong : function(Reader)
    {
        if (true === Reader.GetBool())
            return undefined;
        else if (true === Reader.GetBool())
            return null;
        else
            return Reader.GetLong();
    },

    private_ReadUndefinedNullBool : function(Reader)
    {
        if (true === Reader.GetBool())
            return undefined;
        else if (true === Reader.GetBool())
            return null;
        else
            return Reader.GetBool();
    },

    Read_FromBinary2 : function(Reader)
    {
        // String   : Id
        // Bool(und) -> Bool(null) -> String : Name
        // Bool(und) -> Bool(null) -> String : BasedOn
        // Bool(und) -> Bool(null) -> String : Next
        // Long : Type
        // Bool(und) -> Bool(null) -> Long : uiPriority
        // Bool(und) -> Bool(null) -> Bool : qFormat
        // Bool(und) -> Bool(null) -> Bool : hidden
        // Bool(und) -> Bool(null) -> Bool : semiHidden
        // Bool(und) -> Bool(null) -> Bool : unhideWhenUsed
        // Variable(CTextPr)      : TextPr
        // Variable(CParaPr)      : ParaPr
        // Variable(CTablePr)     : TablePr
        // Variable(CTableRowPr)  : TableRowPr
        // Variable(CTableCellPr) : TableCellPr
		// Bool : ConditionTableStyles?
        // Variable(CTableStylePr) : TableBand1Horz
        // Variable(CTableStylePr) : TableBand1Vert
        // Variable(CTableStylePr) : TableBand2Horz
        // Variable(CTableStylePr) : TableBand2Vert
        // Variable(CTableStylePr) : TableFirstCol
        // Variable(CTableStylePr) : TableFirstRow
        // Variable(CTableStylePr) : TableLastCol
        // Variable(CTableStylePr) : TableLastRow
        // Variable(CTableStylePr) : TableTLCell
        // Variable(CTableStylePr) : TableTRCell
        // Variable(CTableStylePr) : TableBLCell
        // Variable(CTableStylePr) : TableBRCell
        // Variable(CTableStylePr) : TableWholeTable

        this.Id = Reader.GetString2();

        this.Name    = this.private_ReadUndefinedNullString(Reader);
        this.BasedOn = this.private_ReadUndefinedNullString(Reader);
        this.Next    = this.private_ReadUndefinedNullString(Reader);

        this.Type = Reader.GetLong();

        this.uiPriority     = this.private_ReadUndefinedNullLong(Reader);
        this.qFormat        = this.private_ReadUndefinedNullBool(Reader);
        this.hidden         = this.private_ReadUndefinedNullBool(Reader);
        this.semiHidden     = this.private_ReadUndefinedNullBool(Reader);
        this.unhideWhenUsed = this.private_ReadUndefinedNullBool(Reader);

        this.TextPr.Read_FromBinary(Reader);
        this.ParaPr.Read_FromBinary(Reader);

        this.TablePr.Read_FromBinary(Reader);
        this.TableRowPr.Read_FromBinary(Reader);
        this.TableCellPr.Read_FromBinary(Reader);
		
		if (Reader.GetBool())
		{
			this.InitConditionalTableStyles();
			
			this.TableBand1Horz.Read_FromBinary(Reader);
			this.TableBand1Vert.Read_FromBinary(Reader);
			this.TableBand2Horz.Read_FromBinary(Reader);
			this.TableBand2Vert.Read_FromBinary(Reader);
			this.TableFirstCol.Read_FromBinary(Reader);
			this.TableFirstRow.Read_FromBinary(Reader);
			this.TableLastCol.Read_FromBinary(Reader);
			this.TableLastRow.Read_FromBinary(Reader);
			this.TableTLCell.Read_FromBinary(Reader);
			this.TableTRCell.Read_FromBinary(Reader);
			this.TableBLCell.Read_FromBinary(Reader);
			this.TableBRCell.Read_FromBinary(Reader);
			this.TableWholeTable.Read_FromBinary(Reader);
		}
    },
	
	Process_EndLoad : function(LinkData)
    {
        if (true === LinkData.StyleUpdate)
        {
            var LogicDocument = private_GetWordLogicDocument();
            if (!LogicDocument)
                return;

            var Styles = LogicDocument.Get_Styles();
            if (!Styles)
                return;

            var AllParagraphs = [];
            if (this.Id != Styles.Default.Paragraph)
            {
                var AllStylesId = Styles.private_GetAllBasedStylesId(this.Id);
                AllParagraphs = LogicDocument.GetAllParagraphsByStyle(AllStylesId);
                LogicDocument.Add_ChangedStyle(AllStylesId);
            }
            else
            {
                AllParagraphs = LogicDocument.GetAllParagraphs({All : true});
                LogicDocument.Add_ChangedStyle([this.Id]);
            }

            var Count = AllParagraphs.length;
            for ( var Index = 0; Index < Count; Index++ )
            {
                var Para = AllParagraphs[Index];
                Para.Recalc_CompiledPr();
                Para.Recalc_RunsCompiledPr();
            }
        }
    }
};
CStyle.fromObject = function(obj)
{
	if (!obj || !obj.Name)
		return null;
	
	let style = new CStyle(obj.Name, null, null, obj.Type, true);
	if (!obj)
		return style;
	
	style.SetFromObject(obj);
	return style;
};
CStyle.prototype.SetFromObject = function(obj)
{
	if (undefined !== obj.QFormat)
		this.SetQFormat(obj.QFormat);
	
	if (undefined !== obj.UiPriority)
		this.SetUiPriority(obj.UiPriority);
	
	if (undefined !== obj.Hidden)
		this.SetHidden(obj.Hidden);
	
	if (undefined !== obj.SemiHidden)
		this.SetSemiHidden(obj.SemiHidden);
	
	if (undefined !== obj.UnhideWhenUsed)
		this.SetUnhideWhenUsed(obj.UnhideWhenUsed);
	
	if (undefined !== obj.TextPr)
		this.SetTextPr(obj.TextPr);
	
	if (undefined !== obj.ParaPr)
		this.SetParaPr(obj.ParaPr);
	
	if (undefined !== obj.TablePr)
		this.Set_TablePr(obj.TablePr);
	
	if (undefined !== obj.TableRowPr)
		this.Set_TableRowPr(obj.TableRowPr);
	
	if (undefined !== obj.TableCellPr)
		this.Set_TableCellPr(obj.TableCellPr);
	
	// TODO: Table conditional
};
/**
 * @param {?CStyles} parent
 */
CStyle.prototype.SetParent = function(parent)
{
	this.Parent = parent;
};
CStyle.prototype.GetParent = function()
{
	return this.Parent;
};
/**
 * Получаем ссылку на основной класс документа
 * @returns {?AscWord.CDocument}
 */
CStyle.prototype.GetLogicDocument = function()
{
	if (!this.Parent)
		return null;
	
	return this.Parent.GetLogicDocument();
};
/**
 * Устанавливаем стиль, от которого данный наследуется
 * @param styleId
 */
CStyle.prototype.SetBasedOn = function(styleId)
{
	return this.Set_BasedOn(styleId);
};
/**
 * Получаем родительский стиль в иерархии наследования
 * @returns {null | string}
 */
CStyle.prototype.GetBasedOn = function()
{
	return this.BasedOn;
};
CStyle.prototype.SetNext = function(styleId)
{
	return this.Set_Next(styleId);
};
CStyle.prototype.GetNext = function()
{
	return this.Next;
};
CStyle.prototype.SetLink = function(styleId)
{
	return this.Set_Link(styleId);
};
CStyle.prototype.GetLink = function()
{
	return this.Link;
};
/**
 * Устаналиваем форматный идентификатор стиля
 * @param styleId
 * @constructor
 */
CStyle.prototype.SetStyleId = function(styleId)
{
	if (styleId === this.StyleId)
		return;

	History.Add(new CChangesStyleStyleId(this, this.StyleId, styleId));
	this.StyleId = styleId;
};
CStyle.prototype.GetStyleId = function()
{
	return this.StyleId;
};
/**
 * Выставляем текстовые настройки
 * @param {CTextPr | Object} oTextPr
 */
CStyle.prototype.SetTextPr = function(oTextPr)
{
	this.Set_TextPr(oTextPr);
};
/**
 * Получаем текстовые настройки
 * @returns {CTextPr}
 */
CStyle.prototype.GetTextPr = function()
{
	return this.TextPr;
};
/**
 * Выставляем настройки параграфа
 * @param {CParaPr | Object} oParaPr
 */
CStyle.prototype.SetParaPr = function(oParaPr)
{
	this.Set_ParaPr(oParaPr);
};
/**
 * Связываем данный стиль с заданной нумерацией
 * @param {string} numId Для удаления используем null
 * @param {number} iLvl
 */
CStyle.prototype.SetNumPr = function(numId, iLvl)
{
	let paraPr = this.GetParaPr().Copy();
	
	if (null !== numId)
		paraPr.NumPr = new AscWord.NumPr(numId, iLvl);
	else
		paraPr.NumPr = undefined;
	
	this.SetParaPr(paraPr);
};
/**
 * Получаем настройки параграфа
 * @returns {CParaPr}
 */
CStyle.prototype.GetParaPr = function()
{
	return this.ParaPr;
};
/**
 * Выставляем значение приоритета данного стиля
 * @param {number} nUiPriority
 */
CStyle.prototype.SetUiPriority = function(nUiPriority)
{
	this.Set_UiPriority(nUiPriority);
};
/**
 * Получаем значение приоритета данного стиля
 * @returns {number}
 */
CStyle.prototype.GetUiPriority = function()
{
	return this.uiPriority;
};
/**
 * Выставляем параметр примарность данного стиля
 * @param {boolean} isQFormat
 */
CStyle.prototype.SetQFormat = function(isQFormat)
{
	this.Set_QFormat(isQFormat);
};
/**
 * Получаем параметр примарности стиля
 * @returns {boolean}
 */
CStyle.prototype.GetQFormat = function()
{
	return this.qFormat;
};
/**
 * @param {boolean} isHidden
 */
CStyle.prototype.SetHidden = function(isHidden)
{
	this.Set_Hidden(isHidden);
};
/**
 * @returns {boolean}
 */
CStyle.prototype.GetHidden = function()
{
	return this.hidden;
};
/**
 * @param {boolean} isSemiHidden
 */
CStyle.prototype.SetSemiHidden = function(isSemiHidden)
{
	this.Set_SemiHidden(isSemiHidden);
};
/**
 * @returns {boolean}
 */
CStyle.prototype.GetSemiHidden = function()
{
	return this.semiHidden;
};
/**
 * Нужно ли показывать стиль в панели при его использовании
 * @param {boolean} isUnhide
 */
CStyle.prototype.SetUnhideWhenUsed = function(isUnhide)
{
	this.Set_UnhideWhenUsed(isUnhide);
};
/**
 * Нужно ли показывать стиль в панели при его использовании
 * @returns {boolean}
 */
CStyle.prototype.GetUnhideWhenUsed = function()
{
	return this.unhideWhenUsed;
};
CStyle.prototype.GetName = function(isSimplify)
{
	return this.Get_Name(isSimplify);
};
CStyle.prototype.GetBasedOn = function()
{
	return this.Get_BasedOn();
};
CStyle.prototype.GetNext = function()
{
	return this.Get_Next()
};
CStyle.prototype.GetType = function()
{
	return this.Get_Type();
};
/**
 * Выставляем линкованный стиль
 * @param {string} sLink
 */
CStyle.prototype.SetLink = function(sLink)
{
	this.Set_Link(sLink);
};
/**
 * Возвращаем идентефикатор линкованного стиля
 * @returns {null|string}
 */
CStyle.prototype.GetLink = function()
{
	return this.Link;
};
CStyle.prototype.IsEqual = function(oOtherStyle)
{
	return this.Is_Equal(oOtherStyle);
};
/**
 * Создаем слинкованный символьный шрифт
 * @param {string} sStyleName
 * @param {string} sBasedOn
 * @returns {CStyle}
 */
CStyle.prototype.CreateLinkedCharacterStyle = function(sStyleName, sBasedOn)
{
	var oStyle = new CStyle(sStyleName, sBasedOn, null, styletype_Character, true);

	oStyle.SetTextPr(this.GetTextPr());
	oStyle.SetUiPriority(this.GetUiPriority());
	oStyle.SetLink(this.GetId());

	this.SetLink(oStyle.GetId());

	return oStyle;
};
/**
 * Дефолтовые настройки для стиля Normal
 */
CStyle.prototype.CreateNormal = function()
{
	this.SetQFormat(true);
};
/**
 * Дефолтовые настройки для стиля Default Paragraph Font
 */
CStyle.prototype.CreateDefaultParagraphFont = function()
{
	this.SetUiPriority(1);
	this.SetSemiHidden(true);
	this.SetUnhideWhenUsed(true);
};
/**
 * Дефолтовые настройки для стиля No List
 */
CStyle.prototype.CreateNoList = function()
{
	this.SetUiPriority(99);
	this.SetSemiHidden(true);
	this.SetUnhideWhenUsed(true);
};
CStyle.prototype.CreateTOC = function(nLvl, nType)
{
	var ParaPr = {},
		TextPr = {};

	if (undefined === nType || null === nType || Asc.c_oAscTOCStylesType.Simple === nType)
	{
		ParaPr = {
			Spacing : {
				After : 57 / 20 * g_dKoef_pt_to_mm
			},

			Ind : {
				Left      : 0,
				Right     : 0,
				FirstLine : 0
			}
		};

		if (1 === nLvl)
			ParaPr.Ind.Left = 283 / 20 * g_dKoef_pt_to_mm;
		else if (2 === nLvl)
			ParaPr.Ind.Left = 567 / 20 * g_dKoef_pt_to_mm;
		else if (3 === nLvl)
			ParaPr.Ind.Left = 850 / 20 * g_dKoef_pt_to_mm;
		else if (4 === nLvl)
			ParaPr.Ind.Left = 1134 / 20 * g_dKoef_pt_to_mm;
		else if (5 === nLvl)
			ParaPr.Ind.Left = 1417 / 20 * g_dKoef_pt_to_mm;
		else if (6 === nLvl)
			ParaPr.Ind.Left = 1701 / 20 * g_dKoef_pt_to_mm;
		else if (7 === nLvl)
			ParaPr.Ind.Left = 1984 / 20 * g_dKoef_pt_to_mm;
		else if (8 === nLvl)
			ParaPr.Ind.Left = 2268 / 20 * g_dKoef_pt_to_mm;
	}
	else if (Asc.c_oAscTOCStylesType.Standard === nType)
	{
		ParaPr = {
			Spacing : {
				After : 57 / 20 * g_dKoef_pt_to_mm
			},

			Ind : {
				Left      : 0,
				Right     : 0,
				FirstLine : 0
			}
		};

		if (0 === nLvl)
		{
			TextPr.Bold     = true;
			TextPr.FontSize = 14;
		}
		else if (1 === nLvl)
		{
			ParaPr.Ind.Left = 283 / 20 * g_dKoef_pt_to_mm;
			TextPr.Bold     = true;
			TextPr.FontSize = 13;
		}
		else if (2 === nLvl)
		{
			ParaPr.Ind.Left = 567 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 13;
		}
		else if (3 === nLvl)
		{
			ParaPr.Ind.Left = 850 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 11;
		}
		else if (4 === nLvl)
		{
			ParaPr.Ind.Left = 1134 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 11;
		}
		else if (5 === nLvl)
		{
			ParaPr.Ind.Left = 1417 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 11;
		}
		else if (6 === nLvl)
		{
			ParaPr.Ind.Left = 1701 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 11;
		}
		else if (7 === nLvl)
		{
			ParaPr.Ind.Left = 1984 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 11;
		}
		else if (8 === nLvl)
		{
			ParaPr.Ind.Left = 2268 / 20 * g_dKoef_pt_to_mm;
			TextPr.FontSize = 11;
		}
	}
	else if (Asc.c_oAscTOCStylesType.Modern === nType)
	{
		ParaPr = {
			Ind : {
				Left      : 0,
				Right     : 0,
				FirstLine : 0
			}
		};

		if (0 === nLvl)
		{
			ParaPr.Spacing = {
				After : 170 / 20 * g_dKoef_pt_to_mm
			};

			ParaPr.Brd = {
				Bottom : {
					Color : {r : 0x00, g : 0x00, b : 0x00},
					Space : 0,
					Size  : 1 * g_dKoef_pt_to_mm,
					Value : border_Single
				}
			};


			TextPr.Bold     = true;
			TextPr.FontSize = 14;
		}
		else if (1 === nLvl)
		{
			ParaPr.Spacing = {
				After : 57 / 20 * g_dKoef_pt_to_mm
			};

			TextPr.Bold     = true;
			TextPr.FontSize = 13;
		}
		else if (2 === nLvl)
		{
			ParaPr.Spacing = {
				After : 57 / 20 * g_dKoef_pt_to_mm
			};

			TextPr.FontSize = 13;
		}
		else
		{
			ParaPr.Spacing = {
				After : 57 / 20 * g_dKoef_pt_to_mm
			};

			TextPr.FontSize = 11;
		}
	}
	else if (Asc.c_oAscTOCStylesType.Classic === nType)
	{
		ParaPr.Spacing = {
			After : 57 / 20 * g_dKoef_pt_to_mm
		};

		ParaPr.Ind = {
			Left      : 0,
			Right     : 0,
			FirstLine : 0
		};

		TextPr.FontSize = 11;

		if (0 === nLvl)
		{
			ParaPr.Spacing.After = 170 / 20 * g_dKoef_pt_to_mm;

			TextPr.Bold     = true;
			TextPr.FontSize = 14;
		}
		else if (1 === nLvl)
		{
			TextPr.Bold     = true;
			TextPr.FontSize = 13;
		}
		else if (2 === nLvl)
		{
			ParaPr.Ind.FirstLine = 283 / 20 * g_dKoef_pt_to_mm;

			TextPr.FontSize = 13;
		}
		else if (3 === nLvl)
		{
			ParaPr.Ind.FirstLine = 567 / 20 * g_dKoef_pt_to_mm;
		}
		else if (4 === nLvl)
		{
			ParaPr.Ind.FirstLine = 850 / 20 * g_dKoef_pt_to_mm;
		}
		else if (5 === nLvl)
		{
			ParaPr.Ind.FirstLine = 1134 / 20 * g_dKoef_pt_to_mm;
		}
		else if (6 === nLvl)
		{
			ParaPr.Ind.FirstLine = 1417 / 20 * g_dKoef_pt_to_mm;
		}
		else if (7 === nLvl)
		{
			ParaPr.Ind.FirstLine = 1701 / 20 * g_dKoef_pt_to_mm;
		}
		else if (8 === nLvl)
		{
			ParaPr.Ind.FirstLine = 1984 / 20 * g_dKoef_pt_to_mm;
		}
	}
	else if (Asc.c_oAscTOCStylesType.Web === nType)
	{
		ParaPr.Spacing = {
			After : 57 / 20 * g_dKoef_pt_to_mm
		};

		ParaPr.Ind = {};

		TextPr.Underline = true;
		TextPr.Color     = {r : 0x00, g : 0xC8, b : 0xC3};
		TextPr.Unifill   = AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorHyperlink, null, null);

		if (0 === nLvl)
		{
			ParaPr.Spacing.After = 100 / 20 * g_dKoef_pt_to_mm;
		}
		else if (1 === nLvl)
		{
			ParaPr.Spacing.After = 100 / 20 * g_dKoef_pt_to_mm;
			ParaPr.Ind.Left      = 220 / 20 * g_dKoef_pt_to_mm;
		}
		else if (2 === nLvl)
		{
			ParaPr.Spacing.After = 100 / 20 * g_dKoef_pt_to_mm;
			ParaPr.Ind.Left      = 440 / 20 * g_dKoef_pt_to_mm;
		}
		else if (3 === nLvl)
		{
			ParaPr.Ind.Left = 850 / 20 * g_dKoef_pt_to_mm;
		}
		else if (4 === nLvl)
		{
			ParaPr.Ind.Left = 1134 / 20 * g_dKoef_pt_to_mm;
		}
		else if (5 === nLvl)
		{
			ParaPr.Ind.Left = 1417 / 20 * g_dKoef_pt_to_mm;
		}
		else if (6 === nLvl)
		{
			ParaPr.Ind.Left = 1701 / 20 * g_dKoef_pt_to_mm;
		}
		else if (7 === nLvl)
		{
			ParaPr.Ind.Left = 1984 / 20 * g_dKoef_pt_to_mm;
		}
		else if (8 === nLvl)
		{
			ParaPr.Ind.Left = 2268 / 20 * g_dKoef_pt_to_mm;
		}
	}

	this.Set_UiPriority(39);
	this.Set_UnhideWhenUsed(true);
	this.Set_ParaPr(ParaPr);
	this.Set_TextPr(TextPr);
};
CStyle.prototype.CreateTOCHeading = function()
{
	var ParaPr = {};

	this.Set_UiPriority(39);
	this.Set_UnhideWhenUsed(true);
	this.Set_ParaPr(ParaPr);
};
CStyle.prototype.CreateTOF = function(nType)
{
	var nType_;
	var oParaPr = new CParaPr(), oTextPr = new CTextPr();
	oParaPr.Spacing.Set_FromObject({After: 0, AfterAutoSpacing: 0});
	if(Asc.c_oAscTOFStylesType)
	{
		if(nType === undefined || nType === null)
		{
			nType_ = Asc.c_oAscTOFStylesType.Formal;
		}
		else
		{
			nType_ = nType;
		}
		switch(nType_)
		{
			case Asc.c_oAscTOFStylesType.Classic:
			{
				oTextPr.SetCaps(true);
				break;
			}
			case Asc.c_oAscTOFStylesType.Distinctive:
			{
				oTextPr.SetItalic(true);
				break;
			}
			case Asc.c_oAscTOFStylesType.Centered:
			{
				oTextPr.SetBold(true);
				oTextPr.SetItalic(true);
				oParaPr.SetJc(AscCommon.align_Center);
				break;
			}
			case Asc.c_oAscTOFStylesType.Formal:
			{
				break;
			}
			case Asc.c_oAscTOFStylesType.Simple:
			{
				oTextPr.SetBold(true);
				break;
			}
			case Asc.c_oAscTOFStylesType.Web:
			{
				oTextPr.SetUnderline(true);
				oTextPr.SetUnifill(AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorHyperlink, null, null));
				break;
			}
		}
	}
	this.Set_UiPriority(99);
	this.Set_UnhideWhenUsed(true);
	this.Set_ParaPr(oParaPr);
	this.Set_TextPr(oTextPr);
};
/**
 * Конвертируем стиль в Asc.CAscStyle
 * @returns {Asc.CAscStyle}
 */
CStyle.prototype.ToAscStyle = function()
{
	var oAscStyle = new Asc.CAscStyle();

	oAscStyle.StyleId    = this.Id;

	oAscStyle.Name       = this.Name;
	oAscStyle.Type       = this.Type;
	oAscStyle.qFormat    = null === this.qFormat || undefined === this.qFormat ? false : this.qFormat;
	oAscStyle.uiPriority = null === this.qFormat || undefined === this.qFormat ? -1 : this.uiPriority;

	return oAscStyle;
};
/**
 * Получаем текстовые настройки
 * @returns {CTextPr}
 */
CStyle.prototype.GetTextPr = function()
{
	return this.TextPr;
};
/**
 * Получаем настройки для параграфа
 * @returns {CParaPr}
 */
CStyle.prototype.GetParaPr = function()
{
	return this.ParaPr;
};
/**
 * Показывать ли данный стиль в верхней панели
 * @param {CStyles} oStyles
 * @returns {boolean}
 */
CStyle.prototype.IsExpressStyle = function(oStyles)
{
	if (true === this.qFormat)
		return true;

	if (oStyles
		&& oStyles.Default
		&& (oStyles.Default.Header === this.Id
		|| oStyles.Default.Footer === this.Id
		|| oStyles.Default.FootnoteText === this.Id
		|| oStyles.Default.EndnoteText === this.Id))
		return true;

	return false;
};
/**
 * Выставляем является ли данный стиль кастомным
 * @param {boolean} isCustom
 */
CStyle.prototype.SetCustom = function(isCustom)
{
	History.Add(new CChangesStyleCustom(this, this.Name, isCustom));
	this.Custom = isCustom;
};
/**
 * Проверяем является ли данный стиль кастомным
 * @returns {boolean}
 */
CStyle.prototype.IsCustom = function()
{
	return this.Custom;
};
/**
 * Проверяем является ли данный стиль, стилем для параграфа
 * @returns {boolean}
 */
CStyle.prototype.IsParagraphStyle = function()
{
	return (this.Type === styletype_Paragraph);
};
CStyle.prototype.Document_Is_SelectionLocked = function(CheckType)
{
	switch ( CheckType )
	{
		case AscCommon.changestype_Paragraph_Content:
		case AscCommon.changestype_Paragraph_Properties:
		case AscCommon.changestype_Paragraph_AddText:
		case AscCommon.changestype_Paragraph_TextProperties:
		case AscCommon.changestype_ContentControl_Add:
		case AscCommon.changestype_Document_Content:
		case AscCommon.changestype_Document_Content_Add:
		case AscCommon.changestype_Image_Properties:
		case AscCommon.changestype_Remove:
		case AscCommon.changestype_Delete:
		case AscCommon.changestype_Document_SectPr:
		case AscCommon.changestype_Table_Properties:
		case AscCommon.changestype_Table_RemoveCells:
		case AscCommon.changestype_HdrFtr:
		{
			AscCommon.CollaborativeEditing.Add_CheckLock(true);
			break;
		}
	}
};
/**
 * @returns {boolean}
 */
CStyle.prototype.IsTableStyle = function()
{
	return (this.Type === styletype_Table);
};
/**
 * Задана ли в стиле нумерация
 * @returns {boolean}
 */
CStyle.prototype.HaveNumbering = function()
{
	return (!!this.ParaPr.NumPr);
};

CStyle.prototype.wholeToTablePr = function() {

	let oWhole = this.TableWholeTable;
	if(!oWhole) {
		return
	}
	let oWholeBorders = oWhole.TablePr && oWhole.TablePr.TableBorders;
	if(!oWholeBorders) {
		return;
	}
	let oWholeCellBorders = oWhole.TableCellPr && oWhole.TableCellPr.TableCellBorders;
	if(!oWholeCellBorders) {
		return;
	}

	let oTablePBorders = this.TablePr && this.TablePr.TableBorders;
	if(!oTablePBorders) {
		return;
	}

	if(oWholeBorders.InsideH) {
		oTablePBorders.InsideH = oWholeBorders.InsideH;
		delete oWholeBorders.InsideH;
	}
	if(oWholeBorders.InsideV) {
		oTablePBorders.InsideV = oWholeBorders.InsideV;
		delete oWholeBorders.InsideV;
	}
	if(oWholeCellBorders.Top) {
		oTablePBorders.Top = oWholeCellBorders.Top;
		delete oWholeCellBorders.Top;
	}
	if(oWholeCellBorders.Bottom) {
		oTablePBorders.Bottom = oWholeCellBorders.Bottom;
		delete oWholeCellBorders.Bottom;
	}
	if(oWholeCellBorders.Left) {
		oTablePBorders.Left = oWholeCellBorders.Left;
		delete oWholeCellBorders.Left;
	}
	if(oWholeCellBorders.Right) {
		oTablePBorders.Right = oWholeCellBorders.Right;
		delete oWholeCellBorders.Right;
	}
};
/**
 * Получаем список параграфов, использующих данный стиль, либо стиль, основанный на данном
 * @returns {AscWord.Paragraph[]}
 */
CStyle.prototype.GetRelatedParagraphs = function()
{
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument)
		return [];
	
	let styleManager = logicDocument.GetStyles();
	if (!styleManager)
		return [];
	
	let styleId = this.GetId();
	let paragraphs;
	if (styleId !== styleManager.GetDefaultParagraph())
	{
		let relatedStylesId = styleManager.private_GetAllBasedStylesId(styleId);
		paragraphs = logicDocument.GetAllParagraphsByStyle(relatedStylesId);
	}
	else
	{
		paragraphs = logicDocument.GetAllParagraphs();
	}
	
	return paragraphs;
};
/**
 * Обновляем коллекцию нумераций для параграфов, связанных с данным стилем
 */
CStyle.prototype.UpdateNumberingCollection = function()
{
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument)
		return;

	let numberingCollection = logicDocument.GetNumberingCollection();
	this.GetRelatedParagraphs().forEach(function(paragraph)
	{
		paragraph.RecalcCompiledPr();
		numberingCollection.CheckParagraph(paragraph);
	});
};
/**
 * Инициализируем условные табличные стили
 */
CStyle.prototype.InitConditionalTableStyles = function()
{
	this.TableBand1Horz  = new CTableStylePr();
	this.TableBand1Vert  = new CTableStylePr();
	this.TableBand2Horz  = new CTableStylePr();
	this.TableBand2Vert  = new CTableStylePr();
	this.TableFirstCol   = new CTableStylePr();
	this.TableFirstRow   = new CTableStylePr();
	this.TableLastCol    = new CTableStylePr();
	this.TableLastRow    = new CTableStylePr();
	this.TableTLCell     = new CTableStylePr();
	this.TableTRCell     = new CTableStylePr();
	this.TableBLCell     = new CTableStylePr();
	this.TableBRCell     = new CTableStylePr();
	this.TableWholeTable = new CTableStylePr();
};

function CStyles(bCreateDefault)
{
	this.ValidDefaultEastAsiaFont = false;

    if (bCreateDefault !== false)
    {
        this.Id = AscCommon.g_oIdCounter.Get_NewId();
        this.Lock = new AscCommon.CLock();

		this.Default = {
			ParaPr      : new CParaPr(),
			TextPr      : new CTextPr(),
			TablePr     : new CTablePr(),
			TableRowPr  : new CTableRowPr(),
			TableCellPr : new CTableCellPr(),

			Paragraph         : null,
			Character         : null,
			Numbering         : null,
			Table             : null,
			TableGrid         : null,
			Headings          : [],
			ParaList          : null,
			Header            : null,
			Footer            : null,
			Hyperlink         : null,
			FollowedHyperlink : null,
			FootnoteText      : null,
			FootnoteTextChar  : null,
			FootnoteReference : null,
			NoSpacing         : null,
			Title             : null,
			Subtitle          : null,
			Quote             : null,
			IntenseQuote      : null,
			TOC               : [],
			TOCHeading        : null,
			TOF               : null,
			Caption           : null,
			EndnoteText       : null,
			EndnoteTextChar   : null,
			EndnoteReference  : null,
			PlaceholderText   : null
		};

        // Заполняем значения по умолчанию
        this.Default.ParaPr.InitDefault();
        this.Default.TextPr.InitDefault();
        this.Default.TablePr.InitDefault();
        this.Default.TableRowPr.InitDefault();
        this.Default.TableCellPr.InitDefault();

        this.Style = [];
		
		// TODO: Сейчас почти все стили параграфа и текста заданы в AscWord.DEFAULT_STYLE_LIST
		//       Стили таблиц нужно переделать также
		
		// Создаем стандартный стиль для таблиц
		var Style_Table = new CStyle("Normal Table", null, null, styletype_Table);
		Style_Table.Create_NormalTable();
		this.Default.Table = this.Add(Style_Table);
		
        var fUF = AscCommonWord.CreateThemeUnifill;

		// Создаем стиль для таблиц, который будет применяться к новым таблицам
		var Style_TableGrid = new CStyle("Table Grid", this.Default.Table, null, styletype_Table);
		Style_TableGrid.Create_TableGrid();
		this.Add(Style_TableGrid);

        var Style_TableGridLight = new CStyle("Table Grid Light", this.Default.Table, null, styletype_Table);
		Style_TableGridLight.Create_TableGrid_Light(fUF(EThemeColor.themecolorText1, 0x50, null));
        this.Add(Style_TableGridLight);

        var Style_Plain_Table_1 = new CStyle("Plain Table 1", this.Default.Table, null, styletype_Table);
		Style_Plain_Table_1.Create_Table_Plain_1(fUF(EThemeColor.themecolorText1, 0x50, null), fUF(EThemeColor.themecolorText1, 0x0D, null));
        this.Add(Style_Plain_Table_1);

        var Style_Plain_Table_2 = new CStyle("Plain Table 2", this.Default.Table, null, styletype_Table);
		Style_Plain_Table_2.Create_Table_Plain_2( fUF(EThemeColor.themecolorText1, null, null));
        this.Add(Style_Plain_Table_2);

        var Style_Plain_Table_3 = new CStyle("Plain Table 3", this.Default.Table, null, styletype_Table);
		Style_Plain_Table_3.Create_Table_Plain_3(fUF(EThemeColor.themecolorText1, null, null), fUF(EThemeColor.themecolorText1, 0x0D, null));
        this.Add(Style_Plain_Table_3);

        var Style_Plain_Table_4 = new CStyle("Plain Table 4", this.Default.Table, null, styletype_Table);
		Style_Plain_Table_4.Create_Table_Plain_4( fUF(EThemeColor.themecolorText1, 0x0D, null));
        this.Add(Style_Plain_Table_4);

        var Style_Plain_Table_5 = new CStyle("Plain Table 5", this.Default.Table, null, styletype_Table);
		Style_Plain_Table_5.Create_Table_Plain_5(fUF(EThemeColor.themecolorNone, null, null), fUF(EThemeColor.themecolorText1, 0x0D, null));
        this.Add(Style_Plain_Table_5);

        /*
         // Создаем стандартный стиль для таблиц
         var Style_Table = new CStyle("LightShading", this.Default.Table, null, styletype_Table );
         Style_Table.Create_Table_LightShading();
         this.Add( Style_Table );

         // Создаем стандартный стиль для таблиц
         var Style_Table = new CStyle("ColorfulListAccent6", this.Default.Table, null, styletype_Table );
         Style_Table.Create_Table_ColorfulListAccent6();
         this.Add( Style_Table );
         */


        var Style_Table_Grid_1_Accent = new CStyle("Grid Table 1 Light", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent.Create_Table_Grid_1(fUF(EThemeColor.themecolorText1, 0x95, null),  fUF(EThemeColor.themecolorText1, 0x67, null));
		this.Add( Style_Table_Grid_1_Accent );

		var Style_Table_Grid_1_Accent_1 = new CStyle("Grid Table 1 Light - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent_1.Create_Table_Grid_1(fUF(EThemeColor.themecolorAccent1, 0x95, null), fUF(EThemeColor.themecolorAccent1, 0x67, null));
		this.Add( Style_Table_Grid_1_Accent_1 );

		var Style_Table_Grid_1_Accent_2 = new CStyle("Grid Table 1 Light - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent_2.Create_Table_Grid_1(fUF(EThemeColor.themecolorAccent2, 0x95, null), fUF(EThemeColor.themecolorAccent2, 0x67, null));
		this.Add( Style_Table_Grid_1_Accent_2 );

		var Style_Table_Grid_1_Accent_3 = new CStyle("Grid Table 1 Light - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent_3.Create_Table_Grid_1(fUF(EThemeColor.themecolorAccent3, 0x95, null), fUF(EThemeColor.themecolorAccent3, 0x67, null));
		this.Add( Style_Table_Grid_1_Accent_3 );

		var Style_Table_Grid_1_Accent_4 = new CStyle("Grid Table 1 Light - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent_4.Create_Table_Grid_1(fUF(EThemeColor.themecolorAccent4, 0x95, null), fUF(EThemeColor.themecolorAccent4, 0x67, null));
		this.Add( Style_Table_Grid_1_Accent_4 );

		var Style_Table_Grid_1_Accent_5 = new CStyle("Grid Table 1 Light - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent_5.Create_Table_Grid_1(fUF(EThemeColor.themecolorAccent5, 0x95, null), fUF(EThemeColor.themecolorAccent5, 0x67, null));
		this.Add( Style_Table_Grid_1_Accent_5 );

		var Style_Table_Grid_1_Accent_6 = new CStyle("Grid Table 1 Light - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_1_Accent_6.Create_Table_Grid_1(fUF(EThemeColor.themecolorAccent6, 0x95, null), fUF(EThemeColor.themecolorAccent6, 0x67, null));
        this.Add( Style_Table_Grid_1_Accent_6 );

		var Style_Table_Grid_Accent = new CStyle("Grid Table 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent.Create_Table_Grid_2(fUF(EThemeColor.themecolorText1, 0x95, null), fUF(EThemeColor.themecolorText1, 0x34, null));
		this.Add( Style_Table_Grid_Accent);

		var Style_Table_Grid_Accent_1 = new CStyle("Grid Table 2 - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_1.Create_Table_Grid_2(fUF(EThemeColor.themecolorAccent1, 0xEA, null), fUF(EThemeColor.themecolorAccent1, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_1 );

		var Style_Table_Grid_Accent_2 = new CStyle("Grid Table 2 - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_2.Create_Table_Grid_2(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null),fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_2 );

		var Style_Table_Grid_Accent_3 = new CStyle("Grid Table 2 - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_3.Create_Table_Grid_2(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_3 );

		var Style_Table_Grid_Accent_4 = new CStyle("Grid Table 2 - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_4.Create_Table_Grid_2(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_4 );

		var Style_Table_Grid_Accent_5 = new CStyle("Grid Table 2 - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_5.Create_Table_Grid_2(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_5 );

		var Style_Table_Grid_Accent_6 = new CStyle("Grid Table 2 - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_6.Create_Table_Grid_2(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_6 );

		var Style_Table_Grid_Accent = new CStyle("Grid Table 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent.Create_Table_Grid_3(fUF(EThemeColor.themecolorText1, 0x95, null), fUF(EThemeColor.themecolorText1, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent);

		var Style_Table_Grid_Accent_1 = new CStyle("Grid Table 3 - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_1.Create_Table_Grid_3(fUF(EThemeColor.themecolorAccent1, 0xEA, null), fUF(EThemeColor.themecolorAccent1, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_1 );

		var Style_Table_Grid_Accent_2 = new CStyle("Grid Table 3 - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_2.Create_Table_Grid_3(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_2 );

		var Style_Table_Grid_Accent_3 = new CStyle("Grid Table 3 - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_3.Create_Table_Grid_3(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_3 );

		var Style_Table_Grid_Accent_4 = new CStyle("Grid Table 3 - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_4.Create_Table_Grid_3(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_4 );

		var Style_Table_Grid_Accent_5 = new CStyle("Grid Table 3 - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_5.Create_Table_Grid_3(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_5 );

		var Style_Table_Grid_Accent_6 = new CStyle("Grid Table 3 - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_6.Create_Table_Grid_3(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_Grid_Accent_6 );

		var Style_Table_Grid_Accent = new CStyle("Grid Table 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent.Create_Table_Grid_4(fUF(EThemeColor.themecolorText1, null, null), fUF(EThemeColor.themecolorText1, 0x34, null), fUF(EThemeColor.themecolorText1, 0x90, null));
		this.Add( Style_Table_Grid_Accent);

		var Style_Table_Grid_Accent_1 = new CStyle("Grid Table 4 - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_1.Create_Table_Grid_4(fUF(EThemeColor.themecolorAccent1, 0xEA, null), fUF(EThemeColor.themecolorAccent1, 0x32, null), fUF(EThemeColor.themecolorAccent1, 0x90, null));
		this.Add( Style_Table_Grid_Accent_1 );

		var Style_Table_Grid_Accent_2 = new CStyle("Grid Table 4 - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_2.Create_Table_Grid_4(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorAccent2, 0x90, null));
		this.Add( Style_Table_Grid_Accent_2 );

		var Style_Table_Grid_Accent_3 = new CStyle("Grid Table 4 - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_3.Create_Table_Grid_4(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorAccent3, 0x90, null));
		this.Add( Style_Table_Grid_Accent_3 );

		var Style_Table_Grid_Accent_4 = new CStyle("Grid Table 4 - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_4.Create_Table_Grid_4(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorAccent4, 0x90, null));
		this.Add( Style_Table_Grid_Accent_4 );

		var Style_Table_Grid_Accent_5 = new CStyle("Grid Table 4 - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_5.Create_Table_Grid_4(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorAccent5, 0x90, null));
		this.Add( Style_Table_Grid_Accent_5 );

		var Style_Table_Grid_Accent_6 = new CStyle("Grid Table 4 - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_6.Create_Table_Grid_4(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorAccent6, 0x90, null));
		this.Add( Style_Table_Grid_Accent_6 );

		var Style_Table_Grid_Accent = new CStyle("Grid Table 5 Dark", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent.Create_Table_Grid_5(fUF(EThemeColor.themecolorText1, null, null), fUF(EThemeColor.themecolorText1, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorText1, 0x75, null));
		this.Add( Style_Table_Grid_Accent);

		var Style_Table_Grid_Accent_1 = new CStyle("Grid Table 5 Dark- Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_1.Create_Table_Grid_5(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent1, 0x75, null));
		this.Add( Style_Table_Grid_Accent_1 );

		var Style_Table_Grid_Accent_2 = new CStyle("Grid Table 5 Dark - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_2.Create_Table_Grid_5(fUF(EThemeColor.themecolorAccent2, null, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent2, 0x75, null));
		this.Add( Style_Table_Grid_Accent_2 );

		var Style_Table_Grid_Accent_3 = new CStyle("Grid Table 5 Dark - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_3.Create_Table_Grid_5(fUF(EThemeColor.themecolorAccent3, null, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent3, 0x75, null));
		this.Add( Style_Table_Grid_Accent_3 );

		var Style_Table_Grid_Accent_4 = new CStyle("Grid Table 5 Dark- Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_4.Create_Table_Grid_5(fUF(EThemeColor.themecolorAccent4, null, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent4, 0x75, null));
		this.Add( Style_Table_Grid_Accent_4 );

		var Style_Table_Grid_Accent_5 = new CStyle("Grid Table 5 Dark - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_5.Create_Table_Grid_5(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent5, 0x75, null));
		this.Add( Style_Table_Grid_Accent_5 );

		var Style_Table_Grid_Accent_6 = new CStyle("Grid Table 5 Dark - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_6.Create_Table_Grid_5(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent6, 0x75, null));
		this.Add( Style_Table_Grid_Accent_6 );

		var Style_Table_Grid_Accent = new CStyle("Grid Table 6 Colorful", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent.Create_Table_Grid_6(fUF(EThemeColor.themecolorText1, 0x80, null), fUF(EThemeColor.themecolorText1, 0x34, null), fUF(EThemeColor.themecolorText1, 0x80, 0x95));
		this.Add( Style_Table_Grid_Accent);

		var Style_Table_Grid_Accent_1 = new CStyle("Grid Table 6 Colorful - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_1.Create_Table_Grid_6(fUF(EThemeColor.themecolorAccent1, 0x80, null), fUF(EThemeColor.themecolorAccent1, 0x34, null), fUF(EThemeColor.themecolorAccent1, 0x80, 0x95));
		this.Add( Style_Table_Grid_Accent_1 );

		var Style_Table_Grid_Accent_2 = new CStyle("Grid Table 6 Colorful - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_2.Create_Table_Grid_6(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorAccent2, 0x97, 0x95));
		this.Add( Style_Table_Grid_Accent_2 );

		var Style_Table_Grid_Accent_3 = new CStyle("Grid Table 6 Colorful - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_3.Create_Table_Grid_6(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorAccent3, 0xFE, 0x95));
		this.Add( Style_Table_Grid_Accent_3 );

		var Style_Table_Grid_Accent_4 = new CStyle("Grid Table 6 Colorful - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_4.Create_Table_Grid_6(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorAccent4, 0x9A, 0x95));
		this.Add( Style_Table_Grid_Accent_4 );

		var Style_Table_Grid_Accent_5 = new CStyle("Grid Table 6 Colorful - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_5.Create_Table_Grid_6(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorAccent5, null, 0x95));
		this.Add( Style_Table_Grid_Accent_5 );

		var Style_Table_Grid_Accent_6 = new CStyle("Grid Table 6 Colorful - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_6.Create_Table_Grid_6(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null),  fUF(EThemeColor.themecolorAccent5, null, 0x95));
		this.Add( Style_Table_Grid_Accent_6 );

		var Style_Table_Grid_Accent = new CStyle("Grid Table 7 Colorful", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent.Create_Table_Grid_7(fUF(EThemeColor.themecolorText1, 0x80, null), fUF(EThemeColor.themecolorText1, 0x0D, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorText1, 0x80, 0x95));
		this.Add( Style_Table_Grid_Accent);

		var Style_Table_Grid_Accent_1 = new CStyle("Grid Table 7 Colorful - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_1.Create_Table_Grid_7(fUF(EThemeColor.themecolorAccent1, 0x80, null), fUF(EThemeColor.themecolorAccent1, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent1, 0x80, 0x95));
		this.Add( Style_Table_Grid_Accent_1 );

		var Style_Table_Grid_Accent_2 = new CStyle("Grid Table 7 Colorful - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_2.Create_Table_Grid_7(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent2, 0x97, 0x95));
		this.Add( Style_Table_Grid_Accent_2 );

		var Style_Table_Grid_Accent_3 = new CStyle("Grid Table 7 Colorful - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_3.Create_Table_Grid_7(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent3, 0xFE, 0x95));
		this.Add( Style_Table_Grid_Accent_3 );

		var Style_Table_Grid_Accent_4 = new CStyle("Grid Table 7 Colorful - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_4.Create_Table_Grid_7(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent4, 0x9A, 0x95));
		this.Add( Style_Table_Grid_Accent_4 );

		var Style_Table_Grid_Accent_5 = new CStyle("Grid Table 7 Colorful - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_5.Create_Table_Grid_7(fUF(EThemeColor.themecolorAccent5, 0x90, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent5, null, 0x95));
		this.Add( Style_Table_Grid_Accent_5 );

		var Style_Table_Grid_Accent_6 = new CStyle("Grid Table 7 Colorful - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Grid_Accent_6.Create_Table_Grid_7(fUF(EThemeColor.themecolorAccent6, 0x90, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent6, null, 0x95));
        this.Add( Style_Table_Grid_Accent_6 );

        var Style_Table_List_1_Accent = new CStyle("List Table 1 Light", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent.Create_Table_List_1(fUF(EThemeColor.themecolorText1, null, null),  fUF(EThemeColor.themecolorText1, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_List_1_Accent );

		var Style_Table_List_1_Accent_1 = new CStyle("List Table 1 Light - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent_1.Create_Table_List_1(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_List_1_Accent_1 );

		var Style_Table_List_1_Accent_2 = new CStyle("List Table 1 Light - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent_2.Create_Table_List_1(fUF(EThemeColor.themecolorAccent2, null, null), fUF(EThemeColor.themecolorAccent2, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_List_1_Accent_2 );

		var Style_Table_List_1_Accent_3 = new CStyle("List Table 1 Light - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent_3.Create_Table_List_1(fUF(EThemeColor.themecolorAccent3, null, null), fUF(EThemeColor.themecolorAccent3, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_List_1_Accent_3 );

		var Style_Table_List_1_Accent_4 = new CStyle("List Table 1 Light - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent_4.Create_Table_List_1(fUF(EThemeColor.themecolorAccent4, null, null), fUF(EThemeColor.themecolorAccent4, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_List_1_Accent_4 );

		var Style_Table_List_1_Accent_5 = new CStyle("List Table 1 Light - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent_5.Create_Table_List_1(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
		this.Add( Style_Table_List_1_Accent_5 );

		var Style_Table_List_1_Accent_6 = new CStyle("List Table 1 Light - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_1_Accent_6.Create_Table_List_1(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x40, null), fUF(EThemeColor.themecolorNone, null, null));
        this.Add( Style_Table_List_1_Accent_6 );

        var Style_Table_List_2_Accent = new CStyle("List Table 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent.Create_Table_List_2(fUF(EThemeColor.themecolorText1, 0x90, null),  fUF(EThemeColor.themecolorText1, 0x40, null));
		this.Add( Style_Table_List_2_Accent );

		var Style_Table_List_2_Accent_1 = new CStyle("List Table 2 - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent_1.Create_Table_List_2(fUF(EThemeColor.themecolorAccent1, 0x90, null), fUF(EThemeColor.themecolorAccent1, 0x40, null));
		this.Add( Style_Table_List_2_Accent_1 );

		var Style_Table_List_2_Accent_2 = new CStyle("List Table 2 - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent_2.Create_Table_List_2(fUF(EThemeColor.themecolorAccent2, 0x90, null), fUF(EThemeColor.themecolorAccent2, 0x40, null));
		this.Add( Style_Table_List_2_Accent_2 );

		var Style_Table_List_2_Accent_3 = new CStyle("List Table 2 - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent_3.Create_Table_List_2(fUF(EThemeColor.themecolorAccent3, 0x90, null), fUF(EThemeColor.themecolorAccent3, 0x40, null));
		this.Add( Style_Table_List_2_Accent_3 );

		var Style_Table_List_2_Accent_4 = new CStyle("List Table 2 - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent_4.Create_Table_List_2(fUF(EThemeColor.themecolorAccent4, 0x90, null), fUF(EThemeColor.themecolorAccent4, 0x40, null));
		this.Add( Style_Table_List_2_Accent_4 );

		var Style_Table_List_2_Accent_5 = new CStyle("List Table 2 - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent_5.Create_Table_List_2(fUF(EThemeColor.themecolorAccent5, 0x90, null), fUF(EThemeColor.themecolorAccent5, 0x40, null));
		this.Add( Style_Table_List_2_Accent_5 );

		var Style_Table_List_2_Accent_6 = new CStyle("List Table 2 - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_2_Accent_6.Create_Table_List_2(fUF(EThemeColor.themecolorAccent6, 0x90, null), fUF(EThemeColor.themecolorAccent6, 0x40, null));
        this.Add( Style_Table_List_2_Accent_6 );

        var Style_Table_List_3_Accent = new CStyle("List Table 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent.Create_Table_List_3(fUF(EThemeColor.themecolorText1, null, null),  fUF(EThemeColor.themecolorText1, 0x26, null));
		this.Add( Style_Table_List_3_Accent );

		var Style_Table_List_3_Accent_1 = new CStyle("List Table 3 - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent_1.Create_Table_List_3(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x67, null));
		this.Add( Style_Table_List_3_Accent_1 );

		var Style_Table_List_3_Accent_2 = new CStyle("List Table 3 - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent_2.Create_Table_List_3(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x67, null));
		this.Add( Style_Table_List_3_Accent_2 );

		var Style_Table_List_3_Accent_3 = new CStyle("List Table 3 - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent_3.Create_Table_List_3(fUF(EThemeColor.themecolorAccent3, 0x98, null), fUF(EThemeColor.themecolorAccent3, 0x67, null));
		this.Add( Style_Table_List_3_Accent_3 );

		var Style_Table_List_3_Accent_4 = new CStyle("List Table 3 - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent_4.Create_Table_List_3(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x67, null));
		this.Add( Style_Table_List_3_Accent_4 );

		var Style_Table_List_3_Accent_5 = new CStyle("List Table 3 - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent_5.Create_Table_List_3(fUF(EThemeColor.themecolorAccent5, 0x9A, null), fUF(EThemeColor.themecolorAccent5, 0x67, null));
		this.Add( Style_Table_List_3_Accent_5 );

		var Style_Table_List_3_Accent_6 = new CStyle("List Table 3 - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_3_Accent_6.Create_Table_List_3(fUF(EThemeColor.themecolorAccent6, 0x98, null), fUF(EThemeColor.themecolorAccent6, 0x67, null));
        this.Add( Style_Table_List_3_Accent_6 );

        var Style_Table_List_4_Accent = new CStyle("List Table 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent.Create_Table_List_4(fUF(EThemeColor.themecolorText1, null, null),  fUF(EThemeColor.themecolorText1, 0x40, null), fUF(EThemeColor.themecolorText1, null, null));
		this.Add( Style_Table_List_4_Accent );

		var Style_Table_List_4_Accent_1 = new CStyle("List Table 4 - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent_1.Create_Table_List_4(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x40, null), fUF(EThemeColor.themecolorAccent1, 0x90, null));
		this.Add( Style_Table_List_4_Accent_1 );

		var Style_Table_List_4_Accent_2 = new CStyle("List Table 4 - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent_2.Create_Table_List_4(fUF(EThemeColor.themecolorAccent2, null, null), fUF(EThemeColor.themecolorAccent2, 0x40, null), fUF(EThemeColor.themecolorAccent2, 0x90, null));
		this.Add( Style_Table_List_4_Accent_2 );

		var Style_Table_List_4_Accent_3 = new CStyle("List Table 4 - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent_3.Create_Table_List_4(fUF(EThemeColor.themecolorAccent3, null, null), fUF(EThemeColor.themecolorAccent3, 0x40, null), fUF(EThemeColor.themecolorAccent3, 0x90, null));
		this.Add( Style_Table_List_4_Accent_3 );

		var Style_Table_List_4_Accent_4 = new CStyle("List Table 4 - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent_4.Create_Table_List_4(fUF(EThemeColor.themecolorAccent4, null, null), fUF(EThemeColor.themecolorAccent4, 0x40, null), fUF(EThemeColor.themecolorAccent4, 0x90, null));
		this.Add( Style_Table_List_4_Accent_4 );

		var Style_Table_List_4_Accent_5 = new CStyle("List Table 4 - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent_5.Create_Table_List_4(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x40, null), fUF(EThemeColor.themecolorAccent5, 0x90, null));
		this.Add( Style_Table_List_4_Accent_5 );

		var Style_Table_List_4_Accent_6 = new CStyle("List Table 4 - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_4_Accent_6.Create_Table_List_4(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x40, null), fUF(EThemeColor.themecolorAccent6, 0x90, null));
        this.Add( Style_Table_List_4_Accent_6 );

        var Style_Table_List_5_Accent = new CStyle("List Table 5 Dark", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent.Create_Table_List_5(fUF(EThemeColor.themecolorText1, 0x80, null), fUF(EThemeColor.themecolorLight1, null, null));
		this.Add( Style_Table_List_5_Accent );

		var Style_Table_List_5_Accent_1 = new CStyle("List Table 5 Dark - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent_1.Create_Table_List_5(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorLight1, null, null));
		this.Add( Style_Table_List_5_Accent_1 );

		var Style_Table_List_5_Accent_2 = new CStyle("List Table 5 Dark - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent_2.Create_Table_List_5(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorLight1, null, null));
		this.Add( Style_Table_List_5_Accent_2 );

		var Style_Table_List_5_Accent_3 = new CStyle("List Table 5 Dark - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent_3.Create_Table_List_5(fUF(EThemeColor.themecolorAccent3, 0x98, null), fUF(EThemeColor.themecolorLight1, null, null));
		this.Add( Style_Table_List_5_Accent_3 );

		var Style_Table_List_5_Accent_4 = new CStyle("List Table 5 Dark - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent_4.Create_Table_List_5(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorLight1, null, null));
		this.Add( Style_Table_List_5_Accent_4 );

		var Style_Table_List_5_Accent_5 = new CStyle("List Table 5 Dark - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent_5.Create_Table_List_5(fUF(EThemeColor.themecolorAccent5, 0x9A, null), fUF(EThemeColor.themecolorLight1, null, null));
		this.Add( Style_Table_List_5_Accent_5 );

		var Style_Table_List_5_Accent_6 = new CStyle("List Table 5 Dark - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_5_Accent_6.Create_Table_List_5(fUF(EThemeColor.themecolorAccent6, 0x98, null), fUF(EThemeColor.themecolorLight1, null, null));
        this.Add( Style_Table_List_5_Accent_6 );

        var Style_Table_List_6_Accent = new CStyle("List Table 6 Colorful", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent.Create_Table_List_6(fUF(EThemeColor.themecolorText1, 0x80, null),  fUF(EThemeColor.themecolorText1, 0x40, null), fUF(EThemeColor.themecolorText1, null, null));
		this.Add( Style_Table_List_6_Accent );

		var Style_Table_List_6_Accent_1 = new CStyle("List Table 6 Colorful - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent_1.Create_Table_List_6(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x40, null), fUF(EThemeColor.themecolorAccent1, null, 0x95));
		this.Add( Style_Table_List_6_Accent_1 );

		var Style_Table_List_6_Accent_2 = new CStyle("List Table 6 Colorful - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent_2.Create_Table_List_6(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x40, null), fUF(EThemeColor.themecolorAccent2, 0x97, 0x95));
		this.Add( Style_Table_List_6_Accent_2 );

		var Style_Table_List_6_Accent_3 = new CStyle("List Table 6 Colorful - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent_3.Create_Table_List_6(fUF(EThemeColor.themecolorAccent3, 0x98, null), fUF(EThemeColor.themecolorAccent3, 0x40, null), fUF(EThemeColor.themecolorAccent3, 0x98, 0x95));
		this.Add( Style_Table_List_6_Accent_3 );

		var Style_Table_List_6_Accent_4 = new CStyle("List Table 6 Colorful - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent_4.Create_Table_List_6(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x40, null), fUF(EThemeColor.themecolorAccent4, 0x9A, 0x95));
		this.Add( Style_Table_List_6_Accent_4 );

		var Style_Table_List_6_Accent_5 = new CStyle("List Table 6 Colorful - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent_5.Create_Table_List_6(fUF(EThemeColor.themecolorAccent5, 0x9A, null), fUF(EThemeColor.themecolorAccent5, 0x40, null), fUF(EThemeColor.themecolorAccent5, 0x9A, 0x95));
		this.Add( Style_Table_List_6_Accent_5 );

		var Style_Table_List_6_Accent_6 = new CStyle("List Table 6 Colorful - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_6_Accent_6.Create_Table_List_6(fUF(EThemeColor.themecolorAccent6, 0x98, null), fUF(EThemeColor.themecolorAccent6, 0x40, null), fUF(EThemeColor.themecolorAccent6, 0x98, 0x95));
        this.Add( Style_Table_List_6_Accent_6 );

        var Style_Table_List_7_Accent = new CStyle("List Table 7 Colorful", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent.Create_Table_List_7(fUF(EThemeColor.themecolorText1, 0x80, null),  fUF(EThemeColor.themecolorText1, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorText1, 0x80, 0x95));
		this.Add( Style_Table_List_7_Accent );

		var Style_Table_List_7_Accent_1 = new CStyle("List Table 7 Colorful - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent_1.Create_Table_List_7(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent1, null, 0x95));
		this.Add( Style_Table_List_7_Accent_1 );

		var Style_Table_List_7_Accent_2 = new CStyle("List Table 7 Colorful - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent_2.Create_Table_List_7(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent2, 0x97, 0x95));
		this.Add( Style_Table_List_7_Accent_2 );

		var Style_Table_List_7_Accent_3 = new CStyle("List Table 7 Colorful - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent_3.Create_Table_List_7(fUF(EThemeColor.themecolorAccent3, 0x98, null), fUF(EThemeColor.themecolorAccent3, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent3, 0x98, 0x95));
		this.Add( Style_Table_List_7_Accent_3 );

		var Style_Table_List_7_Accent_4 = new CStyle("List Table 7 Colorful - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent_4.Create_Table_List_7(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent4, 0x9A, 0x95));
		this.Add( Style_Table_List_7_Accent_4 );

		var Style_Table_List_7_Accent_5 = new CStyle("List Table 7 Colorful - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent_5.Create_Table_List_7(fUF(EThemeColor.themecolorAccent5, 0x9A, null), fUF(EThemeColor.themecolorAccent5, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent5, 0x9A, 0x95));
		this.Add( Style_Table_List_7_Accent_5 );

		var Style_Table_List_7_Accent_6 = new CStyle("List Table 7 Colorful - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_List_7_Accent_6.Create_Table_List_7(fUF(EThemeColor.themecolorAccent6, 0x98, null), fUF(EThemeColor.themecolorAccent6, 0x40, null), fUF(EThemeColor.themecolorLight1, null, null), fUF(EThemeColor.themecolorAccent6, 0x98, 0x95));
        this.Add( Style_Table_List_7_Accent_6 );

         // Стандартные стили таблиц
		var Style_Table_Lined_Accent = new CStyle("Lined - Accent", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent.Create_Table_Lined(fUF(EThemeColor.themecolorText1, 0x80, null), fUF(EThemeColor.themecolorText1, 0x0D, null));
		this.Add( Style_Table_Lined_Accent );

		var Style_Table_Lined_Accent1 = new CStyle("Lined - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent1.Create_Table_Lined(fUF(EThemeColor.themecolorAccent1, 0xEA, null), fUF(EThemeColor.themecolorAccent1, 0x50, null));
		this.Add( Style_Table_Lined_Accent1 );

		var Style_Table_Lined_Accent2 = new CStyle("Lined - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent2.Create_Table_Lined(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null));
		this.Add( Style_Table_Lined_Accent2 );

		var Style_Table_Lined_Accent3 = new CStyle("Lined - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent3.Create_Table_Lined(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null));
		this.Add( Style_Table_Lined_Accent3 );

		var Style_Table_Lined_Accent4 = new CStyle("Lined - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent4.Create_Table_Lined(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null));
		this.Add( Style_Table_Lined_Accent4 );

		var Style_Table_Lined_Accent5 = new CStyle("Lined - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent5.Create_Table_Lined(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null));
		this.Add( Style_Table_Lined_Accent5 );

		var Style_Table_Lined_Accent6 = new CStyle("Lined - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Lined_Accent6.Create_Table_Lined(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null));
		this.Add( Style_Table_Lined_Accent6 );

        var Style_Table_BorderedLined_Accent0 = new CStyle("Bordered & Lined - Accent", this.Default.Table, null, styletype_Table );
		Style_Table_BorderedLined_Accent0.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorText1, 0x80, null), fUF(EThemeColor.themecolorText1, 0x0D, null), fUF(EThemeColor.themecolorText1, 0x0D, null), fUF(EThemeColor.themecolorText1, 0xA6, null));
		this.Add( Style_Table_BorderedLined_Accent0 );

        var Style_Table_BorderedLined_Accent1 = new CStyle("Bordered & Lined - Accent 1", this.Default.Table, null, styletype_Table );
        Style_Table_BorderedLined_Accent1.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorAccent1, 0xEA, null), fUF(EThemeColor.themecolorAccent1, 0x50, null), fUF(EThemeColor.themecolorAccent1, 0x50, null), fUF(EThemeColor.themecolorAccent1, null, 0x95));
        this.Add( Style_Table_BorderedLined_Accent1 );

        var Style_Table_BorderedLined_Accent2 = new CStyle("Bordered & Lined - Accent 2", this.Default.Table, null, styletype_Table );
        Style_Table_BorderedLined_Accent2.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorAccent2, 0x32, null), fUF(EThemeColor.themecolorAccent2, null, 0x95));
        this.Add( Style_Table_BorderedLined_Accent2 );

		var Style_Table_BorderedLined_Accent3 = new CStyle("Bordered & Lined - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_BorderedLined_Accent3.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorAccent3, 0xFE, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorAccent3, 0x34, null), fUF(EThemeColor.themecolorAccent3, null, 0x95));
		this.Add( Style_Table_BorderedLined_Accent3 );

        var Style_Table_BorderedLined_Accent4 = new CStyle("Bordered & Lined - Accent 4", this.Default.Table, null, styletype_Table );
        Style_Table_BorderedLined_Accent4.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorAccent4, 0x34, null), fUF(EThemeColor.themecolorAccent4, null, 0x95));
        this.Add( Style_Table_BorderedLined_Accent4 );

        var Style_Table_BorderedLined_Accent5 = new CStyle("Bordered & Lined - Accent 5", this.Default.Table, null, styletype_Table );
        Style_Table_BorderedLined_Accent5.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorAccent5, null, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorAccent5, 0x34, null), fUF(EThemeColor.themecolorAccent5, null, 0x95));
        this.Add( Style_Table_BorderedLined_Accent5 );

        var Style_Table_BorderedLined_Accent6 = new CStyle("Bordered & Lined - Accent 6", this.Default.Table, null, styletype_Table );
        Style_Table_BorderedLined_Accent6.Create_Table_BorderedAndLined(fUF(EThemeColor.themecolorAccent6, null, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorAccent6, 0x34, null), fUF(EThemeColor.themecolorAccent6, null, 0x95));
		this.Add( Style_Table_BorderedLined_Accent6 );

		var Style_Table_Bordered_Accent = new CStyle("Bordered", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent.Create_Grid_Table_Light(fUF(EThemeColor.themecolorText1, 0x80, null),  fUF(EThemeColor.themecolorText1, 0x26, null));
		this.Add( Style_Table_Bordered_Accent );

		var Style_Table_Bordered_Accent_1 = new CStyle("Bordered - Accent 1", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent_1.Create_Grid_Table_Light(fUF(EThemeColor.themecolorAccent1, null, null), fUF(EThemeColor.themecolorAccent1, 0x67, null));
		this.Add( Style_Table_Bordered_Accent_1 );

		var Style_Table_Bordered_Accent_2 = new CStyle("Bordered - Accent 2", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent_2.Create_Grid_Table_Light(fUF(EThemeColor.themecolorAccent2, 0x97, null), fUF(EThemeColor.themecolorAccent2, 0x67, null));
		this.Add( Style_Table_Bordered_Accent_2 );

		var Style_Table_Bordered_Accent_3 = new CStyle("Bordered - Accent 3", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent_3.Create_Grid_Table_Light(fUF(EThemeColor.themecolorAccent3, 0x98, null), fUF(EThemeColor.themecolorAccent3, 0x67, null));
		this.Add( Style_Table_Bordered_Accent_3 );

		var Style_Table_Bordered_Accent_4 = new CStyle("Bordered - Accent 4", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent_4.Create_Grid_Table_Light(fUF(EThemeColor.themecolorAccent4, 0x9A, null), fUF(EThemeColor.themecolorAccent4, 0x67, null));
		this.Add( Style_Table_Bordered_Accent_4 );

		var Style_Table_Bordered_Accent_5 = new CStyle("Bordered - Accent 5", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent_5.Create_Grid_Table_Light(fUF(EThemeColor.themecolorAccent5, 0x9A, null), fUF(EThemeColor.themecolorAccent5, 0x67, null));
		this.Add( Style_Table_Bordered_Accent_5 );

		var Style_Table_Bordered_Accent_6 = new CStyle("Bordered - Accent 6", this.Default.Table, null, styletype_Table );
		Style_Table_Bordered_Accent_6.Create_Grid_Table_Light(fUF(EThemeColor.themecolorAccent6, 0x98, null), fUF(EThemeColor.themecolorAccent6, 0x67, null));
		this.Add( Style_Table_Bordered_Accent_6 );
		
		
		this.AddStylesFromObject(AscWord.DEFAULT_STYLE_LIST);
		this.UpdateDefaultStyleLinks();

		for (var nLvl = 0; nLvl <= 8; ++nLvl)
		{
			var oStyleTOC = new CStyle("toc " + (nLvl + 1), this.Default.Paragraph, this.Default.Paragraph, styletype_Paragraph);
			oStyleTOC.CreateTOC(nLvl);
		}

		var oStyleTOCHeading = new CStyle("TOC Heading", this.Default.TOCHeading);
		oStyleTOCHeading.CreateTOCHeading();
		this.Default.TOCHeading = this.Add(oStyleTOCHeading);

		var oStyleTOF = new CStyle("table of figures", this.Default.Paragraph, this.Default.Paragraph, styletype_Paragraph);
		oStyleTOF.CreateTOF();
		this.Default.TOF = this.Add(oStyleTOF);
		

        // Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
        AscCommon.g_oTableId.Add( this, this.Id );
    }
	else
	{
		this.Default = {
			ParaPr      : new CParaPr(),
			TextPr      : new CTextPr(),
			TablePr     : new CTablePr(),
			TableRowPr  : new CTableRowPr(),
			TableCellPr : new CTableCellPr(),

			Paragraph         : null,
			Character         : null,
			Numbering         : null,
			Table             : null,
			TableGrid         : null,
			Headings          : [],
			ParaList          : null,
			Header            : null,
			Footer            : null,
			Hyperlink         : null,
			FollowedHyperlink : null,
			FootnoteText      : null,
			FootnoteTextChar  : null,
			FootnoteReference : null,

			TOC               : [],
			TOCHeading        : null,
			TOF               : null,
			Caption           : null,
			EndnoteText       : null,
			EndnoteTextChar   : null,
			EndnoteReference  : null
		};

		// Заполняем значения по умолчанию
		this.Default.ParaPr.InitDefault();
		this.Default.TextPr.InitDefault();
		this.Default.TablePr.InitDefault();
		this.Default.TableRowPr.InitDefault();
		this.Default.TableCellPr.InitDefault();

		this.Style = [];
	}

    this.LogicDocument = null;
	this.OnChangeDefaultTextPr();
}

CStyles.prototype =
{
//-----------------------------------------------------------------------------------
// Работаем с Id данного объекта
//-----------------------------------------------------------------------------------
    GetId : function()
    {
        return this.Id;
    },

    Get_Id : function()
    {
        return this.GetId();
    },
//-----------------------------------------------------------------------------------
// Базовые функции для работы со стилем
//-----------------------------------------------------------------------------------
	Add : function(Style)
	{
		var Id = Style.Get_Id();
		History.Add(new CChangesStylesAdd(this, Id, Style));
		this.Style[Id] = Style;
		Style.SetParent(this);
		this.Update_Interface(Id);
		return Id;
	},

	GetStyleByStyleId: function (sStyleId)
	{
		for (let sId in this.Style)
		{
			const oStyle = this.Style[sId];
			if (oStyle.GetStyleId() === sStyleId)
				return oStyle;
		}
	},

	Remove : function(Id)
	{
		if (!this.Style[Id])
			return;
		
		this.Style[Id].SetParent(null);
		History.Add(new CChangesStylesRemove(this, Id, this.Style[Id]));
		delete this.Style[Id];
		this.Update_Interface(Id);
	},

	RemapIdReferences : function(OldId, NewId)
	{
		for (var Id in this.Style)
		{
			this.Style[Id].RemapIdReferences(OldId, NewId);
		}
	},

    Copy : function()
    {
        var Styles = new CStyles();

        Styles.Default.ParaPr      = this.Default.ParaPr.Copy();
        Styles.Default.TextPr      = this.Default.TextPr.Copy();
        Styles.Default.TablePr     = this.Default.TablePr.Copy();
        Styles.Default.TableRowPr  = this.Default.TableRowPr.Copy();
        Styles.Default.TableCellPr = this.Default.TableCellPr.Copy();

        for (var StyleId in this.Style)
        {
            Styles.Style[StyleId] = this.Style[StyleId].Copy();
        }

        Styles.OnChangeDefaultTextPr();
		Styles.UpdateDefaultStyleLinks();
        return Styles;
    },

    CopyStyle : function()
    {
        var res = [];
        for (var StyleId in this.Style)
        {
            res[StyleId] = this.Style[StyleId].Copy();
        }
        return res;
    },

    Get_DefaultParaPr : function()
    {
        return this.Default.ParaPr;
    },

	Set_DefaultParaPr : function(ParaPr)
	{
		History.Add(new CChangesStylesChangeDefaultParaPr(this, this.Default.ParaPr, ParaPr));
		this.Default.ParaPr.InitDefault();
		this.Default.ParaPr.Merge(ParaPr);

		// TODO: Пока данная функция используется только в билдере, как только будет использоваться в самом редакторе,
		//       надо будет сделать, чтобы происходил пересчет всех стилей.
	},

    Get_DefaultTextPr : function()
    {
        return this.Default.TextPr;
    },

	Set_DefaultTextPr : function(TextPr)
	{
		History.Add(new CChangesStylesChangeDefaultTextPr(this, this.Default.TextPr, TextPr));
		this.Default.TextPr = new AscWord.CTextPr();
		this.Default.TextPr.InitDefault();
		this.Default.TextPr.Merge(TextPr);
		this.OnChangeDefaultTextPr();
		// TODO: Пока данная функция используется только в билдере, как только будет использоваться в самом редакторе,
		//       надо будет сделать, чтобы происходил пересчет всех стилей.
	},
//-----------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------
    Set_LogicDocument : function(LogicDocument)
    {
        this.LogicDocument = LogicDocument;
    },

    Get_Pr : function(StyleId, Type, TableStyle, ShapeStyle)
    {
        var Pr = {TextPr : new CTextPr(), ParaPr : new CParaPr()};
        // Сначала копируем параметры заданные в табличном стиле
        switch (Type)
        {
            case styletype_Paragraph:
            {
                if (undefined === StyleId || null === StyleId)
                    StyleId = this.Default.Paragraph;

                if (TableStyle != null || ShapeStyle != null)
                {
                    if (ShapeStyle != null)
                    {
                        Pr.TextPr.Merge(ShapeStyle.TextPr);
                        if(!TableStyle)
                        {
                            Pr.ParaPr.Merge(ShapeStyle.ParaPr);
                        }
                    }
                    if (TableStyle != null)
                    {
                        Pr.TextPr.Merge(TableStyle.TextPr);
                        Pr.ParaPr.Merge(TableStyle.ParaPr);
                    }
                }
                else
                {
                    Pr.TextPr.Merge(this.Default.TextPr);
                    Pr.ParaPr.Merge(this.Default.ParaPr);
                }
                break;
            }
            case styletype_Table:
            {
                if (undefined === StyleId || null === StyleId)
                    StyleId = this.Default.Table;

                // Сначала копируем параметры по умолчанию
                Pr.TextPr = this.Default.TextPr.Copy();
                Pr.ParaPr = this.Default.ParaPr.Copy();

                // В таблицах мы не учитываем настройки дефолтового параграфа, т.е. стиля Normal (баг 31469)

                Pr.TablePr     = this.Default.TablePr.Copy();
                Pr.TableRowPr  = this.Default.TableRowPr.Copy();
                Pr.TableCellPr = this.Default.TableCellPr.Copy();

                Pr.TableFirstCol   = new CTableStylePr();
                Pr.TableFirstRow   = new CTableStylePr();
                Pr.TableLastCol    = new CTableStylePr();
                Pr.TableLastRow    = new CTableStylePr();
                Pr.TableBand1Horz  = new CTableStylePr();
                Pr.TableBand1Vert  = new CTableStylePr();
                Pr.TableBand2Horz  = new CTableStylePr();
                Pr.TableBand2Vert  = new CTableStylePr();
                Pr.TableTLCell     = new CTableStylePr();
                Pr.TableTRCell     = new CTableStylePr();
                Pr.TableBLCell     = new CTableStylePr();
                Pr.TableBRCell     = new CTableStylePr();
                Pr.TableWholeTable = new CTableStylePr();

                break;
            }
            case styletype_Character:
            {
                if (undefined === StyleId || null === StyleId)
                    StyleId = this.Default.Character;

                Pr.TextPr = new CTextPr();
                break;
            }
        }

        // Рассчитываем стиль
        this.Internal_Get_Pr(Pr, StyleId, Type, ( null === TableStyle && null == ShapeStyle ? true : false ), [], StyleId);

        if (styletype_Table === Type)
        {
            // В таблицах мы не учитываем настройки дефолтового параграфа, т.е. стиля Normal (баг 31469)

            // Соединим настройки для всей таблицы в одну общую настройку и удалим одну из них за ненадобностью
            Pr.ParaPr.Merge(Pr.TableWholeTable.ParaPr);
            Pr.TextPr.Merge(Pr.TableWholeTable.TextPr);
            Pr.TablePr.Merge(Pr.TableWholeTable.TablePr);
            Pr.TableRowPr.Merge(Pr.TableWholeTable.TableRowPr);
            Pr.TableCellPr.Merge(Pr.TableWholeTable.TableCellPr);
        }

        return Pr;
    },

    Get_Next : function(StyleId)
    {
        var NextId = this.Style[StyleId].Next;

        if (null !== NextId && undefined !== this.Style[NextId])
            return NextId;

        return null;
    },

    Get_Name : function(StyleId)
    {
        return this.GetName(StyleId);
    },
	
	GetName : function(styleId)
	{
		let style = this.Get(styleId);
		return style ? style.GetName() : "";
	},

    Get_Default_Paragraph : function()
    {
        return this.Default.Paragraph;
    },

    Get_Default_Numbering : function()
    {
        return this.Default.Numbering;
    },

    Get_Default_Table : function()
    {
        return this.Default.Table;
    },

    Get_Default_TableGrid : function()
    {
        return this.Default.TableGrid;
    },

    Get_Default_Heading : function(Lvl)
    {
        Lvl = Math.max( Math.min( Lvl, 8 ), 0 );
        return this.Default.Headings[Lvl];
    },

    Get_Default_ParaList : function()
    {
        return this.Default.ParaList;
    },

    Get_Default_Header : function()
    {
        return this.Default.Header;
    },

    Get_Default_Footer : function()
    {
        return this.Default.Footer;
    },

    Internal_Get_Pr : function(Pr, StyleId, Type, bUseDefault, PassedStyles, StartStyleId)
	{
		// Делаем проверку от зацикливания, среди уже пройденных стилей ищем текущий стриль.
		for (var nIndex = 0, nCount = PassedStyles.length; nIndex < nCount; nIndex++)
		{
			if (PassedStyles[nIndex] == StyleId)
				return;
		}
		PassedStyles.push(StyleId);

		var Style = this.Style[StyleId];
		if (Style && Style.Type !== Type)
			Style = undefined;
		
		if (!StyleId || !Style)
		{
			if (true === bUseDefault)
			{
				// Копируем свойства по умолчанию для данного типа
				switch (Type)
				{
					case styletype_Paragraph:
					{
						var DefId = this.Default.Paragraph;

						if (undefined != this.Style[DefId])
						{
							Pr.ParaPr.Merge(this.Style[DefId].ParaPr);
							Pr.TextPr.Merge(this.Style[DefId].TextPr);
						}

						break;
					}
					case styletype_Numbering:
					{
						var DefId = this.Default.Numbering;
						break;
					}
					case styletype_Table:
					{
						var DefId = this.Default.Table;

						if (undefined != this.Style[DefId])
						{
							Pr.ParaPr.Merge(this.Style[DefId].ParaPr);
							Pr.TextPr.Merge(this.Style[DefId].TextPr);
							Pr.TablePr.Merge(this.Styles[DefId].TablePr);
							Pr.TableRowPr.Merge(this.Styles[DefId].TableRowPr);
							Pr.TableCellPr.Merge(this.Styles[DefId].TableCellPr);
						}

						break;
					}
					case styletype_Character:
					{
						var DefId = this.Default.Character;

						if (undefined != this.Style[DefId])
						{
							Pr.TextPr.Merge(this.Style[DefId].TextPr);
						}

						break;
					}
				}
			}

			return;
		}

		if (null === Style.BasedOn)
		{
			// TODO: Проверить нужно ли копировать стили по умолчанию для данного типа стиля, когда сам стиль задан
			//       Для параграфа, вроде как не нужно (см. ivanova_veronica.docx стиль "Colon")

			// Копируем свойства по умолчанию для данного типа
			if (true === bUseDefault)
			{
				// Копируем свойства по умолчанию для данного типа
				switch (Type)
				{
					case styletype_Paragraph:
					{
						//                        var DefId = this.Default.Paragraph;
						//
						//                        Pr.ParaPr.Merge( this.Style[DefId].ParaPr );
						//                        Pr.TextPr.Merge( this.Style[DefId].TextPr );

						break;
					}
					case styletype_Numbering:
					{
						var DefId = this.Default.Numbering;
						break;
					}
					case styletype_Table:
					{
						var DefId = this.Default.Table;

						/*
						 Pr.ParaPr.Merge( this.Style[DefId].ParaPr );
						 Pr.TextPr.Merge( this.Style[DefId].TextPr );
						 Pr.TablePr.Merge( this.Styles[DefId].TablePr );
						 Pr.TableRowPr.Merge( this.Styles[DefId].TableRowPr );
						 Pr.TableCellPr.Merge( this.Styles[DefId].TableCellPr );
						 */

						break;
					}
					case styletype_Character:
					{
						var DefId = this.Default.Character;

						if (undefined != this.Style[DefId])
						{
							Pr.TextPr.Merge(this.Style[DefId].TextPr);
						}

						break;
					}
				}
			}
		}
		else
		{
			// Копируем свойства родительского стиля
			this.Internal_Get_Pr(Pr, Style.BasedOn, Type, false, PassedStyles, StartStyleId);
		}

		// Копируем свойства из стиля нумерации
		var oLogicDocument = private_GetWordLogicDocument();
		if (oLogicDocument
			&& Style.ParaPr.NumPr
			&& (styletype_Paragraph === Type || styletype_Table === Type))
		{
			let oNumbering = oLogicDocument.GetNumbering();
			let sNumId     = Style.ParaPr.NumPr.NumId;
			if (0 !== sNumId && "0" !== sNumId)
			{
				if (undefined === sNumId && Pr.ParaPr.NumPr)
					sNumId = Pr.ParaPr.NumPr.NumId;

				let oNum = oNumbering.GetNum(sNumId);
				if (oNum)
				{
					var nLvl = oNum.GetLvlByStyle(StyleId);
					if (-1 !== nLvl)
						Pr.ParaPr.Merge(oNumbering.GetParaPr(sNumId, nLvl));
					else if (undefined !== Style.ParaPr.NumPr.Lvl)
						Pr.ParaPr.Merge(oNumbering.GetParaPr(sNumId, Style.ParaPr.NumPr.Lvl));
					else
						Pr.ParaPr.NumPr = undefined;
				}
			}
			else
			{
				// Word значение 0 для NumId воспринимает как отсутствие нумерации и обнуляет левый отступ и
				// отступ первой строки
				Pr.ParaPr.NumPr         = undefined;
				Pr.ParaPr.Ind.Left      = 0;
				Pr.ParaPr.Ind.FirstLine = 0;
			}
		}

		// Копируем свойства текущего стиля
		switch (Type)
		{
			case styletype_Paragraph:
			{
				Pr.ParaPr.Merge(Style.ParaPr);
				Pr.TextPr.Merge(Style.TextPr);

				break;
			}
			case styletype_Numbering:
			{
				break;
			}
			case styletype_Table:
			{
				Pr.ParaPr.Merge(Style.ParaPr);
				Pr.TextPr.Merge(Style.TextPr);

				if (undefined !== Style.TablePr)
				{
					Pr.TablePr.Merge(Style.TablePr);
					Pr.TableRowPr.Merge(Style.TableRowPr);
					Pr.TableCellPr.Merge(Style.TableCellPr);

					Pr.TableBand1Horz.Merge(Style.TableBand1Horz);
					Pr.TableBand1Vert.Merge(Style.TableBand1Vert);
					Pr.TableBand2Horz.Merge(Style.TableBand2Horz);
					Pr.TableBand2Vert.Merge(Style.TableBand2Vert);
					Pr.TableFirstCol.Merge(Style.TableFirstCol);
					Pr.TableFirstRow.Merge(Style.TableFirstRow);
					Pr.TableLastCol.Merge(Style.TableLastCol);
					Pr.TableLastRow.Merge(Style.TableLastRow);
					Pr.TableTLCell.Merge(Style.TableTLCell);
					Pr.TableTRCell.Merge(Style.TableTRCell);
					Pr.TableBLCell.Merge(Style.TableBLCell);
					Pr.TableBRCell.Merge(Style.TableBRCell);
					Pr.TableWholeTable.Merge(Style.TableWholeTable);
				}

				break;
			}
			case styletype_Character:
			{
				Pr.TextPr.Merge(Style.TextPr);

				break;
			}
		}
	},

    Document_Get_AllFontNames : function(AllFonts)
    {
        for ( var Id in this.Style )
        {
            var Style = this.Style[Id];
            Style.Document_Get_AllFontNames(AllFonts);
        }

        this.Default.TextPr.Document_Get_AllFontNames(AllFonts);
    },

    Update_Interface : function(StyleId)
    {
        if (null != this.LogicDocument && undefined !== this.LogicDocument)
        {
            // Данный стиль может быть базовым для других стилей, поэтому нам нужно пересчитать все параграфы, не только у
            // которых выставлен данный стиль, но и у которых выставлен стиль, для которого данный будет базовым (в любом поколе��ии).

            this.LogicDocument.Add_ChangedStyle(this.private_GetAllBasedStylesId(StyleId));
        }
    },

    private_GetAllBasedStylesId : function(StyleId)
    {
        var arrStyles = [];

        // Отдельно добавляем StyleId, т.к. данная функция вызывается и после удаления стиля из списка,
        // но при этом в данный массив стиль должен попасть.
        arrStyles.push(StyleId);

        for (var CurStyleId in this.Style)
        {
            if (CurStyleId == StyleId)
            {
                arrStyles.push(StyleId);
            }

            var oStyle = this.Style[CurStyleId];
            var BaseId = oStyle.Get_BasedOn();
            var PassedStyles = [];
            while (null != BaseId && undefined != BaseId)
            {
                var bBreak = false;
                // Делаем проверку от зацикливания, среди уже пройденных стилей ищем текущий стриль.
                for (var nIndex = 0, nCount = PassedStyles.length; nIndex < nCount; nIndex++)
                {
                    if (PassedStyles[nIndex] == BaseId)
                    {
                        bBreak = true;
                        break;
                    }
                }

                if (true === bBreak)
                    break;

                PassedStyles.push(BaseId);

                if (BaseId == StyleId)
                {
                    arrStyles.push(CurStyleId);
                    break;
                }

                var BaseStyle = this.Style[BaseId];
                if (!BaseStyle)
                    break;

                BaseId = BaseStyle.Get_BasedOn();
            }
        }

        return arrStyles;
    },

    Check_StyleNumberingOnLoad : function(Numbering)
    {
        // TODO: Похоже Word сначала пробегается по дефолтовым стилям, типа Heading, потом по остальным.
        for (var StyleId in this.Style)
        {
            var Style = this.Style[StyleId];

            var oNumPr = Style.ParaPr.NumPr;
            if (!oNumPr || !oNumPr.NumId)
                continue;

            var oNum = Numbering.GetNum(oNumPr.NumId);
            if (!oNum)
                continue;

            var nLvl = (oNumPr.Lvl ? oNumPr.Lvl : 0);
            var oLvl = oNum.GetLvl(nLvl);

            if (!oLvl || oLvl.GetPStyle())
                continue;

            var oNewLvl = oLvl.Copy();
            oNewLvl.PStyle = StyleId;
            oNum.SetLvl(oNewLvl, nLvl);
        }
    },
//-----------------------------------------------------------------------------------
// Undo/Redo функции
//-----------------------------------------------------------------------------------
	GetSelectionState : function()
    {
    },

	SetSelectionState : function(State, StateIndex)
    {
    },

    Get_ParentObject_or_DocumentPos : function()
    {
        return { Type : AscDFH.historyitem_recalctype_Inline, Data : 0 };
    },

    Refresh_RecalcData : function(Data)
    {
        var Type = Data.Type;

        var bNeedRecalc = false;

        switch ( Type )
        {
            case AscDFH.historyitem_Styles_Add   :
            case AscDFH.historyitem_Styles_Remove:
            {
                bNeedRecalc = true;
                break;
            }
			case AscDFH.historyitem_Styles_ChangeDefaultTextPr:
			{
				// TODO: Нужно сделать отдельный метод для проверки по стилю рана (в том числе и дефолтовому, как здесь)
				let logicDocument = private_GetWordLogicDocument();
				if (!logicDocument || !logicDocument.IsDocumentEditor())
					return;
				
				let paragraphs = logicDocument.GetAllParagraphs({All : true});
				for (let i = 0, count = paragraphs.length; i < count; ++i)
				{
					paragraphs[i].Recalc_CompiledPr();
					paragraphs[i].Recalc_RunsCompiledPr();
				}
			}
        }

        if ( true === bNeedRecalc )
        {
            // Сообщаем родительскому классу, что изменения произошли в элементе с номером this.Index и на странице this.PageNum
            return this.Refresh_RecalcData2(Data.Id);
        }
    },

    Refresh_RecalcData2 : function(StyleId)
    {
        if (undefined != StyleId)
        {
            // TODO: Надо сделать механизм, чтобы данное действие не вызывалось много раз подряд, а только 1.
            var LogicDocument = private_GetWordLogicDocument();
			if(!LogicDocument)
				return;

            var AllParagraphs = [];

            if (StyleId != this.Default.Paragraph)
            {
                var AllStylesId = this.private_GetAllBasedStylesId(StyleId);
                AllParagraphs = LogicDocument.GetAllParagraphsByStyle(AllStylesId);
            }
            else
            {
                AllParagraphs = LogicDocument.GetAllParagraphs({All : true});
            }

            var Count = AllParagraphs.length;
            for (var Index = 0; Index < Count; Index++)
            {
                var Para = AllParagraphs[Index];
                Para.Refresh_RecalcData({ Type : AscDFH.historyitem_Paragraph_PStyle });
            }
        }
    },
//-----------------------------------------------------------------------------------
// Функции для совместного редактирования
//-----------------------------------------------------------------------------------
	Process_EndLoad : function(LinkData)
    {
        if (undefined !== LinkData.UpdateStyleId)
        {
            var StyleId = LinkData.UpdateStyleId;

            var LogicDocument = private_GetWordLogicDocument();
            if (!LogicDocument)
                return;

            var AllParagraphs = [];

            if (StyleId != this.Default.Paragraph)
            {
                var AllStylesId = this.private_GetAllBasedStylesId(StyleId);
                AllParagraphs = LogicDocument.GetAllParagraphsByStyle(AllStylesId);
            }
            else
            {
                AllParagraphs = LogicDocument.GetAllParagraphs({All : true});
            }

            var Count = AllParagraphs.length;
            for (var Index = 0; Index < Count; Index++)
            {
                var Para = AllParagraphs[Index];
                Para.Recalc_CompiledPr();
                Para.Recalc_RunsCompiledPr();
            }
        }
    }
};
/**
 * Получаем стиль по идентификатору
 * @param sStyleId {string}
 * @returns {?CStyle}
 */
CStyles.prototype.Get = function(sStyleId)
{
	return (this.Style[sStyleId] ? this.Style[sStyleId] : null);
};
CStyles.prototype.GetAllTableStyles = function()
{
	let arrTableStyles = [];
	for (let Id in this.Style)
	{
		let oStyle = this.Style[Id];
		if (oStyle.IsTableStyle())
			arrTableStyles.push(oStyle);
	}

	return arrTableStyles;
};
CStyles.prototype.GetAllStyles = function()
{
	let result = [];
	for (let id in this.Style)
	{
		result.push(this.Style[id]);
	}

	return result;
};
CStyles.prototype.GetRelatedStyles = function(styleId)
{
	let result = [];
	for (let id in this.Style)
	{
		if (id === styleId)
			continue;

		let style = this.Style[id];

		if (styleId === style.GetBasedOn()
			|| styleId === style.GetLink()
			|| styleId === style.GetNext())
			result.push(style);
	}

	return result;
};
/**
 * Получаем идентификатор стиля по его имени
 * @param {string} sName
 * @param {boolean} [isReturnParaDefault=false] Возвращать ли дефолтовый стиль для параграфа, если стиль не найден
 * @param {boolean} [isOnlyCharChecking=false] Проверять ли имя стиля только по символам
 * @returns {?string}
 */
CStyles.prototype.GetStyleIdByName = function(sName, isReturnParaDefault, isSimplify)
{
	for (var sId in this.Style)
	{
		var oStyle = this.Style[sId];
		if (sName === oStyle.GetName(isSimplify))
			return sId;
	}

	if (isReturnParaDefault)
		return this.Default.Paragraph;

	return null;
};
CStyles.prototype.Create_StyleFromInterface = function(oAscStyle, bCheckLink)
{
	var sStyleName = oAscStyle.get_Name();
	var sStyleId   = this.GetStyleIdByName(sStyleName);
	if (null !== sStyleId)
	{
		var oStyle = this.Style[sStyleId];

		var NewStyleParaPr = oAscStyle.get_ParaPr();
		var NewStyleTextPr = oAscStyle.get_TextPr();

		var BasedOnId = this.GetStyleIdByName(oAscStyle.get_BasedOn());
		var NextId    = this.GetStyleIdByName(oAscStyle.get_Next());

		oStyle.Set_Type(oAscStyle.get_Type());

		if (BasedOnId === sStyleId || sStyleId === this.Default.Paragraph)
		{
			if (sStyleId !== this.Default.Paragraph)
			{
				var oBaseStyle      = this.Get(BasedOnId);
				var oBasedBasesOnId = this.Get_Default_Paragraph();
				if (oBaseStyle)
				{
					oBasedBasesOnId = oBaseStyle.Get_BasedOn();
					if (oBaseStyle.Get_BasedOn() !== sStyleId)
						oBasedBasesOnId = oBaseStyle.Get_BasedOn();
				}

				oStyle.Set_BasedOn(oBasedBasesOnId);
			}
			else
			{
				oStyle.Set_BasedOn(null);
			}

			var OldStyleParaPr = oStyle.ParaPr.Copy();
			var OldStyleTextPr = oStyle.TextPr.Copy();
			OldStyleParaPr.Merge(NewStyleParaPr);
			OldStyleTextPr.Merge(NewStyleTextPr);
			NewStyleParaPr = OldStyleParaPr;
			NewStyleTextPr = OldStyleTextPr;
		}
		else
		{
			oStyle.Set_BasedOn(BasedOnId);
		}

		if (null === oStyle.Get_Next() || (null !== NextId && NextId !== sStyleId))
		{
			if (NextId === sStyleId)
				oStyle.Set_Next(null);
			else
				oStyle.Set_Next(NextId);
		}

		var oAscLink   = oAscStyle.get_Link();
		var sOldLinkId = oStyle.Get_Link();
		if (sOldLinkId && this.Style[sOldLinkId])
			oAscLink.put_Name(this.Style[sOldLinkId].GetName());
		else
			bCheckLink = false;

		if (false != bCheckLink && null != oAscLink && undefined !== oAscLink)
		{
			var oLinkedStyle = this.Create_StyleFromInterface(oAscLink, false);
			oStyle.Set_Link(oLinkedStyle.Get_Id());
			oLinkedStyle.Set_Link(oStyle.Get_Id());
		}

		oStyle.Set_TextPr(NewStyleTextPr);
		oStyle.Set_ParaPr(NewStyleParaPr, true);
		
		let numPr = oStyle.GetParaPr().NumPr;
		oStyle.GetRelatedParagraphs().forEach(function(paragraph)
		{
			if (numPr && numPr.IsEqual(paragraph.GetNumPr()))
				paragraph.SetNumPr(null);

			paragraph.RecalcCompiledPr();
		});
		
		return oStyle;
	}
	else
	{
		var oStyle = new CStyle();

		var BasedOnId = this.GetStyleIdByName(oAscStyle.get_BasedOn());
		oStyle.Set_BasedOn(BasedOnId);
		oStyle.Set_Next(this.GetStyleIdByName(oAscStyle.get_Next()));
		oStyle.Set_Type(oAscStyle.get_Type());
		oStyle.Set_TextPr(oAscStyle.get_TextPr());
		oStyle.Set_ParaPr(oAscStyle.get_ParaPr(), true);
		oStyle.Set_Name(sStyleName);
		oStyle.SetCustom(true);

		if (styletype_Paragraph === oStyle.Get_Type())
			oStyle.Set_QFormat(true);

		var oAscLink = oAscStyle.get_Link();
		if (false != bCheckLink && null != oAscLink && undefined !== oAscLink)
		{
			var oLinkedStyle = this.Create_StyleFromInterface(oAscLink, false);
			oStyle.Set_Link(oLinkedStyle.Get_Id());
			oLinkedStyle.Set_Link(oStyle.Get_Id());
		}

		this.Add(oStyle);
		return oStyle;
	}
};
CStyles.prototype.Remove_StyleFromInterface = function(styleId)
{
	// Если этот стиль не один из стилей по умолчанию, тогда мы просто удаляем этот стиль
	// и очищаем все параграфы с сылкой на этот стиль.

	let style = this.Style[styleId];
	if (!style)
		return;
	
	let styleName = this.GetName(styleId);
	let defaultObject = AscWord.getDefaultStyleObject(styleName);
	if (defaultObject)
	{
		style.Clear(defaultObject.Name, null, null, defaultObject.Type);
		style.SetFromObject(defaultObject);
		
		if (defaultObject.Link)
			style.SetLink(this.GetStyleIdByName(AscWord.getDefaultStyleName(defaultObject.Link)));
		if (defaultObject.BasedOn)
			style.SetBasedOn(this.GetStyleIdByName(AscWord.getDefaultStyleName(defaultObject.BasedOn)));
		if (defaultObject.Next)
			style.SetNext(this.GetStyleIdByName(AscWord.getDefaultStyleName(defaultObject.Next)));
	}
	else
	{
		this.Remove(styleId);

		if (this.LogicDocument)
		{
			var AllParagraphs = this.LogicDocument.GetAllParagraphsByStyle([styleId]);
			var Count = AllParagraphs.length;
			for (var Index = 0; Index < Count; Index++)
			{
				var Para = AllParagraphs[Index];
				Para.Style_Remove();
			}
		}
	}
	this.Update_Interface(styleId);
};
CStyles.prototype.Remove_AllCustomStylesFromInterface = function()
{
	for (var StyleId in this.Style)
	{
		var Style = this.Style[StyleId];
		if ((styletype_Paragraph === Style.GetType() || styletype_Character === Style.GetType()) && true === Style.GetQFormat())
		{
			this.Remove_StyleFromInterface(StyleId);
		}
	}
};
CStyles.prototype.IsStyleDefaultByName = function(styleName, isSimplify)
{
	var styleId = this.GetStyleIdByName(styleName, false, isSimplify);
	if (!styleId)
		return false;

	return this.IsStyleDefaultById(styleId);
};
CStyles.prototype.IsStyleDefaultById = function(styleId)
{
	let styleName = this.GetName(styleId);
	return !!AscWord.getDefaultStyleObject(styleName);
};
CStyles.prototype.Is_StyleDefaultOOXML = function(sStyleName)
{
	var StyleId = this.GetStyleIdByName(sStyleName);
	if (null === StyleId)
		return false;

	if (StyleId == this.Default.Paragraph
		|| StyleId == this.Default.Character
		|| StyleId == this.Default.Numbering
		|| StyleId == this.Default.Table)
	{
		return true;
	}
	return false;
};
CStyles.prototype.Is_DefaultStyleChanged = function(styleName)
{
	let styleId = this.GetStyleIdByName(styleName);
	if (!styleId)
		return false;
	
	let currentStyle  = this.Style[styleId];
	let defaultObject = AscWord.getDefaultStyleObject(styleName);
	if (!defaultObject)
		return false;
	
	let _t = this;
	return AscCommon.ExecuteNoHistory(function()
	{
		let style = CStyle.fromObject(defaultObject);
		
		if (defaultObject.Link)
			style.SetLink(_t.GetStyleIdByName(AscWord.getDefaultStyleName(defaultObject.Link)));
		if (defaultObject.BasedOn)
			style.SetBasedOn(_t.GetStyleIdByName(AscWord.getDefaultStyleName(defaultObject.BasedOn)));
		if (defaultObject.Next)
			style.SetNext(_t.GetStyleIdByName(AscWord.getDefaultStyleName(defaultObject.Next)));
		
		return (true !== style.Is_Equal(currentStyle));
	});
};
CStyles.prototype.AddStylesFromObject = function(obj)
{
	let styleMap = {};
	
	for (let i = 0; i < obj.length; ++i)
	{
		let style = AscWord.CStyle.fromObject(obj[i]);
		if (!style || !obj[i].StyleId)
			continue;
		
		let styleId = this.Add(style);
		styleMap[obj[i].StyleId] = {
			style   : style,
			styleId : styleId,
			next    : undefined !== obj[i].Next ? obj[i].Next : undefined,
			basedOn : undefined !== obj[i].BasedOn ? obj[i].BasedOn : undefined,
			link    : undefined !== obj[i].Link ? obj[i].Link : undefined
		};
		
		if (obj[i].Default)
		{
			switch (obj[i].Type)
			{
				case styletype_Paragraph:
					this.Default.Paragraph = styleId;
					break;
				case styletype_Character:
					this.Default.Character = styleId;
					break;
				case styletype_Numbering:
					this.Default.Numbering = styleId;
					break;
				case styletype_Table:
					this.Default.Table = styleId;
					break;
			}
		}
	}
	
	for (let styleId in styleMap)
	{
		let entry = styleMap[styleId];
		let style = entry.style;
		
		if (entry.next && styleMap[entry.next])
			style.SetNext(styleMap[entry.next].styleId);
		
		if (entry.basedOn && styleMap[entry.basedOn])
			style.SetBasedOn(styleMap[entry.basedOn].styleId);
		
		if (entry.link && styleMap[entry.link])
			style.SetLink(styleMap[entry.link].styleId);
	}
};
/**
 * Обновляем ссылки на все дефолтные стили
 */
CStyles.prototype.UpdateDefaultStyleLinks = function()
{
	// TODO: Если данный метод будет слишком часто вызываться, то нужно переделать, чтобы реальное обновление срабатывало
	//       при запросе к одному из дефолтовых стилей GetDefaultParagraph, например
	
	let localHyperlink  = AscCommon.translateManager.getValue("Hyperlink").toLowerCase().replace(/\s/g,"");
	let localFHyperlink = AscCommon.translateManager.getValue("FollowedHyperlink").toLowerCase().replace(/\s/g,"");
	for (let styleId in this.Style)
	{
		let name = this.Style[styleId].GetName().toLowerCase().replace(/\s/g,"");
		switch (name)
		{
			case "normal":
				this.Default.Paragraph = styleId;
				break;
			case "defaultparagraphfont":
				this.Default.Character = styleId;
				break;
			case "nolist":
				this.Default.Numbering = styleId;
				break;
			case "normaltable":
				this.Default.Table = styleId;
				break;
			case "tablegrid":
				this.Default.TableGrid = styleId;
				break;
			case "heading1":
				this.Default.Headings[0] = styleId;
				break;
			case "heading2":
				this.Default.Headings[1] = styleId;
				break;
			case "heading3":
				this.Default.Headings[2] = styleId;
				break;
			case "heading4":
				this.Default.Headings[3] = styleId;
				break;
			case "heading5":
				this.Default.Headings[4] = styleId;
				break;
			case "heading6":
				this.Default.Headings[5] = styleId;
				break;
			case "heading7":
				this.Default.Headings[6] = styleId;
				break;
			case "heading8":
				this.Default.Headings[7] = styleId;
				break;
			case "heading9":
				this.Default.Headings[8] = styleId;
				break;
			case "listparagraph":
				this.Default.ParaList = styleId;
				break;
			case "header":
				this.Default.Header = styleId;
				break;
			case "footer":
				this.Default.Footer = styleId;
				break;
			case "hyperlink":
			case localHyperlink:
				this.Default.Hyperlink = styleId;
				break;
			case "followedhyperlink":
			case localFHyperlink:
				this.Default.FollowedHyperlink = styleId;
				break;
			case "footnotetext":
				this.Default.FootnoteText = styleId;
				break
			case "footnotetextchar":
				this.Default.FootnoteTextChar = styleId;
				break;
			case "footnotereference":
				this.Default.FootnoteReference = styleId;
				break;
			case "endnotetext":
				this.Default.EndnoteText = styleId;
				break;
			case "endnotetextchar":
				this.Default.EndnoteTextChar = styleId;
				break;
			case "endnotereference":
				this.Default.EndnoteReference = styleId;
				break;
			case "toc1":
				this.Default.TOC[0] = styleId;
				break;
			case "toc2":
				this.Default.TOC[1] = styleId;
				break;
			case "toc3":
				this.Default.TOC[2] = styleId;
				break;
			case "toc4":
				this.Default.TOC[3] = styleId;
				break;
			case "toc5":
				this.Default.TOC[4] = styleId;
				break;
			case "toc6":
				this.Default.TOC[5] = styleId;
				break;
			case "toc7":
				this.Default.TOC[6] = styleId;
				break;
			case "toc8":
				this.Default.TOC[7] = styleId;
				break;
			case "toc9":
				this.Default.TOC[8] = styleId;
				break;
			case "tocheading":
				this.Default.TOCHeading = styleId;
				break;
			case "tableoffigures":
				this.Default.TOF = styleId;
				break;
			case "placeholdertext":
				this.Default.PlaceholderText = styleId;
				break;
		}
	}
};
/**
 * Получаем идентификатор стиля по умолчанию для параграфов
 * @returns {string}
 */
CStyles.prototype.GetDefaultParagraph = function()
{
	return this.Default.Paragraph;
};
/**
 * Получаем идентификатор стиля по умолчанию для ранов
 * @returns {string}
 */
CStyles.prototype.GetDefaultCharacter = function()
{
	return this.Default.Character;
};
CStyles.prototype.GetDefaultFootnoteText = function()
{
	return this.Default.FootnoteText;
};
CStyles.prototype.GetDefaultFootnoteTextChar = function()
{
	return this.Default.FootnoteTextChar;
};
CStyles.prototype.GetDefaultFootnoteReference = function()
{
	return this.Default.FootnoteReference;
};
CStyles.prototype.GetDefaultEndnoteText = function()
{
	return this.Default.EndnoteText;
};
CStyles.prototype.GetDefaultEndnoteTextChar = function()
{
	return this.Default.EndnoteTextChar;
};
CStyles.prototype.GetDefaultEndnoteReference = function()
{
	return this.Default.EndnoteReference;
};
CStyles.prototype.GetDefaultTOC = function(nLvl)
{
	nLvl = Math.max(Math.min(nLvl, 8), 0);
	return this.Default.TOC[nLvl];
};
CStyles.prototype.GetDefaultTOCHeading = function()
{
	return this.Default.TOCHeading;
};
CStyles.prototype.GetDefaultTOF = function()
{
	return this.Default.TOF;
};
CStyles.prototype.GetDefaultHyperlink = function()
{
	return this.Default.Hyperlink;
};
CStyles.prototype.GetDefaultFollowedHyperlink = function()
{
	return this.Default.FollowedHyperlink;
};
CStyles.prototype.GetDefaultHeading = function(nLvl)
{
	return this.Default.Headings[Math.max(Math.min(nLvl, 8), 0)];
};
CStyles.prototype.GetDefaultPlaceholderText = function()
{
	return this.Default.PlaceholderText;
};
CStyles.prototype.HaveHeadingsNum = function()
{
	for (let index = 0; index <= 8; ++index)
	{
		let style = this.Get(this.GetDefaultHeading(index));
		let numPr;
		if (style
			&& (numPr = style.GetParaPr().GetNumPr())
			&& numPr.IsValid())
			return true;
	}

	return false;
};
CStyles.prototype.GetHeadingLevelByName = function(sStyleName)
{
	let sId = this.GetStyleIdByName(sStyleName);
	if (!sId)
		return -1;

	return this.GetHeadingLevelById(sId);
};
CStyles.prototype.GetHeadingLevelById = function(sId)
{
	for (let nIndex = 0; nIndex <= 8; ++nIndex)
	{
		if (sId === this.Default.Headings[nIndex])
			return nIndex;
	}

	return -1;
};
/**
 * Получаем тип набора стилей для Table of Contents
 * @returns {Asc.c_oAscTOCStylesType}
 */
CStyles.prototype.GetTOCStylesType = function()
{
	if (this.private_CheckTOCStyles(Asc.c_oAscTOCStylesType.Simple))
		return Asc.c_oAscTOCStylesType.Simple;
	else if (this.private_CheckTOCStyles(Asc.c_oAscTOCStylesType.Standard))
		return Asc.c_oAscTOCStylesType.Standard;
	else if (this.private_CheckTOCStyles(Asc.c_oAscTOCStylesType.Modern))
		return Asc.c_oAscTOCStylesType.Modern;
	else if (this.private_CheckTOCStyles(Asc.c_oAscTOCStylesType.Classic))
		return Asc.c_oAscTOCStylesType.Classic;
	else if (this.private_CheckTOCStyles(Asc.c_oAscTOCStylesType.Web))
		return Asc.c_oAscTOCStylesType.Web;

	return Asc.c_oAscTOCStylesType.Current;
};
CStyles.prototype.private_CheckTOCStyles = function(nType)
{
	for (var nLvl = 0; nLvl <= 8; ++nLvl)
	{
		if (!this.private_CheckTOCStyle(nLvl, nType))
			return false;
	}

	return true;
};
CStyles.prototype.private_CheckTOCStyle = function(nLvl, nType)
{
	var oTOCStyle = this.Get(this.GetDefaultTOC(nLvl));

	if (!this.LogicDocument || !oTOCStyle)
		return false;

	this.LogicDocument.TurnOffHistory();
	var oCheckStyle = new CStyle();
	oCheckStyle.Clear(oTOCStyle.GetName(), oTOCStyle.GetBasedOn(), oTOCStyle.GetNext(), oTOCStyle.GetType());
	oCheckStyle.CreateTOC(nLvl, nType);
	this.LogicDocument.TurnOnHistory();

	return (!!oCheckStyle.IsEqual(oTOCStyle));
};

/**
 * Get table of figures style type
 * @returns {Asc.c_oAscTOCStylesType}
 */
CStyles.prototype.GetTOFStyleType = function()
{
	if(this.private_CheckTOFStyle(Asc.c_oAscTOFStylesType.Classic))
	{
		return Asc.c_oAscTOFStylesType.Classic;
	}
	else if(this.private_CheckTOFStyle(Asc.c_oAscTOFStylesType.Distinctive))
	{
		return Asc.c_oAscTOFStylesType.Distinctive;
	}
	else if(this.private_CheckTOFStyle(Asc.c_oAscTOFStylesType.Centered))
	{
		return Asc.c_oAscTOFStylesType.Centered;
	}
	else if(this.private_CheckTOFStyle(Asc.c_oAscTOFStylesType.Formal))
	{
		return Asc.c_oAscTOFStylesType.Formal;
	}
	else if(this.private_CheckTOFStyle(Asc.c_oAscTOFStylesType.Simple))
	{
		return Asc.c_oAscTOFStylesType.Simple;
	}
	else if(this.private_CheckTOFStyle(Asc.c_oAscTOFStylesType.Web))
	{
		return Asc.c_oAscTOFStylesType.Web;
	}
	return Asc.c_oAscTOFStylesType.Current;
};
CStyles.prototype.private_CheckTOFStyle = function(nType)
{
	var oTOCStyle = this.Get(this.GetDefaultTOF());

	if (!this.LogicDocument || !oTOCStyle)
		return false;

	this.LogicDocument.TurnOffHistory();
	var oCheckStyle = new CStyle();
	oCheckStyle.Clear(oTOCStyle.GetName(), oTOCStyle.GetBasedOn(), oTOCStyle.GetNext(), oTOCStyle.GetType());
	oCheckStyle.CreateTOF(nType);
	this.LogicDocument.TurnOnHistory();

	return (!!oCheckStyle.IsEqual(oTOCStyle));
};

/**
 * Переделываем стили для Table of Contents на заданную коллекцию стилей
 * @param nType {Asc.c_oAscTOCStylesType}
 */
CStyles.prototype.SetTOCStylesType = function(nType)
{
	if (Asc.c_oAscTOCStylesType.Current === nType)
		return;

	for (var nLvl = 0; nLvl <= 8; ++nLvl)
	{
		var oStyle = this.Get(this.GetDefaultTOC(nLvl));
		if (!oStyle)
			continue;

		oStyle.Clear(oStyle.GetName(), oStyle.GetBasedOn(), oStyle.GetNext(), oStyle.GetType());
		oStyle.CreateTOC(nLvl, nType);
	}
};


/**
 * Recreate table of figures style according to current style type
 * @param nType {Asc.c_oAscTOCStylesType}
 */
CStyles.prototype.SetTOFStyleType = function(nType)
{
	if (Asc.c_oAscTOCStylesType.Current === nType)
		return;
	var oStyle = this.Get(this.GetDefaultTOF());
	if (!oStyle)
		return;

	oStyle.Clear(oStyle.GetName(), oStyle.GetBasedOn(), oStyle.GetNext(), oStyle.GetType());
	oStyle.CreateTOF(nType);
};

/**
 * Получаем массив стилей в виде классов Asc.CAscStyle
 * @returns {Asc.CAscStyle[]}
 */
CStyles.prototype.GetAscStylesArray = function()
{
	var arrStyles = [];
	for (var sId in this.Style)
	{
		arrStyles.push(this.Style[sId].ToAscStyle());
	}

	return arrStyles;
};
/**
 * Получаем ссылку на класс документа
 * @returns {?CDocument}
 */
CStyles.prototype.private_GetLogicDocument = function()
{
	return (editor && editor.WordControl && editor.WordControl.m_oLogicDocument ? editor.WordControl.m_oLogicDocument : null);
};
CStyles.prototype.GetLogicDocument = function()
{
	return this.LogicDocument;
};
CStyles.prototype.Document_Is_SelectionLocked = function(CheckType)
{
	switch ( CheckType )
	{
		case AscCommon.changestype_Paragraph_Content:
		case AscCommon.changestype_Paragraph_Properties:
		case AscCommon.changestype_Paragraph_AddText:
		case AscCommon.changestype_Paragraph_TextProperties:
		case AscCommon.changestype_ContentControl_Add:
		case AscCommon.changestype_Document_Content:
		case AscCommon.changestype_Document_Content_Add:
		case AscCommon.changestype_Image_Properties:
		case AscCommon.changestype_Remove:
		case AscCommon.changestype_Delete:
		case AscCommon.changestype_Document_SectPr:
		case AscCommon.changestype_Table_Properties:
		case AscCommon.changestype_Table_RemoveCells:
		case AscCommon.changestype_HdrFtr:
		{
			AscCommon.CollaborativeEditing.Add_CheckLock(true);
			break;
		}
	}
};
CStyles.prototype.UpdateDefaultsDependingOnCompatibility = function(nCompatibilityMode)
{
	this.Default.ParaPr.InitDefault(nCompatibilityMode);
	this.Default.TextPr.InitDefault(nCompatibilityMode);
	this.Default.TablePr.InitDefault(nCompatibilityMode);
	this.Default.TableRowPr.InitDefault(nCompatibilityMode);
	this.Default.TableCellPr.InitDefault(nCompatibilityMode);

	g_oDocumentDefaultTextPr.InitDefault(nCompatibilityMode);
	g_oDocumentDefaultParaPr.InitDefault(nCompatibilityMode);
	g_oDocumentDefaultTablePr.InitDefault(nCompatibilityMode);
	g_oDocumentDefaultTableCellPr.InitDefault(nCompatibilityMode);
	g_oDocumentDefaultTableRowPr.InitDefault(nCompatibilityMode);
	g_oDocumentDefaultTableStylePr.InitDefault(nCompatibilityMode);

	this.OnChangeDefaultTextPr();
};
CStyles.prototype.GetDefaultParaPrForWrite = function()
{
	var oParaPr = new CParaPr();
	oParaPr.InitDefault();
	return this.Default.ParaPr.GetDiff(oParaPr);
};
CStyles.prototype.GetDefaultTextPrForWrite = function()
{
	var oTextPr = new CTextPr();
	oTextPr.InitDefault();
	return this.Default.TextPr.GetDiff(oTextPr);
};
CStyles.prototype.OnChangeDefaultTextPr = function()
{
	this.ValidDefaultEastAsiaFont = false;
	if (this.Default.TextPr.RFonts && this.Default.TextPr.RFonts.EastAsia)
		this.ValidDefaultEastAsiaFont = AscCommon.IsEastAsianFont(this.Default.TextPr.RFonts.EastAsia.Name);
};
CStyles.prototype.IsValidDefaultEastAsiaFont = function()
{
	return this.ValidDefaultEastAsiaFont;
};
CStyles.prototype.OnEndDocumentLoad = function(oDocument)
{
	// Специальная проверка плохо заданных нумераций через стиль. Когда ссылка на нумерацию в стиле есть,
	// а обратной ссылки в нумерации на стиль - нет.
	this.Check_StyleNumberingOnLoad(oDocument.GetNumbering());

	this.OnChangeDefaultTextPr();
};

function CDocumentColor(r,g,b, Auto)
{
    this.r = r;
    this.g = g;
    this.b = b;

    this.Auto = ( Auto === undefined ? false : Auto );
}

CDocumentColor.prototype =
{
    Copy : function()
    {
        return new CDocumentColor(this.r, this.g, this.b, this.Auto);
    },

    Write_ToBinary : function(Writer)
    {
        // Byte : r
        // Byte : g
        // Byte : b
        // Bool : Auto

        Writer.WriteByte( this.r );
        Writer.WriteByte( this.g );
        Writer.WriteByte( this.b );
        Writer.WriteBool( this.Auto );
    },

    Read_FromBinary : function(Reader)
    {
        // Byte : r
        // Byte : g
        // Byte : b
        // Bool : Auto

        this.r = Reader.GetByte();
        this.g = Reader.GetByte();
        this.b = Reader.GetByte();
        this.Auto = Reader.GetBool();
    },

    Set : function(r, g, b, Auto)
    {
        this.r = r;
        this.g = g;
        this.b = b;
        this.Auto = ( Auto === undefined ? false : Auto );
    },

    Check_BlackAutoColor : function()
    {
        // TODO: Коэффициенты подобраны опытным путем. В некоторых "пограничных" случаях
        //       может быть несовпадение с Word (когда изменение на 1 одного из каналов
        //       меняет цвет), чтобы такого не было надо более точно подобрать коэффициенты.
        if ( 0.5 * this.r + this.g + 0.195 * this.b < 103 )
            return false;

        return true;
    }
};
CDocumentColor.prototype.isBlackAutoColor = function()
{
	return this.Check_BlackAutoColor();
};
CDocumentColor.prototype.WriteToBinary = function(oWriter)
{
	this.Write_ToBinary(oWriter);
};
CDocumentColor.prototype.ReadFromBinary = function(oReader)
{
	this.Read_FromBinary(oReader);
};
CDocumentColor.prototype.IsEqual = function(oColor)
{
	if (!oColor || this.Auto !== oColor.Auto)
		return false;

	if (true === this.Auto)
		return true;

	return (this.r === oColor.r && this.g === oColor.g && this.b === oColor.b);
};
CDocumentColor.prototype.Compare = function(Color)
{
	return this.IsEqual(Color);
};
CDocumentColor.prototype.Is_Equal = function(Color)
{
	return this.IsEqual(Color);
};
CDocumentColor.prototype.IsAuto = function()
{
	return this.Auto;
};
CDocumentColor.prototype.IsEqualRGB = function(oColor)
{
	return (oColor && this.r === oColor.r && this.g === oColor.g && this.b === oColor.b);
};
CDocumentColor.prototype.SetFromColor = function(oColor)
{
	if (!oColor)
		return;

	this.r    = undefined !== oColor.r ? oColor.r : 0;
	this.g    = undefined !== oColor.g ? oColor.g : 0;
	this.b    = undefined !== oColor.b ? oColor.b : 0;
	this.Auto = undefined !== oColor.Auto ? oColor.Auto : false;
};
CDocumentColor.prototype.SetFromHexColor = function(val) {
	this.Auto = false;
	if ("auto" === val || "none" === val) {
		this.Auto = true;
	} else if (AscFormat.mapPrstColor[val]) {
		var rgb = AscFormat.mapPrstColor[val];
		this.r = (rgb >> 16) & 0xFF;
		this.g = (rgb >> 8) & 0xFF;
		this.b = rgb & 0xFF;
	} else if (6 <= val.length || 3 === val.length) {
		var rgba = AscCommon.RgbaHexToRGBA(val);
		this.r = rgba.R;
		this.g = rgba.G;
		this.b = rgba.B;
	} else {
		this.Auto = true;
	}
};
CDocumentColor.prototype.ToHexColor = function() {
	if (this.Auto) {
		return "auto";
	} else {
		return AscCommon.ByteToHex(this.r) + AscCommon.ByteToHex(this.g) + AscCommon.ByteToHex(this.b);
	}
};
CDocumentColor.prototype.ToHighlightColor = function()
{
	// 17.18.40 ST_HighlightColor
	let val = (((this.r & 0xFF) << 16) | ((this.g & 0xFF) << 8) | (this.b & 0xFF));
	
	switch (val)
	{
		case 0x000000: return "black";
		case 0x0000FF: return "blue";
		case 0x00FFFF: return "cyan";
		case 0x00008B: return "darkBlue";
		case 0x008B8B: return "darkCyan";
		case 0xA9A9A9: return "darkGray";
		case 0x006400: return "darkGreen";
		case 0x800080: return "darkMagenta";
		case 0x8B0000: return "darkRed";
		case 0x808000: return "darkYellow";
		case 0x00FF00: return "green";
		case 0xD3D3D3: return "lightGray";
		case 0xFF00FF: return "magenta";
		case 0xFF0000: return "red";
		case 0xFFFFFF: return "white";
		case 0xFFFF00: return "yellow";
	}
	
	return "";
};

CDocumentColor.prototype.ConvertToUniColor = function()
{
	return AscFormat.CreateUniColorRGB(this.r, this.g, this.b);
};
AscWord.CDocumentColor = CDocumentColor;

(function()
{
	/**
	 * @param r
	 * @param g
	 * @param b
	 * @param a
	 * @constructor
	 */
	function CDocumentColorA(r, g, b, a)
	{
		this.r = r;
		this.g = g;
		this.b = b;
		this.a = a;
	}
	CDocumentColorA.prototype.Copy = function()
	{
		return new CDocumentColorA(this.r, this.g, this.b, this.a);
	};
	CDocumentColorA.fromBinary = function(reader)
	{
		let r = reader.GetByte();
		let g = reader.GetByte();
		let b = reader.GetByte();
		let a = reader.GetByte();
		return new CDocumentColorA(r, g, b, a);
	};
	CDocumentColorA.fromObjectRgb = function(obj)
	{
		return new CDocumentColorA(obj.r, obj.g, obj.b, 255);
	};
	CDocumentColorA.fromObjectRgba = function(obj)
	{
		return new CDocumentColorA(obj.r, obj.g, obj.b, obj.a);
	};
	CDocumentColorA.prototype.toBinary = function(writer)
	{
		writer.WriteByte(this.r);
		writer.WriteByte(this.g);
		writer.WriteByte(this.b);
		writer.WriteByte(this.a);
	};
	CDocumentColorA.prototype.isEqual = function(color)
	{
		return (this.r === color.r
			&& this.g === color.g
			&& this.b === color.b
			&& this.a === color.a);
	};
	CDocumentColorA.prototype.isEqualRgb = function(color)
	{
		return (this.r === color.r
			&& this.g === color.g
			&& this.b === color.b
			&& 255 === this.a);
	};
	CDocumentColorA.prototype.setRgb = function(color)
	{
		this.r = color.r;
		this.g = color.g;
		this.b = color.b;
		this.a = 255;
	};
	CDocumentColorA.prototype.setRgba = function(color)
	{
		this.r = color.r;
		this.g = color.g;
		this.b = color.b;
		this.a = color.a;
	};
	/**
	 * @return {Asc.asc_CColor}
	 */
	CDocumentColorA.prototype.getAscColor = function()
	{
		return new Asc.asc_CColor(this.r, this.g, this.b, this.a);
	};
	CDocumentColorA.prototype.SetFromHexColor = function(val)
	{
		if (AscFormat.mapPrstColor[val])
		{
			let rgb = AscFormat.mapPrstColor[val];
			this.r = (rgb >> 16) & 0xFF;
			this.g = (rgb >> 8) & 0xFF;
			this.b = rgb & 0xFF;
			this.a = 255;
		}
		else if (val.length >= 8)
		{
			let index = "#" === val[0] ? 1 : 0;
			this.r = parseInt(val.substring(index, index + 2), 16);
			this.g = parseInt(val.substring(index + 2, index + 4), 16);
			this.b = parseInt(val.substring(index + 4, index + 6), 16);
			this.a = parseInt(val.substring(index + 6, index + 8), 16);
		}
		else
		{
			this.r = 0;
			this.g = 0;
			this.b = 0;
			this.a = 255;
		}
	};
	CDocumentColorA.prototype.ToHexColor = function()
	{
		return AscCommon.ByteToHex(this.r) + AscCommon.ByteToHex(this.g) + AscCommon.ByteToHex(this.b) + AscCommon.ByteToHex(this.a);
	};
	CDocumentColorA.prototype.WriteToBinary = function(writer)
	{
		this.toBinary(writer);
	};
	CDocumentColorA.prototype.ReadFromBinary = function(reader)
	{
		this.r = reader.GetByte();
		this.g = reader.GetByte();
		this.b = reader.GetByte();
		this.a = reader.GetByte();
	};
	CDocumentColorA.prototype.Write_ToBinary = function(writer)
	{
		return this.WriteToBinary(writer);
	};
	CDocumentColorA.prototype.Read_FromBinary = function(reader)
	{
		return this.ReadFromBinary(reader);
	};
	//------------------------------------------------------------------------------------------------------------------
	AscWord.CDocumentColorA = CDocumentColorA;
})();

function CDocumentShd()
{
	this.Value   = Asc.c_oAscShd.Nil;
	this.Color   = new CDocumentColor(255, 255, 255);
	this.Fill    = undefined;
	this.Unifill = undefined;
	this.FillRef = undefined;
	this.ThemeFill = undefined;

	// TODO:
	//  1. this.Color по умолчанию должен быть undefined
	//  2. Переименовать Unifill в ThemeColor
}

CDocumentShd.prototype =
{
    Get_Color : function(Paragraph)
    {
        if ( undefined !== this.Unifill )
        {
            this.Unifill.check(Paragraph.Get_Theme(), Paragraph.Get_ColorMap());
            var RGBA = this.Unifill.getRGBAColor();
            return new CDocumentColor( RGBA.R, RGBA.G, RGBA.B, false );
        }
		else if( undefined !== this.Fill )
		{
			return this.Fill;
		}
        else
            return this.Color;
    },

    Get_Color2 : function(Theme, ColorMap)
    {
        if ( undefined !== this.Unifill )
        {
            this.Unifill.check(Theme, ColorMap);
            var RGBA = this.Unifill.getRGBAColor();
            return new CDocumentColor( RGBA.R, RGBA.G, RGBA.B, false );
        }
		else if( undefined !== this.Fill )
		{
			return this.Fill;
		}
        else
            return this.Color;
    },

    Get_Color3 : function(Theme, ColorMap)
    {
        if ( undefined !== this.Unifill )
        {
            this.Unifill.check(Theme, ColorMap);
            return this.Unifill.getRGBAColor();
        }
        else
        {
            return {R: 255, G: 255, B: 255, A: 255};
        }
    },

    Check_PresentationPr : function(Theme)
    {
        if(this.FillRef && Theme)
        {
            this.Unifill = Theme.getFillStyle(this.FillRef.idx, this.FillRef.Color);
            this.FillRef = undefined;
        }
    }
};
CDocumentShd.fromObject = function(val)
{
	let shd = new CDocumentShd();
	shd.Set_FromObject(val);
	return shd;
};
CDocumentShd.FromObject = function(val)
{
	let shd = new CDocumentShd();
	shd.Set_FromObject(val);
	return shd;
};
CDocumentShd.prototype.Copy = function()
{
	var Shd = new CDocumentShd();

	Shd.Value = this.Value;

	if (undefined !== this.Color)
		Shd.Color.Set(this.Color.r, this.Color.g, this.Color.b, this.Color.Auto);

	if (undefined !== this.Unifill)
		Shd.Unifill = this.Unifill.createDuplicate();

	if (undefined !== this.FillRef)
		Shd.FillRef = this.FillRef.createDuplicate();

	if (undefined !== this.Fill)
		Shd.Fill = new CDocumentColor(this.Fill.r, this.Fill.g, this.Fill.b, this.Fill.Auto);

	if (undefined !== this.ThemeFill)
		Shd.ThemeFill = this.ThemeFill.createDuplicate();

	return Shd;
};
CDocumentShd.prototype.Compare = function(oShd)
{
	return this.IsEqual(oShd);
};
CDocumentShd.prototype.IsEqual = function(oShd)
{
	if (!oShd || this.Value !== oShd.Value)
		return false;

	if (Asc.c_oAscShd.Nil === this.Value)
		return true;

	return (IsEqualStyleObjects(this.Color, oShd.Color)
		&& IsEqualStyleObjects(this.Fill, oShd.Fill)
		&& IsEqualStyleObjects(this.Unifill, oShd.Unifill)
		&& IsEqualStyleObjects(this.ThemeFill, oShd.ThemeFill));
};
CDocumentShd.prototype.Is_Equal = function(Shd)
{
	return this.IsEqual(Shd);
};
CDocumentShd.prototype.InitDefault = function()
{
	this.Value     = Asc.c_oAscShd.Nil;
	this.Color     = new CDocumentColor(0, 0, 0, false);
	this.Unifill   = undefined;
	this.FillRef   = undefined;
	this.Fill      = undefined;
	this.ThemeFill = undefined;
};
CDocumentShd.prototype.Set_FromObject = function(oShd)
{
	if (!oShd)
	{
		this.Value = Asc.c_oAscShd.Nil;
		return;
	}

	this.Value = oShd.Value;

	if (Asc.c_oAscShd.Nil !== oShd.Value)
	{
		if (oShd.Color)
			this.Color = new CDocumentColor(oShd.Color.r, oShd.Color.g, oShd.Color.b, oShd.Color.Auto);

		if (oShd.ThemeColor)
			this.Unifill = oShd.ThemeColor.createDuplicate();
		else if (oShd.Unifill)
			this.Unifill = oShd.Unifill.createDuplicate();


		if (oShd.FillRef)
			this.FillRef = oShd.FillRef.createDuplicate();

		if (oShd.Fill)
			this.Fill = new CDocumentColor(oShd.Fill.r, oShd.Fill.g, oShd.Fill.b, oShd.Fill.Auto);

		if (oShd.ThemeFill)
			this.ThemeFill = oShd.ThemeFill.createDuplicate();
	}
	else if (oShd.Color)
	{
		this.Color = undefined;
	}
};
CDocumentShd.prototype.IsNil = function()
{
	return (Asc.c_oAscShd.Nil === this.Value);
};
CDocumentShd.prototype.GetSimpleColor = function(oTheme, oColorMap)
{
	var oFillColor   = g_oDocumentDefaultFillColor;
	var oStrokeColor = g_oDocumentDefaultStrokeColor;

	// TODO: Пока у нас неправильно работает сохранение и открытие в DOCX, поэтому считаем, что
	//       цвет, заданный в теме влияет на оба цвета, чтобы работало нормально в текущей схеме

	if (undefined !== this.ThemeFill)
	{
		if (oTheme && oColorMap)
			this.ThemeFill.check(oTheme, oColorMap);

		var RGBA = this.ThemeFill.getRGBAColor();
		oFillColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
	}
	else if (this.Fill && !this.Fill.IsAuto())
	{
		oFillColor = this.Fill;
	}

	if (undefined !== this.Unifill)
	{
		if (oTheme && oColorMap)
			this.Unifill.check(oTheme, oColorMap);

		var RGBA = this.Unifill.getRGBAColor();
		oStrokeColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
	}
	else if (this.Color && !this.Color.IsAuto())
	{
		oStrokeColor = this.Color;
	}

	var oResultColor;

	switch (this.Value)
	{
		case Asc.c_oAscShd.Clear:
		{
			oResultColor = oFillColor;
			break;
		}
		case Asc.c_oAscShd.Pct5:
		{
			oResultColor = this.private_GetPctShdColor(0.05, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct10:
		{
			oResultColor = this.private_GetPctShdColor(0.1, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct12:
		{
			oResultColor = this.private_GetPctShdColor(0.12, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct15:
		{
			oResultColor = this.private_GetPctShdColor(0.15, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct20:
		{
			oResultColor = this.private_GetPctShdColor(0.2, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct25:
		{
			oResultColor = this.private_GetPctShdColor(0.25, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct30:
		{
			oResultColor = this.private_GetPctShdColor(0.3, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct35:
		{
			oResultColor = this.private_GetPctShdColor(0.35, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct37:
		{
			oResultColor = this.private_GetPctShdColor(0.37, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct40:
		{
			oResultColor = this.private_GetPctShdColor(0.4, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct45:
		{
			oResultColor = this.private_GetPctShdColor(0.45, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct50:
		{
			oResultColor = this.private_GetPctShdColor(0.5, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct55:
		{
			oResultColor = this.private_GetPctShdColor(0.55, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct60:
		{
			oResultColor = this.private_GetPctShdColor(0.6, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct62:
		{
			oResultColor = this.private_GetPctShdColor(0.62, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct65:
		{
			oResultColor = this.private_GetPctShdColor(0.65, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct70:
		{
			oResultColor = this.private_GetPctShdColor(0.7, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct75:
		{
			oResultColor = this.private_GetPctShdColor(0.75, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct80:
		{
			oResultColor = this.private_GetPctShdColor(0.8, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct85:
		{
			oResultColor = this.private_GetPctShdColor(0.85, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct87:
		{
			oResultColor = this.private_GetPctShdColor(0.87, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct90:
		{
			oResultColor = this.private_GetPctShdColor(0.9, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.Pct95:
		{
			oResultColor = this.private_GetPctShdColor(0.95, oStrokeColor, oFillColor);
			break;
		}

		case Asc.c_oAscShd.DiagCross:
		{
			oResultColor = this.private_GetPctShdColor(0.75, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.ThinDiagStripe:
		case Asc.c_oAscShd.ThinHorzStripe:
		case Asc.c_oAscShd.ThinReverseDiagStripe:
		case Asc.c_oAscShd.ThinVertStripe:
		{
			oResultColor = this.private_GetPctShdColor(0.25, oStrokeColor, oFillColor);
			break;
		}
		case Asc.c_oAscShd.DiagStripe:
		case Asc.c_oAscShd.HorzCross:
		case Asc.c_oAscShd.HorzStripe:
		case Asc.c_oAscShd.ReverseDiagStripe:
		case Asc.c_oAscShd.ThinDiagCross:
		case Asc.c_oAscShd.ThinHorzCross:
		case Asc.c_oAscShd.VertStripe:
		{
			oResultColor = this.private_GetPctShdColor(0.5, oStrokeColor, oFillColor);
			break;
		}

		case Asc.c_oAscShd.Solid:
		{
			oResultColor = oStrokeColor;
			break;
		}

		default:
		{
			oResultColor = oFillColor;
			break;
		}
	}

	return oResultColor;
};
CDocumentShd.prototype.private_GetPctShdColor = function(nPct, strokeColor, fillColor)
{
	var _nPct = 1 - nPct;
	return new CDocumentColor(
		(strokeColor.r * nPct + fillColor.r * _nPct) | 0,
		(strokeColor.g * nPct + fillColor.g * _nPct) | 0,
		(strokeColor.b * nPct + fillColor.b * _nPct) | 0,
		false
	);
};
CDocumentShd.prototype.Write_ToBinary = function(Writer)
{
	// Byte : Value
	//
	// Если не Asc.c_oAscShd.Nil
	// Variable : Color

	Writer.WriteByte(this.Value);

	if (Asc.c_oAscShd.Nil !== this.Value)
	{
		if (this.Color)
		{
			Writer.WriteBool(true);
			this.Color.Write_ToBinary(Writer);
		}
		else
		{
			Writer.WriteBool(false);
		}

		if (this.Unifill)
		{
			Writer.WriteBool(true);
			this.Unifill.Write_ToBinary(Writer);
		}
		else
		{
			Writer.WriteBool(false);
		}

		if (this.FillRef)
		{
			Writer.WriteBool(true);
			this.FillRef.Write_ToBinary(Writer);
		}
		else
		{
			Writer.WriteBool(false);
		}

		if (this.Fill)
		{
			Writer.WriteBool(true);
			this.Fill.Write_ToBinary(Writer);
		}
		else
		{
			Writer.WriteBool(false);
		}

		if (this.ThemeFill)
		{
			Writer.WriteBool(true);
			this.ThemeFill.Write_ToBinary(Writer);
		}
		else
		{
			Writer.WriteBool(false);
		}
	}
};
CDocumentShd.prototype.Read_FromBinary = function(Reader)
{
	// Byte : Value
	//
	// Если не Asc.c_oAscShd.Nil
	// Variable : Color

	this.Value = Reader.GetByte();

	if (Asc.c_oAscShd.Nil !== this.Value)
	{
		if (Reader.GetBool())
		{
			this.Color = new CDocumentColor();
			this.Color.Read_FromBinary(Reader);
		}

		if (Reader.GetBool())
		{
			this.Unifill = new AscFormat.CUniFill();
			this.Unifill.Read_FromBinary(Reader);
		}

		if (Reader.GetBool())
		{
			this.FillRef = new AscFormat.StyleRef();
			this.FillRef.Read_FromBinary(Reader);
		}

		if (Reader.GetBool())
		{
			this.Fill = new CDocumentColor();
			this.Fill.Read_FromBinary(Reader);
		}

		if (Reader.GetBool())
		{
			this.ThemeFill = new AscFormat.CUniFill();
			this.ThemeFill.Read_FromBinary(Reader);
		}
	}
	else
	{
		this.Color = new CDocumentColor(0, 0, 0, false);
	}
};
CDocumentShd.prototype.WriteToBinary = function(oWriter)
{
	return this.Write_ToBinary(oWriter);
};
CDocumentShd.prototype.ReadFromBinary = function(oReader)
{
	return this.Read_FromBinary(oReader);
};

AscWord.BorderType = {
	// No Border
	none : -1,
	nil  : 0,
	
	// Line Border
	single                 : 1,
	thick                  : 2,
	double                 : 3,
	dotted                 : 4,
	dashed                 : 5,
	dotDash                : 6,
	dotDotDash             : 7,
	triple                 : 8,
	thinThickSmallGap      : 9,
	thickThinSmallGap      : 10,
	thinThickThinSmallGap  : 11,
	thinThickMediumGap     : 12,
	thickThinMediumGap     : 13,
	thinThickThinMediumGap : 14,
	thinThickLargeGap      : 15,
	thickThinLargeGap      : 16,
	thinThickThinLargeGap  : 17,
	wave                   : 18,
	doubleWave             : 19,
	dashSmallGap           : 20,
	dashDotStroked         : 21,
	threeDEmboss           : 22,
	threeDEngrave          : 23,
	outset                 : 24,
	inset                  : 25,
	
	// ArtBorder
	apples             : 101,
	archedScallops     : 102,
	babyPacifier       : 103,
	babyRattle         : 104,
	balloons3Colors    : 105,
	balloonsHotAir     : 106,
	basicBlackDashes   : 107,
	basicBlackDots     : 108,
	basicBlackSquares  : 109,
	basicThinLines     : 110,
	basicWhiteDashes   : 111,
	basicWhiteDots     : 112,
	basicWhiteSquares  : 113,
	basicWideInline    : 114,
	basicWideMidline   : 115,
	basicWideOutline   : 116,
	bats               : 117,
	birds              : 118,
	birdsFlight        : 119,
	cabins             : 120,
	cakeSlice          : 121,
	candyCorn          : 122,
	celticKnotwork     : 123,
	certificateBanner  : 124,
	chainLink          : 125,
	champagneBottle    : 126,
	checkedBarBlack    : 127,
	checkedBarColor    : 128,
	checkered          : 129,
	christmasTree      : 130,
	circlesLines       : 131,
	circlesRectangles  : 132,
	classicalWave      : 133,
	clocks             : 134,
	compass            : 135,
	confetti           : 136,
	confettiGrays      : 137,
	confettiOutline    : 138,
	confettiStreamers  : 139,
	confettiWhite      : 140,
	cornerTriangles    : 141,
	couponCutoutDashes : 142,
	couponCutoutDots   : 143,
	crazyMaze          : 144,
	creaturesButterfly : 145,
	creaturesFish      : 146,
	creaturesInsects   : 147,
	creaturesLadyBug   : 148,
	crossStitch        : 149,
	cup                : 150,
	custom             : 151,
	decoArch           : 152,
	decoArchColor      : 153,
	decoBlocks         : 154,
	diamondsGray       : 155,
	doubleD            : 156,
	doubleDiamonds     : 157,
	earth1             : 158,
	earth2             : 159,
	earth3             : 160,
	eclipsingSquares1  : 161,
	eclipsingSquares2  : 162,
	eggsBlack          : 163,
	fans               : 164,
	film               : 165,
	firecrackers       : 166,
	flowersBlockPrint  : 167,
	flowersDaisies     : 168,
	flowersModern1     : 169,
	flowersModern2     : 170,
	flowersPansy       : 171,
	flowersRedRose     : 172,
	flowersRoses       : 173,
	flowersTeacup      : 174,
	flowersTiny        : 175,
	gems               : 176,
	gingerbreadMan     : 177,
	gradient           : 178,
	handmade1          : 179,
	handmade2          : 180,
	heartBalloon       : 181,
	heartGray          : 182,
	hearts             : 183,
	heebieJeebies      : 184,
	holly              : 185,
	houseFunky         : 186,
	hypnotic           : 187,
	iceCreamCones      : 188,
	lightBulb          : 189,
	lightning1         : 190,
	lightning2         : 191,
	mapleLeaf          : 192,
	mapleMuffins       : 193,
	mapPins            : 194,
	marquee            : 195,
	marqueeToothed     : 196,
	moons              : 197,
	mosaic             : 198,
	musicNotes         : 199,
	northwest          : 200,
	ovals              : 201,
	packages           : 202,
	palmsBlack         : 203,
	palmsColor         : 204,
	paperClips         : 205,
	papyrus            : 206,
	partyFavor         : 207,
	partyGlass         : 208,
	pencils            : 209,
	people             : 210,
	peopleHats         : 211,
	peopleWaving       : 212,
	poinsettias        : 213,
	postageStamp       : 214,
	pumpkin1           : 215,
	pushPinNote1       : 216,
	pushPinNote2       : 217,
	pyramids           : 218,
	pyramidsAbove      : 219,
	quadrants          : 220,
	rings              : 221,
	safari             : 222,
	sawtooth           : 223,
	sawtoothGray       : 224,
	scaredCat          : 225,
	seattle            : 226,
	sharksTeeth        : 227,
	shadowedSquares    : 228,
	shapes1            : 229,
	shapes2            : 230,
	shorebirdTracks    : 231,
	skyrocket          : 232,
	snowflakeFancy     : 233,
	snowflakes         : 234,
	sombrero           : 235,
	southwest          : 236,
	stars              : 237,
	stars3d            : 238,
	starsBlack         : 239,
	starsShadowed      : 240,
	starsTop           : 241,
	sun                : 242,
	swirligig          : 243,
	tornPaper          : 244,
	tornPaperBlack     : 245,
	trees              : 246,
	triangle1          : 247,
	triangle2          : 248,
	triangleCircle1    : 249,
	triangleCircle2    : 250,
	triangleParty      : 251,
	triangles          : 252,
	twistedLines1      : 253,
	twistedLines2      : 254,
	vine               : 255,
	waveline           : 256,
	weavingAngles      : 257,
	weavingBraid       : 258,
	weavingRibbon      : 259,
	weavingStrips      : 260,
	whiteFlowers       : 261,
	woodwork           : 262,
	xIllusions         : 263,
	zanyTriangles      : 264,
	zigZag             : 265,
	zigZagStitch       : 266
};

// For compatibility
var border_None   = AscWord.BorderType.none;
var border_Single = AscWord.BorderType.single;

function CDocumentBorder()
{
	this.Color   = new CDocumentColor(0, 0, 0);
	this.Unifill = undefined;
	this.LineRef = undefined;
	this.Space   = 0;                      // Это значение учитывается всегда, даже когда Value = none (поэтому важно, что по умолчанию 0)
	this.Size    = 0.5 * g_dKoef_pt_to_mm; // Размер учитываем в зависимости от Value
	this.Value   = AscWord.BorderType.none;
}
CDocumentBorder.FromObject = function(obj)
{
	let border = new CDocumentBorder();
	border.Set_FromObject(obj);
	return border;
};
CDocumentBorder.prototype =
{
    Copy : function()
    {
        var Border = new CDocumentBorder();

        if ( undefined === this.Color )
            Border.Color = undefined;
        else
            Border.Color.Set(this.Color.r, this.Color.g, this.Color.b);

        if(undefined === this.Unifill)
            Border.Unifill = undefined;
        else
            Border.Unifill = this.Unifill.createDuplicate();

        if(undefined === this.LineRef)
        {
            Border.LineRef = undefined;
        }
        else
        {
            Border.LineRef = this.LineRef.createDuplicate();
        }

        if ( undefined === this.Space )
            Border.Space = undefined;
        else
            Border.Space = this.Space;

        if ( undefined === this.Size )
            Border.Size = undefined;
        else
            Border.Size  = this.Size;

        if ( undefined === this.Value )
            Border.Value = undefined;
        else
            Border.Value = this.Value;

        return Border;
    },

    Compare : function(Border)
    {
        if ( false === this.Color.Compare(Border.Color) )
            return false;

        if(AscFormat.CompareUnifillBool(this.Unifill, Border.Unifill) === false)
            return false;

        if(this.LineRef !== undefined && Border.LineRef === undefined || Border.LineRef !== undefined && this.LineRef === undefined)
            return false;

        if(this.LineRef !== undefined && !this.LineRef.compare(Border.LineRef))
        {
            return false;
        }

        if ( Math.abs( this.Size - Border.Size ) > 0.001 )
            return false;

        if ( Math.abs( this.Space - Border.Space ) > 0.001 )
            return false;

        if ( this.Value != Border.Value )
            return false;

        return true;
    },

    Is_Equal : function(Border)
    {
    	return this.IsEqual(Border);
    },

    Get_Color : function(Paragraph)
    {
        if ( undefined !== this.Unifill )
        {
            this.Unifill.check(Paragraph.Get_Theme(), Paragraph.Get_ColorMap());
            var RGBA = this.Unifill.getRGBAColor();
            return new CDocumentColor( RGBA.R, RGBA.G, RGBA.B, false );
        }
        else
            return this.Color;
    },

    Get_Color2 : function(Theme, ColorMap)
    {
        if ( undefined !== this.Unifill )
        {
            this.Unifill.check(Theme, ColorMap);
            var RGBA = this.Unifill.getRGBAColor();
            return new CDocumentColor( RGBA.R, RGBA.G, RGBA.B, false );
        }
        else
            return this.Color;
    },

    Check_PresentationPr : function(Theme)
    {
        if(this.LineRef && Theme)
        {
            var pen = Theme.getLnStyle(this.LineRef.idx, this.LineRef.Color);

            this.Unifill = pen.Fill;
            this.LineRef = undefined;
            this.Size = AscFormat.isRealNumber(pen.w) ? pen.w / 36000 : 12700 /36000;
        }
        if(!this.Unifill || !this.Unifill.isVisible())
        {
            this.Value = border_None;
        }
    },

    Set_FromObject : function(Border)
    {
        this.Space = Border.Space;
        this.Size  = Border.Size;
        this.Value = Border.Value;

        if ( undefined != Border.Color )
            this.Color = new CDocumentColor( Border.Color.r, Border.Color.g, Border.Color.b );
        else
            this.Color = undefined;

        if(undefined != Border.Unifill)
        {
            this.Unifill = Border.Unifill.createDuplicate();
        }
        if(undefined != Border.LineRef)
        {
            this.LineRef = Border.LineRef.createDuplicate();
        }
    },

    Check_Null : function()
    {
        if ( undefined === this.Space || undefined === this.Size || undefined === this.Value || undefined === this.Color || undefined === this.Unifill || undefined === this.LineRef)
            return false;

        return true;
    },

    Write_ToBinary : function(Writer)
    {
        // Double   : Size
        // Long     : Space
        // Long     : Value
        // Variable : Color

        Writer.WriteDouble( this.Size );
        Writer.WriteLong( this.Space );
        Writer.WriteLong( this.Value );
        this.Color.Write_ToBinary( Writer );
        if(this.Unifill)
        {
            Writer.WriteBool(true);
            this.Unifill.Write_ToBinary(Writer);
        }
        else
        {
            Writer.WriteBool(false);
        }
        if(this.LineRef)
        {
            Writer.WriteBool(true);
            this.LineRef.Write_ToBinary(Writer);
        }
        else
        {
            Writer.WriteBool(false);
        }

    },

    Read_FromBinary : function(Reader)
    {
        // Double   : Size
        // Long     : Space
        // Byte     : Value
        // Variable : Color

        this.Size  = Reader.GetDouble();
        this.Space = Reader.GetLong();
        this.Value = Reader.GetLong();
        this.Color.Read_FromBinary( Reader );
        if(Reader.GetBool())
        {
            this.Unifill = new AscFormat.CUniFill();
            this.Unifill.Read_FromBinary(Reader);
        }
        if(Reader.GetBool())
        {
            this.LineRef = new AscFormat.StyleRef();
            this.LineRef.Read_FromBinary(Reader);
        }
    }
};
CDocumentBorder.prototype.IsNone = function()
{
	return (this.Value === border_None);
};
CDocumentBorder.prototype.SetNone = function()
{
	this.Value = border_None;
};
CDocumentBorder.prototype.GetSize = function()
{
	return this.Size;
};
CDocumentBorder.prototype.setSizeIn8Point = function(val)
{
	if(null !== val && undefined !== val) {
		this.Size = g_dKoef_pt_to_mm * val / 8;
	}
};
CDocumentBorder.prototype.getSizeIn8Point = function()
{
	return undefined !== this.Size ? Math.round(8 * this.Size * g_dKoef_mm_to_pt) : undefined;
};
CDocumentBorder.prototype.setSpaceInPoint = function(val)
{
	if(null !== val && undefined !== val) {
		this.Space = g_dKoef_pt_to_mm * val;
	}
};
CDocumentBorder.prototype.getSpaceInPoint = function(val)
{
	return undefined !== this.Space ? Math.round(this.Space * g_dKoef_mm_to_pt) : undefined;
};
CDocumentBorder.prototype.setValue = function(value)
{
	// Пока мы работаем только с такими типами
	if (value === AscWord.BorderType.nil || value === AscWord.BorderType.none)
		this.Value = AscWord.BorderType.none;
	else
		this.Value = AscWord.BorderType.single;
};
CDocumentBorder.prototype.getValue = function()
{
	return this.Value;
};
/**
 * Получаем рассчитанную толщину линии в зависимости от типа.
 * @returns {number}
 */
CDocumentBorder.prototype.GetWidth = function()
{
	if (border_None === this.Value)
		return 0;

	return this.Size;
};
CDocumentBorder.prototype.SetSimpleColor = function(r, g, b)
{
	this.Color   = new CDocumentColor(r, g, b);
	this.Unifill = undefined;
	this.LineRef = undefined;
};
CDocumentBorder.prototype.GetColor = function(oParagraph)
{
	return this.Get_Color(oParagraph);
};
CDocumentBorder.prototype.IsEqual = function(oBorder)
{
	if (!oBorder || this.Value !== oBorder.Value)
		return false;

	if (this.IsNone())
		return true;

	return (IsEqualStyleObjects(this.Color, oBorder.Color)
		&& IsEqualStyleObjects(this.Unifill, oBorder.Unifill)
		&& this.Space === oBorder.Space
		&& this.Size === oBorder.Size);
};
CDocumentBorder.prototype.WriteToBinary = function(oWriter)
{
	return this.Write_ToBinary(oWriter);
};
CDocumentBorder.prototype.ReadFromBinary = function(oReader)
{
	return this.Read_FromBinary(oReader);
};

function CTableMeasurement(Type, W)
{
    this.Type = Type;
    this.W    = W;
}

CTableMeasurement.prototype =
{
    Copy : function()
    {
        return new CTableMeasurement(this.Type, this.W);
    },

    Is_Equal : function(Other)
    {
        if (this.Type !== Other.Type
            || this.W !== Other.W)
            return false;

        return true;
    },

    Write_ToBinary : function(Writer)
    {
    	this.WriteToBinary(Writer);
    },

    Read_FromBinary : function(Reader)
    {
    	return this.ReadFromBinary(Reader);
    },

    Set_FromObject : function(Obj)
    {
        this.W    = Obj.W;
        this.Type = Obj.Type;
    }
};
/**
 * Измерение в миллиметрах?
 * @returns {boolean}
 */
CTableMeasurement.prototype.IsMM = function()
{
	return (tblwidth_Mm === this.Type);
};
/**
 * Измерение в процентах?
 * @returns {boolean}
 */
CTableMeasurement.prototype.IsPercent = function()
{
	return (tblwidth_Pct === this.Type);
};
/**
 * Ширина должна быть расчитана автоматически
 * @returns {boolean}
 */
CTableMeasurement.prototype.IsAuto = function()
{
	return (tblwidth_Auto === this.Type);
};
/**
 * Получаем значение ширины в процентах или в миллиметрах
 * @returns {number}
 */
CTableMeasurement.prototype.GetValue = function()
{
	return this.W;
};
/**
 * Выставляем значение ширины
 * @param nValue {number}
 */
CTableMeasurement.prototype.SetValue = function(nValue)
{
	this.W = nValue;
};
/**
 * Расчитываем ширину с учетом заданного типа
 * @param {number} nFullWidth - ширина родительского элемента для расчета значения процентах
 */
CTableMeasurement.prototype.GetCalculatedValue = function(nFullWidth)
{
	if (this.IsMM())
		return this.W;
	else if (this.IsPercent())
		return this.W * nFullWidth / 100;

	return 0;
};
CTableMeasurement.prototype.ReadFromBinary = function(oReader)
{
	// Double : W
	// Long   : Type

	this.W    = oReader.GetDouble();
	this.Type = oReader.GetLong();
};
CTableMeasurement.prototype.WriteToBinary = function(oWriter)
{
	// Double : W
	// Long   : Type
	oWriter.WriteDouble(this.W);
	oWriter.WriteLong(this.Type);
};
CTableMeasurement.prototype.SetValueByType = function(val)
{
	if (null !== val && undefined !== val) {
		if (tblwidth_Mm == this.Type) {
			this.W = g_dKoef_twips_to_mm * val;
		} else if (tblwidth_Pct == this.Type) {
			this.W = 2 * val / 100;
		} else {
			this.W = val;
		}
	}
};
CTableMeasurement.prototype.GetValueByType = function() {
	var res = this.W;
	if (null !== res && undefined !== res) {
		if (tblwidth_Mm == this.Type) {
			res = Math.round(AscCommonWord.g_dKoef_mm_to_twips * this.W);
		} else if (tblwidth_Pct == this.Type) {
			res = Math.round(100 * this.W / 2);
		}
	}
	return res;
};

function CTablePr()
{
    this.TableStyleColBandSize = undefined;
    this.TableStyleRowBandSize = undefined;
    this.Jc                    = undefined;
    this.Shd                   = undefined;
    this.TableBorders          =
    {
        Bottom  : undefined,
        Left    : undefined,
        Right   : undefined,
        Top     : undefined,
        InsideH : undefined,
        InsideV : undefined
    };
    this.TableCellMar          =
    {
        Bottom : undefined,
        Left   : undefined,
        Right  : undefined,
        Top    : undefined
    };
    this.TableCellSpacing      = undefined;
    this.TableInd              = undefined;
    this.TableW                = undefined;
    this.TableLayout           = undefined;
    this.TableDescription      = undefined;
    this.TableCaption          = undefined;
    this.PrChange              = undefined;
    this.ReviewInfo            = undefined;
}

CTablePr.prototype.Copy = function(bCopyPrChange)
{
	var TablePr = new CTablePr();

	TablePr.TableStyleColBandSize = this.TableStyleColBandSize;
	TablePr.TableStyleRowBandSize = this.TableStyleRowBandSize;
	TablePr.Jc                    = this.Jc;

	if (undefined != this.Shd)
		TablePr.Shd = this.Shd.Copy();

	// TableBorders
	if (undefined != this.TableBorders.Bottom)
		TablePr.TableBorders.Bottom = this.TableBorders.Bottom.Copy();

	if (undefined != this.TableBorders.Left)
		TablePr.TableBorders.Left = this.TableBorders.Left.Copy();

	if (undefined != this.TableBorders.Right)
		TablePr.TableBorders.Right = this.TableBorders.Right.Copy();

	if (undefined != this.TableBorders.Top)
		TablePr.TableBorders.Top = this.TableBorders.Top.Copy();

	if (undefined != this.TableBorders.InsideH)
		TablePr.TableBorders.InsideH = this.TableBorders.InsideH.Copy();

	if (undefined != this.TableBorders.InsideV)
		TablePr.TableBorders.InsideV = this.TableBorders.InsideV.Copy();

	// TableCellMar
	if (undefined != this.TableCellMar.Bottom)
		TablePr.TableCellMar.Bottom = this.TableCellMar.Bottom.Copy();

	if (undefined != this.TableCellMar.Left)
		TablePr.TableCellMar.Left = this.TableCellMar.Left.Copy();

	if (undefined != this.TableCellMar.Right)
		TablePr.TableCellMar.Right = this.TableCellMar.Right.Copy();

	if (undefined != this.TableCellMar.Top)
		TablePr.TableCellMar.Top = this.TableCellMar.Top.Copy();

	TablePr.TableCellSpacing = this.TableCellSpacing;
	TablePr.TableInd         = this.TableInd;

	if (undefined != this.TableW)
		TablePr.TableW = this.TableW.Copy();

	TablePr.TableLayout = this.TableLayout;

	TablePr.TableDescription = this.TableDescription;
	TablePr.TableCaption     = this.TableCaption;

	if (true === bCopyPrChange && undefined !== this.PrChange)
	{
		TablePr.PrChange   = this.PrChange.Copy();
		TablePr.ReviewInfo = this.ReviewInfo.Copy();
	}

	return TablePr;
};
CTablePr.prototype.Merge = function(TablePr)
{
	if (undefined != TablePr.TableStyleColBandSize)
		this.TableStyleColBandSize = TablePr.TableStyleColBandSize;

	if (undefined != TablePr.TableStyleRowBandSize)
		this.TableStyleRowBandSize = TablePr.TableStyleRowBandSize;

	if (undefined != TablePr.Jc)
		this.Jc = TablePr.Jc;

	if (undefined != TablePr.Shd)
		this.Shd = TablePr.Shd.Copy();

	// TableBorders
	if (undefined != TablePr.TableBorders.Bottom)
		this.TableBorders.Bottom = TablePr.TableBorders.Bottom.Copy();

	if (undefined != TablePr.TableBorders.Left)
		this.TableBorders.Left = TablePr.TableBorders.Left.Copy();

	if (undefined != TablePr.TableBorders.Right)
		this.TableBorders.Right = TablePr.TableBorders.Right.Copy();

	if (undefined != TablePr.TableBorders.Top)
		this.TableBorders.Top = TablePr.TableBorders.Top.Copy();

	if (undefined != TablePr.TableBorders.InsideH)
		this.TableBorders.InsideH = TablePr.TableBorders.InsideH.Copy();

	if (undefined != TablePr.TableBorders.InsideV)
		this.TableBorders.InsideV = TablePr.TableBorders.InsideV.Copy();

	// TableCellMar
	if (undefined != TablePr.TableCellMar.Bottom)
		this.TableCellMar.Bottom = TablePr.TableCellMar.Bottom.Copy();

	if (undefined != TablePr.TableCellMar.Left)
		this.TableCellMar.Left = TablePr.TableCellMar.Left.Copy();

	if (undefined != TablePr.TableCellMar.Right)
		this.TableCellMar.Right = TablePr.TableCellMar.Right.Copy();

	if (undefined != TablePr.TableCellMar.Top)
		this.TableCellMar.Top = TablePr.TableCellMar.Top.Copy();

	if (undefined != TablePr.TableCellSpacing)
		this.TableCellSpacing = TablePr.TableCellSpacing;

	if (undefined != TablePr.TableInd)
		this.TableInd = TablePr.TableInd;

	if (undefined != TablePr.TableW)
		this.TableW = TablePr.TableW.Copy();

	if (undefined != TablePr.TableLayout)
		this.TableLayout = TablePr.TableLayout;

	if (undefined !== TablePr.TableDescription)
		this.TableDescription = TablePr.TableDescription;

	if (undefined !== TablePr.TableCaption)
		this.TableCaption = TablePr.TableCaption;
};
CTablePr.prototype.Is_Equal = function(TablePr)
{
	if (this.TableStyleColBandSize !== TablePr.TableStyleColBandSize
		|| this.TableStyleRowBandSize !== TablePr.TableStyleRowBandSize
		|| this.Jc !== TablePr.Jc
		|| true !== IsEqualStyleObjects(this.TableBorders.Bottom, TablePr.TableBorders.Bottom)
		|| true !== IsEqualStyleObjects(this.TableBorders.Left, TablePr.TableBorders.Left)
		|| true !== IsEqualStyleObjects(this.TableBorders.Right, TablePr.TableBorders.Right)
		|| true !== IsEqualStyleObjects(this.TableBorders.Top, TablePr.TableBorders.Top)
		|| true !== IsEqualStyleObjects(this.TableBorders.InsideH, TablePr.TableBorders.InsideH)
		|| true !== IsEqualStyleObjects(this.TableBorders.InsideV, TablePr.TableBorders.InsideV)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Bottom, TablePr.TableCellMar.Bottom)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Left, TablePr.TableCellMar.Left)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Right, TablePr.TableCellMar.Right)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Top, TablePr.TableCellMar.Top)
		|| this.TableCellSpacing !== TablePr.TableCellSpacing
		|| this.TableInd !== TablePr.TableInd
		|| true !== IsEqualStyleObjects(this.TableW, TablePr.TableW)
		|| this.TableLayout !== TablePr.TableLayout)
		return false;

	return true;
};
CTablePr.prototype.InitDefault = function(nCompatibilityMode)
{
	if (undefined === nCompatibilityMode)
		nCompatibilityMode = AscCommon.document_compatibility_mode_Word12;

	this.TableStyleColBandSize = 1;
	this.TableStyleRowBandSize = 1;
	this.Jc                    = align_Left;
	this.Shd                   = new CDocumentShd();
	this.TableBorders.Bottom   = new CDocumentBorder();
	this.TableBorders.Left     = new CDocumentBorder();
	this.TableBorders.Right    = new CDocumentBorder();
	this.TableBorders.Top      = new CDocumentBorder();
	this.TableBorders.InsideH  = new CDocumentBorder();
	this.TableBorders.InsideV  = new CDocumentBorder();
	this.TableCellMar.Bottom   = new CTableMeasurement(tblwidth_Mm, 0);
	this.TableCellMar.Left     = nCompatibilityMode <= AscCommon.document_compatibility_mode_Word12 ? new CTableMeasurement(tblwidth_Mm, 0.5 * g_dKoef_pt_to_mm) : new CTableMeasurement(tblwidth_Mm, 1.9/*5.4 * g_dKoef_pt_to_mm*/); // 5.4pt
	this.TableCellMar.Right    = nCompatibilityMode <= AscCommon.document_compatibility_mode_Word12 ? new CTableMeasurement(tblwidth_Mm, 0.5 * g_dKoef_pt_to_mm) : new CTableMeasurement(tblwidth_Mm, 1.9/*5.4 * g_dKoef_pt_to_mm*/); // 5.4pt
	this.TableCellMar.Top      = new CTableMeasurement(tblwidth_Mm, 0);
	this.TableCellSpacing      = null;
	this.TableInd              = 0;
	this.TableW                = new CTableMeasurement(tblwidth_Auto, 0);
	this.TableLayout           = tbllayout_AutoFit;
	this.TableDescription      = "";
	this.TableCaption          = "";
	this.PrChange              = undefined;
	this.ReviewInfo            = undefined;
};
CTablePr.prototype.Set_FromObject = function(TablePr)
{
	this.TableStyleColBandSize = TablePr.TableStyleColBandSize;
	this.TableStyleRowBandSize = TablePr.TableStyleRowBandSize;
	this.Jc                    = TablePr.Jc;

	if (undefined != TablePr.Shd)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Set_FromObject(TablePr.Shd);
	}
	else
		this.Shd = undefined;

	if (undefined != TablePr.TableBorders)
	{
		if (undefined != TablePr.TableBorders.Bottom)
		{
			this.TableBorders.Bottom = new CDocumentBorder();
			this.TableBorders.Bottom.Set_FromObject(TablePr.TableBorders.Bottom);
		}
		else
			this.TableBorders.Bottom = undefined;

		if (undefined != TablePr.TableBorders.Left)
		{
			this.TableBorders.Left = new CDocumentBorder();
			this.TableBorders.Left.Set_FromObject(TablePr.TableBorders.Left);
		}
		else
			this.TableBorders.Left = undefined;

		if (undefined != TablePr.TableBorders.Right)
		{
			this.TableBorders.Right = new CDocumentBorder();
			this.TableBorders.Right.Set_FromObject(TablePr.TableBorders.Right);
		}
		else
			this.TableBorders.Right = undefined;

		if (undefined != TablePr.TableBorders.Top)
		{
			this.TableBorders.Top = new CDocumentBorder();
			this.TableBorders.Top.Set_FromObject(TablePr.TableBorders.Top);
		}
		else
			this.TableBorders.Top = undefined;

		if (undefined != TablePr.TableBorders.InsideH)
		{
			this.TableBorders.InsideH = new CDocumentBorder();
			this.TableBorders.InsideH.Set_FromObject(TablePr.TableBorders.InsideH);
		}
		else
			this.TableBorders.InsideH = undefined;

		if (undefined != TablePr.TableBorders.InsideV)
		{
			this.TableBorders.InsideV = new CDocumentBorder();
			this.TableBorders.InsideV.Set_FromObject(TablePr.TableBorders.InsideV);
		}
		else
			this.TableBorders.InsideV = undefined;
	}
	else
	{
		this.TableBorders.Bottom  = undefined;
		this.TableBorders.Left    = undefined;
		this.TableBorders.Right   = undefined;
		this.TableBorders.Top     = undefined;
		this.TableBorders.InsideH = undefined;
		this.TableBorders.InsideV = undefined;
	}

	if (undefined != TablePr.TableCellMar)
	{
		if (undefined != TablePr.TableCellMar.Bottom)
			this.TableCellMar.Bottom = new CTableMeasurement(TablePr.TableCellMar.Bottom.Type, TablePr.TableCellMar.Bottom.W);
		else
			this.TableCellMar.Bottom = undefined;

		if (undefined != TablePr.TableCellMar.Left)
			this.TableCellMar.Left = new CTableMeasurement(TablePr.TableCellMar.Left.Type, TablePr.TableCellMar.Left.W);
		else
			this.TableCellMar.Left = undefined;

		if (undefined != TablePr.TableCellMar.Right)
			this.TableCellMar.Right = new CTableMeasurement(TablePr.TableCellMar.Right.Type, TablePr.TableCellMar.Right.W);
		else
			this.TableCellMar.Right = undefined;

		if (undefined != TablePr.TableCellMar.Top)
			this.TableCellMar.Top = new CTableMeasurement(TablePr.TableCellMar.Top.Type, TablePr.TableCellMar.Top.W);
		else
			this.TableCellMar.Top = undefined;
	}
	else
	{
		this.TableCellMar.Bottom = undefined;
		this.TableCellMar.Left   = undefined;
		this.TableCellMar.Right  = undefined;
		this.TableCellMar.Top    = undefined;
	}

	this.TableCellSpacing = TablePr.TableCellSpacing;
	this.TableInd         = TablePr.TableInd;

	if (undefined != TablePr.TableW)
		this.TableW = new CTableMeasurement(TablePr.TableW.Type, TablePr.TableW.W);
	else
		this.TableW = undefined;

	this.TableLayout = TablePr.TableLayout;

	this.TableDescription = TablePr.TableDescription;
	this.TableCaption     = TablePr.TableCaption;
};
CTablePr.prototype.Check_PresentationPr = function(Theme)
{
	if (this.Shd)
	{
		this.Shd.Check_PresentationPr(Theme);
	}
	if (this.TableBorders.Bottom)
	{
		this.TableBorders.Bottom.Check_PresentationPr(Theme);
	}
	if (this.TableBorders.Left)
	{
		this.TableBorders.Left.Check_PresentationPr(Theme);
	}
	if (this.TableBorders.Right)
	{
		this.TableBorders.Right.Check_PresentationPr(Theme);
	}
	if (this.TableBorders.Top)
	{
		this.TableBorders.Top.Check_PresentationPr(Theme);
	}
	if (this.TableBorders.InsideH)
	{
		this.TableBorders.InsideH.Check_PresentationPr(Theme);
	}
	if (this.TableBorders.InsideV)
	{
		this.TableBorders.InsideV.Check_PresentationPr(Theme);
	}
};
CTablePr.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);
	var Flags = 0;

	if (undefined != this.TableStyleColBandSize)
	{
		Writer.WriteLong(this.TableStyleColBandSize);
		Flags |= 1;
	}

	if (undefined != this.TableStyleRowBandSize)
	{
		Writer.WriteLong(this.TableStyleRowBandSize);
		Flags |= 2;
	}

	if (undefined != this.Jc)
	{
		Writer.WriteLong(this.Jc);
		Flags |= 4;
	}

	if (undefined != this.Shd)
	{
		this.Shd.Write_ToBinary(Writer);
		Flags |= 8;
	}

	if (undefined != this.TableBorders.Bottom)
	{
		this.TableBorders.Bottom.Write_ToBinary(Writer);
		Flags |= 16;
	}

	if (undefined != this.TableBorders.Left)
	{
		this.TableBorders.Left.Write_ToBinary(Writer);
		Flags |= 32;
	}

	if (undefined != this.TableBorders.Right)
	{
		this.TableBorders.Right.Write_ToBinary(Writer);
		Flags |= 64;
	}

	if (undefined != this.TableBorders.Top)
	{
		this.TableBorders.Top.Write_ToBinary(Writer);
		Flags |= 128;
	}

	if (undefined != this.TableBorders.InsideH)
	{
		this.TableBorders.InsideH.Write_ToBinary(Writer);
		Flags |= 256;
	}

	if (undefined != this.TableBorders.InsideV)
	{
		this.TableBorders.InsideV.Write_ToBinary(Writer);
		Flags |= 512;
	}

	if (undefined != this.TableCellMar.Bottom)
	{
		this.TableCellMar.Bottom.Write_ToBinary(Writer);
		Flags |= 1024;
	}

	if (undefined != this.TableCellMar.Left)
	{
		this.TableCellMar.Left.Write_ToBinary(Writer);
		Flags |= 2048;
	}

	if (undefined != this.TableCellMar.Right)
	{
		this.TableCellMar.Right.Write_ToBinary(Writer);
		Flags |= 4096;
	}

	if (undefined != this.TableCellMar.Top)
	{
		this.TableCellMar.Top.Write_ToBinary(Writer);
		Flags |= 8192;
	}

	if (undefined != this.TableCellSpacing)
	{
		if (null === this.TableCellSpacing)
			Writer.WriteBool(true);
		else
		{
			Writer.WriteBool(false);
			Writer.WriteDouble(this.TableCellSpacing);
		}

		Flags |= 16384;
	}

	if (undefined != this.TableInd)
	{
		Writer.WriteDouble(this.TableInd);

		Flags |= 32768;
	}

	if (undefined != this.TableW)
	{
		this.TableW.Write_ToBinary(Writer);
		Flags |= 65536;
	}

	if (undefined != this.TableLayout)
	{
		Writer.WriteLong(this.TableLayout);
		Flags |= 131072;
	}

	if (undefined !== this.TableDescription)
	{
		Writer.WriteString2(this.TableDescription);
		Flags |= 262144;
	}

	if (undefined !== this.TableCaption)
	{
		Writer.WriteString2(this.TableCaption);
		Flags |= 524288;
	}

	if (undefined !== this.PrChange && undefined !== this.ReviewInfo)
	{
		this.PrChange.WriteToBinary(Writer);
		this.ReviewInfo.WriteToBinary(Writer);
		Flags |= 1048576;
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CTablePr.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	if (1 & Flags)
		this.TableStyleColBandSize = Reader.GetLong();

	if (2 & Flags)
		this.TableStyleRowBandSize = Reader.GetLong();

	if (4 & Flags)
		this.Jc = Reader.GetLong();

	if (8 & Flags)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Read_FromBinary(Reader);
	}

	if (16 & Flags)
	{
		this.TableBorders.Bottom = new CDocumentBorder();
		this.TableBorders.Bottom.Read_FromBinary(Reader);
	}

	if (32 & Flags)
	{
		this.TableBorders.Left = new CDocumentBorder();
		this.TableBorders.Left.Read_FromBinary(Reader);
	}

	if (64 & Flags)
	{
		this.TableBorders.Right = new CDocumentBorder();
		this.TableBorders.Right.Read_FromBinary(Reader);
	}

	if (128 & Flags)
	{
		this.TableBorders.Top = new CDocumentBorder();
		this.TableBorders.Top.Read_FromBinary(Reader);
	}

	if (256 & Flags)
	{
		this.TableBorders.InsideH = new CDocumentBorder();
		this.TableBorders.InsideH.Read_FromBinary(Reader);
	}

	if (512 & Flags)
	{
		this.TableBorders.InsideV = new CDocumentBorder();
		this.TableBorders.InsideV.Read_FromBinary(Reader);
	}

	if (1024 & Flags)
	{
		this.TableCellMar.Bottom = new CTableMeasurement(tblwidth_Auto, 0);
		this.TableCellMar.Bottom.Read_FromBinary(Reader);
	}

	if (2048 & Flags)
	{
		this.TableCellMar.Left = new CTableMeasurement(tblwidth_Auto, 0);
		this.TableCellMar.Left.Read_FromBinary(Reader);
	}

	if (4096 & Flags)
	{
		this.TableCellMar.Right = new CTableMeasurement(tblwidth_Auto, 0);
		this.TableCellMar.Right.Read_FromBinary(Reader);
	}

	if (8192 & Flags)
	{
		this.TableCellMar.Top = new CTableMeasurement(tblwidth_Auto, 0);
		this.TableCellMar.Top.Read_FromBinary(Reader);
	}

	if (16384 & Flags)
	{
		if (true === Reader.GetBool())
			this.TableCellSpacing = null;
		else
			this.TableCellSpacing = Reader.GetDouble()
	}

	if (32768 & Flags)
		this.TableInd = Reader.GetDouble();

	if (65536 & Flags)
	{
		this.TableW = new CTableMeasurement(tblwidth_Auto, 0);
		this.TableW.Read_FromBinary(Reader);
	}

	if (131072 & Flags)
		this.TableLayout = Reader.GetLong();

	if (262144 & Flags)
		this.TableDescription = Reader.GetString2();

	if (524288 & Flags)
		this.TableCaption = Reader.GetString2();

	if (1048576 & Flags)
	{
		this.PrChange   = new CTablePr();
		this.ReviewInfo = new AscWord.ReviewInfo();

		this.PrChange.ReadFromBinary(Reader);
		this.ReviewInfo.ReadFromBinary(Reader);
	}
};
CTablePr.prototype.WriteToBinary = function(oWriter)
{
	this.Write_ToBinary(oWriter);
};
CTablePr.prototype.ReadFromBinary = function(oReader)
{
	this.Read_FromBinary(oReader);
};
CTablePr.prototype.HavePrChange = function()
{
	if (undefined === this.PrChange || null === this.PrChange)
		return false;

	return true;
};
CTablePr.prototype.AddPrChange = function()
{
	this.PrChange   = this.Copy(false);
	this.ReviewInfo = new AscWord.ReviewInfo();
	this.ReviewInfo.Update();
};
CTablePr.prototype.SetPrChange = function(oPrChange, oReviewInfo)
{
	this.PrChange   = oPrChange;
	this.ReviewInfo = oReviewInfo;
};
CTablePr.prototype.RemovePrChange = function()
{
	delete this.PrChange;
	delete this.ReviewInfo;
};

function CTableRowHeight(Value, HRule)
{
    this.Value = Value;
    this.HRule = HRule;
}

CTableRowHeight.prototype =
{
    Copy : function()
    {
        return new CTableRowHeight(this.Value, this.HRule);
    },

    Is_Equal : function(Other)
    {
        if (this.Value !== Other.Value
            || this.HRule !== Other.HRule)
            return false;

        return true;
    },

    Write_ToBinary : function(Writer)
    {
        // Double : Value
        // Long   : HRule
        Writer.WriteDouble( this.Value );
        Writer.WriteLong( this.HRule );
    },

    Read_FromBinary : function(Reader)
    {
        // Double : Value
        // Long   : HRule

        this.Value = Reader.GetDouble();
        this.HRule = Reader.GetLong();
    }
};
CTableRowHeight.prototype.IsAuto = function()
{
	return (this.HRule === Asc.linerule_Auto ? true : false);
};
CTableRowHeight.prototype.GetValue = function()
{
	return this.Value;
};
CTableRowHeight.prototype.GetRule = function()
{
	return this.HRule;
};

function CTableRowPr()
{
    this.CantSplit        = undefined;
    this.GridAfter        = undefined;
    this.GridBefore       = undefined;
    this.Jc               = undefined;
    this.TableCellSpacing = undefined;
    this.Height           = undefined;
    this.WAfter           = undefined;
    this.WBefore          = undefined;
    this.TableHeader      = undefined;
    this.PrChange         = undefined;
    this.ReviewInfo       = undefined;
}

CTableRowPr.prototype.Copy = function(bCopyPrChange)
{
	var RowPr = new CTableRowPr();

	RowPr.CantSplit        = this.CantSplit;
	RowPr.GridAfter        = this.GridAfter;
	RowPr.GridBefore       = this.GridBefore;
	RowPr.Jc               = this.Jc;
	RowPr.TableCellSpacing = this.TableCellSpacing;

	if (undefined != this.Height)
		RowPr.Height = this.Height.Copy();

	if (undefined != this.WAfter)
		RowPr.WAfter = this.WAfter.Copy();

	if (undefined != this.WBefore)
		RowPr.WBefore = this.WBefore.Copy();

	RowPr.TableHeader = this.TableHeader;

	if (true === bCopyPrChange && undefined !== this.PrChange)
	{
		RowPr.PrChange   = this.PrChange.Copy();
		RowPr.ReviewInfo = this.ReviewInfo.Copy();
	}

	return RowPr;
};
CTableRowPr.prototype.Merge = function(RowPr)
{
	if (undefined !== RowPr.CantSplit)
		this.CantSplit = RowPr.CantSplit;

	if (undefined !== RowPr.GridAfter)
		this.GridAfter = RowPr.GridAfter;

	if (undefined !== RowPr.GridBefore)
		this.GridBefore = RowPr.GridBefore;

	if (undefined !== RowPr.Jc)
		this.Jc = RowPr.Jc;

	if (undefined !== RowPr.TableCellSpacing)
		this.TableCellSpacing = RowPr.TableCellSpacing;

	if (undefined !== RowPr.Height)
		this.Height = RowPr.Height.Copy();

	if (undefined !== RowPr.WAfter)
		this.WAfter = RowPr.WAfter.Copy();

	if (undefined !== RowPr.WBefore)
		this.WBefore = RowPr.WBefore.Copy();

	if (undefined !== RowPr.TableHeader)
		this.TableHeader = RowPr.TableHeader;
};
CTableRowPr.prototype.Is_Equal = function(RowPr)
{
	if (this.CantSplit !== RowPr.CantSplit
		|| this.GridAfter !== RowPr.GridAfter
		|| this.GridBefore !== RowPr.GridBefore
		|| this.Jc !== RowPr.Jc
		|| this.TableCellSpacing !== RowPr.TableCellSpacing
		|| true !== IsEqualStyleObjects(this.Height, RowPr.Height)
		|| true !== IsEqualStyleObjects(this.WAfter, RowPr.WAfter)
		|| true !== IsEqualStyleObjects(this.WBefore, RowPr.WBefore)
		|| this.TableHeader !== RowPr.TableHeader)
		return false;

	return true;
};
CTableRowPr.prototype.InitDefault = function(nCompatibilityMode)
{
	this.CantSplit        = false;
	this.GridAfter        = 0;
	this.GridBefore       = 0;
	this.Jc               = align_Left;
	this.TableCellSpacing = null;
	this.Height           = new CTableRowHeight(0, Asc.linerule_Auto);
	this.WAfter           = new CTableMeasurement(tblwidth_Auto, 0);
	this.WBefore          = new CTableMeasurement(tblwidth_Auto, 0);
	this.TableHeader      = false;
	this.PrChange         = undefined;
	this.ReviewInfo       = undefined;
};
CTableRowPr.prototype.Set_FromObject = function(RowPr)
{
	this.CantSplit        = RowPr.CantSplit;
	this.GridAfter        = RowPr.GridAfter;
	this.GridBefore       = RowPr.GridBefore;
	this.Jc               = RowPr.Jc;
	this.TableCellSpacing = RowPr.TableCellSpacing;

	if (undefined != RowPr.Height)
		this.Height = new CTableRowHeight(RowPr.Height.Value, RowPr.Height.HRule);
	else
		this.Height = undefined;

	if (undefined != RowPr.WAfter)
		this.WAfter = new CTableMeasurement(RowPr.WAfter.Type, RowPr.WAfter.W);
	else
		this.WAfter = undefined;

	if (undefined != RowPr.WBefore)
		this.WBefore = new CTableMeasurement(RowPr.WBefore.Type, RowPr.WBefore.W);
	else
		this.WBefore = undefined;

	this.TableHeader = RowPr.TableHeader;
};
CTableRowPr.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);
	var Flags = 0;

	if (undefined != this.CantSplit)
	{
		Writer.WriteBool(this.CantSplit);
		Flags |= 1;
	}

	if (undefined != this.GridAfter)
	{
		Writer.WriteLong(this.GridAfter);
		Flags |= 2;
	}

	if (undefined != this.GridBefore)
	{
		Writer.WriteLong(this.GridBefore);
		Flags |= 4;
	}

	if (undefined != this.Jc)
	{
		Writer.WriteLong(this.Jc);
		Flags |= 8;
	}

	if (undefined != this.TableCellSpacing)
	{
		if (null === this.TableCellSpacing)
			Writer.WriteBool(true);
		else
		{
			Writer.WriteBool(false);
			Writer.WriteDouble(this.TableCellSpacing);
		}

		Flags |= 16;
	}

	if (undefined != this.Height)
	{
		this.Height.Write_ToBinary(Writer);
		Flags |= 32;
	}

	if (undefined != this.WAfter)
	{
		this.WAfter.Write_ToBinary(Writer);
		Flags |= 64;
	}

	if (undefined != this.WBefore)
	{
		this.WBefore.Write_ToBinary(Writer);
		Flags |= 128;
	}

	if (undefined != this.TableHeader)
	{
		Writer.WriteBool(this.TableHeader);
		Flags |= 256;
	}

	if (undefined !== this.PrChange && undefined !== this.ReviewInfo)
	{
		this.PrChange.WriteToBinary(Writer);
		this.ReviewInfo.WriteToBinary(Writer);
		Flags |= 512;
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CTableRowPr.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	if (1 & Flags)
		this.CantSplit = Reader.GetBool();

	if (2 & Flags)
		this.GridAfter = Reader.GetLong();

	if (4 & Flags)
		this.GridBefore = Reader.GetLong();

	if (8 & Flags)
		this.Jc = Reader.GetLong();

	if (16 & Flags)
	{
		if (true === Reader.GetBool())
			this.TableCellSpacing = Reader.GetLong();
		else
			this.TableCellSpacing = Reader.GetDouble();
	}

	if (32 & Flags)
	{
		this.Height = new CTableRowHeight(0, Asc.linerule_Auto);
		this.Height.Read_FromBinary(Reader);
	}

	if (64 & Flags)
	{
		this.WAfter = new CTableMeasurement(tblwidth_Auto, 0);
		this.WAfter.Read_FromBinary(Reader);
	}

	if (128 & Flags)
	{
		this.WBefore = new CTableMeasurement(tblwidth_Auto, 0);
		this.WBefore.Read_FromBinary(Reader);
	}

	if (256 & Flags)
		this.TableHeader = Reader.GetBool();

	if (512 & Flags)
	{
		this.PrChange   = new CTableRowPr();
		this.ReviewInfo = new AscWord.ReviewInfo();

		this.PrChange.ReadFromBinary(Reader);
		this.ReviewInfo.ReadFromBinary(Reader);
	}
};
CTableRowPr.prototype.WriteToBinary = function(oWriter)
{
	this.Write_ToBinary(oWriter);
};
CTableRowPr.prototype.ReadFromBinary = function(oReader)
{
	this.Read_FromBinary(oReader);
};
CTableRowPr.prototype.HavePrChange = function()
{
	if (undefined === this.PrChange || null === this.PrChange)
		return false;

	return true;
};
CTableRowPr.prototype.AddPrChange = function()
{
	this.PrChange   = this.Copy(false);
	this.ReviewInfo = new AscWord.ReviewInfo();
	this.ReviewInfo.Update();
};
CTableRowPr.prototype.SetPrChange = function(oPrChange, oReviewInfo)
{
	this.PrChange   = oPrChange;
	this.ReviewInfo = oReviewInfo;
};
CTableRowPr.prototype.RemovePrChange = function()
{
	delete this.PrChange;
	delete this.ReviewInfo;
};

function CTableCellPr()
{
    this.GridSpan         = undefined;
    this.Shd              = undefined;
    this.TableCellMar     = undefined; // undefined/null/{Top, Left, Right, Bottom}
    this.TableCellBorders =
    {
        Bottom : undefined,
        Left   : undefined,
        Right  : undefined,
        Top    : undefined
    };
    this.TableCellW       = undefined;
    this.VAlign           = undefined;
    this.VMerge           = undefined;
    this.TextDirection    = undefined;
    this.NoWrap           = undefined;
    this.HMerge           = undefined;
    this.PrChange         = undefined;
    this.ReviewInfo       = undefined;
}

CTableCellPr.prototype.Copy = function(bCopyPrChange)
{
	var CellPr = new CTableCellPr();

	CellPr.GridSpan = this.GridSpan;

	if (undefined != this.Shd)
		CellPr.Shd = this.Shd.Copy();

	if (undefined === this.TableCellMar)
		CellPr.TableCellMar = undefined;
	else if (null === this.TableCellMar)
		CellPr.TableCellMar = null;
	else
	{
		CellPr.TableCellMar        = {};
		CellPr.TableCellMar.Bottom = undefined != this.TableCellMar.Bottom ? this.TableCellMar.Bottom.Copy() : undefined;
		CellPr.TableCellMar.Left   = undefined != this.TableCellMar.Left ? this.TableCellMar.Left.Copy() : undefined;
		CellPr.TableCellMar.Right  = undefined != this.TableCellMar.Right ? this.TableCellMar.Right.Copy() : undefined;
		CellPr.TableCellMar.Top    = undefined != this.TableCellMar.Top ? this.TableCellMar.Top.Copy() : undefined;
	}

	if (undefined != this.TableCellBorders.Bottom)
		CellPr.TableCellBorders.Bottom = this.TableCellBorders.Bottom.Copy();

	if (undefined != this.TableCellBorders.Left)
		CellPr.TableCellBorders.Left = this.TableCellBorders.Left.Copy();

	if (undefined != this.TableCellBorders.Right)
		CellPr.TableCellBorders.Right = this.TableCellBorders.Right.Copy();

	if (undefined != this.TableCellBorders.Top)
		CellPr.TableCellBorders.Top = this.TableCellBorders.Top.Copy();

	if (undefined != this.TableCellW)
		CellPr.TableCellW = this.TableCellW.Copy();

	CellPr.VAlign        = this.VAlign;
	CellPr.VMerge        = this.VMerge;
	CellPr.TextDirection = this.TextDirection;
	CellPr.NoWrap        = this.NoWrap;
	CellPr.HMerge        = this.HMerge;

	if (true === bCopyPrChange && undefined !== this.PrChange)
	{
		CellPr.PrChange   = this.PrChange.Copy();
		CellPr.ReviewInfo = this.ReviewInfo.Copy();
	}

	return CellPr;
};
CTableCellPr.prototype.Merge = function(CellPr)
{
	if (undefined != CellPr.GridSpan)
		this.GridSpan = CellPr.GridSpan;

	if (undefined != CellPr.Shd)
		this.Shd = CellPr.Shd.Copy();

	if (undefined === CellPr.TableCellMar)
	{
	}
	else if (null === CellPr.TableCellMar)
		this.TableCellMar = null;
	else
	{
		this.TableCellMar        = {};
		this.TableCellMar.Bottom = undefined != CellPr.TableCellMar.Bottom ? CellPr.TableCellMar.Bottom.Copy() : undefined;
		this.TableCellMar.Left   = undefined != CellPr.TableCellMar.Left ? CellPr.TableCellMar.Left.Copy() : undefined;
		this.TableCellMar.Right  = undefined != CellPr.TableCellMar.Right ? CellPr.TableCellMar.Right.Copy() : undefined;
		this.TableCellMar.Top    = undefined != CellPr.TableCellMar.Top ? CellPr.TableCellMar.Top.Copy() : undefined;
	}

	if (undefined != CellPr.TableCellBorders.Bottom)
		this.TableCellBorders.Bottom = CellPr.TableCellBorders.Bottom.Copy();

	if (undefined != CellPr.TableCellBorders.Left)
		this.TableCellBorders.Left = CellPr.TableCellBorders.Left.Copy();

	if (undefined != CellPr.TableCellBorders.Right)
		this.TableCellBorders.Right = CellPr.TableCellBorders.Right.Copy();

	if (undefined != CellPr.TableCellBorders.Top)
		this.TableCellBorders.Top = CellPr.TableCellBorders.Top.Copy();

	if (undefined != CellPr.TableCellW)
		this.TableCellW = CellPr.TableCellW.Copy();

	if (undefined != CellPr.VAlign)
		this.VAlign = CellPr.VAlign;

	if (undefined != CellPr.VMerge)
		this.VMerge = CellPr.VMerge;

	if (undefined != CellPr.TextDirection)
		this.TextDirection = CellPr.TextDirection;

	if (undefined !== CellPr.NoWrap)
		this.NoWrap = CellPr.NoWrap;

	if (undefined != CellPr.HMerge)
		this.HMerge = CellPr.HMerge;
};
CTableCellPr.prototype.Is_Equal = function(CellPr)
{
	if (this.GridSpan !== CellPr.GridSpan
		|| true !== IsEqualStyleObjects(this.Shd, CellPr.Shd)
		|| (this.TableCellMar !== undefined
		&& CellPr.TableCellMar === undefined)
		|| (CellPr.TableCellMar !== undefined
		&& this.TableCellMar === undefined)
		|| (this.TableCellMar !== null
		&& CellPr.TableCellMar === null)
		|| (CellPr.TableCellMar !== null
		&& this.TableCellMar === null)
		|| (this.TableCellMar !== undefined
		&& this.TableCellMar !== null
		&& (true !== IsEqualStyleObjects(this.TableCellMar.Top, CellPr.TableCellMar.Top)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Left, CellPr.TableCellMar.Left)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Right, CellPr.TableCellMar.Right)
		|| true !== IsEqualStyleObjects(this.TableCellMar.Bottom, CellPr.TableCellMar.Bottom)))
		|| true !== IsEqualStyleObjects(this.TableCellBorders.Bottom, CellPr.TableCellBorders.Bottom)
		|| true !== IsEqualStyleObjects(this.TableCellBorders.Left, CellPr.TableCellBorders.Left)
		|| true !== IsEqualStyleObjects(this.TableCellBorders.Right, CellPr.TableCellBorders.Right)
		|| true !== IsEqualStyleObjects(this.TableCellBorders.Top, CellPr.TableCellBorders.Top)
		|| true !== IsEqualStyleObjects(this.TableCellW, CellPr.TableCellW)
		|| this.VAlign !== CellPr.VAlign
		|| this.VMerge !== CellPr.VMerge
		|| this.TextDirection !== CellPr.TextDirection
		|| this.NoWrap !== CellPr.NoWrap
		|| this.HMerge !== CellPr.HMerge)
		return false;

	return true;
};
CTableCellPr.prototype.InitDefault = function(nCompatibilityMode)
{
	this.GridSpan                = 1;
	this.Shd                     = new CDocumentShd();
	this.TableCellMar            = null;
	this.TableCellBorders.Bottom = undefined;
	this.TableCellBorders.Left   = undefined;
	this.TableCellBorders.Right  = undefined;
	this.TableCellBorders.Top    = undefined;
	this.TableCellW              = new CTableMeasurement(tblwidth_Auto, 0);
	this.VAlign                  = vertalignjc_Top;
	this.VMerge                  = vmerge_Restart;
	this.TextDirection           = textdirection_LRTB;
	this.NoWrap                  = false;
	this.HMerge                  = vmerge_Restart;
	this.PrChange                = undefined;
	this.ReviewInfo              = undefined;
};
CTableCellPr.prototype.Set_FromObject = function(CellPr)
{
	this.GridSpan = CellPr.GridSpan;

	if (undefined != CellPr.Shd)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Set_FromObject(CellPr.Shd);
	}
	else
		this.Shd = undefined;

	if (undefined === CellPr.TableCellMar)
		this.TableCellMar = undefined;
	else if (null === CellPr.TableCellMar)
		this.TableCellMar = null;
	else
	{
		this.TableCellMar = {};

		if (undefined != CellPr.TableCellMar.Bottom)
			this.TableCellMar.Bottom = new CTableMeasurement(CellPr.TableCellMar.Bottom.Type, CellPr.TableCellMar.Bottom.W);
		else
			this.TableCellMar.Bottom = undefined;

		if (undefined != CellPr.TableCellMar.Left)
			this.TableCellMar.Left = new CTableMeasurement(CellPr.TableCellMar.Left.Type, CellPr.TableCellMar.Left.W);
		else
			this.TableCellMar.Left = undefined;

		if (undefined != CellPr.TableCellMar.Right)
			this.TableCellMar.Right = new CTableMeasurement(CellPr.TableCellMar.Right.Type, CellPr.TableCellMar.Right.W);
		else
			this.TableCellMar.Right = undefined;

		if (undefined != CellPr.TableCellMar.Top)
			this.TableCellMar.Top = new CTableMeasurement(CellPr.TableCellMar.Top.Type, CellPr.TableCellMar.Top.W);
		else
			this.TableCellMar.Top = undefined;
	}

	if (undefined != CellPr.TableCellBorders)
	{
		if (undefined != CellPr.TableCellBorders.Bottom)
		{
			this.TableCellBorders.Bottom = new CDocumentBorder();
			this.TableCellBorders.Bottom.Set_FromObject(CellPr.TableCellBorders.Bottom);
		}
		else
			this.TableCellBorders.Bottom = undefined;

		if (undefined != CellPr.TableCellBorders.Left)
		{
			this.TableCellBorders.Left = new CDocumentBorder();
			this.TableCellBorders.Left.Set_FromObject(CellPr.TableCellBorders.Left);
		}
		else
			this.TableCellBorders.Left = undefined;

		if (undefined != CellPr.TableCellBorders.Right)
		{
			this.TableCellBorders.Right = new CDocumentBorder();
			this.TableCellBorders.Right.Set_FromObject(CellPr.TableCellBorders.Right);
		}
		else
			this.TableCellBorders.Right = undefined;

		if (undefined != CellPr.TableCellBorders.Top)
		{
			this.TableCellBorders.Top = new CDocumentBorder();
			this.TableCellBorders.Top.Set_FromObject(CellPr.TableCellBorders.Top);
		}
		else
			this.TableCellBorders.Top = undefined;
	}
	else
	{
		this.TableCellBorders.Bottom = undefined;
		this.TableCellBorders.Left   = undefined;
		this.TableCellBorders.Right  = undefined;
		this.TableCellBorders.Top    = undefined;
	}

	if (undefined != CellPr.TableCellW)
		this.TableCellW = new CTableMeasurement(CellPr.TableCellW.Type, CellPr.TableCellW.W);
	else
		this.TableCellW = undefined;


	this.VAlign        = CellPr.VAlign;
	this.VMerge        = CellPr.VMerge;
	this.TextDirection = CellPr.TextDirection;
	this.NoWrap        = CellPr.NoWrap;
	this.HMerge        = CellPr.HMerge;

};
CTableCellPr.prototype.Check_PresentationPr = function(Theme)
{
	if (this.Shd)
	{
		this.Shd.Check_PresentationPr(Theme);
	}
	if (this.TableCellBorders.Bottom)
	{
		this.TableCellBorders.Bottom.Check_PresentationPr(Theme);
	}
	if (this.TableCellBorders.Left)
	{
		this.TableCellBorders.Left.Check_PresentationPr(Theme);
	}
	if (this.TableCellBorders.Right)
	{
		this.TableCellBorders.Right.Check_PresentationPr(Theme);
	}
	if (this.TableCellBorders.Top)
	{
		this.TableCellBorders.Top.Check_PresentationPr(Theme);
	}
};
CTableCellPr.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);
	var Flags = 0;

	if (undefined != this.GridSpan)
	{
		Writer.WriteLong(this.GridSpan);
		Flags |= 1;
	}

	if (undefined != this.Shd)
	{
		this.Shd.Write_ToBinary(Writer);
		Flags |= 2;
	}

	if (undefined != this.TableCellMar)
	{
		if (null === this.TableCellMar)
		{
			Flags |= 4;
		}
		else
		{
			if (undefined != this.TableCellMar.Bottom)
			{
				this.TableCellMar.Bottom.Write_ToBinary(Writer);
				Flags |= 8;
			}

			if (undefined != this.TableCellMar.Left)
			{
				this.TableCellMar.Left.Write_ToBinary(Writer);
				Flags |= 16;
			}

			if (undefined != this.TableCellMar.Right)
			{
				this.TableCellMar.Right.Write_ToBinary(Writer);
				Flags |= 32;
			}

			if (undefined != this.TableCellMar.Top)
			{
				this.TableCellMar.Top.Write_ToBinary(Writer);
				Flags |= 64;
			}

			Flags |= 128;
		}
	}

	if (undefined != this.TableCellBorders.Bottom)
	{
		this.TableCellBorders.Bottom.Write_ToBinary(Writer);
		Flags |= 256;
	}

	if (undefined != this.TableCellBorders.Left)
	{
		this.TableCellBorders.Left.Write_ToBinary(Writer);
		Flags |= 512;
	}

	if (undefined != this.TableCellBorders.Right)
	{
		this.TableCellBorders.Right.Write_ToBinary(Writer);
		Flags |= 1024;
	}

	if (undefined != this.TableCellBorders.Top)
	{
		this.TableCellBorders.Top.Write_ToBinary(Writer);
		Flags |= 2048;
	}

	if (undefined != this.TableCellW)
	{
		this.TableCellW.Write_ToBinary(Writer);
		Flags |= 4096;
	}

	if (undefined != this.VAlign)
	{
		Writer.WriteLong(this.VAlign);
		Flags |= 8192;
	}

	if (undefined != this.VMerge)
	{
		Writer.WriteLong(this.VMerge);
		Flags |= 16384;
	}

	if (undefined !== this.TextDirection)
	{
		Writer.WriteLong(this.TextDirection);
		Flags |= 32768;
	}

	if (undefined !== this.NoWrap)
	{
		Writer.WriteBool(this.NoWrap);
		Flags |= 65536;
	}

	if (undefined !== this.HMerge)
	{
		Writer.WriteLong(this.HMerge);
		Flags |= 131072;
	}

	if (undefined !== this.PrChange && undefined !== this.ReviewInfo)
	{
		this.PrChange.WriteToBinary(Writer);
		this.ReviewInfo.WriteToBinary(Writer);
		Flags |= 262144;
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CTableCellPr.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	if (1 & Flags)
		this.GridSpan = Reader.GetLong();

	if (2 & Flags)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Read_FromBinary(Reader);
	}

	if (4 & Flags)
		this.TableCellMar = null;
	else if (128 & Flags)
	{
		this.TableCellMar = {};
		if (8 & Flags)
		{
			this.TableCellMar.Bottom = new CTableMeasurement(tblwidth_Auto, 0);
			this.TableCellMar.Bottom.Read_FromBinary(Reader);
		}

		if (16 & Flags)
		{
			this.TableCellMar.Left = new CTableMeasurement(tblwidth_Auto, 0);
			this.TableCellMar.Left.Read_FromBinary(Reader);
		}

		if (32 & Flags)
		{
			this.TableCellMar.Right = new CTableMeasurement(tblwidth_Auto, 0);
			this.TableCellMar.Right.Read_FromBinary(Reader);
		}

		if (64 & Flags)
		{
			this.TableCellMar.Top = new CTableMeasurement(tblwidth_Auto, 0);
			this.TableCellMar.Top.Read_FromBinary(Reader);
		}
	}

	if (256 & Flags)
	{
		this.TableCellBorders.Bottom = new CDocumentBorder();
		this.TableCellBorders.Bottom.Read_FromBinary(Reader);
	}

	if (512 & Flags)
	{
		this.TableCellBorders.Left = new CDocumentBorder();
		this.TableCellBorders.Left.Read_FromBinary(Reader);
	}

	if (1024 & Flags)
	{
		this.TableCellBorders.Right = new CDocumentBorder();
		this.TableCellBorders.Right.Read_FromBinary(Reader);
	}

	if (2048 & Flags)
	{
		this.TableCellBorders.Top = new CDocumentBorder();
		this.TableCellBorders.Top.Read_FromBinary(Reader);
	}

	if (4096 & Flags)
	{
		this.TableCellW = new CTableMeasurement(tblwidth_Auto, 0);
		this.TableCellW.Read_FromBinary(Reader);
	}

	if (8192 & Flags)
		this.VAlign = Reader.GetLong();

	if (16384 & Flags)
		this.VMerge = Reader.GetLong();

	if (32768 & Flags)
		this.TextDirection = Reader.GetLong();

	if (65536 & Flags)
		this.NoWrap = Reader.GetBool();

	if (131072 & Flags)
		this.HMerge = Reader.GetLong();

	if (262144 & Flags)
	{
		this.PrChange   = new CTableCellPr();
		this.ReviewInfo = new AscWord.ReviewInfo();

		this.PrChange.ReadFromBinary(Reader);
		this.ReviewInfo.ReadFromBinary(Reader);
	}
};
CTableCellPr.prototype.WriteToBinary = function(oWriter)
{
	this.Write_ToBinary(oWriter);
};
CTableCellPr.prototype.ReadFromBinary = function(oReader)
{
	this.Read_FromBinary(oReader);
};
CTableCellPr.prototype.Is_Empty = function()
{
	if (undefined !== this.GridSpan
		|| undefined !== this.Shd
		|| undefined !== this.TableCellMar
		|| undefined !== this.TableCellBorders.Bottom
		|| undefined !== this.TableCellBorders.Left
		|| undefined !== this.TableCellBorders.Right
		|| undefined !== this.TableCellBorders.Top
		|| undefined !== this.TableCellW
		|| undefined !== this.VAlign
		|| undefined !== this.VMerge
		|| undefined !== this.TextDirection
		|| undefined !== this.NoWrap
		|| undefined !== this.HMerge)
		return false;

	return true;
};
CTableCellPr.prototype.HavePrChange = function()
{
	if (undefined === this.PrChange || null === this.PrChange)
		return false;

	return true;
};
CTableCellPr.prototype.AddPrChange = function()
{
	this.PrChange   = this.Copy(false);
	this.ReviewInfo = new AscWord.ReviewInfo();
	this.ReviewInfo.Update();
};
CTableCellPr.prototype.SetPrChange = function(oPrChange, oReviewInfo)
{
	this.PrChange   = oPrChange;
	this.ReviewInfo = oReviewInfo;
};
CTableCellPr.prototype.RemovePrChange = function()
{
	delete this.PrChange;
	delete this.ReviewInfo;
};

function CRFonts()
{
    this.Ascii    = undefined;
    this.EastAsia = undefined;
    this.HAnsi    = undefined;
    this.CS       = undefined;
    this.Hint     = undefined;

    this.AsciiTheme    = undefined;
    this.EastAsiaTheme = undefined;
    this.HAnsiTheme    = undefined;
    this.CSTheme       = undefined;
}

CRFonts.prototype.Is_Empty = function()
{
	return (undefined === this.Ascii
		&& undefined === this.EastAsia
		&& undefined === this.HAnsi
		&& undefined === this.CS
		&& undefined === this.Hint
		&& undefined === this.AsciiTheme
		&& undefined === this.EastAsiaTheme
		&& undefined === this.HAnsiTheme
		&& undefined === this.CSTheme);
};
CRFonts.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CRFonts.prototype.Copy = function()
{
	var oRFonts = new CRFonts();

	if (undefined !== this.Ascii)
	{
		oRFonts.Ascii = {
			Name  : this.Ascii.Name,
			Index : this.Ascii.Index
		};
	}

	if (undefined !== this.EastAsia)
	{
		oRFonts.EastAsia = {
			Name  : this.EastAsia.Name,
			Index : this.EastAsia.Index
		};
	}

	if (undefined !== this.HAnsi)
	{
		oRFonts.HAnsi = {
			Name  : this.HAnsi.Name,
			Index : this.HAnsi.Index
		};
	}

	if (undefined !== this.CS)
	{
		oRFonts.CS = {
			Name  : this.CS.Name,
			Index : this.CS.Index
		};
	}

	oRFonts.Hint          = this.Hint;
	oRFonts.AsciiTheme    = this.AsciiTheme;
	oRFonts.EastAsiaTheme = this.EastAsiaTheme;
	oRFonts.HAnsiTheme    = this.HAnsiTheme;
	oRFonts.CSTheme       = this.CSTheme;

	return oRFonts;
};
CRFonts.prototype.Merge = function(oRFonts)
{
	if (oRFonts.AsciiTheme)
	{
		this.AsciiTheme = oRFonts.AsciiTheme;
		this.Ascii      = undefined;
	}
	else if (oRFonts.Ascii && oRFonts.Ascii.Name)
	{
		this.Ascii      = oRFonts.Ascii;
		this.AsciiTheme = undefined;
	}

	if (oRFonts.EastAsiaTheme)
	{
		this.EastAsiaTheme = oRFonts.EastAsiaTheme;
		this.EastAsia      = undefined;
	}
	else if (oRFonts.EastAsia && oRFonts.EastAsia.Name)
	{
		this.EastAsia      = oRFonts.EastAsia;
		this.EastAsiaTheme = undefined;
	}

	if (oRFonts.HAnsiTheme)
	{
		this.HAnsiTheme = oRFonts.HAnsiTheme;
		this.HAnsi      = undefined;
	}
	else if (oRFonts.HAnsi && oRFonts.HAnsi.Name)
	{
		this.HAnsi      = oRFonts.HAnsi;
		this.HAnsiTheme = undefined;
	}

	if (oRFonts.CSTheme)
	{
		this.CSTheme = oRFonts.CSTheme;
		this.CS      = undefined;
	}
	else if (oRFonts.CS && oRFonts.CS.Name)
	{
		this.CS      = oRFonts.CS;
		this.CSTheme = undefined;
	}

	if (oRFonts.Hint)
		this.Hint = oRFonts.Hint;
};
CRFonts.prototype.InitDefault = function()
{
	this.Ascii = {
		Name  : "Times New Roman",
		Index : -1
	};

	this.EastAsia = {
		Name  : "Times New Roman",
		Index : -1
	};

	this.HAnsi = {
		Name  : "Times New Roman",
		Index : -1
	};

	this.CS = {
		Name  : "Times New Roman",
		Index : -1
	};

	this.Hint          = AscWord.fonthint_Default;
	this.AsciiTheme    = undefined;
	this.EastAsiaTheme = undefined;
	this.HAnsiTheme    = undefined;
	this.CSTheme       = undefined;
};
CRFonts.prototype.Set_FromObject = function(oRFonts, isUndefinedToNull)
{
	if (undefined !== oRFonts.Ascii)
	{
		this.Ascii = {
			Name  : oRFonts.Ascii.Name,
			Index : oRFonts.Ascii.Index
		};
	}
	else
	{
		this.Ascii = isUndefinedToNull ? null : undefined;
	}

	if (undefined !== oRFonts.EastAsia)
	{
		this.EastAsia = {
			Name  : oRFonts.EastAsia.Name,
			Index : oRFonts.EastAsia.Index
		};
	}
	else
	{
		this.EastAsia = isUndefinedToNull ? null : undefined;
	}

	if (undefined !== oRFonts.HAnsi)
	{
		this.HAnsi = {
			Name  : oRFonts.HAnsi.Name,
			Index : oRFonts.HAnsi.Index
		};
	}
	else
	{
		this.HAnsi = isUndefinedToNull ? null : undefined;
	}

	if (undefined !== oRFonts.CS)
	{
		this.CS = {
			Name  : oRFonts.CS.Name,
			Index : oRFonts.CS.Index
		};
	}
	else
	{
		this.CS = isUndefinedToNull ? null : undefined;
	}

	this.Hint          = CheckUndefinedToNull(isUndefinedToNull, oRFonts.Hint);
	this.AsciiTheme    = CheckUndefinedToNull(isUndefinedToNull, oRFonts.AsciiTheme);
	this.EastAsiaTheme = CheckUndefinedToNull(isUndefinedToNull, oRFonts.EastAsiaTheme);
	this.HAnsiTheme    = CheckUndefinedToNull(isUndefinedToNull, oRFonts.HAnsiTheme);
	this.CSTheme       = CheckUndefinedToNull(isUndefinedToNull, oRFonts.CSTheme);
};
CRFonts.prototype.SetAll = function(sFontName, nFontIndex)
{
	if (undefined === nFontIndex)
		nFontIndex = -1;

	this.Ascii = {
		Name  : sFontName,
		Index : nFontIndex
	};

	this.EastAsia = {
		Name  : sFontName,
		Index : nFontIndex
	};

	this.HAnsi = {
		Name  : sFontName,
		Index : nFontIndex
	};

	this.CS = {
		Name  : sFontName,
		Index : nFontIndex
	};

	this.Hint = AscWord.fonthint_Default;

	this.AsciiTheme    = undefined;
	this.EastAsiaTheme = undefined;
	this.HAnsiTheme    = undefined;
	this.CSTheme       = undefined;
};
CRFonts.prototype.SetFontStyle = function(nFontStyleIdx)
{
	var sFirstPart;
	if (nFontStyleIdx === AscFormat.fntStyleInd_major)
	{
		sFirstPart = "+mj-";
	}
	else
	{
		sFirstPart = "+mn-";
	}
	var oRFonts = {};
	oRFonts.Ascii = {Name: sFirstPart + "lt", Index: -1};
	oRFonts.EastAsia = {Name: sFirstPart + "ea", Index: -1};
	oRFonts.CS = {Name: sFirstPart + "cs", Index: -1};
	oRFonts.HAnsi = {Name: sFirstPart + "lt", Index: -1};
	this.Set_FromObject(oRFonts, false);
};
CRFonts.prototype.IsEqual = function(oRFonts)
{
	return (this.private_IsEqual(this.Ascii, oRFonts.Ascii)
		&& this.private_IsEqual(this.EastAsia, oRFonts.EastAsia)
		&& this.private_IsEqual(this.HAnsi, oRFonts.HAnsi)
		&& this.private_IsEqual(this.CS, oRFonts.CS)
		&& this.Hint === oRFonts.Hint
		&& this.AsciiTheme === oRFonts.AsciiTheme
		&& this.EastAsiaTheme === oRFonts.EastAsiaTheme
		&& this.HAnsiTheme === oRFonts.HAnsiTheme
		&& this.CSTheme === oRFonts.CSTheme);
};
CRFonts.prototype.IsEqualSlot = function(rFonts, fontSlot)
{
	switch (fontSlot)
	{
		case AscWord.fontslot_ASCII:
			return (this.private_IsEqual(this.Ascii, rFonts.Ascii) && this.AsciiTheme === rFonts.AsciiTheme);
		case AscWord.fontslot_EastAsia:
			return (this.private_IsEqual(this.EastAsia, rFonts.EastAsia) && this.EastAsiaTheme === rFonts.EastAsiaTheme);
		case AscWord.fontslot_HAnsi:
			return (this.private_IsEqual(this.HAnsi, rFonts.HAnsi) && this.HAnsiTheme === rFonts.HAnsiTheme);
		case AscWord.fontslot_CS:
			return (this.private_IsEqual(this.CS, rFonts.CS) && this.CSTheme === rFonts.CSTheme);
	}
	
	return this.IsEqual(rFonts);
};
CRFonts.prototype.private_IsEqual = function(oFont1, oFont2)
{
	return ((undefined === oFont1 && undefined === oFont2)
		|| (undefined !== oFont1 && undefined !== oFont2 && oFont1.Name === oFont2.Name));
};
CRFonts.prototype.Is_Equal = function(oRFonts)
{
	return this.IsEqual(oRFonts);
};
CRFonts.prototype.Compare = function(rFonts)
{
	if (!this.private_IsEqual(this.Ascii, rFonts.Ascii))
		this.Ascii = {Name : undefined, Index : -1};
	
	if (!this.private_IsEqual(this.EastAsia, rFonts.EastAsia))
		this.EastAsia = {Name : undefined, Index : -1};
	
	if (!this.private_IsEqual(this.HAnsi, rFonts.HAnsi))
		this.HAnsi = {Name : undefined, Index : -1};
	
	if (!this.private_IsEqual(this.CS, rFonts.CS))
		this.CS = {Name : undefined, Index : -1};
	
	if (this.AsciiTheme !== rFonts.AsciiTheme)
		this.AsciiTheme = undefined;
	
	if (this.EastAsiaTheme !== rFonts.EastAsiaTheme)
		this.EastAsiaTheme = undefined;
	
	if (this.HAnsiTheme !== rFonts.HAnsiTheme)
		this.HAnsiTheme = undefined;
	
	if (this.CSTheme !== rFonts.CSTheme)
		this.CSTheme = undefined;
};
CRFonts.prototype.Write_ToBinary = function(oWriter)
{
	var nStartPos = oWriter.GetCurPosition();
	oWriter.Skip(4);
	var nFlags = 0;

	if (undefined !== this.Ascii && null !== this.Ascii)
	{
		oWriter.WriteString2(this.Ascii.Name);
		nFlags |= 1;
	}

	if (undefined !== this.EastAsia && null !== this.EastAsia)
	{
		oWriter.WriteString2(this.EastAsia.Name);
		nFlags |= 2;
	}

	if (undefined !== this.HAnsi && null !== this.HAnsi)
	{
		oWriter.WriteString2(this.HAnsi.Name);
		nFlags |= 4;
	}

	if (undefined !== this.CS && null !== this.CS)
	{
		oWriter.WriteString2(this.CS.Name);
		nFlags |= 8;
	}

	if (undefined !== this.Hint && null !== this.Hint)
	{
		oWriter.WriteLong(this.Hint);
		nFlags |= 16;
	}

	if (undefined !== this.AsciiTheme && null !== this.AsciiTheme)
	{
		oWriter.WriteString2(this.AsciiTheme);
		nFlags |= 32;
	}

	if (undefined !== this.EastAsiaTheme && null !== this.EastAsiaTheme)
	{
		oWriter.WriteString2(this.EastAsiaTheme);
		nFlags |= 64;
	}

	if (undefined !== this.HAnsiTheme && null !== this.HAnsiTheme)
	{
		oWriter.WriteString2(this.HAnsiTheme);
		nFlags |= 128;
	}

	if (undefined !== this.CSTheme && null !== this.CSTheme)
	{
		oWriter.WriteString2(this.CSTheme);
		nFlags |= 256;
	}

	var nEndPos = oWriter.GetCurPosition();
	oWriter.Seek(nStartPos);
	oWriter.WriteLong(nFlags);
	oWriter.Seek(nEndPos);
};
CRFonts.prototype.Read_FromBinary = function(oReader)
{
	var nFlags = oReader.GetLong();

	if (nFlags & 1)
		this.Ascii = {Name : oReader.GetString2(), Index : -1};

	if (nFlags & 2)
		this.EastAsia = {Name : oReader.GetString2(), Index : -1};

	if (nFlags & 4)
		this.HAnsi = {Name : oReader.GetString2(), Index : -1};

	if (nFlags & 8)
		this.CS = {Name : oReader.GetString2(), Index : -1};

	if (nFlags & 16)
		this.Hint = oReader.GetLong();

	if (nFlags & 32)
		this.AsciiTheme = oReader.GetString2();

	if (nFlags & 64)
		this.EastAsiaTheme = oReader.GetString2();

	if (nFlags & 128)
		this.HAnsiTheme = oReader.GetString2();

	if (nFlags & 256)
		this.CSTheme = oReader.GetString2();
};
CRFonts.prototype.HaveThemeFonts = function()
{
	return (undefined !== this.AsciiTheme
		|| undefined !== this.EastAsiaTheme
		|| undefined !== this.HAnsiTheme
		|| undefined !== this.CSTheme);
};

function CLang()
{
    this.Bidi     = undefined;
    this.EastAsia = undefined;
    this.Val      = undefined;
}

CLang.prototype =
{
    Copy : function()
    {
        var Lang = new CLang();
        Lang.Bidi     = this.Bidi;
        Lang.EastAsia = this.EastAsia;
        Lang.Val      = this.Val;
        return Lang;
    },

    Merge : function(Lang)
    {
        if ( undefined !== Lang.Bidi )
            this.Bidi = Lang.Bidi;

        if ( undefined !== Lang.EastAsia )
            this.EastAsia = Lang.EastAsia;

        if ( undefined !== Lang.Val )
            this.Val = Lang.Val;
    },

	InitDefault : function()
    {
        this.Bidi     = lcid_arSA;
        this.EastAsia = lcid_zhCN;
        this.Val      = lcid_enUS;
    },

    Set_FromObject : function(Lang, isUndefinedToNull)
	{
		this.Bidi     = CheckUndefinedToNull(isUndefinedToNull, Lang.Bidi);
		this.EastAsia = CheckUndefinedToNull(isUndefinedToNull, Lang.EastAsia);
		this.Val      = CheckUndefinedToNull(isUndefinedToNull, Lang.Val);
	},

    Compare : function(Lang)
    {
        // Bidi
        if ( undefined !== this.Bidi && this.Bidi !== Lang.Bidi )
            this.Bidi = undefined;

        // EastAsia
        if ( undefined !== this.EastAsia && this.EastAsia !== Lang.EastAsia )
            this.EastAsia = undefined;

        // Val
        if ( undefined !== this.Val && this.Val !== Lang.Val )
            this.Val = undefined;
    },

    Write_ToBinary : function(Writer)
    {
        var StartPos = Writer.GetCurPosition();
        Writer.Skip(4);
        var Flags = 0;

        if ( undefined != this.Bidi )
        {
            Writer.WriteLong( this.Bidi );
            Flags |= 1;
        }

        if ( undefined != this.EastAsia )
        {
            Writer.WriteLong( this.EastAsia );
            Flags |= 2;
        }

        if ( undefined != this.Val )
        {
            Writer.WriteLong( this.Val );
            Flags |= 4;
        }

        var EndPos = Writer.GetCurPosition();
        Writer.Seek( StartPos );
        Writer.WriteLong( Flags );
        Writer.Seek( EndPos );
    },

    Read_FromBinary : function(Reader)
    {
        var Flags = Reader.GetLong();

        // Bidi
        if ( Flags & 1 )
            this.Bidi = Reader.GetLong();

        // EastAsia
        if ( Flags & 2 )
            this.EastAsia = Reader.GetLong();

        // Val
        if ( Flags & 4 )
            this.Val = Reader.GetLong();
    }
};
CLang.prototype.Clear = function()
{
	this.Bidi     = undefined;
	this.EastAsia = undefined;
	this.Val      = undefined;
};
CLang.prototype.Apply = function(lang)
{
	if (null === lang.Bidi)
		this.Bidi = undefined;
	else if (undefined !== lang.Bidi)
		this.Bidi = lang.Bidi;
	
	if (null === lang.EastAsia)
		this.EastAsia = undefined;
	else if (undefined !== lang.EastAsia)
		this.EastAsia = lang.EastAsia;
	
	if (null === lang.Val)
		this.Val = undefined;
	else if (undefined !== lang.Val)
		this.Val = lang.Val;
};
CLang.prototype.Is_Empty = function()
{
	if (undefined !== this.Bidi
		|| undefined !== this.EastAsia
		|| undefined !== this.Val)
		return false;

	return true;
};
CLang.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CLang.prototype.IsEqual = function(oLang)
{
	return (this.Bidi === oLang.Bidi && this.EastAsia === oLang.EastAsia && this.Val === oLang.Val);
};
CLang.prototype.Is_Equal = function(Lang)
{
	return this.IsEqual(Lang);
}

function CTextPr()
{
    this.Bold       = undefined; // Жирный текст
    this.Italic     = undefined; // Наклонный текст
    this.Strikeout  = undefined; // Зачеркивание
    this.Underline  = undefined;
    this.FontFamily = undefined;
    this.FontSize   = undefined;
    this.Color      = undefined;
    this.VertAlign  = undefined;
    this.HighLight  = undefined; // highlight_None/Color
    this.RStyle     = undefined;
    this.Spacing    = undefined; // Дополнительное расстояние между символвами
    this.DStrikeout = undefined; // Двойное зачеркивание
    this.Caps       = undefined;
    this.SmallCaps  = undefined;
    this.Position   = undefined; // Смещение по Y

    this.RFonts     = new CRFonts();
    this.BoldCS     = undefined;
    this.ItalicCS   = undefined;
    this.FontSizeCS = undefined;
    this.CS         = undefined;
    this.RTL        = undefined;
    this.Lang       = new CLang();
    this.Unifill    = undefined;
    this.FontRef    = undefined;

    this.Shd        = undefined;
    this.Vanish     = undefined;
	this.Ligatures  = undefined;

    this.TextOutline    = undefined;
    this.TextFill       = undefined;
	this.HighlightColor = undefined;
	this.FontScale      = undefined;
	this.FontSizeOrig   = undefined;
	this.FontSizeCSOrig = undefined;

	this.PrChange       = undefined;
	this.ReviewInfo     = undefined;
}

CTextPr.prototype.Clear = function()
{
	this.Bold       = undefined;
	this.Italic     = undefined;
	this.Strikeout  = undefined;
	this.Underline  = undefined;
	this.FontFamily = undefined;
	this.FontSize   = undefined;
	this.Color      = undefined;
	this.VertAlign  = undefined;
	this.HighLight  = undefined;
	this.RStyle     = undefined;
	this.Spacing    = undefined;
	this.DStrikeout = undefined;
	this.Caps       = undefined;
	this.SmallCaps  = undefined;
	this.Position   = undefined;

	this.RFonts         = new CRFonts();
	this.BoldCS         = undefined;
	this.ItalicCS       = undefined;
	this.FontSizeCS     = undefined;
	this.CS             = undefined;
	this.RTL            = undefined;
	this.Lang           = new CLang();
	this.Unifill        = undefined;
	this.FontRef        = undefined;
	this.Shd            = undefined;
	this.Vanish         = undefined;
	this.Ligatures      = undefined;
	this.TextOutline    = undefined;
	this.TextFill       = undefined;
	this.HighlightColor = undefined;
	this.AscFill        = undefined;
	this.AscUnifill     = undefined;
	this.AscLine        = undefined;
	this.FontScale      = undefined;
	this.FontSizeOrig   = undefined;
	this.FontSizeCSOrig = undefined;

	this.PrChange   = undefined;
	this.ReviewInfo = undefined;
};
CTextPr.prototype.Copy = function(bCopyPrChange, oPr)
{
	var TextPr       = new CTextPr();
	TextPr.Bold      = this.Bold;
	TextPr.Italic    = this.Italic;
	TextPr.Strikeout = this.Strikeout;
	TextPr.Underline = this.Underline;

	if (undefined != this.FontFamily)
	{
		TextPr.FontFamily       = {};
		TextPr.FontFamily.Name  = this.FontFamily.Name;
		TextPr.FontFamily.Index = this.FontFamily.Index;
	}

	TextPr.FontSize = this.FontSize;

	if (undefined != this.Color)
		TextPr.Color = new CDocumentColor(this.Color.r, this.Color.g, this.Color.b, this.Color.Auto);

	TextPr.VertAlign = this.VertAlign;
	TextPr.HighLight = this.Copy_HighLight();

	TextPr.RStyle     = this.RStyle;
	if(oPr && oPr.Comparison && TextPr.RStyle)
	{
		TextPr.RStyle = oPr.Comparison.copyStyleById(TextPr.RStyle);
	}
	TextPr.Spacing    = this.Spacing;
	TextPr.DStrikeout = this.DStrikeout;
	TextPr.Caps       = this.Caps;
	TextPr.SmallCaps  = this.SmallCaps;
	TextPr.Position   = this.Position;
	TextPr.RFonts     = this.RFonts.Copy();
	TextPr.BoldCS     = this.BoldCS;
	TextPr.ItalicCS   = this.ItalicCS;
	TextPr.FontSizeCS = this.FontSizeCS;
	TextPr.CS         = this.CS;
	TextPr.RTL        = this.RTL;
	TextPr.Lang       = this.Lang.Copy();
	if (undefined != this.Unifill)
		TextPr.Unifill = this.Unifill.createDuplicate();
	if (undefined != this.FontRef)
		TextPr.FontRef = this.FontRef.createDuplicate();

	if (undefined !== this.Shd)
		TextPr.Shd = this.Shd.Copy();
	if (undefined !== this.FontScale)
		TextPr.FontScale = this.FontScale;

	TextPr.Vanish    = this.Vanish;
	TextPr.Ligatures = this.Ligatures;

	if (true === bCopyPrChange && undefined !== this.PrChange)
	{
		TextPr.PrChange   = this.PrChange.Copy();
		TextPr.ReviewInfo = this.ReviewInfo.Copy();
	}
	if (undefined != this.TextOutline)
	{
		TextPr.TextOutline = this.TextOutline.createDuplicate();
	}
	if (undefined != this.TextFill)
	{
		TextPr.TextFill = this.TextFill.createDuplicate();
	}
	if (undefined !== this.HighlightColor)
	{
		TextPr.HighlightColor = this.HighlightColor.createDuplicate();
	}

	return TextPr;
};
CTextPr.prototype.createDuplicateForSmartArt = function(oPr)
{
	var TextPr       = new CTextPr();

	TextPr.Bold      = this.Bold;
	TextPr.Italic    = this.Italic;
	TextPr.Strikeout = this.Strikeout;
	TextPr.DStrikeout = this.DStrikeout;
	TextPr.Caps       = this.Caps;
	TextPr.SmallCaps  = this.SmallCaps;

	TextPr.Underline = this.Underline;
	TextPr.Lang       = this.Lang.Copy();
	TextPr.Spacing    = this.Spacing;
	TextPr.RFonts     = this.RFonts.Copy()

	TextPr.Vanish    = this.Vanish;
	TextPr.Ligatures = this.Ligatures;

	if (oPr.custT) {
		TextPr.FontSize = this.FontSize;
	}

	if (undefined != this.FontFamily)
	{
		TextPr.FontFamily       = {};
		TextPr.FontFamily.Name  = this.FontFamily.Name;
		TextPr.FontFamily.Index = this.FontFamily.Index;
	}

	if (undefined != this.Color)
		TextPr.Color = new CDocumentColor(this.Color.r, this.Color.g, this.Color.b, this.Color.Auto);

	TextPr.VertAlign = this.VertAlign;
	TextPr.HighLight = this.Copy_HighLight();

	if (undefined != this.Unifill)
		TextPr.Unifill = this.Unifill.createDuplicate();

	if (undefined != this.FontRef)
		TextPr.FontRef = this.FontRef.createDuplicate();

	if (undefined !== this.Shd)
		TextPr.Shd = this.Shd.Copy();

	if (undefined !== this.FontScale)
		TextPr.FontScale = this.FontScale;

	if (undefined != this.TextOutline)
	{
		TextPr.TextOutline = this.TextOutline.createDuplicate();
	}
	if (undefined != this.TextFill)
	{
		TextPr.TextFill = this.TextFill.createDuplicate();
	}
	if (undefined !== this.HighlightColor)
	{
		TextPr.HighlightColor = this.HighlightColor.createDuplicate();
	}
	return TextPr;
};
CTextPr.prototype.Copy_HighLight = function()
{
	if (undefined === this.HighLight)
		return undefined;
	else if (highlight_None === this.HighLight)
		return highlight_None;
	else
		return this.HighLight.Copy();

	return undefined;
};
CTextPr.prototype.Merge = function(TextPr)
{
	if (undefined != TextPr.Bold)
		this.Bold = TextPr.Bold;

	if (undefined != TextPr.Italic)
		this.Italic = TextPr.Italic;

	if (undefined != TextPr.Strikeout)
		this.Strikeout = TextPr.Strikeout;

	if (undefined != TextPr.Underline)
		this.Underline = TextPr.Underline;

	if (undefined != TextPr.FontFamily && undefined !== TextPr.FontFamily.Name)
	{
		this.FontFamily       = {};
		this.FontFamily.Name  = TextPr.FontFamily.Name;
		this.FontFamily.Index = TextPr.FontFamily.Index;
	}

	if (undefined != TextPr.FontSize)
		this.FontSize = TextPr.FontSize;

	if (undefined != TextPr.Color)
		this.Color = TextPr.Color.Copy();

	if (undefined != TextPr.VertAlign)
		this.VertAlign = TextPr.VertAlign;

	if (undefined === TextPr.HighLight || null === TextPr.HighLight)
	{
	}
	else if (highlight_None === TextPr.HighLight)
		this.HighLight = highlight_None;
	else
		this.HighLight = TextPr.HighLight.Copy();

	if (undefined != TextPr.RStyle)
		this.RStyle = TextPr.RStyle;

	if (undefined != TextPr.Spacing)
		this.Spacing = TextPr.Spacing;

	if (undefined != TextPr.DStrikeout)
		this.DStrikeout = TextPr.DStrikeout;

	if (undefined != TextPr.SmallCaps)
		this.SmallCaps = TextPr.SmallCaps;

	if (undefined != TextPr.Caps)
		this.Caps = TextPr.Caps;

	if (undefined != TextPr.Position)
		this.Position = TextPr.Position;

	this.RFonts.Merge(TextPr.RFonts);

	if (undefined != TextPr.BoldCS)
		this.BoldCS = TextPr.BoldCS;

	if (undefined != TextPr.ItalicCS)
		this.ItalicCS = TextPr.ItalicCS;

	if (undefined != TextPr.FontSizeCS)
		this.FontSizeCS = TextPr.FontSizeCS;

	if (undefined != TextPr.CS)
		this.CS = TextPr.CS;

	if (undefined != TextPr.RTL)
		this.RTL = TextPr.RTL;

	if (TextPr.Lang)
		this.Lang.Merge(TextPr.Lang);

	if (TextPr.Unifill)
	{
		this.Unifill = TextPr.Unifill.createDuplicate();
		this.TextFill = undefined;
	}
	else if (undefined != TextPr.Color)
		this.Unifill = undefined;

	if (undefined != TextPr.FontRef)
	{
		this.FontRef = TextPr.FontRef.createDuplicate();
	}

	if (TextPr.Shd)
		this.Shd = TextPr.Shd.Copy();

	if (undefined !== TextPr.Vanish && null !== TextPr.Vanish)
		this.Vanish = TextPr.Vanish;

	if (undefined !== TextPr.Ligatures && null !== TextPr.Ligatures)
		this.Ligatures = TextPr.Ligatures;

	if (TextPr.TextOutline)
		this.TextOutline = TextPr.TextOutline.createDuplicate();

	if (TextPr.TextFill)
	{
		this.TextFill = TextPr.TextFill.createDuplicate();
		this.Unifill = undefined;
	}

	if (TextPr.HighlightColor)
		this.HighlightColor = TextPr.HighlightColor.createDuplicate();
	if (undefined !== TextPr.FontScale)
	{
		this.FontScale = TextPr.FontScale;
	}
};
/**
 * Накатываем на данные настройки те, которые пришли, причем если какое-то значение null, то это поле мы делаем undefined
 * @param {CTextPr} textPr
 */
CTextPr.prototype.Apply = function(textPr)
{
	if (null === textPr.Bold)
		this.Bold = undefined;
	else if (undefined !== textPr.Bold)
		this.Bold = textPr.Bold;
	
	if (null === textPr.BoldCS)
		this.BoldCS = undefined;
	else if (undefined !== textPr.BoldCS)
		this.BoldCS = textPr.BoldCS;
	
	if (null === textPr.Italic)
		this.Italic = undefined;
	else if (undefined !== textPr.Italic)
		this.Italic = textPr.Italic;
	
	if (null === textPr.ItalicCS)
		this.ItalicCS = undefined;
	else if (undefined !== textPr.ItalicCS)
		this.ItalicCS = textPr.ItalicCS;
	
	if (null === textPr.Strikeout)
		this.Strikeout = undefined;
	else if (undefined !== textPr.Strikeout)
		this.Strikeout = textPr.Strikeout;
	
	if (null === textPr.Underline)
		this.Underline = undefined;
	else if (undefined !== textPr.Underline)
		this.Underline = textPr.Underline;

	if (null === textPr.FontFamily)
	{
		this.FontFamily = undefined;
	}
	else if (undefined !== textPr.FontFamily)
	{
		this.FontFamily = {
			Name  : textPr.FontFamily.Name,
			Index : textPr.FontFamily.Index,
		};
	}
	
	if (null === textPr.FontSize)
		this.FontSize = undefined;
	else if (undefined !== textPr.FontSize)
		this.FontSize = textPr.FontSize;
	
	if (null === textPr.FontSizeCS)
		this.FontSizeCS = undefined;
	else if (undefined !== textPr.FontSizeCS)
		this.FontSizeCS = textPr.FontSizeCS;
	
	if (null === textPr.Color)
		this.Color = undefined;
	else if (undefined !== textPr.Color)
		this.Color = textPr.Color.Copy();
	
	if (null === textPr.VertAlign)
		this.VertAlign = undefined;
	else if (undefined !== textPr.VertAlign)
		this.VertAlign = textPr.VertAlign;
	
	if (null === textPr.HighLight)
		this.HighLight = undefined;
	else if (highlight_None === textPr.HighLight)
		this.HighLight = highlight_None;
	else if (undefined !== textPr.HighLight)
		this.HighLight = textPr.HighLight.Copy();
	
	if (null === textPr.RStyle)
		this.RStyle = undefined;
	else if (undefined !== textPr.RStyle)
		this.RStyle = textPr.RStyle;
	
	if (null === textPr.Spacing)
		this.Spacing = undefined;
	else if (undefined !== textPr.Spacing)
		this.Spacing = textPr.Spacing;
	
	if (null === textPr.DStrikeout)
		this.DStrikeout = undefined;
	else if (undefined !== textPr.DStrikeout)
		this.DStrikeout = textPr.DStrikeout;
	
	if (null === textPr.SmallCaps)
		this.SmallCaps = undefined;
	else if (undefined !== textPr.SmallCaps)
		this.SmallCaps = textPr.SmallCaps;
	
	if (null === textPr.Caps)
		this.Caps = undefined;
	else if (undefined !== textPr.Caps)
		this.Caps = textPr.Caps;
	
	if (null === textPr.Position)
		this.Position = undefined;
	else if (undefined !== textPr.Position)
		this.Position = textPr.Position;
	
	if (textPr.RFonts)
		this.RFonts.Merge(textPr.RFonts);
	
	if (null === textPr.CS)
		this.CS = undefined;
	else if (undefined !== textPr.CS)
		this.CS = textPr.CS;
	
	if (null === textPr.RTL)
		this.RTL = undefined;
	else if (undefined !== textPr.RTL)
		this.RTL = textPr.RTL;
	
	if (null === textPr.Lang)
		this.Lang.Clear();
	else if (textPr.Lang)
		this.Lang.Apply(textPr.Lang);
	
	if (null === textPr.Unifill)
		this.Unifill = undefined;
	else if (textPr.Unifill)
		this.Unifill = textPr.Unifill.createDuplicate();
	
	if (null === textPr.FontRef)
		this.FontRef = undefined;
	else if (textPr.FontRef)
		this.FontRef = textPr.FontRef.createDuplicate();
	
	if (null === textPr.TextOutline)
		this.TextOutline = undefined;
	else if (textPr.TextOutline)
		this.TextOutline = textPr.TextOutline.createDuplicate();
	
	if (null === textPr.TextFill)
		this.TextFill = undefined;
	else if (textPr.TextFill)
		this.TextFill = textPr.TextFill.createDuplicate();
	
	if (null === textPr.HighlightColor)
		this.HighlightColor = undefined;
	else if (textPr.HighlightColor)
		this.HighlightColor = textPr.HighlightColor.createDuplicate();
	
	if (null === textPr.Shd)
		this.Shd = undefined;
	else if (undefined !== textPr.Shd)
		this.Shd = textPr.Shd.Copy();
	
	if (null === textPr.Vanish)
		this.Vanish = undefined;
	else if (undefined !== textPr.Vanish)
		this.Vanish = textPr.Vanish;
	
	if (null === textPr.Ligatures)
		this.Ligatures = undefined;
	else if (undefined !== textPr.Ligatures)
		this.Ligatures = textPr.Ligatures;
	
	if (null === textPr.FontScale)
		this.FontScale = undefined;
	else if (undefined !== textPr.FontScale)
		this.FontScale = textPr.FontScale;
};
CTextPr.prototype.InitDefault = function(nCompatibilityMode)
{
	this.Bold       = false;
	this.Italic     = false;
	this.Underline  = false;
	this.Strikeout  = false;
	this.FontFamily = {
		Name  : "Times New Roman",
		Index : -1
	};
	this.FontSize   = 10;
	this.Color      = new CDocumentColor(0, 0, 0, true);
	this.VertAlign  = AscCommon.vertalign_Baseline;
	this.HighLight  = highlight_None;
	this.RStyle     = undefined;
	this.Spacing    = 0;
	this.DStrikeout = false;
	this.SmallCaps  = false;
	this.Caps       = false;
	this.Position   = 0;
	this.RFonts.InitDefault();
	this.BoldCS     = false;
	this.ItalicCS   = false;
	this.FontSizeCS = 10;
	this.CS         = false;
	this.RTL        = false;
	this.Lang.InitDefault();
	this.Unifill = undefined;
	this.FontRef = undefined;
	this.Shd     = undefined;
	this.Vanish  = false;
	this.Ligatures  = Asc.LigaturesType.None;

	this.TextOutline    = undefined;
	this.TextFill       = undefined;
	this.HighlightColor = undefined;
	this.FontScale      = undefined;
	this.FontSizeOrig   = undefined;
	this.FontSizeCSOrig = undefined;

	this.PrChange   = undefined;
	this.ReviewInfo = undefined;
};
CTextPr.prototype.Set_FromObject = function(TextPr, isUndefinedToNull)
{
	this.Bold      = CheckUndefinedToNull(isUndefinedToNull, TextPr.Bold);
	this.Italic    = CheckUndefinedToNull(isUndefinedToNull, TextPr.Italic);
	this.Strikeout = CheckUndefinedToNull(isUndefinedToNull, TextPr.Strikeout);
	this.Underline = CheckUndefinedToNull(isUndefinedToNull, TextPr.Underline);

	if (undefined !== TextPr.FontFamily)
	{
		this.FontFamily       = {};
		this.FontFamily.Name  = TextPr.FontFamily.Name;
		this.FontFamily.Index = TextPr.FontFamily.Index;
	}
	else
	{
		if (isUndefinedToNull)
			this.FontFamily = null;
		else
			this.FontFamily = undefined;
	}

	this.FontSize = CheckUndefinedToNull(isUndefinedToNull, TextPr.FontSize);

	if (null === TextPr.Color || undefined === TextPr.Color)
		this.Color = CheckUndefinedToNull(isUndefinedToNull, TextPr.Color);
	else
		this.Color = new CDocumentColor(TextPr.Color.r, TextPr.Color.g, TextPr.Color.b, TextPr.Color.Auto);

	this.VertAlign = CheckUndefinedToNull(isUndefinedToNull, TextPr.VertAlign);

	if (undefined === TextPr.HighLight || null === TextPr.HighLight)
		this.HighLight = CheckUndefinedToNull(isUndefinedToNull, undefined);
	else if (highlight_None === TextPr.HighLight)
		this.HighLight = highlight_None;
	else
		this.HighLight = new CDocumentColor(TextPr.HighLight.r, TextPr.HighLight.g, TextPr.HighLight.b);

	this.RStyle     = CheckUndefinedToNull(isUndefinedToNull, TextPr.RStyle);
	this.Spacing    = CheckUndefinedToNull(isUndefinedToNull, TextPr.Spacing);
	this.DStrikeout = CheckUndefinedToNull(isUndefinedToNull, TextPr.DStrikeout);
	this.Caps       = CheckUndefinedToNull(isUndefinedToNull, TextPr.Caps);
	this.SmallCaps  = CheckUndefinedToNull(isUndefinedToNull, TextPr.SmallCaps);
	this.Position   = CheckUndefinedToNull(isUndefinedToNull, TextPr.Position);

	if (undefined !== TextPr.RFonts)
		this.RFonts.Set_FromObject(TextPr.RFonts, isUndefinedToNull);

	this.BoldCS     = CheckUndefinedToNull(isUndefinedToNull, TextPr.BoldCS);
	this.ItalicCS   = CheckUndefinedToNull(isUndefinedToNull, TextPr.ItalicCS);
	this.FontSizeCS = CheckUndefinedToNull(isUndefinedToNull, TextPr.FontSizeCS);
	this.CS         = CheckUndefinedToNull(isUndefinedToNull, TextPr.CS);
	this.RTL        = CheckUndefinedToNull(isUndefinedToNull, TextPr.RTL);

	if (undefined !== TextPr.Lang)
		this.Lang.Set_FromObject(TextPr.Lang, isUndefinedToNull);

	if (undefined !== TextPr.Shd)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Set_FromObject(TextPr.Shd);
	}
	else
	{
		this.Shd = isUndefinedToNull ? null : undefined;
	}

	this.Vanish         = CheckUndefinedToNull(isUndefinedToNull, TextPr.Vanish);
	this.Ligatures      = CheckUndefinedToNull(isUndefinedToNull, TextPr.Ligatures);
	this.Unifill        = CheckUndefinedToNull(isUndefinedToNull, TextPr.Unifill);
	this.FontRef        = CheckUndefinedToNull(isUndefinedToNull, TextPr.FontRef);
	this.TextFill       = CheckUndefinedToNull(isUndefinedToNull, TextPr.TextFill);
	this.HighlightColor = CheckUndefinedToNull(isUndefinedToNull, TextPr.HighlightColor);
	this.TextOutline    = CheckUndefinedToNull(isUndefinedToNull, TextPr.TextOutline);
	this.AscFill        = CheckUndefinedToNull(isUndefinedToNull, TextPr.AscFill);
	this.AscUnifill     = CheckUndefinedToNull(isUndefinedToNull, TextPr.AscUnifill);
	this.AscLine        = CheckUndefinedToNull(isUndefinedToNull, TextPr.AscLine);
	this.FontScale      = CheckUndefinedToNull(isUndefinedToNull, TextPr.FontScale);
};
CTextPr.prototype.Check_PresentationPr = function()
{
	if (this.FontRef && !this.Unifill)
	{
		this.RFonts.SetFontStyle(this.FontRef.idx);
		if (this.FontRef.Color && !this.Unifill)
		{
			this.Unifill = AscFormat.CreateUniFillByUniColorCopy(this.FontRef.Color);
		}
	}
};
CTextPr.prototype.Compare = function(TextPr)
{
	// Bold
	if (undefined !== this.Bold && this.Bold !== TextPr.Bold)
		this.Bold = undefined;

	// Italic
	if (undefined !== this.Italic && this.Italic !== TextPr.Italic)
		this.Italic = undefined;

	// Strikeout
	if (undefined !== this.Strikeout && this.Strikeout !== TextPr.Strikeout)
		this.Strikeout = undefined;

	// Underline
	if (undefined !== this.Underline && this.Underline !== TextPr.Underline)
		this.Underline = undefined;

	// FontFamily
	if (undefined !== this.FontFamily && ( undefined === TextPr.FontFamily || this.FontFamily.Name !== TextPr.FontFamily.Name ))
		this.FontFamily = undefined;

	// FontSize
	if (undefined !== this.FontSize && ( undefined === TextPr.FontSize || Math.abs(this.FontSize - TextPr.FontSize) >= 0.001 ))
		this.FontSize = undefined;

	// Color
	if (undefined !== this.Color && ( undefined === TextPr.Color || true !== this.Color.Compare(TextPr.Color) ))
		this.Color = undefined;

	// VertAlign
	if (undefined !== this.VertAlign && this.VertAlign !== TextPr.VertAlign)
		this.VertAlign = undefined;

	// HighLight
	if (undefined !== this.HighLight && ( undefined === TextPr.HighLight || ( highlight_None === this.HighLight && highlight_None !== TextPr.HighLight ) || ( highlight_None !== this.HighLight && highlight_None === TextPr.HighLight ) || ( highlight_None !== this.HighLight && highlight_None !== TextPr.HighLight && true !== this.HighLight.Compare(TextPr.HighLight) ) ))
		this.HighLight = undefined;

	// RStyle
	if (undefined !== this.RStyle && ( undefined === TextPr.RStyle || this.RStyle !== TextPr.RStyle ))
		this.RStyle = undefined;

	// Spacing
	if (undefined !== this.Spacing && ( undefined === TextPr.Spacing || Math.abs(this.Spacing - TextPr.Spacing) >= 0.001 ))
		this.Spacing = undefined;

	// DStrikeout
	if (undefined !== this.DStrikeout && ( undefined === TextPr.DStrikeout || this.DStrikeout !== TextPr.DStrikeout ))
		this.DStrikeout = undefined;

	// Caps
	if (undefined !== this.Caps && ( undefined === TextPr.Caps || this.Caps !== TextPr.Caps ))
		this.Caps = undefined;

	// SmallCaps
	if (undefined !== this.SmallCaps && ( undefined === TextPr.SmallCaps || this.SmallCaps !== TextPr.SmallCaps ))
		this.SmallCaps = undefined;

	// Position
	if (undefined !== this.Position && ( undefined === TextPr.Position || Math.abs(this.Position - TextPr.Position) >= 0.001 ))
		this.Position = undefined;

	// RFonts
	this.RFonts.Compare(TextPr.RFonts);

	// BoldCS
	if (undefined !== this.BoldCS && this.BoldCS !== TextPr.BoldCS)
		this.BoldCS = undefined;

	// ItalicCS
	if (undefined !== this.ItalicCS && this.ItalicCS !== TextPr.ItalicCS)
		this.ItalicCS = undefined;

	// FontSizeCS
	if (undefined !== this.FontSizeCS && ( undefined === TextPr.FontSizeCS || Math.abs(this.FontSizeCS - TextPr.FontSizeCS) >= 0.001 ))
		this.FontSizeCS = undefined;

	// CS
	if (undefined !== this.CS && this.CS !== TextPr.CS)
		this.CS = undefined;

	// RTL
	if (undefined !== this.RTL && this.RTL !== TextPr.RTL)
		this.RTL = undefined;

	// Lang
	this.Lang.Compare(TextPr.Lang);
	//Result_TextPr.Unifill = CompareUniFill(this.Unifill, TextPr.Unifill);

	// Vanish
	if (undefined !== this.Vanish && this.Vanish !== TextPr.Vanish)
		this.Vanish = undefined;

	if (undefined !== this.Ligatures && this.Ligatures !== TextPr.Ligatures)
		this.Ligatures = undefined;

	if (undefined !== this.FontScale && this.FontScale !== TextPr.FontScale)
		this.FontScale = undefined;

	if (undefined !== this.Unifill && !this.Unifill.IsIdentical(TextPr.Unifill))
	{
		this.Unifill = AscFormat.CompareUniFill(this.Unifill, TextPr.Unifill);
		if (null === this.Unifill)
		{
			this.Unifill = undefined;
		}
		this.Color    = undefined;
		this.TextFill = undefined;
	}


	if (undefined !== this.TextFill && !this.TextFill.IsIdentical(TextPr.TextFill))
	{
		this.Unifill  = undefined;
		this.Color    = undefined;
		this.TextFill = AscFormat.CompareUniFill(this.TextFill, TextPr.TextFill);
		if (null === this.TextFill)
		{
			this.TextFill = undefined;
		}
	}

	if (undefined !== this.HighlightColor && !this.HighlightColor.IsIdentical(TextPr.HighlightColor))
	{
		this.HighlightColor = this.HighlightColor.compare(TextPr.HighlightColor);
		if (null === this.HighlightColor)
		{
			this.HighlightColor = undefined;
		}
	}


	if (undefined !== this.TextOutline && !this.TextOutline.IsIdentical(TextPr.TextOutline))
	{
		if (TextPr.TextOutline !== undefined)
		{
			this.TextOutline = this.TextOutline.compare(TextPr.TextOutline);
		}
		else
		{
			this.TextOutline = undefined;
		}

	}

	return this;
};
CTextPr.prototype.ReplaceThemeFonts = function(oFontScheme)
{
	if (this.RFonts && oFontScheme)
	{
		if (this.RFonts.AsciiTheme)
		{
			this.RFonts.Ascii = {
				Name  : oFontScheme.checkFont(this.RFonts.AsciiTheme),
				Index : -1
			};

			this.RFonts.AsciiTheme = undefined;
		}
		else if (this.RFonts.Ascii)
		{
			this.RFonts.Ascii.Name  = oFontScheme.checkFont(this.RFonts.Ascii.Name);
			this.RFonts.Ascii.Index = -1;
		}

		if (this.RFonts.EastAsiaTheme)
		{
			this.RFonts.EastAsia = {
				Name  : oFontScheme.checkFont(this.RFonts.EastAsiaTheme),
				Index : -1
			};

			this.RFonts.EastAsiaTheme = undefined;
		}
		else if (this.RFonts.EastAsia)
		{
			this.RFonts.EastAsia.Name  = oFontScheme.checkFont(this.RFonts.EastAsia.Name);
			this.RFonts.EastAsia.Index = -1;
		}

		if (this.RFonts.HAnsiTheme)
		{
			this.RFonts.HAnsi = {
				Name  : oFontScheme.checkFont(this.RFonts.HAnsiTheme),
				Index : -1
			};

			this.RFonts.HAnsiTheme = undefined;
		}
		else if (this.RFonts.HAnsi)
		{
			this.RFonts.HAnsi.Name  = oFontScheme.checkFont(this.RFonts.HAnsi.Name);
			this.RFonts.HAnsi.Index = -1;
		}

		if (this.RFonts.CSTheme)
		{
			this.RFonts.CS = {
				Name  : oFontScheme.checkFont(this.RFonts.CSTheme),
				Index : -1
			};

			this.RFonts.CSTheme = undefined;
		}
		else if (this.RFonts.CS)
		{
			this.RFonts.CS.Name  = oFontScheme.checkFont(this.RFonts.CS.Name);
			this.RFonts.CS.Index = -1;
		}
	}

	if (this.FontFamily && oFontScheme)
	{
		this.FontFamily.Name  = oFontScheme.checkFont(this.FontFamily.Name);
		this.FontFamily.Index = -1;
	}
};
CTextPr.prototype.GetSimpleTextColor = function(oTheme, oColorMap)
{
	if (this.Unifill)
	{
		if (oTheme && oColorMap)
			this.Unifill.check(oTheme, oColorMap);

		var oRGBA = this.Unifill.getRGBAColor();
		return new CDocumentColor(oRGBA.R, oRGBA.G, oRGBA.B);
	}
	else if (this.Color)
	{
		return this.Color;
	}

	return new CDocumentColor();
}
CTextPr.prototype.GetIncDecFontSize = function(IncFontSize)
{
	var FontSize = this.FontSize;
	if(this.FontScale !== null &&
		this.FontScale !== undefined &&
		this.FontSizeOrig !== null &&
		this.FontSizeOrig !== undefined)
	{
		FontSize = this.FontSizeOrig;
	}
	return FontSize_IncreaseDecreaseValue(IncFontSize, FontSize);
};
CTextPr.prototype.GetIncDecFontSizeCS = function(IncFontSize)
{
	var FontSize = this.FontSizeCS;
	if(this.FontScale !== null &&
		this.FontScale !== undefined &&
		this.FontSizeCSOrig !== null &&
		this.FontSizeCSOrig !== undefined)
	{
		FontSize = this.FontSizeCSOrig;
	}
	return FontSize_IncreaseDecreaseValue(IncFontSize, FontSize);
};
CTextPr.prototype.CheckFontScale = function()
{
	if (this.FontScale !== null && this.FontScale !== undefined)
	{
		this.FontSizeOrig   = this.FontSize;
		this.FontSizeCSOrig = this.FontSizeCS;
		this.FontSize      *= this.FontScale;
		this.FontSize       = (this.FontSize + 0.5) >> 0;
		this.FontSize       = Math.max(1, this.FontSize);
		this.FontSizeCS    *= this.FontScale;
		this.FontSizeCS     = (this.FontSizeCS + 0.5) >> 0;
		this.FontSizeCS     = Math.max(1, this.FontSizeCS);
	}
};
CTextPr.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);
	var Flags = 0;

	if (undefined != this.Bold)
	{
		Writer.WriteBool(this.Bold);
		Flags |= 1;
	}

	if (undefined != this.Italic)
	{
		Writer.WriteBool(this.Italic);
		Flags |= 2;
	}

	if (undefined != this.Underline)
	{
		Writer.WriteBool(this.Underline);
		Flags |= 4;
	}

	if (undefined != this.Strikeout)
	{
		Writer.WriteBool(this.Strikeout);
		Flags |= 8;
	}

	if (undefined != this.FontFamily)
	{
		Writer.WriteString2(this.FontFamily.Name);
		Flags |= 16;
	}

	if (undefined != this.FontSize)
	{
		Writer.WriteDouble(this.FontSize);
		Flags |= 32;
	}

	if (undefined != this.Color)
	{
		this.Color.Write_ToBinary(Writer);
		Flags |= 64;
	}

	if (undefined != this.VertAlign)
	{
		Writer.WriteLong(this.VertAlign);
		Flags |= 128;
	}

	if (undefined != this.HighLight)
	{
		if (highlight_None === this.HighLight)
		{
			Writer.WriteLong(highlight_None);
		}
		else
		{
			Writer.WriteLong(0);
			this.HighLight.Write_ToBinary(Writer);
		}

		Flags |= 256;
	}

	if (undefined != this.RStyle)
	{
		Writer.WriteString2(this.RStyle);
		Flags |= 512;
	}

	if (undefined != this.Spacing)
	{
		Writer.WriteDouble(this.Spacing);
		Flags |= 1024;
	}

	if (undefined != this.DStrikeout)
	{
		Writer.WriteBool(this.DStrikeout);
		Flags |= 2048;
	}

	if (undefined != this.Caps)
	{
		Writer.WriteBool(this.Caps);
		Flags |= 4096;
	}

	if (undefined != this.SmallCaps)
	{
		Writer.WriteBool(this.SmallCaps);
		Flags |= 8192;
	}

	if (undefined != this.Position)
	{
		Writer.WriteDouble(this.Position);
		Flags |= 16384;
	}

	if (undefined != this.RFonts)
	{
		this.RFonts.Write_ToBinary(Writer);
		Flags |= 32768;
	}

	if (undefined != this.BoldCS)
	{
		Writer.WriteBool(this.BoldCS);
		Flags |= 65536;
	}

	if (undefined != this.ItalicCS)
	{
		Writer.WriteBool(this.ItalicCS);
		Flags |= 131072;
	}

	if (undefined != this.FontSizeCS)
	{
		Writer.WriteDouble(this.FontSizeCS);
		Flags |= 262144;
	}

	if (undefined != this.CS)
	{
		Writer.WriteBool(this.CS);
		Flags |= 524288;
	}

	if (undefined != this.RTL)
	{
		Writer.WriteBool(this.RTL);
		Flags |= 1048576;
	}

	if (undefined != this.Lang)
	{
		this.Lang.Write_ToBinary(Writer);
		Flags |= 2097152;
	}

	if (undefined != this.Unifill)
	{
		this.Unifill.Write_ToBinary(Writer);
		Flags |= 4194304;
	}

	if (undefined != this.Shd)
	{
		this.Shd.Write_ToBinary(Writer);
		Flags |= 8388608;
	}

	if (undefined != this.Vanish)
	{
		Writer.WriteBool(this.Vanish);
		Flags |= 16777216;
	}

	if (undefined != this.FontRef)
	{
		this.FontRef.Write_ToBinary(Writer);
		Flags |= 33554432;
	}

	if (undefined != this.PrChange)
	{
		this.PrChange.Write_ToBinary(Writer);
		Flags |= 67108864;
	}
	if (undefined != this.TextOutline)
	{
		this.TextOutline.Write_ToBinary(Writer);
		Flags |= 134217728;
	}
	if (undefined != this.TextFill)
	{
		this.TextFill.Write_ToBinary(Writer);
		Flags |= 268435456;
	}

	if (undefined != this.PrChange)
	{
		this.PrChange.WriteToBinary(Writer);
		this.ReviewInfo.WriteToBinary(Writer);
		Flags |= 536870912;
	}

	if (undefined != this.HighlightColor)
	{
		this.HighlightColor.Write_ToBinary(Writer);
		Flags |= 1073741824;
	}

	if (undefined != this.Ligatures)
	{
		Writer.WriteByte(this.Ligatures);
		Flags |= (1 << 31);
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CTextPr.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	// Bold
	if (Flags & 1)
		this.Bold = Reader.GetBool();

	// Italic
	if (Flags & 2)
		this.Italic = Reader.GetBool();

	// Underline
	if (Flags & 4)
		this.Underline = Reader.GetBool();

	// Strikeout
	if (Flags & 8)
		this.Strikeout = Reader.GetBool();

	// FontFamily
	if (Flags & 16)
		this.FontFamily = {Name : Reader.GetString2(), Index : -1};

	// FontSize
	if (Flags & 32)
		this.FontSize = Reader.GetDouble();

	// Color
	if (Flags & 64)
	{
		this.Color = new CDocumentColor(0, 0, 0);
		this.Color.Read_FromBinary(Reader);
	}

	// VertAlign
	if (Flags & 128)
		this.VertAlign = Reader.GetLong();

	// HighLight
	if (Flags & 256)
	{
		var HL_type = Reader.GetLong();
		if (highlight_None == HL_type)
			this.HighLight = highlight_None;
		else
		{
			this.HighLight = new CDocumentColor(0, 0, 0);
			this.HighLight.Read_FromBinary(Reader);
		}
	}

	// RStyle
	if (Flags & 512)
		this.RStyle = Reader.GetString2();

	// Spacing
	if (Flags & 1024)
		this.Spacing = Reader.GetDouble();

	// DStrikeout
	if (Flags & 2048)
		this.DStrikeout = Reader.GetBool();

	// Caps
	if (Flags & 4096)
		this.Caps = Reader.GetBool();

	// SmallCaps
	if (Flags & 8192)
		this.SmallCaps = Reader.GetBool();

	// Position
	if (Flags & 16384)
		this.Position = Reader.GetDouble();

	// RFonts
	if (Flags & 32768)
		this.RFonts.Read_FromBinary(Reader);

	// BoldCS
	if (Flags & 65536)
		this.BoldCS = Reader.GetBool();

	// ItalicCS
	if (Flags & 131072)
		this.ItalicCS = Reader.GetBool();

	// FontSizeCS
	if (Flags & 262144)
		this.FontSizeCS = Reader.GetDouble();

	// CS
	if (Flags & 524288)
		this.CS = Reader.GetBool();

	// RTL
	if (Flags & 1048576)
		this.RTL = Reader.GetBool();

	// Lang
	if (Flags & 2097152)
		this.Lang.Read_FromBinary(Reader);

	// Unifill
	if (Flags & 4194304)
	{
		this.Unifill = new AscFormat.CUniFill()
		this.Unifill.Read_FromBinary(Reader);
	}

	// Shd
	if (Flags & 8388608)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Read_FromBinary(Reader);
	}

	// Vanish
	if (Flags & 16777216)
		this.Vanish = Reader.GetBool();

	if (Flags & 33554432)
	{
		this.FontRef = new AscFormat.FontRef();
		this.FontRef.Read_FromBinary(Reader);
	}

	if (Flags & 67108864)
	{
		this.PrChange = new CTextPr();
		this.PrChange.Read_FromBinary(Reader);
	}

	if (Flags & 134217728)
	{
		this.TextOutline = new AscFormat.CLn();
		this.TextOutline.Read_FromBinary(Reader);
	}

	if (Flags & 268435456)
	{
		this.TextFill = new AscFormat.CUniFill();
		this.TextFill.Read_FromBinary(Reader);
	}

	if (Flags & 536870912)
	{
		this.PrChange   = new CTextPr();
		this.ReviewInfo = new AscWord.ReviewInfo();
		this.PrChange.ReadFromBinary(Reader);
		this.ReviewInfo.ReadFromBinary(Reader);
	}

	if (Flags & 1073741824)
	{
		this.HighlightColor = new AscFormat.CUniColor();
		this.HighlightColor.Read_FromBinary(Reader);
	}

	if (Flags & (1 << 31))
		this.Ligatures = Reader.GetByte();
};
CTextPr.prototype.Check_NeedRecalc = function()
{
	// Потому что в параграфе внутри Internal_recalculate_0 кэшируются ParaTextPr
	return true;

	if (undefined != this.Bold)
		return true;

	if (undefined != this.Italic)
		return true;

	if (undefined != this.FontFamily)
		return true;

	if (undefined != this.FontSize)
		return true;

	if (undefined != this.VertAlign)
		return true;

	if (undefined != this.Spacing)
		return true;

	if (undefined != this.Caps)
		return true;

	if (undefined != this.SmallCaps)
		return true;

	if (undefined != this.Position)
		return true;

	if (undefined != this.RFonts.Ascii)
		return true;

	if (undefined != this.RFonts.EastAsia)
		return true;

	if (undefined != this.RFonts.HAnsi)
		return true;

	if (undefined != this.RFonts.CS)
		return true;

	if (undefined != this.RTL || undefined != this.CS || undefined != this.BoldCS || undefined != this.ItalicCS || undefined != this.FontSizeCS)
		return true;

	if (undefined != this.Lang.Val)
		return true;

	// Потому что в параграфе внутри Internal_recalculate_0 кэшируются ParaTextPr
	if (undefined != this.Color)
		return true;

	if (undefined != this.HighLight)
		return true;

	if (undefined != this.Shd)
		return true;

	return false;
};
CTextPr.prototype.getFontCoef = function()
{
	return (AscCommon.vertalign_SubScript === this.VertAlign || AscCommon.vertalign_SubScript === this.VertAlign ? AscCommon.vaKSize : 1);
};
CTextPr.prototype.Document_Get_AllFontNames = function(AllFonts)
{
	if (undefined != this.RFonts.Ascii)
		AllFonts[this.RFonts.Ascii.Name] = true;

	if (undefined != this.RFonts.HAnsi)
		AllFonts[this.RFonts.HAnsi.Name] = true;

	if (undefined != this.RFonts.EastAsia)
		AllFonts[this.RFonts.EastAsia.Name] = true;

	if (undefined != this.RFonts.CS)
		AllFonts[this.RFonts.CS.Name] = true;
};
CTextPr.prototype.Document_CreateFontMap = function(FontMap, FontScheme)
{
	var Style   = ( true === this.Bold ? 1 : 0 ) + ( true === this.Italic ? 2 : 0 );
	var StyleCS = ( true === this.BoldCS ? 1 : 0 ) + ( true === this.ItalicCS ? 2 : 0 );
	var Size    = this.FontSize;
	var SizeCS  = this.FontSizeCS;

	var RFonts = this.RFonts;
	var CheckedName;
	if (undefined != RFonts.Ascii)
	{
		CheckedName  = FontScheme.checkFont(RFonts.Ascii.Name);
		var Key      = "" + CheckedName + "_" + Style + "_" + Size;
		FontMap[Key] =
			{
				Name  : CheckedName,
				Style : Style,
				Size  : Size
			};
	}

	if (undefined != RFonts.EastAsia)
	{
		CheckedName  = FontScheme.checkFont(RFonts.EastAsia.Name);
		var Key      = "" + CheckedName + "_" + Style + "_" + Size;
		FontMap[Key] =
			{
				Name  : CheckedName,
				Style : Style,
				Size  : Size
			};
	}

	if (undefined != RFonts.HAnsi)
	{
		CheckedName  = FontScheme.checkFont(RFonts.HAnsi.Name);
		var Key      = "" + CheckedName + "_" + Style + "_" + Size;
		FontMap[Key] =
			{
				Name  : CheckedName,
				Style : Style,
				Size  : Size
			};
	}

	if (undefined != RFonts.CS)
	{

		CheckedName  = FontScheme.checkFont(RFonts.CS.Name);
		var Key      = "" + CheckedName + "_" + StyleCS + "_" + SizeCS;
		FontMap[Key] =
			{
				Name  : CheckedName,
				Style : StyleCS,
				Size  : SizeCS
			};
	}
};
CTextPr.prototype.isEqual = function(TextPrOld, TextPrNew)
{
	if (TextPrOld == undefined || TextPrNew == undefined)
		return false;
	for (var TextPr in TextPrOld)
	{
		if (typeof TextPrOld[TextPr] == 'object')
		{
			/*for(var cpPr in TextPrOld[TextPr])
			 {
			 if(TextPrOld[TextPr][cpPr] != TextPrNew[TextPr][cpPr])
			 return false;
			 }*/
			this.isEqual(TextPrOld[TextPr], TextPrNew[TextPr]);
		}
		else
		{
			if (typeof TextPrOld[TextPr] == "number" && typeof TextPrNew[TextPr] == "number")
			{
				if (Math.abs(TextPrOld[TextPr] - TextPrNew[TextPr]) > 0.001)
					return false;
			}
			else if (TextPrOld[TextPr] != TextPrNew[TextPr])
				return false;
		}
	}
	return true;
};
CTextPr.prototype.Is_Equal = function(TextPr)
{
	if (!TextPr)
		return false;

	if (this.Bold !== TextPr.Bold)
		return false;

	if (this.Italic !== TextPr.Italic)
		return false;

	if (this.Strikeout !== TextPr.Strikeout)
		return false;

	if (this.Underline !== TextPr.Underline)
		return false;

	if ((undefined === this.FontFamily && undefined !== TextPr.FontFamily) || (undefined !== this.FontFamily && (undefined === TextPr.FontFamily || this.FontFamily.Name !== TextPr.FontFamily.Name)))
		return false;

	if ((undefined === this.FontSize
		&& undefined !== TextPr.FontSize)
		|| (undefined !== this.FontSize
			&& (undefined === TextPr.FontSize
				|| Math.abs(this.FontSize - TextPr.FontSize) >= 0.001)))
		return false;

	if ((undefined === this.Color && undefined !== TextPr.Color) || (undefined !== this.Color && (undefined === TextPr.Color || true !== this.Color.Compare(TextPr.Color))))
		return false;

	if (this.VertAlign !== TextPr.VertAlign)
		return false;

	if ((undefined === this.HighLight && undefined !== TextPr.HighLight) || (undefined !== this.HighLight && (undefined === TextPr.HighLight || (highlight_None === this.HighLight && highlight_None !== TextPr.HighLight) || (highlight_None !== this.HighLight && highlight_None === TextPr.HighLight) || (highlight_None !== this.HighLight && highlight_None !== TextPr.HighLight && true !== this.HighLight.Compare(TextPr.HighLight)))))
		return false;

	if (this.RStyle !== TextPr.RStyle)
		return false;

	if ((undefined === this.Spacing && undefined !== TextPr.Spacing) || (undefined !== this.Spacing && (undefined === TextPr.Spacing || Math.abs(this.Spacing - TextPr.Spacing) >= 0.001)))
		return false;

	if (this.DStrikeout !== TextPr.DStrikeout)
		return false;

	if (this.Caps !== TextPr.Caps)
		return false;

	if (this.SmallCaps !== TextPr.SmallCaps)
		return false;

	if ((undefined === this.Position && undefined !== TextPr.Position) || (undefined !== this.Position && (undefined === TextPr.Position || Math.abs(this.Position - TextPr.Position) >= 0.001)))
		return false;

	if (true !== this.RFonts.Is_Equal(TextPr.RFonts))
		return false;

	if (this.BoldCS !== TextPr.BoldCS)
		return false;

	if (this.ItalicCS !== TextPr.ItalicCS)
		return false;

	if ((undefined === this.FontSizeCS
		&& undefined !== TextPr.FontSizeCS)
		|| (undefined !== this.FontSizeCS
			&& (undefined === TextPr.FontSizeCS
				|| Math.abs(this.FontSizeCS - TextPr.FontSizeCS) >= 0.001)))
		return false;

	if (this.CS !== TextPr.CS)
		return false;

	if (this.RTL !== TextPr.RTL)
		return false;

	if (true !== this.Lang.Is_Equal(TextPr.Lang))
		return false;

	if ((undefined === this.Unifill && undefined !== TextPr.Unifill) || (undefined !== this.Unifill && (undefined === TextPr.Unifill || true !== this.Unifill.IsIdentical(TextPr.Unifill))))
		return false;

	if ((undefined === this.TextOutline && undefined !== TextPr.TextOutline) || (undefined !== this.TextOutline && (undefined === TextPr.TextOutline || true !== this.TextOutline.IsIdentical(TextPr.TextOutline))))
		return false;

	if ((undefined === this.TextFill && undefined !== TextPr.TextFill) || (undefined !== this.TextFill && (undefined === TextPr.TextFill || true !== this.TextFill.IsIdentical(TextPr.TextFill))))
		return false;

	if ((undefined === this.HighlightColor && undefined !== TextPr.HighlightColor) || (undefined !== this.HighlightColor && (undefined === TextPr.HighlightColor || true !== this.HighlightColor.IsIdentical(TextPr.HighlightColor))))
		return false;

	if (this.Vanish !== TextPr.Vanish)
		return false;

	if (this.Ligatures !== TextPr.Ligatures)
		return false;

	if (!IsEqualStyleObjects(this.Shd, TextPr.Shd))
		return false;
	if (undefined != TextPr.AscLine)
	{
		return false;
	}
	if (undefined != TextPr.AscUnifill)
	{
		return false;
	}
	if (undefined != TextPr.AscFill)
	{
		return false;
	}

	return true;
};
CTextPr.prototype.IsEqual = function(oTextPr)
{
	return this.Is_Equal(oTextPr);
};
CTextPr.prototype.Is_Empty = function()
{
	if (undefined !== this.Bold
		|| undefined !== this.Italic
		|| undefined !== this.Strikeout
		|| undefined !== this.Underline
		|| undefined !== this.FontFamily
		|| undefined !== this.FontSize
		|| undefined !== this.Color
		|| undefined !== this.VertAlign
		|| undefined !== this.HighLight
		|| undefined !== this.RStyle
		|| undefined !== this.Spacing
		|| undefined !== this.DStrikeout
		|| undefined !== this.Caps
		|| undefined !== this.SmallCaps
		|| undefined !== this.Position
		|| true !== this.RFonts.Is_Empty()
		|| undefined !== this.BoldCS
		|| undefined !== this.ItalicCS
		|| undefined !== this.FontSizeCS
		|| undefined !== this.CS
		|| undefined !== this.RTL
		|| true !== this.Lang.Is_Empty()
		|| undefined !== this.Unifill
		|| undefined !== this.FontRef
		|| undefined !== this.Shd
		|| undefined !== this.Vanish
		|| undefined !== this.Ligatures
		|| undefined !== this.TextOutline
		|| undefined !== this.TextFill
		|| undefined !== this.HighlightColor)
		return false;

	return true;
};
CTextPr.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
/**
 * Сравниваем данные настройки с заданными, если настройка совпала ставим undefined, если нет, то берем из текущей
 * @param oTextPr {CTextPr}
 * @returns {CTextPr}
 */
CTextPr.prototype.GetDiff = function(oTextPr)
{
	var oResultTextPr = new CTextPr();

	if (this.Bold !== oTextPr.Bold)
		oResultTextPr.Bold = this.Bold;

	if (this.Italic !== oTextPr.Italic)
		oResultTextPr.Italic = this.Italic;

	if (this.Strikeout !== oTextPr.Strikeout)
		oResultTextPr.Strikeout = this.Strikeout;

	if (this.Underline !== oTextPr.Underline)
		oResultTextPr.Underline = this.Underline;

	if (undefined !== this.FontSize && !IsEqualNullableFloatNumbers(this.FontSize, oTextPr.FontSize))
		oResultTextPr.FontSize = this.FontSize;

	if (undefined !== this.Color && !this.Color.IsEqual(oTextPr.Color))
		oResultTextPr.Color = this.Color.Copy();

	if (this.VertAlign !== oTextPr.VertAlign)
		oResultTextPr.VertAlign = this.VertAlign;

	if (undefined === this.HighLight)
	{
		oResultTextPr.HighLight = undefined;
	}
	else if (highlight_None === this.HighLight)
	{
		if (highlight_None !== oTextPr.HighLight)
			oResultTextPr.HighLight = highlight_None;
	}
	else
	{
		if (!this.HighLight.IsEqual(oTextPr.HighLight))
			oResultTextPr.HighLight = this.HighLight.Copy();
	}

	if (this.RStyle !== oTextPr.RStyle)
		oResultTextPr.RStyle = this.RStyle;

	if (undefined !== this.Spacing && !IsEqualNullableFloatNumbers(this.Spacing, oTextPr.Spacing))
		oResultTextPr.Spacing = this.Spacing;

	if (this.DStrikeout !== oTextPr.DStrikeout)
		oResultTextPr.DStrikeout = this.DStrikeout;

	if (this.Caps !== oTextPr.Caps)
		oResultTextPr.Caps = this.Caps;

	if (this.SmallCaps !== oTextPr.SmallCaps)
		oResultTextPr.SmallCaps = this.SmallCaps;

	if (undefined !== this.Position && !IsEqualNullableFloatNumbers(this.Position, oTextPr.Position))
		oResultTextPr.Position = this.Position;

	if (!this.RFonts.IsEqual(oTextPr.RFonts))
		oResultTextPr.RFonts = this.RFonts.Copy();

	if (this.BoldCS !== oTextPr.BoldCS)
		oResultTextPr.BoldCS = this.BoldCS;

	if (this.ItalicCS !== oTextPr.ItalicCS)
		oResultTextPr.ItalicCS = this.ItalicCS;

	if (undefined !== this.FontSizeCS && !IsEqualNullableFloatNumbers(this.FontSizeCS, oTextPr.FontSizeCS))
		oResultTextPr.FontSizeCS = this.FontSizeCS;

	if (this.CS !== oTextPr.CS)
		oResultTextPr.CS = this.CS;

	if (this.RTL !== oTextPr.RTL)
		oResultTextPr.RTL = this.RTL;

	if (!this.Lang.IsEqual(oTextPr.Lang))
		oResultTextPr.Lang = this.Lang.Copy();

	if (this.Unifill && !this.Unifill.IsIdentical(oTextPr.Unifill))
		oResultTextPr.Unifill = this.Unifill.createDuplicate();

	if (this.TextOutline && !this.TextOutline.IsIdentical(oTextPr.TextOutline))
		oResultTextPr.TextOutline = this.TextOutline.createDuplicate();

	if (this.TextFill && !this.TextFill.IsIdentical(oTextPr.TextFill))
		oResultTextPr.TextFill = this.TextFill.createDuplicate();

	if (this.HighlightColor && !this.HighlightColor.IsIdentical(oTextPr.HighlightColor))
		oResultTextPr.HighlightColor = this.HighlightColor.createDuplicate();

	if (this.Vanish !== oTextPr.Vanish)
		oResultTextPr.Vanish = this.Vanish;

	if (this.Ligatures !== oTextPr.Ligatures)
		oResultTextPr.Ligatures = this.Ligatures;

	if (this.Shd && !this.Shd.IsEqual(oTextPr.Shd))
		oResultTextPr.Shd = this.Shd.Copy();

	return oResultTextPr;
};
CTextPr.prototype.GetBold = function()
{
    return this.Bold;
};
CTextPr.prototype.SetBold = function(isBold)
{
	this.Bold = isBold;
};
CTextPr.prototype.GetItalic = function()
{
    return this.Italic;
};
CTextPr.prototype.SetItalic = function(isItalic)
{
	this.Italic = isItalic;
};
CTextPr.prototype.GetStrikeout = function()
{
    return this.Strikeout;
};
CTextPr.prototype.SetStrikeout = function(isStrikeout)
{
	this.Strikeout = isStrikeout;
};
CTextPr.prototype.GetUnderline = function()
{
    return this.Underline;
};
CTextPr.prototype.SetUnderline = function(isUnderling)
{
	this.Underline = isUnderling;
};
CTextPr.prototype.GetColor = function()
{
    return this.Color;
};
CTextPr.prototype.SetColor = function(nR, nG, nB, isAuto)
{
	if (undefined === nR)
		this.Color = undefined;
	else
		this.Color = new CDocumentColor(nR, nG, nB, isAuto);
};
CTextPr.prototype.GetAscColor = function()
{
	if (this.Unifill && this.Unifill.fill && this.Unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID && this.Unifill.fill.color)
	{
		return AscCommon.CreateAscColor(this.Unifill.fill.color);
	}
	else if (this.Color)
	{
		return AscCommon.CreateAscColorCustom(this.Color.r, this.Color.g, this.Color.b, this.Color.Auto);
	}

	return undefined;
};
CTextPr.prototype.SetAscColor = function(oAscColor)
{
	if (!oAscColor)
	{
		this.Color   = undefined;
		this.Unifill = undefined;
	}
	else if (true === oAscColor.Auto)
	{
		this.Color   = new CDocumentColor(0, 0, 0, true);
		this.Unifill = undefined;
	}
	else
	{
		this.Color              = undefined;
		this.Unifill            = new AscFormat.CUniFill();
		this.Unifill.fill       = new AscFormat.CSolidFill();
		this.Unifill.fill.color = AscFormat.CorrectUniColor(oAscColor, this.Unifill.fill.color, 1);

		var oLogicDocument = private_GetWordLogicDocument();
		if (oLogicDocument)
			this.Unifill.check(oLogicDocument.GetTheme(), oLogicDocument.GetColorMap());

	}
};
CTextPr.prototype.SetUnifill = function(oUnifill)
{
	this.Unifill = oUnifill;
};
CTextPr.prototype.FillFromExcelFont = function(oFont)
{
	if (!oFont)
		return;

	var nSchemeFont = oFont.getScheme();
	switch (nSchemeFont)
	{
		case Asc.EFontScheme.fontschemeMajor:
		{
			this.RFonts.SetFontStyle(AscFormat.fntStyleInd_major);
			break;
		}
		case Asc.EFontScheme.fontschemeMinor:
		{
			this.RFonts.SetFontStyle(AscFormat.fntStyleInd_minor);
			break;
		}
		case Asc.EFontScheme.fontschemeNone:
		{
			this.RFonts.SetAll(oFont.getName(), -1);
			break;
		}
	}
	this.SetFontSize(oFont.getSize());
	this.SetBold(oFont.getBold());
	this.SetItalic(oFont.getItalic());
	let bUnderline = (oFont.getUnderline() !== Asc.EUnderline.underlineNone);
	this.SetUnderline(bUnderline);
	var oColor = oFont.getColor();
	this.SetUnifill(AscFormat.CreateSolidFillRGBA(oColor.getR(), oColor.getG(), oColor.getB(), 255));
};
CTextPr.prototype.GetVertAlign = function()
{
    return this.VertAlign;
};
CTextPr.prototype.SetVertAlign = function(nVertAlign)
{
	this.VertAlign = nVertAlign;
};
CTextPr.prototype.GetHighlight = function()
{
    return this.HighLight;
};
CTextPr.prototype.SetHighlight = function(nHighlight)
{
	this.HighLight = nHighlight;
};
CTextPr.prototype.GetSpacing = function()
{
    return this.Spacing;
};
CTextPr.prototype.SetSpacing = function(nSpacing)
{
	this.Spacing = nSpacing;
};
CTextPr.prototype.GetDStrikeout = function()
{
    return this.DStrikeout;
};
CTextPr.prototype.SetDStrikeout = function(isDStrikeout)
{
	this.DStrikeout = isDStrikeout;
};
CTextPr.prototype.GetCaps = function()
{
    return this.Caps;
};
CTextPr.prototype.SetCaps = function(isCaps)
{
	this.Caps = isCaps;
};
CTextPr.prototype.GetSmallCaps = function()
{
    return this.SmallCaps;
};
CTextPr.prototype.SetSmallCaps = function(isSmallCaps)
{
	this.SmallCaps = isSmallCaps;
};
CTextPr.prototype.GetPosition = function()
{
    return this.Position;
};
CTextPr.prototype.SetPosition = function(nPosition)
{
	this.Position = nPosition;
};
CTextPr.prototype.GetFontFamily = function(fontSlot)
{
	if (!this.RFonts)
		return undefined;
	
	if (AscWord.fontslot_CS === fontSlot && this.RFonts.CS)
		return this.RFonts.CS.Name;
	else if (AscWord.fontslot_HAnsi === fontSlot && this.RFonts.HAnsi)
		return this.RFonts.HAnsi.Name;
	else if (AscWord.fontslot_EastAsia === fontSlot && this.RFonts.EastAsia)
		return this.RFonts.EastAsia.Name;
	else if (this.RFonts.Ascii)
		return this.RFonts.Ascii.Name;
	
	return undefined;
};
CTextPr.prototype.SetFontFamily = function(sFontName)
{
	if (!this.RFonts || typeof sFontName !== 'string')
		this.RFonts = new CRFonts();

	if (typeof sFontName === 'string')
		this.RFonts.SetAll(sFontName);
};
CTextPr.prototype.GetFontSize = function()
{
    return this.FontSize;
};
CTextPr.prototype.SetFontSize = function(nFontSize)
{
	this.FontSize = nFontSize;
};
CTextPr.prototype.GetLang = function()
{
    if (this.Lang)
        return this.Lang.Val;

    return undefined;
};
CTextPr.prototype.SetLang = function(nLang)
{
	if (!this.Lang)
		this.Lang = new CLang();

	this.Lang.Val = nLang;
};
CTextPr.prototype.GetShd = function()
{
    return this.Shd;
};
CTextPr.prototype.SetShd = function(nType, nR, nG, nB)
{
	if (!this.Shd)
		this.Shd = new CDocumentShd();

	this.Shd.Value = nType;

	if (Asc.c_oAscShdNil === nType)
		return;

	this.Shd.Color.Set(nR, nG, nB);
};
CTextPr.prototype.SetLigatures = function(nType)
{
	this.Ligatures = nType;
};
CTextPr.prototype.GetLigatures = function()
{
	return this.Ligatures;
};
CTextPr.prototype.WriteToBinary = function(oWriter)
{
	return this.Write_ToBinary(oWriter);
};
CTextPr.prototype.ReadFromBinary = function(oReader)
{
	return this.Read_FromBinary(oReader);
};
CTextPr.prototype.GetAllFontNames = function(arrAllFonts)
{
	return this.Document_Get_AllFontNames(arrAllFonts);
};
CTextPr.prototype.HavePrChange = function()
{
	if (undefined === this.PrChange || null === this.PrChange)
		return false;

	return true;
};
CTextPr.prototype.AddPrChange = function()
{
	this.PrChange   = this.Copy();
	this.ReviewInfo = new AscWord.ReviewInfo();
	this.ReviewInfo.Update();
};
CTextPr.prototype.SetPrChange = function(PrChange, ReviewInfo)
{
	this.PrChange   = PrChange;
	this.ReviewInfo = ReviewInfo;
};
CTextPr.prototype.RemovePrChange = function()
{
	delete this.PrChange;
	delete this.ReviewInfo;
};
CTextPr.prototype.GetDiffPrChange = function()
{
	var TextPr = new CTextPr();

	if (false === this.HavePrChange())
		return TextPr;

	var PrChange = this.PrChange;
	if (this.Bold !== PrChange.Bold)
		TextPr.Bold = this.Bold;

	if (this.Italic !== PrChange.Italic)
		TextPr.Italic = this.Italic;

	if (this.Strikeout !== PrChange.Strikeout)
		TextPr.Strikeout = this.Strikeout;

	if (this.Underline !== PrChange.Underline)
		TextPr.Underline = this.Underline;

	if (undefined !== this.FontFamily && (undefined === PrChange.FontFamily || this.FontFamily.Name !== PrChange.FontFamily.Name))
		TextPr.FontFamily = {Name : this.FontFamily.Name, Index : -1};

	if (undefined !== this.FontSize && (undefined === PrChange.FontSize || Math.abs(this.FontSize - PrChange.FontSize) >= 0.001))
		TextPr.FontSize = this.FontSize;

	if (undefined !== this.Color && (undefined === PrChange.Color || true !== this.Color.Compare(PrChange.Color)))
		TextPr.Color = this.Color.Copy();

	if (this.VertAlign !== PrChange.VertAlign)
		TextPr.VertAlign = this.VertAlign;

	if (highlight_None === this.HighLight)
	{
		if (highlight_None !== PrChange.HighLight)
			TextPr.HighLight = highlight_None;
	}
	else if (undefined !== this.HighLight)
	{
		if (undefined === PrChange.HighLight || highlight_None === PrChange.HighLight || true !== this.HighLight.Compare(PrChange.HighLight))
			TextPr.HighLight = this.HighLight.Copy();
	}

	if (this.RStyle !== PrChange.RStyle)
		TextPr.RStyle = this.RStyle;

	if (undefined !== this.Spacing && (undefined === PrChange.Spacing || Math.abs(this.Spacing - PrChange.Spacing) >= 0.001))
		TextPr.Spacing = this.Spacing;

	if (this.DStrikeout !== PrChange.DStrikeout)
		TextPr.DStrikeout = this.DStrikeout;

	if (this.Caps !== PrChange.Caps)
		TextPr.Caps = this.Caps;

	if (this.SmallCaps !== PrChange.SmallCaps)
		TextPr.SmallCaps = this.SmallCaps;

	if (undefined !== this.Position && (undefined === PrChange.Position || Math.abs(this.Position - PrChange.Position) >= 0.001))
		TextPr.Position = this.Position;

	if (undefined !== this.RFonts && (undefined === PrChange.RFonts || true !== this.RFonts.Is_Equal(TextPr.RFonts)))
		TextPr.RFonts = this.RFonts.Copy();

	if (this.BoldCS !== PrChange.BoldCS)
		TextPr.BoldCS = this.BoldCS;

	if (this.ItalicCS !== PrChange.ItalicCS)
		TextPr.ItalicCS = this.ItalicCS;

	if (undefined !== this.FontSizeCS && (undefined === PrChange.FontSizeCS || Math.abs(this.FontSizeCS - PrChange.FontSizeCS) >= 0.001))
		TextPr.FontSizeCS = this.FontSizeCS;

	if (this.CS !== PrChange.CS)
		TextPr.CS = this.CS;

	if (this.RTL !== PrChange.RTL)
		TextPr.RTL = this.RTL;

	if (undefined !== this.Lang && (undefined === PrChange.Lang || true !== this.Lang.Is_Equal(PrChange.Lang)))
		TextPr.Lang = this.Lang.Copy();

	if (undefined !== this.Unifill && (undefined === PrChange.Unifill || true !== this.Unifill.IsIdentical(PrChange.Unifill)))
		TextPr.Unifill = this.Unifill.createDuplicate();

	if (undefined !== this.TextOutline && (undefined === PrChange.TextOutline || true !== this.TextOutline.IsIdentical(PrChange.TextOutline)))
		TextPr.TextOutline = this.TextOutline.createDuplicate();

	if (undefined !== this.TextFill && (undefined === PrChange.TextFill || true !== this.TextFill.IsIdentical(PrChange.TextFill)))
		TextPr.TextFill = this.TextFill.createDuplicate();

	if (undefined !== this.HighlightColor && (undefined === PrChange.HighlightColor || true !== this.HighlightColor.IsIdentical(PrChange.HighlightColor)))
		TextPr.HighlightColor = this.HighlightColor.createDuplicate();

	if (this.Vanish !== PrChange.Vanish)
		TextPr.Vanish = this.Vanish;

	// TODO: Shd

	return TextPr;
};
CTextPr.prototype.GetDescription = function()
{
	var Description = "Text formatting: ";

	if (undefined !== this.Bold)
		Description += this.Bold ? "Bold; " : "No Bold; ";

	if (undefined !== this.Italic)
		Description += this.Italic ? "Italic; " : "No Italic; ";

	if (undefined !== this.Strikeout)
		Description += this.Strikeout ? "Strikeout; " : "No Strikeout; ";

	if (undefined !== this.DStrikeout)
		Description += this.DStrikeout ? "Double Strikeout; " : "No Double Strikeout; ";

	if (undefined !== this.FontSize)
		Description += this.FontSize + "FontSize; ";

	//        if (undefined !== this.Color && (undefined === PrChange.Color || true !== this.Color.Compare(PrChange.Color)))
	//            TextPr.Color = this.Color.Copy();
	//
	//        if (this.VertAlign !== PrChange.VertAlign)
	//            TextPr.VertAlign = this.VertAlign;
	//
	//        if (highlight_None === this.HighLight)
	//        {
	//            if (highlight_None !== PrChange.HighLight)
	//                TextPr.HighLight = highlight_None;
	//        }
	//        else if (undefined !== this.HighLight)
	//        {
	//            if (undefined === PrChange.HighLight || highlight_None === PrChange.HighLight || true !== this.HighLight.Compare(PrChange.HighLight))
	//                TextPr.HighLight = this.HighLight.Copy();
	//        }
	//
	//        if (this.RStyle !== PrChange.RStyle)
	//            TextPr.RStyle = this.RStyle;
	//
	//        if (undefined !== this.Spacing && (undefined === PrChange.Spacing || Math.abs(this.Spacing - PrChange.Spacing) >= 0.001))
	//            TextPr.Spacing = this.Spacing;
	//
	//        if (this.Caps !== PrChange.Caps)
	//            TextPr.Caps = this.Caps;
	//
	//        if (this.SmallCaps !== PrChange.SmallCaps)
	//            TextPr.SmallCaps = this.SmallCaps;
	//
	//        if (undefined !== this.Position && (undefined === PrChange.Position || Math.abs(this.Position - PrChange.Position) >= 0.001))
	//            TextPr.Position = this.Position;
	//
	//        if (undefined !== this.RFonts && (undefined === PrChange.RFonts || true !== this.RFonts.Is_Equal(TextPr.RFonts)))
	//            TextPr.RFonts = this.RFonts.Copy();
	//
	//        if (this.BoldCS !== PrChange.BoldCS)
	//            TextPr.BoldCS = this.BoldCS;
	//
	//        if (this.ItalicCS !== PrChange.ItalicCS)
	//            TextPr.ItalicCS = this.ItalicCS;
	//
	//        if (undefined !== this.FontSizeCS && (undefined === PrChange.FontSizeCS || Math.abs(this.FontSizeCS - PrChange.FontSizeCS) >= 0.001))
	//            TextPr.FontSizeCS = this.FontSizeCS;
	//
	//        if (this.CS !== PrChange.CS)
	//            TextPr.CS = this.CS;
	//
	//        if (this.RTL !== PrChange.RTL)
	//            TextPr.RTL = this.RTL;
	//
	//        if (undefined !== this.Lang && (undefined === PrChange.Lang || true !== this.Lang.Is_Equal(PrChange.Lang)))
	//            TextPr.Lang = this.Lang.Copy();
	//
	//        if (undefined !== this.Unifill && (undefined === PrChange.Unifill || true !== this.Unifill.IsIdentical(PrChange.Unifill)))
	//            TextPr.Unifill = this.Unifill.createDuplicate();
	//
	//        if (undefined !== this.TextOutline && (undefined === PrChange.TextOutline || true !== this.TextOutline.IsIdentical(PrChange.TextOutline)))
	//            TextPr.TextOutline = this.TextOutline.createDublicate();
	//
	//        if (undefined !== this.TextFill && (undefined === PrChange.TextFill || true !== this.TextFill.IsIdentical(PrChange.TextFill)))
	//            TextPr.TextFill = this.TextFill.createDublicate();
	//
	//        if (this.Vanish !== PrChange.Vanish)
	//            TextPr.Vanish = this.Vanish;

	return Description;
};
CTextPr.prototype.GetTextMetrics = function(nFontFlags, oTheme)
{
	let oMetrics = new CTextMetrics();

	let oTextPr = this;
	if (this.RFonts.HaveThemeFonts())
	{
		oTextPr = this.Copy();
		oTextPr.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
	}

	if ((nFontFlags & AscWord.fontslot_ASCII) && oTextPr.RFonts.Ascii)
		oMetrics.UpdateByFontInfo(oTextPr.GetFontInfo(AscWord.fontslot_ASCII));

	if ((nFontFlags & AscWord.fontslot_CS) && oTextPr.RFonts.CS)
		oMetrics.UpdateByFontInfo(oTextPr.GetFontInfo(AscWord.fontslot_CS));

	if ((nFontFlags & AscWord.fontslot_HAnsi) && oTextPr.RFonts.HAnsi)
		oMetrics.UpdateByFontInfo(oTextPr.GetFontInfo(AscWord.fontslot_HAnsi));

	if ((nFontFlags & AscWord.fontslot_EastAsia) && oTextPr.RFonts.EastAsia)
		oMetrics.UpdateByFontInfo(oTextPr.GetFontInfo(AscWord.fontslot_EastAsia));

	return oMetrics;
};
CTextPr.prototype.GetFontInfo = function(nFontSlot)
{
	let sFontName = this.RFonts.Ascii.Name;
	let isBold    = this.Bold;
	let isItalic  = this.Italic;
	let nFontSize = this.FontSize;

	switch (nFontSlot)
	{
		case AscWord.fontslot_CS:
		{
			sFontName = this.RFonts.CS.Name;
			isBold    = this.BoldCS;
			isItalic  = this.ItalicCS;
			nFontSize = this.FontSizeCS;
			break;
		}
		case AscWord.fontslot_EastAsia:
		{
			sFontName = this.RFonts.EastAsia.Name;
			break;
		}
		case AscWord.fontslot_HAnsi:
		{
			sFontName = this.RFonts.HAnsi.Name;
			break;
		}
	}

	return new AscFonts.CTextFontInfo(sFontName, (isBold ? 1  : 0) | (isItalic ? 2 : 0), nFontSize);
};

function CTextMetrics()
{
	this.Ascent  = 0;
	this.Descent = 0;
	this.LineGap = 0;
	this.Height  = 0;
}
/**
 * @param {AscFonts.CTextFontInfo} fontInfo
 */
CTextMetrics.prototype.UpdateByFontInfo = function(fontInfo)
{
	this.Update(fontInfo.Name, fontInfo.Size, fontInfo.Style);
};
CTextMetrics.prototype.Update = function(fontName, fontSize, fontStyle)
{
	g_oTextMeasurer.SetFontInternal(fontName, fontSize, fontStyle);
	
	let nHeight  = g_oTextMeasurer.GetHeight();
	let nAscent  = g_oTextMeasurer.GetAscender();
	let nDescent = Math.abs(g_oTextMeasurer.GetDescender());
	
	let _nHeight  = nHeight;
	let _nDescent = nDescent;
	let _nAscent  = Math.min(nAscent, nHeight - nDescent);
	let _nLineGap = Math.max(0, nHeight - nAscent - nDescent);
	
	if (this.Height < _nHeight)
		this.Height = _nHeight;
	
	if (this.Descent < _nDescent)
		this.Descent = _nDescent;
	
	if (this.Ascent < _nAscent)
		this.Ascent = _nAscent;
	
	if (this.LineGap < _nLineGap)
		this.LineGap = _nLineGap;
};
//----------------------------------------------------------------------------------------------------------------------
// CTextPr Export
//----------------------------------------------------------------------------------------------------------------------
CTextPr.prototype['get_Bold']       = CTextPr.prototype.get_Bold       = CTextPr.prototype['Get_Bold']       = CTextPr.prototype.GetBold;
CTextPr.prototype['put_Bold']       = CTextPr.prototype.put_Bold       = CTextPr.prototype.SetBold;
CTextPr.prototype['get_Italic']     = CTextPr.prototype.get_Italic     = CTextPr.prototype['Get_Italic']     = CTextPr.prototype.GetItalic;
CTextPr.prototype['put_Italic']     = CTextPr.prototype.put_Italic     = CTextPr.prototype.SetItalic;
CTextPr.prototype['get_Strikeout']  = CTextPr.prototype.get_Strikeout  = CTextPr.prototype['Get_Strikeout']  = CTextPr.prototype.GetStrikeout;
CTextPr.prototype['put_Strikeout']  = CTextPr.prototype.put_Strikeout  = CTextPr.prototype.SetStrikeout;
CTextPr.prototype['get_Underline']  = CTextPr.prototype.get_Underline  = CTextPr.prototype['Get_Underline']  = CTextPr.prototype.GetUnderline;
CTextPr.prototype['put_Underline']  = CTextPr.prototype.put_Underline  = CTextPr.prototype.SetUnderline;
CTextPr.prototype['get_Color']      = CTextPr.prototype.get_Color      = CTextPr.prototype['Get_Color']      = CTextPr.prototype.GetAscColor;
CTextPr.prototype['put_Color']      = CTextPr.prototype.put_Color      = CTextPr.prototype.SetAscColor;
CTextPr.prototype['get_VertAlign']  = CTextPr.prototype.get_VertAlign  = CTextPr.prototype['Get_VertAlign']  = CTextPr.prototype.GetVertAlign;
CTextPr.prototype['put_VertAlign']  = CTextPr.prototype.put_VertAlign  = CTextPr.prototype.SetVertAlign;
CTextPr.prototype['get_Highlight']  = CTextPr.prototype.get_Highlight  = CTextPr.prototype['Get_Highlight']  = CTextPr.prototype.GetHighlight;
CTextPr.prototype['put_Highlight']  = CTextPr.prototype.put_Highlight  = CTextPr.prototype.SetHighlight;
CTextPr.prototype['get_Spacing']    = CTextPr.prototype.get_Spacing    = CTextPr.prototype['Get_Spacing']    = CTextPr.prototype.GetSpacing;
CTextPr.prototype['put_Spacing']    = CTextPr.prototype.put_Spacing    = CTextPr.prototype.SetSpacing;
CTextPr.prototype['get_DStrikeout'] = CTextPr.prototype.get_DStrikeout = CTextPr.prototype['Get_DStrikeout'] = CTextPr.prototype.GetDStrikeout;
CTextPr.prototype['put_DStrikeout'] = CTextPr.prototype.put_DStrikeout = CTextPr.prototype.SetDStrikeout;
CTextPr.prototype['get_Caps']       = CTextPr.prototype.get_Caps       = CTextPr.prototype['Get_Caps']       = CTextPr.prototype.GetCaps;
CTextPr.prototype['put_Caps']       = CTextPr.prototype.put_Caps       = CTextPr.prototype.SetCaps;
CTextPr.prototype['get_SmallCaps']  = CTextPr.prototype.get_SmallCaps  = CTextPr.prototype['Get_SmallCaps']  = CTextPr.prototype.GetSmallCaps;
CTextPr.prototype['put_SmallCaps']  = CTextPr.prototype.put_SmallCaps  = CTextPr.prototype.SetSmallCaps;
CTextPr.prototype['get_Position']   = CTextPr.prototype.get_Position   = CTextPr.prototype['Get_Position']   = CTextPr.prototype.GetPosition;
CTextPr.prototype['put_Position']   = CTextPr.prototype.put_Position   = CTextPr.prototype.SetPosition;
CTextPr.prototype['get_FontFamily'] = CTextPr.prototype.get_FontFamily = CTextPr.prototype['Get_FontFamily'] = CTextPr.prototype.GetFontFamily;
CTextPr.prototype['put_FontFamily'] = CTextPr.prototype.put_FontFamily = CTextPr.prototype.SetFontFamily;
CTextPr.prototype['get_FontSize']   = CTextPr.prototype.get_FontSize   = CTextPr.prototype['Get_FontSize']   = CTextPr.prototype.GetFontSize;
CTextPr.prototype['put_FontSize']   = CTextPr.prototype.put_FontSize   = CTextPr.prototype.SetFontSize;
CTextPr.prototype['get_Lang']       = CTextPr.prototype.get_Lang       = CTextPr.prototype['Get_Lang']       = CTextPr.prototype.GetLang;
CTextPr.prototype['put_Lang']       = CTextPr.prototype.put_Lang       = CTextPr.prototype.SetLang;
CTextPr.prototype['get_Shd']        = CTextPr.prototype.get_Shd        = CTextPr.prototype['Get_Shd']        = CTextPr.prototype.GetShd;
CTextPr.prototype['put_Shd']        = CTextPr.prototype.put_Shd        = CTextPr.prototype.SetShd;
CTextPr.prototype['get_Ligatures']  = CTextPr.prototype.GetLigatures;
CTextPr.prototype['put_Ligatures']  = CTextPr.prototype.SetLigatures;
//----------------------------------------------------------------------------------------------------------------------

function CParaTab(Value, Pos, Leader)
{
	this.Value  = Value;
	this.Pos    = Pos;
	this.Leader = undefined !== Leader ? Leader : Asc.c_oAscTabLeader.None;
}
CParaTab.prototype.Copy = function()
{
	return new CParaTab(this.Value, this.Pos, this.Leader);
};
CParaTab.prototype.Is_Equal = function(Tab)
{
	return this.IsEqual(Tab);
};
CParaTab.prototype.IsEqual = function(Tab)
{
	// TODO: Если таб точно такого же типа и в той же позиции, то неясно нужно ли проверять совпадение Tab.Leader
	if (this.Value !== Tab.Value
		|| this.Pos !== Tab.Pos)
		return false;

	return true;
};
CParaTab.prototype.IsRightTab = function()
{
	return this.Value === tab_Right;
};
CParaTab.prototype.IsLeftTab = function()
{
	return this.Value === tab_Left;
};
CParaTab.prototype.IsCenterTab = function()
{
	return this.Value === tab_Center;
};
CParaTab.prototype.GetLeader = function()
{
	return this.Leader;
};
CParaTab.prototype.IsValid = function()
{
	return null != this.Pos && null != this.Value;
};

function CParaTabs()
{
	this.Tabs = [];
}
CParaTabs.prototype.Add = function(_Tab)
{
	var Index = 0;
	for (Index = 0; Index < this.Tabs.length; Index++)
	{
		var Tab = this.Tabs[Index];

		if (Math.abs(Tab.Pos - _Tab.Pos) < 0.001)
		{
			this.Tabs.splice(Index, 1, _Tab);
			break;
		}

		if (Tab.Pos > _Tab.Pos)
			break;
	}

	if (-1 != Index)
		this.Tabs.splice(Index, 0, _Tab);
};
CParaTabs.prototype.Merge = function(Tabs)
{
	var _Tabs = Tabs.Tabs;

	for (var Index = 0; Index < _Tabs.length; Index++)
	{
		var _Tab = _Tabs[Index];

		var Index2 = 0;
		var Flag   = 0;
		for (Index2 = 0; Index2 < this.Tabs.length; Index2++)
		{
			var Tab = this.Tabs[Index2];

			if (Math.abs(Tab.Pos - _Tab.Pos) < 0.001)
			{
				if (tab_Clear === _Tab.Value)
					Flag = -2; // таб нужно удалить
				else if (Tab.Value !== _Tab.Value || Tab.Leader !== _Tab.Leader)
					Flag = -3; // таб нужно заменить
				else
					Flag = -1; // табы совпали, не надо новый добавлять

				break;
			}

			if (Tab.Pos > _Tab.Pos)
				break;
		}

		if (-2 === Flag)
			this.Tabs.splice(Index2, 1);
		else if (-3 === Flag)
			this.Tabs.splice(Index2, 1, _Tab);
		else if (-1 != Flag)
			this.Tabs.splice(Index2, 0, _Tab);
	}
};
CParaTabs.prototype.IsEqual = function(Tabs)
{
	if (this.Tabs.length !== Tabs.Tabs.length)
		return false;

	for (var CurTab = 0, TabsCount = this.Tabs.length; CurTab < TabsCount; CurTab++)
	{
		if (true !== this.Tabs[CurTab].IsEqual(Tabs.Tabs[CurTab]))
			return false;
	}

	return true;
};
CParaTabs.prototype.Is_Equal = function(Tabs)
{
	return this.IsEqual(Tabs);
};
CParaTabs.prototype.Copy = function()
{
	var Tabs  = new CParaTabs();
	var Count = this.Tabs.length;

	for (var Index = 0; Index < Count; Index++)
		Tabs.Add(this.Tabs[Index].Copy());

	return Tabs;
};
CParaTabs.prototype.Set_FromObject = function(Tabs)
{
	if (Tabs instanceof Array)
	{
		for (var nIndex = 0, nCount = Tabs.length; nIndex < nCount; ++nIndex)
			this.Add(new CParaTab(Tabs[nIndex].Value, Tabs[nIndex].Pos, Tabs[nIndex].Leader));
	}
};
CParaTabs.prototype.GetCount = function()
{
	return this.Tabs.length;
};
CParaTabs.prototype.Get_Count = function()
{
	return this.GetCount();
};
CParaTabs.prototype.Get = function(Index)
{
	return this.Tabs[Index];
};
CParaTabs.prototype.Get_Value = function(Pos)
{
	var Count = this.Tabs.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Tab = this.Tabs[Index];
		if (Math.abs(Tab.Pos - Pos) < 0.001)
			return Tab.Value;
	}

	return -1;
};
CParaTabs.prototype.Write_ToBinary = function(Writer)
{
	// Long : количество (если 0, удаляем элемент)
	// Массив
	// Byte   : Value
	// Double : Pos
	// Long   : Leader

	var Count = this.Tabs.length;
	Writer.WriteLong(Count);

	for (var Index = 0; Index < Count; Index++)
	{
		Writer.WriteByte(this.Tabs[Index].Value);
		Writer.WriteDouble(this.Tabs[Index].Pos);
		Writer.WriteLong(this.Tabs[Index].Leader);
	}
};
CParaTabs.prototype.Read_FromBinary = function(Reader)
{
	// Long : количество (если 0, удаляем элемент)
	// Массив
	// Byte   : Value
	// Double : Pos
	// Long   : Leader

	var Count = Reader.GetLong();
	this.Tabs = [];

	for (var Index = 0; Index < Count; Index++)
	{
		var Value  = Reader.GetByte();
		var Pos    = Reader.GetDouble();
		var Leader = Reader.GetLong();
		this.Add(new CParaTab(Value, Pos, Leader));
	}
};

function CParaInd()
{
    this.Left      = undefined; // Левый отступ
    this.Right     = undefined; // Правый отступ
    this.FirstLine = undefined; // Первая строка
}

CParaInd.prototype =
{
    Copy : function()
    {
        var Ind = new CParaInd();
        Ind.Left      = this.Left;
        Ind.Right     = this.Right;
        Ind.FirstLine = this.FirstLine;
        return Ind;
    },

    Merge : function(Ind)
    {
        if ( undefined != Ind.Left )
            this.Left = Ind.Left;

        if ( undefined != Ind.Right )
            this.Right = Ind.Right;

        if ( undefined != Ind.FirstLine )
            this.FirstLine = Ind.FirstLine;
    },

    Is_Equal  : function(Ind)
    {
    	return this.IsEqual(Ind);
    },

    Set_FromObject : function(Ind)
    {
        if ( undefined != Ind.Left )
            this.Left = Ind.Left;
        else
            this.Left = undefined;

        if ( undefined != Ind.Right )
            this.Right = Ind.Right;
        else
            this.Right = undefined;

        if ( undefined != Ind.FirstLine )
            this.FirstLine = Ind.FirstLine;
        else
            this.FirstLine = undefined;
    },

    Write_ToBinary : function(Writer)
    {
        var StartPos = Writer.GetCurPosition();
        Writer.Skip(4);
        var Flags = 0;

        if ( undefined != this.Left )
        {
            Writer.WriteDouble( this.Left );
            Flags |= 1;
        }

        if ( undefined != this.Right )
        {
            Writer.WriteDouble( this.Right );
            Flags |= 2;
        }

        if ( undefined != this.FirstLine )
        {
            Writer.WriteDouble( this.FirstLine );
            Flags |= 4;
        }

        var EndPos = Writer.GetCurPosition();
        Writer.Seek( StartPos );
        Writer.WriteLong( Flags );
        Writer.Seek( EndPos );
    },

    Read_FromBinary : function(Reader)
    {
        var Flags = Reader.GetLong();

        if ( Flags & 1 )
            this.Left = Reader.GetDouble();

        if ( Flags & 2 )
            this.Right = Reader.GetDouble();

        if ( Flags & 4 )
            this.FirstLine = Reader.GetDouble();
    }
};
CParaInd.prototype.Is_Empty = function()
{
	if (undefined !== this.Left
		|| undefined !== this.Right
		|| undefined !== this.FirstLine)
		return false;

	return true;
};
CParaInd.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CParaInd.prototype.Get_Diff = function(Ind)
{
    var DiffInd = new CParaInd();

    if (this.Left !== Ind.Left)
        DiffInd.Left = this.Left;

    if (this.Left !== Ind.Right)
        DiffInd.Right = this.Right;

    if (this.FirstLine !== Ind.FirstLine)
        DiffInd.FirstLine = this.FirstLine;

    return DiffInd;
};
CParaInd.prototype.IsEqual = function(oInd)
{
	return (IsEqualNullableFloatNumbers(this.Left, oInd.Left) && IsEqualNullableFloatNumbers(this.Right, oInd.Right) && IsEqualNullableFloatNumbers(this.FirstLine, oInd.FirstLine));
};

function CParaSpacing()
{
    this.Line              = undefined; // Расстояние между строками внутри абзаца
    this.LineRule          = undefined; // Тип расстрояния между строками
    this.Before            = undefined; // Дополнительное расстояние до абзаца
    this.BeforePct         = undefined; // Расстояние до абзаца в процентах от высоты строки
    this.BeforeAutoSpacing = undefined; // Использовать ли автоматический расчет расстояния до параграфа
    this.After             = undefined; // Дополнительное расстояние после абзаца
    this.AfterPct          = undefined; // Расстояние после абзаца в процентах от высоты строки
    this.AfterAutoSpacing  = undefined; // Использовать ли автоматический расчет расстояния после параграфа
}

CParaSpacing.prototype =
{
    Copy : function()
    {
        var Spacing = new CParaSpacing();
        Spacing.Line              = this.Line;
        Spacing.LineRule          = this.LineRule;
        Spacing.Before            = this.Before;
        Spacing.BeforeAutoSpacing = this.BeforeAutoSpacing;
        Spacing.After             = this.After;
        Spacing.AfterAutoSpacing  = this.AfterAutoSpacing;
        Spacing.BeforePct         = this.BeforePct;
        Spacing.AfterPct          = this.AfterPct;
        return Spacing;
    },

    Merge : function(Spacing)
    {
        if ( undefined != Spacing.Line )
            this.Line = Spacing.Line;

        if ( undefined != Spacing.LineRule )
            this.LineRule = Spacing.LineRule;

        if ( undefined != Spacing.Before )
            this.Before = Spacing.Before;

        if ( undefined != Spacing.BeforeAutoSpacing )
            this.BeforeAutoSpacing = Spacing.BeforeAutoSpacing;

        if ( undefined != Spacing.After )
            this.After = Spacing.After;

        if ( undefined != Spacing.AfterAutoSpacing )
            this.AfterAutoSpacing = Spacing.AfterAutoSpacing;

        if ( undefined != Spacing.BeforePct )
            this.BeforePct = Spacing.BeforePct;

        if ( undefined != Spacing.AfterPct )
            this.AfterPct = Spacing.AfterPct;
    },

	Is_Equal : function(Spacing)
	{
		return this.IsEqual(Spacing);
	},

    Set_FromObject : function(Spacing)
    {
        this.Line              = Spacing.Line;
        this.LineRule          = Spacing.LineRule;
        this.Before            = Spacing.Before;
        this.BeforeAutoSpacing = Spacing.BeforeAutoSpacing;
        this.After             = Spacing.After;
        this.AfterAutoSpacing  = Spacing.AfterAutoSpacing;
        this.BeforePct         = Spacing.BeforePct;
        this.AfterPct          = Spacing.AfterPct;
    },

    Write_ToBinary : function(Writer)
    {
        var StartPos = Writer.GetCurPosition();
        Writer.Skip(4);
        var Flags = 0;

        if ( undefined != this.Line )
        {
            Writer.WriteDouble( this.Line );
            Flags |= 1;
        }

        if ( undefined != this.LineRule )
        {
            Writer.WriteByte( this.LineRule );
            Flags |= 2;
        }

        if ( undefined != this.Before )
        {
            Writer.WriteDouble( this.Before );
            Flags |= 4;
        }

        if ( undefined != this.After )
        {
            Writer.WriteDouble( this.After );
            Flags |= 8;
        }

        if ( undefined != this.AfterAutoSpacing )
        {
            Writer.WriteBool( this.AfterAutoSpacing );
            Flags |= 16;
        }

        if ( undefined != this.BeforeAutoSpacing )
        {
            Writer.WriteBool( this.BeforeAutoSpacing );
            Flags |= 32;
        }

        if ( undefined != this.BeforePct )
        {
            Writer.WriteLong( this.BeforePct );
            Flags |= 64;
        }

        if ( undefined != this.AfterPct )
        {
            Writer.WriteLong( this.AfterPct );
            Flags |= 128;
        }

        var EndPos = Writer.GetCurPosition();
        Writer.Seek( StartPos );
        Writer.WriteLong( Flags );
        Writer.Seek( EndPos );
    },

    Read_FromBinary : function(Reader)
    {
        var Flags = Reader.GetLong();

        if ( Flags & 1 )
            this.Line = Reader.GetDouble();

        if ( Flags & 2 )
            this.LineRule = Reader.GetByte();

        if ( Flags & 4 )
            this.Before = Reader.GetDouble();

        if ( Flags & 8 )
            this.After = Reader.GetDouble();

        if ( Flags & 16 )
            this.AfterAutoSpacing = Reader.GetBool();

        if ( Flags & 32 )
            this.BeforeAutoSpacing = Reader.GetBool();

        if ( Flags & 64 )
            this.BeforePct = Reader.GetLong();

        if ( Flags & 128 )
            this.AfterPct = Reader.GetLong();
    }
};
CParaSpacing.prototype.Get_Diff = function(Spacing)
{
    var DiffSpacing = new CParaSpacing();

    if (this.Line !== Spacing.Line)
        DiffSpacing.Line = this.Line;

    if (this.LineRule !== Spacing.LineRule)
        DiffSpacing.LineRule = this.LineRule;

    if (this.Before !== Spacing.Before)
        DiffSpacing.Before = this.Before;

    if (this.BeforeAutoSpacing !== Spacing.BeforeAutoSpacing)
        DiffSpacing.BeforeAutoSpacing = this.BeforeAutoSpacing;

    if (this.After !== Spacing.After)
        DiffSpacing.After = this.After;

    if (this.AfterAutoSpacing !== Spacing.AfterAutoSpacing)
        DiffSpacing.AfterAutoSpacing = this.AfterAutoSpacing;

    if (this.BeforePct !== Spacing.BeforePct)
        DiffSpacing.BeforePct = this.BeforePct;

    if (this.AfterPct !== Spacing.AfterPct)
        DiffSpacing.AfterPct = this.AfterPct;


    return DiffSpacing;
};
CParaSpacing.prototype.Is_Empty = function()
{
	if (undefined !== this.Line
		|| undefined !== this.LineRule
		|| undefined !== this.Before
		|| undefined !== this.BeforeAutoSpacing
		|| undefined !== this.After
		|| undefined !== this.AfterAutoSpacing
		|| undefined !== this.BeforePct
		|| undefined !== this.AfterPct)
		return false;

	return true;
};
CParaSpacing.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CParaSpacing.prototype.IsEqual = function(oSpacing)
{
	return (this.Line === oSpacing.Line
		&& this.LineRule === oSpacing.LineRule
		&& IsEqualNullableFloatNumbers(this.Before, oSpacing.Before)
		&& IsEqualNullableFloatNumbers(this.After, oSpacing.After)
		&& IsEqualNullableFloatNumbers(this.AfterPct, oSpacing.AfterPct)
		&& IsEqualNullableFloatNumbers(this.BeforePct, oSpacing.BeforePct)
		&& this.BeforeAutoSpacing === oSpacing.BeforeAutoSpacing
		&& this.AfterAutoSpacing === oSpacing.AfterAutoSpacing);
};
CParaSpacing.prototype.SetLineTwips = function (val) {
	if (null !== val && undefined !== val) {
		if (val < 0) {
			val = Math.abs(val);
			this.LineRule = Asc.linerule_Exact;
		}
		if (Asc.linerule_Auto === this.LineRule) {
			this.Line = val / 240;
		} else {
			this.Line = g_dKoef_twips_to_mm * val;
		}
	}
};

function CNumPr(numId, iLvl)
{
    this.NumId = numId;
    this.Lvl   = iLvl;
}

CNumPr.prototype =
{
    Merge : function(NumPr)
    {
        if ( undefined != NumPr.NumId )
            this.NumId = NumPr.NumId;

        if ( undefined != NumPr.Lvl )
            this.Lvl = NumPr.Lvl;
    },

    Is_Equal : function(NumPr)
    {
    	return this.IsEqual(NumPr);
    },

    Set : function(NumId, Lvl)
    {
        this.NumId = NumId;
        this.Lvl   = Lvl;
    },

    Set_FromObject : function(NumPr)
    {
        this.NumId = NumPr.NumId;
        this.Lvl   = NumPr.Lvl;
    },

    Write_ToBinary : function(Writer)
    {
        // Bool   : Is NumId undefined
        // String : NumId
        // Bool   : IsLvl undefined
        // Byte   : Lvl

        if ( undefined === this.NumId )
            Writer.WriteBool( true );
        else
        {
            Writer.WriteBool( false );
            Writer.WriteString2( this.NumId );
        }

        if ( undefined === this.Lvl )
            Writer.WriteBool( true );
        else
        {
            Writer.WriteBool( false );
            Writer.WriteByte( this.Lvl );
        }
    },

    Read_FromBinary : function(Reader)
    {
        // Bool   : Is NumId undefined
        // String : NumId
        // Byte   : Lvl

        if ( true === Reader.GetBool() )
            this.NumId = undefined;
        else
            this.NumId = Reader.GetString2();

        if ( true === Reader.GetBool() )
            this.Lvl = undefined;
        else
            this.Lvl = Reader.GetByte();
    }
};
CNumPr.prototype.Copy = function()
{
	return new CNumPr(this.NumId, this.Lvl);
};
CNumPr.prototype.IsValid = function()
{
	return (!this.IsZero() && undefined !== this.NumId && null !== this.NumId);
};
/**
 * Нулевая нумерация используется для сброса нумерации в иерархии
 * @returns {boolean}
 */
CNumPr.prototype.IsZero = function()
{
	return (0 === this.NumId || "0" === this.NumId);
};
CNumPr.prototype.IsEqual = function(oNumPr)
{
	if (!oNumPr || this.NumId !== oNumPr.NumId)
		return false;

	if (0 === this.NumId)
		return true;

	return (this.NumId === oNumPr.NumId && this.Lvl === oNumPr.Lvl);
};

var wrap_Around    = 0x01;
var wrap_Auto      = 0x02;
var wrap_None      = 0x03;
var wrap_NotBeside = 0x04;
var wrap_Through   = 0x05;
var wrap_Tight     = 0x06;

function CFramePr()
{
    this.DropCap = undefined; // Является ли данный элемент буквицей
    this.H       = undefined;
    this.HAnchor = undefined;
    this.HRule   = undefined;
    this.HSpace  = undefined;
    this.Lines   = undefined;
    this.VAnchor = undefined;
    this.VSpace  = undefined;
    this.W       = undefined;
    this.Wrap    = undefined;
    this.X       = undefined;
    this.XAlign  = undefined;
    this.Y       = undefined;
    this.YAlign  = undefined;
}

CFramePr.prototype =
{
    Copy : function()
    {
        var FramePr = new CFramePr();

        FramePr.DropCap = this.DropCap;
        FramePr.H       = this.H;
        FramePr.HAnchor = this.HAnchor;
        FramePr.HRule   = this.HRule;
        FramePr.HSpace  = this.HSpace;
        FramePr.Lines   = this.Lines;
        FramePr.VAnchor = this.VAnchor;
        FramePr.VSpace  = this.VSpace;
        FramePr.W       = this.W;
        FramePr.Wrap    = this.Wrap;
        FramePr.X       = this.X;
        FramePr.XAlign  = this.XAlign;
        FramePr.Y       = this.Y;
        FramePr.YAlign  = this.YAlign;

        return FramePr;
    },

    Set_FromObject : function(FramePr)
    {
        this.DropCap = FramePr.DropCap;
        this.H       = FramePr.H;
        this.HAnchor = FramePr.HAnchor;
        this.HRule   = FramePr.HRule;
        this.HSpace  = FramePr.HSpace;
        this.Lines   = FramePr.Lines;
        this.VAnchor = FramePr.VAnchor;
        this.VSpace  = FramePr.VSpace;
        this.W       = FramePr.W;
        this.Wrap    = FramePr.Wrap;
        this.X       = FramePr.X;
        this.XAlign  = FramePr.XAlign;
        this.Y       = FramePr.Y;
        this.YAlign  = FramePr.YAlign;
    },

	Write_ToBinary : function(Writer)
	{
		var StartPos = Writer.GetCurPosition();
		Writer.Skip(4);
		var Flags = 0;

		if (undefined != this.DropCap)
		{
			Writer.WriteLong(this.DropCap);
			Flags |= 1;
		}

		if (undefined != this.H)
		{
			Writer.WriteDouble(this.H);
			Flags |= 2;
		}

		if (undefined != this.HAnchor)
		{
			Writer.WriteLong(this.HAnchor);
			Flags |= 4;
		}

		if (undefined != this.HRule)
		{
			Writer.WriteLong(this.HRule);
			Flags |= 8;
		}

		if (undefined != this.HSpace)
		{
			Writer.WriteDouble(this.HSpace);
			Flags |= 16;
		}

		if (undefined != this.Lines)
		{
			Writer.WriteLong(this.Lines);
			Flags |= 32;
		}

		if (undefined != this.VAnchor)
		{
			Writer.WriteLong(this.VAnchor);
			Flags |= 64;
		}

		if (undefined != this.VSpace)
		{
			Writer.WriteDouble(this.VSpace);
			Flags |= 128;
		}

		if (undefined != this.W)
		{
			Writer.WriteDouble(this.W);
			Flags |= 256;
		}

		if (undefined != this.X)
		{
			Writer.WriteDouble(this.X);
			Flags |= 512;
		}

		if (undefined != this.XAlign)
		{
			Writer.WriteLong(this.XAlign);
			Flags |= 1024;
		}

		if (undefined != this.Y)
		{
			Writer.WriteDouble(this.Y);
			Flags |= 2048;
		}

		if (undefined != this.YAlign)
		{
			Writer.WriteLong(this.YAlign);
			Flags |= 4096;
		}

		if (undefined !== this.Wrap)
		{
			Writer.WriteLong(this.Wrap);
			Flags |= 8192;
		}

		var EndPos = Writer.GetCurPosition();
		Writer.Seek(StartPos);
		Writer.WriteLong(Flags);
		Writer.Seek(EndPos);
	},

	Read_FromBinary : function(Reader)
	{
		var Flags = Reader.GetLong();

		if (Flags & 1)
			this.DropCap = Reader.GetLong();

		if (Flags & 2)
			this.H = Reader.GetDouble();

		if (Flags & 4)
			this.HAnchor = Reader.GetLong();

		if (Flags & 8)
			this.HRule = Reader.GetLong();

		if (Flags & 16)
			this.HSpace = Reader.GetDouble();

		if (Flags & 32)
			this.Lines = Reader.GetLong();

		if (Flags & 64)
			this.VAnchor = Reader.GetLong();

		if (Flags & 128)
			this.VSpace = Reader.GetDouble();

		if (Flags & 256)
			this.W = Reader.GetDouble();

		if (Flags & 512)
			this.X = Reader.GetDouble();

		if (Flags & 1024)
			this.XAlign = Reader.GetLong();

		if (Flags & 2048)
			this.Y = Reader.GetDouble();

		if (Flags & 4096)
			this.YAlign = Reader.GetLong();

		if (Flags & 8192)
			this.Wrap = Reader.GetLong();
	},

    Init_Default_DropCap : function(bInside)
    {
        this.DropCap = ( true === bInside ? c_oAscDropCap.Drop : c_oAscDropCap.Margin );
        this.Lines   = 3;
        this.Wrap    = wrap_Around;
        this.VAnchor = Asc.c_oAscVAnchor.Text;
        this.HAnchor = ( true === bInside ? Asc.c_oAscHAnchor.Text : Asc.c_oAscHAnchor.Page );

        this.X       = undefined;
        this.XAlign  = undefined;
        this.Y       = undefined;
        this.YAlign  = undefined;

        this.H       = undefined;
        this.W       = undefined;
        this.HRule   = undefined;
    },

    Get_W : function()
    {
        return this.W;
    },

    Get_H : function()
    {
        return this.H;
    },

    Is_DropCap : function()
    {
        if ( c_oAscDropCap.Margin === this.DropCap || c_oAscDropCap.Drop === this.DropCap )
            return true;

        return false;
    }
};
CFramePr.prototype.GetW = function()
{
	return this.W;
};
CFramePr.prototype.IsEqual = function(oFramePr)
{
	if (!oFramePr)
		return false;

	return (this.DropCap === oFramePr.DropCap
		&& IsEqualNullableFloatNumbers(this.H, oFramePr.H)
		&& this.HAnchor === oFramePr.HAnchor
		&& this.HRule === oFramePr.HRule
		&& this.HSpace === oFramePr.HSpace
		&& this.Lines === oFramePr.Lines
		&& this.VAnchor === oFramePr.VAnchor
		&& this.VSpace === oFramePr.VSpace
		&& IsEqualNullableFloatNumbers(this.W, oFramePr.W)
		&& this.Wrap === oFramePr.Wrap
		&& IsEqualNullableFloatNumbers(this.X, oFramePr.X)
		&& this.XAlign === oFramePr.XAlign
		&& IsEqualNullableFloatNumbers(this.Y, oFramePr.Y)
		&& this.YAlign === oFramePr.YAlign);
};
CFramePr.prototype.Compare = function(FramePr)
{
	return this.IsEqual(FramePr);
};
CFramePr.prototype.Is_Equal = function(FramePr)
{
	return this.IsEqual(FramePr);
};
CFramePr.prototype.Merge = function(oFramePr)
{
	// Некоторые свойства завязаны друг на друга. Если одно присутствует, то второе берется из текущего объекта
	// (даже если оно не задано), а не по иерархии выше. H / HRule, X / XAlign, Y / YAlign

	if (null !== oFramePr.DropCap && undefined !== oFramePr.DropCap)
		this.DropCap = oFramePr.DropCap;

	if (null !== oFramePr.H && undefined !== oFramePr.H)
	{
		if (null === oFramePr.HRule || undefined === oFramePr.HRule)
			this.HRule = undefined;

		this.H = oFramePr.H;
	}

	if (null !== oFramePr.HAnchor && undefined !== oFramePr.HAnchor)
		this.HAnchor = oFramePr.HAnchor;

	if (null !== oFramePr.HRule && undefined !== oFramePr.HRule)
	{
		if (null === oFramePr.H || undefined === oFramePr.H)
			this.H = undefined;

		this.HRule = oFramePr.HRule;
	}

	if (null !== oFramePr.HSpace && undefined !== oFramePr.HSpace)
		this.HSpace = oFramePr.HSpace;

	if (null !== oFramePr.Lines && undefined !== oFramePr.Lines)
		this.Lines = oFramePr.Lines;

	if (null !== oFramePr.VAnchor && undefined !== oFramePr.VAnchor)
		this.VAnchor = oFramePr.VAnchor;

	if (null !== oFramePr.VSpace && undefined !== oFramePr.VSpace)
		this.VSpace = oFramePr.VSpace;

	if (null !== oFramePr.W && undefined !== oFramePr.W)
		this.W = oFramePr.W;

	if (null !== oFramePr.Wrap && undefined !== oFramePr.Wrap)
		this.Wrap = oFramePr.Wrap;

	if (null !== oFramePr.X && undefined !== oFramePr.X)
	{
		if (null === oFramePr.XAlign || undefined === oFramePr.XAlign)
			this.XAlign = undefined;

		this.X = oFramePr.X;
	}

	if (null !== oFramePr.XAlign && undefined !== oFramePr.XAlign)
	{
		if (null === oFramePr.X || undefined === oFramePr.X)
			this.X = undefined;

		this.XAlign = oFramePr.XAlign;
	}

	if (null !== oFramePr.Y && undefined !== oFramePr.Y)
	{
		if (null === oFramePr.YAlign || undefined === oFramePr.YAlign)
			this.YAlign = undefined;

		this.Y = oFramePr.Y;
	}

	if (null !== oFramePr.YAlign && undefined !== oFramePr.YAlign)
	{
		if (null === oFramePr.Y || undefined === oFramePr.Y)
			this.Y = undefined;

		this.YAlign = oFramePr.YAlign;
	}
};
/**
 * Является ли рамка инлайновой
 * @returns {boolean}
 */
CFramePr.prototype.IsInline = function()
{
	return ((undefined === this.XAlign || c_oAscXAlign.Left === this.XAlign)
		&& (undefined === this.YAlign || Asc.c_oAscYAlign.Inline === this.YAlign)
		&& (undefined === this.HAnchor || Asc.c_oAscHAnchor.Margin === this.HAnchor)
		&& (undefined === this.VAnchor || Asc.c_oAscVAnchor.Margin === this.VAnchor)
		&& (undefined === this.X || 0 === AscCommon.MMToTwips(this.X))
		&& (undefined === this.Y || 0 === AscCommon.MMToTwips(this.Y))
		&& (undefined === this.W || 0 === AscCommon.MMToTwips(this.W))
		&& (undefined === this.W || 0 === AscCommon.MMToTwips(this.W)));
};

function CCalculatedFrame(FramePr, L, T, W, H, L2, T2, W2, H2, PageIndex, Index, FlowCount)
{
	this.FramePr    = FramePr;
	this.L          = undefined !== L ? L : 0; // Внутренний рект, по которому идет рассчет
	this.T          = undefined !== T ? T : 0;
	this.W          = undefined !== W ? W : 0;
	this.H          = undefined !== H ? H : 0;
	this.L2         = undefined !== L2 ? L2 : 0; // Внешний рект, с учетом границ
	this.T2         = undefined !== T2 ? T2 : 0;
	this.W2         = undefined !== W2 ? W2 : 0;
	this.H2         = undefined !== H2 ? H2 : 0;
	this.PageIndex  = undefined !== PageIndex ? PageIndex : 0;
	this.StartIndex = undefined !== Index ? Index : 0;
	this.FlowCount  = undefined !== FlowCount ? FlowCount : 0;

	this.Paragraphs = [];
}
CCalculatedFrame.prototype.AddParagraph = function(oPara)
{
	this.Paragraphs.push(oPara);
};
CCalculatedFrame.prototype.GetParagraphs = function()
{
	return this.Paragraphs;
};
CCalculatedFrame.prototype.GetFramePr = function()
{
	return this.FramePr;
};

function CParaPr()
{
	this.Bidi              = undefined;
	this.ContextualSpacing = undefined;          // Удалять ли интервал между параграфами одинакового стиля
	this.Ind               = new CParaInd();     // Отступы
	this.Jc                = undefined;          // Прилегание параграфа
	this.KeepLines         = undefined;          // Неразрывный параграф
	this.KeepNext          = undefined;          // Не разъединять со следующим параграфом
	this.PageBreakBefore   = undefined;          // Разрыв страницы перед параграфом
	this.Spacing           = new CParaSpacing(); // Расстояния между строками внутри параграфа и между параграфами
	this.Shd               = undefined;          // Заливка параграфа
	this.Brd               = {
		First   : undefined,            // Является ли данный параграф первым в группе параграфов с одинаковыми краями и настройками границ
		Last    : undefined,            // Является ли данный параграф последним в группе параграфов с одинаковыми краями и настройками границ
		Between : undefined,
		Bottom  : undefined,
		Left    : undefined,
		Right   : undefined,
		Top     : undefined
	};
	this.WidowControl      = undefined; // Запрет висячих строк
	this.Tabs              = undefined; // Заданные табы
	this.NumPr             = undefined; // Нумерация
	this.PStyle            = undefined; // Стиль параграфа
	this.FramePr           = undefined;
	this.OutlineLvl        = undefined; // Для TableOfContents
	this.DefaultRunPr      = undefined;
	this.Bullet            = undefined;
	this.Lvl               = undefined;
	this.DefaultTab        = undefined;
	this.LnSpcReduction    = undefined;
	this.PrChange          = undefined;
	this.ReviewInfo        = undefined;

	this.SuppressLineNumbers = undefined;
}

CParaPr.fromObject = function(obj)
{
	let paraPr = new CParaPr();
	paraPr.Set_FromObject(obj);
	return paraPr;
};
CParaPr.prototype.Copy = function(bCopyPrChange, oPr)
{
	var ParaPr = new CParaPr();

	ParaPr.Bidi = this.Bidi;
	ParaPr.ContextualSpacing = this.ContextualSpacing;

	if (undefined != this.Ind)
		ParaPr.Ind = this.Ind.Copy();

	ParaPr.Jc              = this.Jc;
	ParaPr.KeepLines       = this.KeepLines;
	ParaPr.KeepNext        = this.KeepNext;
	ParaPr.PageBreakBefore = this.PageBreakBefore;

	if (undefined != this.Spacing)
		ParaPr.Spacing = this.Spacing.Copy();

	if (undefined != this.Shd)
		ParaPr.Shd = this.Shd.Copy();

	ParaPr.Brd.First = this.Brd.First;
	ParaPr.Brd.Last  = this.Brd.Last;

	if (undefined != this.Brd.Between)
		ParaPr.Brd.Between = this.Brd.Between.Copy();

	if (undefined != this.Brd.Bottom)
		ParaPr.Brd.Bottom = this.Brd.Bottom.Copy();

	if (undefined != this.Brd.Left)
		ParaPr.Brd.Left = this.Brd.Left.Copy();

	if (undefined != this.Brd.Right)
		ParaPr.Brd.Right = this.Brd.Right.Copy();

	if (undefined != this.Brd.Top)
		ParaPr.Brd.Top = this.Brd.Top.Copy();

	ParaPr.WidowControl = this.WidowControl;

	if (undefined != this.Tabs)
		ParaPr.Tabs = this.Tabs.Copy();

	if (undefined != this.NumPr)
	{
		ParaPr.NumPr = this.NumPr.Copy();
		if(oPr && oPr.Comparison)
		{
			ParaPr.NumPr.NumId = oPr.Comparison.getCopyNumId(ParaPr.NumPr.NumId);
		}
	}

	if (undefined != this.PStyle)
	{
		ParaPr.PStyle = this.PStyle;
		if(oPr && oPr.Comparison)
		{
			ParaPr.PStyle = oPr.Comparison.copyStyleById(ParaPr.PStyle);
		}
	}

	if (undefined != this.FramePr)
		ParaPr.FramePr = this.FramePr.Copy();
	else
		ParaPr.FramePr = undefined;

	if (undefined != this.DefaultRunPr)
		ParaPr.DefaultRunPr = this.DefaultRunPr.Copy();

	if (undefined != this.Bullet)
		ParaPr.Bullet = this.Bullet.createDuplicate();

	if (undefined != this.Lvl)
		ParaPr.Lvl = this.Lvl;

	if (undefined != this.DefaultTab)
		ParaPr.DefaultTab = this.DefaultTab;

	if (undefined != this.LnSpcReduction)
		ParaPr.LnSpcReduction = this.LnSpcReduction;

	if (true === bCopyPrChange && undefined !== this.PrChange)
	{
		ParaPr.PrChange   = this.PrChange.Copy();
		ParaPr.ReviewInfo = this.ReviewInfo.Copy();
	}

	if (undefined !== this.OutlineLvl)
		ParaPr.OutlineLvl = this.OutlineLvl;

	if (undefined !== this.OutlineLvlStyle)
		ParaPr.OutlineLvlStyle = this.OutlineLvlStyle;

	if (undefined !== this.Locked)
		ParaPr.Locked = this.Locked;

	if (undefined !== this.SuppressLineNumbers)
		ParaPr.SuppressLineNumbers = this.SuppressLineNumbers;

	return ParaPr;
};
CParaPr.prototype.createDuplicateForSmartArt = function (bCopyPrChange, oPr) {
	var ParaPr = new CParaPr();
	ParaPr.Jc              = this.Jc;

	if (undefined != this.Spacing)
		ParaPr.Spacing = this.Spacing.Copy();

	// if (undefined != this.Ind) // TODO: apply only changed ind
	// 	ParaPr.Ind = this.Ind.Copy();

	if (undefined != this.Tabs)
		ParaPr.Tabs = this.Tabs.Copy();

	if (undefined != this.DefaultTab)
		ParaPr.DefaultTab = this.DefaultTab;
	return ParaPr;
};
CParaPr.prototype.Merge = function(ParaPr)
{
	if (undefined !== ParaPr.Bidi)
		this.Bidi = ParaPr.Bidi;
	
	if (undefined != ParaPr.ContextualSpacing)
		this.ContextualSpacing = ParaPr.ContextualSpacing;

	if (undefined != ParaPr.Ind)
		this.Ind.Merge(ParaPr.Ind);

	if (undefined != ParaPr.Jc)
		this.Jc = ParaPr.Jc;

	if (undefined != ParaPr.KeepLines)
		this.KeepLines = ParaPr.KeepLines;

	if (undefined != ParaPr.KeepNext)
		this.KeepNext = ParaPr.KeepNext;

	if (undefined != ParaPr.PageBreakBefore)
		this.PageBreakBefore = ParaPr.PageBreakBefore;

	if (undefined != ParaPr.Spacing)
		this.Spacing.Merge(ParaPr.Spacing);

	if (undefined != ParaPr.Shd && (!this.Shd || !ParaPr.Shd.IsNil()))
		this.Shd = ParaPr.Shd.Copy();
	
	if (ParaPr.Brd)
	{
		if (undefined != ParaPr.Brd.First)
			this.Brd.First = ParaPr.Brd.First;
		
		if (undefined != ParaPr.Brd.Last)
			this.Brd.Last = ParaPr.Brd.Last;
		
		if (undefined != ParaPr.Brd.Between)
			this.Brd.Between = ParaPr.Brd.Between.Copy();
		
		if (undefined != ParaPr.Brd.Bottom)
			this.Brd.Bottom = ParaPr.Brd.Bottom.Copy();
		
		if (undefined != ParaPr.Brd.Left)
			this.Brd.Left = ParaPr.Brd.Left.Copy();
		
		if (undefined != ParaPr.Brd.Right)
			this.Brd.Right = ParaPr.Brd.Right.Copy();
		
		if (undefined != ParaPr.Brd.Top)
			this.Brd.Top = ParaPr.Brd.Top.Copy();
	}

	if (undefined != ParaPr.WidowControl)
		this.WidowControl = ParaPr.WidowControl;

	if (undefined != ParaPr.Tabs)
	{
		if (undefined === this.Tabs)
			this.Tabs = ParaPr.Tabs.Copy();
		else
			this.Tabs.Merge(ParaPr.Tabs);
	}

	if (undefined != ParaPr.NumPr)
	{
		if (undefined === this.NumPr)
			this.NumPr = ParaPr.NumPr.Copy();
		else
			this.NumPr.Merge(ParaPr.NumPr);

		if (undefined != this.NumPr && this.NumPr.Lvl > 8)
			this.NumPr = undefined;
	}

	if (undefined != ParaPr.PStyle)
		this.PStyle = ParaPr.PStyle;

	if (null !== ParaPr.FramePr && undefined !== ParaPr.FramePr)
	{
		if (!this.FramePr)
			this.FramePr = ParaPr.FramePr.Copy();
		else
			this.FramePr.Merge(ParaPr.FramePr);
	}

	if (undefined != ParaPr.DefaultRunPr)
	{
		if (undefined == this.DefaultRunPr)
			this.DefaultRunPr = new CTextPr();
		this.DefaultRunPr.Merge(ParaPr.DefaultRunPr);
	}

	if (undefined != ParaPr.Bullet)
	{
		if (ParaPr.Bullet.isBullet())
		{
			if (!this.Bullet)
			{
				this.Bullet = new AscFormat.CBullet();
			}
			var PrBullet = ParaPr.Bullet;
			if (PrBullet.bulletColor)
			{
				this.Bullet.bulletColor = PrBullet.bulletColor.createDuplicate();
			}
			if (PrBullet.bulletSize)
			{
				this.Bullet.bulletSize = PrBullet.bulletSize.createDuplicate();
			}
			if (PrBullet.bulletTypeface)
			{
				this.Bullet.bulletTypeface = PrBullet.bulletTypeface.createDuplicate();
			}
			if (PrBullet.bulletType)
			{
				this.Bullet.bulletType = PrBullet.bulletType.createDuplicate();
			}
			this.Bullet.Bullet = PrBullet.Bullet;

		}
		else
		{
			if (this.Bullet && this.Bullet.isBullet())
			{
				if (ParaPr.Bullet.bulletColor)
				{
					this.Bullet.bulletColor = ParaPr.Bullet.bulletColor.createDuplicate();
				}
				if (ParaPr.Bullet.bulletSize)
				{
					this.Bullet.bulletSize = ParaPr.Bullet.bulletSize.createDuplicate();
				}
				if (ParaPr.Bullet.bulletTypeface)
				{
					this.Bullet.bulletTypeface = ParaPr.Bullet.bulletTypeface.createDuplicate();
				}
			}
		}
	}

	if (undefined != ParaPr.Lvl)
		this.Lvl = ParaPr.Lvl;

	if (undefined != ParaPr.DefaultTab)
		this.DefaultTab = ParaPr.DefaultTab;

	if (undefined != ParaPr.LnSpcReduction)
		this.LnSpcReduction = ParaPr.LnSpcReduction;

	if (undefined !== ParaPr.OutlineLvl)
		this.OutlineLvl = ParaPr.OutlineLvl;

	if (undefined !== ParaPr.SuppressLineNumbers)
		this.SuppressLineNumbers = ParaPr.SuppressLineNumbers;
};
CParaPr.prototype.InitDefault = function(nCompatibilityMode)
{
	this.Bidi                      = false;
	this.ContextualSpacing         = false;
	this.Ind                       = new CParaInd();
	this.Ind.Left                  = 0;
	this.Ind.Right                 = 0;
	this.Ind.FirstLine             = 0;
	this.Jc                        = align_Left;
	this.KeepLines                 = false;
	this.KeepNext                  = false;
	this.PageBreakBefore           = false;
	this.Spacing                   = new CParaSpacing();
	this.Spacing.Line              = 1;
	this.Spacing.LineRule          = linerule_Auto;
	this.Spacing.Before            = 0;
	this.Spacing.BeforeAutoSpacing = false;
	this.Spacing.After             = 0;
	this.Spacing.AfterAutoSpacing  = false;
	this.Shd                       = new CDocumentShd();
	this.Brd.First                 = true;
	this.Brd.Last                  = true;
	this.Brd.Between               = new CDocumentBorder();
	this.Brd.Bottom                = new CDocumentBorder();
	this.Brd.Left                  = new CDocumentBorder();
	this.Brd.Right                 = new CDocumentBorder();
	this.Brd.Top                   = new CDocumentBorder();
	this.WidowControl              = true;
	this.Tabs                      = new CParaTabs();
	this.NumPr                     = undefined;
	this.PStyle                    = undefined;
	this.FramePr                   = undefined;
	this.OutlineLvl                = undefined;
	this.SuppressLineNumbers       = false;

	this.DefaultRunPr   = undefined;
	this.Bullet         = undefined;
	this.DefaultTab     = undefined;
	this.LnSpcReduction = undefined;
	this.PrChange       = undefined;
	this.ReviewInfo     = undefined
};
CParaPr.prototype.Set_FromObject = function(ParaPr)
{
	this.Bidi              = ParaPr.Bidi;
	this.ContextualSpacing = ParaPr.ContextualSpacing;

	this.Ind = new CParaInd();
	if (undefined != ParaPr.Ind)
		this.Ind.Set_FromObject(ParaPr.Ind);

	this.Jc              = ParaPr.Jc;
	this.KeepLines       = ParaPr.KeepLines;
	this.KeepNext        = ParaPr.KeepNext;
	this.PageBreakBefore = ParaPr.PageBreakBefore;

	this.Spacing = new CParaSpacing();
	if (undefined != ParaPr.Spacing)
		this.Spacing.Set_FromObject(ParaPr.Spacing);

	if (undefined != ParaPr.Shd)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Set_FromObject(ParaPr.Shd);
	}
	else
		this.Shd = undefined;

	if (undefined != ParaPr.Brd)
	{
		if (undefined != ParaPr.Brd.Between)
		{
			this.Brd.Between = new CDocumentBorder();
			this.Brd.Between.Set_FromObject(ParaPr.Brd.Between);
		}
		else
			this.Brd.Between = undefined;

		if (undefined != ParaPr.Brd.Bottom)
		{
			this.Brd.Bottom = new CDocumentBorder();
			this.Brd.Bottom.Set_FromObject(ParaPr.Brd.Bottom);
		}
		else
			this.Brd.Bottom = undefined;

		if (undefined != ParaPr.Brd.Left)
		{
			this.Brd.Left = new CDocumentBorder();
			this.Brd.Left.Set_FromObject(ParaPr.Brd.Left);
		}
		else
			this.Brd.Left = undefined;

		if (undefined != ParaPr.Brd.Right)
		{
			this.Brd.Right = new CDocumentBorder();
			this.Brd.Right.Set_FromObject(ParaPr.Brd.Right);
		}
		else
			this.Brd.Right = undefined;

		if (undefined != ParaPr.Brd.Top)
		{
			this.Brd.Top = new CDocumentBorder();
			this.Brd.Top.Set_FromObject(ParaPr.Brd.Top);
		}
		else
			this.Brd.Top = undefined;
	}
	else
	{
		this.Brd.Between = undefined;
		this.Brd.Bottom  = undefined;
		this.Brd.Left    = undefined;
		this.Brd.Right   = undefined;
		this.Brd.Top     = undefined;
	}

	this.WidowControl = ParaPr.WidowControl;

	if (undefined != ParaPr.Tabs)
	{
		this.Tabs = new CParaTabs();
		this.Tabs.Set_FromObject(ParaPr.Tabs.Tabs);
	}
	else
		this.Tabs = undefined;

	if (undefined != ParaPr.NumPr)
	{
		this.NumPr = new AscWord.NumPr();
		this.NumPr.Set_FromObject(ParaPr.NumPr);
	}
	else
		this.NumPr = undefined;

	if (undefined != ParaPr.FramePr)
	{
		this.FramePr = new CFramePr();
		this.FramePr.Set_FromObject(ParaPr.FramePr);
	}
	else
		this.FramePr = undefined;

	if (undefined != ParaPr.DefaultRunPr)
	{
		this.DefaultRunPr = new CTextPr();
		this.DefaultRunPr.Set_FromObject(ParaPr.DefaultRunPr);
	}

	if (undefined != ParaPr.Bullet)
	{
		this.Bullet = new AscFormat.CBullet();
		this.Bullet.Set_FromObject(ParaPr.Bullet);
	}

	if (undefined != ParaPr.DefaultTab)
	{
		this.DefaultTab = ParaPr.DefaultTab;
	}

	if (undefined !== ParaPr.OutlineLvl)
		this.OutlineLvl = ParaPr.OutlineLvl;

	if (undefined !== ParaPr.SuppressLineNumbers)
		this.SuppressLineNumbers = ParaPr.SuppressLineNumbers;
};
CParaPr.prototype.SetFromObject = function(oPr)
{
	return this.Set_FromObject(oPr);
};
CParaPr.prototype.Compare = function(ParaPr)
{
	// При сравнении добавляем 1 элемент Locked
	var Result_ParaPr    = new CParaPr();
	Result_ParaPr.Locked = false;
	
	if (ParaPr.Bidi === this.Bidi)
		Result_ParaPr.Bidi = ParaPr.Bidi;

	if (ParaPr.ContextualSpacing === this.ContextualSpacing)
		Result_ParaPr.ContextualSpacing = ParaPr.ContextualSpacing;

	Result_ParaPr.Ind = new CParaInd();
	if (undefined != ParaPr.Ind && undefined != this.Ind)
	{
		if (undefined != ParaPr.Ind.Left && undefined != this.Ind.Left && Math.abs(ParaPr.Ind.Left - this.Ind.Left) < 0.001)
			Result_ParaPr.Ind.Left = ParaPr.Ind.Left;

		if (undefined != ParaPr.Ind.Right && undefined != this.Ind.Right && Math.abs(ParaPr.Ind.Right - this.Ind.Right) < 0.001)
			Result_ParaPr.Ind.Right = ParaPr.Ind.Right;

		if (undefined != ParaPr.Ind.FirstLine && undefined != this.Ind.FirstLine && Math.abs(ParaPr.Ind.FirstLine - this.Ind.FirstLine) < 0.001)
			Result_ParaPr.Ind.FirstLine = ParaPr.Ind.FirstLine;
	}

	if (ParaPr.Jc === this.Jc)
		Result_ParaPr.Jc = ParaPr.Jc;

	if (ParaPr.KeepLines === this.KeepLines)
		Result_ParaPr.KeepLines = ParaPr.KeepLines;

	if (ParaPr.KeepNext === this.KeepNext)
		Result_ParaPr.KeepNext = ParaPr.KeepNext;

	if (ParaPr.PageBreakBefore === this.PageBreakBefore)
		Result_ParaPr.PageBreakBefore = ParaPr.PageBreakBefore;

	Result_ParaPr.Spacing = new CParaSpacing();
	if (undefined != this.Spacing && undefined != ParaPr.Spacing)
	{
		if (undefined != this.Spacing.After && undefined != ParaPr.Spacing.After && Math.abs(this.Spacing.After - ParaPr.Spacing.After) < 0.001)
			Result_ParaPr.Spacing.After = ParaPr.Spacing.After;

		if (this.Spacing.AfterAutoSpacing === ParaPr.Spacing.AfterAutoSpacing)
			Result_ParaPr.Spacing.AfterAutoSpacing = ParaPr.Spacing.AfterAutoSpacing;

		if (undefined != this.Spacing.Before && undefined != ParaPr.Spacing.Before && Math.abs(this.Spacing.Before - ParaPr.Spacing.Before) < 0.001)
			Result_ParaPr.Spacing.Before = ParaPr.Spacing.Before;

		if (this.Spacing.BeforeAutoSpacing === ParaPr.Spacing.BeforeAutoSpacing)
			Result_ParaPr.Spacing.BeforeAutoSpacing = ParaPr.Spacing.BeforeAutoSpacing;

		if (undefined != this.Spacing.Line && undefined != ParaPr.Spacing.Line && Math.abs(this.Spacing.Line - ParaPr.Spacing.Line) < 0.001)
			Result_ParaPr.Spacing.Line = ParaPr.Spacing.Line;

		if (this.Spacing.LineRule === ParaPr.Spacing.LineRule)
			Result_ParaPr.Spacing.LineRule = ParaPr.Spacing.LineRule;
	}

	if (undefined != this.Shd && undefined != ParaPr.Shd && true === this.Shd.Compare(ParaPr.Shd))
		Result_ParaPr.Shd = ParaPr.Shd.Copy();

	if (undefined != this.Brd.Between && undefined != ParaPr.Brd.Between && true === this.Brd.Between.Compare(ParaPr.Brd.Between))
		Result_ParaPr.Brd.Between = ParaPr.Brd.Between.Copy();

	if (undefined != this.Brd.Bottom && undefined != ParaPr.Brd.Bottom && true === this.Brd.Bottom.Compare(ParaPr.Brd.Bottom))
		Result_ParaPr.Brd.Bottom = ParaPr.Brd.Bottom.Copy();

	if (undefined != this.Brd.Left && undefined != ParaPr.Brd.Left && true === this.Brd.Left.Compare(ParaPr.Brd.Left))
		Result_ParaPr.Brd.Left = ParaPr.Brd.Left.Copy();

	if (undefined != this.Brd.Right && undefined != ParaPr.Brd.Right && true === this.Brd.Right.Compare(ParaPr.Brd.Right))
		Result_ParaPr.Brd.Right = ParaPr.Brd.Right.Copy();

	if (undefined != this.Brd.Top && undefined != ParaPr.Brd.Top && true === this.Brd.Top.Compare(ParaPr.Brd.Top))
		Result_ParaPr.Brd.Top = ParaPr.Brd.Top.Copy();

	if (ParaPr.WidowControl === this.WidowControl)
		Result_ParaPr.WidowControl = ParaPr.WidowControl;

	// PStyle
	if (undefined != this.PStyle && undefined != ParaPr.PStyle && this.PStyle === ParaPr.PStyle)
		Result_ParaPr.PStyle = ParaPr.PStyle;

	// NumPr
	if (undefined != this.NumPr && undefined != ParaPr.NumPr && this.NumPr.NumId === ParaPr.NumPr.NumId)
	{
		Result_ParaPr.NumPr       = new AscWord.NumPr();
		Result_ParaPr.NumPr.NumId = ParaPr.NumPr.NumId;
		Result_ParaPr.NumPr.Lvl   = Math.max(this.NumPr.Lvl, ParaPr.NumPr.Lvl);
	}

	// Locked
	if (undefined != this.Locked && undefined != ParaPr.Locked)
	{
		if (this.Locked != ParaPr.Locked)
			Result_ParaPr.Locked = true;
		else
			Result_ParaPr.Locked = ParaPr.Locked;
	}

	// FramePr
	if (undefined != this.FramePr && undefined != ParaPr.FramePr && true === this.FramePr.Compare(ParaPr.FramePr))
		Result_ParaPr.FramePr = this.FramePr;

	if (undefined != this.Bullet && undefined != ParaPr.Bullet)
		Result_ParaPr.Bullet = AscFormat.CompareBullets(ParaPr.Bullet, this.Bullet);

	if (undefined != this.DefaultRunPr && undefined != ParaPr.DefaultRunPr)
		Result_ParaPr.DefaultRunPr = this.DefaultRunPr.Compare(ParaPr.DefaultRunPr);

	if (undefined != this.Lvl && undefined != ParaPr.Lvl && ParaPr.Lvl === this.Lvl)
		Result_ParaPr.Lvl = this.Lvl;


	if (undefined != this.DefaultTab && undefined != ParaPr.DefaultTab && ParaPr.DefaultTab === this.DefaultTab)
		Result_ParaPr.DefaultTab = this.DefaultTab;

	if (undefined !== this.Tabs && undefined !== ParaPr.Tabs && this.Tabs.Is_Equal(ParaPr.Tabs))
		Result_ParaPr.Tabs = this.Tabs.Copy();

	if (this.OutlineLvl === ParaPr.OutlineLvl)
		Result_ParaPr.OutlineLvl = this.OutlineLvl;

	if (this.OutlineLvlStyle || ParaPr.OutlineLvlStyle)
		Result_ParaPr.OutlineLvlStyle = true;

	if (this.SuppressLineNumbers === ParaPr.SuppressLineNumbers)
		Result_ParaPr.SuppressLineNumbers = this.SuppressLineNumbers;

	return Result_ParaPr;
};
CParaPr.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);
	var Flags = 0;

	if (undefined != this.ContextualSpacing)
	{
		Writer.WriteBool(this.ContextualSpacing);
		Flags |= 1;
	}

	if (undefined != this.Ind)
	{
		this.Ind.Write_ToBinary(Writer);
		Flags |= 2;
	}

	if (undefined != this.Jc)
	{
		Writer.WriteByte(this.Jc);
		Flags |= 4;
	}

	if (undefined != this.KeepLines)
	{
		Writer.WriteBool(this.KeepLines);
		Flags |= 8;
	}

	if (undefined != this.KeepNext)
	{
		Writer.WriteBool(this.KeepNext);
		Flags |= 16;
	}

	if (undefined != this.PageBreakBefore)
	{
		Writer.WriteBool(this.PageBreakBefore);
		Flags |= 32;
	}

	if (undefined != this.Spacing)
	{
		this.Spacing.Write_ToBinary(Writer);
		Flags |= 64;
	}

	if (undefined != this.Shd)
	{
		this.Shd.Write_ToBinary(Writer);
		Flags |= 128;
	}

	if (undefined != this.Brd.Between)
	{
		this.Brd.Between.Write_ToBinary(Writer);
		Flags |= 256;
	}

	if (undefined != this.Brd.Bottom)
	{
		this.Brd.Bottom.Write_ToBinary(Writer);
		Flags |= 512;
	}

	if (undefined != this.Brd.Left)
	{
		this.Brd.Left.Write_ToBinary(Writer);
		Flags |= 1024;
	}

	if (undefined != this.Brd.Right)
	{
		this.Brd.Right.Write_ToBinary(Writer);
		Flags |= 2048;
	}

	if (undefined != this.Brd.Top)
	{
		this.Brd.Top.Write_ToBinary(Writer);
		Flags |= 4096;
	}

	if (undefined != this.WidowControl)
	{
		Writer.WriteBool(this.WidowControl);
		Flags |= 8192;
	}

	if (undefined != this.Tabs)
	{
		this.Tabs.Write_ToBinary(Writer);
		Flags |= 16384;
	}

	if (undefined != this.NumPr)
	{
		this.NumPr.Write_ToBinary(Writer);
		Flags |= 32768;
	}

	if (undefined != this.PStyle)
	{
		Writer.WriteString2(this.PStyle);
		Flags |= 65536;
	}

	if (undefined != this.FramePr)
	{
		this.FramePr.Write_ToBinary(Writer);
		Flags |= 131072;
	}

	if (undefined != this.DefaultRunPr)
	{
		this.DefaultRunPr.Write_ToBinary(Writer);
		Flags |= 262144;
	}

	if (undefined != this.Bullet)
	{
		this.Bullet.Write_ToBinary(Writer);
		Flags |= 524288;
	}

	if (undefined != this.Lvl)
	{
		Writer.WriteByte(this.Lvl);
		Flags |= 1048576;
	}

	if (undefined != this.DefaultTab)
	{
		Writer.WriteDouble(this.DefaultTab);
		Flags |= 2097152;
	}

	if (undefined !== this.OutlineLvl)
	{
		Writer.WriteByte(this.OutlineLvl);
		Flags |= 4194304;
	}

	if (undefined !== this.PrChange)
	{
		this.PrChange.WriteToBinary(Writer);
		this.ReviewInfo.WriteToBinary(Writer);
		Flags |= 8388608;
	}

	if (undefined !== this.SuppressLineNumbers)
	{
		Writer.WriteBool(this.SuppressLineNumbers);
		Flags |= 16777216;
	}
	
	if (undefined !== this.Bidi)
	{
		Writer.WriteBool(this.Bidi);
		Flags |= (1 << 25);
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CParaPr.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	if (Flags & 1)
		this.ContextualSpacing = Reader.GetBool();

	if (Flags & 2)
	{
		this.Ind = new CParaInd();
		this.Ind.Read_FromBinary(Reader);
	}

	if (Flags & 4)
		this.Jc = Reader.GetByte();

	if (Flags & 8)
		this.KeepLines = Reader.GetBool();

	if (Flags & 16)
		this.KeepNext = Reader.GetBool();

	if (Flags & 32)
		this.PageBreakBefore = Reader.GetBool();

	if (Flags & 64)
	{
		this.Spacing = new CParaSpacing();
		this.Spacing.Read_FromBinary(Reader);
	}

	if (Flags & 128)
	{
		this.Shd = new CDocumentShd();
		this.Shd.Read_FromBinary(Reader);
	}

	if (Flags & 256)
	{
		this.Brd.Between = new CDocumentBorder();
		this.Brd.Between.Read_FromBinary(Reader);
	}

	if (Flags & 512)
	{
		this.Brd.Bottom = new CDocumentBorder();
		this.Brd.Bottom.Read_FromBinary(Reader);
	}

	if (Flags & 1024)
	{
		this.Brd.Left = new CDocumentBorder();
		this.Brd.Left.Read_FromBinary(Reader);
	}

	if (Flags & 2048)
	{
		this.Brd.Right = new CDocumentBorder();
		this.Brd.Right.Read_FromBinary(Reader);
	}

	if (Flags & 4096)
	{
		this.Brd.Top = new CDocumentBorder();
		this.Brd.Top.Read_FromBinary(Reader);
	}

	if (Flags & 8192)
		this.WidowControl = Reader.GetBool();

	if (Flags & 16384)
	{
		this.Tabs = new CParaTabs();
		this.Tabs.Read_FromBinary(Reader);
	}

	if (Flags & 32768)
	{
		this.NumPr = new AscWord.NumPr();
		this.NumPr.Read_FromBinary(Reader);
	}

	if (Flags & 65536)
		this.PStyle = Reader.GetString2();

	if (Flags & 131072)
	{
		this.FramePr = new CFramePr();
		this.FramePr.Read_FromBinary(Reader);
	}

	if (Flags & 262144)
	{
		this.DefaultRunPr = new CTextPr();
		this.DefaultRunPr.Read_FromBinary(Reader);
	}

	if (Flags & 524288)
	{
		this.Bullet = new AscFormat.CBullet();
		this.Bullet.Read_FromBinary(Reader);
	}

	if (Flags & 1048576)
	{
		this.Lvl = Reader.GetByte();
	}

	if (Flags & 2097152)
	{
		this.DefaultTab = Reader.GetDouble();
	}

	if (Flags & 4194304)
		this.OutlineLvl = Reader.GetByte();

	if (Flags & 8388608)
	{
		this.PrChange   = new CParaPr();
		this.ReviewInfo = new AscWord.ReviewInfo();
		this.PrChange.ReadFromBinary(Reader);
		this.ReviewInfo.ReadFromBinary(Reader);
	}

	if (Flags & 16777216)
		this.SuppressLineNumbers = Reader.GetBool();
	
	if (Flags & (1 << 25))
		this.Bidi = Reader.GetBool();
};
CParaPr.prototype.isEqual = function(ParaPrUOld,ParaPrNew)
{
	if (ParaPrUOld == undefined || ParaPrNew == undefined)
		return false;
	for (var pPr in ParaPrUOld)
	{
		if (typeof ParaPrUOld[pPr] == 'object')
		{
			if (!this.isEqual(ParaPrUOld[pPr], ParaPrNew[pPr]))
				return false
		}
		else
		{
			if (typeof ParaPrUOld[pPr] == "number" && typeof ParaPrNew[pPr] == "number")
			{
				if (Math.abs(ParaPrUOld[pPr] - ParaPrNew[pPr]) > 0.001)
					return false;
			}
			else if (ParaPrUOld[pPr] != ParaPrNew[pPr])
				return false;
		}
	}
	return true;
};
CParaPr.prototype.Is_Equal = function(ParaPr)
{
	return !(this.ContextualSpacing !== ParaPr.ContextualSpacing
		|| true !== IsEqualStyleObjects(this.Ind, ParaPr.Ind)
		|| this.Jc !== ParaPr.Jc
		|| this.KeepLines !== ParaPr.KeepLines
		|| this.KeepNext !== ParaPr.KeepNext
		|| this.PageBreakBefore !== ParaPr.PageBreakBefore
		|| true !== IsEqualStyleObjects(this.Spacing, ParaPr.Spacing)
		|| true !== IsEqualStyleObjects(this.Shd, ParaPr.Shd)
		|| true !== IsEqualStyleObjects(this.Brd.Between, ParaPr.Brd.Between)
		|| true !== IsEqualStyleObjects(this.Brd.Bottom, ParaPr.Brd.Bottom)
		|| true !== IsEqualStyleObjects(this.Brd.Left, ParaPr.Brd.Left)
		|| true !== IsEqualStyleObjects(this.Brd.Right, ParaPr.Brd.Right)
		|| true !== IsEqualStyleObjects(this.Brd.Top, ParaPr.Brd.Top)
		|| this.WidowControl !== ParaPr.WidowControl
		|| true !== IsEqualStyleObjects(this.Tabs, ParaPr.Tabs)
		|| true !== IsEqualStyleObjects(this.NumPr, ParaPr.NumPr)
		|| this.PStyle !== ParaPr.PStyle
		|| true !== IsEqualStyleObjects(this.FramePr, ParaPr.FramePr)
		|| this.OutlineLvl !== ParaPr.OutlineLvl
		|| this.SuppressLineNumbers !== ParaPr.SuppressLineNumbers
		|| this.Bidi !== ParaPr.Bidi
	);
};
CParaPr.prototype.IsEqual = function(paraPr)
{
	return this.Is_Equal(paraPr);
};
/**
 * Сравниваем данные настройки с заданными, если настройка совпала ставим undefined, если нет, то берем из текущей
 * @param oParaPr {CParaPr}
 * @returns {CParaPr}
 */
CParaPr.prototype.GetDiff = function(oParaPr)
{
	var oResultParaPr = new CParaPr();

	if (this.ContextualSpacing !== oParaPr.ContextualSpacing)
		oResultParaPr.ContextualSpacing = this.ContextualSpacing;

	if (!this.Ind.IsEqual(oParaPr.Ind))
		oResultParaPr.Ind = this.Ind.Copy();

	if (this.Jc !== oParaPr.Jc)
		oResultParaPr.Jc = this.Jc;

	if (this.KeepLines !== oParaPr.KeepLines)
		oResultParaPr.KeepLines = this.KeepLines;

	if (this.KeepNext !== oParaPr.KeepNext)
		oResultParaPr.KeepNext = this.KeepNext;

	if (this.PageBreakBefore !== oParaPr.PageBreakBefore)
		oResultParaPr.PageBreakBefore = this.PageBreakBefore;

	if (!this.Spacing.IsEqual(oParaPr.Spacing))
		oResultParaPr.Spacing = this.Spacing.Copy();

	if (this.Shd && !this.Shd.IsEqual(oParaPr.Shd))
		oResultParaPr.Shd = this.Shd.Copy();

	if (this.Brd.Between && !this.Brd.Between.IsEqual(oParaPr.Brd.Between))
		oResultParaPr.Brd.Between = this.Brd.Between.Copy();

	if (this.Brd.Bottom && !this.Brd.Between.IsEqual(oParaPr.Brd.Bottom))
		oResultParaPr.Brd.Bottom = this.Brd.Bottom.Copy();

	if (this.Brd.Left && !this.Brd.Between.IsEqual(oParaPr.Brd.Left))
		oResultParaPr.Brd.Left = this.Brd.Left.Copy();

	if (this.Brd.Right && !this.Brd.Between.IsEqual(oParaPr.Brd.Right))
		oResultParaPr.Brd.Right = this.Brd.Right.Copy();

	if (this.Brd.Top && !this.Brd.Between.IsEqual(oParaPr.Brd.Top))
		oResultParaPr.Brd.Top = this.Brd.Top.Copy();

	if (this.WidowControl !== oParaPr.WidowControl)
		oResultParaPr.WidowControl = this.WidowControl;

	if (this.Tabs && !this.Tabs.IsEqual(oParaPr.Tabs))
		oResultParaPr.Tabs = this.Tabs.Copy();

	if (this.NumPr && !this.NumPr.IsEqual(oParaPr.NumPr))
		oResultParaPr.NumPr = this.NumPr.Copy();

	if (this.PStyle !== oParaPr.PStyle)
		oResultParaPr.PStyle = this.PStyle;

	if (this.FramePr && !this.FramePr.IsEqual(oParaPr.FramePr))
		oResultParaPr.FramePr = this.FramePr.Copy();

	if (this.OutlineLvl !== oParaPr.OutlineLvl)
		oResultParaPr.OutlineLvl = this.OutlineLvl;

	if (this.DefaultRunPr && !this.DefaultRunPr.IsEqual(oParaPr.DefaultRunPr))
		oResultParaPr.DefaultRunPr = this.DefaultRunPr.Copy();

	if (this.SuppressLineNumbers !== oParaPr.SuppressLineNumbers)
		oResultParaPr.SuppressLineNumbers = this.SuppressLineNumbers;
	
	if (this.Bidi !== oParaPr.Bidi)
		oResultParaPr.Bidi = this.Bidi;

	return oResultParaPr;
};
CParaPr.prototype.Get_PresentationBullet = function(theme, colorMap)
{
	var Bullet = new CPresentationBullet();
	if (this.Bullet && this.Bullet.isBullet())
	{
		switch (this.Bullet.bulletType.type)
		{
			case AscFormat.BULLET_TYPE_BULLET_CHAR:
			{
				Bullet.m_nType = AscFormat.numbering_presentationnumfrmt_Char;
				if (typeof this.Bullet.bulletType.Char === "string" && this.Bullet.bulletType.Char.length > 0)
				{
					Bullet.m_sChar = this.Bullet.bulletType.Char.substring(0, 1);
				}
				else
				{
					Bullet.m_sChar = "•";
				}
				break;
			}

			case AscFormat.BULLET_TYPE_BULLET_AUTONUM :
			{
				Bullet.m_nType    = this.Bullet.bulletType.AutoNumType;
				if(this.Bullet.bulletType.startAt === null)
				{
					Bullet.m_nStartAt = 1;
				}
				else
				{
					Bullet.m_nStartAt = this.Bullet.bulletType.startAt;
				}
				break;
			}
			case AscFormat.BULLET_TYPE_BULLET_NONE :
			{
				break;
			}
			case AscFormat.BULLET_TYPE_BULLET_BLIP :
			{
				Bullet.m_nType = AscFormat.numbering_presentationnumfrmt_Blip;

				var imageBulletURL = this.Bullet.getImageBulletURL();
				if (imageBulletURL)
				{
					Bullet.m_sSrc = AscCommon.getFullImageSrc2(imageBulletURL);
				}
				break;
			}
		}

		if (this.Bullet.bulletColor)
		{
			if (this.Bullet.bulletColor.type === AscFormat.BULLET_TYPE_COLOR_NONE)
			{
				Bullet.m_bColorTx = false;
				Bullet.m_oColor.a = 0;
			}
			if (this.Bullet.bulletColor.type === AscFormat.BULLET_TYPE_COLOR_CLR)
			{
				if (this.Bullet.bulletColor.UniColor && this.Bullet.bulletColor.UniColor.color && theme && colorMap)
				{
					Bullet.m_bColorTx = false;
					Bullet.Unifill    = AscFormat.CreateUniFillByUniColorCopy(this.Bullet.bulletColor.UniColor);
				}
			}
		}
		if (this.Bullet.bulletTypeface)
		{
			if (this.Bullet.bulletTypeface.type === AscFormat.BULLET_TYPE_TYPEFACE_BUFONT
				&& this.Bullet.bulletType.type === AscFormat.BULLET_TYPE_BULLET_CHAR)
			{
				Bullet.m_bFontTx = false;
				Bullet.m_sFont   = this.Bullet.bulletTypeface.typeface;
			}

		}
		if (this.Bullet.bulletSize)
		{
			switch (this.Bullet.bulletSize.type)
			{
				case AscFormat.BULLET_TYPE_SIZE_TX:
				{
					Bullet.m_bSizeTx = true;
					break;
				}
				case AscFormat.BULLET_TYPE_SIZE_PCT:
				{
					Bullet.m_bSizeTx  = false;
					Bullet.m_bSizePct = true;
					Bullet.m_dSize    = this.Bullet.bulletSize.val / 100000.0;
					break;
				}
				case AscFormat.BULLET_TYPE_SIZE_PTS:
				{
					Bullet.m_bSizeTx  = false;
					Bullet.m_bSizePct = false;
					Bullet.m_dSize    = this.Bullet.bulletSize.val / 100.0;
					break;
				}
			}
		}
	}
	return Bullet;
};
CParaPr.prototype.Is_Empty = function(oPr)
{
	const bIsSingleLvlPresetJSON = !!(oPr && oPr.isSingleLvlPresetJSON);
	return !(undefined !== this.ContextualSpacing
		|| true !== (bIsSingleLvlPresetJSON || this.Ind.Is_Empty())
		|| undefined !== this.Jc
		|| undefined !== this.KeepLines
		|| undefined !== this.KeepNext
		|| undefined !== this.PageBreakBefore
		|| true !== this.Spacing.Is_Empty()
		|| undefined !== this.Shd
		|| undefined !== this.Brd.First
		|| undefined !== this.Brd.Last
		|| undefined !== this.Brd.Between
		|| undefined !== this.Brd.Bottom
		|| undefined !== this.Brd.Left
		|| undefined !== this.Brd.Right
		|| undefined !== this.Brd.Top
		|| undefined !== this.WidowControl
		|| (undefined !== this.Tabs && !bIsSingleLvlPresetJSON)
		|| undefined !== this.NumPr
		|| undefined !== this.PStyle
		|| undefined !== this.OutlineLvl
		|| undefined !== this.SuppressLineNumbers
		|| undefined !== this.Bidi
	);
};
CParaPr.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CParaPr.prototype.IsEmptyBorders = function()
{
	return (!this.Brd
		|| (!this.Brd.First
			&& !this.Brd.Last
			&& !this.Brd.Between
			&& !this.Brd.Bottom
			&& !this.Brd.Left
			&& !this.Brd.Right
			&& !this.Brd.Top));
};
CParaPr.prototype.GetDiffPrChange = function()
{
	var ParaPr = new CParaPr();

	if (false === this.HavePrChange())
		return ParaPr;

	var PrChange = this.PrChange;

	if (this.ContextualSpacing !== PrChange.ContextualSpacing)
		ParaPr.ContextualSpacing = this.ContextualSpacing;

	ParaPr.Ind = this.Ind.Get_Diff(PrChange.Ind);

	if (this.Jc !== PrChange.Jc)
		ParaPr.Jc = this.Jc;

	if (this.KeepLines !== PrChange.KeepLines)
		ParaPr.KeepLines = this.KeepLines;

	if (this.KeepNext !== PrChange.KeepNext)
		ParaPr.KeepNext = this.KeepNext;

	if (this.PageBreakBefore !== PrChange.PageBreakBefore)
		ParaPr.PageBreakBefore = this.PageBreakBefore;

	ParaPr.Spacing = this.Spacing.Get_Diff(PrChange.Spacing);

	// TODO: Shd
	// TODO: Brd

	if (this.WidowControl !== PrChange.WidowControl)
		ParaPr.WidowControl = this.WidowControl;

	if (this.Tabs !== PrChange.Tabs)
		ParaPr.Tabs = this.Tabs;

	if (this.NumPr !== PrChange.NumPr)
		ParaPr.NumPr = this.NumPr;

	if (this.PStyle !== PrChange.PStyle)
		ParaPr.PStyle = this.PStyle;

	return ParaPr;
};
CParaPr.prototype.HavePrChange = function()
{
	if (undefined === this.PrChange || null === this.PrChange)
		return false;

	return true;
};
CParaPr.prototype.GetPrChangeNumPr = function()
{
	if (!this.HavePrChange() || !this.PrChange.NumPr)
		return null;

	return this.PrChange.NumPr;
};
CParaPr.prototype.AddPrChange = function()
{
	this.PrChange   = this.Copy();
	this.ReviewInfo = new AscWord.ReviewInfo();
	this.ReviewInfo.Update();
};
CParaPr.prototype.SetPrChange = function(PrChange, ReviewInfo)
{
	this.PrChange   = PrChange;
	this.ReviewInfo = ReviewInfo;
};
CParaPr.prototype.RemovePrChange = function()
{
	delete this.PrChange;
	delete this.ReviewInfo;
};
CParaPr.prototype.GetBidi = function()
{
	return this.Bidi;
};
CParaPr.prototype.SetBidi = function(isBidi)
{
	this.Bidi = isBidi;
};
CParaPr.prototype.GetContextualSpacing = function()
{
    return this.ContextualSpacing;
};
CParaPr.prototype.SetContextualSpacing = function(isContextualSpacing)
{
	this.ContextualSpacing = isContextualSpacing;
};
CParaPr.prototype.GetIndLeft = function()
{
    return this.Ind.Left;
};
CParaPr.prototype.GetIndRight = function()
{
    return this.Ind.Right;
};
CParaPr.prototype.GetIndFirstLine = function()
{
    return this.Ind.FirstLine;
};
CParaPr.prototype.SetInd = function(nFirst, nLeft, nRight)
{
	if (undefined !== nFirst)
		this.Ind.FirstLine = nFirst;

	if (undefined !== nLeft)
		this.Ind.Left = nLeft;

	if (undefined !== nRight)
		this.Ind.Right = nRight;
};
CParaPr.prototype.GetJc = function()
{
    return this.Jc;
};
CParaPr.prototype.SetJc = function(nAlign)
{
	this.Jc = nAlign;
};
CParaPr.prototype.GetKeepLines = function()
{
    return this.KeepLines;
};
CParaPr.prototype.SetKeepLines = function(isKeepLines)
{
	this.KeepLines = isKeepLines;
};
CParaPr.prototype.GetKeepNext = function()
{
    return this.KeepNext;
};
CParaPr.prototype.SetKeepNext = function(isKeepNext)
{
	this.KeepNext = isKeepNext;
};
CParaPr.prototype.GetPageBreakBefore = function()
{
    return this.PageBreakBefore;
};
CParaPr.prototype.SetPageBreakBefore = function(isPageBreakBefore)
{
	this.PageBreakBefore = isPageBreakBefore;
};
CParaPr.prototype.GetSpacingLine = function()
{
    return this.Spacing.Line;
};
CParaPr.prototype.GetSpacingLineRule = function()
{
    return this.Spacing.LineRule;
};
CParaPr.prototype.GetSpacingBefore = function()
{
    return this.Spacing.Before;
};
CParaPr.prototype.GetSpacingBeforeAutoSpacing = function()
{
    return this.Spacing.BeforeAutoSpacing;
};
CParaPr.prototype.GetSpacingAfter = function()
{
    return this.Spacing.After;
};
CParaPr.prototype.GetSpacingAfterAutoSpacing = function()
{
    return this.Spacing.AfterAutoSpacing;
};
CParaPr.prototype.SetSpacing = function(nLine, nLineRule, nBefore, nAfter, isBeforeAuto, isAfterAuto)
{
	if (undefined !== nLine)
		this.Spacing.Line = nLine;

	if (undefined !== nLineRule)
		this.Spacing.LineRule = nLineRule;

	if (undefined !== nBefore)
		this.Spacing.Before = nBefore;

	if (undefined !== nAfter)
		this.Spacing.After = nAfter;

	if (undefined !== isBeforeAuto)
		this.Spacing.BeforeAutoSpacing = isBeforeAuto;

	if (undefined !== isAfterAuto)
		this.Spacing.AfterAutoSpacing = isAfterAuto;
};
CParaPr.prototype.GetWidowControl = function()
{
    return this.WidowControl;
};
CParaPr.prototype.SetWidowControl = function(isWidowControl)
{
	this.WidowControl = isWidowControl;
};
CParaPr.prototype.GetTabs = function()
{
    return this.Tabs;
};
CParaPr.prototype.GetNumPr = function()
{
    return this.NumPr;
};
CParaPr.prototype.SetNumPr = function(sNumId, nLvl)
{
	if (undefined === sNumId)
		this.NumPr = undefined;
	else
		this.NumPr = new AscWord.NumPr(sNumId, nLvl);
};
CParaPr.prototype.GetPStyle = function()
{
    return this.PStyle;
};
CParaPr.prototype.SetPStyle = function(sPStyle)
{
	this.PStyle = sPStyle;
};
CParaPr.prototype.GetOutlineLvl = function()
{
	return this.OutlineLvl;
};
CParaPr.prototype.SetOutlineLvl = function(nOutlineLvl)
{
	this.OutlineLvl = nOutlineLvl;
};
CParaPr.prototype.GetSuppressLineNumbers = function()
{
	return this.SuppressLineNumbers;
};
CParaPr.prototype.SetSuppressLineNumbers = function(isSuppress)
{
	this.SuppressLineNumbers = isSuppress;
};
CParaPr.prototype.WriteToBinary = function(oWriter)
{
	return this.Write_ToBinary(oWriter);
};
CParaPr.prototype.ReadFromBinary = function(oReader)
{
	return this.Read_FromBinary(oReader);
};
CParaPr.prototype.private_CorrectBorderSpace = function(nValue)
{
	var dKoef = (32 * 25.4 / 72);
	return (nValue - Math.floor(nValue / dKoef) * dKoef);
};
CParaPr.prototype.CheckBorderSpaces = function()
{
	// MSWordHack: Специальная заглушка под MS Word
	// В Word значение Space ограничивается 0..31пт. Судя по всему у них это значение реализуется ровно 5 битами, т.к.
	// значение 32пт, уже воспринимается как 0, 33=1,34=2 и т.д.

	if (this.Brd.Top)
		this.Brd.Top.Space = this.private_CorrectBorderSpace(this.Brd.Top.Space);

	if (this.Brd.Bottom)
		this.Brd.Bottom.Space = this.private_CorrectBorderSpace(this.Brd.Bottom.Space);

	if (this.Brd.Left)
		this.Brd.Left.Space = this.private_CorrectBorderSpace(this.Brd.Left.Space);

	if (this.Brd.Right)
		this.Brd.Right.Space = this.private_CorrectBorderSpace(this.Brd.Right.Space);

	if (this.Brd.Between)
		this.Brd.Between.Space = this.private_CorrectBorderSpace(this.Brd.Between.Space);
};

//----------------------------------------------------------------------------------------------------------------------
// CParaPr Export
//----------------------------------------------------------------------------------------------------------------------
CParaPr.prototype['get_Bidi']                     = CParaPr.prototype.get_Bidi                     = CParaPr.prototype['Get_Bidi']                     = CParaPr.prototype.GetBidi;
CParaPr.prototype['put_Bidi']                     = CParaPr.prototype.put_Bidi                     = CParaPr.prototype.SetBidi;
CParaPr.prototype['get_ContextualSpacing']        = CParaPr.prototype.get_ContextualSpacing        = CParaPr.prototype['Get_ContextualSpacing']        = CParaPr.prototype.GetContextualSpacing;
CParaPr.prototype['put_ContextualSpacing']        = CParaPr.prototype.put_ContextualSpacing        = CParaPr.prototype.SetContextualSpacing;
CParaPr.prototype['get_IndLeft']                  = CParaPr.prototype.get_IndLeft                  = CParaPr.prototype['Get_IndLeft']                  = CParaPr.prototype.GetIndLeft;
CParaPr.prototype['get_IndRight']                 = CParaPr.prototype.get_IndRight                 = CParaPr.prototype['Get_IndRight']                 = CParaPr.prototype.GetIndRight;
CParaPr.prototype['get_IndFirstLine']             = CParaPr.prototype.get_IndFirstLine             = CParaPr.prototype['Get_IndFirstLine']             = CParaPr.prototype.GetIndFirstLine;
CParaPr.prototype['put_Ind']                      = CParaPr.prototype.put_Ind                      = CParaPr.prototype.SetInd;
CParaPr.prototype['get_Jc']                       = CParaPr.prototype.get_Jc                       = CParaPr.prototype['Get_Jc']                       = CParaPr.prototype.GetJc;
CParaPr.prototype['put_Jc']                       = CParaPr.prototype.put_Jc                       = CParaPr.prototype.SetJc;
CParaPr.prototype['get_KeepLines']                = CParaPr.prototype.get_KeepLines                = CParaPr.prototype['Get_KeepLines']                = CParaPr.prototype.GetKeepLines;
CParaPr.prototype['put_KeepLines']                = CParaPr.prototype.put_KeepLines                = CParaPr.prototype.SetKeepLines;
CParaPr.prototype['get_KeepNext']                 = CParaPr.prototype.get_KeepNext                 = CParaPr.prototype['Get_KeepNext']                 = CParaPr.prototype.GetKeepNext;
CParaPr.prototype['put_KeepNext']                 = CParaPr.prototype.put_KeepNext                 = CParaPr.prototype.SetKeepNext;
CParaPr.prototype['get_PageBreakBefore']          = CParaPr.prototype.get_PageBreakBefore          = CParaPr.prototype['Get_PageBreakBefore']          = CParaPr.prototype.GetPageBreakBefore;
CParaPr.prototype['put_PageBreakBefore']          = CParaPr.prototype.put_PageBreakBefore          = CParaPr.prototype.SetPageBreakBefore;
CParaPr.prototype['get_SpacingLine']              = CParaPr.prototype.get_SpacingLine              = CParaPr.prototype['Get_SpacingLine']              = CParaPr.prototype.GetSpacingLine;
CParaPr.prototype['get_SpacingLineRule']          = CParaPr.prototype.get_SpacingLineRule          = CParaPr.prototype['Get_SpacingLineRule']          = CParaPr.prototype.GetSpacingLineRule;
CParaPr.prototype['get_SpacingBefore']            = CParaPr.prototype.get_SpacingBefore            = CParaPr.prototype['Get_SpacingBefore']            = CParaPr.prototype.GetSpacingBefore;
CParaPr.prototype['get_SpacingBeforeAutoSpacing'] = CParaPr.prototype.get_SpacingBeforeAutoSpacing = CParaPr.prototype['Get_SpacingBeforeAutoSpacing'] = CParaPr.prototype.GetSpacingBeforeAutoSpacing;
CParaPr.prototype['get_SpacingAfter']             = CParaPr.prototype.get_SpacingAfter             = CParaPr.prototype['Get_SpacingAfter']             = CParaPr.prototype.GetSpacingAfter;
CParaPr.prototype['get_SpacingAfterAutoSpacing']  = CParaPr.prototype.get_SpacingAfterAutoSpacing  = CParaPr.prototype['Get_SpacingAfterAutoSpacing']  = CParaPr.prototype.GetSpacingAfterAutoSpacing;
CParaPr.prototype['put_Spacing']                  = CParaPr.prototype.put_Spacing                  = CParaPr.prototype.SetSpacing;
CParaPr.prototype['get_WidowControl']             = CParaPr.prototype.get_WidowControl             = CParaPr.prototype['Get_WidowControl']             = CParaPr.prototype.GetWidowControl;
CParaPr.prototype['put_WidowControl']             = CParaPr.prototype.put_WidowControl             = CParaPr.prototype.SetWidowControl;
CParaPr.prototype['get_Tabs']                     = CParaPr.prototype.get_Tabs                     = CParaPr.prototype['Get_Tabs']                     = CParaPr.prototype.GetTabs;
CParaPr.prototype['get_NumPr']                    = CParaPr.prototype.get_NumPr                    = CParaPr.prototype['Get_NumPr']                    = CParaPr.prototype.GetNumPr;
CParaPr.prototype['put_NumPr']                    = CParaPr.prototype.put_NumPr                    = CParaPr.prototype.SetNumPr;
CParaPr.prototype['get_PStyle']                   = CParaPr.prototype.get_PStyle                   = CParaPr.prototype['Get_PStyle']                   = CParaPr.prototype.GetPStyle;
CParaPr.prototype['put_PStyle']                   = CParaPr.prototype.put_PStyle                   = CParaPr.prototype.SetPStyle;
CParaPr.prototype['get_OutlineLvl']               = CParaPr.prototype.get_OutlineLvl               = CParaPr.prototype['Get_OutlineLvl']               = CParaPr.prototype.GetOutlineLvl;
CParaPr.prototype['put_OutlineLvl']               = CParaPr.prototype.put_OutlineLvl               = CParaPr.prototype.SetOutlineLvl;
CParaPr.prototype['get_SuppressLineNumbers']      = CParaPr.prototype.get_SuppressLineNumbers      = CParaPr.prototype['Get_SuppressLineNumbers']      = CParaPr.prototype.GetSuppressLineNumbers;
CParaPr.prototype['pet_SuppressLineNumbers']      = CParaPr.prototype.put_SuppressLineNumbers      = CParaPr.prototype.SetSuppressLineNumbers;
//----------------------------------------------------------------------------------------------------------------------

function Copy_Bounds(Bounds)
{
    if ( undefined === Bounds )
        return {};
    var Bounds_new = {};
    Bounds_new.Bottom = Bounds.Bottom;
    Bounds_new.Left   = Bounds.Left;
    Bounds_new.Right  = Bounds.Right;
    Bounds_new.Top    = Bounds.Top;
    return Bounds_new;
}

function asc_CStyle()
{
    this.Name    = "";

    this.BasedOn = "";
    this.Next    = "";
    this.Link    = "";
    this.Type    = styletype_Paragraph;

    this.TextPr = new CTextPr();
    this.ParaPr = new CParaPr();
}
asc_CStyle.prototype.get_Name = function()
{
    return this.Name;
};
asc_CStyle.prototype.put_Name = function(sName)
{
    this.Name = sName;
};
asc_CStyle.prototype.put_BasedOn = function(Name)
{
    this.BasedOn = Name;
};
asc_CStyle.prototype.get_BasedOn = function()
{
    return this.BasedOn;
};
asc_CStyle.prototype.put_Next = function(Name)
{
    this.Next = Name;
};
asc_CStyle.prototype.get_Next = function()
{
    return this.Next;
};
asc_CStyle.prototype.put_Type = function(Type)
{
    this.Type = Type;
};
asc_CStyle.prototype.get_Type = function()
{
    return this.Type;
};
asc_CStyle.prototype.put_Link = function(LinkStyle)
{
    this.Link = LinkStyle;
};
asc_CStyle.prototype.get_Link = function()
{
    return this.Link;
};
asc_CStyle.prototype.fill_ParaPr = function(oPr)
{
    this.ParaPr = oPr.Copy();
};
asc_CStyle.prototype.get_ParaPr = function()
{
    return this.ParaPr;
};
asc_CStyle.prototype.fill_TextPr = function(oPr)
{
    this.TextPr = oPr.Copy();
};
asc_CStyle.prototype.get_TextPr = function()
{
    return this.TextPr;
};
//---------------------------------------------------------export---------------------------------------------------
window['Asc'] = window['Asc'] || {};
window['AscWord'] = window['AscWord'] || {};
window['AscCommonWord'] = window['AscCommonWord'] || {};
window["Asc"]["asc_CStyle"] = window["Asc"].asc_CStyle = asc_CStyle;
asc_CStyle.prototype["get_Name"]    = asc_CStyle.prototype.get_Name;
asc_CStyle.prototype["put_Name"]    = asc_CStyle.prototype.put_Name;
asc_CStyle.prototype["get_BasedOn"] = asc_CStyle.prototype.get_BasedOn;
asc_CStyle.prototype["put_BasedOn"] = asc_CStyle.prototype.put_BasedOn;
asc_CStyle.prototype["get_Next"]    = asc_CStyle.prototype.get_Next;
asc_CStyle.prototype["put_Next"]    = asc_CStyle.prototype.put_Next;
asc_CStyle.prototype["get_Type"]    = asc_CStyle.prototype.get_Type;
asc_CStyle.prototype["put_Type"]    = asc_CStyle.prototype.put_Type;
asc_CStyle.prototype["get_Link"]    = asc_CStyle.prototype.get_Link;
asc_CStyle.prototype["put_Link"]    = asc_CStyle.prototype.put_Link;

window["AscCommonWord"].CDocumentColor = CDocumentColor;
window["AscCommonWord"].CStyle = CStyle;
window["AscCommonWord"].CTextPr = CTextPr;
window["AscCommonWord"].CLang = CLang;
window["AscCommonWord"].CParaPr = CParaPr;
window["AscCommonWord"].CParaTabs = CParaTabs;
window["AscCommonWord"].CParaTab = CParaTab;
window["AscCommonWord"].CDocumentShd = CDocumentShd;
window["AscCommonWord"].g_dKoef_pt_to_mm = g_dKoef_pt_to_mm;
window["AscCommonWord"].g_dKoef_pc_to_mm = g_dKoef_pc_to_mm;
window["AscCommonWord"].g_dKoef_in_to_mm = g_dKoef_in_to_mm;
window["AscCommonWord"].g_dKoef_twips_to_mm = g_dKoef_twips_to_mm;
window["AscCommonWord"].g_dKoef_emu_to_mm = g_dKoef_emu_to_mm;
window["AscCommonWord"].g_dKoef_emu_to_twips = g_dKoef_emu_to_twips;
window["AscCommonWord"].g_dKoef_mm_to_twips = g_dKoef_mm_to_twips;
window["AscCommonWord"].g_dKoef_mm_to_pt = g_dKoef_mm_to_pt;
window["AscCommonWord"].g_dKoef_mm_to_emu = g_dKoef_mm_to_emu;
window["AscCommonWord"].g_dKoef_twips_to_pt = g_dKoef_twips_to_pt;
window["AscCommonWord"].g_dKoef_twips_to_emu = g_dKoef_twips_to_emu;
window["AscCommonWord"].g_dKoef_em_to_mm = g_dKoef_em_to_mm;
window["AscCommonWord"].g_dKoef_pt_to_twips = g_dKoef_pt_to_twips;
window["AscCommonWord"].border_Single = border_Single;
window["AscCommonWord"].Default_Tab_Stop = Default_Tab_Stop;
window["AscCommonWord"].highlight_None = highlight_None;
window["AscCommonWord"].spacing_Auto = spacing_Auto;
window["AscCommonWord"].wrap_Around = wrap_Around;
window["AscCommonWord"].wrap_Auto = wrap_Auto;
window["AscCommonWord"].wrap_None = wrap_None;
window["AscCommonWord"].wrap_NotBeside = wrap_NotBeside;
window["AscCommonWord"].wrap_Through = wrap_Through;
window["AscCommonWord"].wrap_Tight = wrap_Tight;

window["AscWord"].CTextPr = CTextPr;
window["AscWord"].CParaPr = CParaPr;
window["AscWord"].CStyle  = CStyle;
window["AscWord"].CNumPr  = CNumPr;
window["AscWord"].NumPr   = CNumPr;
window["AscWord"].CBorder = CDocumentBorder;
window["AscWord"].CShd    = CDocumentShd;
window["AscWord"].CStyles = CStyles;


// Создаем глобальные дефолтовые стили, чтобы быстро можно было отдать дефолтовые настройки
var g_oDocumentDefaultTextPr = new CTextPr();
var g_oDocumentDefaultParaPr = new CParaPr();
var g_oDocumentDefaultTablePr = new CTablePr();
var g_oDocumentDefaultTableCellPr = new CTableCellPr();
var g_oDocumentDefaultTableRowPr = new CTableRowPr();
var g_oDocumentDefaultTableStylePr = new CTableStylePr();
g_oDocumentDefaultTextPr.InitDefault();
g_oDocumentDefaultParaPr.InitDefault();
g_oDocumentDefaultTablePr.InitDefault();
g_oDocumentDefaultTableCellPr.InitDefault();
g_oDocumentDefaultTableRowPr.InitDefault();
g_oDocumentDefaultTableStylePr.InitDefault();

window["AscWord"].DEFAULT_TEXT_PR        = g_oDocumentDefaultTextPr;
window["AscWord"].DEFAULT_PARA_PR        = g_oDocumentDefaultParaPr;
window["AscWord"].DEFAULT_TABLE_PR       = g_oDocumentDefaultTablePr;
window["AscWord"].DEFAULT_TABLE_CELL_PR  = g_oDocumentDefaultTableCellPr;
window["AscWord"].DEFAULT_TABLE_ROW_PR   = g_oDocumentDefaultTableRowPr;
window["AscWord"].DEFAULT_TABLE_STYLE_PR = g_oDocumentDefaultTableStylePr;

window["AscWord"].styletype_Paragraph = styletype_Paragraph;
window["AscWord"].styletype_Numbering = styletype_Numbering;
window["AscWord"].styletype_Table     = styletype_Table;
window["AscWord"].styletype_Character = styletype_Character;

AscWord.BLACK_COLOR = new AscWord.CDocumentColor(0, 0, 0, false);
AscWord.WHITE_COLOR = new AscWord.CDocumentColor(255, 255, 255, false);

var g_oDocumentDefaultFillColor   = new CDocumentColor(255, 255, 255, true);
var g_oDocumentDefaultStrokeColor = new CDocumentColor(0, 0, 0, true);

window["AscWord"].DEFAULT_STYLES = new CStyles(false);
window["AscCommonWord"].g_oDocumentDefaultFillColor   = g_oDocumentDefaultFillColor;
window["AscCommonWord"].g_oDocumentDefaultStrokeColor = g_oDocumentDefaultStrokeColor;

// ----------------------------------------------------------------
