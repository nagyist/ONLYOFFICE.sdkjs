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
	/**
	 *
	 * @constructor
	 */
	function DocumentPage()
	{
		this.Width   = 0;
		this.Height  = 0;
		this.Margins = {
			Left   : 0,
			Right  : 0,
			Top    : 0,
			Bottom : 0
		};
		
		this.Bounds = new AscWord.CDocumentBounds(0,0,0,0);
		this.Pos    = 0;
		this.EndPos = 0;
		
		this.X      = 0;
		this.Y      = 0;
		this.XLimit = 0;
		this.YLimit = 0;
		
		this.OriginX      = 0; // Начальные значения X, Y без учета использования функции Shift
		this.OriginY      = 0; // Используется, например, при расчете позиции автофигуры внутри ячейки таблицы,
		this.OriginXLimit = 0; // которая имеет вертикальное выравнивание по центру или по низу
		this.OriginYLimit = 0;
		
		this.Sections = [];
		
		this.EndSectionParas = [];
		
		this.ResetStartElement = false;
		this.NextPageNewElement = false;
		
		this.Frames     = [];
		this.FlowTables = [];
	}
	
	DocumentPage.prototype.GetStartPos = function()
	{
		return this.Pos;
	};
	DocumentPage.prototype.GetEndPos = function()
	{
		return this.EndPos;
	};
	DocumentPage.prototype.GetSection = function(nIndex)
	{
		return this.Sections[nIndex] ? this.Sections[nIndex] : null;
	};
	DocumentPage.prototype.GetFirstSectPr = function()
	{
		if (!this.Sections.length)
			return null;
		
		return this.Sections[0].GetSectPr();
	};
	DocumentPage.prototype.GetLastSectPr = function()
	{
		if (!this.Sections.length)
			return null;
		
		return this.Sections[this.Sections.length - 1].GetSectPr();
	};
	DocumentPage.prototype.Update_Limits = function(Limits)
	{
		this.X      = Limits.X;
		this.XLimit = Limits.XLimit;
		this.Y      = Limits.Y;
		this.YLimit = Limits.YLimit;
		
		this.OriginX      = Limits.X;
		this.OriginY      = Limits.Y;
		this.OriginXLimit = Limits.XLimit;
		this.OriginYLimit = Limits.YLimit;
	};
	DocumentPage.prototype.Shift = function(Dx, Dy)
	{
		this.X      += Dx;
		this.XLimit += Dx;
		this.Y      += Dy;
		this.YLimit += Dy;
		
		this.Bounds.Shift(Dx, Dy);
		
		for (var SectionIndex = 0, Count = this.Sections.length; SectionIndex < Count; ++SectionIndex)
		{
			this.Sections[SectionIndex].Shift(Dx, Dy);
		}
	};
	DocumentPage.prototype.Check_EndSectionPara = function(Element)
	{
		var Count = this.EndSectionParas.length;
		for ( var Index = 0; Index < Count; Index++ )
		{
			if ( Element === this.EndSectionParas[Index] )
				return true;
		}
		
		return false;
	};
	DocumentPage.prototype.Copy = function()
	{
		var NewPage = new DocumentPage();
		
		NewPage.Width          = this.Width;
		NewPage.Height         = this.Height;
		NewPage.Margins.Left   = this.Margins.Left;
		NewPage.Margins.Right  = this.Margins.Right;
		NewPage.Margins.Top    = this.Margins.Top;
		NewPage.Margins.Bottom = this.Margins.Bottom;
		
		NewPage.Bounds.CopyFrom(this.Bounds);
		NewPage.Pos    = this.Pos;
		NewPage.EndPos = this.EndPos;
		NewPage.X      = this.X;
		NewPage.Y      = this.Y;
		NewPage.XLimit = this.XLimit;
		NewPage.YLimit = this.YLimit;
		
		for (var SectionIndex = 0, Count = this.Sections.length; SectionIndex < Count; ++SectionIndex)
		{
			NewPage.Sections[SectionIndex] = this.Sections[SectionIndex].Copy();
		}
		
		return NewPage;
	};
	DocumentPage.prototype.AddFrame = function(oFrame)
	{
		if (-1 !== this.private_GetFrameIndex(oFrame.StartIndex))
			return -1;
		
		this.Frames.push(oFrame);
	};
	DocumentPage.prototype.RemoveFrame = function(nStartIndex)
	{
		var nPos = this.private_GetFrameIndex(nStartIndex);
		if (-1 === nPos)
			return;
		
		this.Frames.splice(nPos, 1);
	};
	DocumentPage.prototype.private_GetFrameIndex = function(nStartIndex)
	{
		for (var nIndex = 0, nCount = this.Frames.length; nIndex < nCount; ++nIndex)
		{
			if (nStartIndex === this.Frames[nIndex].StartIndex)
				return nIndex;
		}
		
		return -1;
	};
	DocumentPage.prototype.AddFlowTable = function(oTable)
	{
		if (-1 !== this.private_GetFlowTableIndex(oTable))
			return;
		
		this.FlowTables.push(oTable);
	};
	DocumentPage.prototype.RemoveFlowTable = function(oTable)
	{
		var nPos = this.private_GetFlowTableIndex(oTable);
		if (-1 === nPos)
			return;
		
		this.FlowTables.splice(nPos, 1);
	};
	DocumentPage.prototype.private_GetFlowTableIndex = function(oTable)
	{
		for (var nIndex = 0, nCount = this.FlowTables.length; nIndex < nCount; ++nIndex)
		{
			if (oTable === this.FlowTables[nIndex])
				return nIndex;
		}
		
		return -1;
	};
	DocumentPage.prototype.IsFlowTable = function(oElement)
	{
		return (-1 !== this.private_GetFlowTableIndex(oElement));
	};
	DocumentPage.prototype.IsFrame = function(oElement)
	{
		var nIndex = oElement.GetIndex();
		for (var nFrameIndex = 0, nFramesCount = this.Frames.length; nFrameIndex < nFramesCount; ++nFrameIndex)
		{
			if (this.Frames[nFrameIndex].StartIndex <= nIndex && nIndex < this.Frames[nFrameIndex].StartIndex + this.Frames[nFrameIndex].FlowCount)
				return true;
		}
		
		return false;
	};
	DocumentPage.prototype.CheckFrameClipStart = function(nIndex, oGraphics, oDrawingDocument)
	{
		for (var sId in this.Frames)
		{
			var oFrame = this.Frames[sId];
			
			if (oFrame.StartIndex === nIndex)
			{
				var nPixelError = oDrawingDocument.GetMMPerDot(1);
				
				var nL = oFrame.CalculatedFrame.L2 - nPixelError;
				var nT = oFrame.CalculatedFrame.T2 - nPixelError;
				var nH = oFrame.CalculatedFrame.H2 + 2 * nPixelError;
				var nW = oFrame.CalculatedFrame.W2 + 2 * nPixelError;
				
				oGraphics.SaveGrState();
				oGraphics.AddClipRect(nL, nT, nW, nH);
				return;
			}
		}
	};
	DocumentPage.prototype.CheckFrameClipStart = function(nIndex, oGraphics)
	{
		for (var sId in this.Frames)
		{
			var oFrame = this.Frames[sId];
			
			if (oFrame.StartIndex + oFrame.FlowCount - 1 === nIndex)
			{
				oGraphics.RestoreGrState();
				return
			}
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentPage = DocumentPage;
	
})();
