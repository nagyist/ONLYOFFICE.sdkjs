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

(function()
{
	/**
	 * Class for managing document sections and their header/footer content.
	 *
	 * @param {AscWord.Document} logicDocument - The parent document object
	 * @constructor
	 */
	function DocumentSections(logicDocument)
	{
		this.logicDocument = logicDocument;
		this.Elements = [];
	}
	DocumentSections.prototype.Add = function(SectPr, Index)
	{
		this.Elements.push(new DocumentSection(SectPr, Index));
	};
	DocumentSections.prototype.GetSectionsCount = function()
	{
		return this.Elements.length;
	};
	DocumentSections.prototype.Clear = function()
	{
		this.Elements.length = 0;
	};
	DocumentSections.prototype.Find_ByHdrFtr = function(HdrFtr)
	{
		if (!HdrFtr)
			return -1;
		
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (HdrFtr === SectPr.Get_Header_First()
				|| HdrFtr === SectPr.Get_Header_Default()
				|| HdrFtr === SectPr.Get_Header_Even()
				|| HdrFtr === SectPr.Get_Footer_First()
				|| HdrFtr === SectPr.Get_Footer_Default()
				|| HdrFtr === SectPr.Get_Footer_Even())
				return Index;
		}
		
		return -1;
	};
	DocumentSections.prototype.Reset_HdrFtrRecalculateCache = function()
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.Reset_RecalculateCache();
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.Reset_RecalculateCache();
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.Reset_RecalculateCache();
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.Reset_RecalculateCache();
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.Reset_RecalculateCache();
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.Reset_RecalculateCache();
		}
	};
	DocumentSections.prototype.GetAllParagraphs = function(Props, ParaArray)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.GetAllParagraphs(Props, ParaArray);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.GetAllParagraphs(Props, ParaArray);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.GetAllParagraphs(Props, ParaArray);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.GetAllParagraphs(Props, ParaArray);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.GetAllParagraphs(Props, ParaArray);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.GetAllParagraphs(Props, ParaArray);
		}
	};
	DocumentSections.prototype.GetAllTables = function(oProps, arrTables)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.GetAllTables(oProps, arrTables);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.GetAllTables(oProps, arrTables);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.GetAllTables(oProps, arrTables);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.GetAllTables(oProps, arrTables);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.GetAllTables(oProps, arrTables);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.GetAllTables(oProps, arrTables);
		}
	};
	DocumentSections.prototype.GetAllDrawingObjects = function(arrDrawings)
	{
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var SectPr = this.Elements[nIndex].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.GetAllDrawingObjects(arrDrawings);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.GetAllDrawingObjects(arrDrawings);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.GetAllDrawingObjects(arrDrawings);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.GetAllDrawingObjects(arrDrawings);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.GetAllDrawingObjects(arrDrawings);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.GetAllDrawingObjects(arrDrawings);
		}
	};
	DocumentSections.prototype.UpdateBookmarks = function(oBookmarkManager)
	{
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var SectPr = this.Elements[nIndex].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.UpdateBookmarks(oBookmarkManager);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.UpdateBookmarks(oBookmarkManager);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.UpdateBookmarks(oBookmarkManager);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.UpdateBookmarks(oBookmarkManager);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.UpdateBookmarks(oBookmarkManager);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.UpdateBookmarks(oBookmarkManager);
		}
	};
	DocumentSections.prototype.Document_CreateFontMap = function(FontMap)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.Document_CreateFontMap(FontMap);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.Document_CreateFontMap(FontMap);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.Document_CreateFontMap(FontMap);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.Document_CreateFontMap(FontMap);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.Document_CreateFontMap(FontMap);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.Document_CreateFontMap(FontMap);
		}
	};
	DocumentSections.prototype.Document_CreateFontCharMap = function(FontCharMap)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.Document_CreateFontCharMap(FontCharMap);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.Document_CreateFontCharMap(FontCharMap);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.Document_CreateFontCharMap(FontCharMap);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.Document_CreateFontCharMap(FontCharMap);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.Document_CreateFontCharMap(FontCharMap);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.Document_CreateFontCharMap(FontCharMap);
		}
	};
	DocumentSections.prototype.Document_Get_AllFontNames = function(AllFonts)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.Document_Get_AllFontNames(AllFonts);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.Document_Get_AllFontNames(AllFonts);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.Document_Get_AllFontNames(AllFonts);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.Document_Get_AllFontNames(AllFonts);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.Document_Get_AllFontNames(AllFonts);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.Document_Get_AllFontNames(AllFonts);
		}
	};
	DocumentSections.prototype.Get_Index = function(Index)
	{
		var Count = this.Elements.length;
		
		for (var Pos = 0; Pos < Count; Pos++)
		{
			if (Index <= this.Elements[Pos].Index)
				return Pos;
		}
		
		// Последний элемент здесь это всегда конечная секция документа
		return (Count - 1);
	};
	DocumentSections.prototype.Get_Count = function()
	{
		return this.Elements.length;
	};
	DocumentSections.prototype.Get_SectPr = function(Index)
	{
		return this.GetByContentPos(Index);
	};
	DocumentSections.prototype.Get_SectPr2 = function(Index)
	{
		return this.Elements[Index];
	}
	DocumentSections.prototype.Find = function(SectPr)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var Element = this.Elements[Index];
			if (Element.SectPr === SectPr)
				return Index;
		}
		
		return -1;
	};
	DocumentSections.prototype.Update_OnAdd = function(Pos, Items)
	{
		var Count = Items.length;
		var Len   = this.Elements.length;
		
		// Сначала обновим старые метки
		for (var Index = 0; Index < Len; Index++)
		{
			if (this.Elements[Index].Index >= Pos)
				this.Elements[Index].Index += Count;
		}
		
		// Если среди новых элементов были параграфы с настройками секции, тогда добавим их здесь
		for (var Index = 0; Index < Count; Index++)
		{
			var Item   = Items[Index];
			var SectPr = (type_Paragraph === Item.GetType() ? Item.Get_SectionPr() : undefined);
			
			if (undefined !== SectPr)
			{
				var TempPos = 0;
				for (; TempPos < Len; TempPos++)
				{
					if (Pos + Index <= this.Elements[TempPos].Index)
						break;
				}
				
				this.Elements.splice(TempPos, 0, new DocumentSection(SectPr, Pos + Index));
				Len++;
			}
		}
	};
	DocumentSections.prototype.Update_OnRemove = function(Pos, Count, bCheckHdrFtr)
	{
		var Len = this.Elements.length;
		
		for (var Index = 0; Index < Len; Index++)
		{
			var CurPos = this.Elements[Index].Index;
			
			if (CurPos >= Pos && CurPos < Pos + Count)
			{
				// Копируем поведение Word: Если у следующей секции не задан вообще ни один колонтитул,
				// тогда копируем ссылки на колонтитулы из удаляемой секции. Если задан хоть один колонтитул,
				// тогда этого не делаем.
				if (true === bCheckHdrFtr && Index < Len - 1)
				{
					var CurrSectPr = this.Elements[Index].SectPr;
					var NextSectPr = this.Elements[Index + 1].SectPr;
					if (true === NextSectPr.IsAllHdrFtrNull() && true !== CurrSectPr.IsAllHdrFtrNull())
					{
						NextSectPr.Set_Header_First(CurrSectPr.Get_Header_First());
						NextSectPr.Set_Header_Even(CurrSectPr.Get_Header_Even());
						NextSectPr.Set_Header_Default(CurrSectPr.Get_Header_Default());
						NextSectPr.Set_Footer_First(CurrSectPr.Get_Footer_First());
						NextSectPr.Set_Footer_Even(CurrSectPr.Get_Footer_Even());
						NextSectPr.Set_Footer_Default(CurrSectPr.Get_Footer_Default());
					}
				}
				
				this.Elements.splice(Index, 1);
				Len--;
				Index--;
				
				
			}
			else if (CurPos >= Pos + Count)
				this.Elements[Index].Index -= Count;
		}
	};
	DocumentSections.prototype.GetCount = function()
	{
		return this.Elements.length;
	};
	/**
	 * Получаем секцию по заданному номеру
	 * @param {number} nIndex
	 * @returns {DocumentSection}
	 */
	DocumentSections.prototype.Get = function(nIndex)
	{
		return this.Elements[nIndex];
	};
	/**
	 * Получаем секцию по заданной позиции контента
	 * @param {number} nContentPos
	 * @returns {DocumentSection}
	 */
	DocumentSections.prototype.GetByContentPos = function(nContentPos)
	{
		var nCount = this.Elements.length;
		for (var nPos = 0; nPos < nCount; ++nPos)
		{
			if (nContentPos <= this.Elements[nPos].Index)
				return this.Elements[nPos];
		}
		
		// Последний элемент здесь это всегда конечная секция документа
		return this.Elements[nCount - 1];
	};
	/**
	 * Получаем массив всех колонтитулов, используемых в данном документе
	 * @returns {Array.CHeaderFooter}
	 */
	DocumentSections.prototype.GetAllHdrFtrs = function()
	{
		var HdrFtrs = [];
		
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			SectPr.GetAllHdrFtrs(HdrFtrs);
		}
		
		return HdrFtrs;
	};
	DocumentSections.prototype.GetAllContentControls = function(arrContentControls)
	{
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var SectPr = this.Elements[nIndex].SectPr;
			
			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.GetAllContentControls(arrContentControls);
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.GetAllContentControls(arrContentControls);
			
			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.GetAllContentControls(arrContentControls);
			
			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.GetAllContentControls(arrContentControls);
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.GetAllContentControls(arrContentControls);
			
			if (null != SectPr.FooterEven)
				SectPr.FooterEven.GetAllContentControls(arrContentControls);
		}
	};
	/**
	 * Обновляем заданную секцию
	 * @param oSectPr {CSectionPr} - Секция, которую нужно обновить
	 * @param oNewSectPr {?CSectionPr} - Либо новое значение секции, либо undefined для удалении секции
	 * @param isCheckHdrFtr {boolean} - Нужно ли проверять колонтитулы при удалении секции
	 * @returns {boolean} Если не смогли обновить, возвращаем false
	 */
	DocumentSections.prototype.UpdateSection = function(oSectPr, oNewSectPr, isCheckHdrFtr)
	{
		if (oSectPr === oNewSectPr || !oSectPr)
			return false;
		
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			if (oSectPr === this.Elements[nIndex].SectPr)
			{
				if (!oNewSectPr)
				{
					// Копируем поведение Word: Если у следующей секции не задан вообще ни один колонтитул,
					// тогда копируем ссылки на колонтитулы из удаляемой секции. Если задан хоть один колонтитул,
					// тогда этого не делаем.
					if (true === isCheckHdrFtr && nIndex < nCount - 1)
					{
						var oCurrSectPr = this.Elements[nIndex].SectPr;
						var oNextSectPr = this.Elements[nIndex + 1].SectPr;
						
						if (true === oNextSectPr.IsAllHdrFtrNull() && true !== oCurrSectPr.IsAllHdrFtrNull())
						{
							oNextSectPr.Set_Header_First(oCurrSectPr.Get_Header_First());
							oNextSectPr.Set_Header_Even(oCurrSectPr.Get_Header_Even());
							oNextSectPr.Set_Header_Default(oCurrSectPr.Get_Header_Default());
							oNextSectPr.Set_Footer_First(oCurrSectPr.Get_Footer_First());
							oNextSectPr.Set_Footer_Even(oCurrSectPr.Get_Footer_Even());
							oNextSectPr.Set_Footer_Default(oCurrSectPr.Get_Footer_Default());
						}
					}
					
					this.Elements.splice(nIndex, 1);
				}
				else
				{
					this.Elements[nIndex].SectPr = oNewSectPr;
				}
				
				return true;
			}
		}
		
		return false;
	};
	DocumentSections.prototype.private_GetHdrFtrsArray = function(oCurHdrFtr)
	{
		var isEvenOdd = EvenAndOddHeaders;
		
		var nCurPos    = -1;
		var arrHdrFtrs = [];
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oSectPr = this.Elements[nIndex].SectPr;
			var isFirst = oSectPr.Get_TitlePage();
			
			var oHeaderFirst   = oSectPr.Get_Header_First();
			var oHeaderEven    = oSectPr.Get_Header_Even();
			var oHeaderDefault = oSectPr.Get_Header_Default();
			var oFooterFirst   = oSectPr.Get_Footer_First();
			var oFooterEven    = oSectPr.Get_Footer_Even();
			var oFooterDefault = oSectPr.Get_Footer_Default();
			
			if (oHeaderFirst && isFirst)
				arrHdrFtrs.push(oHeaderFirst);
			
			if (oHeaderEven && isEvenOdd)
				arrHdrFtrs.push(oHeaderEven);
			
			if (oHeaderDefault)
				arrHdrFtrs.push(oHeaderDefault);
			
			if (oFooterFirst && isFirst)
				arrHdrFtrs.push(oFooterFirst);
			
			if (oFooterEven && isEvenOdd)
				arrHdrFtrs.push(oFooterEven);
			
			if (oFooterDefault)
				arrHdrFtrs.push(oFooterDefault);
		}
		
		if (oCurHdrFtr)
		{
			for (var nIndex = 0, nCount = arrHdrFtrs.length; nIndex < nCount; ++nIndex)
			{
				if (oCurHdrFtr === arrHdrFtrs[nIndex])
				{
					nCurPos = nIndex;
					break;
				}
			}
		}
		
		return {
			HdrFtrs : arrHdrFtrs,
			CurPos  : nCurPos
		};
	};
	DocumentSections.prototype.FindNextFillingForm = function(isNext, oCurHdrFtr)
	{
		var oInfo = this.private_GetHdrFtrsArray(oCurHdrFtr);
		
		var arrHdrFtrs = oInfo.HdrFtrs;
		var nCurPos    = oInfo.CurPos;
		
		var nCount = arrHdrFtrs.length;
		
		var isCurrent = true;
		if (-1 === nCurPos)
		{
			isCurrent = false;
			nCurPos   = isNext ? 0 : arrHdrFtrs.length - 1;
			if (arrHdrFtrs[nCurPos])
				oCurHdrFtr = arrHdrFtrs[nCurPos];
		}
		
		if (nCurPos >= 0 && nCurPos <= nCount - 1)
		{
			var oRes = oCurHdrFtr.GetContent().FindNextFillingForm(isNext, isCurrent, isCurrent);
			if (oRes)
				return oRes;
			
			if (isNext)
			{
				for (var nIndex = nCurPos + 1; nIndex < nCount; ++nIndex)
				{
					oRes = arrHdrFtrs[nIndex].GetContent().FindNextFillingForm(isNext, false);
					
					if (oRes)
						return oRes;
				}
			}
			else
			{
				for (var nIndex = nCurPos - 1; nIndex >= 0; --nIndex)
				{
					oRes = arrHdrFtrs[nIndex].GetContent().FindNextFillingForm(isNext, false);
					
					if (oRes)
						return oRes;
				}
			}
		}
		
		return null;
	};
	DocumentSections.prototype.RestartSpellCheck = function()
	{
		var bEvenOdd = EvenAndOddHeaders;
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var SectPr = this.Elements[nIndex].SectPr;
			var bFirst = SectPr.Get_TitlePage();
			
			if (null != SectPr.HeaderFirst && true === bFirst)
				SectPr.HeaderFirst.RestartSpellCheck();
			
			if (null != SectPr.HeaderEven && true === bEvenOdd)
				SectPr.HeaderEven.RestartSpellCheck();
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.RestartSpellCheck();
			
			if (null != SectPr.FooterFirst && true === bFirst)
				SectPr.FooterFirst.RestartSpellCheck();
			
			if (null != SectPr.FooterEven && true === bEvenOdd)
				SectPr.FooterEven.RestartSpellCheck();
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.RestartSpellCheck();
		}
	};
	DocumentSections.prototype.RemoveEmptyHdrFtrs = function()
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oSectPr = this.Elements[nIndex].SectPr;
			oSectPr.RemoveEmptyHdrFtrs();
		}
	};
	DocumentSections.prototype.CheckRunContent = function(fCheck)
	{
		let headers = this.GetAllHdrFtrs();
		for (let index = 0, count = headers.length; index < count; ++index)
		{
			headers[index].GetContent().CheckRunContent(fCheck);
		}
	};
	//------------------------------------------------------------------------------------------------------------------
	// Search
	//------------------------------------------------------------------------------------------------------------------
	DocumentSections.prototype.Search = function(oSearchEngine)
	{
		var bEvenOdd = EvenAndOddHeaders;
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oSectPr = this.Elements[nIndex].SectPr;
			var bFirst  = oSectPr.Get_TitlePage();
			
			if (oSectPr.HeaderFirst && true === bFirst)
				oSectPr.HeaderFirst.Search(oSearchEngine, search_Header);
			
			if (oSectPr.HeaderEven && true === bEvenOdd)
				oSectPr.HeaderEven.Search(oSearchEngine, search_Header);
			
			if (oSectPr.HeaderDefault)
				oSectPr.HeaderDefault.Search(oSearchEngine, search_Header);
			
			if (oSectPr.FooterFirst && true === bFirst)
				oSectPr.FooterFirst.Search(oSearchEngine, search_Footer);
			
			if (oSectPr.FooterEven && true === bEvenOdd)
				oSectPr.FooterEven.Search(oSearchEngine, search_Footer);
			
			if (oSectPr.FooterDefault)
				oSectPr.FooterDefault.Search(oSearchEngine, search_Footer);
		}
	};
	DocumentSections.prototype.GetSearchElementId = function(bNext, CurHdrFtr)
	{
		var HdrFtrs = [];
		var CurPos  = -1;
		
		var bEvenOdd = EvenAndOddHeaders;
		var Count    = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			var bFirst = SectPr.Get_TitlePage();
			
			if (null != SectPr.HeaderFirst && true === bFirst)
			{
				HdrFtrs.push(SectPr.HeaderFirst);
				
				if (CurHdrFtr === SectPr.HeaderFirst)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.HeaderEven && true === bEvenOdd)
			{
				HdrFtrs.push(SectPr.HeaderEven);
				
				if (CurHdrFtr === SectPr.HeaderEven)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.HeaderDefault)
			{
				HdrFtrs.push(SectPr.HeaderDefault);
				
				if (CurHdrFtr === SectPr.HeaderDefault)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.FooterFirst && true === bFirst)
			{
				HdrFtrs.push(SectPr.FooterFirst);
				
				if (CurHdrFtr === SectPr.FooterFirst)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.FooterEven && true === bEvenOdd)
			{
				HdrFtrs.push(SectPr.FooterEven);
				
				if (CurHdrFtr === SectPr.FooterEven)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.FooterDefault)
			{
				HdrFtrs.push(SectPr.FooterDefault);
				
				if (CurHdrFtr === SectPr.FooterDefault)
					CurPos = HdrFtrs.length - 1;
			}
		}
		
		var Count = HdrFtrs.length;
		
		var isCurrent = true;
		if (-1 === CurPos)
		{
			isCurrent = false;
			CurPos    = bNext ? 0 : HdrFtrs.length - 1;
			if (HdrFtrs[CurPos])
				CurHdrFtr = HdrFtrs[CurPos];
		}
		
		if (CurPos >= 0 && CurPos <= HdrFtrs.length - 1)
		{
			var Id = CurHdrFtr.GetSearchElementId(bNext, isCurrent);
			if (null != Id)
				return Id;
			
			if (true === bNext)
			{
				for (var Index = CurPos + 1; Index < Count; Index++)
				{
					Id = HdrFtrs[Index].GetSearchElementId(bNext, false);
					
					if (null != Id)
						return Id;
				}
			}
			else
			{
				for (var Index = CurPos - 1; Index >= 0; Index--)
				{
					Id = HdrFtrs[Index].GetSearchElementId(bNext, false);
					
					if (null != Id)
						return Id;
				}
			}
		}
		
		return null;
	};
	
	
	//----------------------------------------------------------------------------------------------------------------------
	
	/**
	 * Represents a document section associated with a specific paragraph.
	 *
	 * @param {AscWord.SectPr} sectPr - Section properties object containing formatting settings
	 * @param {AscWord.Paragraph} paragraph - The paragraph object that ends this section
	 * @constructor
	 */
	function DocumentSection(sectPr, paragraph)
	{
		this.SectPr = sectPr;
		this.Paragraph = paragraph;
		
		this.Index = paragraph;
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentSections = DocumentSections;
	AscWord.DocumentSection  = DocumentSection;
	
})();
