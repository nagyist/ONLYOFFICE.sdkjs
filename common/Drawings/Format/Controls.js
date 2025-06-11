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

(function() {
	function CControl() {
		AscFormat.CShape.call(this);
		this.name = null;
		this.progId = null;
		this.dvAspect = null;
		this.link = null;
		this.oleUpdate = null;
		this.autoLoad = null;
		this.shapeId = null;
		this.rId = null;
		this.controlPr = new CControlPr();
		this.formControlPr = new CFormControlPr();
	}
	AscFormat.InitClass(CControl, AscFormat.CShape, AscDFH.historyitem_type_Shape);
	CControl.prototype.superclass = AscFormat.CGraphicObjectBase
	CControl.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		const oMainTransfrom = transform || this.transform;
		const checkBoxTransform = oMainTransfrom.CreateDublicate();
		AscFormat.CShape.prototype.draw.call(this, graphics, transform, transformText, pageIndex, opt);
		graphics.SaveGrState();
		checkBoxTransform.tx += 1;
		checkBoxTransform.ty += (this.extY - 3) / 2;
		graphics.transform3(checkBoxTransform);
		graphics.b_color1(255, 255, 255, 255);
		graphics.p_color(0, 0, 0, 255);
		graphics.p_width(0);
		graphics._s();
		graphics._m(0, 0);
		graphics._l(0, 3);
		graphics._l(3, 3);
		graphics._l(3, 0);
		graphics._z();
		graphics.ds();
		graphics.df();
		graphics._e();
		graphics.p_color(0, 0, 0, 255);
		graphics.p_width(400);
		graphics._m(2.5, 0.75);
		graphics._l(1, 2.25);
		graphics._l(0.5, 1.75);
		graphics.ds();
		graphics._e();
		graphics.RestoreGrState();
	};
	
	function CControlPr() {
		this.altText = null;
		this.autoFill = null;
		this.autoLine = null;
		this.autoPict = null;
		this.dde = null;
		this.defaultSize = null;
		this.disabled = null;
		this.cf = null;
		this.linkedCell = null;
		this.listFillRange = null;
		this.rId = null;
		this.locked = null;
		this.macro = null;
		this.print = null;
		this.recalcAlways = null;
		this.uiObject = null;
		this.anchor = null;
	}
	
	function CFormControlPr() {
		this.dropLines = null;
		this.objectType = null;
		this.checked = null;
		this.dropStyle = null;
		this.dx = null;
		this.inc = null;
		this.min = null;
		this.max = null;
		this.page = null;
		this.sel = null;
		this.selType = null;
		this.textHAlign = null;
		this.textVAlign = null;
		this.val = null;
		this.widthMin = null;
		this.editVal = null;
		this.fmlaGroup = null;
		this.fmlaLink = null;
		this.fmlaRange = null;
		this.fmlaTxbx = null;
		this.colored = null;
		this.firstButton = null;
		this.horiz = null;
		this.justLastX = null;
		this.lockText = null;
		this.multiSel = null;
		this.noThreeD = null;
		this.noThreeD2 = null;
		this.multiLine = null;
		this.verticalBar = null;
		this.passwordEdit = null;
		this.itemLst = [];
	}

	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CControl = CControl;
})();
