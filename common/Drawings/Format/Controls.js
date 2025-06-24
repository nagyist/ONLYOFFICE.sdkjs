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

	const CFormControlPr_checked_unchecked = 0;
	const CFormControlPr_checked_checked   = 1;
	const CFormControlPr_checked_mixed     = 2;

	const CFormControlPr_objectType_button = 0;
	const CFormControlPr_objectType_checkBox = 1;
	const CFormControlPr_objectType_drop = 2;
	const CFormControlPr_objectType_gBox = 3;
	const CFormControlPr_objectType_label = 4;
	const CFormControlPr_objectType_list = 5;
	const CFormControlPr_objectType_radio = 6;
	const CFormControlPr_objectType_scroll = 7;
	const CFormControlPr_objectType_spin = 8;
	const CFormControlPr_objectType_editBox = 9;
	const CFormControlPr_objectType_dialog = 10;
	const CFormControlPr_objectType_toggleButton = 11;
	const CFormControlPr_objectType_tabStrip = 12;
	const CFormControlPr_objectType_image = 13;

	AscDFH.changesFactory[AscDFH.historyitem_Control_ControlPr] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_Control_FormControlPr] = AscDFH.CChangesDrawingsObject;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Control_ControlPr] = function(oClass, pr) {
		oClass.controlPr = pr;
	}
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Control_FormControlPr] = function(oClass, pr) {
		oClass.formControlPr = pr;
	}
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
		this.controller = null;
	}
	AscFormat.InitClass(CControl, AscFormat.CShape, AscDFH.historyitem_type_Control);
	CControl.prototype.superclass = AscFormat.CGraphicObjectBase;
	CControl.prototype.initController = function() {
		switch (this.formControlPr.objectType) {
			case CFormControlPr_objectType_checkBox: {
				this.controller = new CCheckBoxController(this);
				break;
			}
			default: {
				return false;
			}
		}
		// this.controller.init();
		return true;
	}
	CControl.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		this.controller.draw(graphics, transform, transformText, pageIndex, opt);
	};
	CControl.prototype.hitInInnerArea = function (x, y) {
		const oInvertTransform = this.getInvertTransform();
		const nX = oInvertTransform.TransformPointX(x, y);
		const nY = oInvertTransform.TransformPointY(x, y);
		return nX > 0 && nX < this.extX && nY > 0 && nY < this.extY;
	}
	CControl.prototype.hitInPath = CControl.prototype.hitInInnerArea;
	CControl.prototype.hitInTextRect = function(x, y) {
		if (this.selected) {
			return AscFormat.CShape.prototype.hitInTextRect.call(this, x, y);
		}
		return false;
	};
	CControl.prototype.isControl = function () {
		return true;
	}
	CControl.prototype.onClick = function(oController, nX, nY) {
		this.controller.onClick(oController, nX, nY);
	}
	CControl.prototype.updateFromRanges = function(aRanges) {
		this.controller.updateFromRanges(aRanges);
	};
	CControl.prototype.getTextRect = function () {
		return this.controller.getTextRect();
	};
	CControl.prototype.canRotate = function () {
		return false;
	};
	CControl.prototype.setControlPr = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Control_ControlPr, this.controlPr, pr));
		this.controlPr = pr;
	};
	CControl.prototype.setFormControlPr = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Control_FormControlPr, this.formControlPr, pr));
		this.formControlPr = pr;
	};

	function CControlControllerBase(oControl) {
		this.control = oControl;
	}
	CControlControllerBase.prototype.getFormControlPr = function() {
		return this.control.formControlPr;
	};
	CControlControllerBase.prototype.getWorksheet = function() {
		return this.control.Get_Worksheet();
	};
	CControlControllerBase.prototype.draw = function(graphics, transform, transformText, pageIndex, opt) {};
	CControlControllerBase.prototype.onClick = function(oController, nX, nY) {};
	CControlControllerBase.prototype.init = function() {};
	CControlControllerBase.prototype.updateFromRanges = function(aRanges) {};
	CControlControllerBase.prototype.getBodyPr = function() {return null;};
	CControlControllerBase.prototype.getTextRect = function() {
		return AscFormat.CShape.prototype.getTextRect.call(this.control);
	};

	const CHECKBOX_SIDE_SIZE = 3;
	const CHECKBOX_X_OFFSET = 1.5;
	const CHECKBOX_BODYPR_INSETS_L = 27432 / 36000;
	const CHECKBOX_BODYPR_INSETS_R = 0;
	const CHECKBOX_BODYPR_INSETS_T = 32004 / 36000;
	const CHECKBOX_BODYPR_INSETS_B = 32004 / 36000;
	const CHECKBOX_OFFSET_X = CHECKBOX_SIDE_SIZE + (CHECKBOX_X_OFFSET * 2 - CHECKBOX_BODYPR_INSETS_L);
	function CCheckBoxController(oControl) {
		CControlControllerBase.call(this, oControl);
	};
	AscFormat.InitClassWithoutType(CCheckBoxController, CControlControllerBase);
	CCheckBoxController.prototype.getBodyPr = function () {
		const oBodyPr = new AscFormat.CBodyPr();
		oBodyPr.setInsets(CHECKBOX_BODYPR_INSETS_L, CHECKBOX_BODYPR_INSETS_T, CHECKBOX_BODYPR_INSETS_R, CHECKBOX_BODYPR_INSETS_B);
		oBodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
		oBodyPr.vertOverflow = AscFormat.nVOTClip;
		oBodyPr.wrap = AscFormat.nTWTSquare;
		oBodyPr.upright = true;
		return oBodyPr;
	};
	CCheckBoxController.prototype.draw = function(graphics, transform, transformText, pageIndex, opt) {
		const oControl = this.control;
		const oMainTransfrom = transform || oControl.transform;
		const checkBoxTransform = oMainTransfrom.CreateDublicate();
		AscFormat.CShape.prototype.draw.call(oControl, graphics, transform, transformText, pageIndex, opt);
		graphics.SaveGrState();
		graphics.AddClipRect(oControl.x, oControl.y, oControl.extX, oControl.extY);
		checkBoxTransform.tx += CHECKBOX_X_OFFSET;
		checkBoxTransform.ty += (oControl.extY - CHECKBOX_SIDE_SIZE) / 2;
		graphics.transform3(checkBoxTransform);
		graphics.b_color1(255, 255, 255, 255);
		graphics.p_color(0, 0, 0, 255);
		graphics.p_width(0);
		graphics._s();
		graphics._m(0, 0);
		graphics._l(0, CHECKBOX_SIDE_SIZE);
		graphics._l(CHECKBOX_SIDE_SIZE, CHECKBOX_SIDE_SIZE);
		graphics._l(CHECKBOX_SIDE_SIZE, 0);
		graphics._z();
		graphics.ds();
		graphics.df();
		graphics._e();
		if (this.isChecked()) {
			graphics.p_color(0, 0, 0, 255);
			graphics.p_width(400);
			graphics._m(2.5, 0.75);
			graphics._l(1, 2.25);
			graphics._l(0.5, 1.75);
			graphics.ds();
			graphics._e();
		} else if (this.isMixed()) {
			graphics.b_color1(0, 0, 0, 255);
			graphics._s();
			const nRectCount = 7;
			const nRectWidth = CHECKBOX_SIDE_SIZE / nRectCount;
			for (let i = 0; i < nRectCount; i += 1) {
				const nX = i * nRectWidth;
				for (let j = 0; j < nRectCount; j += 1) {
					const nY = j * nRectWidth;
					if ((i % 2) === (j % 2)) {
						graphics.TableRect(nX, nY, nRectWidth, nRectWidth);

					}
				}
			}
			graphics._e();
		}

		graphics.RestoreGrState();
	};
	CCheckBoxController.prototype.isChecked = function() {
		const oFormControlPr = this.getFormControlPr();
		return oFormControlPr.checked === CFormControlPr_checked_checked;
	};
	CCheckBoxController.prototype.isMixed = function() {
		const oFormControlPr = this.getFormControlPr();
		return oFormControlPr.checked === CFormControlPr_checked_mixed;
	};
	CCheckBoxController.prototype.isEmpty = function() {
		return !(this.isChecked() || this.isMixed());
	};
	CCheckBoxController.prototype.onClick = function(oController, nX, nY) {
		const oThis = this;
		oController.checkObjectsAndCallback(function() {
			const oFormControlPr = oThis.getFormControlPr();
			if (oThis.isMixed() || oThis.isChecked()) {
				oFormControlPr.setChecked(CFormControlPr_checked_unchecked);
			} else {
				oFormControlPr.setChecked(CFormControlPr_checked_checked);
			}
			oThis.updateCellFromControl(oController);
		}, [], false, AscDFH.historydescription_Spreadsheet_SwitchCheckbox, [this.control]);
	};
	CCheckBoxController.prototype.getParsedRef = function() {
		const oFormControlPr = this.getFormControlPr();
		if (oFormControlPr.fmlaLink) {
			const oWs = this.getWorksheet();
			let aParsedRef = AscCommonExcel.getRangeByRef(oFormControlPr.fmlaLink, oWs, true, true, true);
			const oRef = aParsedRef[0];
			if (oRef) {
				return new AscCommonExcel.Range(oRef.worksheet, oRef.bbox.r1, oRef.bbox.c1, oRef.bbox.r1, oRef.bbox.c1);
			}
		}
		return null;
	};
	CCheckBoxController.prototype.init = function() {
		const oFormControlPr = this.getFormControlPr();
		const oRef = this.getParsedRef();
		if (oRef) {
			oRef._foreachNoEmpty(function(oCell) {
				if (oCell) {
					const bValue = oCell.getBoolValue();
					if (oCell.type === AscCommon.CellValueType.Bool || oCell.type === AscCommon.CellValueType.Number) {
						oFormControlPr.setChecked(bValue ? CFormControlPr_checked_checked : CFormControlPr_checked_unchecked);
					} else if (oCell.type === AscCommon.CellValueType.Error) {
						oFormControlPr.setChecked(CFormControlPr_checked_mixed);
					}
				}
			});
		}
	};
	CCheckBoxController.prototype.updateCellFromControl = function (oController) {
		const oThis = this;
		const oRef = this.getParsedRef();
		if (oRef) {
			oRef._foreachNoEmpty(function(oCell) {
				if (oCell) {
					const oCellValue = new AscCommonExcel.CCellValue();
					if (oThis.isChecked()) {
						oCellValue.type = AscCommon.CellValueType.Bool;
						oCellValue.number = 1;
					} else if (oThis.isMixed()) {
						oCellValue.type = AscCommon.CellValueType.Error;
						oCellValue.text = AscCommonExcel.cError.prototype.getStringFromErrorType(cErrorType.not_available);
					} else {
						oCellValue.type = AscCommon.CellValueType.Bool;
						oCellValue.number = 0;
					}
					oCell.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));

				}
			});
			const oWb = Asc.editor.wb;
			const nWorksheetIndex = oRef.worksheet.getIndex();
			const oWs = oWb && oWb.getWorksheet(nWorksheetIndex, true);
			if (oWs) {
				oWs._updateRange(oRef.bbox);
				if (oWb.wsActive === nWorksheetIndex) {
					oWs.draw();
				}
			}
		}
	};
	CCheckBoxController.prototype.updateFromRanges = function(aRanges) {
		const oMainRange = this.getParsedRef();
		if (oMainRange) {
			for (let i = 0; i < aRanges.length; i += 1) {
				const oRange = aRanges[i];
				if (oRange.isIntersect(oMainRange)) {
					this.init();
				}
			}
		}
	};
	CCheckBoxController.prototype.getTextRect = function () {
		const oTextRect = AscFormat.CShape.prototype.getTextRect.call(this);
		oTextRect.l += CHECKBOX_OFFSET_X;
		oTextRect.r += CHECKBOX_OFFSET_X;
		return oTextRect;
	};



	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AltText] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AutoFill] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AutoLine] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AutoPict] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Dde] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_DefaultSize] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Disabled] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Cf] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_LinkedCell] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_ListFillRange] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_RId] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Locked] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Macro] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Print] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_RecalcAlways] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_UiObject] = AscDFH.CChangesDrawingsBool;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AltText] = function(oClass, value) {
		this.altText = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AutoFill] = function(oClass, value) {
		this.autoFill = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AutoLine] = function(oClass, value) {
		this.autoLine = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AutoPict] = function(oClass, value) {
		this.autoPict = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Dde] = function(oClass, value) {
		this.dde = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_DefaultSize] = function(oClass, value) {
		this.defaultSize = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Disabled] = function(oClass, value) {
		this.disabled = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Cf] = function(oClass, value) {
		this.cf = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_LinkedCell] = function(oClass, value) {
		this.linkedCell = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_ListFillRange] = function(oClass, value) {
		this.listFillRange = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_RId] = function(oClass, value) {
		this.rId = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Locked] = function(oClass, value) {
		this.locked = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Macro] = function(oClass, value) {
		this.macro = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Print] = function(oClass, value) {
		this.print = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_RecalcAlways] = function(oClass, value) {
		this.recalcAlways = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_UiObject] = function(oClass, value) {
		this.uiObject = value;
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
	}
	CControlPr.prototype.setAltText = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_AltText, this.altText, pr));
		this.altText = pr;
	}
	CControlPr.prototype.getAltText = function() {
		return this.altText;
	};
	CControlPr.prototype.setAutoFill = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_AutoFill, this.autoFill, pr));
		this.autoFill = pr;
	}
	CControlPr.prototype.getAutoFill = function() {
		return this.autoFill;
	};
	CControlPr.prototype.setAutoLine = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_AutoLine, this.autoLine, pr));
		this.autoLine = pr;
	}
	CControlPr.prototype.getAutoLine = function() {
		return this.autoLine;
	};
	CControlPr.prototype.setAutoPict = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_AutoPict, this.autoPict, pr));
		this.autoPict = pr;
	}
	CControlPr.prototype.getAutoPict = function() {
		return this.autoPict;
	};
	CControlPr.prototype.setDde = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Dde, this.dde, pr));
		this.dde = pr;
	}
	CControlPr.prototype.getDde = function() {
		return this.dde;
	};
	CControlPr.prototype.setDefaultSize = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_DefaultSize, this.defaultSize, pr));
		this.defaultSize = pr;
	}
	CControlPr.prototype.getDefaultSize = function() {
		return this.defaultSize;
	};
	CControlPr.prototype.setDisabled = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Disabled, this.disabled, pr));
		this.disabled = pr;
	}
	CControlPr.prototype.getDisabled = function() {
		return this.disabled;
	};
	CControlPr.prototype.setCf = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_Cf, this.cf, pr));
		this.cf = pr;
	}
	CControlPr.prototype.getCf = function() {
		return this.cf;
	};
	CControlPr.prototype.setLinkedCell = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_LinkedCell, this.linkedCell, pr));
		this.linkedCell = pr;
	}
	CControlPr.prototype.getLinkedCell = function() {
		return this.linkedCell;
	};
	CControlPr.prototype.setListFillRange = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_ListFillRange, this.listFillRange, pr));
		this.listFillRange = pr;
	}
	CControlPr.prototype.getListFillRange = function() {
		return this.listFillRange;
	};
	CControlPr.prototype.setRId = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_ControlPr_RId, this.rId, pr));
		this.rId = pr;
	}
	CControlPr.prototype.getRId = function() {
		return this.rId;
	};
	CControlPr.prototype.setLocked = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Locked, this.locked, pr));
		this.locked = pr;
	}
	CControlPr.prototype.getLocked = function() {
		return this.locked;
	};
	CControlPr.prototype.setMacro = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_Macro, this.macro, pr));
		this.macro = pr;
	}
	CControlPr.prototype.getMacro = function() {
		return this.macro;
	};
	CControlPr.prototype.setPrint = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Print, this.print, pr));
		this.print = pr;
	}
	CControlPr.prototype.getPrint = function() {
		return this.print;
	};
	CControlPr.prototype.setRecalcAlways = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_RecalcAlways, this.recalcAlways, pr));
		this.recalcAlways = pr;
	}
	CControlPr.prototype.getRecalcAlways = function() {
		return this.recalcAlways;
	};
	CControlPr.prototype.setUiObject = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_UiObject, this.uiObject, pr));
		this.uiObject = pr;
	};
	CControlPr.prototype.getUiObject = function() {
		return this.uiObject;
	};

	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_DropLines] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_ObjectType] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Checked] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_DropStyle] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Dx] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Inc] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Min] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Max] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Page] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Sel] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_SelType] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_TextHAlign] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_TextVAlign] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Val] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_WidthMin] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_EditVal] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaGroup] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaLink] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaRange] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaTxbx] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Colored] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FirstButton] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Horiz] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_JustLastX] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_LockText] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_MultiSel] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_NoThreeD] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_NoThreeD2] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_MultiLine] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_VerticalBar] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_PasswordEdit] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_AddItemToLst] = AscDFH.CChangesDrawingsContentString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_RemoveItemFromLst] = AscDFH.CChangesDrawingsContentString;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_DropLines] = function (oClass, value) {
		oClass.dropLines = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_ObjectType] = function (oClass, value) {
		oClass.objectType = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Checked] = function (oClass, value) {
		oClass.checked = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_DropStyle] = function (oClass, value) {
		oClass.dropStyle = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Dx] = function (oClass, value) {
		oClass.dx = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Inc] = function (oClass, value) {
		oClass.inc = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Min] = function (oClass, value) {
		oClass.min = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Max] = function (oClass, value) {
		oClass.max = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Page] = function (oClass, value) {
		oClass.page = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Sel] = function (oClass, value) {
		oClass.sel = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_SelType] = function (oClass, value) {
		oClass.selType = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_TextHAlign] = function (oClass, value) {
		oClass.textHAlign = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_TextVAlign] = function (oClass, value) {
		oClass.textVAlign = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Val] = function (oClass, value) {
		oClass.val = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_WidthMin] = function (oClass, value) {
		oClass.widthMin = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_EditVal] = function (oClass, value) {
		oClass.editVal = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaGroup] = function (oClass, value) {
		oClass.fmlaGroup = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaLink] = function (oClass, value) {
		oClass.fmlaLink = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaRange] = function (oClass, value) {
		oClass.fmlaRange = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaTxbx] = function (oClass, value) {
		oClass.fmlaTxbx = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Colored] = function (oClass, value) {
		oClass.colored = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FirstButton] = function (oClass, value) {
		oClass.firstButton = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Horiz] = function (oClass, value) {
		oClass.horiz = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_JustLastX] = function (oClass, value) {
		oClass.justLastX = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_LockText] = function (oClass, value) {
		oClass.lockText = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_MultiSel] = function (oClass, value) {
		oClass.multiSel = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_NoThreeD] = function (oClass, value) {
		oClass.noThreeD = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_NoThreeD2] = function (oClass, value) {
		oClass.noThreeD2 = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_MultiLine] = function (oClass, value) {
		oClass.multiLine = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_VerticalBar] = function (oClass, value) {
		oClass.verticalBar = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_PasswordEdit] = function (oClass, value) {
		oClass.passwordEdit = value;
	};
	AscDFH.drawingContentChanges[AscDFH.historyitem_FormControlPr_AddItemToLst] = function (oClass) {
		return oClass.itemLst;
	};
	AscDFH.drawingContentChanges[AscDFH.historyitem_FormControlPr_RemoveItemFromLst] = function (oClass) {
		return oClass.itemLst;
	};
	function CFormControlPr() {
		AscFormat.CBaseFormatObject.call(this);
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
	AscFormat.InitClass(CFormControlPr, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_FormControlPr);
	CFormControlPr.prototype.setDropLines = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_DropLines, this.dropLines, pr));
		this.dropLines = pr;
	}
	CFormControlPr.prototype.getDropLines = function() {
		return this.dropLines;
	}
	CFormControlPr.prototype.setObjectType = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_ObjectType, this.objectType, pr));
		this.objectType = pr;
	}
	CFormControlPr.prototype.getObjectType = function() {
		return this.objectType;
	}
	CFormControlPr.prototype.setChecked = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Checked, this.checked, pr));
		this.checked = pr;
	}
	CFormControlPr.prototype.getChecked = function() {
		return this.checked;
	}
	CFormControlPr.prototype.setDropStyle = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_DropStyle, this.dropStyle, pr));
		this.dropStyle = pr;
	}
	CFormControlPr.prototype.getDropStyle = function() {
		return this.dropStyle;
	}
	CFormControlPr.prototype.setDx = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Dx, this.dx, pr));
		this.dx = pr;
	}
	CFormControlPr.prototype.getDx = function() {
		return this.dx;
	}
	CFormControlPr.prototype.setInc = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Inc, this.inc, pr));
		this.inc = pr;
	}
	CFormControlPr.prototype.getInc = function() {
		return this.inc;
	}
	CFormControlPr.prototype.setMin = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Min, this.min, pr));
		this.min = pr;
	}
	CFormControlPr.prototype.getMin = function() {
		return this.min;
	}
	CFormControlPr.prototype.setMax = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Max, this.max, pr));
		this.max = pr;
	}
	CFormControlPr.prototype.getMax = function() {
		return this.max;
	}
	CFormControlPr.prototype.setPage = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Page, this.page, pr));
		this.page = pr;
	}
	CFormControlPr.prototype.getPage = function() {
		return this.page;
	}
	CFormControlPr.prototype.setSel = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Sel, this.sel, pr));
		this.sel = pr;
	}
	CFormControlPr.prototype.getSel = function() {
		return this.sel;
	}
	CFormControlPr.prototype.setSelType = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_SelType, this.selType, pr));
		this.selType = pr;
	}
	CFormControlPr.prototype.getSelType = function() {
		return this.selType;
	}
	CFormControlPr.prototype.setTextHAlign = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_TextHAlign, this.textHAlign, pr));
		this.textHAlign = pr;
	}
	CFormControlPr.prototype.getTextHAlign = function() {
		return this.textHAlign;
	}
	CFormControlPr.prototype.setTextVAlign = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_TextVAlign, this.textVAlign, pr));
		this.textVAlign = pr;
	}
	CFormControlPr.prototype.getTextVAlign = function() {
		return this.textVAlign;
	}
	CFormControlPr.prototype.setVal = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Val, this.val, pr));
		this.val = pr;
	}
	CFormControlPr.prototype.getVal = function() {
		return this.val;
	}
	CFormControlPr.prototype.setWidthMin = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_WidthMin, this.widthMin, pr));
		this.widthMin = pr;
	}
	CFormControlPr.prototype.getWidthMin = function() {
		return this.widthMin;
	}
	CFormControlPr.prototype.setEditVal = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_EditVal, this.editVal, pr));
		this.editVal = pr;
	}
	CFormControlPr.prototype.getEditVal = function() {
		return this.editVal;
	}
	CFormControlPr.prototype.setFmlaGroup = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaGroup, this.fmlaGroup, pr));
		this.fmlaGroup = pr;
	}
	CFormControlPr.prototype.getFmlaGroup = function() {
		return this.fmlaGroup;
	}
	CFormControlPr.prototype.setFmlaLink = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaLink, this.fmlaLink, pr));
		this.fmlaLink = pr;
	}
	CFormControlPr.prototype.getFmlaLink = function() {
		return this.fmlaLink;
	}
	CFormControlPr.prototype.setFmlaRange = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaRange, this.fmlaRange, pr));
		this.fmlaRange = pr;
	}
	CFormControlPr.prototype.getFmlaRange = function() {
		return this.fmlaRange;
	}
	CFormControlPr.prototype.setFmlaTxbx = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaTxbx, this.fmlaTxbx, pr));
		this.fmlaTxbx = pr;
	}
	CFormControlPr.prototype.getFmlaTxbx = function() {
		return this.fmlaTxbx;
	}
	CFormControlPr.prototype.setColored = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_Colored, this.colored, pr));
		this.colored = pr;
	}
	CFormControlPr.prototype.getColored = function() {
		return this.colored;
	}
	CFormControlPr.prototype.setFirstButton = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_FirstButton, this.firstButton, pr));
		this.firstButton = pr;
	}
	CFormControlPr.prototype.getFirstButton = function() {
		return this.firstButton;
	}
	CFormControlPr.prototype.setHoriz = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_Horiz, this.horiz, pr));
		this.horiz = pr;
	}
	CFormControlPr.prototype.getHoriz = function() {
		return this.horiz;
	}
	CFormControlPr.prototype.setJustLastX = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_JustLastX, this.justLastX, pr));
		this.justLastX = pr;
	}
	CFormControlPr.prototype.getJustLastX = function() {
		return this.justLastX;
	}
	CFormControlPr.prototype.setLockText = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_LockText, this.lockText, pr));
		this.lockText = pr;
	}
	CFormControlPr.prototype.getLockText = function() {
		return this.lockText;
	}
	CFormControlPr.prototype.setMultiSel = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_MultiSel, this.multiSel, pr));
		this.multiSel = pr;
	}
	CFormControlPr.prototype.getMultiSel = function() {
		return this.multiSel;
	}
	CFormControlPr.prototype.setNoThreeD = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_NoThreeD, this.noThreeD, pr));
		this.noThreeD = pr;
	}
	CFormControlPr.prototype.getNoThreeD = function() {
		return this.noThreeD;
	}
	CFormControlPr.prototype.setNoThreeD2 = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_NoThreeD2, this.noThreeD2, pr));
		this.noThreeD2 = pr;
	}
	CFormControlPr.prototype.getNoThreeD2 = function() {
		return this.noThreeD2;
	}
	CFormControlPr.prototype.setMultiLine = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_MultiLine, this.multiLine, pr));
		this.multiLine = pr;
	}
	CFormControlPr.prototype.getMultiLine = function() {
		return this.multiLine;
	}
	CFormControlPr.prototype.setVerticalBar = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_VerticalBar, this.verticalBar, pr));
		this.verticalBar = pr;
	}
	CFormControlPr.prototype.getVerticalBar = function() {
		return this.verticalBar;
	}
	CFormControlPr.prototype.setPasswordEdit = function(pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_PasswordEdit, this.passwordEdit, pr));
		this.passwordEdit = pr;
	}
	CFormControlPr.prototype.getPasswordEdit = function() {
		return this.passwordEdit;
	}
	CFormControlPr.prototype.addItemToLst = function (nIdx, sPr) {
		var nInsertIdx = Math.min(this.itemLst.length, Math.max(0, nIdx));
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsContentString(this, AscDFH.historyitem_FormControlPr_AddItemToLst, nInsertIdx, [sPr], true));
		this.itemLst.splice(nInsertIdx, 0, sPr);
	};
	CFormControlPr.prototype.removeItemFromLst = function (nIdx) {
		if (nIdx > -1 && nIdx < this.itemLst.length) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsContentString(this, AscDFH.historyitem_CCommonDataListRemove, nIdx, [this.itemLst[nIdx]], false));
			this.itemLst.splice(nIdx, 1);
		}
	};

	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CControl = CControl;
})();
