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
	CControl.prototype.hitInTextRect = function() {
		//todo
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

	function CCheckBoxController(oControl) {
		CControlControllerBase.call(this, oControl);
	};
	AscFormat.InitClassWithoutType(CCheckBoxController, CControlControllerBase);
	CCheckBoxController.prototype.draw = function(graphics, transform, transformText, pageIndex, opt) {
		const oControl = this.control;
		const nSide = 3;
		const nXOffset = 1.5;
		const oMainTransfrom = transform || oControl.transform;
		const checkBoxTransform = oMainTransfrom.CreateDublicate();
		const oMainTextTransform = transformText || oControl.transformText;
		const oCheckBoxTextTransform = oMainTextTransform.CreateDublicate();
		oCheckBoxTextTransform.tx += nSide + nXOffset * 2;
		AscFormat.CShape.prototype.draw.call(oControl, graphics, transform, oCheckBoxTextTransform, pageIndex, opt);
		graphics.SaveGrState();

		checkBoxTransform.tx += nXOffset;
		checkBoxTransform.ty += (oControl.extY - nSide) / 2;
		graphics.transform3(checkBoxTransform);
		graphics.b_color1(255, 255, 255, 255);
		graphics.p_color(0, 0, 0, 255);
		graphics.p_width(0);
		graphics._s();
		graphics._m(0, 0);
		graphics._l(0, nSide);
		graphics._l(nSide, nSide);
		graphics._l(nSide, 0);
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
			const nRectWidth = nSide / nRectCount;
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
	}


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
	CFormControlPr.prototype.setDropLines = function(pr) {
		this.dropLines = pr;
	}
	CFormControlPr.prototype.getDropLines = function() {
		return this.dropLines;
	}
	CFormControlPr.prototype.setObjectType = function(pr) {
		this.objectType = pr;
	}
	CFormControlPr.prototype.getObjectType = function() {
		return this.objectType;
	}
	CFormControlPr.prototype.setChecked = function(pr) {
		this.checked = pr;
	}
	CFormControlPr.prototype.getChecked = function() {
		return this.checked;
	}
	CFormControlPr.prototype.setDropStyle = function(pr) {
		this.dropStyle = pr;
	}
	CFormControlPr.prototype.getDropStyle = function() {
		return this.dropStyle;
	}
	CFormControlPr.prototype.setDx = function(pr) {
		this.dx = pr;
	}
	CFormControlPr.prototype.getDx = function() {
		return this.dx;
	}
	CFormControlPr.prototype.setInc = function(pr) {
		this.inc = pr;
	}
	CFormControlPr.prototype.getInc = function() {
		return this.inc;
	}
	CFormControlPr.prototype.setMin = function(pr) {
		this.min = pr;
	}
	CFormControlPr.prototype.getMin = function() {
		return this.min;
	}
	CFormControlPr.prototype.setMax = function(pr) {
		this.max = pr;
	}
	CFormControlPr.prototype.getMax = function() {
		return this.max;
	}
	CFormControlPr.prototype.setPage = function(pr) {
		this.page = pr;
	}
	CFormControlPr.prototype.getPage = function() {
		return this.page;
	}
	CFormControlPr.prototype.setSel = function(pr) {
		this.sel = pr;
	}
	CFormControlPr.prototype.getSel = function() {
		return this.sel;
	}
	CFormControlPr.prototype.setSelType = function(pr) {
		this.selType = pr;
	}
	CFormControlPr.prototype.getSelType = function() {
		return this.selType;
	}
	CFormControlPr.prototype.setTextHAlign = function(pr) {
		this.textHAlign = pr;
	}
	CFormControlPr.prototype.getTextHAlign = function() {
		return this.textHAlign;
	}
	CFormControlPr.prototype.setTextVAlign = function(pr) {
		this.textVAlign = pr;
	}
	CFormControlPr.prototype.getTextVAlign = function() {
		return this.textVAlign;
	}
	CFormControlPr.prototype.setVal = function(pr) {
		this.val = pr;
	}
	CFormControlPr.prototype.getVal = function() {
		return this.val;
	}
	CFormControlPr.prototype.setWidthMin = function(pr) {
		this.widthMin = pr;
	}
	CFormControlPr.prototype.getWidthMin = function() {
		return this.widthMin;
	}
	CFormControlPr.prototype.setEditVal = function(pr) {
		this.editVal = pr;
	}
	CFormControlPr.prototype.getEditVal = function() {
		return this.editVal;
	}
	CFormControlPr.prototype.setFmlaGroup = function(pr) {
		this.fmlaGroup = pr;
	}
	CFormControlPr.prototype.getFmlaGroup = function() {
		return this.fmlaGroup;
	}
	CFormControlPr.prototype.setFmlaLink = function(pr) {
		this.fmlaLink = pr;
	}
	CFormControlPr.prototype.getFmlaLink = function() {
		return this.fmlaLink;
	}
	CFormControlPr.prototype.setFmlaRange = function(pr) {
		this.fmlaRange = pr;
	}
	CFormControlPr.prototype.getFmlaRange = function() {
		return this.fmlaRange;
	}
	CFormControlPr.prototype.setFmlaTxbx = function(pr) {
		this.fmlaTxbx = pr;
	}
	CFormControlPr.prototype.getFmlaTxbx = function() {
		return this.fmlaTxbx;
	}
	CFormControlPr.prototype.setColored = function(pr) {
		this.colored = pr;
	}
	CFormControlPr.prototype.getColored = function() {
		return this.colored;
	}
	CFormControlPr.prototype.setFirstButton = function(pr) {
		this.firstButton = pr;
	}
	CFormControlPr.prototype.getFirstButton = function() {
		return this.firstButton;
	}
	CFormControlPr.prototype.setHoriz = function(pr) {
		this.horiz = pr;
	}
	CFormControlPr.prototype.getHoriz = function() {
		return this.horiz;
	}
	CFormControlPr.prototype.setJustLastX = function(pr) {
		this.justLastX = pr;
	}
	CFormControlPr.prototype.getJustLastX = function() {
		return this.justLastX;
	}
	CFormControlPr.prototype.setLockText = function(pr) {
		this.lockText = pr;
	}
	CFormControlPr.prototype.getLockText = function() {
		return this.lockText;
	}
	CFormControlPr.prototype.setMultiSel = function(pr) {
		this.multiSel = pr;
	}
	CFormControlPr.prototype.getMultiSel = function() {
		return this.multiSel;
	}
	CFormControlPr.prototype.setNoThreeD = function(pr) {
		this.noThreeD = pr;
	}
	CFormControlPr.prototype.getNoThreeD = function() {
		return this.noThreeD;
	}
	CFormControlPr.prototype.setNoThreeD2 = function(pr) {
		this.noThreeD2 = pr;
	}
	CFormControlPr.prototype.getNoThreeD2 = function() {
		return this.noThreeD2;
	}
	CFormControlPr.prototype.setMultiLine = function(pr) {
		this.multiLine = pr;
	}
	CFormControlPr.prototype.getMultiLine = function() {
		return this.multiLine;
	}
	CFormControlPr.prototype.setVerticalBar = function(pr) {
		this.verticalBar = pr;
	}
	CFormControlPr.prototype.getVerticalBar = function() {
		return this.verticalBar;
	}
	CFormControlPr.prototype.setPasswordEdit = function(pr) {
		this.passwordEdit = pr;
	}
	CFormControlPr.prototype.getPasswordEdit = function() {
		return this.passwordEdit;
	}


	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CControl = CControl;
})();
