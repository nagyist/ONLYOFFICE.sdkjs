/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(){

    //------------------------------------------------------------------------------------------------------------------
	//
	// Internal
	//
	//------------------------------------------------------------------------------------------------------------------

    let FIELDS_HIGHLIGHT = {
        r: 201, 
        g: 200,
        b: 255
    }
    let BUTTON_PRESSED = {
        r: 153,
        g: 193,
        b: 218
    }
    
    let FIELD_TYPE = {
        button:         "button",
        checkbox:       "checkbox",
        combobox:       "combobox",
        listbox:        "listbox",
        radiobutton:    "radiobutton",
        signature:      "signature",
        text:           "text"
    }

    let ACTION_TRIGGER_TYPES = {
        MouseUp:    0,
        MouseDown:  1,
        MouseEnter: 2,
        MouseExit:  3,
        OnFocus:    4,
        OnBlur:     5,
        Keystroke:  6,
        Validate:   7,
        Calculate:  8,
        Format:     9
    }
    //------------------------------------------------------------------------------------------------------------------
	//
	// pdf api types
	//
	//------------------------------------------------------------------------------------------------------------------
    
    let ALIGN_TYPE = {
        left:   "left",
        center: "center",
        right:  "right"
    }

    let border = {
        "s": "solid",
        "b": "beveled",
        "d": "dashed",
        "i": "inset",
        "u": "underline"
    }

    let position = {
        "textOnly":   0,
        "iconOnly":   1,
        "iconTextV":  2,
        "textIconV":  3,
        "iconTextH":  4,
        "textIconH":  5,
        "overlay":    6
    }

    let scaleHow = {
        "proportional":   0,
        "anamorphic":     1
    }

    let scaleWhen = {
        "always":     0,
        "never":      1,
        "tooBig":     2,
        "tooSmall":   3
    }

    const CHAR_LIM_MAX = 500; // to do проверить

    let display = {
        "visible":  0,
        "hidden":   1,
        "noPrint":  2,
        "noView":   3
    }

    // For Span attributes (start)
    let FONT_STRETCH = ["ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "normal",
        "semi-expanded", "expanded", "extra-expanded", "ultra-expanded"];

    let FONT_STYLE = {
        italic: "italic",
        normal: "normal"
    }

    let FONT_WEIGHT = [100, 200, 300, 400, 500, 600, 700, 800, 900];

    // for CSpan (end)

    
    // default availible colors
    let color = {
        "transparent":  [ "T" ],
        "black":        [ "G", 0 ],
        "white":        [ "G", 1 ],
        "red":          [ "RGB", 1,0,0 ],
        "green":        [ "RGB", 0,1,0 ],
        "blue":         [ "RGB", 0, 0, 1 ],
        "cyan":         [ "CMYK", 1,0,0,0 ],
        "magenta":      [ "CMYK", 0,1,0,0 ],
        "yellow":       [ "CMYK", 0,0,1,0 ],
        "dkGray":       [ "G", 0.25 ],  // version 4.0
        "gray":         [ "G", 0.5 ],   // version 4.0
        "ltGray":       [ "G", 0.75 ]   // version 4.0
    }

    // please use copy of this object
    let DEFAULT_SPAN = {
        "alignment":        ALIGN_TYPE.left,
        "fontFamily":       ["sans-serif"],
        "fontStretch":      "normal",
        "fontStyle":        "normal",
        "fontWeight":       400,
        "strikethrough":    false,
        "subscript":        false,
        "superscript":      false,
        "text":             "",
        "color":            color["black"],
        "textSize":         12.0,
        "underline":        false
    }

    // Defines how a button reacts when a user clicks it.
    // The four highlight modes supported are:
    let highlight = {
        "n": "none",
        "i": "invert",
        "p": "push",
        "o": "outline"
    }
    
    let LINE_WIDTH = {
        "none":   0,
        "thin":   1,
        "medium": 2,
        "thick":  3
    }

    let VALID_ROTATIONS = [0, 90, 180, 270];

    // Allows the user to set the glyph style of a check box or radio button.
    // The glyph style is the graphic used to indicate that the item has been selected.
    let style = {
        "ch": "check",
        "cr": "cross",
        "di": "diamond",
        "ci": "circle",
        "st": "star",
        "sq": "square"
    }

    const MAX_TEXT_SIZE = 32767;

    // freeze objects
    Object.freeze(FIELDS_HIGHLIGHT);
    Object.freeze(FIELD_TYPE);
    Object.freeze(ALIGN_TYPE);
    Object.freeze(border);
    Object.freeze(position);
    Object.freeze(scaleHow);
    Object.freeze(scaleWhen);
    Object.freeze(FONT_STRETCH);
    Object.freeze(FONT_STYLE);
    Object.freeze(FONT_WEIGHT);
    Object.freeze(color);
    Object.freeze(highlight);
    Object.freeze(VALID_ROTATIONS);
    Object.freeze(style);
    Object.freeze(ACTION_TRIGGER_TYPES);

    // base form class with attributes and method for all types of forms
	function CBaseField(sName, sType, nPage, aRect)
    {
        this.type = sType;
        
        this._borderStyle   = border.s;
        this._delay         = false;
        this._display       = display["visible"];
        this._doc           = null;
        this._fillColor     = ["RGB",1,1,1];
        this._bgColor       = undefined;          // prop for old versions (fillColor)
        this._hidden        = false;             // This property has been superseded by the display property and its use is discouraged.
        this._lineWidth     = LINE_WIDTH.thin;  // In older versions of this specification, this property was borderWidth
        this._borderWidth   = undefined;       
        this._name          = sName;         // to do
        this._page          = nPage;        // integer | array
        this._print         = true;        // This property has been superseded by the display property and its use is discouraged.
        this._readonly      = false;
        this._rect          = aRect;
        this._required      = false;       // for all except button
        this._rotation      = 0;
        this._strokeColor   = ["T"];     // In older versions of this specification, this property was borderColor. The use of borderColor is now discouraged,
                                        // although it is still valid for backward compatibility.
        this._borderColor   = undefined;
        this._submitName    = "";
        this._textColor     = ["RGB",0,0,0];
        this._fgColor       = undefined;
        this._textSize      = 10; // 0 == max text size // to do
        this._userName      = ""; // It is intended to be used as tooltip text whenever the cursor enters a field. 
        //It can also be used as a user-friendly name, instead of the field name, when generating error messages.

        this._actions = new CFormActions();

        // internal
        this._formRelRectMM = {
            X: 0,
            Y: 0,
            W: 0,
            H: 0,
            Page: nPage
        }
        this._oldContentPos = {X: 0, Y: 0, XLimit: 0, YLimit: 0};
        this._curShiftView = {
            x: 0,
            y: 0
        }

        this._needDrawHighlight = true;
        this._wasChanged = false; // было ло изменено содержимое формы

        private_getViewer().ImageMap = {};
        private_getViewer().InitDocument = function() {return};

        this._partialName = sName;
        this._apiForm = this.GetApiForm();
    }
    
    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.GetField = function(sName) {
        for (let i = 0; i < this._kids.length; i++) {
            if (this._kids[i]._partialName == sName)
                return this._kids[i];
        }

        return null;
    };

    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.AddKid = function(oField) {
        this._kids.push(oField);
        oField._parent = this;
    };
    /**
	 * Removes field from kids.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @param {CBaseField} oField - the field to remove.
	 * @returns {boolean} - returns false if field isn't in the field kids.
	 */
    CBaseField.prototype.RemoveKid = function(oField) {
        let nIndex = this._kids.indexOf(oField);
        if (nIndex != -1) {
            this._kids.splice(nIndex, 1);
            oField._parent = null;
            return true;
        }

        return false;
    };

    CBaseField.prototype.getFormRelRect = function()
    {
        return this._formRelRectMM;
    };
    CBaseField.prototype.IntersectWithRect = function(X, Y, W, H, nPageAbs)
    {
        var arrRects = [];
        
        var oBounds = this.getFormRelRect();

        if (nPageAbs === oBounds.Page)
        {
            var nLeft   = Math.max(X, oBounds.X);
            var nRight  = Math.min(X + W, oBounds.X + oBounds.W);
            var nTop    = Math.max(Y, oBounds.Y);
            var nBottom = Math.min(Y + H, oBounds.Y + oBounds.H);

            if (nLeft < nRight && nTop < nBottom)
            {
                arrRects.push({
                    X : nLeft,
                    Y : nTop,
                    W : nRight - nLeft,
                    H : nBottom - nTop
                });
            }
        }

        return arrRects;
    };

    /**
	 * Sets the JavaScript action of the field for a given trigger.
     * Note: This method will overwrite any action already defined for the chosen trigger.
	 * @memberof CBaseField
     * @param {cTrigger} cTrigger - A string that sets the trigger for the action.
     * @param {string} cScript - The JavaScript code to be executed when the trigger is activated.
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.setAction = function(cTrigger, cScript) {
        switch (cTrigger) {
            case "MouseUp":
                ACTION_TRIGGER_TYPES.MouseUp;
                break;
            case "MouseDown":
                ACTION_TRIGGER_TYPES.MouseDown;
                break;
            case "MouseEnter":
                ACTION_TRIGGER_TYPES.MouseEnter;
                break;
            case "MouseExit":
                ACTION_TRIGGER_TYPES.MouseExit;
                break;
            case "OnFocus":
                ACTION_TRIGGER_TYPES.OnFocus;
                break;
            case "OnBlur":
                ACTION_TRIGGER_TYPES.OnBlur;
                break;
            case "Keystroke":
                this._actions.Keystroke = new CFormAction(ACTION_TRIGGER_TYPES.Keystroke, cScript);
                break;
            case "Validate":
                ACTION_TRIGGER_TYPES.Validate;
                break;
            case "Calculate":
                ACTION_TRIGGER_TYPES.Calculate;
                break;
            case "Format":
                this._actions.Format = new CFormAction(ACTION_TRIGGER_TYPES.Format, cScript);
                break;
        }
    };

    CBaseField.prototype.DrawHighlight = function(oCtx) {
        oCtx = private_getViewer().canvasFormsHighlight.getContext("2d");
        if (this.type == "button" && this._buttonPressed) {
            oCtx.fillStyle = `rgb(${BUTTON_PRESSED.r}, ${BUTTON_PRESSED.g}, ${BUTTON_PRESSED.b})`;
            oCtx.fillRect(this._pagePos.realX, this._pagePos.realY, this._pagePos.w, this._pagePos.h);
        }
        else {
            oCtx.fillStyle = `rgb(${FIELDS_HIGHLIGHT.r}, ${FIELDS_HIGHLIGHT.g}, ${FIELDS_HIGHLIGHT.b})`;
            oCtx.fillRect(this._pagePos.realX, this._pagePos.realY, this._pagePos.w, this._pagePos.h);
        }
    };

    CBaseField.prototype.DrawBorders = function(oCtx, X, Y, nWidth, nHeight) {
        let oViewer = private_getViewer();
        let nLineWidth = 1 * oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio * this._lineWidth;
        oCtx.lineWidth = nLineWidth;
        if (nLineWidth == 0) {
            return;
        }

        switch (this._borderStyle) {
            case "solid":
                oCtx.setLineDash([]);
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                break;
            case "beveled":
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                oCtx.closePath();

                // bottom part
                oCtx.beginPath();
                oCtx.moveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                
                oCtx.moveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);

                oCtx.fillStyle = "gray";
                oCtx.closePath();
                oCtx.fill();

                // right part
                oCtx.beginPath();
                oCtx.moveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                
                oCtx.moveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);

                oCtx.fillStyle = "gray";
                oCtx.closePath();
                oCtx.fill();

                break;
            case "dashed":
                oCtx.setLineDash([5 * oViewer.zoom]);
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                break;
            case "inset":
                oCtx.setLineDash([]);
                oCtx.beginPath();
                oCtx.rect(X, Y, nWidth, nHeight);
                oCtx.stroke();
                oCtx.closePath();

                // left part
                oCtx.beginPath();
                oCtx.moveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                
                oCtx.moveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);

                oCtx.fillStyle = "gray";
                oCtx.closePath();
                oCtx.fill();

                // top part
                oCtx.beginPath();
                oCtx.moveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                
                oCtx.moveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);

                oCtx.fillStyle = "gray";
                oCtx.closePath();
                oCtx.fill();

                // bottom part
                oCtx.beginPath();
                oCtx.moveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                oCtx.lineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                
                oCtx.moveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);

                oCtx.fillStyle = "rgb(191, 191, 191)";
                oCtx.closePath();
                oCtx.fill();

                // right part
                oCtx.beginPath();
                oCtx.moveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                
                oCtx.moveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                oCtx.lineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                oCtx.lineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);

                oCtx.fillStyle = "rgb(191, 191, 191)";
                oCtx.closePath();
                oCtx.fill();

                break;
            case "underline":
                oCtx.setLineDash([]);
                oCtx.beginPath();
                oCtx.moveTo(X, Y + nHeight);
                oCtx.lineTo(X + nWidth, Y + nHeight);
                oCtx.stroke();
                break;
        }

        // draw comb cells
        if ((this._borderStyle == "solid" || this._borderStyle == "dashed") && (this.type == "text" && this._comb == true)) {
            let nCombWidth = nWidth / this._charLimit;
            let nIndentX = nCombWidth;

            for (let i = 0; i < this._charLimit - 1; i++) {
                oCtx.moveTo(X + nIndentX, Y);
                oCtx.lineTo(X + nIndentX, Y + nHeight);
                oCtx.stroke();
                nIndentX += nCombWidth;
            }
        }
    };

    CBaseField.prototype.GetType = function() {
        return this.type;
    };

    CBaseField.prototype.SetReadOnly = function(bReadOnly) {
        this._readonly = bReadOnly;
    };
    
    CBaseField.prototype.SetRequired = function(bRequired) {
        if (this.type != "button")
            this._required = bRequired;
    };

    /**
	 * Gets Api class for this form.
	 * @memberof CTextField
     * @param {number} nIdx - The 0-based index of the item in the list or -1 for the last item in the list.
     * @param {boolean} [bExportValue=true] - Specifies whether to return an export value.
	 * @typeofeditors ["PDF"]
     * @returns {ApiBaseField}
	 */
    CBaseField.prototype.GetApiForm = function() {
        if (this._apiForm)
            return this._apiForm;

        switch (this.type) {
            case "text":
                return new AscPDFEditor.ApiTextField(this);
            case "combobox":
                return new AscPDFEditor.ApiComboBoxField(this);
            case "listbox":
                return new AscPDFEditor.ApiListBoxField(this);
            case "checkbox":
                return new AscPDFEditor.ApiCheckBoxField(this);
            case "radiobutton":
                return new AscPDFEditor.ApiRadioButtonField(this);
            case "button":
                return new AscPDFEditor.ApiPushButtonField(this);
        }
    };

          
    function CPushButtonField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.button, nPage, aRect);

        this._buttonAlignX = 50; // must be integer
        this._buttonAlignY = 50; // must be integer
        this._buttonFitBounds = undefined;
        this._buttonPosition = {
            _textOnly:   undefined,
            _iconOnly:   undefined,
            _iconTextV:  undefined,
            _textIconV:  undefined,
            _iconTextH:  undefined,
            _textIconH:  undefined,
            _overlay:    undefined
        };
        this._buttonScaleHow    = undefined;
        this._highlight         = highlight["p"];
        this._textFont          = "ArialMT";
        
        this._buttonPressed = false;
        // internal
        TurnOffHistory();
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);

    }
    CPushButtonField.prototype = Object.create(CBaseField.prototype);
	CPushButtonField.prototype.constructor = CPushButtonField;

    CPushButtonField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = private_getViewer();

        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        this.DrawBorders(oCtx, X, Y, nWidth, nHeight);
        let oMargins = this.GetBordersWidth();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02 + oMargins.left) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98 - oMargins.right) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm / scaleCoef;
        
        contentY = Y * g_dKoef_pix_to_mm / scaleCoef;

        this._formRelRectMM.X = contentX;
        this._formRelRectMM.Y = contentY;
        this._formRelRectMM.W = contentXLimit - contentX;
        this._formRelRectMM.H = contentYLimit - contentY;

        // подгоняем размер галочки
        //this.ProcessAutoFitContent();

        // выставляем текст посередине
        // let nContentH = this._content.GetElement(0).Get_EmptyHeight();
        // contentY = contentY + ((contentYLimit - contentY) - nContentH) / 2;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
        }
        else if (this._wasChanged) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this._wasChanged = false;
        }

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        oGraphics.init(oCtx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && (oViewer.fieldFillingMode || this.type == "combobox"))
            this._content.RecalculateCurPos();
        
        oGraphics.RemoveClip();
        this._pageIndX = pageIndX;
        this._pageIndY = pageIndY;
        
        // save pos in page.
        this._pagePos = {
            x: X - pageIndX,
            y: Y - pageIndY,
            w: nWidth,
            h: nHeight,
            realX: X,
            realY: Y
        };
    };
    CPushButtonField.prototype.onMouseDown = function() {
        let oViewer = private_getViewer();
        this._buttonPressed = true;
        this._needDrawHighlight = true;

        oViewer._paintFormsHighlight();
    };
    CPushButtonField.prototype.onMouseUp = function() {
        let oViewer = private_getViewer();
        this._buttonPressed = false;
        this._needDrawHighlight = true;

        this.buttonImportIcon();
        oViewer._paintFormsHighlight();
    };
    CPushButtonField.prototype.buttonImportIcon = function() {
        this._apiForm.buttonImportIcon();
    };


    function CBaseCheckBoxField(sName, sType, nPage, aRect)
    {
        CBaseField.call(this, sName, sType, nPage, aRect);

        this._exportValues  = ["Yes"];
        this._value         = "Off";
        this._exportValue   = "Yes";
        
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        
        let oPara = this._content.GetElement(0);
        oPara.Recalculate_Page(0);
        oPara.SetParagraphAlign(align_Center);
        oPara.CompiledPr.NeedRecalc = false;
        oPara.CompiledPr.Pr.ParaPr.Jc = align_Center;
    }
    
    CBaseCheckBoxField.prototype = Object.create(CBaseField.prototype);
	CBaseCheckBoxField.prototype.constructor = CBaseCheckBoxField;

    CBaseCheckBoxField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = private_getViewer();

        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        if (this.type == "checkbox")
            this.DrawBorders(oCtx, X, Y, nWidth, nHeight);
        let oMargins = this.GetBordersWidth();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02 + oMargins.left) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98 - oMargins.right) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm / scaleCoef;
        
        contentY = Y * g_dKoef_pix_to_mm / scaleCoef;

        this._formRelRectMM.X = contentX;
        this._formRelRectMM.Y = contentY;
        this._formRelRectMM.W = contentXLimit - contentX;
        this._formRelRectMM.H = contentYLimit - contentY;

        // подгоняем размер галочки
        this.ProcessAutoFitContent();

        // выставляем текст посередине
        // let nContentH = this._content.GetElement(0).Get_EmptyHeight();
        // contentY = contentY + ((contentYLimit - contentY) - nContentH) / 2;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
        }
        else if (this._wasChanged) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this._wasChanged = false;
        }

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        oGraphics.init(oCtx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && (oViewer.fieldFillingMode || this.type == "combobox"))
            this._content.RecalculateCurPos();
        
        oGraphics.RemoveClip();
        this._pageIndX = pageIndX;
        this._pageIndY = pageIndY;
        
        // save pos in page.
        this._pagePos = {
            x: X - pageIndX,
            y: Y - pageIndY,
            w: nWidth,
            h: nHeight,
            realX: X,
            realY: Y
        };
    };

    CBaseCheckBoxField.prototype.ProcessAutoFitContent = function() {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        let oTextPr = oRun.Get_CompiledPr(true);
        let oBounds = this.getFormRelRect();

        g_oTextMeasurer.SetTextPr(oTextPr, null);
	    g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);

        var nTextHeight = g_oTextMeasurer.GetHeight();
	    var nMaxWidth   = oPara.RecalculateMinMaxContentWidth(false).Max;
	    var nFontSize   = oTextPr.FontSize;

        if (nMaxWidth < 0.001 || nTextHeight < 0.001 || oBounds.W < 0.001 || oBounds.H < 0.001)
		    return;

	    var nNewFontSize = nFontSize;

        nNewFontSize = oBounds.H / g_dKoef_pt_to_mm;
        oRun.SetFontSize(nNewFontSize);
    };
    
    // for radiobutton
    const CheckedSymbol   = 0x25C9;
	const UncheckedSymbol = 0x25CB;

    function CCheckBoxField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.checkbox, nPage, aRect);

        this._style = style.ch;
    }
    CCheckBoxField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CCheckBoxField.prototype.constructor = CCheckBoxField;

    CCheckBoxField.prototype.onMouseDown = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThis = this;
        aFields.forEach(function(field) {
            if (field == oThis) {
                CreateNewHistoryPointForField(oThis);
            }
            else
                TurnOffHistory();

            let oRun = field._content.GetElement(0).GetElement(0);
            if (field._value != "Off") {
                oRun.ClearContent();
                field._value = "Off";
            }
            else {
                field._value = field._exportValue;
                oRun.AddText('✓');
            }

            field._wasChanged = true;
        });
        
        private_getViewer()._paintForms();
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.ApplyValueForAll = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThisPara = this._content.GetElement(0);
        let oThisRun = oThisPara.GetElement(0);

        this.CheckValue();
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
             // пропускаем текущее поле, т.к. уже были изменения после redo/undo
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oFieldRun = oFieldPara.GetElement(0);

            oFieldRun.ClearContent();
            for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
            }

            aFields[i]._wasChanged = true;
            aFields[i].CheckValue();
        }
    };
    /**
	 * Checks value of the field and corrects it.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.CheckValue = function() {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        if (oRun.GetText() != "")
            this._value = this._exportValue;
        else
            this._value = "Off";
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let nThisIdx = aFields.indexOf(this);
        
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                this._exportValues = aFields[i]._exportValues.slice();

                if (this._exportValues[nThisIdx])
                    this._exportValue = this._exportValues[nThisIdx];
                else
                    this._exportValue = "Yes";

                if (aFields[i]._value != "Off")
                    this._value = this._exportValue;
                break;
            }
        }
    };

    CCheckBoxField.prototype.SetValue = function(sValue) {
        this._exportValue = sValue;
    };

    function CRadioButtonField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.radiobutton, nPage, aRect);
        
        let oRun = this._content.GetElement(0).GetElement(0);
        //oRun.AddText(String.fromCharCode(UncheckedSymbol));
        oRun.AddText("〇");

        this._radiosInUnison = false;
        this._noToggleToOff = true;

        this._style = style.ci;
    }
    CRadioButtonField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CRadioButtonField.prototype.constructor = CRadioButtonField;
    
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let nThisIdx = aFields.indexOf(this);
                
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._radiosInUnison = aFields[i]._radiosInUnison;
                this._exportValues = aFields[i]._exportValues.slice();
                if (this._exportValues[nThisIdx])
                    this._exportValue = this._exportValues[nThisIdx];
                else
                    this._exportValue = "Yes";

                if (aFields[i]._value != "Off")
                    this._value = this._exportValue;

                if (this._radiosInUnison && this._exportValue == aFields[i]._exportValue) {
                    let oPara = this._content.GetElement(0);
                    let oParaToCopy = aFields[i]._content.GetElement(0);

                    oPara.ClearContent();
                    for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                        oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                    }
                    oPara.CheckParaEnd();
                
                    break;
                }
            }
        }
    };
    CRadioButtonField.prototype.onMouseDown = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThis = this;

        CreateNewHistoryPointForField(oThis);
        if (false == this._radiosInUnison) {
            if (this._value != "Off") {
                if (this._noToggleToOff == false) {
                    this.SetChecked(false);
                    this._wasChanged = true;
                }
                if (AscCommon.History.Is_LastPointEmpty())
                    AscCommon.History.Remove_LastPoint();

                return;
            }
            else {
                this.SetChecked(true);
                this._wasChanged = true;
            }

            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field._value != "Off") {
                    field.SetChecked(false);
                    field._wasChanged = true;
                }
            }); 
        }
        else {
            if (this._value != "Off") {
                if (this._noToggleToOff == false) {
                    this.SetChecked(false);
                    this._wasChanged = true;
                }
            }
            else {
                this.SetChecked(true);
                this._wasChanged = true;
            }

            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field._exportValue != oThis._exportValue && field._value != "Off") {
                    field.SetChecked(false);
                    field._wasChanged = true;
                }
                else if (field._exportValue == oThis._exportValue && oThis._value == "Off") {
                    field.SetChecked(false);
                    field._wasChanged = true;
                }
                else if (field._exportValue == oThis._exportValue && field._value == "Off") {
                    field.SetChecked(true);
                    field._wasChanged = true;
                }
            });

            if (AscCommon.History.Is_LastPointEmpty())
                AscCommon.History.Remove_LastPoint();

        }
        
        private_getViewer()._paintForms();
    };

    /**
	 * Updates all field with this field name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.UpdateAll = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        if (this._radiosInUnison) {
            // отмечаем все radiobuttons с тем же экспортом, что и отмеченные
            let sExportValue;
            for (let i = 0; i < aFields.length; i++) {
                if (!sExportValue && aFields[i]._value != "Off") {
                    sExportValue = aFields[i]._exportValue;
                    break;
                }
            }
            if (!sExportValue) {
                aFields.forEach(function(field) {
                    field.SetChecked(false);
                });
            }
            else {
                aFields.forEach(function(field) {
                    if (field._exportValue != sExportValue) {
                        field.SetChecked(false);
                    }
                    else {
                        field.SetChecked(true);
                    }
                });
            }
        }
        else {
            let oCheckedFld = null;
            // оставляем активной первую отмеченную radiobutton
            for (let i = 0; i < aFields.length; i++) {
                if (!oCheckedFld && aFields[i]._value != "Off") {
                    oCheckedFld = aFields[i];
                    continue;
                }
                if (oCheckedFld) {
                    aFields[i].SetChecked(false);
                }
            }
        }
    };
    /**
	 * Set checked to this field (not for all with the same name).
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.SetChecked = function(bChecked) {
        let oRun = this._content.GetElement(0).GetElement(0);
        if (bChecked) {
            oRun.ClearContent();
            oRun.AddText(String.fromCharCode(CheckedSymbol));
            this._value = this._exportValue;
        }
        else {
            oRun.ClearContent();
            oRun.AddText("〇");
            this._value = "Off";
        }
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.ApplyValueForAll = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        for (let i = 0; i < aFields.length; i++) {
            aFields[i].CheckValue();
            aFields[i]._wasChanged = true;
        }
    };
    /**
	 * Checks value of the field and corrects it.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.CheckValue = function() {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        if (oRun.GetText() != "〇")
            this._value = this._exportValue;
        else
            this._value = "Off";
    };
    CRadioButtonField.prototype.SetNoTogleToOff = function(bValue) {
        this._noToggleToOff = bValue;
    };
    CRadioButtonField.prototype.SetValue = function() {

    };

    function CTextField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.text, nPage, aRect);
        
        this._alignment         = ALIGN_TYPE.left;
        this._calcOrderIndex    = 0;
        this._charLimit         = 0; // to do
        this._comb              = false;
        this._defaultStyle      = Object.assign({}, DEFAULT_SPAN); // to do (must not be fileSelect flag)
        this._doNotScroll       = false;
        this._doNotSpellCheck   = false;
        this._multiline         = false;
        this._password          = false;
        this._richText          = false; // to do связанные свойства, методы
        this._richValue         = [];
        this._textFont          = "ArialMT";
        this._fileSelect        = false;

        // internal
        TurnOffHistory();
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);

        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
        this._contentFormat = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._contentFormat.ParentPDF = this;
        this._contentFormat.SetUseXLimit(false);

        this._scrollInfo = null;
    }
    CTextField.prototype = Object.create(CBaseField.prototype);
	CTextField.prototype.constructor = CTextField;
    
    CTextField.prototype.SetAlign = function(nAlgnType) {
        this._content.SetApplyToAll(true);
        this._contentFormat.SetApplyToAll(true);
        
        this._content.SetParagraphAlign(nAlgnType);
        this._content.SetApplyToAll(false);

        this._contentFormat.SetParagraphAlign(nAlgnType);
        this._contentFormat.SetApplyToAll(false);

        this._content.GetElement(0).private_CompileParaPr(true);
        this._contentFormat.GetElement(0).private_CompileParaPr(true);
    };
    CTextField.prototype.SetComb = function(bComb) {
        if (bComb == true) {
            this._comb = true;
            this._doNotScroll = true;
        }
        else {
            this._comb = false;
            this._content.GetElement(0).Content.forEach(function(run) {
                run.RecalcInfo.Measure = true;
            });
            this._contentFormat.GetElement(0).Content.forEach(function(run) {
                run.RecalcInfo.Measure = true;
            });
        }
    };
    CTextField.prototype.SetDoNotScroll = function(bNot) {
        this._doNotScroll = bNot;
    };
    CTextField.prototype.SetDoNotSpellCheck = function(bNot) {
        this._doNotSpellCheck = bNot;
    };
    CTextField.prototype.SetFileSelect = function(bFileSelect) {
        if (bFileSelect === true && this._multiline != true && this._charLimit === 0
            && this.password != true && this.defaultValue == "") {
                this._fileSelect = true;
            }
        else if (bFileSelect === false) {
            this._fileSelect = false;
        }
    };
    CTextField.prototype.SetMultiline = function(bMultiline) {
        if (bMultiline == true && this.fileSelect != true) {
            this._content.SetUseXLimit(true);
            this._contentFormat.SetUseXLimit(true);
            this._multiline = true;
        }
        else if (bMultiline === false) {
            this._content.SetUseXLimit(false);
            this._contentFormat.SetUseXLimit(false);
            this._multiline = false;
        }
    };
    CTextField.prototype.SetPassword = function(bPassword) {
        if (bPassword === true && this.fileSelect != true) {
            this._password = true;
        }
        else if (bPassword === false) {
            this._password = false;
        }
    };
    CTextField.prototype.SetRichText = function(bRichText) {
        this._richText = bRichText;
    };
    CTextField.prototype.SetValue = function(sValue) {
        this._content.GetElement(0).GetElement(0).AddText(sValue);
    };


    CTextField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = private_getViewer();

        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        this.DrawBorders(oCtx, X, Y, nWidth, nHeight);
        let oMargins = this.GetBordersWidth();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.01 + oMargins.left) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.99 - oMargins.right) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm / scaleCoef;

        let oContentToDraw = this._actions.Format && oViewer.mouseDownFieldObject != this ? this._contentFormat : this._content;

        if ((this.borderStyle == "solid" || this.borderStyle == "dashed") && 
        this._comb == true && this._charLimit > 1) {
            contentX = (X) * g_dKoef_pix_to_mm / scaleCoef;
            contentXLimit = (X + nWidth) * g_dKoef_pix_to_mm / scaleCoef;
        }
        
        if (this._multiline == false) {
            // выставляем текст посередине
            let nContentH = this._content.GetElement(0).Get_EmptyHeight();
            contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm / scaleCoef - nContentH / 2;
        }

        this._formRelRectMM.X = contentX;
        this._formRelRectMM.Y = contentY;
        this._formRelRectMM.W = contentXLimit - contentX;
        this._formRelRectMM.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._contentFormat.X = this._oldContentPos.X = contentX;
            this._content.Y      = this._contentFormat.Y = this._oldContentPos.Y = contentY;
            this._content.XLimit = this._contentFormat.XLimit = this._oldContentPos.XLimit = contentXLimit;
            this._content.YLimit = this._contentFormat.YLimit = this._oldContentPos.YLimit = 20000;
            this._content.Recalculate_Page(0, true);
            this._contentFormat.Recalculate_Page(0, true);
        }
        else if (this._wasChanged) {
            oContentToDraw.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this._wasChanged = false;
        }
        
        if (this._multiline == true) {
            oContentToDraw.ResetShiftView();
            oContentToDraw.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        if (this._needShiftContentView)
            this.CheckFormViewWindow();

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        oGraphics.init(oCtx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        oGraphics.AddClipRect(oContentToDraw.X, oContentToDraw.Y, oContentToDraw.XLimit - oContentToDraw.X, contentYLimit - contentY);

        oContentToDraw.Draw(0, oGraphics);

        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && oContentToDraw.IsSelectionUse() == false && oViewer.fieldFillingMode)
            oContentToDraw.RecalculateCurPos();
        
        oGraphics.RemoveClip();
        this._pageIndX = pageIndX;
        this._pageIndY = pageIndY;
        
        // save pos in page.
        this._pagePos = {
            x: X - pageIndX,
            y: Y - pageIndY,
            w: nWidth,
            h: nHeight,
            realX: X,
            realY: Y
        };

        if (this._doNotScroll == false && this._multiline == true) {
            if (this._wasChanged == false)
                this.UpdateScroll(true);
            else
                this.UpdateScroll(false, true);
        }

        this._wasChanged = false;
    };
    CTextField.prototype.onMouseDown = function(x, y, e) {
        let oViewer = private_getViewer();
                
        let mouseXInPage = x - oViewer.x;
        let mouseYInPage = y - oViewer.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        
        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1;
        
        this._content.Selection_SetStart(X, Y, 0, e);
        //this._content.RemoveSelection();
        this._content.RecalculateCurPos();
        if (this._doNotScroll == false && this._multiline)
            this.UpdateScroll(false, true);
    };
    CTextField.prototype.SelectionSetStart = function(x, y, e) {
        let oViewer = private_getViewer();
        
        let mouseXInPage = x - oViewer.x;
        let mouseYInPage = y - oViewer.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oViewer.zoom;

        this._content.Selection_SetStart(X, Y, 0, e);
    };
    CTextField.prototype.SelectionSetEnd = function(x, y, e) {
        let oViewer = private_getViewer();
        
        let mouseXInPage = x - oViewer.x;
        let mouseYInPage = y - oViewer.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oViewer.zoom;

        this._content.Selection_SetEnd(X, Y, 0, e);
    };
    CTextField.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey)
    {
        this._content.MoveCursorLeft(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey)
    {
        this._content.MoveCursorRight(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        this._content.MoveCursorDown(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        this._content.MoveCursorUp(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.EnterText = function(aChars)
    {
        if (aChars.length > 0)
            CreateNewHistoryPointForField(this);

        let isCanEnter = private_doKeystrokeAction(this, aChars);
        if (isCanEnter == false)
            return false;

        let oPara = this._content.GetElement(0);
        if (this._content.IsSelectionUse() && this._content.IsSelectionEmpty())
            this._content.RemoveSelection();

        let nChars = 0;
        function getCharsCount(oRun) {
            var nCurPos = oRun.Content.length;
			for (var nPos = 0; nPos < nCurPos; ++nPos)
			{
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type)
                    nChars++;
            }
        }
        
        if (this._content.IsSelectionUse()) {
            // Если у нас что-то заселекчено и мы вводим текст или пробел
			// и т.д., тогда сначала удаляем весь селект.
            this._content.Remove(1, true, false, true);
        }

        if (this._charLimit != 0)
            this._content.CheckRunContent(getCharsCount);

        let nMaxCharsToAdd = this._charLimit != 0 ? this._charLimit - nChars : aChars.length;
        for (let index = 0, count = Math.min(nMaxCharsToAdd, aChars.length); index < count; ++index) {
            let codePoint = aChars[index];
            oPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        if (aChars.length > 0) {
            this._wasChanged = true;
            this._needApplyToAll = true; // флаг что значение будет применено к остальным формам с таким именем
            this._needShiftContentView = true && this._doNotScroll == false;
        }

        return true;
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CTextField
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.ApplyValueForAll = function(bUnionPoints) {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThisPara = this._content.GetElement(0);
        
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        TurnOffHistory();

        if (this._actions.Format) {
            this._wasChanged = true;
            this._doc.activeForm = this;
            let isValidFormat = eval(this._actions.Format.script);
            // проверка для форматов, не ограниченных на какие-либо символы своей функцией keystroke
            // например для даты, маски
            if (isValidFormat === false && this.value != "") {
                // отменяем все изменения сделанные в форме, т.к. не подходят формату 
                this.UnionLastHistoryPoints();
                let nPoint = AscCommon.History.Index;
                AscCommon.History.Undo();
                
                // удаляем точки
                AscCommon.History.Points.length = nPoint;

                // to do выдать предупреждение, что строка не подходит по формату
                return;
            }
        }
            
        if (bUnionPoints)
            this.UnionLastHistoryPoints();

        if (aFields.length == 1)
            this._needApplyToAll = false;

        for (let i = 0; i < aFields.length; i++) {
            aFields[i]._content.MoveCursorToStartPos();

            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oThisPara.Content.length - 1; nItem++) {
                oThisRun = oThisPara.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i]._wasChanged = true;
        }

        let oParaFromFormat = this._contentFormat.GetElement(0);
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._contentFormat.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oParaFromFormat.Content.length - 1; nItem++) {
                oThisRun = oParaFromFormat.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i]._wasChanged = true;
        }
    };
    
    /**
	 * Unions the last history points of this form.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.UnionLastHistoryPoints = function() {
        let oTmpPoint;
        let oResultPoint = {
            State      : undefined,
            Items      : [],
            Time       : new Date().getTime(),
            Additional : {FormFilling: this, CanUnion: false},
            Description: undefined
        };

        let i = 0;
        for (i = AscCommon.History.Points.length - 1; i >= 0 ; i--) {
            oTmpPoint = AscCommon.History.Points[i];
            if (oTmpPoint.Additional.FormFilling === this && oTmpPoint.Additional.CanUnion != false) {
                oResultPoint.Items = oTmpPoint.Items.concat(oResultPoint.Items);
            }
            else {
                break;
            }
        }

        i++; // индекс точки, в которую поместим результирующее значение.
        
        // объединяем только больше 2х точек
        if (i < AscCommon.History.Points.length - 1) {
            AscCommon.History.Index = i;
            AscCommon.History.Points.splice(i, AscCommon.History.Points.length - i, oResultPoint);
        }
        else
            AscCommon.History.Points[i].Additional.CanUnion = false; // запрещаем объединять последнюю добавленную точку
    };
    /**
	 * Removes all history points, which were done before form was applied.
	 * @memberof CTextField
     * @param {number} [nCurPoint=AscCommon.History.Index]
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.RemoveNotAppliedChangesPoints = function(nCurPoint) {
        nCurPoint = nCurPoint != undefined ? nCurPoint : AscCommon.History.Index + 1;

        if (!AscCommon.History.Points[nCurPoint + 1] || AscCommon.History.Points[nCurPoint + 1].Additional.CanUnion === false) {
            return;
        }
        AscCommon.History.Points.splice(nCurPoint + 1, AscCommon.History.Points.length - 1);
    };
    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.Remove = function(nDirection, bWord) {
        CreateNewHistoryPointForField(this);

        this._content.Remove(nDirection, true, false, false, bWord);
        
        if (AscCommon.History.Is_LastPointEmpty())
            AscCommon.History.Remove_LastPoint();
        else {
            this._wasChanged = true;
            this._needApplyToAll = true;
        }
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        TurnOffHistory();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._alignment         = aFields[i]._alignment;
                this._calcOrderIndex    = aFields[i]._calcOrderIndex;
                this._charLimit         = aFields[i]._charLimit;
                this._comb              = aFields[i]._comb;
                this._doNotScroll       = aFields[i]._doNotScroll;
                this._doNotSpellCheckl  = aFields[i]._doNotSpellCheckl;
                this._fileSelect        = aFields[i]._fileSelect;
                this._multiline         = aFields[i]._multiline;
                this._password          = aFields[i]._password;
                this._richText          = aFields[i]._richText;
                this._richValue         = aFields[i]._richValue.slice();
                this._textFont          = aFields[i]._textFont;
                this._borderStyle       = aFields[i]._borderStyle;

                this._actions = aFields[i]._actions ? aFields[i]._actions.Copy(this) : null;

                if (this._multiline)
                    this._content.SetUseXLimit(true);

                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();

                // format content
                oPara = this._contentFormat.GetElement(0);
                oParaToCopy = aFields[i]._contentFormat.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                break;
            }
        }
    };

    // pdf api methods

    /**
	 * A string that sets the trigger for the action. Values are:
	 * @typedef {"MouseUp" | "MouseDown" | "MouseEnter" | "MouseExit" | "OnFocus" | "OnBlur" | "Keystroke" | "Validate" | "Calculate" | "Format"} cTrigger
	 * For a list box, use the Keystroke trigger for the Selection Change event.
     */
    
    function CBaseListField(sName, sType, nPage, aRect)
    {
        CBaseField.call(this, sName, sType, nPage, aRect);

        this._commitOnSelChange     = false;
        this._currentValueIndices   = undefined;
        this._textFont              = "ArialMT";
        this._options               = [];
        
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);
    }
    CBaseListField.prototype = Object.create(CBaseField.prototype);
	CBaseListField.prototype.constructor = CBaseListField;

    /**
	 * Unions the last history points of this form.
	 * @memberof CBaseListField
     * @param {boolean} [bForbidToUnion=true] - wheter to forbid to merge the points united by this iteration
	 * @typeofeditors ["PDF"]
	 */
    CBaseListField.prototype.UnionLastHistoryPoints = function(bForbidToUnion) {
        if (bForbidToUnion == undefined)
            bForbidToUnion = true;
            
        let oTmpPoint;
        let oResultPoint = {
            State      : undefined,
            Items      : [],
            Time       : new Date().getTime(),
            Additional : {FormFilling: this, CanUnion: !bForbidToUnion},
            Description: undefined
        };

        let i = 0;
        for (i = AscCommon.History.Points.length - 1; i >= 0 ; i--) {
            oTmpPoint = AscCommon.History.Points[i];
            if (oTmpPoint.Additional.FormFilling === this && oTmpPoint.Additional.CanUnion != false) {
                oResultPoint.Items = oTmpPoint.Items.concat(oResultPoint.Items);
            }
            else {
                break;
            }
        }

        i++; // индекс точки, в которую поместим результирующее значение.
        
        // объединяем только больше 2х точек
        if (i < AscCommon.History.Points.length - 1) {
            AscCommon.History.Index = i;
            AscCommon.History.Points.splice(i, AscCommon.History.Points.length - i, oResultPoint);
        }
        else
            AscCommon.History.Points[i].Additional.CanUnion = !bForbidToUnion; // запрещаем объединять последнюю добавленную точку
    };

    CBaseListField.prototype.SetCommitOnSelChange = function(bValue) {
        this._commitOnSelChange = bValue;
    };

    function CComboBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.combobox, nPage, aRect);

        this._calcOrderIndex    = 0;
        this._doNotSpellCheck   = false;
        this._editable          = false;

        // internal
        this._id = AscCommon.g_oIdCounter.Get_NewId();

        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
        this._contentFormat = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._contentFormat.ParentPDF = this;
        this._contentFormat.SetUseXLimit(false);
    };
    CComboBoxField.prototype = Object.create(CBaseListField.prototype);
	CComboBoxField.prototype.constructor = CComboBoxField;

    CComboBoxField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = private_getViewer();

        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        // маркер dropdown
        let nMarkX = X + nWidth * 0.95 + (nWidth * 0.025) - (nWidth * 0.025)/4;
        let nMarkWidth = nWidth * 0.025;
        let nMarkHeight = nMarkWidth/ 2;
        oCtx.beginPath();
        oCtx.moveTo(nMarkX, Y + nHeight/2 + nMarkHeight/2);
        oCtx.lineTo(nMarkX + nMarkWidth/2, Y + nHeight/2 - nMarkHeight/2);
        oCtx.lineTo(nMarkX - nMarkWidth/2, Y + nHeight/2 - nMarkHeight/2);
        oCtx.fill();

        this._markRect = {
            x1: (nMarkX - nMarkWidth/2) - ((X + nWidth) - (nMarkX + nMarkWidth/2)),
            y1: Y,
            x2: X + nWidth,
            y2: Y + nHeight
        }

        this.DrawBorders(oCtx, X, Y, nWidth, nHeight);
        let oMargins = this.GetBordersWidth();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02 + oMargins.left) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98 - oMargins.right) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm / scaleCoef;
        
        let oContentToDraw = this._actions.Format && oViewer.mouseDownFieldObject != this ? this._contentFormat : this._content;

        // ограничиваем контент позицией маркера
        contentXLimit = this._markRect.x1 * g_dKoef_pix_to_mm / scaleCoef; 
        let nContentH = this._content.GetElement(0).Get_EmptyHeight();
        contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm / scaleCoef - nContentH / 2;

        this._formRelRectMM.X = contentX;
        this._formRelRectMM.Y = contentY;
        this._formRelRectMM.W = contentXLimit - contentX;
        this._formRelRectMM.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._contentFormat.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._contentFormat.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._contentFormat.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._contentFormat.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
            this._contentFormat.Recalculate_Page(0, true);
        }
        else if (this._wasChanged) {
            oContentToDraw.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this._wasChanged = false;
        }
        
        this.CheckFormViewWindow();

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        oGraphics.init(oCtx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        
        oGraphics.AddClipRect(oContentToDraw.X, oContentToDraw.Y, oContentToDraw.XLimit - oContentToDraw.X, contentYLimit - contentY);

        oContentToDraw.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && oContentToDraw.IsSelectionUse() == false && oViewer.fieldFillingMode)
            oContentToDraw.RecalculateCurPos();
        
        oGraphics.RemoveClip();
        this._pageIndX = pageIndX;
        this._pageIndY = pageIndY;
        
        // save pos in page.
        this._pagePos = {
            x: X - pageIndX,
            y: Y - pageIndY,
            w: nWidth,
            h: nHeight,
            realX: X,
            realY: Y
        };
    };
    CComboBoxField.prototype.onMouseDown = function(x, y, e) {
        let oViewer = private_getViewer();
        let X = (x - oViewer.x) * AscCommon.AscBrowser.retinaPixelRatio;
        let Y = (y - oViewer.y) * AscCommon.AscBrowser.retinaPixelRatio;
        
        let mouseXInPage = x - oViewer.x;
        let mouseYInPage = y - oViewer.y;

        let XInContent = mouseXInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        let YInContent = mouseYInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        
        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0; // to do
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1; //

        if (X >= this._markRect.x1 && X <= this._markRect.x2 && Y >= this._markRect.y1 && Y <= this._markRect.y2 && this._options.length != 0) {
            editor.sendEvent("asc_onShowPDFFormsActions", this, x, y);
            this._content.MoveCursorToStartPos();
        }
        else {
            this._content.Selection_SetStart(XInContent, YInContent, 0, e);
            this._content.RemoveSelection();
        }
        
        this._content.RecalculateCurPos();
    };
    
    /**
	 * Selects the specified option.
	 * @memberof CComboBoxField
     * @param {boolean} [bAddToHistory=true] - whether to add change to history.
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.SelectOption = function(nIdx, bAddToHistory) {
        if (bAddToHistory == undefined)
            bAddToHistory = true;

        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);

        this._currentValueIndices = nIdx;

        if (bAddToHistory) {
            CreateNewHistoryPointForField(this);
        }
        else
            TurnOffHistory();

        oRun.ClearContent();
        if (Array.isArray(this._options[nIdx])) {
            oRun.AddText(this._options[nIdx][0]);
            this._value = this._options[nIdx][0];
        }
        else {
            oRun.AddText(this._options[nIdx]);
            this._value = this._options[nIdx];
        }

        this._wasChanged = true;
        this._needApplyToAll = true;
    };

    CComboBoxField.prototype.SetValue = function(sValue) {
        let sTextToAdd = "";
        for (let i = 0; i < this._options.length; i++) {
            if (this._options[i][1] && this._options[i][1] == sValue) {
                sTextToAdd = this._options[i][1];
                break;
            }
        }
        if (sTextToAdd == "") {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i] == sValue) {
                    sTextToAdd = this._options[i];
                    break;
                }
            }
        }
        
        if (sTextToAdd == "")
            sTextToAdd = sValue;

        this._content.GetElement(0).GetElement(0).AddText(sTextToAdd);
        this.CheckCurValueIndex();
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {

                this._calcOrderIndex    = aFields[i]._calcOrderIndex;
                this._doNotSpellCheck   = aFields[i]._doNotSpellCheck;
                this._editable          = aFields[i]._editable;

                let oPara = this._content.GetElement(0);
                let oParaToCopy = aFields[i]._content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                this._options = aFields[i]._options.slice();
                break;
            }
        }
    };
    CComboBoxField.prototype.EnterText = function(aChars, bForce)
    {
        if (this._editable == false && !bForce)
            return false;

        if (aChars.length > 0)
            CreateNewHistoryPointForField(this);
        else
            return false;

        let isCanEnter = private_doKeystrokeAction(this, aChars);
        if (isCanEnter == false)
            return false;

        let oPara = this._content.GetElement(0);
        if (this._content.IsSelectionUse()) {
            // Если у нас что-то заселекчено и мы вводим текст или пробел
			// и т.д., тогда сначала удаляем весь селект.
            this._content.Remove(1, true, false, true);
        }
        
        for (let index = 0, count = aChars.length; index < count; ++index) {
            let codePoint = aChars[index];
            oPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        this.CheckCurValueIndex();
        this._wasChanged = true;
        this._needApplyToAll = true; // флаг что значение будет применено к остальным формам с таким именем
        
        return true;
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CComboBoxField
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.ApplyValueForAll = function(bUnionPoints) {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThisPara = this._content.GetElement(0);
        
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        TurnOffHistory();

        if (this._actions.Format) {
            this._wasChanged = true;
            this._doc.activeForm = this;
            let isValidFormat = eval(this._actions.Format.script);
            // проверка для форматов, не ограниченных на какие-либо символы своей функцией keystroke
            // например для даты, маски
            if (isValidFormat === false && this.value != "") {
                // отменяем все изменения сделанные в форме, т.к. не подходят формату 
                this.UnionLastHistoryPoints();
                let nPoint = AscCommon.History.Index;
                AscCommon.History.Undo();
                
                // удаляем точки
                AscCommon.History.Points.length = nPoint;

                // to do выдать предупреждение, что строка не подходит по формату
                return;
            }
        }

        this.CheckCurValueIndex();

        if (bUnionPoints)
            this.UnionLastHistoryPoints(true);

        if (aFields.length == 1)
            this._needApplyToAll = false;

        for (let i = 0; i < aFields.length; i++) {
            aFields[i]._content.MoveCursorToStartPos();

            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i]._content.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oThisPara.Content.length - 1; nItem++) {
                oThisRun = oThisPara.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, AscCommon.IsSpace(oThisRun.Content[nRunPos].Value) ? new AscWord.CRunSpace(oThisRun.Content[nRunPos].Value) : new AscWord.CRunText(oThisRun.Content[nRunPos].Value));
                }
            }

            aFields[i]._currentValueIndices = this._currentValueIndices;
            aFields[i]._wasChanged = true;
        }
    };
    /**
	 * Checks curValueIndex, corrects it and return.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CComboBoxField.prototype.CheckCurValueIndex = function() {
        let sValue = this._content.GetElement(0).GetText({ParaEndToSpace: false});
        this._value = sValue;
        let nIdx = -1;
        if (Array.isArray(this._options) == true) {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i][0] === sValue) {
                    nIdx = i;
                    break;
                }
            }
        }
        else {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i] === sValue) {
                    nIdx = i;
                    break;
                }
            }
        }

        this._currentValueIndices = nIdx;
        return nIdx;
    };

    CComboBoxField.prototype.SetEditable = function(bValue) {
        this._editable = bValue;
    };
    CComboBoxField.prototype.SetOptions = function(aOpt) {
        let aOptToPush = [];
        for (let i = 0; i < aOpt.length; i++) {
            if (aOpt[i] == null)
                continue;
            if (typeof(aOpt[i]) == "string" && aOpt[i] != "")
                aOptToPush.push(aOpt[i]);
            else if (Array.isArray(aOpt[i]) && aOpt[i][0] != undefined && aOpt[i][1] != undefined) {
                if (aOpt[i][0].toString && aOpt[i][1].toString) {
                    aOptToPush.push([aOpt[i][0].toString(), aOpt[i][1].toString()])
                }
            }
            else if (typeof(aOpt[i]) != "string" && aOpt[i].toString) {
                aOptToPush.push(aOpt[i].toString());
            }
        }

        this._options = aOptToPush;
    };
    
    function CListBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.listbox, nPage, aRect);

        this._multipleSelection = false;

        // internal
        this._scrollInfo = null;
    };
    CListBoxField.scrollCount = 0;
    CListBoxField.prototype = Object.create(CBaseListField.prototype);
	CListBoxField.prototype.constructor = CListBoxField;

    CListBoxField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = private_getViewer();

        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        this.DrawBorders(oCtx, X, Y, nWidth, nHeight);
        let oMargins = this.GetBordersWidth();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02 + oMargins.left) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01 + oMargins.top) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98 - oMargins.right) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01 - oMargins.bottom) * g_dKoef_pix_to_mm / scaleCoef;
        
        this._formRelRectMM.X = contentX;
        this._formRelRectMM.Y = contentY;
        this._formRelRectMM.W = contentXLimit - contentX;
        this._formRelRectMM.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
        }
        else if (this._wasChanged) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this._wasChanged = false;
        }
        
        this._content.ResetShiftView();
        this._content.ShiftView(this._curShiftView.x, this._curShiftView.y);

        if (this._needShiftContentView)
            this.CheckFormViewWindow();

        let oGraphics = new AscCommon.CGraphics();
        let widthPx = oViewer.canvas.width;
        let heightPx = oViewer.canvas.height;
        
        oGraphics.init(oCtx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);
        
        oGraphics.RemoveClip();
        this._pageIndX = pageIndX;
        this._pageIndY = pageIndY;
        
        // save pos in page.
        this._pagePos = {
            x: X - pageIndX,
            y: Y - pageIndY,
            w: nWidth,
            h: nHeight,
            realX: X,
            realY: Y
        };

        this.UpdateScroll(true);
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.SyncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._multipleSelection = aFields[i]._multipleSelection;
                this._content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++) {
                    this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                }
                
                this._options = aFields[i]._options.slice();
                this._currentValueIndices = aFields.multipleSelection ? aFields[i]._currentValueIndices.slice() : aFields[i]._currentValueIndices;

                let oPara;
                for (let i = 0; i < this._content.Content.length; i++) {
                    oPara = this._content.GetElement(i);
                    if (oPara.Pr.Shd && oPara.Pr.Shd.IsNil() == false)
                        oPara.private_CompileParaPr(true);
                }
                break;
            }
        }
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CListBoxField
     * @param {CListBoxField} [oFieldToSkip] - field to don't be apply changes.
     * @param {boolean} [bUnionPoints=true] - whether to union last changes maked in this form to one history point.
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.ApplyValueForAll = function(oFieldToSkip, bUnionPoints) {
        if (bUnionPoints == undefined)
            bUnionPoints = true;

        let aFields = this._doc.getWidgetsByName(this.name);
        let oThis = this;
        
        
        if (bUnionPoints)
            this.UnionLastHistoryPoints();

        aFields.forEach(function(field) {
            field._wasChanged = true;

            if (oThis == field || oFieldToSkip == field)
                return;
            
            field._curShiftView.y = oThis._curShiftView.y;
            field._needShiftContentView = true;
            if (oThis._multipleSelection) {
                // снимаем выделение с тех, которые не присутсвуют в поле, от которого применяем ко всем
                for (let i = 0; i < field._currentValueIndices.length; i++) {
                    if (oThis._currentValueIndices.includes(field._currentValueIndices[i]) == false) {
                        field.UnselectOption(field._currentValueIndices[i]);
                    }
                }
                
                for (let i = 0; i < oThis._currentValueIndices.length; i++) {
                    // добавляем выделение тем, которые не присутсвуют в текущем поле, но присутсвуют в том, от которого применяем
                    if (field._currentValueIndices.includes(oThis._currentValueIndices[i]) == false) {
                        field.SelectOption(oThis._currentValueIndices[i], false, false);
                    }
                }
                field._currentValueIndices = oThis._currentValueIndices.slice();
            }
            else {
                field._currentValueIndices = oThis._currentValueIndices;
                field.SelectOption(field._currentValueIndices, true, false);
            }
        });
    };

    CListBoxField.prototype.SetMultipleSelection = function(bValue) {
        if (bValue == true) {
            this._multipleSelection = true;
            this._currentValueIndices = [this._currentValueIndices];
        }
        else {
            this._multipleSelection = false;
            this._currentValueIndices = this._currentValueIndices[0];
            this._currentValueIndices != -1 && this.SelectOption(this._currentValueIndices, true);
        }
    };

    CListBoxField.prototype.SelectOption = function(nIdx, isSingleSelect, bAddToHistory) {
        if (bAddToHistory == undefined)
            bAddToHistory = true;

        let oPara = this._content.GetElement(nIdx);
        let oApiPara;
        
        if (bAddToHistory) {
            CreateNewHistoryPointForField(this);
        }
        else
            TurnOffHistory();

        this._content.Set_CurrentElement(nIdx);
        if (isSingleSelect) {
            this._content.Content.forEach(function(para) {
                oApiPara = editor.private_CreateApiParagraph(para);
                if (para.Pr.Shd && para.Pr.Shd.IsNil() == false) {
                    oApiPara.SetShd('nil');
                    para.private_CompileParaPr(true);
                }
            });
        }

        if (oPara) {
            oApiPara = editor.private_CreateApiParagraph(oPara);
            oApiPara.SetShd('clear', 0, 112, 192);
            oApiPara.Paragraph.private_CompileParaPr(true);
        }

        this._needApplyToAll = true;
    };
    CListBoxField.prototype.UnselectOption = function(nIdx) {
        let oApiPara = editor.private_CreateApiParagraph(this._content.GetElement(nIdx));
        oApiPara.SetShd('nil');
        oApiPara.Paragraph.private_CompileParaPr(true);
    };
    CListBoxField.prototype.SetOptions = function(aOpt) {
        this._content.Internal_Content_RemoveAll();
        for (let i = 0; i < aOpt.length; i++) {
            if (aOpt[i] == null)
                continue;
            sCaption = "";
            if (typeof(aOpt[i]) == "string" && aOpt[i] != "") {
                sCaption = aOpt[i];
                this._options.push(aOpt[i]);
            }
            else if (Array.isArray(aOpt[i]) && aOpt[i][0] != undefined && aOpt[i][1] != undefined) {
                if (aOpt[i][0].toString && aOpt[i][1].toString) {
                    this._options.push([aOpt[i][0].toString(), aOpt[i][1].toString()]);
                    sCaption = aOpt[i][0].toString();
                }
            }
            else if (typeof(aOpt[i]) != "string" && aOpt[i].toString) {
                this._options.push(aOpt[i].toString());
                sCaption = aOpt[i].toString();
            }

            if (sCaption != "") {
                oPara = new AscCommonWord.Paragraph(this._content.DrawingDocument, this._content, false);
                oRun = new AscCommonWord.ParaRun(oPara, false);
                this._content.Internal_Content_Add(i, oPara);
                oPara.Add(oRun);
                oRun.AddText(sCaption);
            }
        }
    };
    CListBoxField.prototype.SetValue = function(value) {
        let aIndexes = [];
        if (Array.isArray(value)) {
            for (let sVal of value) {
                let isFound = false;
                for (let i = 0; i < this._options.length; i++) {
                    if (this._options[i][1] && this._options[i][1] == sVal) {
                        if (aIndexes.includes(i))
                            continue;
                        else {
                            isFound = true;
                            aIndexes.push(i);
                            break;
                        }
                    }
                }
                if (isFound == false) {
                    for (let i = 0; i < this._options.length; i++) {
                        if (this._options[i] == sVal) {
                            if (aIndexes.includes(i))
                                continue;
                            else {
                                aIndexes.push(i);
                                break;
                            }
                        }
                    }
                }
            }
        }
        else {
            for (let i = 0; i < this._options.length; i++) {
                if (this._options[i][1] && this._options[i][1] == value) {
                    aIndexes.push(i);
                    break;
                }
            }
            if (aIndexes.length == 0) {
                for (let i = 0; i < this._options.length; i++) {
                    if (this._options[i] == value) {
                        aIndexes.push(i);
                        break;
                    }
                }
            }
        }
        
        let oPara, oApiPara;
        for (let idx of aIndexes) {
            oPara = this._content.GetElement(idx);
            oApiPara = editor.private_CreateApiParagraph(oPara);
            oApiPara.SetShd('clear', 0, 112, 192);
            oPara.private_CompileParaPr(true);
        }

        this._currentValueIndices = aIndexes;
    };

    CListBoxField.prototype.onMouseDown = function(x, y, e) {
        if (this._options.length == 0)
            return;

        let oHTMLPage = private_getViewer();
        let mouseXInPage = x - oHTMLPage.x;
        let mouseYInPage = y - oHTMLPage.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oHTMLPage.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oHTMLPage.zoom;
        
        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1;
        
        let nPos = this._content.Internal_GetContentPosByXY(X, Y, 0);

        if (this._multipleSelection == true) {
            if (e.ctrlKey == true) {
                if (this._currentValueIndices.includes(nPos)) {
                    this.UnselectOption(nPos);
                    this._currentValueIndices.splice(this._currentValueIndices.indexOf(nPos), 1);
                }
                else {
                    this.SelectOption(nPos, false);
                    this._currentValueIndices.push(nPos);
                    this._currentValueIndices.sort();
                }
            }
            else {
                this.SelectOption(nPos, true);
                this._currentValueIndices = [nPos];
            }
        }
        else {
            if (nPos == this._currentValueIndices) {
                this.UpdateScroll(false, true);
                return;
            }
                
            this.SelectOption(nPos, true);
            this._currentValueIndices = nPos;
        }

        this._needShiftContentView = true;

        this.UnionLastHistoryPoints(false);

        private_getViewer()._paintForms();
        this.UpdateScroll(false, true);
    };
    CListBoxField.prototype.MoveSelectDown = function() {
        this._needShiftContentView = true;
        this._content.MoveCursorDown();

        if (this._multipleSelection == true) {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = [this._content.CurPos.ContentPos];
        }
        else {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = this._content.CurPos.ContentPos;
        }
        
        private_getViewer()._paintForms();
        this.UpdateScroll();
    };
    CListBoxField.prototype.MoveSelectUp = function() {
        this._needShiftContentView = true;
        this._content.MoveCursorUp();

        if (this._multipleSelection == true) {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = [this._content.CurPos.ContentPos];
        }
        else {
            this.SelectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = this._content.CurPos.ContentPos;
        }

        private_getViewer()._paintForms();
        this.UpdateScroll();
    };
    CListBoxField.prototype.UpdateScroll = function(bUpdateOnlyPos, bShow) {
        let oContentBounds = this._content.GetContentBounds(0);
        let oFieldBounds = this._content.ParentPDF.getFormRelRect();
        let oScroll, oScrollDocElm, oScrollSettings;
        let nContentH = oContentBounds.Bottom - oContentBounds.Top;
        
        if (nContentH < oFieldBounds.H || bShow == false)
            return;

        if (typeof(bShow) != "boolean" && this._scrollInfo)
            bShow = this._scrollInfo.scroll.canvas.style.display == "none" ? false : true;

        let oViewer = private_getViewer();
        let nKoeff = g_dKoef_pix_to_mm / (oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio);

        if (!bUpdateOnlyPos && !this._scrollInfo && oContentBounds.Bottom - oContentBounds.Top > oFieldBounds.H) {
            CListBoxField.scrollCount++;
            oScrollDocElm = document.createElement('div');
            document.getElementById('editor_sdk').appendChild(oScrollDocElm);
            oScrollDocElm.id = "formScroll_" + CListBoxField.scrollCount;
            oScrollDocElm.style.top = (this._pagePos.realY - 1) / AscCommon.AscBrowser.retinaPixelRatio + 'px';
            oScrollDocElm.style.left = (this._pagePos.realX + this._pagePos.w) / AscCommon.AscBrowser.retinaPixelRatio + 'px';
            oScrollDocElm.style.position = "absolute";
            oScrollDocElm.style.display = "block";
			oScrollDocElm.style.width = "14px";
			oScrollDocElm.style.height = (this._pagePos.h + 2) / AscCommon.AscBrowser.retinaPixelRatio + "px";

            oScrollSettings = editor.WordControl.CreateScrollSettings();
            oScrollSettings.isHorizontalScroll = false;
		    oScrollSettings.isVerticalScroll = true;
		    oScrollSettings.contentH = (oContentBounds.Bottom - oContentBounds.Top) / nKoeff;
            oScrollSettings.screenH = 0;
            oScrollSettings.scrollerMinHeight = 5;
            
            oScroll = new AscCommon.ScrollObject(oScrollDocElm.id, oScrollSettings);
            let oThis = this;
            oScroll.bind("scrollvertical", function(evt) {
                oThis.ScrollVertical(evt.scrollD, evt.maxScrollY);
            });

            let nMaxShiftY = oFieldBounds.H - nContentH;
            let nScrollCoeff = this._curShiftView.y / nMaxShiftY;
            oScroll.scrollVCurrentY = oScroll.maxScrollY * nScrollCoeff;

            this._scrollInfo = {
                scroll: oScroll,
                docElem: oScrollDocElm,
                baseYPos: parseInt(oScrollDocElm.style.top),
                oldZoom: oViewer.zoom,
                scrollCoeff: nScrollCoeff // проскроленная часть
            }

            oScroll.Repos(oScrollSettings, false);
        }
        else if (this._scrollInfo) {
            if (bUpdateOnlyPos) {
                this._scrollInfo.docElem.style.top = (this._pagePos.realY - 1) / AscCommon.AscBrowser.retinaPixelRatio  + 'px';
                this._scrollInfo.docElem.style.left = (this._pagePos.realX + this._pagePos.w) / AscCommon.AscBrowser.retinaPixelRatio + 'px';
                this._scrollInfo.docElem.style.height = (this._pagePos.h + 2) / AscCommon.AscBrowser.retinaPixelRatio  + 'px';
            }
            
            if (this._scrollInfo.oldZoom != oViewer.zoom) {
                this._scrollInfo.oldZoom = oViewer.zoom;
                oScrollSettings = editor.WordControl.CreateScrollSettings();
                oScrollSettings.isHorizontalScroll = false;
                oScrollSettings.isVerticalScroll = true;
                oScrollSettings.contentH = (oContentBounds.Bottom - oContentBounds.Top) / nKoeff;
                oScrollSettings.screenH = 0;
                oScrollSettings.scrollerMinHeight = 5;
                this._scrollInfo.scroll.scrollVCurrentY = this._scrollInfo.scroll.maxScrollY * this._scrollInfo.scrollCoeff;
                this._scrollInfo.scroll.Repos(oScrollSettings, false);
            }

            if (!bUpdateOnlyPos) {
                oScrollSettings = editor.WordControl.CreateScrollSettings();
                oScrollSettings.isHorizontalScroll = false;
                oScrollSettings.contentH = (oContentBounds.Bottom - oContentBounds.Top) / nKoeff;
                oScrollSettings.screenH = 0;
                oScrollSettings.scrollerMinHeight = 5;

                let nMaxShiftY = oFieldBounds.H - nContentH;
                let nScrollCoeff = this._curShiftView.y / nMaxShiftY;
                this._scrollInfo.scroll.scrollVCurrentY = this._scrollInfo.scroll.maxScrollY * nScrollCoeff;
                this._scrollInfo.scroll.Repos(oScrollSettings, false);
                this._scrollInfo.scrollCoeff = nScrollCoeff;
            }

            if (bShow === true)
                this._scrollInfo.scroll.canvas.style.display = "";
            if (bShow === false)
                this._scrollInfo.scroll.canvas.style.display = "none";
        }
    };
    CListBoxField.prototype.ScrollVertical = function(scrollY, maxYscroll) {
        let oContentBounds = this._content.GetContentBounds(0);

        let oFormBounds = this.getFormRelRect();
        let nContentH = oContentBounds.Bottom - oContentBounds.Top;
        let nMaxShiftY = oFormBounds.H - nContentH;

        let nScrollCoeff = scrollY / maxYscroll;
        this._curShiftView.y = nMaxShiftY * nScrollCoeff;
        this._needShiftContentView = false;
        this._scrollInfo.scrollCoeff = nScrollCoeff;
        private_getViewer()._paintForms();
    };
    /**
	 * Checks curValueIndices, corrects it and return.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CListBoxField.prototype.CheckCurValueIndex = function() {
        let nIdx;
        if (this._multipleSelection)
            nIdx = []
        else
            nIdx = -1;

        let oCurPara;
        for (let i = 0; i < this._content.Content.length; i++) {
            oCurPara = this._content.GetElement(i);
            if (oCurPara.Pr.Shd && oCurPara.Pr.Shd.IsNil() == false) {
                if (this._multipleSelection)
                    nIdx.push(i);
                else {
                    nIdx = i;
                    break;
                }
            }
        }

        this._currentValueIndices = nIdx;
        return nIdx;
    };

    function CSignatureField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.signature, nPage, aRect);
    };

    function CFormActions() {
        this.MouseUp = null; 
        this.MouseDown = null; 
        this.MouseEnter = null; 
        this.MouseExit = null; 
        this.OnFocus = null; 
        this.OnBlur = null; 
        this.Keystroke = null; 
        this.Validate = null; 
        this.Calculate = null; 
        this.Format = null;
    }
    CFormActions.prototype.Copy = function() {
        let newObj = new CFormActions();
        if (this.MouseUp != null)
            newObj.MouseUp = this.MouseUp.Copy(); 
        if (this.MouseDown != null)
            newObj.MouseDown = this.MouseDown.Copy(); 
        if (this.MouseEnter != null)
            newObj.MouseEnter = this.MouseEnter.Copy(); 
        if (this.MouseExit != null)
            newObj.MouseExit = this.MouseExit.Copy(); 
        if (this.OnFocus != null)
            newObj.OnFocus = this.OnFocus.Copy(); 
        if (this.OnBlur != null)
            newObj.OnBlur = this.OnBlur.Copy(); 
        if (this.Keystroke != null)
            newObj.Keystroke = this.Keystroke.Copy(); 
        if (this.Validate != null)
            newObj.Validate = this.Validate.Copy(); 
        if (this.Calculate != null)
            newObj.Calculate = this.Calculate.Copy(); 
        if (this.Format != null)
            newObj.Format = this.Format.Copy();

        return newObj;
    }

    function CFormAction(type, sScript) {
        this.type = type;
        this.script = sScript;
    }
    CFormAction.prototype.Copy = function() {
        return new CFormAction(this.type, this.script);
    }
    
    CBaseField.prototype.CheckFormViewWindow = function()
    {
        let oParagraph  = this._content.GetElement(this._content.CurPos.ContentPos);
        let nEndPos     = this._content.Pages[0].EndPos;

        // размеры до текущего параграффа
        this._content.Pages[0].EndPos = this._content.CurPos.ContentPos;
        let oPageBoundsToCurPara = this._content.GetContentBounds(0);
        this._content.Pages[0].EndPos = nEndPos;

        // размеры всего контента
        let oPageBounds     = this._content.GetContentBounds(0);
        let oCurParaHeight  = oParagraph.Lines[0].Bottom - oParagraph.Lines[0].Top;

        let oFormBounds = this.getFormRelRect();

        let nDx = 0, nDy = 0;

        if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
        {
            if (oPageBounds.Left > oFormBounds.X)
                nDx = -oPageBounds.Left + oFormBounds.X;
            else if (oPageBounds.Right < oFormBounds.X + oFormBounds.W)
                nDx = oFormBounds.X + oFormBounds.W - oPageBounds.Right;
        }
        else
        {
            nDx = -this._content.ShiftViewX;
        }

        // если высота контента больше чем высота формы (для нескольких параграфов)
        if (this.type == "listbox") {
            if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
                if (oPageBoundsToCurPara.Bottom > oFormBounds.Y + oFormBounds.H)
                    nDy = oFormBounds.Y + oFormBounds.H - oPageBoundsToCurPara.Bottom;
                else if (oPageBoundsToCurPara.Bottom - oCurParaHeight < oFormBounds.Y)
                    nDy = oFormBounds.Y - (oPageBoundsToCurPara.Bottom - oCurParaHeight);
                else if (oPageBoundsToCurPara.Bottom < oFormBounds.Y)
                    nDy = oCurParaHeight;
            }
        }

        if (Math.abs(nDx) > 0.001 || Math.abs(nDy))
        {
            this._content.ShiftView(nDx, nDy);
            this._curShiftView = {
                x: this._content.ShiftViewX,
                y: this._content.ShiftViewY
            }
        }

        var oCursorPos  = oParagraph.GetCalculatedCurPosXY();
        var oLineBounds = oParagraph.GetLineBounds(oCursorPos.Internal.Line);
        var oLastLineBounds = oParagraph.GetLineBounds(oParagraph.GetLinesCount() - 1);

	    nDx = 0;
	    nDy = 0;

        var nCursorT = Math.min(oCursorPos.Y, oLineBounds.Top);
        var nCursorB = Math.max(oCursorPos.Y + oCursorPos.Height, oLineBounds.Bottom);
        var nCursorH = Math.max(0, nCursorB - nCursorT);

        if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
        {
            if (oCursorPos.X < oFormBounds.X)
                nDx = oFormBounds.X - oCursorPos.X;
            else if (oCursorPos.X > oFormBounds.X + oFormBounds.W)
                nDx = oFormBounds.X + oFormBounds.W - oCursorPos.X;
        }

        if (this._multiline) {
            // если высота контента больше чем высота формы
            if (oParagraph.IsSelectionUse()) {
                if (oParagraph.GetSelectDirection() == 1) {
                    if (nCursorT + nCursorH > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else {
                    if (nCursorT < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                }
            }
            else {
                if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
                    if (oLastLineBounds.Bottom - Math.floor(((oFormBounds.Y + oFormBounds.H) * 1000)) / 1000 < 0)
                        nDy = oFormBounds.Y + oFormBounds.H - oLastLineBounds.Bottom;
                    else if (nCursorT < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                    else if (nCursorT + nCursorH > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else
                    nDy = -this._content.ShiftViewY;
            }
        }

        if (this.type == "listbox") nDx = 0;

        if (Math.abs(nDx) > 0.001 || Math.round(nDy) != 0)
        {
            this._content.ShiftView(nDx, nDy);
            this._curShiftView = {
                x: this._content.ShiftViewX,
                y: this._content.ShiftViewY
            }
        }
    };
    CBaseField.prototype.GetBordersWidth = function() {
        let oViewer = private_getViewer();
        let nLineWidth = 1 * oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio * this._lineWidth;

        if (nLineWidth == 0 || this.type == "radiobutton") {
            return {
                left:     0,
                top:      0,
                right:    0,
                bottom:   0
            }
        }

        switch (this._borderStyle) {
            case "solid":
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    nLineWidth,
                    bottom:   nLineWidth
                }
            case "beveled":
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    2 * nLineWidth,
                    bottom:   2 * nLineWidth
                }             
            case "dashed":
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    nLineWidth,
                    bottom:   nLineWidth
                }
            case "inset":
                return {
                    left:     2 * nLineWidth,
                    top:      2 * nLineWidth,
                    right:    2 * nLineWidth,
                    bottom:   2 * nLineWidth
                }
            case "underline":
                return {
                    left:     0,
                    top:      0,
                    right:    0,
                    bottom:   nLineWidth
                }
        }
    };
    CBaseField.prototype.HasShiftView = function() {
        if (this._content.ShiftViewX != 0 || this._content.ShiftViewY != 0)
            return true;

        return false;
    };
    CBaseField.prototype.MoveCursorToStartPos = function() {
        this._content.MoveCursorToStartPos();
    };

    // for format

    /**
	 * Convert field value to specific number format.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
     * @param {number} negStyle = 0 black minus / 1 red minus / 2 parens black / 3 parens red /
     * @param {number} currStyle = reserved
     * @param {string} strCurrency = string of currency to display
     * @param {boolean} bCurrencyPrepend = true = pre pend / false = post pend
	 * @typeofeditors ["PDF"]
	 */
    function AFNumber_Format(nDec, sepStyle, negStyle, currStyle, strCurrency, bCurrencyPrepend) {
        let oCurForm = oDoc.activeForm;

        let oInfoObj = {
            decimalPlaces: nDec,
            separator: true,
            symbol: null,
            type: Asc.c_oAscNumFormatType.Number
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[oInfoObj.symbol]);
        switch (sepStyle) {
            case 0:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = ",";
                break;
            case 1:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 2:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = ".";
                break;
            case 3:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 4:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "'";
                break;
        }

        oCultureInfo.NumberGroupSizes = [3];
        
        let aFormats = AscCommon.getFormatCells(oInfoObj);
        let oNumFormat = AscCommon.oNumFormatCache.get(aFormats[0]);
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        let sCurValue = oCurForm.value;
        if (sCurValue == "") {
            oTargetRun.ClearContent();
            return;
        }
            
        let sRes = oNumFormat.format(sCurValue, 0, AscCommon.gc_nMaxDigCount, true, oCultureInfo, true)[0].text;

        if (bCurrencyPrepend)
            sRes = strCurrency + sRes;
        else
            sRes = sRes + strCurrency;

        if (sRes.indexOf("-") != - 1) {
            sRes = sRes.replace("-", "");
            switch (negStyle) {
                case 0:
                    oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
                    break;
                case 1:
                    oTargetRun.Pr.Color = private_GetColor(255, 0, 0, false);
                    break;
                case 2:
                    oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
                    sRes = "(" + sRes + ")";
                    break;
                case 3:
                    oTargetRun.Pr.Color = private_GetColor(255, 0, 0, false);
                    sRes = "(" + sRes + ")";
                    break;
            }
        }
        else {
            oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
        }
        
        oTargetRun.RecalcInfo.TextPr = true
        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
     * @param {number} negStyle = 0 black minus / 1 red minus / 2 parens black / 3 parens red /
     * @param {number} currStyle = reserved
     * @param {string} strCurrency = string of currency to display
     * @param {boolean} bCurrencyPrepend = true = pre pend / false = post pend
	 * @typeofeditors ["PDF"]
	 */
    function AFNumber_Keystroke(nDec, sepStyle, negStyle, currStyle, strCurrency, bCurrencyPrepend) {
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        function isValidNumber(str) {
            return !isNaN(str) && isFinite(str);
        }

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});

        // разделитель дробной части, который можно ввести
        switch (sepStyle) {
            case 0:
            case 1:
            case 4:
                if (sResultText.indexOf(",") != -1)
                    return false;

                if (isValidNumber(sResultText) == false)
                    return false;
                break;
            case 2:
            case 3:
                if (sResultText.indexOf(".") != -1)
                    return false;

                sResultText = sResultText.replace(/\,/g, ".");
                if (isValidNumber(sResultText) == false)
                    return false;
                break;
        }

        return true;
    }

    /**
	 * Convert field value to specific percent format.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
	 * @typeofeditors ["PDF"]
	 */
    function AFPercent_Format(nDec, sepStyle) {
        let oCurForm = oDoc.activeForm;

        let oInfoObj = {
            decimalPlaces: nDec,
            separator: true,
            symbol: null,
            type: Asc.c_oAscNumFormatType.Number
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[oInfoObj.symbol]);
        switch (sepStyle) {
            case 0:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = ",";
                break;
            case 1:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 2:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = ".";
                break;
            case 3:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 4:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "'";
                break;
        }
        oCultureInfo.NumberGroupSizes = [3];

        let aFormats = AscCommon.getFormatCells(oInfoObj);
        let oNumFormat = AscCommon.oNumFormatCache.get(aFormats[0]);
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        let sCurValue = oCurForm.value;
        sCurValue.replace(",", ".");
        if (sCurValue == "")
            sCurValue = 0;
            
        sCurValue = (parseFloat(sCurValue) * 100).toString();
        let sRes = oNumFormat.format(sCurValue, 0, AscCommon.gc_nMaxDigCount, true, oCultureInfo, true)[0].text;
        sRes = sRes + "%";

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
	 * @typeofeditors ["PDF"]
	 */
    function AFPercent_Keystroke(nDec, sepStyle) {
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        function isValidNumber(str) {
            return !isNaN(str) && isFinite(str);
        }

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});

        // разделитель дробной части, который можно ввести
        switch (sepStyle) {
            case 0:
            case 1:
            case 4:
                if (sResultText.indexOf(",") != -1)
                    return false;

                if (isValidNumber(sResultText) == false)
                    return false;
                break;
            case 2:
            case 3:
                if (sResultText.indexOf(".") != -1)
                    return false;

                sResultText = sResultText.replace(/\,/g, ".");
                if (isValidNumber(sResultText) == false)
                    return false;
                break;
        }

        return true;
    }

    /**
	 * Convert field value to specific date format.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFDate_Format(cFormat) {
        let oCurForm = oDoc.activeForm;

        let oNumFormat = AscCommon.oNumFormatCache.get(cFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        let sCurValue = oCurForm.value;
        let oFormatParser = AscCommon.g_oFormatParser;

        function getShortPattern(aRawFormat) {
            let dayDone     = false;
            let monthDone   = false;
            let yearDone    = false;
            
            let sPattern = "";

            let numFormat_Year = 12;
            let numFormat_Month = 13;
            let numFormat_Day = 16;

            for (let obj of aRawFormat) {
                switch (obj.type) {
                    case numFormat_Day:
                        if (dayDone == false) {
                            sPattern += 1;
                            dayDone = true;
                        }
                        break;
                    case numFormat_Month:
                        if (monthDone == false) {
                            sPattern += 3;
                            monthDone = true;
                        }
                        break;
                    case numFormat_Year:
                        if (yearDone == false) {
                            if (obj.val > 2)
                                sPattern += 5;
                            else
                                sPattern += 4;
                            yearDone = true;
                        }
                        break;
                            
                }
            }
            return sPattern;
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);
        if (null == oNumFormat.oTextFormat.ShortDatePattern) {
            oNumFormat.oTextFormat.ShortDatePattern = getShortPattern(oNumFormat.oTextFormat.aRawFormat);
            oNumFormat.oTextFormat._prepareFormatDatePDF();
        }
        oCultureInfo.ShortDatePattern = oNumFormat.oTextFormat.ShortDatePattern;

        if (oCultureInfo.ShortDatePattern.indexOf("1") == -1)
            oNumFormat.oTextFormat.bDay = false;

        oCultureInfo.AbbreviatedMonthNames.length = 12;
        oCultureInfo.MonthNames.length = 12;

        let oResParsed = oFormatParser.parseDatePDF(sCurValue, oCultureInfo, oNumFormat);
        
        if (sCurValue == "")
            oTargetRun.ClearContent();
        if (!oResParsed) {
            return false;
        }

        oNumFormat.oTextFormat.formatType = AscCommon.NumFormatType.PDFFormDate;
        let sRes = oNumFormat.oTextFormat.format(oResParsed.value, 0, AscCommon.gc_nMaxDigCount, oCultureInfo)[0].text;

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFDate_Keystroke(cFormat) {
        return true;
    }
    let AFDate_FormatEx = AFDate_Format;
    let AFDate_KeystrokeEx = AFDate_Keystroke;

    /**
	 * Convert field value to specific time format.
	 * @memberof CTextField
     * @param {number} ptf - time format
     *  0 = 24HR_MM [ 14:30 ]
     *  1 = 12HR_MM [ 2:30 PM ]
     *  2 = 24HR_MM_SS [ 14:30:15 ]
     *  3 = 12HR_MM_SS [ 2:30:15 PM ]
	 * @typeofeditors ["PDF"]
	 */
    function AFTime_Format(ptf) {
        let oCurForm = oDoc.activeForm;
        if (!oCurForm)
            return;

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);

        let sFormat;
        let fIsValidTime = null;
        switch (ptf) {
            case 0:
                sFormat = "HH:MM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*$/;
                    return pattern.test(time);
                }
                break;
            case 1:
                sFormat = "h:MM AM/PM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
            case 2:
                sFormat = "HH:MM:ss";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*$/;
                    return pattern.test(time);
                }
                break;
            case 3:
                sFormat = "h:MM:ss AM/PM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
        }

        let oNumFormat = AscCommon.oNumFormatCache.get(sFormat);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);
        let sCurValue = oCurForm.value;
        
        if (sCurValue == "")
            oTargetRun.ClearContent();
        else if (fIsValidTime(sCurValue) == false)
            return false;

        let oFormatParser = AscCommon.g_oFormatParser;
        let oResParsed = oFormatParser.parseDatePDF(sCurValue, AscCommon.g_aCultureInfos[9]);

        if (!oResParsed) {
            oTargetRun.ClearContent();
            return false;
        }
        
        let sRes = oNumFormat.format(oResParsed.value, 0, AscCommon.gc_nMaxDigCount, true, undefined, true)[0].text;

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFTime_Keystroke(cFormat) {
        return true;
    }

    let AFTime_FormatEx = AFDate_FormatEx;
    let AFTime_KeystrokeEx = AFTime_Keystroke;

    /**
	 * Convert field value to specific special format.
	 * @memberof CTextField
     * @param {number} psf – psf is the type of formatting to use:
     *  0 = zip code
     *  1 = zip + 4
     *  2 = phone
     *  3 = SSN
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_Format(psf) {
        let oCurForm = oDoc.activeForm;
        if (!oCurForm)
            return;

        let sFormValue = oCurForm.value;
        let oTargetRun = oCurForm._contentFormat.GetElement(0).GetElement(0);

        function isValidZipCode(zipCode) {
            let regex = /^\d{5}$/;
            return regex.test(zipCode);
        }
        function isValidZipCode4(zip) {
            let regex = /^\d{5}[-\s.]?(\d{4})?$/;
            return regex.test(zip);
        }
        function isValidPhoneNumber(number) {
            let regex = /^\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$/;
            return regex.test(number);
        }
        function isValidSSN(ssn) {
            let regex = /^\d{3}[-\s.]?\d{2}[-\s.]?\d{4}$/;
            return regex.test(ssn);
        }

        if (sFormValue == "")
            oTargetRun.ClearContent();
            
        switch (psf) {
            case 0:
                if (isValidZipCode(sFormValue) == false)
                    return false;
                break;
            case 1:
                if (isValidZipCode4(sFormValue) == false)
                    return false;
                break;
            case 2:
                if (isValidPhoneNumber(sFormValue) == false)
                    return false;
                break;
            case 3:
                if (isValidSSN(sFormValue) == false)
                    return false;
                break;
        }

        sFormValue = sFormValue.replace(/\D/g, ""); // delete all except no digit chars
        let sFormatValue = "";

        oTargetRun.ClearContent();

        switch (psf) {
            case 0:
                sFormatValue = sFormValue.substring(0, 5);
                break;
            case 1:
                sFormatValue = sFormValue.substring(0, 9);
                if (sFormatValue[4])
                    sFormatValue = sFormValue.substring(0, 5) + "-" + sFormValue.substring(5);
                break;
            case 2: 
                let x = sFormValue.substring(0, 10);
                sFormatValue = x.length > 6
                    ? "(" + x.substring(0, 3) + ") " + x.substring(3, 6) + "-" + x.substring(6, 10)
                    : x.length > 3
                    ? "(" + x.substring(0, 3) + ") " + x.substring(3)
                    : x;
                break;
            case 3:
                let y = sFormValue.substring(0, 9);
                sFormatValue = y.length > 5
                ? y.substring(0, 3) + "-" + y.substring(3, 5) + "-" + y.substring(5, 9)
                : y.length > 2
                ? y.substring(0, 3) + "-" + x.substring(3)
                : x;
                break;
        }

        oTargetRun.AddText(sFormatValue);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} psf – psf is the type of formatting to use:
     *  0 = zip code
     *  1 = zip + 4
     *  2 = phone
     *  3 = SSN
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_Keystroke(psf) {
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        function isValidZipCode(zipCode) {
            let regex = /^\d{0,5}$/;
            return regex.test(zipCode);
        }
        function isValidZipCode4(zip) {
            let regex = /^\d{0,5}[-\s.]?(\d{0,4})?$/;
            return regex.test(zip);
        }
        function isValidPhoneNumber(number) {
            let regex = /^\(?\d{0,3}?\)?[\s.-]?\d{0,3}?[\s.-]?\d{0,4}?$/;
            return regex.test(number);
        }
        function isValidSSN(ssn) {
            let regex = /^\d{0,3}?[-\s.]?\d{0,2}?[-\s.]?\d{0,4}$/;
            return regex.test(ssn);
        }

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});
        
        let canAdd;
        switch (psf) {
            case 0:
                canAdd = isValidZipCode(sResultText);
                break;
            case 1:
                canAdd = isValidZipCode4(sResultText);
                break;
            case 2:
                canAdd = isValidPhoneNumber(sResultText);
                break;
            case 3:
                canAdd = isValidSSN(sResultText);
                break;
        }

        return canAdd;
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} mask - the special mask
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_KeystrokeEx(mask) {
        let oCurForm = oDoc.activeForm;
        let aEnteredChars = oDoc.enteredFormChars;

        if (!oCurForm)
            return true;

        let isHasSelectedText = oCurForm._content.IsSelectionUse() && oCurForm._content.IsSelectionEmpty() == false;

        let oPara = oCurForm._content.GetElement(0);
        let oTempPara = oPara.Copy(null, oPara.DrawingDocument);
        oTempPara.ClearContent();
        for (var nPos = 0; nPos < oPara.Content.length - 1; nPos++) {
            oTempPara.Internal_Content_Add(nPos, oPara.GetElement(nPos).Copy());
        }
        oTempPara.CheckParaEnd();
        let oSelState = oPara.Get_SelectionState2();
        oTempPara.Set_SelectionState2(oSelState);

        if (isHasSelectedText) {
            oTempPara.Remove(-1, true, false, true);
        }

        for (let index = 0; index < aEnteredChars.length; ++index) {
            let codePoint = aEnteredChars[index];
            oTempPara.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
        }

        let sResultText = oTempPara.GetText({ParaEndToSpace: false});
        
        if (typeof(mask) == "string" && mask != "") {
            let oTextFormat = new AscWord.CTextFormFormat();
            let arrBuffer = oTextFormat.GetBuffer(sResultText);

            let oFormMask = new AscWord.CTextFormMask();
            oFormMask.Set(mask);
            return oFormMask.Check(arrBuffer);
        }

        return false;
    }


    
    // private methods

    function private_GetColor(r, g, b, Auto)
	{
		return new AscCommonWord.CDocumentColor(r, g, b, Auto ? Auto : false);
	}
    
    function private_doKeystrokeAction(oField, aChars) {
        let isValid = true;
        if (oField._actions.Keystroke) {
            oField._doc.enteredFormChars = aChars;
            oField._doc.activeForm = oField;
            isValid = eval(oField._actions.Keystroke.script);
        }

        return isValid;
    }

    function private_GetFieldAlign(sJc)
	{
		if ("left" === sJc)
			return align_Left;
		else if ("right" === sJc)
			return align_Right;
		else if ("center" === sJc)
			return align_Center;

		return undefined;
	}

    function private_getViewer() {
        return editor.getDocumentRenderer();
    }
    function CreateNewHistoryPointForField(oField) {
        if (AscCommon.History.IsOn() == false)
            AscCommon.History.TurnOn();
        AscCommon.History.Create_NewPoint();
        AscCommon.History.SetAdditionalFormFilling(oField);
    }

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    CComboBoxField.prototype.Remove                 = CTextField.prototype.Remove;
    CComboBoxField.prototype.MoveCursorLeft         = CTextField.prototype.MoveCursorLeft;
    CComboBoxField.prototype.MoveCursorRight        = CTextField.prototype.MoveCursorRight;
    CComboBoxField.prototype.SelectionSetStart      = CTextField.prototype.SelectionSetStart;
    CComboBoxField.prototype.SelectionSetEnd        = CTextField.prototype.SelectionSetEnd;

    CComboBoxField.prototype.SetAlign = CTextField.prototype.SetAlign;
    CComboBoxField.prototype.SetDoNotSpellCheck = CTextField.prototype.SetDoNotSpellCheck;

    CComboBoxField.prototype.RemoveNotAppliedChangesPoints = CTextField.prototype.RemoveNotAppliedChangesPoints;
    CTextField.prototype.UpdateScroll               = CListBoxField.prototype.UpdateScroll;
    CTextField.prototype.ScrollVertical             = CListBoxField.prototype.ScrollVertical;
    
    if (!window["AscPDFEditor"])
	    window["AscPDFEditor"] = {};
        
	window["AscPDFEditor"].FIELD_TYPE           = FIELD_TYPE;
	window["AscPDFEditor"].CPushButtonField     = CPushButtonField;
	window["AscPDFEditor"].CBaseField           = CBaseField;
	window["AscPDFEditor"].CTextField           = CTextField;
	window["AscPDFEditor"].CCheckBoxField       = CCheckBoxField;
	window["AscPDFEditor"].CComboBoxField       = CComboBoxField;
	window["AscPDFEditor"].CListBoxField        = CListBoxField;
	window["AscPDFEditor"].CRadioButtonField    = CRadioButtonField;
	window["AscPDFEditor"].CSignatureField      = CSignatureField;
})();
