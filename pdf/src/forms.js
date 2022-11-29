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

    // field types

    let FIELD_TYPE = {
        button:         "button",
        checkbox:       "checkbox",
        combobox:       "combobox",
        listbox:        "listbox",
        radiobutton:    "radiobutton",
        signature:      "signature",
        text:           "text"
    }
    
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
        none:   0,
        thin:   1,
        medium: 2,
        thick:  3
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
        this._textSize      = 12; // 0 == max text size
        this._userName      = ""; // It is intended to be used as tooltip text whenever the cursor enters a field. 
        //It can also be used as a user-friendly name, instead of the field name, when generating error messages.

        this._oldContentPos = {X: 0, Y: 0, XLimit: 0, YLimit: 0};

        editor.getDocumentRenderer().ImageMap = {};
        editor.getDocumentRenderer().InitDocument = function() {return};

        this._partialName = sName;
    }
    Object.defineProperties(CBaseField.prototype, {
        // private
        "_parent": {
            enumerable: false,
            writable: true,
            value: null
        },
        "_pagePos": {
            writable: true,
            enumerable: false,
            value: {
                x: 0,
                y: 0,
                w: 0,
                h: 0
            }
        },
        "_kids": {
            enumerable: false,
            value: [],
        },
        "_partialName": {
            writable: true,
            enumerable: false
        },


        // common
        "borderStyle": {
            set(sValue) {
                if (Object.values(border).includes(sValue))
                    this._borderStyle = sValue;
            },
            get() {
                return this._borderStyle;
            }
        },
        "delay": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._delay = bValue;
            },
            get() {
                return this._delay;
            }
        },
        "display": {
            set(nValue) {
                if (Object.values(display).includes(nValue))
                    this._display = nValue;
            },
            get() {
                return this._display;
            }
        },
        "doc": {
            get() {
                return this._doc;
            }
        },
        "fillColor": {
            set (aColor) {
                if (Array.isArray(aColor))
                    this._fillColor = aColor;
            },
            get () {
                return this._fillColor;
            }
        },
        "bgColor": {
            set (aColor) {
                if (Array.isArray(aColor))
                    this._bgColor = aColor;
            },
            get () {
                return this._bgColor;
            }
        },
        "hidden": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._hidden = bValue;
            },
            get() {
                return this._hidden;
            }
        },
        "lineWidth": {
            set(nValue) {
                nValue = parseInt(nValue);
                if (!isNaN(nValue))
                    this._lineWidth = nValue;
            },
            get() {
                return this._lineWidth;
            }
        },
        "borderWidth": {
            set(nValue) {
                nValue = parseInt(nValue);
                if (!isNaN(nValue))
                    this._borderWidth = nValue;
            },
            get() {
                return this._borderWidth;
            }
        },
        "name": {
            get() {
                if (this._parent)
                {
                    if (this._partialName != "")
                        return `${this._parent.name}.${this._partialName}`
                    else
                        return this._parent.name;
                }

                return this._partialName ? this._partialName : "";
            }
        },
        "page": {
            get() {
                return this._page;
            }
        },
        "print": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._print = bValue;
            },
            get() {
                return this._print;
            }
        },
        "readonly": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._readonly = bValue;
            },
            get() {
                return this._readonly;
            }
        },
        "rect": {
            set(aRect) {
                if (Array.isArray(aRect)) {
                    let isValidRect = true;
                    for (let i = 0; i < 4; i++) {
                        if (typeof(aRect[i]) != "number") {
                            isValidRect = false;
                            break;
                        }
                    }
                  
                    if (isValidRect)
                        this._rect = aRect;
                }
            },
            get() {
                return this._rect;
            }
        },
        "required": {
            set(bValue) {
                if (typeof(bValue) == "boolean" && this.type != "button")
                    this._required = bValue;
            },
            get() {
                if (this.type != "button")
                    return this._required;

                return undefined;
            }
        },
        "rotation": {
            set(nValue) {
                if (VALID_ROTATIONS.includes(nValue))
                    this._rotation = nValue;
            },
            get() {
                return this._rotation;
            }
        },
        "strokeColor": {
            set(aColor) {
                if (Array.isArray(aColor))
                    this._strokeColor = aColor;
            },
            get() {
                return this._strokeColor;
            }
        },
        "borderColor": {
            set(aColor) {
                if (Array.isArray(aColor))
                    this._borderColor = aColor;
            },
            get() {
                return this._borderColor;
            }
        },
        "submitName": {
            set(sValue) {
                if (typeof(sValue) == "string")
                    this._submitName = sValue;
            },
            get() {
                return this._submitName;
            }
        },
        "textColor": {
            set (aColor) {
                if (Array.isArray(aColor))
                    this._textColor = aColor;
            },
            get () {
                return this._textColor;
            }
        },
        "fgColor": {
            set (aColor) {
                if (Array.isArray(aColor))
                    this._fgColor = aColor;
            },
            get () {
                return this._fgColor;
            }
        },
        "textSize": {
            set(nValue) {
                if (typeof(nValue) != "number" && nValue >= 0 && nValue < MAX_TEXT_SIZE)
                    this._textSize = nValue;
            },
            get() {
                return 
            }
        },
        "userName": {
            set(sValue) {
                if (typeof(sValue) == "string")
                    this._userName = sValue;
            },
            get() {
                return this._userName;
            }
        }
    });
    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.privat_getField = function(sName) {
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
    CBaseField.prototype.private_addKid = function(oField) {
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
    CBaseField.prototype.private_removeKid = function(oField) {
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
        var nX = this._content.X, nW = this._content.XLimit - this._content.X;
	    var nY = this._content.Y, nH = this._content.YLimit - this._content.Y;
      
        var aX = [nX, nW];
        var aY = [nY, nH];
        var fX0, fY0;

        return {
            X    : nX,
            Y    : nY,
            W    : nW,
            H    : nH,
            Page : this._content.CurPage
        };

        var aRelX = [], aRelY = [];
        for(var nX = 0; nX < aX.length; ++nX) {
            fX0 = aX[nX];
            for(var nY = 0; nY < aY.length; ++nY) {
                fY0 = aY[nY];
                var fX = oSpTransform.TransformPointX(fX0, fY0);
                var fY = oSpTransform.TransformPointY(fX0, fY0);
                var fRelX = oInvTextTransform.TransformPointX(fX, fY);
                var fRelY = oInvTextTransform.TransformPointY(fX, fY);
                aRelX.push(fRelX);
                aRelY.push(fRelY);
            }
        }

        return {
            X    : Math.min.apply(Math, aRelX),
            Y    : Math.min.apply(Math, aRelY),
            W    : nW,
            H    : nH,
            Page : this.parent.PageNum
        };
    };

    CBaseField.prototype.Draw = function(viewer, pageIndX, pageIndY) {
        if (!viewer)
            return;

        function round(nValue) {
            return (nValue + 0.5) >> 0;
        }
        
        let ctx = viewer.canvas.getContext("2d");
        
        // draw border
        ctx.beginPath();

        let X = pageIndX + (this._rect[0] * viewer.zoom);
        let Y = pageIndY + (this._rect[1] * viewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * viewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * viewer.zoom;

        switch (this._borderStyle) {
            case "solid":
                //ctx.setLineDash([]);
                break;
            case "beveled":
                break;
            case "dashed":
                //ctx.setLineDash([5 * viewer.zoom]);
                break;
            case "inset":
                break;
            case "underline":
                break;
        }

        ctx.rect(X, Y, nWidth, nHeight);
        ctx.stroke();

        let scaleCoef = viewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nHeight / 8) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight) * g_dKoef_pix_to_mm / scaleCoef;
        
        if (this.type == "checkbox" || this.type == "radiobutton") {
            contentY = Y * g_dKoef_pix_to_mm / scaleCoef;
            this.ProcessAutoFitContent(); // подгоняем размер галочки
        }

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit || contentYLimit != this._oldContentPos.YLimit) {
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = contentYLimit;
            this._content.Recalculate_Page(0, true);
        }
        else {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
        }
        
        this._content.CheckFormViewWindowPDF();
        let oGraphics = new AscCommon.CGraphics();
        let widthPx = viewer.canvas.width;
        let heightPx = viewer.canvas.height;
        
        oGraphics.init(ctx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.endGlobalAlphaColor = [255, 255, 255];
        oGraphics.transform(1, 0, 0, 1, 0, 0);

        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, this._content.YLimit - this._content.Y);

        this._content.Draw(0, oGraphics);
        // redraw target cursor if field is selected
        if (viewer.mouseDownFieldObject == this)
            this._content.RecalculateCurPos();
        
        oGraphics.RemoveClip();
        this._pageIndX = pageIndX;
        this._pageIndY = pageIndY;
        
        // save pos in page.
        this._pagePos = {x: X - pageIndX, y: Y - pageIndY, w: nWidth, h: nHeight};
        
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
        
    }
    CPushButtonField.prototype = Object.create(CBaseField.prototype);
	CPushButtonField.prototype.constructor = CPushButtonField;
    Object.defineProperties(CPushButtonField.prototype, {
        "buttonAlignX": {
            set(nValue) {
                if (typeof(nValue) == "number")
                    this._buttonAlignX = Math.round(nValue);
            },
            get() {
                return this._buttonAlignX;
            },
        },
        "buttonAlignY": {
            set(nValue) {
                if (typeof(nValue) == "number")
                    this._buttonAlignY = Math.round(nValue);
            },
            get() {
                return this._buttonAlignY;
            }
        },
        "buttonFitBounds": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._buttonFitBounds = bValue;
            },
            get() {
                return this._buttonFitBounds;
            }
        },
        "buttonPosition": {
            set(nValue) {
                if (Object.values(position).includes(nValue))
                    this._buttonPosition = nValue;
            },
            get() {
                return this._buttonPosition;
            }
        },
        "buttonScaleHow": {
            set(nValue) {
                if (Object.values(scaleHow).includes(nValue))
                    this._buttonScaleHow = nValue;
            },
            get() {
                return this._buttonScaleHow;
            }
                
        },
        "buttonScaleWhen": {
            set(nValue) {
                if (Object.values(scaleWhen).includes(nValue))
                    this._buttonScaleWhen = nValue;
            },
            get() {
                return this._buttonScaleWhen;
            }
                
        },
        "highlight": {
            set(sValue) {
                if (Object.values(highlight).includes(sValue))
                    this._highlight = sValue;
            },
            get() {
                return this._highlight;
            }
        },
        "textFont": {
            set(sValue) {
                if (typeof(sValue) == "string" && sValue !== "")
                    this._textFont = sValue;
            },
            get() {
                return this.textFont;
            }
        },
        "value": {
            get() {
                return undefined;
            }
        }
    });

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
    Object.defineProperties(CBaseCheckBoxField.prototype, {
        "exportValues": {
            set(arrValues) {
                for (let i = 0; i < arrValues.length; i++)
                    if (typeof(arrValues[i]) !== "string")
                        arrValues[i] = String(arrValues[i]);
                    else if (arrValues[i] === "")
                        arrValues[i] = "Yes";
            },
            get() {
                return this._exportValues;
            }
        },
        "style": {
            set(sStyle) {
                if (Object.values(style).includes(sStyle))
                    this._style = sStyle;
            },
            get() {
                return this._style;
            }
        },
        "value": {
            set(sValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (this._exportValues.includes(sValue)) {
                    aFields.forEach(function(field) {
                        field._value = sValue;
                    }) 
                }
                else
                    aFields.forEach(function(field) {
                        field._value = "Off";
                    }) 
            },
            get() {
                return this._value;
            }
        }
    });

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

    const CheckedSymbol   = 0x2611;
	const UncheckedSymbol = 0x2610;

    function CCheckBoxField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.checkbox, nPage, aRect);

        this._style = style.ch;
    }
    CCheckBoxField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CCheckBoxField.prototype.constructor = CCheckBoxField;

    CCheckBoxField.prototype.onMouseDown = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        aFields.forEach(function(field) {
            let oRun = field._content.GetElement(0).GetElement(0);
            if (field._value != "Off") {
                oRun.ClearContent();
                field._value = "Off";
            }
            else {
                field._value = field._exportValue;
                oRun.AddText('✓');
            }
        });
        
        editor.getDocumentRenderer()._paint();
    };
    

    function CRadioButtonField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.radiobutton, nPage, aRect);
        
        this._radiosInUnison = false;
        this._style = style.ci;
    }
    CRadioButtonField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CRadioButtonField.prototype.constructor = CRadioButtonField;
    Object.defineProperties(CRadioButtonField.prototype, {
        "radiosInUnison": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._radiosInUnison = bValue;
            },
            get() {
                return this._radiosInUnison;
            }
        },
    });


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

        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);

    }
    CTextField.prototype = Object.create(CBaseField.prototype);
	CTextField.prototype.constructor = CTextField;
    Object.defineProperties(CTextField.prototype, {
        "alignment": {
            set(sValue) {
                if (Object.values(ALIGN_TYPE).includes(sValue)) {
                    this._alignment = sValue;
                    this._content.SetApplyToAll(true);
                    var nJcType = private_GetFieldAlign(sValue);
                    this._content.SetParagraphAlign(nJcType);
                    this._content.Content.forEach(function(para) {
                        para.CompiledPr.NeedRecalc = false;
                        para.CompiledPr.Pr.ParaPr.Jc = nJcType;
                    });
                    this._content.SetApplyToAll(false);
                }
            },
            get() {
                return this._alignment;
            }
        },
        "calcOrderIndex": {
            set(nValue) {
                if (typeof(nValue) == "number")
                    this._calcOrderIndex = nValue;
            },
            get() {
                return this._calcOrderIndex;
            }
        },
        "charLimit": {
            set(nValue) {
                if (typeof(nValue) == "number" && nValue <= 500 && nValue > 0 && this.fileSelect === false)
                    this._charLimit = Math.round(nValue);
            },
            get() {
                return this._charLimit;
            }
        },
        "comb": {
            set(bValue) {
                if (bValue === true) {
                    this._comb = true;
                    this._doNotScroll = true;
                }
                else if (bValue === false)
                    this._comb = false;
            },
            get() {
                return this._comb;
            }
        },
        "doNotScroll": {
            set(bValue) {
                if (typeof(bValue) === "boolean")
                    this._doNotScroll = bValue;
            },
            get() {
                return this._doNotScroll;
            }
        },
        "doNotSpellCheck": {
            set(bValue) {
                if (typeof(bValue) === "boolean")
                    this._doNotSpellCheck = bValue;
            },
            get() {
                return this._doNotSpellCheckl;
            }
        },
        "fileSelect": {
            set(bValue) {
                if (bValue === true && this.multiline == false && this.charLimit === 0
                    && this.password == false && this.defaultValue == "")
                    this._fileSelect = true;
                else if (bValue === false)
                    this._fileSelect = false;
            },
            get() {
                return this._fileSelect;
            }
        },
        "multiline": {
            set(bValue) {
                if (bValue === true && this.fileSelect === false)
                    this._multiline = true;
                else if (bValue === false)
                    this._multiline = false;
            },
            get() {
                return this._multiline;
            }
        },
        "password": {
            set (bValue) {
                if (bValue === true && this.fileSelect === false)
                    this._password = true;
                else if (bValue === false)
                    this._password = false;
            },
            get() {
                return this._password;
            }
        },
        "richText": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._richText = bValue;
            },
            get() {
                return this._richText;
            }
        },
        "richValue": {
            set(aSpans) {
                if (Array.isArray(aSpans)) {
                    let aCorrectVals = aSpans.filter(function(item) {
                        if (Array.isArray(item) == false && typeof(item) == "object" && item != null)
                            return item;
                    });

                    this._richValue = aCorrectVals;
                }
            },
            get() {
                return this._richValue;
            }
        },
        "textFont": {
            set(sValue) {
                if (typeof(sValue) == "string" && sValue !== "")
                    this._textFont = sValue;
            },
            get() {
                return this.textFont;
            }
        }

    });
    
    CTextField.prototype.onMouseDown = function(x, y, e) {
        let oHTMLPage = editor.getDocumentRenderer();
                
        let mouseXInPage = x - oHTMLPage.x;
        let mouseYInPage = y - oHTMLPage.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oHTMLPage.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oHTMLPage.zoom;
        
        editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
        editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
        editor.WordControl.m_oDrawingDocument.m_lPagesCount = 1;
        
        let aFields = this._doc.getWidgetsByName(this.name);
        aFields.forEach(function(field) {
            field._content.Selection_SetStart(X, field._content.Y, 0, e);
            field._content.RemoveSelection();
        });
        
        this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey)
    {
        let aFields = this._doc.getWidgetsByName(this.name);
        aFields.forEach(function(field) {
            field._content.MoveCursorLeft(isShiftKey, isCtrlKey);
        });

        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey)
    {
        let aFields = this._doc.getWidgetsByName(this.name);
        aFields.forEach(function(field) {
            field._content.MoveCursorRight(isShiftKey, isCtrlKey);
        });

        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.EnterText = function(aChars)
    {
        if (aChars.length == 0)
            return;

        let aFields = this._doc.getWidgetsByName(this.name);
        aFields.forEach(function(field) {
            let oPara = field._content.GetElement(0);
            let oRun = oPara.GetElement(0);
            for (let index = 0, count = aChars.length; index < count; ++index)
            {
                let codePoint = aChars[index];
                oRun.Add(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
            }
        });

        editor.getDocumentRenderer()._paint();
    };
    
    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.Remove = function(nDirection) {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        aFields.forEach(function(field) {
            field._content.Remove(nDirection, true, false, false, false);
        });
    };

    function CBaseListField(sName, sType, nPage, aRect)
    {
        CBaseField.call(this, sName, sType, nPage, aRect);

        this._commitOnSelChange     = false;
        this._currentValueIndices   = undefined;
        this._numItems              = 0;
        this._textFont              = "ArialMT";
    }
    CBaseListField.prototype = Object.create(CBaseField.prototype);
	CBaseListField.prototype.constructor = CBaseListField;
    Object.defineProperties(CBaseListField.prototype, {
        "commitOnSelChange": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._commitOnSelChange = bValue;
            },
            get() {
                return this._commitOnSelChange;
            }
        },
        "currentValueIndices": {
            set(value) {
                if (Array.isArray(value) && this.multipleSelection === true)
                {
                    let isValid = true;
                    for (let i = 0; i < value.length; i++) {
                        if (typeof(value[i]) != "number") {
                            isValid = false;
                            break;
                        }
                    }

                    if (isValid)
                        this._currentValueIndices = value;
                }
                else if (typeof(value) === "number" && this.getItemAt(value) !== undefined)
                    this._currentValueIndices = value;
            },
            get() {
                return this._currentValueIndices;
            }
        },
        "numItems": {
            get() {
                return this._numItems;
            }
        },
        "textFont": {
            set(sValue) {
                if (typeof(sValue) == "string" && sValue !== "")
                    this._textFont = sValue;
            },
            get() {
                return this.textFont;
            }
        }
    });

    function CComboBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.combobox, nPage, aRect);

        this._calcOrderIndex    = 0;
        this._doNotSpellCheck   = false;
        this._editable          = false;
    };
    CComboBoxField.prototype = Object.create(CBaseListField.prototype);
	CComboBoxField.prototype.constructor = CComboBoxField;
    Object.defineProperties(CComboBoxField.prototype, {
        "calcOrderIndex": {
            set(nValue) {
                if (typeof(nValue) == "number")
                    this._calcOrderIndex = nValue;
            },
            get() {
                return this._calcOrderIndex;
            }
        },
        "doNotSpellCheck": {
            set(bValue) {
                if (typeof(bValue) === "boolean")
                    this._doNotSpellCheck = bValue;
            },
            get() {
                return this._doNotSpellCheckl;
            }
        },
        "editable": {
            set(bValue) {
                if (typeof(bValue) === "boolean")
                    this._editable = bValue;
            },
            get() {
                return this._editablel;
            }
        }
        
    });

    function CListBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.listbox, nPage, aRect);

        this._multipleSelection = false;
        
    };
    CListBoxField.prototype = Object.create(CBaseListField.prototype);
	CListBoxField.prototype.constructor = CListBoxField;
    Object.defineProperties(CListBoxField.prototype, {
        "multipleSelection": {
            set(bValue) {
                if (typeof(bValue) == "boolean")
                    this._multipleSelection = bValue;
            },
            get() {
                return this._multipleSelection;
            }

        }
    });

    function CSignatureField(sName, nPage, aRect)
    {
        CBaseField.call(this, sName, FIELD_TYPE.signature, nPage, aRect);
        
    };

    function CSpan()
    {
        this._alignment = ALIGN_TYPE.left;
        this._fontFamily = ["sans-serif"];
        this._fontStretch = "normal";
        this._fontStyle = "normal";
        this._fontWeight = 400;
        this._strikethrough = false;
        this._subscript = false;
        this._superscript = false;

        Object.defineProperties(this, {
            "alignment": {
                set(sValue) {
                    if (Object.values(ALIGN_TYPE).includes(sValue))
                        this._alignment = sValue;
                },
                get() {
                    return this._alignment;
                }
            },
            "fontFamily": {
                set(arrValue) {
                    if (Array.isArray(arrValue))
                    {
                        let aCorrectFonts = [];

                        if (arrValue[0] !== undefined && typeof(arrValue[0]) == "string" && arrValue[0] === "")
                            aCorrectFonts.push(arrValue[0]);
                        if (arrValue[1] !== undefined && typeof(arrValue[1]) == "string" && arrValue[1] === "")
                            aCorrectFonts.push(arrValue[1]);

                        this._fontFamily = aCorrectFonts;
                    }
                }
            },
            "fontStretch": {
                set(sValue) {
                    if (FONT_STRETCH.includes(sValue))
                        this._fontStretch = sValue;
                },
                get() {
                    return this._fontStretch;
                }
            },
            "fontStyle": {
                set(sValue) {
                    if (Object.values(FONT_STYLE).includes(sValue))
                        this._fontStyle = sValue;
                },
                get() {
                    return this._fontStyle;
                }
            },
            "fontWeight": {
                set(nValue) {
                    if (FONT_WEIGHT.includes(nValue))
                        this._fontWeight = nValue;
                },
                get() {
                    return this._fontWeight;
                }
            },
            "strikethrough": {
                set(bValue) {
                    if (typeof(bValue) == "boolean")
                        this._strikethrough = bValue;
                },
                get() {
                    return this._strikethrough;
                }
            },
            "subscript": {
                set(bValue) {
                    if (typeof(bValue) == "boolean")
                        this._subscript = bValue;
                },
                get() {
                    return this._subscript;
                }
            },
            "superscript": {
                set(bValue) {
                    if (typeof(bValue) == "boolean")
                        this._superscript = bValue;
                },
                get() {
                    return this._superscript;
                }
            },

        });
    }

    // private methods
    
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
