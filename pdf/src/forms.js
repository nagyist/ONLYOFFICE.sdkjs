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

    let FIELDS_HIGHLIGHT = {
        r: 201, 
        g: 200,
        b: 255
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
                if (typeof(nValue) == "number" && nValue >= 0 && nValue < MAX_TEXT_SIZE) {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    let oField;
                    for (var i = 0; i < aFields.length; i++) {
                        oField = aFields[i];
                        oField._textSize = nValue;

                        let aParas = oField._content.Content;
                        aParas.forEach(function(para) {
                           para.SetApplyToAll(true);
                           para.Add(new AscCommonWord.ParaTextPr({FontSize : nValue}));
                           para.SetApplyToAll(false);
                        });
                    }
                }
                    
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

    // CBaseField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
    //     let oViewer = editor.getDocumentRenderer();

    //     function round(nValue) {
    //         return (nValue + 0.5) >> 0;
    //     }
        
    //     let X = pageIndX + (this._rect[0] * oViewer.zoom);
    //     let Y = pageIndY + (this._rect[1] * oViewer.zoom);
    //     let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
    //     let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

    //     switch (this._borderStyle) {
    //         case "solid":
    //             //oCtx.setLineDash([]);
    //             break;
    //         case "beveled":
    //             break;
    //         case "dashed":
    //             //oCtx.setLineDash([5 * oViewer.zoom]);
    //             break;
    //         case "inset":
    //             break;
    //         case "underline":
    //             break;
    //     }

    //     // draw border
    //     if (this.type != "radiobutton") {
    //         oCtx.beginPath();
    //         oCtx.rect(X, Y, nWidth, nHeight);
    //         oCtx.stroke();
    //     }

    //     // маркер списка
    //     if (this.type == "combobox") {
    //         let nMarkX = X + nWidth * 0.95 + (nWidth * 0.025) - (nWidth * 0.025)/4;
    //         let nMarkWidth = nWidth * 0.025;
    //         let nMarkHeight = nMarkWidth/ 2;
    //         oCtx.beginPath();
    //         oCtx.moveTo(nMarkX, Y + nHeight/2 + nMarkHeight/2);
    //         oCtx.lineTo(nMarkX + nMarkWidth/2, Y + nHeight/2 - nMarkHeight/2);
    //         oCtx.lineTo(nMarkX - nMarkWidth/2, Y + nHeight/2 - nMarkHeight/2);
    //         oCtx.fill();

    //         this._markRect = {
    //             x1: (nMarkX - nMarkWidth/2) - ((X + nWidth) - (nMarkX + nMarkWidth/2)),
    //             y1: Y,
    //             x2: X + nWidth,
    //             y2: Y + nHeight
    //         }
    //     }

    //     let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

    //     let contentX = (X + nWidth * 0.02) * g_dKoef_pix_to_mm / scaleCoef;
    //     let contentY = (Y + nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
    //     let contentXLimit = (X + nWidth * 0.98) * g_dKoef_pix_to_mm / scaleCoef;
    //     let contentYLimit = (Y + nHeight - nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        
    //     if (this.type == "checkbox" || this.type == "radiobutton") {
    //         contentY = Y * g_dKoef_pix_to_mm / scaleCoef;
    //         this.ProcessAutoFitContent(); // подгоняем размер галочки
    //     }
    //     else if (this.type == "combobox") {
    //         contentXLimit = this._markRect.x1 * g_dKoef_pix_to_mm / scaleCoef; // ограничиваем контент позицией маркера
    //         let nContentH = this._content.GetElement(0).Get_EmptyHeight();
    //         contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm / scaleCoef - nContentH / 2;
    //     }
    //     else if (this.type == "text" && !this.multiline) {
    //         // выставляем текст посередине
    //         let nContentH = this._content.GetElement(0).Get_EmptyHeight();
    //         contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm / scaleCoef - nContentH / 2;
    //     }

    //     this._formRelRectMM.X = contentX;
    //     this._formRelRectMM.Y = contentY;
    //     this._formRelRectMM.W = contentXLimit - contentX;
    //     this._formRelRectMM.H = contentYLimit - contentY;

    //     if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
    //     contentXLimit != this._oldContentPos.XLimit) {
    //         this._content.X      = this._oldContentPos.X        = contentX;
    //         this._content.Y      = this._oldContentPos.Y        = contentY;
    //         this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
    //         this._content.YLimit = this._oldContentPos.YLimit   = 20000;
    //         this._content.Recalculate_Page(0, true);
    //     }
    //     else if (this._wasChanged) {
    //         this._content.Content.forEach(function(element) {
    //             element.Recalculate_Page(0);
    //         });
    //         this._wasChanged = false;
    //     }
        
    //     if (this.type == "listbox" || (this.type == "text" && this._multiline == true)) {
    //         this._content.ResetShiftView();
    //         this._content.ShiftView(this._curShiftView.x, this._curShiftView.y);
    //     }

    //     if (this.type == "text" || this.type == "combobox" || this._needShiftContentView)
    //         this.CheckFormViewWindow();

    //     let oGraphics = new AscCommon.CGraphics();
    //     let widthPx = oViewer.canvas.width;
    //     let heightPx = oViewer.canvas.height;
        
    //     oGraphics.init(oCtx, widthPx * scaleCoef, heightPx * scaleCoef, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
	// 	oGraphics.m_oFontManager = AscCommon.g_fontManager;
	// 	oGraphics.endGlobalAlphaColor = [255, 255, 255];
    //     oGraphics.transform(1, 0, 0, 1, 0, 0);
        
    //     oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

    //     this._content.Draw(0, oGraphics);
    //     // redraw target cursor if field is selected
    //     if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && (oViewer.fieldFillingMode || this.type == "combobox"))
    //         this._content.RecalculateCurPos();
        
    //     oGraphics.RemoveClip();
    //     this._pageIndX = pageIndX;
    //     this._pageIndY = pageIndY;
        
    //     // save pos in page.
    //     this._pagePos = {
    //         x: X - pageIndX,
    //         y: Y - pageIndY,
    //         w: nWidth,
    //         h: nHeight,
    //         realX: X,
    //         realY: Y
    //     };

    //     if (this.type == "listbox" || this.type == "text" || this._doNotScroll == false)
    //         this.private_updateScroll(true);
    // };
    CBaseField.prototype.DrawHighlight = function(oCtx) {
        oCtx = editor.getDocumentRenderer().canvasFormsHighlight.getContext("2d");
        oCtx.fillStyle = `rgb(${FIELDS_HIGHLIGHT.r}, ${FIELDS_HIGHLIGHT.g}, ${FIELDS_HIGHLIGHT.b})`;
        oCtx.fillRect(this._pagePos.realX, this._pagePos.realY, this._pagePos.w, this._pagePos.h);
    };
    CBaseField.prototype.private_GetType = function() {
        return this.type;
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

                let aFields = this._doc.getWidgetsByName(this.name);
                for (var i = 0; i < aFields.length; i++) {
                    aFields[i]._exportValues = arrValues;
                    if (arrValues[i])
                        aFields[i]._exportValue = arrValues[i];
                    else
                        aFields[i]._exportValue = "Yes";
                }
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
        }
        
    });

    CBaseCheckBoxField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = editor.getDocumentRenderer();

        function round(nValue) {
            return (nValue + 0.5) >> 0;
        }
        
        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        switch (this._borderStyle) {
            case "solid":
                //oCtx.setLineDash([]);
                break;
            case "beveled":
                break;
            case "dashed":
                //oCtx.setLineDash([5 * oViewer.zoom]);
                break;
            case "inset":
                break;
            case "underline":
                break;
        }

        // draw border
        oCtx.beginPath();
        oCtx.rect(X, Y, nWidth, nHeight);
        oCtx.stroke();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        
        // подгоняем размер галочки
        contentY = Y * g_dKoef_pix_to_mm / scaleCoef;
        this.ProcessAutoFitContent(); 

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
    Object.defineProperties(CRadioButtonField.prototype, {
        "value": {
            set(sValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (this._exportValues.includes(sValue)) {
                    aFields.forEach(function(field) {
                        field._value = sValue;
                    });
                }
                else {
                    aFields.forEach(function(field) {
                        field._value = "Off";
                    });
                }
                editor.getDocumentRenderer()._paintForms();
            },
            get() {
                return this._value;
            }
        }
    });

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

            field._wasChanged = true;
        });
        
        editor.getDocumentRenderer()._paintForms();
    };
    
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.private_syncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let nThisIdx = aFields.indexOf(this);
        
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++)
                    this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                
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

    function CRadioButtonField(sName, nPage, aRect)
    {
        CBaseCheckBoxField.call(this, sName, FIELD_TYPE.radiobutton, nPage, aRect);
        
        let oRun = this._content.GetElement(0).GetElement(0);
        //oRun.AddText(String.fromCharCode(UncheckedSymbol));
        oRun.AddText("〇");

        this._radiosInUnison = false;
        this._style = style.ci;
    }
    CRadioButtonField.prototype = Object.create(CBaseCheckBoxField.prototype);
	CRadioButtonField.prototype.constructor = CRadioButtonField;
    Object.defineProperties(CRadioButtonField.prototype, {
        "radiosInUnison": {
            set(bValue) {
                if (typeof(bValue) == "boolean") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._radiosInUnison = bValue;
                    });
                    this.private_UpdateAll();
                }
            },
            get() {
                return this._radiosInUnison;
            }
        },
        "value": {
            set(sValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (this._exportValues.includes(sValue)) {
                    aFields.forEach(function(field) {
                        field._value = sValue;
                    });
                }
                else {
                    aFields.forEach(function(field) {
                        field._value = "Off";
                    });
                }
                this.private_UpdateAll();
            },
            get() {
                return this._value;
            }
        }
    });
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.private_syncField = function() {
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
                    this._content.Internal_Content_RemoveAll();
                    for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++)
                        this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                
                    break;
                }
            }
        }
    };
    CRadioButtonField.prototype.onMouseDown = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThis = this;
        if (false == this._radiosInUnison) {
            if (this._value != "Off") {
                return;
            }
            else {
                this.private_SetChecked(true);
                this._wasChanged = true;
            }

            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field._value != "Off") {
                    field.private_SetChecked(false);
                    field._wasChanged = true;
                }
            });
        }
        else {
            aFields.forEach(function(field) {
                if (field._exportValue != oThis._exportValue) {
                    field.private_SetChecked(false);
                    field._wasChanged = true;
                }
                else {
                    field.private_SetChecked(true);
                    field._wasChanged = true;
                }
            });
        }
        
        editor.getDocumentRenderer()._paintForms();
    };
    /**
	 * Updates all field with this field name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.private_UpdateAll = function() {
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
                    field.private_SetChecked(false);
                });
            }
            else {
                aFields.forEach(function(field) {
                    if (field._exportValue != sExportValue) {
                        field.private_SetChecked(false);
                    }
                    else {
                        field.private_SetChecked(true);
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
                    aFields[i].private_SetChecked(false);
                }
            }
        }
    };
    /**
	 * Set checked to this field (not for all with the same name).
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.private_SetChecked = function(isChecked) {
        let oRun = this._content.GetElement(0).GetElement(0);
        if (isChecked) {
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
        this._content = new AscWord.CDocumentContent(null, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, undefined, undefined, false);
        this._content.ParentPDF = this;
        this._content.SetUseXLimit(false);

        this._scrollInfo = null;
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
                if (typeof(nValue) == "number") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._calcOrderIndex = nValue;
                    });
                }
            },
            get() {
                return this._calcOrderIndex;
            }
        },
        "charLimit": {
            set(nValue) {
                if (typeof(nValue) == "number" && nValue <= 500 && nValue > 0 && this.fileSelect === false) {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    nValue = Math.round(nValue);
                    if (this._charLimit != nValue) {
                        let aChars = [];
                        let sText = this._content.GetElement(0).GetText();
                        for (let i = 0; i < sText.length; i++) {
                            aChars.push(sText[i].charCodeAt(0));
                        }

                        aFields.forEach(function(field) {
                            field._charLimit = nValue;
                            field._content.SelectAll();
                            field.EnterText(aChars);
                        });

                        editor.getDocumentRenderer()._paintForms();
                    }
                }
            },
            get() {
                return this._charLimit;
            }
        },
        "comb": {
            set(bValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (bValue === true) {
                    aFields.forEach(function(field) {
                        field._comb = true;
                        field._doNotScroll = true;
                    });
                    
                }
                else if (bValue === false) {
                    aFields.forEach(function(field) {
                        field._comb = false;
                    });
                }
            },
            get() {
                return this._comb;
            }
        },
        "doNotScroll": {
            set(bValue) {
                if (typeof(bValue) === "boolean") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._doNotScroll = bValue;
                        if (editor.getDocumentRenderer().mouseDownFieldObject == field) {
                            if (bValue == true)
                                editor.getDocumentRenderer().mouseDownFieldObject.private_updateScroll(false, false);
                            else
                                editor.getDocumentRenderer().mouseDownFieldObject.private_updateScroll();
                        }
                            
                    });
                }
            },
            get() {
                return this._doNotScroll;
            }
        },
        "doNotSpellCheck": {
            set(bValue) {
                if (typeof(bValue) === "boolean") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._doNotSpellCheck = bValue;
                    });
                }
            },
            get() {
                return this._doNotSpellCheckl;
            }
        },
        "fileSelect": {
            set(bValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (bValue === true && this.multiline == false && this.charLimit === 0
                    && this.password == false && this.defaultValue == "") {
                        aFields.forEach(function(field) {
                            field._fileSelect = true;
                        });
                    }
                else if (bValue === false) {
                    aFields.forEach(function(field) {
                        field._fileSelect = false;
                    });
                }
            },
            get() {
                return this._fileSelect;
            }
        },
        "multiline": {
            set(bValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (bValue === true && this.fileSelect === false) {
                    aFields.forEach(function(field) {
                        field._content.SetUseXLimit(true);
                        field._multiline = true;
                    });
                }
                else if (bValue === false) {
                    aFields.forEach(function(field) {
                        field._content.SetUseXLimit(false);
                        field._multiline = false;
                    });
                }
                editor.getDocumentRenderer()._paintForms();
            },
            get() {
                return this._multiline;
            }
        },
        "password": {
            set (bValue) {
                let aFields = this._doc.getWidgetsByName(this.name);
                if (bValue === true && this.fileSelect === false) {
                    aFields.forEach(function(field) {
                        field._password = true;
                    });
                }
                else if (bValue === false) {
                    aFields.forEach(function(field) {
                        field._password = false;
                    });
                }
            },
            get() {
                return this._password;
            }
        },
        "richText": {
            set(bValue) {
                if (typeof(bValue) == "boolean") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._richText = bValue;
                    });
                }
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

                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._richValue = aCorrectVals;
                    });
                }
            },
            get() {
                return this._richValue;
            }
        },
        "textFont": {
            set(sValue) {
                if (typeof(sValue) == "string" && sValue !== "") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._textFont = sValue;
                    });
                }
            },
            get() {
                return this.textFont;
            }
        },
        "value": {
            set(sText) {
                if (typeof(sText) != "string")
                    return;
                
                let aChars = [];
                for (let i = 0; i < sText.length; i++) {
                    aChars.push(sText[i].charCodeAt(0));
                }

                let aFields = this._doc.getWidgetsByName(this.name);
                aFields.forEach(function(field) {
                    field._content.SelectAll();
                    field.EnterText(aChars);
                });

                editor.getDocumentRenderer()._paintForms();
            }
        }
    });
    
    CTextField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = editor.getDocumentRenderer();

        function round(nValue) {
            return (nValue + 0.5) >> 0;
        }
        
        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        switch (this._borderStyle) {
            case "solid":
                //oCtx.setLineDash([]);
                break;
            case "beveled":
                break;
            case "dashed":
                //oCtx.setLineDash([5 * oViewer.zoom]);
                break;
            case "inset":
                break;
            case "underline":
                break;
        }

        // draw border
        oCtx.beginPath();
        oCtx.rect(X, Y, nWidth, nHeight);
        oCtx.stroke();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        
        if (this.multiline == false) {
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
            this._content.X      = this._oldContentPos.X        = contentX;
            this._content.Y      = this._oldContentPos.Y        = contentY;
            this._content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this._content.YLimit = this._oldContentPos.YLimit   = 20000;
            this._content.Recalculate_Page(0, true);
        }
        else if (true) {
            this._content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
        }
        
        if (this._multiline == true) {
            this._content.ResetShiftView();
            this._content.ShiftView(this._curShiftView.x, this._curShiftView.y);
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
        oGraphics.AddClipRect(this._content.X, this._content.Y, this._content.XLimit - this._content.X, contentYLimit - contentY);

        this._content.Draw(0, oGraphics);

        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && oViewer.fieldFillingMode)
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

        if (this._doNotScroll == false) {
            if (this._wasChanged == false)
                this.private_updateScroll(true);
            else
                this.private_updateScroll(false, true);
        }

        this._wasChanged = false;
    };

    CTextField.prototype.onMouseDown = function(x, y, e) {
        let oViewer = editor.getDocumentRenderer();
                
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
        if (this._doNotScroll == false)
            this.private_updateScroll(false, true);
    };
    CTextField.prototype.private_SelectionSetStart = function(x, y, e) {
        let oViewer = editor.getDocumentRenderer();
        
        let mouseXInPage = x - oViewer.x;
        let mouseYInPage = y - oViewer.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oViewer.zoom;

        this._content.Selection_SetStart(X, Y, 0, e);
    };
    CTextField.prototype.private_SelectionSetEnd = function(x, y, e) {
        let oViewer = editor.getDocumentRenderer();
        
        let mouseXInPage = x - oViewer.x;
        let mouseYInPage = y - oViewer.y;

        let X = mouseXInPage * g_dKoef_pix_to_mm / oViewer.zoom;
        let Y = mouseYInPage * g_dKoef_pix_to_mm / oViewer.zoom;

        this._content.Selection_SetEnd(X, Y, 0, e);
    };

    CTextField.prototype.private_moveCursorLeft = function(isShiftKey, isCtrlKey)
    {
        this._content.MoveCursorLeft(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.private_moveCursorRight = function(isShiftKey, isCtrlKey)
    {
        this._content.MoveCursorRight(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.private_MoveCursorDown = function(isShiftKey, isCtrlKey) {
        this._content.MoveCursorDown(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.private_MoveCursorUp = function(isShiftKey, isCtrlKey) {
        this._content.MoveCursorUp(isShiftKey, isCtrlKey);
        this._needShiftContentView = true && this._doNotScroll == false;
        return this._content.RecalculateCurPos();
    };
    CTextField.prototype.EnterText = function(aChars)
    {
        let oPara = this._content.GetElement(0);
        if (this._content.IsSelectionEmpty())
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

        this._wasChanged = true;
        this._needShiftContentView = true && this._doNotScroll == false;
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.private_applyValueForAll = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        for (let i = 0; i < aFields.length; i++) {
            aFields[i]._content.GetElement(0).MoveCursorToStartPos();

            if (aFields[i] == this)
                continue;

            aFields[i]._content.Internal_Content_RemoveAll();
            for (let nItem = 0; nItem < this._content.Content.length; nItem++)
                aFields[i]._content.Internal_Content_Add(nItem, this._content.Content[nItem].Copy());

            //aFields[i]._wasChanged = true;
            aFields[i]._content.Recalculate_Page(0); // to do check
        }
    };
    
    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.Remove = function(nDirection, bWord) {
        this._content.Remove(nDirection, true, false, false, bWord);
        this._wasChanged = true;
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.private_syncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
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

                if (this._multiline)
                    this._content.SetUseXLimit(true);

                this._content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++)
                    this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                
                break;
            }
        }
    };

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
        "numItems": {
            get() {
                return this._options.length;
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

    /**
	 * Gets the internal value of an item in a combo box or a list box.
	 * @memberof CTextField
     * @param {number} nIdx - The 0-based index of the item in the list or -1 for the last item in the list.
     * @param {boolean} [bExportValue=true] - Specifies whether to return an export value.
	 * @typeofeditors ["PDF"]
     * @returns {string}
	 */
    CBaseListField.prototype.getItemAt = function(nIdx, bExportValue) {
        if (typeof(bExportValue) != "boolean")
            bExportValue = true;

        if (this._options[nIdx]) {
            if (typeof(this._options[nIdx]) == "string")
                return this._options[nIdx];
            else {
                if (bExportValue)
                    return this._options[nIdx][1];

                return this._options[nIdx][0];
            } 
        }
    };

    function CComboBoxField(sName, nPage, aRect)
    {
        CBaseListField.call(this, sName, FIELD_TYPE.combobox, nPage, aRect);

        this._calcOrderIndex    = 0;
        this._doNotSpellCheck   = false;
        this._editable          = false;

        // internal
        this._id = AscCommon.g_oIdCounter.Get_NewId();
    };
    CComboBoxField.prototype = Object.create(CBaseListField.prototype);
	CComboBoxField.prototype.constructor = CComboBoxField;
    Object.defineProperties(CComboBoxField.prototype, {
        "calcOrderIndex": {
            set(nValue) {
                if (typeof(nValue) == "number") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._calcOrderIndex = nValue;
                    });
                }
            },
            get() {
                return this._calcOrderIndex;
            }
        },
        "doNotSpellCheck": {
            set(bValue) {
                if (typeof(bValue) === "boolean") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._doNotSpellCheck = bValue;
                    });
                }
            },
            get() {
                return this._doNotSpellCheckl;
            }
        },
        "editable": {
            set(bValue) {
                if (typeof(bValue) === "boolean") {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._editable = bValue;
                    });
                }
            },
            get() {
                return this._editable;
            }
        },
        "currentValueIndices": {
            set(value) {
                if (typeof(value) === "number" && this.getItemAt(value, false) !== undefined) {
                    let aFields = this._doc.getWidgetsByName(this.name);
                    aFields.forEach(function(field) {
                        field._currentValueIndices = value;
                    });

                    this.private_selectOption(value);
                    this.private_applyValueForAll();
                }
            },
            get() {
                return this._currentValueIndices;
            }
        },
    });

    CComboBoxField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = editor.getDocumentRenderer();

        function round(nValue) {
            return (nValue + 0.5) >> 0;
        }
        
        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        switch (this._borderStyle) {
            case "solid":
                //oCtx.setLineDash([]);
                break;
            case "beveled":
                break;
            case "dashed":
                //oCtx.setLineDash([5 * oViewer.zoom]);
                break;
            case "inset":
                break;
            case "underline":
                break;
        }

        // draw border
        oCtx.beginPath();
        oCtx.rect(X, Y, nWidth, nHeight);
        oCtx.stroke();

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

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        
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
        // redraw target cursor if field is selected
        if (oViewer.mouseDownFieldObject == this && this._content.IsSelectionUse() == false && oViewer.fieldFillingMode)
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
    CComboBoxField.prototype.onMouseDown = function(x, y, e) {
        let oViewer = editor.getDocumentRenderer();
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
            this._content.GetElement(0).MoveCursorToStartPos();
        }
        else {
            this._content.Selection_SetStart(XInContent, YInContent, 0, e);
            this._content.RemoveSelection();
        }
        
        this._content.RecalculateCurPos();
    };
    /**
	 * Sets the list of items for a combo box.
	 * @memberof CComboBoxField
     * @param {string[]} values - An array in which each element is either an object convertible to a string or another array:
        For an element that can be converted to a string, the user and export values for the list item are equal to the string.
        For an element that is an array, the array must have two subelements convertible to strings, where the first is the user value and the second is the export value.
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.setItems = function(values) {
        let aOptToPush = [];
        for (let i = 0; i < values.length; i++) {
            if (values[i] == null)
                continue;
            if (typeof(values[i]) == "string" && values[i] != "")
                aOptToPush.push(values[i]);
            else if (Array.isArray(values[i]) && values[i][0] != undefined && values[i][1] != undefined) {
                if (values[i][0].toString && values[i][1].toString) {
                    aOptToPush.push([values[i][0].toString(), values[i][1].toString()])
                }
            }
            else if (typeof(values[i]) != "string" && values[i].toString) {
                aOptToPush.push(values[i].toString());
            }
        }

        let aFields = this._doc.getWidgetsByName(this.name);
        aFields.forEach(function(field) {
            field._options = aOptToPush.slice();
            field.private_selectOption(0);
        });

        editor.getDocumentRenderer()._paintForms();
    };
    CComboBoxField.prototype.private_selectOption = function(nIdx) {
        let oPara = this._content.GetElement(0);
        let oRun = oPara.GetElement(0);
        oRun.ClearContent();

        this._currentValueIndices = nIdx;

        if (Array.isArray(this._options[nIdx]))
            oRun.AddText(this._options[nIdx][0]);
        else
            oRun.AddText(this._options[nIdx]);

        this._wasChanged = true;
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.private_syncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {

                this._calcOrderIndex    = aFields[i]._calcOrderIndex;
                this._doNotSpellCheck   = aFields[i]._doNotSpellCheck;
                this._editable          = aFields[i]._editable;

                this._content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++) {
                    this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                }
                
                this._options = aFields[i]._options.slice();
                break;
            }
        }
    };
    CComboBoxField.prototype.EnterText = function(aChars)
    {
        if (this._editable == false)
            return;

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

        this._currentValueIndices = -1;
        this._wasChanged = true;
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
    Object.defineProperties(CListBoxField.prototype, {
        "multipleSelection": {
            set(bValue) {
                if (typeof(bValue) == "boolean") {
                    if (bValue == this.multipleSelection)
                        return;

                    let aFields = this._doc.getWidgetsByName(this.name);
                    if (bValue == true) {
                        aFields.forEach(function(field) {
                            field._multipleSelection = true;
                            field._currentValueIndices = [field._currentValueIndices];
                        });
                    }
                    else {
                        aFields.forEach(function(field) {
                            field._multipleSelection = false;
                            field._currentValueIndices = field._currentValueIndices[0];
                            field.private_selectOption(field._currentValueIndices, true);
                        });
                    }
                }
            },
            get() {
                return this._multipleSelection;
            }
        },
        "currentValueIndices": {
            set(value) {
                if (Array.isArray(value) && this.multipleSelection === true)
                {
                    let isValid = true;
                    for (let i = 0; i < value.length; i++) {
                        if (typeof(value[i]) != "number" || this.getItemAt(value[i], false) === undefined) {
                            isValid = false;
                            break;
                        }
                    }

                    if (isValid) {
                        this._needShiftContentView = true;

                        // снимаем выделение с тех, которые не присутсвуютв новых значениях (value)
                        for (let i = 0; i < this._currentValueIndices.length; i++) {
                            if (value.includes(this._currentValueIndices[i]) == false) {
                                this.private_unselectOption(this._currentValueIndices[i]);
                            }
                        }
                        
                        for (let i = 0; i < value.length; i++) {
                            // добавляем выделение тем, которые не присутсвуют в текущем поле
                            if (this._currentValueIndices.includes(value[i]) == false) {
                                this.private_selectOption(value[i], false);
                            }
                        }
                        this._currentValueIndices = value.sort();
                        this.private_applyValueForAll();
                    }
                }
                else if (this.multipleSelection === false && typeof(value) === "number" && this.getItemAt(value, false) !== undefined) {
                    this._currentValueIndices = value;
                    this.private_selectOption(value, true);
                    this.private_applyValueForAll();
                }
            },
            get() {
                return this._currentValueIndices;
            }
        },
    });

    CListBoxField.prototype.Draw = function(oCtx, pageIndX, pageIndY) {
        let oViewer = editor.getDocumentRenderer();

        function round(nValue) {
            return (nValue + 0.5) >> 0;
        }
        
        let X = pageIndX + (this._rect[0] * oViewer.zoom);
        let Y = pageIndY + (this._rect[1] * oViewer.zoom);
        let nWidth = (this._rect[2] - this._rect[0]) * oViewer.zoom;
        let nHeight = (this._rect[3] - this._rect[1]) * oViewer.zoom;

        switch (this._borderStyle) {
            case "solid":
                //oCtx.setLineDash([]);
                break;
            case "beveled":
                break;
            case "dashed":
                //oCtx.setLineDash([5 * oViewer.zoom]);
                break;
            case "inset":
                break;
            case "underline":
                break;
        }

        oCtx.beginPath();
        oCtx.rect(X, Y, nWidth, nHeight);
        oCtx.stroke();

        let scaleCoef = oViewer.zoom * AscCommon.AscBrowser.retinaPixelRatio;

        let contentX = (X + nWidth * 0.02) * g_dKoef_pix_to_mm / scaleCoef;
        let contentY = (Y + nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        let contentXLimit = (X + nWidth * 0.98) * g_dKoef_pix_to_mm / scaleCoef;
        let contentYLimit = (Y + nHeight - nWidth * 0.01) * g_dKoef_pix_to_mm / scaleCoef;
        
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

        this.private_updateScroll(true);
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.private_syncField = function() {
        let aFields = this._doc.getWidgetsByName(this.name);
        
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._multipleSelection = aFields[i]._multipleSelection;
                this._content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i]._content.Content.length; nItem++) {
                    this._content.Internal_Content_Add(nItem, aFields[i]._content.Content[nItem].Copy());
                }
                
                this._options = aFields[i]._options.slice();
                break;
            }
        }
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.private_applyValueForAll = function(oFieldToSkip) {
        let aFields = this._doc.getWidgetsByName(this.name);
        let oThis = this;
        
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
                        field.private_unselectOption(field._currentValueIndices[i]);
                    }
                }
                
                for (let i = 0; i < oThis._currentValueIndices.length; i++) {
                    // добавляем выделение тем, которые не присутсвуют в текущем поле, но присутсвуют в том, от которого применяем
                    if (field._currentValueIndices.includes(oThis._currentValueIndices[i]) == false) {
                        field.private_selectOption(oThis._currentValueIndices[i], false);
                    }
                }
                field._currentValueIndices = oThis._currentValueIndices.slice();
            }
            else {
                field._currentValueIndices = oThis._currentValueIndices;
                field.private_selectOption(field._currentValueIndices, true);
            }
        });
    };
    /**
	 * Sets the list of items for a list box.
	 * @memberof CListBoxField
     * @param {string[]} values - An array in which each element is either an object convertible to a string or another array:
        For an element that can be converted to a string, the user and export values for the list item are equal to the string.
        For an element that is an array, the array must have two subelements convertible to strings, where the first is the user value and the second is the export value.
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.setItems = function(values) {
        let aFields = this._doc.getWidgetsByName(this.name);

        aFields.forEach(function(field) {
            field._options = [];
            field._content.Internal_Content_RemoveAll();
            let sCaption, oPara, oRun;
            
            for (let i = 0; i < values.length; i++) {
                if (values[i] == null)
                    continue;
                sCaption = "";
                if (typeof(values[i]) == "string" && values[i] != "") {
                    sCaption = values[i];
                    field._options.push(values[i]);
                }
                else if (Array.isArray(values[i]) && values[i][0] != undefined && values[i][1] != undefined) {
                    if (values[i][0].toString && values[i][1].toString) {
                        field._options.push([values[i][0].toString(), values[i][1].toString()]);
                        sCaption = values[i][0].toString();
                    }
                }
                else if (typeof(values[i]) != "string" && values[i].toString) {
                    field._options.push(values[i].toString());
                    sCaption = values[i].toString();
                }

                if (sCaption != "") {
                    oPara = new AscCommonWord.Paragraph(field._content.DrawingDocument, field._content, false);
                    oRun = new AscCommonWord.ParaRun(oPara, false);
                    field._content.Internal_Content_Add(i, oPara);
                    oPara.Add(oRun);
                    oRun.AddText(sCaption);
                }
            }

            field._content.Recalculate_Page(0, true);
            field._curShiftView.x = 0;
            field._curShiftView.y = 0;
        });
        if (aFields.length > 0) {
            this.private_selectOption(0, true);
            if (this._multipleSelection)
                this._currentValueIndices = [0];
            else
                this._currentValueIndices = 0;
            this.private_applyValueForAll(this);
        }
    };

    CListBoxField.prototype.private_selectOption = function(nIdx, isSingleSelect) {
        let oPara = this._content.GetElement(nIdx);
        let oApiPara;
        
        this._content.Content.forEach(function(para) {
            oApiPara = editor.private_CreateApiParagraph(para);
            if (oApiPara.Paragraph.CompiledPr.Pr && oApiPara.Paragraph.CompiledPr.Pr.ParaPr == g_oDocumentDefaultParaPr)
                oApiPara.Paragraph.CompiledPr.Pr.ParaPr = g_oDocumentDefaultParaPr.Copy();
        });
        
        this._content.Set_CurrentElement(nIdx);
        if (isSingleSelect) {
            this._content.Content.forEach(function(para){
                oApiPara = editor.private_CreateApiParagraph(para);
                oApiPara.SetShd('nil');
                
                if (oApiPara.Paragraph.CompiledPr.Pr)
                    oApiPara.Paragraph.CompiledPr.Pr.ParaPr.Shd = oApiPara.Paragraph.Pr.Shd.Copy();
            });
        }

        oApiPara = editor.private_CreateApiParagraph(oPara);
        oApiPara.SetShd('clear', 0, 112, 192);

        oApiPara.Paragraph.CompiledPr.Pr.ParaPr.Shd = oApiPara.Paragraph.Pr.Shd.Copy();
    };
    CListBoxField.prototype.private_unselectOption = function(nIdx) {
        let oApiPara = editor.private_CreateApiParagraph(this._content.GetElement(nIdx));
        oApiPara.SetShd('nil');
        oApiPara.Paragraph.CompiledPr.Pr.ParaPr.Shd = oApiPara.Paragraph.Pr.Shd.Copy();
    };

    CListBoxField.prototype.onMouseDown = function(x, y, e) {
        if (this._options.length == 0)
            return;

        let oHTMLPage = editor.getDocumentRenderer();
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
                    this.private_unselectOption(nPos);
                    this._currentValueIndices.splice(this._currentValueIndices.indexOf(nPos), 1);
                }
                else {
                    this.private_selectOption(nPos, false);
                    this._currentValueIndices.push(nPos);
                    this._currentValueIndices.sort();
                }
            }
            else {
                this.private_selectOption(nPos, true);
                this._currentValueIndices = [nPos];
            }
        }
        else {
            if (nPos == this._currentValueIndices) {
                this.private_updateScroll(false, true);
                return;
            }
                
            this.private_selectOption(nPos, true);
            this._currentValueIndices = nPos;
        }

        this._needShiftContentView = true;

        editor.getDocumentRenderer()._paintForms();
        this.private_updateScroll(false, true);
    };
    CListBoxField.prototype.private_MoveSelectDown = function(event) {
        this._needShiftContentView = true;
        this._content.MoveCursorDown();

        if (this._multipleSelection == true) {
            this.private_selectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = [this._content.CurPos.ContentPos];
        }
        else {
            this.private_selectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = this._content.CurPos.ContentPos;
        }
        
        editor.getDocumentRenderer()._paintForms();
        this.private_updateScroll();
    };
    CListBoxField.prototype.private_MoveSelectUp = function() {
        this._needShiftContentView = true;
        this._content.MoveCursorUp();

        if (this._multipleSelection == true) {
            this.private_selectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = [this._content.CurPos.ContentPos];
        }
        else {
            this.private_selectOption(this._content.CurPos.ContentPos, true);
            this._currentValueIndices = this._content.CurPos.ContentPos;
        }

        editor.getDocumentRenderer()._paintForms();
        this.private_updateScroll();
    };
    CListBoxField.prototype.private_updateScroll = function(bUpdateOnlyPos, bShow) {
        let oContentBounds = this._content.GetContentBounds(0);
        let oFieldBounds = this._content.ParentPDF.getFormRelRect();
        let oScroll, oScrollDocElm, oScrollSettings;
        let nContentH = oContentBounds.Bottom - oContentBounds.Top;
        
        if (nContentH < oFieldBounds.H)
            return;

        if (typeof(bShow) != "boolean" && this._scrollInfo)
            bShow = this._scrollInfo.scroll.canvas.style.display == "none" ? false : true;

        let oViewer = editor.getDocumentRenderer();
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
                oThis.private_scrollVertical(evt.scrollD, evt.maxScrollY);
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
    CListBoxField.prototype.private_scrollVertical = function(scrollY, maxYscroll) {
        let oContentBounds = this._content.GetContentBounds(0);

        let oFormBounds = this.getFormRelRect();
        let nContentH = oContentBounds.Bottom - oContentBounds.Top;
        let nMaxShiftY = oFormBounds.H - nContentH;

        let nScrollCoeff = scrollY / maxYscroll;
        this._curShiftView.y = nMaxShiftY * nScrollCoeff;
        this._needShiftContentView = false;
        this._scrollInfo.scrollCoeff = nScrollCoeff;
        editor.getDocumentRenderer()._paintForms();
    };

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

    CComboBoxField.prototype.private_applyValueForAll   = CTextField.prototype.private_applyValueForAll;
    CComboBoxField.prototype.Remove                     = CTextField.prototype.Remove;
    CComboBoxField.prototype.private_moveCursorLeft     = CTextField.prototype.private_moveCursorLeft;
    CComboBoxField.prototype.private_moveCursorRight    = CTextField.prototype.private_moveCursorRight;
    CTextField.prototype.private_updateScroll           = CListBoxField.prototype.private_updateScroll;
    CTextField.prototype.private_scrollVertical         = CListBoxField.prototype.private_scrollVertical;
    
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
