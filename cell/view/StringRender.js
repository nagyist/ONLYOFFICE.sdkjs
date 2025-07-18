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

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function (window, undefined) {


		/*
		 * Import
		 * -----------------------------------------------------------------------------
		 */
		var asc = window["Asc"];
		var asc_debug   = asc.outputDebugStr;
		var asc_typeof  = asc.typeOf;
		var asc_round   = asc.round;

		function LineInfo(lm) {
			this.tw = 0;
			this.th = 0;
			this.bl = 0;
			this.a = 0;
			this.d = 0;
			this.beg = undefined;
			this.end = undefined;
			this.startX = undefined;

			this.assign(lm);
		}

		LineInfo.prototype.assign = function (lm) {
			if (lm) {
				this.th = lm.th;
				this.bl = lm.bl;
				this.a = lm.a;
				this.d = lm.d;
			}
		};


		LineInfo.prototype.initStartX = function (lineWidth, x, maxWidth, align) {
			var x_ = x;
			if (align === AscCommon.align_Right) {
				x_ = x + maxWidth - lineWidth - 1;
			} else if (align === AscCommon.align_Center) {
				x_ = x + 0.5 * (maxWidth - lineWidth);
			}
			this.startX = x_;
			return x_;
		};

		/** @constructor */
		function lineMetrics() {
			this.th = 0;
			this.bl = 0;
			this.bl2 = 0;
			this.a = 0;
			this.d = 0;
		}

		lineMetrics.prototype.clone = function () {
			var oRes = new lineMetrics();
			oRes.th = this.th;
			oRes.bl = this.bl;
			oRes.bl2 = this.bl2;
			oRes.a = this.a;
			oRes.d = this.d;
			return oRes;
		};

		/** @constructor */
		function charProperties() {
			this.grapheme = AscFonts.NO_GRAPHEME;
			this.c = undefined;
			this.lm = undefined;
			this.fm = undefined;
			this.fsz = undefined;
			this.font = undefined;
			this.va = undefined;
			this.nl = undefined;
			this.hp = undefined;
			this.delta = undefined;
			this.skip = undefined;
			this.repeat = undefined;
			this.total = undefined;
			this.wrd = undefined;
		}

		charProperties.prototype.clone = function () {
			var oRes = new charProperties();
			oRes.grapheme = this.grapheme;
			oRes.c = (undefined !== this.c) ? this.c.clone() : undefined;
			oRes.lm = (undefined !== this.lm) ? this.lm.clone() : undefined;
			oRes.fm = (undefined !== this.fm) ? this.fm.clone() : undefined;
			oRes.font = (undefined !== this.font) ? this.font.clone() : undefined;
			oRes.fsz = this.fsz;
			oRes.va = this.va;
			oRes.nl = this.nl;
			oRes.hp = this.hp;
			oRes.delta = this.delta;
			oRes.skip = this.skip;
			oRes.repeat = this.repeat;
			oRes.total = this.total;
			oRes.wrd = this.wrd;
			return oRes;
		};
		
		/**
		 *
		 * @constructor
		 */
		function FragmentShaper() {
			AscFonts.CTextShaper.call(this);
			
			this.font           = null;
			this.stringRenderer = null;
		}
		FragmentShaper.prototype = Object.create(AscFonts.CTextShaper.prototype);
		FragmentShaper.prototype.constructor = FragmentShaper;
		FragmentShaper.prototype.GetCodePoint = function(oItem)
		{
			return oItem.char;
		};
		FragmentShaper.prototype.GetTextScript = function(nUnicode)
		{
			if (0x060C <= nUnicode && nUnicode <= 0x074A)
				return AscFonts.HB_SCRIPT.HB_SCRIPT_ARABIC;

			return AscFonts.hb_get_script_by_unicode(nUnicode);
		};
		FragmentShaper.prototype.shapeFragment = function(chars, font, stringRenderer, beginIndex) {
			this.font           = font;
			this.stringRenderer = stringRenderer;

			this.StartString();
			for (let i = 0; i < chars.length; ++i) {
				let char = chars[i];
				let isNL = stringRenderer.codesHypNL[char];
				let isSP = !isNL ? stringRenderer.codesHypSp[char] : false;

				if (isNL || isSP) {
					this.FlushWord();
					this.AppendToString({idx: beginIndex + i, char: char});
				} else {
					this.AppendToString({idx: beginIndex + i, char: char});
				}
			}

			this.FlushWord();
		};
		FragmentShaper.prototype.GetFontSlot = function() {
			return AscWord.fontslot_ASCII;
		};
		FragmentShaper.prototype.GetFontInfo = function() {
			return {
				Name  : this.font.getName(),
				Size  : this.font.getSize(),
				Style : (this.font.getBold() ? 1 : 0) | (this.font.getItalic() ? 2 : 0)
			};
		};
		FragmentShaper.prototype.FlushGrapheme = function(grapheme, width, codePointCount, isLigature) {
			if (codePointCount <= 0)
				return;

			let charIndex = 0;

			if (this.IsRtlDirection())
			{
				if (this.BufferIndex - codePointCount < 0)
					return;

				this.BufferIndex -= codePointCount;
				charIndex = this.BufferIndex;
			}
			else
			{
				if (this.BufferIndex + codePointCount - 1 >= this.Buffer.length)
					return;

				charIndex = this.BufferIndex;
				this.BufferIndex += codePointCount;
			}
			let _width = asc_round(width * this.font.getSize() / 25.4 * this.stringRenderer.drawingCtx.getPPIY());

			let w = Math.trunc(_width / codePointCount);
			let r = Math.max(0, _width - w * codePointCount);


			if (1 === codePointCount)
			{
				this.private_HandleItem(this.Buffer[charIndex], grapheme, w);
			}
			else
			{
				if (this.IsRtlDirection())
				{
					this.private_HandleItem(this.Buffer[charIndex], AscFonts.NO_GRAPHEME, w);
					this.private_HandleItem(this.Buffer[charIndex + codePointCount - 1], grapheme, w);
				}
				else
				{
					this.private_HandleItem(this.Buffer[charIndex], grapheme, w);
					this.private_HandleItem(this.Buffer[charIndex + codePointCount - 1], AscFonts.NO_GRAPHEME, w);
				}

				for (let nIndex = 1; nIndex < codePointCount - 1; ++nIndex)
				{
					++charIndex;
					this.private_HandleItem(this.Buffer[charIndex], AscFonts.NO_GRAPHEME, w + (r ? 1 : 0));
					if (r)
						--r;
				}
			}
		};
		FragmentShaper.prototype.private_HandleItem = function(oItem, grapheme, w) {

			let st = this.stringRenderer;
			let pr = st._getCharPropAt(oItem.idx);
			pr.grapheme = grapheme;
			pr.idx = oItem.idx;
			st.charWidths[oItem.idx] = w;
			st.chars[oItem.idx] = oItem.char;

		};

		/**
		 * Formatted text render
		 * -----------------------------------------------------------------------------
		 * @constructor
		 * @param {DrawingContext} drawingCtx  Context for drawing on
		 *
		 * @memberOf Asc
		 */
		function StringRender(drawingCtx) {
			this.drawingCtx = drawingCtx;
			this.fragmentShaper = new FragmentShaper();

			this.drawState = new TableCellDrawState(this);

			/** @type Array */
			this.fragments = undefined;

			/** @type Object */
			this.flags = undefined;

			/** @type String */
			this.chars = [];
			this.charWidths = [];
			this.charProps = [];
			this.lines = [];
			this.angle = 0;

			this.codesNL = {0xD: 1, 0xA: 1};

			this.codesSpace = {
				0xA: 1,
				0xD: 1,
				0x2028: 1,
				0x2029: 1,
				0x9: 1,
				0xB: 1,
				0xC: 1,
				0x0020: 1,
				0x2000: 1,
				0x2001: 1,
				0x2002: 1,
				0x2003: 1,
				0x2004: 1,
				0x2005: 1,
				0x2006: 1,
				0x2008: 1,
				0x2009: 1,
				0x200A: 1,
				0x200B: 1,
				0x205F: 1,
				0x3000: 1
			};

			this.codesReplaceNL = {};

			this.codesHypNL = {
				0xA: 1, 0xD: 1, 0x2028: 1, 0x2029: 1
			};

			this.codesHypSp = {
				0x9: 1,
				0xB: 1,
				0xC: 1,
				0x0020: 1,
				0x2000: 1,
				0x2001: 1,
				0x2002: 1,
				0x2003: 1,
				0x2004: 1,
				0x2005: 1,
				0x2006: 1,
				0x2008: 1,
				0x2009: 1,
				0x200A: 1,
				0x200B: 1,
				0x205F: 1,
				0x3000: 1
			};

			this.codesHyphen = {
				0x002D: 1, 0x00AD: 1, 0x2010: 1, 0x2012: 1, 0x2013: 1, 0x2014: 1
			};


			// For replacing invisible chars while rendering
			/** @type RegExp */
			this.reNL = /[\r\n]/;
			/** @type RegExp */
			//this.reSpace = /[\n\r\u2028\u2029\t\v\f\u0020\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2008\u2009\u200A\u200B\u205F\u3000]/;
			/** @type RegExp */
			this.reReplaceNL = /\r?\n|\r/g;

			// For hyphenation
			/** @type RegExp */
			//this.reHypNL =  /[\n\r\u2028\u2029]/;
			/** @type RegExp */
			//this.reHypSp =  /[\t\v\f\u0020\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2008\u2009\u200A\u200B\u205F\u3000]/;
			/** @type RegExp */
			//this.reHyphen = /[\u002D\u00AD\u2010\u2012\u2013\u2014]/;

			return this;
		}

		/**
		 * Setups one or more strings to process on
		 * @param {String|Array} fragments  A simple string or array of formatted strings AscCommonExcel.Fragment
		 * @param {AscCommonExcel.CellFlags} flags  Optional.
		 * @return {StringRender}  Returns 'this' to allow chaining
		 */
		StringRender.prototype.setString = function (fragments, flags) {
			this.fragments = [];
			if (asc_typeof(fragments) === "string") {
				var newFragment = new AscCommonExcel.Fragment();
				newFragment.setFragmentText(fragments);
				newFragment.format = new AscCommonExcel.Font();
				this.fragments.push(newFragment);
			} else {
				for (var i = 0; i < fragments.length; ++i) {
					this.fragments.push(fragments[i].clone());
				}
			}
			this.flags = flags;
			this._reset();
			this._setFont(this.drawingCtx, AscCommonExcel.g_oDefaultFormat.Font);
			return this;
		};

		/**
		 * Применяем только трансформации поворота в области
		 * @param {drawingCtx} drawingCtx
		 * @param {type} angle Угол поворота в градусах
		 * @param {Number} x
		 * @param {Number} y
		 * @param {Number} dx
		 * @param {Number} dy
		 * */
		StringRender.prototype.rotateAtPoint = function (drawingCtx, angle, x, y, dx, dy) {
			var m = new asc.Matrix();
			m.rotate(angle, 0);
			var mbt = new asc.Matrix();

			if (null === drawingCtx) {
				mbt.translate(x + dx, y + dy);

				this.drawingCtx.setTextTransform(m.sx, m.shy, m.shx, m.sy, m.tx, m.ty);
				this.drawingCtx.setTransform(mbt.sx, mbt.shy, mbt.shx, mbt.sy, mbt.tx, mbt.ty);
				this.drawingCtx.updateTransforms();
			} else {

				mbt.translate((x + dx) * AscCommonExcel.vector_koef, (y + dy) * AscCommonExcel.vector_koef);
				mbt.multiply(m, 0);

				drawingCtx.setTransform(mbt.sx, mbt.shy, mbt.shx, mbt.sy, mbt.tx, mbt.ty);
			}

			return this;
		};

		StringRender.prototype.resetTransform = function (drawingCtx) {
			if (null === drawingCtx) {
				this.drawingCtx.resetTransforms();
			} else {
				var m = new asc.Matrix();
				drawingCtx.setTransform(m.sx, m.shy, m.shx, m.sy, m.tx, m.ty);
			}

			this.angle = 0;
		};

		/**
		 * @param {Number} angle
		 * @param {Number} w
		 * @param {Number} h
		 * @param {Number} textW
		 * @param {String} alignHorizontal
		 * @param {String} alignVertical
		 * @param {Number} maxWidth
		 */
		StringRender.prototype.getTransformBound = function (angle, w, h, textW, alignHorizontal, alignVertical, maxWidth) {
			var ctx = this.drawingCtx;

			// TODO: добавить padding по сторонам

			this.angle = 0;  //  angle;

			var dx = 0, dy = 0, offsetX = 0,    // смещение BB

				tm = this._doMeasure(maxWidth),

				mul = (90 - (Math.abs(angle))) / 90,

				angleSin = Math.sin(angle * Math.PI / 180.0),
				angleCos = Math.cos(angle * Math.PI / 180.0),

				posh = (angle === 90 || angle === -90) ? textW : Math.abs(angleSin * textW),
				posv = (angle === 90 || angle === -90) ? 0 : Math.abs(angleCos * textW),

				isHorzLeft = (AscCommon.align_Left === alignHorizontal),
				isHorzCenter = (AscCommon.align_Center === alignHorizontal),
				isHorzRight = (AscCommon.align_Right === alignHorizontal),

				isVertBottom = (Asc.c_oAscVAlign.Bottom === alignVertical),
				isVertCenter = (Asc.c_oAscVAlign.Center === alignVertical || Asc.c_oAscVAlign.Dist === alignVertical || Asc.c_oAscVAlign.Just === alignVertical),
				isVertTop = (Asc.c_oAscVAlign.Top === alignVertical);


			var _height = tm.height * ctx.getZoom();
			if (isVertBottom) {
				if (angle < 0) {
					if (isHorzLeft) {
						dx = -(angleSin * _height);
					} else if (isHorzCenter) {
						dx = (w - angleSin * _height - posv) / 2;
						offsetX = -(w - posv) / 2 - angleSin * _height / 2;
					} else if (isHorzRight) {
						dx = w - posv + 2;
						offsetX = -(w - posv) - angleSin * _height - 2;
					}
				} else {
					if (isHorzLeft) {

					} else if (isHorzCenter) {
						dx = (w - angleSin * _height - posv) / 2;
						offsetX = -(w - posv) / 2 + angleSin * _height / 2;
					} else if (isHorzRight) {
						dx = w - posv + 1 + 1 - _height * angleSin;
						offsetX = -w - posv + 1 + 1 - _height * angleSin;
					}
				}

				if (posh < h) {
					if (angle < 0) {
						dy = h - (posh + angleCos * _height);
					} else {
						dy = h - angleCos * _height;
					}
				} else {
					if (angle > 0) {
						dy = h - angleCos * _height;
					}
				}
			} else if (isVertCenter) {

				if (angle < 0) {
					if (isHorzLeft) {
						dx = -(angleSin * _height);
					} else if (isHorzCenter) {
						dx = (w - angleSin * _height - posv) / 2;
						offsetX = -(w - posv) / 2 - angleSin * _height / 2;
					} else if (isHorzRight) {
						dx = w - posv + 2;
						offsetX = -(w - posv) - angleSin * _height - 2;
					}
				} else {
					if (isHorzLeft) {

					} else if (isHorzCenter) {
						dx = (w - angleSin * _height - posv) / 2;
						offsetX = -(w - posv) / 2 + angleSin * _height / 2;
					} else if (isHorzRight) {
						dx = w - posv + 1 + 1 - _height * angleSin;
						offsetX = -w - posv + 1 + 1 - _height * angleSin;
					}
				}

				//

				if (posh < h) {
					if (angle < 0) {
						dy = (h - posh - angleCos * _height) * 0.5;
					} else {
						dy = (h + posh - angleCos * _height) * 0.5;
					}
				} else {
					if (angle > 0) {
						dy = h - angleCos * _height;
					}
				}
			} else if (isVertTop) {

				if (angle < 0) {
					if (isHorzLeft) {
						dx = -(angleSin * _height);
					} else if (isHorzCenter) {
						dx = (w - angleSin * _height - posv) / 2;
						offsetX = -(w - posv) / 2 - angleSin * _height / 2;
					} else if (isHorzRight) {
						dx = w - posv + 2;
						offsetX = -(w - posv) - angleSin * _height - 2;
					}
				} else {
					if (isHorzLeft) {
					} else if (isHorzCenter) {
						dx = (w - angleSin * _height - posv) / 2;
						offsetX = -(w - posv) / 2 + angleSin * _height / 2;
					} else if (isHorzRight) {
						dx = w - posv + 1 + 1 - _height * angleSin;
						offsetX = -w - posv + 1 + 1 - _height * angleSin;
					}

					dy = Math.min(h + _height * angleCos, posh);
				}
			}

			var bound = {dx: dx, dy: dy, height: 0, width: 0, offsetX: offsetX};

			if (angle === 90 || angle === -90) {
				bound.width = _height;
				bound.height = textW;
			} else {
				bound.height = Math.abs(angleSin * textW) + Math.abs(angleCos * _height);
				bound.width = Math.abs(angleCos * textW) + Math.abs(angleSin * _height);
			}

			return bound;
		};

		/**
		 * Measures string that was setup by 'setString' method
		 * @param {Number} maxWidth  Optional. Text width restriction
		 * @return {Asc.TextMetrics}  Returns text metrics or null. @see Asc.TextMetrics
		 */
		StringRender.prototype.measure = function (maxWidth) {
			return this._doMeasure(maxWidth);
		};

		/**
		 * Draw string that was setup by methods 'setString' or 'measureString'
		 * @param {drawingCtx} drawingCtx
		 * @param {Number} x  Left of the text rect
		 * @param {Number} y  Top of the text rect
		 * @param {Number} maxWidth  Text width restriction
		 * @param {String} textColor  Default text color for formatless string
		 * @return {StringRender}  Returns 'this' to allow chaining
		 */
		StringRender.prototype.render = function (drawingCtx, x, y, maxWidth, textColor) {
			this._doRender(drawingCtx, x, y, maxWidth, textColor);
			return this;
		};

		/**
		 * Measures string
		 * @param {String|Array} fragments  A simple string or array of formatted strings AscCommonExcel.Fragment
		 * @param {AscCommonExcel.CellFlags} [flags]      Optional.
		 * @param {Number} [maxWidth]   Optional. Text width restriction
		 * @return {Asc.TextMetrics}  Returns text metrics or null. @see Asc.TextMetrics
		 */
		StringRender.prototype.measureString = function (fragments, flags, maxWidth) {
			if (fragments) {
				this.setString(fragments, flags);
			}
			return this._doMeasure(maxWidth);
		};

		/**
		 * Returns the width of the widest char in the string has been measured
		 */
		StringRender.prototype.getWidestCharWidth = function () {
			return this.charWidths.reduce(function (p, c) {
				return p < c ? c : p;
			}, 0);
		};

		StringRender.prototype._reset = function () {
			this.chars = [];
			this.charWidths = [];
			this.charProps = [];
			this.lines = [];
		};

		/**
		 * @param {String} fragment
		 * @param {Boolean} wrap
		 * @return {String}  Returns filtered fragment
		 */
		StringRender.prototype._filterText = function (fragment, wrap) {
			var s = fragment;
			if (s.search(this.reNL) >= 0) {
				s = s.replace(this.reReplaceNL, wrap ? "\n" : "");
			}
			return s;
		};

		StringRender.prototype._filterChars = function (chars, wrap) {
			var res = [];
			if (chars) {
				for (var i = 0; i < chars.length; i++) {
					if (0xD === chars[i] && 0xA === chars[i + 1]) {
						//\r\n
						if (wrap) {
							res.push(0xA);
						}
						i++;
					} else if (0xA === chars[i]) {
						//\r
						if (wrap) {
							res.push(0xA);
						}
					} else {
						res.push(chars[i]);
					}
				}
			}
			return res;
		};

		/**
		 * @param {Number} startCh
		 * @param {Number} endCh
		 * @return {Number}
		 */
		StringRender.prototype._calcCharsWidth = function (startCh, endCh) {
			for (var w = 0, i = startCh; i <= endCh; ++i) {
				w += this.charWidths[i];
			}
			return w;
		};
		
		/**
		 * @param {Number} startPos
		 * @param {Number} endPos
		 * @return {Number}
		 */
		StringRender.prototype._calcLineWidth = function (startPos, endPos) {
			var wrap = this.flags && (this.flags.wrapText || this.flags.wrapOnlyNL || this.flags.wrapOnlyCE);
			var isAtEnd, j, chProp, tw;

			if (endPos === undefined || endPos < 0) {
				// search for end of line
				for (j = startPos + 1; j < this.chars.length; ++j) {
					chProp = this.charProps[j];
					if (chProp && (chProp.nl || chProp.hp)) {
						break;
					}
				}
				endPos = j - 1;
			}

			for (j = endPos, tw = 0, isAtEnd = true; j >= startPos; --j) {
				if (isAtEnd) {
					// skip space char at end of line
					if ((wrap) && this.codesSpace[this.chars[j]]) {
						continue;
					}
					isAtEnd = false;
				}
				tw += this.charWidths[j];
			}

			return tw;
		};

		StringRender.prototype._calcLineMetrics = function (f, va, fm) {
			var l = new lineMetrics();

			if (!va) {
				var _a = Math.max(0, asc.ceil(fm.nat_y1 * f / fm.nat_scale));
				var _d = Math.max(0, asc.ceil(-fm.nat_y2 * f / fm.nat_scale)) + 1; // 1 px for border

				l.th = _a + _d;
				l.bl = _a;
				l.a = _a;
				l.d = _d;
			} else {
				var ppi = 96;
				var hpt = f * 1.275;
				var fpx = f * ppi / 72;
				var topt = 72 / ppi;

				var h;
				var a = asc_round(fpx) * topt;
				var d;

				var a_2 = asc_round(fpx / 2) * topt;

				var h_2_3;
				var a_2_3 = asc_round(fpx * 2 / 3) * topt;
				var d_2_3;

				var x = a_2 + a_2_3;

				if (va === AscCommon.vertalign_SuperScript) {
					h = hpt;
					d = h - a;

					l.th = x + d;
					l.bl = x;
					l.bl2 = a_2_3;
					l.a = fm.ascender + a_2;         // >0
					l.d = fm.descender - a_2;        // <0
				} else if (va === AscCommon.vertalign_SubScript) {
					h_2_3 = hpt * 2 / 3;
					d_2_3 = h_2_3 - a_2_3;
					l.th = x + d_2_3;
					l.bl = a;
					l.bl2 = x;
					l.a = fm.ascender + a - x;       // >0
					l.d = fm.descender + x - a;      // >0
				}
			}

			return l;
		};
		StringRender.prototype._calcLineMetrics2 = function (f, va, fm) {
			var l = new lineMetrics();

			var a = Math.max(0, asc.ceil(fm.nat_y1 * f / fm.nat_scale));
			var d = Math.max(0, asc.ceil(-fm.nat_y2 * f / fm.nat_scale)) + 1; // 1 px for border

			/*
			// ToDo
			if (va) {
				var k = (AscCommon.vertalign_SuperScript === va) ? AscCommon.vaKSuper : AscCommon.vaKSub;
				d += asc.ceil((a + d) * k);
				f = asc.ceil(f * 2 / 3 / 0.5) * 0.5; // Round 0.5
				a = Math.max(0, asc.ceil(fm.nat_y1 * f / fm.nat_scale));
			}
			*/

			l.th = a + d;
			l.bl = a;
			l.a = a;
			l.d = d;

			return l;
		};

		StringRender.prototype.calcDelta = function (vnew, vold) {
			return vnew > vold ? vnew - vold : 0;
		};

		/**
		 * @param {Boolean} [dontCalcRepeatChars]
		 * @return {Asc.TextMetrics}
		 */
		StringRender.prototype._calcTextMetrics = function (dontCalcRepeatChars) {
			var self = this, i = 0, p, p_, lm, beg = 0;
			var l = new LineInfo(), TW = 0, TH = 0, BL = 0;

			function addLine(b, e) {
				if (-1 !== b)
					l.tw += self._calcLineWidth(b, e);
				l.beg = b;
				l.end = e < b ? b : e;
				self.lines.push(l);
				if (TW < l.tw) {
					TW = l.tw;
				}
				BL = TH + l.bl;
				TH += l.th;
			}

			if (0 >= this.chars.length) {
				p = this.charProps[0];
				if (p && p.font) {
					lm = this._calcLineMetrics(p.fsz !== undefined ? p.fsz : p.font.getSize(), p.va, p.fm);
					l.assign(lm);
					addLine(-1, -1);
					l.beg = l.end = 0;
				}
			} else {
				for (; i < this.chars.length; ++i) {
					p = this.charProps[i];

					// if font has been changed than calc and update line height and etc.
					if (p && p.font) {
						lm = this._calcLineMetrics(p.fsz !== undefined ? p.fsz : p.font.getSize(), p.va, p.fm);
						if (i === 0) {
							l.assign(lm);
						} else {
							l.th += this.calcDelta(lm.bl, l.bl) + this.calcDelta(lm.th - lm.bl, l.th - l.bl);
							l.bl += this.calcDelta(lm.bl, l.bl);
							l.a += this.calcDelta(lm.a, l.a);
							l.d += this.calcDelta(lm.d, l.d);
						}
						p.lm = lm;
						p_ = p;
					}

					// process 'repeat char' marker
					if (dontCalcRepeatChars && p && p.repeat) {
						l.tw -= this._calcCharsWidth(i, i + p.total);
					}

					// process 'new line' marker
					if (p && (p.nl || p.hp)) {
						addLine(beg, i);
						beg = i + (p.nl ? 1 : 0);
						lm = this._calcLineMetrics(p_.fsz !== undefined ? p_.fsz : p_.font.getSize(), p_.va, p_.fm);
						l = new LineInfo(lm);
					}
				}
				if (beg <= i) {
					// add last line of text
					addLine(beg, i - 1);
				}
			}
			return new asc.TextMetrics(TW, TH, 0, BL, 0, 0);
		};

		StringRender.prototype._getRepeatCharPos = function () {
			var charProp;
			for (var i = 0; i < this.chars.length; ++i) {
				charProp = this.charProps[i];
				if (charProp && charProp.repeat)
					return i;
			}
			return -1;
		};

		/**
		 * @param {Number} maxWidth
		 */
		StringRender.prototype._insertRepeatChars = function (maxWidth) {
			var self = this, width, w, pos, charProp;

			function insertRepeatChars() {
				if (0 === charProp.total)
					return;	// Символ уже изначально лежит в строке и в списке
				var repeatEnd = pos + charProp.total;
				self.chars = [].concat(
					self.chars.slice(0, repeatEnd),
					self.chars.slice(pos, pos + 1),
					self.chars.slice(repeatEnd));

				self.charWidths = [].concat(
					self.charWidths.slice(0, repeatEnd),
					self.charWidths.slice(pos, pos + 1),
					self.charWidths.slice(repeatEnd));
				
				self.charProps = [].concat(
					self.charProps.slice(0, repeatEnd),
					self.charProps.slice(pos, pos + 1),
					self.charProps.slice(repeatEnd));
			}

			function removeRepeatChar() {
				self.chars = [].concat(
					self.chars.slice(0, pos),
					self.chars.slice(pos + 1));

				self.charWidths = [].concat(
					self.charWidths.slice(0, pos),
					self.charWidths.slice(pos + 1));
				
				self.charProps = [].concat(
					self.charProps.slice(0, pos),
					self.charProps.slice(pos + 1));
			}

			width = this._calcTextMetrics(true).width;
			pos = this._getRepeatCharPos();
			if (-1 === pos)
				return;
			w = this._calcCharsWidth(pos, pos);
			charProp = this.charProps[pos];

			while (charProp.total * w + width + w <= maxWidth) {
				insertRepeatChars();
				charProp.total += 1;
				if (w === 0) {
					break;
				}
			}

			if (0 === charProp.total)
				removeRepeatChar();

			this.lines = [];
		};

		StringRender.prototype._getCharPropAt = function (index) {
			var prop = this.charProps[index];
			if (!prop) {
				prop = this.charProps[index] = new charProperties();
			}
			return prop;
		};
		
		StringRender.prototype._getGraphemeDelta = function(grapheme, fontSize) {
			let ppiy = this.drawingCtx.getPPIY();
			let width = AscFonts.GetGraphemeWidth(grapheme) * ppiy / 25.4 * fontSize;
			let bbox = AscFonts.GetGraphemeBBox(grapheme, fontSize, ppiy);
			return bbox.maxX - bbox.minX + 1 - width;
		};

		/**
		 * @param {Number} maxWidth
		 * @return {Asc.TextMetrics}
		 */
		StringRender.prototype._measureChars = function (maxWidth) {
			var self = this;
			var ctx = this.drawingCtx;
			var font = ctx.font;
			var wrap = this.flags && (this.flags.wrapText || this.flags.wrapOnlyCE) && !this.flags.isNumberFormat;
			var wrapNL = this.flags && this.flags.wrapOnlyNL;
			var verticalText = this.flags && this.flags.verticalText;
			var hasRepeats = false;
			var i, j, fr, fmt, chars, p, p_ = {}, pIndex, startCh;
			var tw = 0, nlPos = 0, isEastAsian, hpPos = undefined, isSP_ = true, delta = 0;
			let frShaper = this.fragmentShaper;

			this.drawState.reset(null, null, this.flags, this.angle);
			
			function measureFragment(_chars, format) {
				
				let chPos = self.chars.length;
				let fontSize = format.getSize();
				frShaper.shapeFragment(_chars, format, self, chPos);
				
				var chc, chw, isNL, isSP, isHP;
				for (; chPos < self.chars.length; ++chPos) {
					chc = self.chars[chPos];
					chw = self.charWidths[chPos];
					
					isNL = self.codesHypNL[chc];
					isSP = !isNL ? self.codesHypSp[chc] : false;
					
					// if 'wrap flag' is set
					if (wrap || wrapNL || verticalText) {
						isHP = !isSP && !isNL ? self.codesHyphen[chc] : false;
						isEastAsian = AscCommon.isEastAsianScript(chc);
						if (verticalText) {
							// ToDo verticalText and new line or space
						} else if (isNL) {
							// add new line marker
							nlPos = chPos;
							self._getCharPropAt(nlPos).nl = true;
							self._getCharPropAt(nlPos).delta = delta;
							chc = 0xA0;
							chw = 0;
							tw = 0;
							hpPos = undefined;
							self.charWidths[chPos] = 0;
						} else if (isSP || isHP) {
							// move hyphenation position
							hpPos = chPos + 1;
						} else if (isEastAsian) {
							if (0 !== chPos && !(AscCommon.g_aPunctuation[self.chars[chPos - 1]] &
									AscCommon.PUNCTUATION_FLAG_CANT_BE_AT_END_E) &&
								!(AscCommon.g_aPunctuation[chc] & AscCommon.PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E)) {
								// move hyphenation position
								hpPos = chPos;
							}
						}
						
						if (chPos !== nlPos && ((wrap && !isSP && tw + chw > maxWidth) || (verticalText && !self._isCombinedChar(chPos)))) {
							// add hyphenation marker
							nlPos = hpPos !== undefined ? hpPos : chPos;
							self._getCharPropAt(nlPos).hp = true;
							self._getCharPropAt(nlPos).delta = delta;
							tw = self._calcCharsWidth(nlPos, chPos - 1);
							hpPos = undefined;
						}
						
						if (isEastAsian) {
							// move hyphenation position
							if (chPos < self.chars.length - 1 && !(AscCommon.g_aPunctuation[self.chars[chPos + 1]] &
									AscCommon.PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E) &&
								!(AscCommon.g_aPunctuation[chc] & AscCommon.PUNCTUATION_FLAG_CANT_BE_AT_END_E)) {
								hpPos = chPos + 1;
							}
						}
					}
					
					if (isSP_ && !isSP && !isNL) {
						// add word beginning marker
						self._getCharPropAt(chPos).wrd = true;
					}
					
					tw += chw;
					
					isSP_ = isSP || isNL;
					
					if (isSP || isNL) {
						delta = 0;
					} else if (AscFonts.NO_GRAPHEME !== self._getCharPropAt(chPos).grapheme) {
						delta = self._getGraphemeDelta(self._getCharPropAt(chPos).grapheme, fontSize);
					}
				}
			}
			
			this._reset();
			
			// for each text fragment
			for (i = 0; i < this.fragments.length; ++i) {
				startCh = this.charWidths.length;
				fr = this.fragments[i];
				fmt = fr.format.clone();
				var va = fmt.getVerticalAlign();

				//TODO пока не убрал эту регулярку, сначала перевожу в текст, потом обратно в сиволы
				//TODO избавиться от регулярки!
				if (fr.isInitCharCodes()) {
					fr.initText();
				}
				chars = this._filterChars(fr.getCharCodes(), wrap || wrapNL);
				//fr.initCharCodes();

				pIndex = this.chars.length;
				p = this.charProps[pIndex];
				p = p ? p.clone() : new charProperties();

				// reduce font size for subscript and superscript chars
				if (va === AscCommon.vertalign_SuperScript || va === AscCommon.vertalign_SubScript) {
					p.va = va;
					p.fsz = fmt.getSize();
					fmt.fs = p.fsz * 2 / 3;
					p.font = fmt;
				}
				
				// change font on canvas
				if (!fmt.isEqual(ctx.font)
					|| fmt.getUnderline() !== font.getUnderline()
					|| fmt.getStrikeout() !== font.getStrikeout()
					|| fmt.getColor() !== p_.c) {
					p.font = fmt;
				}
				this._setFont(ctx, fmt);
				
				// add marker in chars flow
				if (i === 0) {
					p.font = fmt;
				}
				if (p.font) {
					p.fm = ctx.getFontMetrics();
					p.c = fmt.getColor();
					this.charProps[pIndex] = p;
					p_ = p;
				}

				if (fmt.getSkip()) {
					this._getCharPropAt(pIndex).skip = chars.length;
				}

				if (fmt.getRepeat()) {
					if (hasRepeats)
						throw new Error("Repeat should occur no more than once");

					this._getCharPropAt(pIndex).repeat = true;
					this._getCharPropAt(pIndex).total = 0;
					hasRepeats = true;
				}

				if (chars.length < 1) {
					continue;
				}
				measureFragment(chars, fmt);

				// для italic текста прибавляем к концу строки разницу между charWidth и BBox
				for (j = startCh; font.getItalic() && j < this.charWidths.length; ++j) {
					if (this.charProps[j] && this.charProps[j].delta && j > 0) {
						if (this.charWidths[j - 1] > 0) {
							this.charWidths[j - 1] += this.charProps[j].delta;
						} else if (j > 1) {
							this.charWidths[j - 2] += this.charProps[j].delta;
						}
					}
				}
			}

			if (0 !== this.chars.length && this.charProps[this.chars.length] !== undefined) {
				delete this.charProps[this.chars.length];
			} else if (font.getItalic()) {
				// для italic текста прибавляем к концу текста разницу между charWidth и BBox
				this.charWidths[this.charWidths.length - 1] += delta;
			}

			if (hasRepeats) {
				if (maxWidth === undefined) {
					throw new Error("Undefined width of cell width Numeric Format");
				}
				this._insertRepeatChars(maxWidth);
			}

			return this._calcTextMetrics();
		};

		/**
		 * @param {Number} maxWidth
		 * @return {Asc.TextMetrics}
		 */
		StringRender.prototype._doMeasure = function (maxWidth) {
			var ratio, format, size, canReduce = true, minSize = 2.5;
			var tm = this._measureChars(maxWidth);
			while (this.flags && this.flags.shrinkToFit && tm.width > maxWidth && canReduce) {
				canReduce = false;
				ratio = maxWidth / tm.width;
				for (var i = 0; i < this.fragments.length; ++i) {
					format = this.fragments[i].format;
					size = Math.max(minSize, Math.floor(format.getSize() * ratio * 2) / 2);
					format.setSize(size);
					if (minSize < size) {
						canReduce = true;
					}
				}
				tm = this._measureChars(maxWidth);
			}
			return tm;
		};
		
		/**
		 * @param {DrawingContext} drawingCtx
		 * @param {Number} x
		 * @param {Number} y
		 * @param {Number} maxWidth
		 * @param {String} textColor
		 */

		StringRender.prototype._doRender = function (drawingCtx, x, y, maxWidth, textColor) {
			let self = this;
			let ctx = drawingCtx || this.drawingCtx;
			let zoom = ctx.getZoom();
			let ppiy = ctx.getPPIY();
			this.drawState.reset(drawingCtx, textColor, this.flags, this.angle);
			let drawState = this.drawState;
			let align = this.getEffectiveAlign();
			let i, j, p, p_, strBeg;
			let n = 0, l = this.lines[0], x1 = l ? this.initStartX(0, l, x, maxWidth) : 0, y1 = y, dx = l ? computeWordDeltaX() : 0;

			ctx.setTextRotated(!!this.angle);
			self.textColor = textColor;


			function computeWordDeltaX() {
				if (align !== AscCommon.align_Justify || n === self.lines.length - 1) {
					return 0;
				}

				if (align === AscCommon.align_Justify) {
					let wordCount = 0;
					let isLastWordSpace = false;
					let lastSpacesWidth = 0;
					let lastSymbolWidth = 0;

					for (let i = l.beg; i <= l.end; ++i) {
						let p = self.charProps[i];
						let isSpace = self.codesHypSp[self.chars[i]];

						if (p && p.wrd && isLastWordSpace) {
							++wordCount;
							if (i !== l.end) {
								lastSpacesWidth = 0;
							} else if (!isSpace) {
								lastSymbolWidth = self.charWidths[i];
							}
						} else if (i === l.end) {
							++wordCount;
						}

						if (isSpace) {
							lastSpacesWidth += self.charWidths[i];
						}

						isLastWordSpace = isSpace;
					}

					if (wordCount <= 1) {
						return 0;
					}

					let rightDiff = 1;
					let availableWidth = maxWidth - rightDiff - (l.tw - lastSymbolWidth - lastSpacesWidth);
					return (availableWidth) / (wordCount - 1);
				} else {
					for (var i = l.beg, c = 0; i <= l.end; ++i) {
						var p = self.charProps[i];
						if (p && p.wrd) {
							++c;
						}
					}
					return c > 1 ? (maxWidth - l.tw) / (c - 1) : 0;
				}
			}

			function renderFragment(begin, end, prop, angle) {
				var dh = prop && prop.lm && prop.lm.bl2 > 0 ? prop.lm.bl2 - prop.lm.bl : 0;
				var dw = self._calcCharsWidth(strBeg, end - 1);
				var so = prop.font.getStrikeout();
				var ul = Asc.EUnderline.underlineNone !== prop.font.getUnderline();
				var isSO = so === true;
				var fsz, x2, y, lw, dy, i, b, cp;
				var bl = asc_round(l.bl * zoom);

				if (begin > end)
					return 0;

				let fontSize = prop.font.getSize();
				y = y1 + bl + dh;

				let startX = drawState.x;
				x1 = startX;
				if (align !== AscCommon.align_Justify || dx < 0.000001) {
					renderGraphemes(begin, end, drawState.x, y, fontSize);
				} else {
					for (i = b = begin; i < end; ++i) {
						cp = self.charProps[i];
						if (cp && cp.wrd && i > b) {
							x1 = drawState.x;
							renderGraphemes(b, i, drawState.x, y, fontSize);
							x1 += self._calcCharsWidth(b, i - 1) + dx;
							drawState.x = x1;
							dw += dx;
							b = i;
						}
					}
					if (i > b) {
						renderGraphemes(b, i, drawState.x, y, fontSize);
					}
				}


				if (isSO || ul) {
					if (angle && window["IS_NATIVE_EDITOR"])
						ctx.nativeTextDecorationTransform(true);

					x2 = startX + dw;
					fsz = prop.font.getSize();
					lw = asc_round(fsz * ppiy / 72 / 18) || 1;
					ctx.setStrokeStyle(prop.c || textColor)
						.setLineWidth(lw)
						.beginPath();
					dy = (lw / 2);
					dy = dy >> 0;
					if (ul) {
						y = asc_round(y1 + bl + prop.lm.d * 0.4 * zoom);
						ctx.lineHor(startX, y + dy, x2 + 1);
					}
					if (isSO) {
						dy += 1;
						y = asc_round(y1 + bl - prop.lm.a * 0.275 * zoom);
						ctx.lineHor(startX, y - dy, x2 + 1);
					}
					ctx.stroke();

					if (angle && window["IS_NATIVE_EDITOR"])
						ctx.nativeTextDecorationTransform(false);
				}

				return dw;
			}

			function renderGraphemes(begin, end, x, y) {
				drawState.y = y;
				drawState.beginFragment(begin, end, p_);
			}


			drawState.beginLine(l, x1, y);
			for (i = 0, strBeg = 0; i < this.chars.length; ++i) {
				p = this.charProps[i];

				if (p && (p.font || p.nl || p.hp || p.skip > 0)) {
					if (strBeg < i) {
						renderFragment(strBeg, i, p_, this.angle);
						strBeg = i;
					}
					if (p.nl) {
						strBeg += 1;
					}

					if (p.font) {
						p_ = p;
					}
					if (p.skip > 0) {
						j = i + p.skip - 1;
						drawState.x += this._calcCharsWidth(i, j);
						strBeg = j + 1;
						i = j;
						continue;
					}
					if (p.nl || p.hp) {
						drawState.endLine();
						y1 += asc_round(l.th * zoom);
						l = self.lines[++n];
						drawState.x = self.initStartX(i, l, x, maxWidth);
						dx = computeWordDeltaX();
						drawState.beginLine(l, drawState.x, y);
					}
				}
			}
			if (strBeg < i) {
				renderFragment(strBeg, i, p_, this.angle);
			}

			drawState.endLine();
		};
		StringRender.prototype.initStartX = function (startPos, l, x, maxWidth, initAllLines) {
			let align = this.getEffectiveAlign();

			if (initAllLines) {
				if (this.lines) {
					for (let i = 0; i < this.lines.length; ++i) {
						let lineWidth = this._calcLineWidth(this.lines[i].beg);
						this.lines[i].initStartX(lineWidth, x, maxWidth, align);
					}
				}
			} else {
				return l.initStartX(this._calcLineWidth(startPos), x, maxWidth, align);
			}
		};
		StringRender.prototype.getInternalState = function () {
			return {
				/** @type Object */
				flags: this.flags,

				chars: this.chars,
				charWidths: this.charWidths,
				charProps: this.charProps,
				lines: this.lines
			};
		};
		StringRender.prototype.restoreInternalState = function (state) {
			this.flags = state.flags;
			this.chars = state.chars;
			this.charWidths = state.charWidths;
			this.charProps = state.charProps;
			this.lines = state.lines;
			return this;
		};
		StringRender.prototype._setFont = function (ctx, font) {
			let oldColor = font.c;
			if(this.textColor) font.c = this.textColor;
			ctx.setFont(font, this.angle);
			font.c = oldColor;
		};
		StringRender.prototype._isCombinedChar = function(pos) {
			let p = this._getCharPropAt(pos);
			let c = this.chars[pos];
			return !p.nl && !this.codesSpace[c] && (AscFonts.NO_GRAPHEME === p.grapheme);
		};
		StringRender.prototype.getEffectiveAlign = function() {
			let align = this.flags ? this.flags.textAlign : null;
			let isRtl = this.drawState.getMainDirection() === AscBidi.TYPE.R;

			if (!isRtl) {
				return align;
			}

			if (align === AscCommon.align_Left) {
				return AscCommon.align_Right;
			} else if (align === AscCommon.align_Right) {
				return AscCommon.align_Left;
			}
			else if(align === null) {
				return AscCommon.align_Right;
			}

			return align;
		};
		//------------------------------------------------------------export---------------------------------------------------
		window['AscCommonExcel'] = window['AscCommonExcel'] || {};
		window["AscCommonExcel"].StringRender = StringRender;


		function TableCellDrawState(stringRender) {
			this.stringRender = stringRender;
			this.bidiFlow = new AscWord.BidiFlow(this);
			this.drawingCtx = this.stringRender.drawingCtx;
			this.x = 0;
			this.y = 0;
			this.baseY = 0;
			this.zoom = 1;
			this.ppiy = 96;
			this.currentFont = null;
			this.currentColor = null;
			this.textColor = null;
			this.angle = 0;
			this.currentLine = null;
			this.startIdx = 0;
		}


		TableCellDrawState.prototype.endLine = function() {
			this.bidiFlow.end();
		};
		TableCellDrawState.prototype.getBidiType = function(char, charProp) {
			if (charProp && charProp.nl) {
				return AscBidi.TYPE.B;
			}
			if (this.stringRender.codesHypSp[char]) {
				return AscBidi.TYPE.WS;
			}
			return  AscBidi.getType(char);
		};
		TableCellDrawState.prototype.beginFragment = function(begin, end, prop) {
			let i = begin;
			while (i < end) {
				let charProp = this.stringRender.charProps[i];
				if (charProp && charProp.skip) {
					i++;
					continue;
				}
				let char = this.stringRender.chars[i];
				let bidiType = this.getBidiType(char, charProp);
				this.stringRender._setFont(this.drawingCtx, prop.font);

				//todo: implement the stack of states in DrawingContext and remove this check
				let textColor = prop.c || this.stringRender.textColor;
				let _r = textColor.getR();
				let _g = textColor.getG();
				let _b = textColor.getB();
				let _a = textColor.getA();
				let setColor = true;
				if (this.drawingCtx.fillColor && this.drawingCtx.fillColor.isEqual(_r, _g, _b, _a)) {
					setColor = false;
				}
				if (setColor) {
					this.drawingCtx.setFillStyle(textColor);
				}
				/////
				this.bidiFlow.add({
					charIndex: i,
					charProp: charProp,
					fragmentProp: prop
				}, bidiType);
				i++;
			}
		};

		TableCellDrawState.prototype.handleBidiFlow = function(data, direction) {
			let charIndex = data.charIndex;
			let charProp = data.charProp;
			let char = this.stringRender.chars[charIndex];
			let width = this.stringRender.charWidths[charIndex];
			let grapheme = charProp ? charProp.grapheme : AscFonts.NO_GRAPHEME;

			if (direction === AscBidi.DIRECTION.R && AscBidi.isPairedBracket(char)) {
				if (grapheme !== AscFonts.NO_GRAPHEME) {
					grapheme = AscBidi.getPairedBracketGrapheme(grapheme);
				}
			}

			let fontSize = data.fragmentProp && data.fragmentProp.font ? data.fragmentProp.font.getSize() : 10;
			let y = this.y;

			if (grapheme !== AscFonts.NO_GRAPHEME) {
				AscFonts.DrawGrapheme(grapheme, this.drawingCtx, this.x, y, fontSize, this.ppiy / 25.4);
			}

			this.x += width;
		};

		TableCellDrawState.prototype.beginLine = function(line, x, y) {
			this.currentLine = line;
			this.x = x;
			this.y = y;
			this.baseY = y;

			this.bidiFlow.begin(this.getMainDirection() === AscBidi.TYPE.R);
		};



		TableCellDrawState.prototype.reset = function(drawingCtx, textColor, flags, angle) {
			this.drawingCtx = drawingCtx || this.stringRender.drawingCtx;
			this.x = 0;
			this.y = 0;
			this.baseY = 0;
			this.currentFont = null;
			this.currentColor = null;
			this.currentLine = null;
			this.startIdx = 0;
			this.textColor = textColor || null;
			this.angle = angle || 0;
			this.zoom = this.drawingCtx.getZoom();
			this.ppiy = this.drawingCtx.getPPIY();
		};
		TableCellDrawState.prototype.getMainDirection = function() {
			let readingOrder = this.stringRender.flags ? this.stringRender.flags.getReadingOrder() : null;
			if (readingOrder === 1) {
				return AscBidi.TYPE.L;
			} else if (readingOrder === 2) {
				return AscBidi.TYPE.R;
			}
			for (let i = 0; i < this.stringRender.chars.length; ++i) {
				let char = this.stringRender.chars[i];
				let type = AscBidi.getType(char);
				if (type & AscBidi.FLAG.STRONG) {
					if (type & AscBidi.FLAG.RTL) {
						return AscBidi.TYPE.R;
					} else {
						return AscBidi.TYPE.L;
					}
				}
			}

			return AscBidi.TYPE.L;
		};
	}


)(window);
