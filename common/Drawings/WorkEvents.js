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

(function(window, undefined)
{
	AscCommon.isTouch 			= false;
	AscCommon.isTouchMove 		= false;
	AscCommon.TouchStartTime 	= -1;
	// Import
	var AscBrowser               = AscCommon.AscBrowser;

	// константы для мыши
	var g_mouse_event_type_down  = 0;
	var g_mouse_event_type_move  = 1;
	var g_mouse_event_type_up    = 2;
	var g_mouse_event_type_wheel = 3;

	var g_mouse_button_left   = 0;
	var g_mouse_button_center = 1;
	var g_mouse_button_right  = 2;

	var MouseUpLock = {
		MouseUpLockedSend : false
	};

	AscCommon.stopEvent = function(e)
	{
		if (!e)
			return;
		if (e.preventDefault)
			e.preventDefault();
		if (e.stopPropagation)
			e.stopPropagation();
	};

	var isUsePointerEvents = true;
	if (AscBrowser.isChrome && (AscBrowser.chromeVersion <= 70)) // xp
		isUsePointerEvents = false;
	else if (AscBrowser.isSafari && (AscBrowser.safariVersion < 15000000))
		isUsePointerEvents = false;
	else if (AscBrowser.isIE)
		isUsePointerEvents = false;

	AscCommon.getPtrEvtName = function (sType)
	{
		return (isUsePointerEvents ? "pointer" : "mouse") + sType;
	};
	AscCommon.getPtrEvtType = function (sType)
	{
		return "on" + AscCommon.getPtrEvtName(sType);
	};

	AscCommon.addMouseEvent = function(elem, type, handler)
	{
		var _type = AscCommon.getPtrEvtType(type);
		elem[_type] = handler;
	};
    AscCommon.removeMouseEvent = function(elem, type)
    {
        var _type = AscCommon.getPtrEvtType(type);
        if (elem[_type])
        	delete elem[_type];
    };
	AscCommon.getMouseEvent = function(elem, type)
	{
		var _type = AscCommon.getPtrEvtType(type);
		return elem[_type];
	};
	AscCommon.capturePointer = function(e, elem)
	{
		if (e.pointerType === "mouse" && elem.setPointerCapture)
		{
			try {
				elem.setPointerCapture(e.pointerId);
			} catch (e) {}
		}
	};

	function CMouseEventHandler()
	{
		this.X = 0;                            // позиция курсора X
		this.Y = 0;                            // позиция курсора Y

		this.Button = g_mouse_button_left;          // кнопка мыши
		this.ButtonOverride = -1;
		this.Type   = g_mouse_event_type_move;      // тип евента

		this.AltKey   = false;                        // нажата ли кнопка alt
		this.CtrlKey  = false;                        // нажата ли кнопка ctrl
		this.ShiftKey = false;                        // нажата ли кнопка shift

		this.Sender = null;                         // от какого html элемента пришел евент

		this.LastClickTime = -1;                       // время последнего mousedown
		this.ClickCount    = 0;                        // количество кликов

		this.WheelDelta = 0;

		// координаты мышки при предыдущем mousedown (для mousemove)
		this.IsPressed = false;                        // была ли зажата кнопка
		this.LastX     = 0;
		this.LastY     = 0;

		this.KoefPixToMM = 1;

		this.IsLocked      = false;
		this.IsLockedEvent = false;

		this.buttonObject = null;

		this.AscHitToHandlesEpsilon = 0;

		this.LockMouse   = function()
		{
			if (!this.IsLocked)
			{
				this.IsLocked = true;

				if (window.captureEvents)
					window.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP);

				if (window.g_asc_plugins)
					window.g_asc_plugins.disablePointerEvents();

				/*
				 var parent = window;
				 while (true)
				 {
				 if (!parent)
				 break;

				 if (parent.captureEvents)
				 parent.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP);

				 if (parent == parent.parent)
				 break;

				 parent = parent.parent;
				 }
				 */

				return true;
			}
			return false;
		};
		this.UnLockMouse = function()
		{
			if (this.IsLocked)
			{
				this.IsLocked = false;

				if (window.releaseEvents)
					window.releaseEvents(Event.MOUSEMOVE);

				if (window.g_asc_plugins)
					window.g_asc_plugins.enablePointerEvents();

				/*
				 var parent = window;
				 while (true)
				 {
				 if (!parent)
				 break;

				 if (parent.releaseEvents)
				 parent.releaseEvents(Event.MOUSEMOVE);

				 if (parent == parent.parent)
				 break;

				 parent = parent.parent;
				 }
				 */

				return true;
			}
			return false;
		};
	}

	function CKeyboardEvent()
	{
		this.AltKey   = false;                        // нажата ли кнопка alt
		this.CtrlKey  = false;                        // нажата ли кнопка ctrl
		this.ShiftKey = false;                        // нажата ли кнопка shift
		this.MacCmdKey = false;
		this.AltGr    = false;

		this.Sender = null;                         // от какого html элемента пришел евент

		this.CharCode = 0;
		this.KeyCode  = 0;
	}

	CKeyboardEvent.prototype.Up = function()
	{
		this.AltKey    = false;
		this.CtrlKey   = false;
		this.ShiftKey  = false;
		this.AltGr     = false;
        this.MacCmdKey = false;
	};
	CKeyboardEvent.prototype.IsCtrl = function()
	{
		return (this.CtrlKey || (this.AltKey && this.AltGr));
	};
	CKeyboardEvent.prototype.IsShortcutCtrl = function()
	{
		return this.IsCtrl();
	};
	CKeyboardEvent.prototype.IsShift = function()
	{
		return this.ShiftKey;
	};
	CKeyboardEvent.prototype.IsAlt = function()
	{
		return this.AltKey;
	};
	CKeyboardEvent.prototype.GetKeyCode = function()
	{
		return this.KeyCode;
	};
	CKeyboardEvent.prototype.IsMacCmd = function() {
		return this.MacCmdKey;
	};



	var global_mouseEvent    = new CMouseEventHandler();
	var global_keyboardEvent = new CKeyboardEvent();

	function check_KeyboardEvent(e)
	{
		global_keyboardEvent.AltKey = e.altKey;
		global_keyboardEvent.AltGr = AscCommon.getAltGr(e);
		global_keyboardEvent.CtrlKey = !global_keyboardEvent.AltGr && (e.metaKey || e.ctrlKey);
        global_keyboardEvent.MacCmdKey = AscCommon.AscBrowser.isMacOs && e.metaKey;

		global_keyboardEvent.ShiftKey = e.shiftKey;

		global_keyboardEvent.Sender = (e.srcElement) ? e.srcElement : e.target;

		global_keyboardEvent.CharCode = e.charCode;
		global_keyboardEvent.KeyCode  = e.keyCode;
		global_keyboardEvent.Which    = e.which;
	}

	function check_KeyboardEvent2(e)
	{
		global_keyboardEvent.AltKey = e.altKey;

		if (e.metaKey !== undefined)
			global_keyboardEvent.CtrlKey = e.ctrlKey || e.metaKey;
		else
			global_keyboardEvent.CtrlKey = e.ctrlKey;

        global_keyboardEvent.MacCmdKey = AscCommon.AscBrowser.isMacOs && e.metaKey;

		global_keyboardEvent.ShiftKey = e.shiftKey;

		global_keyboardEvent.AltGr = (global_keyboardEvent.CtrlKey && global_keyboardEvent.AltKey) ? true : false;

		if (global_keyboardEvent.CtrlKey && global_keyboardEvent.AltKey)
			global_keyboardEvent.CtrlKey = false;
	}

	function check_MouseMoveEvent(e)
	{
		// если мышь залочена, то евент придет от окна.
		if (e.IsLocked && !e.IsLockedEvent)
			return;

		if (e.pageX || e.pageY)
		{
			global_mouseEvent.X = e.pageX;
			global_mouseEvent.Y = e.pageY;
		}
		else if (e.clientX || e.clientY)
		{
			global_mouseEvent.X = e.clientX;
			global_mouseEvent.Y = e.clientY;
		}

		global_mouseEvent.X = (global_mouseEvent.X * AscBrowser.zoom) >> 0;
		global_mouseEvent.Y = (global_mouseEvent.Y * AscBrowser.zoom) >> 0;

		global_mouseEvent.AltKey   = e.altKey;
		global_mouseEvent.ShiftKey = e.shiftKey;
		global_mouseEvent.CtrlKey  = e.ctrlKey || e.metaKey;

		global_mouseEvent.Type = g_mouse_event_type_move;

		if (!global_mouseEvent.IsLocked)
		{
			global_mouseEvent.Sender = (e.srcElement) ? e.srcElement : e.target;
		}

		var _eps = 3 * global_mouseEvent.KoefPixToMM;
		if ((Math.abs(global_mouseEvent.X - global_mouseEvent.LastX) > _eps) || (Math.abs(global_mouseEvent.Y - global_mouseEvent.LastY) > _eps))
		{
			global_mouseEvent.LastClickTime = -1;
			global_mouseEvent.ClickCount    = 0;
		}
	}

	function CreateMouseUpEventObject(x, y)
	{
		var e   = {};
		e.PageX = x;
		e.PageY = y;

		e.altKey   = global_mouseEvent.AltKey;
		e.shiftKey = global_mouseEvent.ShiftKey;
		e.ctrlKey  = global_mouseEvent.CtrlKey;

		e.srcElement = global_mouseEvent.Sender;
		e.button     = 0;
		return e;
	}

	function getMouseButton(e)
	{
		if (-1 !== global_mouseEvent.ButtonOverride)
			return global_mouseEvent.ButtonOverride;
		var res = e.button;
		return (res && -1 !== res) ? res : 0;
	}

	function check_MouseUpEvent(e)
	{
		if (e.pageX || e.pageY)
		{
			global_mouseEvent.X = e.pageX;
			global_mouseEvent.Y = e.pageY;
		}
		else if (e.clientX || e.clientY)
		{
			global_mouseEvent.X = e.clientX;
			global_mouseEvent.Y = e.clientY;
		}

		global_mouseEvent.X = (global_mouseEvent.X * AscBrowser.zoom) >> 0;
		global_mouseEvent.Y = (global_mouseEvent.Y * AscBrowser.zoom) >> 0;

		global_mouseEvent.AltKey   = e.altKey;
		global_mouseEvent.ShiftKey = e.shiftKey;
		global_mouseEvent.CtrlKey  = e.ctrlKey || e.metaKey;

		global_keyboardEvent.AltKey   = global_mouseEvent.AltKey;
		global_keyboardEvent.ShiftKey = global_mouseEvent.ShiftKey;
		global_keyboardEvent.CtrlKey  = global_mouseEvent.CtrlKey;

		global_mouseEvent.Type   = g_mouse_event_type_up;
		global_mouseEvent.Button = getMouseButton(e);

		var lockedElement = null;

		var newSender = (e.srcElement) ? e.srcElement : e.target;
		if (!newSender)
		    newSender = { id : "emulation_oo_id" };

		if (global_mouseEvent.Sender && global_mouseEvent.Sender.id == newSender.id)
		{
			lockedElement = global_mouseEvent.Sender;
		}

		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != newSender && false === MouseUpLock.MouseUpLockedSend)
		{
			Window_OnMouseUp(e);
		}
		MouseUpLock.MouseUpLockedSend = true;
		global_mouseEvent.Sender      = newSender;

		global_mouseEvent.UnLockMouse();

		global_mouseEvent.IsPressed = false;

		return lockedElement;
	}

	function check_MouseClickOnUp()
	{
		// call after check mouseUp
        if (0 == global_mouseEvent.ClickCount)
        	return false;

        var _eps = 3 * global_mouseEvent.KoefPixToMM;
        if ((Math.abs(global_mouseEvent.X - global_mouseEvent.LastX) > _eps) || (Math.abs(global_mouseEvent.Y - global_mouseEvent.LastY) > _eps))
            return false;

        var CurTime = new Date().getTime();
        if (500 < (CurTime - global_mouseEvent.LastClickTime))
        	return false;

        return true;
	}

	function check_MouseDownEvent(e, isClicks)
	{
		if (e.pageX || e.pageY)
		{
			global_mouseEvent.X = e.pageX;
			global_mouseEvent.Y = e.pageY;
		}
		else if (e.clientX || e.clientY)
		{
			global_mouseEvent.X = e.clientX;
			global_mouseEvent.Y = e.clientY;
		}

		global_mouseEvent.X = (global_mouseEvent.X * AscBrowser.zoom) >> 0;
		global_mouseEvent.Y = (global_mouseEvent.Y * AscBrowser.zoom) >> 0;

		var _eps = 3 * global_mouseEvent.KoefPixToMM;
		if ((Math.abs(global_mouseEvent.X - global_mouseEvent.LastX) > _eps) || (Math.abs(global_mouseEvent.Y - global_mouseEvent.LastY) > _eps))
		{
			// not only move!!! (touch - fast click in different places)
			global_mouseEvent.LastClickTime = -1;
			global_mouseEvent.ClickCount    = 0;
		}

		global_mouseEvent.LastX = global_mouseEvent.X;
		global_mouseEvent.LastY = global_mouseEvent.Y;

		global_mouseEvent.AltKey   = e.altKey;
		global_mouseEvent.ShiftKey = e.shiftKey;
		global_mouseEvent.CtrlKey  = e.ctrlKey || e.metaKey;

		global_keyboardEvent.AltKey   = global_mouseEvent.AltKey;
		global_keyboardEvent.ShiftKey = global_mouseEvent.ShiftKey;
		global_keyboardEvent.CtrlKey  = global_mouseEvent.CtrlKey;

		global_mouseEvent.Type   = g_mouse_event_type_down;
		let oldButton = global_mouseEvent.Button;
		global_mouseEvent.Button = getMouseButton(e);

		if (!global_mouseEvent.IsLocked || !global_mouseEvent.Sender)
			global_mouseEvent.Sender = (e.srcElement) ? e.srcElement : e.target;

		AscCommon.capturePointer(e, global_mouseEvent.Sender);

		if (isClicks)
		{
			var CurTime = new Date().getTime();
			if (0 === global_mouseEvent.ClickCount)
			{
				global_mouseEvent.ClickCount    = 1;
				global_mouseEvent.LastClickTime = CurTime;
			}
			else
			{
				if ((500 > (CurTime - global_mouseEvent.LastClickTime)) && oldButton === global_mouseEvent.Button)
				{
					global_mouseEvent.LastClickTime = CurTime;
					global_mouseEvent.ClickCount++;
				}
				else
				{
					global_mouseEvent.ClickCount    = 1;
					global_mouseEvent.LastClickTime = CurTime;
				}
			}
		}
		else
		{
			global_mouseEvent.LastClickTime = -1;
			global_mouseEvent.ClickCount    = 1;
		}

		MouseUpLock.MouseUpLockedSend = false;
	}

	function check_MouseDownEvent2(x, y)
	{
		global_mouseEvent.X = x;
		global_mouseEvent.Y = y;

		global_mouseEvent.LastX = global_mouseEvent.X;
		global_mouseEvent.LastY = global_mouseEvent.Y;

		global_mouseEvent.Type = g_mouse_event_type_down;

		global_mouseEvent.Sender = editor.WordControl.m_oEditor.HtmlElement;

		global_mouseEvent.LastClickTime = -1;
		global_mouseEvent.ClickCount    = 1;

		MouseUpLock.MouseUpLockedSend = false;
	}

	function global_OnMouseWheel(e)
	{
		global_mouseEvent.AltKey   = e.altKey;
		global_mouseEvent.ShiftKey = e.shiftKey;
		global_mouseEvent.CtrlKey  = e.ctrlKey || e.metaKey;

		if (undefined != e.wheelDelta)
			global_mouseEvent.WheelDelta = (e.wheelDelta > 0) ? -45 : 45;
		else
			global_mouseEvent.WheelDelta = (e.detail > 0) ? 45 : -45;

		global_mouseEvent.type = g_mouse_event_type_wheel;

		global_mouseEvent.Sender = (e.srcElement) ? e.srcElement : e.target;

		global_mouseEvent.LastClickTime = -1;
		global_mouseEvent.ClickCount    = 0;
	}

	function InitCaptureEvents()
	{
		AscCommon.addMouseEvent(window, "move", Window_OnMouseMove);
        AscCommon.addMouseEvent(window, "up", Window_OnMouseUp);
	}

	function Window_OnMouseMove(e)
	{
		if (!global_mouseEvent.IsLocked || !global_mouseEvent.Sender)
			return;

        var types = isUsePointerEvents ? ["onpointermove", "onmousemove"] : ["onmousemove", "onpointermove"];
        for (var i = 0; i < 2; i++)
		{
            if (global_mouseEvent.Sender[types[i]])
            {
                global_mouseEvent.IsLockedEvent = true;
                global_mouseEvent.Sender[types[i]](e);
                global_mouseEvent.IsLockedEvent = false;
                break;
            }
		}
	}

	function Window_OnMouseUp(e)
	{
		if (false === MouseUpLock.MouseUpLockedSend)
		{
			MouseUpLock.MouseUpLockedSend = true;
			if (global_mouseEvent.IsLocked && global_mouseEvent.Sender)
			{
                var types = isUsePointerEvents ? ["onpointerup", "onmouseup"] : ["onmouseup", "onpointerup"];
                for (var i = 0; i < 2; i++)
                {
                    if (global_mouseEvent.Sender[types[i]])
                    {
                        global_mouseEvent.Sender[types[i]](e, true);
                        if (global_mouseEvent.IsLocked)
                        	global_mouseEvent.UnLockMouse();
                        break;
                    }
                }
			}
		}

		if (window.g_asc_plugins)
		    window.g_asc_plugins.onExternalMouseUp();
	}

	InitCaptureEvents();

	function button_eventHandlers(disable_pos, norm_pos, over_pos, down_pos, control, click_func_delegate)
	{
		this.state_normal = norm_pos;
		this.state_over   = over_pos;
		this.state_down   = down_pos;

		this.Click_func = click_func_delegate;
		this.Control    = control;
		this.IsPressed  = false;

		var oThis = this;

		this.Control.HtmlElement.onmouseover = function(e)
		{
			check_MouseMoveEvent(e);

			if (global_mouseEvent.IsLocked)
			{
				if (global_mouseEvent.Sender.id != oThis.Control.HtmlElement.id)
				{
					// это не залоченная кнопка
					return;
				}
				// залоченная кнопка
				oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_down;
				return;
			}
			oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_over;
		}
		this.Control.HtmlElement.onmouseout  = function(e)
		{
			check_MouseMoveEvent(e);

			if (global_mouseEvent.IsLocked)
			{
				if (global_mouseEvent.Sender.id != oThis.Control.HtmlElement.id)
				{
					// это не залоченная кнопка
					return;
				}
				// залоченная кнопка
				oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_over;
				return;
			}
			oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_normal;
		}
		this.Control.HtmlElement.onmousedown = function(e)
		{
			check_MouseDownEvent(e);
			global_mouseEvent.LockMouse();
			global_mouseEvent.buttonObject = oThis;

			AscCommon.stopEvent(e);

			if (global_mouseEvent.IsLocked)
			{
				if (global_mouseEvent.Sender.id != oThis.Control.HtmlElement.id)
				{
					// это не залоченная кнопка
					return;
				}
				// залоченная кнопка
				oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_down;
				return;
			}
			oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_down;
		}
		this.Control.HtmlElement.onmouseup   = function(e)
		{
			var lockedElement = check_MouseUpEvent(e);

			if (e.preventDefault)
				e.preventDefault();
			else
				e.returnValue = false;

			if (null != lockedElement && global_mouseEvent.buttonObject != null)
			{
				oThis.Click_func();
			}

			if (null != lockedElement)
			{
				oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_over;
			}
			else
			{
				if (null != global_mouseEvent.buttonObject)
					global_mouseEvent.buttonObject.Control.HtmlElement.style.backgroundPosition = global_mouseEvent.buttonObject.state_normal;

				if ((global_mouseEvent.buttonObject == null) || (oThis.Control.HtmlElement.id != global_mouseEvent.buttonObject.Control.HtmlElement.id))
					oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_over;
			}
			global_mouseEvent.buttonObject = null;
		}

		// теперь touch
		this.Control.HtmlElement.ontouchstart = function(e)
		{
			oThis.Control.HtmlElement.onmousedown(e.touches[0]);
			return false;
		}
		this.Control.HtmlElement.ontouchend   = function(e)
		{
			var lockedElement = check_MouseUpEvent(e.changedTouches[0]);

			if (null != lockedElement)
			{
				oThis.Click_func();
				oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_normal;
			}
			else
			{
				if (null != global_mouseEvent.buttonObject)
					global_mouseEvent.buttonObject.Control.HtmlElement.style.backgroundPosition = global_mouseEvent.buttonObject.state_normal;

				if (oThis.Control.HtmlElement.id != global_mouseEvent.buttonObject.Control.HtmlElement.id)
					oThis.Control.HtmlElement.style.backgroundPosition = oThis.state_normal;
			}
			global_mouseEvent.buttonObject = null;
			return false;
		}
	}

	function emulateKeyDown(_code, _element)
	{
		var oEvent = document.createEvent('KeyboardEvent');

		// Chromium Hack
		Object.defineProperty(oEvent, 'keyCode', {
			get : function()
			{
				return this.keyCodeVal;
			}
		});
		Object.defineProperty(oEvent, 'which', {
			get : function()
			{
				return this.keyCodeVal;
			}
		});
		Object.defineProperty(oEvent, 'shiftKey', {
			get : function()
			{
				return false;
			}
		});
		Object.defineProperty(oEvent, 'altKey', {
			get : function()
			{
				return false;
			}
		});
		Object.defineProperty(oEvent, 'metaKey', {
			get : function()
			{
				return false;
			}
		});
		Object.defineProperty(oEvent, 'ctrlKey', {
			get : function()
			{
				return false;
			}
		});

		if (AscCommon.AscBrowser.isIE)
		{
			oEvent.preventDefault = function () {
				Object.defineProperty(this, "defaultPrevented", {get: function () {return true;}});
			};
		}

		if (oEvent.initKeyboardEvent)
		{
			oEvent.initKeyboardEvent("keydown", true, true, window, false, false, false, false, _code, _code);
		}
		else
		{
			oEvent.initKeyEvent("keydown", true, true, window, false, false, false, false, _code, 0);
		}

		oEvent.keyCodeVal = _code;

		_element.dispatchEvent(oEvent);
		return oEvent.defaultPrevented;
	}

	function PaintMessageLoop(interval, api)
	{
		this.isUseInterval = api.isMobileVersion !== true;
		this.interval = interval || 40;
		this.id = null;

		this.requestAnimationFrame = window.requestAnimationFrame ||
			window.webkitRequestAnimationFrame ||
			window.mozRequestAnimationFrame ||
			window.oRequestAnimationFrame ||
			window.msRequestAnimationFrame || null;
		this.cancelAnimationFrame = window.cancelRequestAnimationFrame ||
			window.webkitCancelAnimationFrame ||
			window.webkitCancelRequestAnimationFrame ||
			window.mozCancelRequestAnimationFrame ||
			window.oCancelRequestAnimationFrame ||
			window.msCancelRequestAnimationFrame || null;

		this.isUseRequestAnimationFrame = AscCommon.AscBrowser.isChrome || AscCommon.AscBrowser.isSafari;
		if (this.isUseRequestAnimationFrame && !this.requestAnimationFrame)
			this.isUseRequestAnimationFrame = false;

		this.requestAnimationOldTime = -1;

		this.engine = null;
		this.step = null;
	}

	PaintMessageLoop.prototype.Start = function(engine)
	{
		this.engine = engine;
		if (null !== this.id)
			return;

		if (this.isUseRequestAnimationFrame)
		{
			this.step = this._animation.bind(this);
		}
		else
		{
			this.step = this._timer.bind(this);
		}

		this.step();
	};
	PaintMessageLoop.prototype.Stop = function()
	{
		if (null === this.id)
			return;

		if (this.isUseRequestAnimationFrame)
		{
			this.cancelAnimationFrame.call(window, this.id);
		}
		else
		{
			this.clearTimeout(this.id);
		}

		this.id = null;
	};

	PaintMessageLoop.prototype._animation = function()
	{
		var now = Date.now();
		if (!this.isUseInterval || -1 === this.requestAnimationOldTime || (now >= (this.requestAnimationOldTime + this.interval)) || (now < this.requestAnimationOldTime))
		{
			this.requestAnimationOldTime = now;
			this.engine();
		}
		this.id = this.requestAnimationFrame.call(window, this.step);
	};

	PaintMessageLoop.prototype._timer = function()
	{
		this.engine();
		this.id = setTimeout(this.step, this.interval);
	};

	PaintMessageLoop.prototype.delayRun = function(_this, _func)
	{
		if (window.requestAnimationFrame)
		{
			if (undefined !== _this._delayRunId)
				window.cancelAnimationFrame(_this._delayRunId);

			_this._delayRunId = window.requestAnimationFrame(function () {
				_func.call(_this);
				delete _this._delayRunId;
			});
		}
		else
		{
			if (undefined !== _this._delayRunId)
				clearTimeout(_this._delayRunId);

			_this._delayRunId = setTimeout(function () {
				_func.call(_this);
				delete _this._delayRunId;
			}, 40);
		}
	};

	function isSupportDoublePx()
	{
		var isSupport = true;

		var oTestSpan       = document.createElement("span");
		oTestSpan.setAttribute("style", "font-size:8pt");
		document.body.appendChild(oTestSpan);
		var defaultView   = oTestSpan.ownerDocument.defaultView;
		var computedStyle = defaultView.getComputedStyle(oTestSpan, null);
		if (null != computedStyle)
		{
			var fontSize = computedStyle.getPropertyValue("font-size");
			if (-1 !== fontSize.indexOf("px") && parseFloat(fontSize) === parseInt(fontSize))
				isSupport = false;
		}
		document.body.removeChild(oTestSpan);

		return isSupport;
	}

	function checkMouseWhell(e, options)
	{
		let isSupportBidirectional = false;
		let isAllowHorizontal = false;
		let isUseMaximumDelta = false;

		if (options)
		{
			isSupportBidirectional = (true === options.isSupportBidirectional);
			isAllowHorizontal = (true === options.isAllowHorizontal);
			isUseMaximumDelta = (true === options.isUseMaximumDelta);
		}

		let delta  = 0;
		let deltaX = 0;
		let deltaY = 0;

		// delta
		if (undefined !== e.wheelDelta && 0 !== e.wheelDelta)
		{
			delta = -45 * e.wheelDelta / 120;
		}
		else if (undefined !== e.detail && 0 !== e.detail)
		{
			delta = 45 * e.detail / 3;
		}

		// y
		if (undefined !== e.wheelDeltaY)
		{
			deltaY = -45 * e.wheelDeltaY / 120;
		}
		else
			deltaY = delta;

		// x
		if (isAllowHorizontal)
		{
			if (undefined !== e.wheelDeltaX)
			{
				deltaX = -45 * e.wheelDeltaX / 120;
			}

			if (e.axis !== undefined && e.axis === e.HORIZONTAL_AXIS)
			{
				deltaY = 0;

				if (0 === deltaX)
					deltaX = delta;
			}
		}

		deltaX >>= 0;
		deltaY >>= 0;

		if (!isSupportBidirectional)
		{
			if (isUseMaximumDelta)
			{
				if (Math.abs(deltaY) >= Math.abs(deltaX))
					deltaX = 0;
				else
					deltaY = 0;
			}
			else
			{
				if (0 !== deltaX)
					deltaY = 0;
			}
		}

		return {
			x : deltaX,
			y : deltaY
		};
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon']                          = window['AscCommon'] || {};
	window['AscCommon'].g_mouse_event_type_down  = g_mouse_event_type_down;
	window['AscCommon'].g_mouse_event_type_move  = g_mouse_event_type_move;
	window['AscCommon'].g_mouse_event_type_up    = g_mouse_event_type_up;
	window['AscCommon'].g_mouse_button_left      = g_mouse_button_left;
	window['AscCommon'].g_mouse_button_center    = g_mouse_button_center;
	window['AscCommon'].g_mouse_button_right     = g_mouse_button_right;
	window['AscCommon'].MouseUpLock              = MouseUpLock;
	window['AscCommon'].CMouseEventHandler       = CMouseEventHandler;
	window['AscCommon'].CKeyboardEvent           = CKeyboardEvent;
	window['AscCommon'].global_mouseEvent        = global_mouseEvent;
	window['AscCommon'].global_keyboardEvent     = global_keyboardEvent;
	window['AscCommon'].check_KeyboardEvent      = check_KeyboardEvent;
	window['AscCommon'].check_KeyboardEvent2     = check_KeyboardEvent2;
	window['AscCommon'].check_MouseMoveEvent     = check_MouseMoveEvent;
	window['AscCommon'].CreateMouseUpEventObject = CreateMouseUpEventObject;
	window['AscCommon'].getMouseButton           = getMouseButton;
	window['AscCommon'].check_MouseUpEvent       = check_MouseUpEvent;
	window['AscCommon'].check_MouseDownEvent     = check_MouseDownEvent;
	window['AscCommon'].Window_OnMouseUp         = Window_OnMouseUp;
	window['AscCommon'].button_eventHandlers     = button_eventHandlers;
	window['AscCommon'].emulateKeyDown 			 = emulateKeyDown;
    window['AscCommon'].check_MouseClickOnUp 	 = check_MouseClickOnUp;

	window['AscCommon'].PaintMessageLoop 	     = PaintMessageLoop;
	window['AscCommon'].isSupportDoublePx 	     = isSupportDoublePx;

	window['AscCommon'].checkMouseWhell 	     = checkMouseWhell;

})(window);
