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
    /**
     * Plugin config.
     * @typed {Object} Config
	 *
     * @pr {string} name
	 * Plugin name which will be visible at the plugin toolbar.
     *
	 * @pr {?Object} nameLocale
	 * Translations for the name field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the plugin name translation for each language. Example: { "fr" : "trFr", "de" : "deTr" }.
     *
	 * @pr {string} guid
	 * Plugin identifier. It must be of the asc.{UUID} type.
	 *
	 * @pr {?string} [baseUrl=""]
	 * Path to the plugin. All the other paths are calculated relative to this path. In case the plugin is installed on the server, an additional parameter (path to the plugins) is added there. If baseUrl == "", the path to all plugins will be used.
     *
	 * @pr {Variation[]} variations
	 * Plugin variations, or subplugins, that are created inside the origin plugin.
     */

    /**
     * The editors which the plugin is available for:
	 * <b>word</b> - text document editor,
	 * <b>cell</b> - spreadsheet editor,
	 * <b>slide</b> - presentation editor,
	 * <b>pdf</b> - pdf editor.
	 * @typedef {("word" | "cell" | "slide" | "pdf")} editorType
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/editorType.js
	 */

    /**
	 * The data type selected in the editor and sent to the plugin:
     * <b>text</b> - the text data,
	 * <b>html</b> - HTML formatted code,
	 * <b>ole</b> - OLE object data,
     * <b>desktop</b> - the desktop editor data,
     * <b>destop-external</b> - the main page data of the desktop app (system messages),
     * <b>none</b> - no data will be send to the plugin from the editor,
	 * <b>sign</b> - the sign for the keychain plugin.
	 * @typedef {("text" | "html" | "ole" | "desktop" | "destop-external" | "none" | "sign")} initDataType
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/initDataType.js
	 */

    /**
     * Plugin event ("onDocumentContentReady", "onTargetPositionChanged", onClick", "onInputHelperClear", "onInputHelperInput", etc.).
     * @typedef {string} EventType
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/EventType.js
	 */

    /**
     * Plugin variations, or subplugins, that are created inside the origin plugin.
     * @typed {Object} Variation
     * @pr {string} description
	 * The description, i.e. what describes your plugin the best way.
	 *
     * @pr {?Object} descriptionLocale
	 * Translations for the description field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the plugin description translation for each language.
	 *
     * @pr {string} url
	 * Plugin entry point, i.e. an HTML file which connects the plugin.js file (the base file needed for work with plugins) and launches the plugin code.
	 *
     * @pr {string[]} icons (with support HiDPI)
	 * Plugin icon image files used in the editors. There can be several scaling types for plugin icons: 100%, 125%, 150%, 175%, 200%, etc.
	 *
     * @pr {?boolean} [isViewer=true]
	 * Specifies if the plugin is working when the document is available in the viewer mode only or not.
	 *
	 * @pr {boolean} isDisplayedInViewer
	 * Specifies if the plugin will be displayed in the viewer mode as well as in the editor mode (isDisplayedInViewer == true) or in the editor mode only (isDisplayedInViewer == false).
	 *
     * @pr {EditorType[]} EditorsSupport
	 * The editors which the plugin is available for ("word" - text document editor, "cell" - spreadsheet editor, "slide" - presentation editor).
	 *
     * @pr {boolean} [isVisual=true]
	 * Specifies if the plugin is visual (will open a window for some action, or introduce some additions to the editor panel interface) or non-visual (will provide a button (or buttons) which is going to apply some transformations or manipulations to the document).
     *
	 * @pr {boolean} [isModal=true]
	 * Specifies if the opened plugin window is modal (used for visual plugins only, and if isInsideMode is not true).
	 *
	 * @pr {boolean} [isInsideMode=false]
	 * Specifies if the plugin must be displayed inside the editor panel instead of its own window.
     *
	 * @pr {boolean} isCustomWindow
	 * Specifies if the plugin uses a custom window, without standard borders and buttons (used for modal plugins only).
	 *
	 * @pr {boolean} isSystem
	 * Specifies if the plugin is not displayed in the editor interface and is started in the background with the server (or desktop editors start) not interfering with the other plugins, so that they can work simultaneously.
     *
	 * @pr {initDataType} initDataType
	 * The data type selected in the editor and sent to the plugin: "text" - the text data, "html" - HTML formatted code, "ole" - OLE object data, "desktop" - the desktop editor data, "destop-external" - the main page data of the desktop app (system messages), "none" - no data will be send to the plugin from the editor.
     *
	 * @pr {string} initData
     * Is usually equal to "" - this is the data which is sent from the editor to the plugin at the plugin start (e.g. if initDataType == "text", the plugin will receive the selected text when run). It may also be equal to encryption in the encryption plugins.
	 *
	 * @pr {?boolean} isUpdateOleOnResize
	 * Specifies if an OLE object must be redrawn when resized in the editor using the vector object draw type or not (used for OLE objects only, i.e. initDataType == "ole").
     *
	 * @pr {?Button[]} buttons
	 * The list of skinnable plugin buttons used in the plugin interface (used for visual plugins with their own window only, i.e. isVisual == true && isInsideMode == false).
	 *
	 * @pr {?boolean} [initOnSelectionChanged=true]
	 * Specifies if the plugin watches the text selection events in the editor window.
	 *
	 * @pr {number[]} size
	 * Plugin window size.
	 *
	 * @pr {EventType[]} events
     * Plugin events.
     */

    /**
	 * The skinnable plugin button used in the plugin interface (used for visual plugins with their own window only, i.e. isVisual == true and isInsideMode == false).
	 * @typedef { Object } Button
	 * @property {string} text - The label which is displayed on the button.
	 * @property {boolean} [primary] - Defines if the button is primary or not. The primary flag affects the button skin only.
	 * @property {boolean} [isViewer] - Defines if the button is shown in the viewer mode only or not.
	 * @property {localeTranslate} [textLocale] - Translations for the text field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the button label translation for each language.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/Button.js
	 */

	/**
	 * The OLE object properties
	 * @typedef {Object} OLEProperties
	 * @property {string} data - OLE object data (internal format).
	 * @property {string} imgSrc - A link to the image (its visual representation) stored in the OLE object and used by the plugin.
	 * @property {string} guid - An identifier of the plugin which can edit the current OLE object and must be of the *asc.{UUID}* type.
	 * @property {number} width - The OLE object width measured in millimeters.
	 * @property {number} height - The OLE object height measured in millimeters.
	 * @property {number} widthPix - The OLE object image width in pixels.
	 * @property {number} heightPix - The OLE object image height in pixels.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/OLEProperties.js
	 */

    /**
     * Base class
     * @global
     * @class
     * @name Api
     */

    var Api = window["AscCommon"].baseEditorsApi;

    /**
     * Returns the editor version.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetVersion
     * @returns {string} - The editor version.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetVersion.js
	 */
    Api.prototype["pluginMethod_GetVersion"] = function() { return this.GetVersion(); };

    /**
     * Adds an OLE object to the current document position.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias AddOleObject
	 * @this Api
     * @param {OLEProperties} data - The OLE object properties.
    * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/AddOleObject.js
	 */
    Api.prototype["pluginMethod_AddOleObject"] = function(data) { return this.asc_addOleObject(data, true); };

    /**
     * Edits an OLE object in the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias EditOleObject
     * @param {OLEProperties} data - The OLE object properties.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/EditOleObject.js
	 */
    Api.prototype["pluginMethod_EditOleObject"] = function(data) { return this.asc_editOleObject(data); };


	/**
	 * Returns an array of the selected OLE objects.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias GetSelectedOleObjects
	 * @returns {OLEProperties[]} - An array of the *OLEProperties* objects containing the data about the OLE object parameters.
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetSelectedOleObjects.js
	 */
	Api.prototype["pluginMethod_GetSelectedOleObjects"] = function()
	{
		let oDrawingsController = this.getGraphicController();
		let aRes = [];
		if(!oDrawingsController) return aRes;
		let aSelectedOle = oDrawingsController.getSelectedOleObjects();
		for(let nIdx = 0; nIdx < aSelectedOle.length; ++nIdx)
		{
			aRes.push(aSelectedOle[nIdx].getPluginDataObject());
		}
		return aRes;
	};

    /**
	 * An object containing the font information.
     * @typedef {Object} FontInfo
     * @property {string} m_wsFontName The font name.
     * @property {string} m_wsFontPath The path to the file with the current font.
	 * @property {number} m_lIndex The font number in the file if there is more than one font in the file.
	 * @property {boolean} m_bBold Specifies if the font characters are bold or not.
	 * @property {boolean} m_bItalic Specifies if the font characters are italic or not.
	 * @property {boolean} m_bIsFixed Specifies if the current font is monospaced or not.
	 * @property {Array.<number>} m_aPanose The PANOSE Typeface Classification Number, a compact 10-byte description of the font critical visual characteristics, such as contrast, weight, and serif style.
	 * @property {number} m_ulUnicodeRange1 The Unicode range encompassed by the font file (Bits 0-31).
	 * @property {number} m_ulUnicodeRange2 The Unicode range encompassed by the font file (Bits 32-63).
	 * @property {number} m_ulUnicodeRange3 The Unicode range encompassed by the font file (Bits 64-95).
	 * @property {number} m_ulUnicodeRange4 The Unicode range encompassed by the font file (Bits 96-127).
	 * @property {number} m_ulCodePageRange1 The code pages encompassed by the font file (Bits 0-31).
	 * @property {number} m_ulCodePageRange2 The code pages encompassed by the font file (Bits 32-63).
	 * @property {number} m_usWeigth The visual weight (stroke blackness or thickness) of the font characters (1-1000).
	 * @property {number} m_usWidth The relative change from the normal aspect ratio (width to height ratio).
	 * @property {number} m_sFamilyClass The font family class which values are assigned by IBM to each font family.
	 * @property {number} m_eFontFormat The specific file type(s) used to store font data: <b>0</b> - *.fon, <b>1</b> - *.ttf, <b>2</b> - *.ttf, *.otf (CFF), <b>3</b> - unknown font format.
	 * @property {number} m_shAvgCharWidth The arithmetic average of the escapement (width) of all non-zero width glyphs in the font.
	 * @property {number} m_shAscent The height above the baseline for a clipping region.
	 * @property {number} m_shDescent The vertical extent below the baseline for a clipping region.
	 * @property {number} m_shLineGap The typographic line gap for the current font.
	 * @property {number} m_shXHeight The distance between the baseline and the approximate height of non-ascending lowercase letters measured in FUnits.
	 * @property {number} m_shCapHeight The distance between the baseline and the approximate height of uppercase letters measured in FUnits.
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/FontInfo.js
	 */

    /**
     * Returns the fonts list.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetFontList
     * @returns {FontInfo[]} - An array of the FontInfo objects containing the data about the used fonts.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetFontList.js
	 */
    Api.prototype["pluginMethod_GetFontList"] = function()
    {
        return AscFonts.g_fontApplication.g_fontSelections.SerializeList();
    };

    /**
     * Inserts text into the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias InputText
	 * @param {string} text - A string value that specifies the text to be inserted into the document.
	 * @param {string} textReplace - A string value that specifies the text to be replaced with a new text.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/InputText.js
	 */
    Api.prototype["pluginMethod_InputText"] = function(text, textReplace)
    {
        if (!this.canEdit() || this.isPdfEditor() || !AscCommon.g_inputContext)
            return;

        if (textReplace)
        {
            for (var i = 0; i < textReplace.length; i++)
                AscCommon.g_inputContext.emulateKeyDownApi(8);
        }

        AscCommon.g_inputContext.addText(text);
        AscCommon.g_inputContext.keyPressInput = "";
    };

	/**
	 * Pastes text in the HTML format into the document.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias PasteHtml
	 * @param {string} htmlText - A string value that specifies the text in the *HTML* format to be pasted into the document.
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/PasteHtml.js
	 */
	Api.prototype["pluginMethod_PasteHtml"] = function (htmlText) {
		if (!AscCommon.g_clipboardBase)
			return null;

		if (!this.canEdit())
			return null;

		let _elem = document.getElementById("pmpastehtml");
		if (_elem)
			return;

		window.g_asc_plugins && window.g_asc_plugins.setPluginMethodReturnAsync();
		_elem = document.createElement("div");
		_elem.id = "pmpastehtml";
		_elem.style.color = "rgb(0,0,0)";

		if (this.editorId === AscCommon.c_oEditorId.Word || this.editorId === AscCommon.c_oEditorId.Presentation) {
			let textPr = this.get_TextProps();
			if (textPr && textPr.TextPr) {
				if (undefined !== textPr.TextPr.FontSize)
					_elem.style.fontSize = textPr.TextPr.FontSize + "pt";

				_elem.style.fontWeight = (true === textPr.TextPr.Bold) ? "bold" : "normal";
				_elem.style.fontStyle = (true === textPr.TextPr.Italic) ? "italic" : "normal";

				let _color = textPr.TextPr.Color;
				if (_color)
					_elem.style.color = "rgb(" + _color.r + "," + _color.g + "," + _color.b + ")";
				else
					_elem.style.color = "rgb(0,0,0)";
			}
		} else if (this.editorId === AscCommon.c_oEditorId.Spreadsheet) {
			let props = this.asc_getCellInfo();

			if (props && props.font) {
				if (undefined != props.font.size)
					_elem.style.fontSize = props.font.size + "pt";

				_elem.style.fontWeight = (true === props.font.bold) ? "bold" : "normal";
				_elem.style.fontStyle = (true === props.font.italic) ? "italic" : "normal";
			}
		}

		_elem.innerHTML = htmlText;
		document.body.appendChild(_elem);
		this.incrementCounterLongAction();
		let b_old_save_format = AscCommon.g_clipboardBase.bSaveFormat;
		AscCommon.g_clipboardBase.bSaveFormat = false;
		let _t = this;

		this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, _elem, undefined, undefined, undefined,
			function () {
				_t.decrementCounterLongAction();

				let fCallback = function () {
					document.body.removeChild(_elem);
					_elem = null;
					AscCommon.g_clipboardBase.bSaveFormat = b_old_save_format;
				};
				if (_t.checkLongActionCallback(fCallback, null)) {
					fCallback();
				}
				window.g_asc_plugins &&	window.g_asc_plugins.onPluginMethodReturn(true);
			}
		);
	};

    /**
     * Pastes text into the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias PasteText
     * @param {string} text - A string value that specifies the text to be pasted into the document.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/PasteText.js
	 */
    Api.prototype["pluginMethod_PasteText"] = function(text)
    {
        if (!AscCommon.g_clipboardBase)
            return null;

        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, text);
    };

    /**
     * An object containing the data about all the macros from the document.
     * @typedef {Object} Macros
     * @property {Array.<string>} macrosArray - An array of macros codes (*[{"name": "Macros1", "value": "{macrosCode}"}]*).
     * @property {number} current - A current macro index.
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/Macros.js
	 */

    /**
     * Returns the document macros.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetMacros
     * @returns {Macros} - The Macros object containing the data about all the macros from the document
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetMacros.js
	 */
    Api.prototype["pluginMethod_GetMacros"] = function()
    {
        return this.asc_getMacros();
    };

    /**
     * Sets macros to the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias SetMacros
     * @param {Macros} data - The *Macros* object containing the data about all the macros from the document.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/SetMacros.js
	 */
    Api.prototype["pluginMethod_SetMacros"] = function(data)
    {
        return this.asc_setMacros(data);
    };

	/**
	 * Returns all VBA macros from the document.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias GetVBAMacros
	 * @returns {string | null} VBA xml macros.
	 * @since 7.3.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetVBAMacros.js
	 */
	Api.prototype["pluginMethod_GetVBAMacros"] = function()
	{
		return (this.vbaProject ? this.vbaProject.vbaXml : null);
	};

    /**
     * Specifies the start action for long operations.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias StartAction
     * @param {number} type - A value which defines an action type which can take <b>0</b> if this is an *Information* action or <b>1</b> if this is a *BlockInteraction* action.
	 * @param {string} description - A string value that specifies the description text for the start action of the operation.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/StartAction.js
	 */
    Api.prototype["pluginMethod_StartAction"] = function(type, description)
    {
        this.sync_StartAction((type == "Block") ? Asc.c_oAscAsyncActionType.BlockInteraction : Asc.c_oAscAsyncActionType.Information, description);
    };

    /**
     * Specifies the end action for long operations.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias EndAction
     * @param {number} type - A value which defines an action type which can take <b>"Block"</b> if this is the *BlockInteraction* action or <b>"Information</b> if this is the *Information* action.
     * @param {string} description - A string value that specifies the description text for the operation end action.
	 * @param {string} status - The error status code. If no error occurs, then an empty string is passed.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/EndAction.js
	 */
    Api.prototype["pluginMethod_EndAction"] = function(type, description, status)
    {
        this.sync_EndAction((type == "Block") ? Asc.c_oAscAsyncActionType.BlockInteraction : Asc.c_oAscAsyncActionType.Information, description);

        if (window["AscDesktopEditor"] && status != null && status != "")
        {
            // error!!!
            if (!window["AscDesktopEditor"]["IsLocalFile"]())
            {
                this.sendEvent("asc_onError", "Encryption error: " + status + ". The file was not compiled.", c_oAscError.Level.Critical);
                window["AscDesktopEditor"]["CryptoMode"] = 0;
            }
            else
            {
                this.sendEvent("asc_onError", "Encryption error: " + status + ". End-to-end encryption mode is disabled.", c_oAscError.Level.NoCritical);
                window["AscDesktopEditor"]["CryptoMode"] = 0;

                if (undefined !== window.LastUserSavedIndex)
                {
                    AscCommon.History.UserSavedIndex = window.LastUserSavedIndex;

                    if (this.editorId == AscCommon.c_oEditorId.Spreadsheet)
                        this.onUpdateDocumentModified(AscCommon.History.Have_Changes());
                    else
                        this.UpdateInterfaceState();
                }
            }

            window.LastUserSavedIndex = undefined;
            setTimeout(function() {

                window["AscDesktopEditor"]["buildCryptedEnd"](false);

            }, 500);

            return;
        }

        window.LastUserSavedIndex = undefined;
        if (this._callbackPluginEndAction)
        {
            this._callbackPluginEndAction.call(this);
        }
    };

    /**
     * Encrypts the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias OnEncryption
     * @param {object} obj - The encryption properties.
     * @param {string} obj.type - The type of encrypting operation:
     * <b>generatePassword</b> - generates a password for the document,
     * <b>getPasswordByFile</b> - sends the password when opening the document,
     * <b>encryptData</b> - encrypts changes when co-editing,
     * <b>decryptData</b> - decrypts changes when co-editing.
     * @param {string} obj.password - A string value specifying the password to access the document.
     * @param {string} obj.data - Encrypted/decrypted changes.
     * @param {boolean} obj.check - Checks if the encryption/decryption operation is successful or not (used only for *encryptData* or *decryptData* types).
     * @param {string} obj.docinfo - An unencrypted part of the encrypted file.
     * @param {string} obj.hash - A string value specifying a file hash (*sha256* by default).
     * @param {string} obj.error - A string value specifying an error that occurs (the "" value means that the operation is successful).
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/OnEncryption.js
	 */
    Api.prototype["pluginMethod_OnEncryption"] = function(obj)
    {
        var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
        switch (obj.type)
        {
            case "generatePassword":
            {
                if ("" == obj["password"])
                {
                    _editor.sendEvent("asc_onError", "There is no connection with the blockchain", c_oAscError.Level.Critical);
                    return;
                }
                if ("no_build" === obj["error"])
				{
					// проблемы - но такие, при которых просто не собираем файл...
					window["AscDesktopEditor"]["buildCryptedEnd"](true);
					return;
				}

                var _ret = _editor.getFileAsFromChanges();
                AscCommon.EncryptionWorker.isPasswordCryptoPresent = true;
                _editor.currentDocumentInfoNext = obj["docinfo"];
                window["AscDesktopEditor"]["buildCryptedStart"](_ret.data, _ret.header, obj["password"], obj["docinfo"] ? obj["docinfo"] : "");
                break;
            }
            case "getPasswordByFile":
            {
                if ("" != obj["password"])
                {
                    var _param = ("<m_sPassword>" + AscCommon.CopyPasteCorrectString(obj["password"]) + "</m_sPassword>");
                    _editor.currentPassword = obj["password"];
                    _editor.currentDocumentHash = obj["hash"];
                    _editor.currentDocumentInfo = obj["docinfo"];

                    AscCommon.EncryptionWorker.isPasswordCryptoPresent = true;

                    if (window.isNativeOpenPassword)
                    {
                        window["AscDesktopEditor"]["NativeViewerOpen"](obj["password"]);
                    }
                    else
                    {
                        window["AscDesktopEditor"]["SetAdvancedOptions"](_param);
                    }
                }
                else
                {
                    this._onNeedParams({ message : obj["message"] }, true);
                }
                break;
            }
            case "encryptData":
            case "decryptData":
            {
                AscCommon.EncryptionWorker.receiveChanges(obj);
                break;
            }
        }
    };

    /**
	 * An object containing the watermark properties.
     * @typedef {Object} watermark_on_draw
     * @property {number} transparent The watermark transparency degree.
     * @property {string} type The {@link /docbuilder/global#ShapeType shape type} which specifies the preset shape geometry for the current watermark.
	 * @property {number} width The watermark width measured in millimeters.
	 * @property {number} height The watermark height measured in millimeters.
	 * @property {number} rotate The watermark rotation angle measured in degrees.
	 * @property {Array.<number>} margins The text margins measured in millimeters in the watermark shape.
	 * @property {Array.<number> | string} fill The watermark fill color in the RGB format, or the URL to image (base64 support: data:image/png;...). The empty array [] means that the watermark has no fill.
     * @property {number} stroke-width The watermark stroke width measured in millimeters.
	 * @property {Array.<number>} stroke The watermark stroke color in the RGB format. The empty array [] means that the watermark stroke has no fill.
	 * @property {number} align The vertical text align in the watermark shape: <b>0</b> - bottom, <b>1</b> - center, <b>4</b> - top.
	 * @property {Array.<object>} paragraphs The array with paragraphs from the current watermark with their properties.
	 * @property {number} paragraphs.align The horizontal text align in the current paragraph: <b>0</b> - right, <b>1</b> - left, <b>2</b> - center, <b>3</b> - justify.
	 * @property {Array.<number>} paragraphs.fill The paragraph highlight in the RGB format. The empty array [] means that the paragraph is not highlighted.
	 * @property {number} paragraphs.linespacing The text linespacing in the current paragraph.
	 * @property {Array.<object>} paragraphs.runs The array with runs from the current paragraph with their properties.
	 * @property {string} paragraphs.runs.text The run text.
	 * @property {Array.<number>} paragraphs.runs.fill The text highlight in the RGB format. The empty array [] means that the text is not highlighted.
	 * @property {string} paragraphs.runs.font-family The text font family.
	 * @property {string} paragraphs.runs.font-size The text font size measured in points (pt).
	 * @property {boolean} paragraphs.runs.bold Defines if the current text is displayed bold or not.
	 * @property {boolean} paragraphs.runs.italic Defines if the current text is displayed italic or not.
	 * @property {boolean} paragraphs.runs.strikeout Defines if the current text is displayed struck through or not.
	 * @property {boolean} paragraphs.runs.underline Defines if the current text is displayed underlined or not.
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/watermark_on_draw.js
	 */

    /**
	 * An object containing the form properties.
     * @typedef {Object} fillForms
     * @property {object} tags The form tags which specify the content for each form type with such a tag.
     * @property {string} tags.text The text field value (some text).
	 * @property {string} tags.checkBox The checkbox form value (<b>true</b> - checked, <b>false</b> - unchecked).
	 * @property {string} tags.picture The image form value (a link to the image).
	 * @property {string} tags.comboBox The combo box form value (one of the items from the combo box list values).
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/fillForms.js
	 */

    /**
     * Sets the properties to the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias SetProperties
     * @param {object} obj - The document properties.
	 * @param {?boolean} obj.copyoutenabled - Disables copying from the editor if it is set to **false**.
	 * @param {?boolean} obj.hideContentControlTrack - Disables tracking the content control if it is set to **true**.
	 * @param {?string} obj.watermark_on_draw - A string value for {@link global#watermark_on_draw watermark properties} in JSON format.
     * @param {?boolean} obj.disableAutostartMacros - Sets a flag that specifies that macros are started automatically when the editor opens.
     * @param {?string} obj.fillForms - Sets rules in JSON format for filling document {@link global#fillForms forms} by tags.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/SetProperties.js
	 */
    Api.prototype["pluginMethod_SetProperties"] = function(obj)
    {
		if (!this.isDocumentLoadComplete && obj)
		{
			if (!this.setPropertiesObj)
			{
				this.setPropertiesObj = obj;
			}
			else
			{
				for (var item in obj)
					this.setPropertiesObj[item] = obj[item];
			}
			return;
		}

		if (this.setPropertiesObj)
		{
			if (!obj)
			{
				obj = this.setPropertiesObj;
			}
			else
			{
				for (var item in this.setPropertiesObj)
					if (!obj[item]) obj[item] = this.setPropertiesObj[item];
			}
			delete this.setPropertiesObj;
		}

        if (!obj)
            return;

        for (var prop in obj)
        {
            switch (prop)
            {
                case "copyoutenabled":
                {
                    this.copyOutEnabled = obj[prop];
                    this.sync_CanCopyCutCallback(this.copyOutEnabled);
                    break;
                }
                case "watermark_on_draw":
                {
                    var sText = "";
                    var tempProp = {};
                    try
                    {
                        tempProp = (typeof obj[prop] === "string") ? JSON.parse(obj[prop]) : obj[prop];
                    }
                    catch (err)
                    {
                        tempProp = {};
                    }
                    if (tempProp["paragraphs"])
                    {
                        tempProp["paragraphs"].forEach(function (el) {
                            if (el["runs"])
                            {
                                sText += el["runs"].reduce(function (accum, curel) {
                                    return accum + (curel["text"] ? curel["text"] : "");
                                }, "");
                            }
                        });
                    }
                    
                    if(!(typeof sText === "string"))
                        sText = "";
                    
                    AscFonts.FontPickerByCharacter.checkText(sText, this, function () {
                        this.watermarkDraw = obj[prop] ? new AscCommon.CWatermarkOnDraw(obj[prop], this) : null;
                        this.watermarkDraw.checkOnReady();
                    });
                    break;
                }
                case "hideContentControlTrack":
                {
                    if (this.editorId === AscCommon.c_oEditorId.Word && this.WordControl && this.WordControl.m_oLogicDocument && !this.isPdfEditor())
                        this.WordControl.m_oLogicDocument.SetForceHideContentControlTrack(obj[prop]);

                    break;
                }
                case "disableAutostartMacros":
                {
                    this.disableAutostartMacros = true;
					break;
                }
				case "fillForms":
				{
					if (this.editorId !== AscCommon.c_oEditorId.Word
						|| !this.WordControl
						|| !this.WordControl.m_oLogicDocument
						|| this.isPdfEditor())
						break;

					let oLogicDocument = this.WordControl.m_oLogicDocument;

					let oMap;
					try
					{
						oMap = (typeof obj[prop] === "string") ? JSON.parse(obj[prop]) : obj[prop];
					}
					catch (err)
					{
						oMap = {};
					}

					if (oMap["tags"])
					{
						let arrControls = oLogicDocument.GetAllContentControls();

						let oTags         = {};
						let arrCheckLocks = [];
						let arrUrls       = [];

						for (let sTag in oMap["tags"])
						{
							oTags[sTag] = [];
							for (let nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
							{
								let oForm = arrControls[nIndex];
								if (oForm
									&& oForm.IsForm()
									&& sTag === oForm.GetTag())
								{
									oTags[sTag].push(oForm);

									let oElement;
									if (oForm.IsInlineLevel())
										oElement = oForm.GetParagraph();
									else
										oElement = oForm;

									if (oElement && -1 === arrCheckLocks.indexOf(oElement))
										arrCheckLocks.push(oElement);

									if (oMap["tags"][sTag]["picture"])
										arrUrls.push(oMap["tags"][sTag]["picture"]);
								}
							}
						}

						function FillForms()
						{
							if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_None, {
								Type      : AscCommon.changestype_2_ElementsArray_and_Type,
								Elements  : arrCheckLocks,
								CheckType : AscCommon.changestype_Paragraph_Content
							}))
							{
								oLogicDocument.StartAction(AscDFH.historydescription_Document_FillFormsByTags);

								for (let sTag in oTags)
								{
									let oValue = oMap["tags"][sTag];
									for (let nFormIndex = 0, nFormsCount = oTags[sTag].length; nFormIndex < nFormsCount; ++nFormIndex)
									{
										let oForm = oTags[sTag][nFormIndex];
										if (oForm.IsComboBox() || oForm.IsDropDownList())
										{
											if (oValue["comboBox"])
												oForm.SelectListItem(oValue["comboBox"]);
											else if (oForm.IsComboBox() && oValue["text"])
												oForm.SetInnerText(oValue["text"]);
										}
										else if (oForm.IsPictureForm())
										{
											let oPicture = oForm.GetPicture();
											if (oValue["picture"] && oPicture)
												oPicture.setBlipFill(AscFormat.CreateBlipFillRasterImageId(oValue["picture"]));
										}
										else if (oForm.IsCheckBox())
										{
											if (oValue["checkBox"])
												oForm.SetCheckBoxChecked(!!oValue["checkBox"]);
										}
										else
										{
											oForm.SetInnerText(oValue["text"]);
										}
									}
								}

								oLogicDocument.Recalculate();
								oLogicDocument.FinalizeAction();
							}
						}

						if (this.ImageLoader && arrUrls.length)
							this.ImageLoader.LoadImagesWithCallback(arrUrls, FillForms, []);
						else
							FillForms();
					}

					break
				}
                default:
                    break;
            }
        }
    };

	/**
	 * The plugin options.
     * @typedef {Object} PluginOptions
     * @property {object} all The parameters which will be set for all plugins ({ "all" : { key, value } }).
     * @property {object} plugin_guid The parameters which will be set for a specific plugin. The plugin must be specified with the plugin GUID of the asc.{UUID} type ({ "plugin_guid" : { keyForSpecificPlugin : valueForSpecificPlugin } }).
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/PluginOptions.js
	 */
	
	/**
	 * Configures plugins from an external source. The settings can be set for all plugins or for a specific plugin.
	 * For example, this method can be used to pass an authorization token to the plugin. This method can be used only with the connector class.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias SetPluginsOptions
	 * @param {PluginOptions} options - Plugin options.
	 * @since 8.1.1
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/SetPluginsOptions.js
	 */
	Api.prototype["pluginMethod_SetPluginsOptions"] = function(options)
	{
		let guid = window.g_asc_plugins.getCurrentPluginGuid();
		let runObject = window.g_asc_plugins.runnedPluginsMap[guid];
		if (!runObject.isConnector)
			return;
		this.setPluginsOptions(options);
	};

    /**
     * Shows the input helper.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias ShowInputHelper
     * @param {string} guid - A string value which specifies a plugin identifier which must be of the *asc.{UUID}* type.
     * @param {number} w - A number which specifies the window width measured in millimeters.
     * @param {number} h - A number which specifies the window height measured in millimeters.
     * @param {boolean} isKeyboardTake - Defines if the keyboard is caught (**true**) or not (**alse**).
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ShowInputHelper.js
	 */
    Api.prototype["pluginMethod_ShowInputHelper"] = function(guid, w, h, isKeyboardTake)
    {
        var _frame = document.getElementById("iframe_" + guid);
        if (!_frame)
            return;

        var _offset = this.getTargetOnBodyCoords();
        if (w > _offset.W)
            w = _offset.W;
        if (h > _offset.H)
            h = _offset.H;

        var _offsetToFrame = 10;
        var _r = _offset.X + _offsetToFrame + w;
        var _t = _offset.Y - _offsetToFrame - h;
        var _b = _offset.Y + _offset.TargetH + _offsetToFrame + h;

        var _x = _offset.X + _offsetToFrame;
        if (_r > _offset.W)
            _x += (_offset.W - _r);

        var _y = 0;

        if (_b < _offset.H)
        {
            _y = _offset.Y + _offset.TargetH + _offsetToFrame;
        }
        else if (_t > 0)
        {
            _y = _t;
        }
        else
        {
            _y = _offset.Y + _offset.TargetH + _offsetToFrame;
            h += (_offset.H - _b);
        }

        _frame.style.left = _x + "px";
        _frame.style.top = _y + "px";
        _frame.style.width = w + "px";
        _frame.style.height = h + "px";

        if (!this.isMobileVersion)
            _frame.style.zIndex = 1000;
        else
            _frame.style.zIndex = 5001;

        if (!_frame.style.boxShadow)
        {
            _frame.style.boxShadow = "0 6px 12px rgba(0, 0, 0, 0.175)";
            _frame.style.webkitBoxShadow = "0 6px 12px rgba(0, 0, 0, 0.175)";
            //_frame.style.borderRadius = "3px";
        }


        if (isKeyboardTake)
        {
            _frame.setAttribute("oo_editor_input", "true");
            _frame.focus();
        }
        else
        {
            _frame.removeAttribute("oo_editor_input");
            if (AscCommon.g_inputContext)
            {
                AscCommon.g_inputContext.isNoClearOnFocus = true;
                AscCommon.g_inputContext.HtmlArea.focus();
            }
        }

        if (AscCommon.g_inputContext)
        {
            AscCommon.g_inputContext.isInputHelpersPresent = true;
            AscCommon.g_inputContext.isInputHelpers[guid] = true;
        }
    };

    /**
     * Unshows the input helper.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias UnShowInputHelper
     * @param {string} guid - A string value which specifies a plugin identifier which must be of the *asc.{UUID}* type.
     * @param {string} isclear - Defines if the input context will be cleared (**true**) or not (**false**).
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/UnShowInputHelper.js
	 */
    Api.prototype["pluginMethod_UnShowInputHelper"] = function(guid, isclear)
    {
        var _frame = document.getElementById("iframe_" + guid);
        if (!_frame)
            return;

        _frame.style.width = "10px";
        _frame.style.height = "10px";
        _frame.removeAttribute("oo_editor_input");

        _frame.style.zIndex = -1000;

        if (AscCommon.g_inputContext && AscCommon.g_inputContext.HtmlArea)
        {
            AscCommon.g_inputContext.HtmlArea.focus();

            if (AscCommon.g_inputContext.isInputHelpers[guid])
                delete AscCommon.g_inputContext.isInputHelpers[guid];

            var count = 0;
            for (var test in AscCommon.g_inputContext.isInputHelpers)
            {
                if (AscCommon.g_inputContext.isInputHelpers[test])
                    count++;
            }

            AscCommon.g_inputContext.isInputHelpersPresent = (0 != count);
        }

        if (AscCommon.g_inputContext && isclear)
        {
            AscCommon.g_inputContext.keyPressInput = "";
        }
    };
    /**
     * Sends a message to the co-authoring chat.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias CoAuthoringChatSendMessage
     * @param {string} sText - Message text.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/CoAuthoringChatSendMessage.js
	 */
    Api.prototype["pluginMethod_CoAuthoringChatSendMessage"] = function(sText)
    {
        return this.CoAuthoringChatSendMessage(sText);
    };

	/**
	 * The current selection type ("none", "text", "drawing", or "slide").
	 * @typedef {("none" | "text" | "drawing" | "slide")} SelectionType
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/SelectionType.js
	 */

	/**
	 * Returns the type of the current selection.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias GetSelectionType
	 * @returns {SelectionType} - The selection type.
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetSelectionType.js
	 */
	Api.prototype["pluginMethod_GetSelectionType"] = function()
	{
		switch (this.editorId)
		{
			case AscCommon.c_oEditorId.Word:
			{
				if (!this.WordControl || !this.WordControl.m_oLogicDocument)
					return "none";
				
				let logicDoc = this.WordControl.m_oLogicDocument;
				if (!logicDoc.IsSelectionUse())
					return "none";

				if (logicDoc.IsTextSelectionUse())
					return "text";
				
				if (logicDoc.DrawingObjects.getSelectedObjectsBounds())
					return "drawing";
				
				return "none";
			}
			case AscCommon.c_oEditorId.Presentation:
			{
				if (!this.WordControl || !this.WordControl.m_oLogicDocument)
					return "none";
				var logicDoc = this.WordControl.m_oLogicDocument;

				if (-1 === logicDoc.CurPage)
					return "none";

				var _controller = logicDoc.GetCurrentSlide().graphicObjects;
				var _elementsCount = _controller.selectedObjects.length;

				var retType = "slide";
				if (!_controller.IsSelectionUse() && _elementsCount > 0)
					retType = "none";

				var selectionBounds = logicDoc.GetSelectionBounds();
				var eps = 0.0001;
				if (selectionBounds && selectionBounds.Start && selectionBounds.End &&
					(Math.abs(selectionBounds.Start.W) > eps) &&
					(Math.abs(selectionBounds.End.W) > eps))
				{
					return "text";
				}
				
				if (retType === "slide" && _controller.getSelectedObjectsBounds())
					retType = "drawing";

				return retType;
			}
			case AscCommon.c_oEditorId.Spreadsheet:
			{
				if (!this.wb || !this.wb.getWorksheet())
					return "none";

				var objectRender = this.wb.getWorksheet().objectRender;
				if (!objectRender)
					return "none";

				var controller = objectRender.controller;
				if (!controller)
					return "none";

				var selection = this.wb.GetSelectionRectsBounds();
				var retType = "none";

				if (!controller.IsSelectionUse() && !selection)
					retType = "none";
				if (controller.GetSelectionBounds() || selection)
					retType = "text";
				if (controller.getSelectedObjectsBounds())
					retType = "drawing";

				return retType;
			}
			default:
				break;
		}
		return "none";
	};
     /**
     * Converts a document to Markdown or HTML text.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias ConvertDocument
     * @param {"markdown" | "html"} [sConvertType="markdown"] - Conversion type ("markdown" or "html").
     * @param {boolean} [bHtmlHeadings=false] - Defines if the HTML headings and IDs will be generated when the Markdown renderer of your target platform does not handle Markdown-style IDs.
	 * @param {boolean} [bBase64img=false] - Defines if the images will be created in the base64 format.
	 * @param {boolean} [bDemoteHeadings=false] - Defines if all heading levels in your document will be demoted to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {boolean} [bRenderHTMLTags=false] - Defines if HTML tags will be preserved in your Markdown. If you just want to use an occasional HTML tag, you can avoid using the opening angle bracket in the following way: \<tag>text\</tag>. By default, the opening angle brackets will be replaced with the special characters.
     * @return {string} - The Markdown/HTML text.
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ConvertDocument.js
	 */
    Api.prototype["pluginMethod_ConvertDocument"] = function(sConvertType, bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags)
    {
        return this.ConvertDocument(sConvertType, bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags);
    };
    /**
     * Returns the selected text from the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CPE", "CSE"]
     * @alias GetSelectedText
     * @param {object} prop - The resulting string display properties.
     * @param {boolean} prop.Numbering - Defines if the resulting string will include numbering or not.
     * @param {boolean} prop.Math - Defines if the resulting string will include mathematical expressions or not.
     * @param {string} [prop.TableCellSeparator='\t'] - Defines how the table cell separator will be specified in the resulting string. Any symbol can be used. The default separator is "\t".
     * @param {string} [prop.TableRowSeparator='\r\n'] - Defines how the table row separator will be specified in the resulting string. Any symbol can be used. The default separator is "\r\n".
     * @param {string} [prop.ParaSeparator='\r\n'] - Defines how the paragraph separator will be specified in the resulting string. Any symbol can be used. The default separator is "\r\n".
     * @param {string} [prop.TabSymbol='\t'] - Defines how the tab will be specified in the resulting string. Any symbol can be used. The default symbol is "\t".
     * @param {string} [prop.NewLineSeparator='\r'] - Defines how the line separator will be specified in the resulting string. Any symbol can be used. The default separator is "\r".
	 * @return {string} - Selected text.
     * @since 7.1.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetSelectedText.js
	 */
    Api.prototype["pluginMethod_GetSelectedText"] = function(prop)
    {
        var properties;
        if (typeof prop === "object")
        {
            properties =
            {
                Numbering : (prop.hasOwnProperty("Numbering")) ? prop["Numbering"] : true,
                Math : (prop.hasOwnProperty("Math")) ? prop["Math"] : true,
                TableCellSeparator: prop["TableCellSeparator"],
                TableRowSeparator: prop["TableRowSeparator"],
                ParaSeparator: prop["ParaSeparator"],
                NewLineSeparator: prop["NewLineSeparator"],
                TabSymbol: prop["TabSymbol"]
            }
        }
        else
        {
            properties =
            {
                Numbering : true
            }
        }

        return this.asc_GetSelectedText(false, properties);
    };
	/**
	 * Returns the selected content in the specified format.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @alias GetSelectedContent
	 * @param {object} prop  - The returned content properties.
	 * @param {"text" | "html"} [prop.type="text"] - The format type of the returned content (text or HTML).
	 * @returns {string} - The selected content.
	 * @since 8.3.1
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetSelectedContent.js
	 */
	Api.prototype["pluginMethod_GetSelectedContent"] = function(prop)
	{
		let type = AscCommon.c_oAscClipboardDataFormat.Text;
		if (prop && "html" === prop["type"])
			type = AscCommon.c_oAscClipboardDataFormat.Html;
			
		return this.getSelectedContent(type);
	};
    /**
     * Replaces each paragraph (or text in cell) in the select with the corresponding text from an array of strings.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias ReplaceTextSmart
     * @param {string[]} arrString - An array of replacement strings.
	 * @param {string} [sParaTab="\t"] - A character which is used to specify the tab in the source text. Any symbol can be used. The default separator is "\t".
     * @param {string} [sParaNewLine="\r\n"] - A character which is used to specify the line break character in the source text. Any symbol can be used. The default separator is "\r\n".
     * @returns {boolean} - Always returns true.
     * @since 7.1.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ReplaceTextSmart.js
	 */
    Api.prototype["pluginMethod_ReplaceTextSmart"] = function(arrString, sParaTab, sParaNewLine)
    {
		window.g_asc_plugins && window.g_asc_plugins.setPluginMethodReturnAsync();
		this.incrementCounterLongAction();

		function ReplaceTextSmart()
		{
			this.asc_canPaste();
			this.ReplaceTextSmart(arrString, sParaTab, sParaNewLine);
			this.asc_Recalculate(true);
			switch (this.editorId)
			{
				case AscCommon.c_oEditorId.Spreadsheet:
					this.asc_endPaste();
					break;
				case AscCommon.c_oEditorId.Word:
				case AscCommon.c_oEditorId.Presentation:
					this.WordControl.m_oLogicDocument.FinalizeAction();
					break;
			}

			this.decrementCounterLongAction();

			window.g_asc_plugins && window.g_asc_plugins.onPluginMethodReturn(true);
		}

		let sOverAll = "";
		for (let nIndex = 0, nCount = arrString.length; nIndex < nCount; ++nIndex)
			sOverAll += arrString[nIndex];

		AscFonts.FontPickerByCharacter.checkText(sOverAll, this, ReplaceTextSmart);
        return true;
    };
	/**
     * Returns the current file to download in the specified format.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetFileToDownload
     * @param {string} [format=" "] - A format in which you need to download a file.
     * @returns {string} - URL to download the file in the specified format or error.
     * @since 7.2.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetFileToDownload.js
	 */
	Api.prototype["pluginMethod_GetFileToDownload"] = function(format)
	{
		window.g_asc_plugins && window.g_asc_plugins.setPluginMethodReturnAsync();
		
		if (format && typeof(format) === "string")
			format = format.toUpperCase();
		
		let dwnldF = Asc.c_oAscFileType[format] || Asc.c_oAscFileType[this.DocInfo.Format.toUpperCase()];
		let opts = new Asc.asc_CDownloadOptions(dwnldF);
		let _t = this;
		opts.callback = function() {
			_t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.DownloadAs);
			_t.fCurCallback = function(res) {
				let data = (res.status == "ok") ? res.data : "error";
				window.g_asc_plugins && window.g_asc_plugins.onPluginMethodReturn(data);
			};
		}
		this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, opts);
	};

	/**
	 * Specifies how to adjust the image object in case of replacing the selected image.
	 * @typedef {("fill" | "fit" | "original" | "stretch")} ReplaceImageMode
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ReplaceImageMode.js
	 */

    /**
     * An object containing the information about the base64 encoded *png* image.
     * @typedef {Object} ImageData
     * @property {string} src The image source in the base64 format.
     * @property {number} width The image width in pixels.
     * @property {number} height The image height in pixels.
     * @property {?ReplaceImageMode} replaceMode Specifies how to adjust the image object in case of replacing the selected image.
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ImageData.js
	 */

	/**
     * Returns the image data from the first of the selected drawings. If there are no drawings selected, the method returns a white rectangle.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetImageDataFromSelection
     * @returns {?ImageData} - The ImageData object containig the information about the base64 encoded png image.
     * @since 7.2.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetImageDataFromSelection.js
	 */
	Api.prototype["pluginMethod_GetImageDataFromSelection"] = function()
	{
		return this.getImageDataFromSelection();
	};
	/**
     * Replaces the first selected drawing with the image specified in the parameters.
     * If there are no drawings selected, the method inserts the image at the current position.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias PutImageDataToSelection
     * @param {ImageData} oImageData - The information about the base64 encoded *png* image.
     * @since 7.2.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/PutImageDataToSelection.js
	 */
	Api.prototype["pluginMethod_PutImageDataToSelection"] = function(oImageData)
	{
		if(!this.canEdit() || this.isPdfEditor())
		{
			return;
		}
		window.g_asc_plugins.setPluginMethodReturnAsync();
		let sImgSrc = oImageData["src"];
		this.asc_checkImageUrlAndAction(sImgSrc, function(oImage)
		{
			let nWidth = oImageData["width"];
			let nHeight = oImageData["height"];
			const isN = AscFormat.isRealNumber;
			if(!isN(nWidth) || !isN(nHeight))
			{
				nWidth = oImage.Image.width;
				nHeight = oImage.Image.height;
			}
			this.putImageToSelection(AscCommon.g_oDocumentUrls.getImageLocal(oImage.src), nWidth, nHeight, oImageData["replaceMode"]);
			window.g_asc_plugins.onPluginMethodReturn();
		});
	};

	function getLocalStorageItem(key)
	{
		try
		{
			return JSON.parse(window.localStorage.getItem(key));
		}
		catch (e)
		{
			return null;
		}
	}
	function setLocalStorageItem(key, value)
	{
		try
		{
			window.localStorage.setItem(key, JSON.stringify(value));
			return true;
		}
		catch (e)
		{
		}
		return false;
	}
	AscCommon.getLocalStorageItem = getLocalStorageItem;
	AscCommon.setLocalStorageItem = setLocalStorageItem;

	function installPlugin(config, loadFuncName)
	{
		if (!config)
		{
			return {
				"type" : loadFuncName,
				"guid" : ""
			};
		}

		window.g_asc_plugins.isUICheckOnInitMessage = true;
		setTimeout(function(){
			if (window.g_asc_plugins.isUICheckOnInitMessage)
				delete window.g_asc_plugins.isUICheckOnInitMessage;
		});

		// desktop detecting (it's necessary when we work with clouds into desktop)
		const isLocal = ( (window["AscDesktopEditor"] !== undefined) && (window.location.protocol.indexOf('file') !== -1) );
		if (isLocal)
		{
			// Отдаём весь конфиг, внутри вычислим путь к deploy
			// TODO: отслеживать возможные ошибки при +/- плагинов: из ++кода отправлять статус операции и на основе его отправлять в менеджер плагинов корректный ответ.
			// UPD: done. Ничего не изменять в менеджере плагинов, если guid пуст

            let result = window["AscDesktopEditor"]["PluginInstall"](JSON.stringify(config));

			if (result && window.g_asc_plugins.isRunned(config["guid"]))
			{
				window.g_asc_plugins.close(config["guid"]);
			}
			
			return {
				"type" : loadFuncName,
				"guid" : result ? config["guid"] : ""
			};
		}

		let currentInstalledPlugins = getLocalStorageItem("asc_plugins_installed");
		if (!currentInstalledPlugins)
			currentInstalledPlugins = {};
		currentInstalledPlugins[config["guid"]] = config;
		setLocalStorageItem("asc_plugins_installed", currentInstalledPlugins);

		let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");
		if (currentRemovedPlugins && currentRemovedPlugins[config["guid"]])
		{
			delete currentRemovedPlugins[config["guid"]];
			setLocalStorageItem("asc_plugins_removed", currentRemovedPlugins);
		}

		window.g_asc_plugins.api.disableCheckInstalledPlugins = true;
		window.g_asc_plugins.loadExtensionPlugins([config], true);
		delete window.g_asc_plugins.api.disableCheckInstalledPlugins;

		return {
			"type" : loadFuncName,
			"guid" : config["guid"]
		};
	}

	Api.prototype.getUsedBackgroundPlugins = function()
	{
		let services = [];
		try
		{
			services = JSON.parse(window.localStorage.getItem("asc_plugins_background"));
			if (!services)
				services = [];
		}
		catch (e)
		{
			services = [];
		}
		return services;
	};
	Api.prototype["getUsedBackgroundPlugins"] = Api.prototype.getUsedBackgroundPlugins;

	Api.prototype.setUsedBackgroundPlugins = function(services)
	{
		window.localStorage.setItem("asc_plugins_background", JSON.stringify(services));
	};

	Api.prototype.checkInstalledPlugins = function()
	{
		if (this.disableCheckInstalledPlugins)
			return;

		const isLocal = ( (window["AscDesktopEditor"] !== undefined) && (window.location.protocol.indexOf('file') !== -1) );
		if (isLocal) {
			// В случае Desktop не работаем с localStorage и extensions, этот метод может быть вызван из интерфейса
			// если по какой-то причине (неактуальный cache) у пользователя есть asc_plugins_installed, asc_plugins_removed, то их нужно игнорировать/удалить
			return;
		}

		let arrayPlugins = [];

		let currentInstalledPlugins = getLocalStorageItem("asc_plugins_installed");
		if (currentInstalledPlugins)
		{
			for (let item in currentInstalledPlugins)
			{
				if (currentInstalledPlugins[item]["guid"])
					arrayPlugins.push(currentInstalledPlugins[item]);
			}
		}

		if (window["Asc"]["extensionPlugins"] && window["Asc"]["extensionPlugins"].length)
		{
			let arrayExtensions = window["Asc"]["extensionPlugins"];
			for (let i = 0, len = arrayExtensions.length; i < len; i++)
				arrayPlugins.push(arrayExtensions[i]);
		}

		let isInstalledPresent = window.g_asc_plugins.loadExtensionPlugins(arrayPlugins, undefined, true);

		let isRemovedPresent = false;
		let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");

		if (currentRemovedPlugins)
		{
			for (let guid in currentRemovedPlugins)
			{
				if (guid)
				{
					if (window.g_asc_plugins.unregister(guid))
						isRemovedPresent = true;
				}
			}
		}

		// этот метод может быть вызван из интерфейса - нужен таймаут для web-apps
		if (isRemovedPresent || isInstalledPresent) {

			setTimeout(function () {

				// в принципе можно не удалять, так как если ничего не поменялось - то не зайдем второй раз сюда.
				// но зачем еще раз парсить
				window.g_asc_plugins.api.disableCheckInstalledPlugins = true;

				if (isRemovedPresent)
					window.g_asc_plugins.api.sendEvent("asc_onPluginsReset");

				if (isInstalledPresent || isRemovedPresent)
					window.g_asc_plugins.updateInterface();

				delete window.g_asc_plugins.api.disableCheckInstalledPlugins;

			}, 10);

		}
	};

    /**
	 * The plugin object.
     * @typedef {Object} PluginData
     * @property {string} url The URL to plugin config.
     * @property {string} guid The plugin identifier. It must be of the *asc.{UUID}* type.
	 * @property {boolean} canRemoved Specifies if the plugin can be removed (**true**) or not (**false**).
     * @property {object} obj The {@link /plugin/config config} of the installed plugin. The version is taken from the config and compared with the current one to check for updates.
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/PluginData.js
	 */

	/**
    * Returns all the installed plugins.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetInstalledPlugins
     * @returns {PluginData[]} - An array of all the installed plugins.
     * @since 7.2.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetInstalledPlugins.js
	 */
	Api.prototype["pluginMethod_GetInstalledPlugins"] = function()
	{
		/*
			формат объекта 
			{
				url: url на конфиг (хотя по факту он не нужен, так как конфиг есть в этом объекте и внутри маркетплейса тоже),
				guid: guid плагина,
				canRemoved: флаг, может ли быть удалён плагин или нет (true/false),
				obj: конфиг установленного плагина (от туда берется версия и сравнивается с текущей для проверки обновлений)
			}
		*/

		const isLocal = ( (window["AscDesktopEditor"] !== undefined) && (window.location.protocol.indexOf('file') !== -1) );

		// В случае Desktop нужно проверить какие плагины нельзя удалять. В UpdateInstallPlugins работаем с двумя типами папок.
		// Пока проверка тут, но грамотнее будет сделать и использовать доп.свойство isSystemInstall класса CPlugin
		// т.к. не будем лишний раз парсить папки, только при +/- плагинов.
		let protectedPlugins = [];

		if (isLocal) {
			var _pluginsTmp = JSON.parse(window["AscDesktopEditor"]["GetInstallPlugins"]());

			var len = _pluginsTmp[0]["pluginsData"].length;
			for (var i = 0; i < len; i++) {
				protectedPlugins.push(_pluginsTmp[0]["pluginsData"][i]["guid"]);
			}
			
			// Также смотрим плагины из папки пользователя, возможно там есть обновленные системные
			len = _pluginsTmp[1]["pluginsData"].length;
			for (var i = 0; i < len; i++) {
				if (_pluginsTmp[1]["pluginsData"][i]["canRemoved"] === false)
					protectedPlugins.push(_pluginsTmp[1]["pluginsData"][i]["guid"]);
			}
		}

		let baseUrl = window.location.href;
		let posQ = baseUrl.indexOf("?");
		if (-1 !== posQ)
			baseUrl = baseUrl.substr(0, posQ);

		let pluginsArray = window.g_asc_plugins.plugins.concat(window.g_asc_plugins.systemPlugins);
		let returnArray = [];

		for (let i = 0, len = pluginsArray.length; i < len; i++)
		{
			if (pluginsArray[i].isConnector)
				continue;
			returnArray.push({
				"baseUrl" : baseUrl,
				"guid" : pluginsArray[i].guid,
				"canRemoved" : protectedPlugins.indexOf(pluginsArray[i].guid) == -1,
				"obj" : pluginsArray[i].serialize(),
				"removed" : false
			});
		}

		if (isLocal)
			return returnArray;

		// нужно послать и удаленные. так как удаленный может не быть в сторе. тогда его никак не установить обратно
		let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");

		if (currentRemovedPlugins)
		{
			for (let guid in currentRemovedPlugins)
			{
				if (currentRemovedPlugins[guid])
				{
					returnArray.push({
						"baseUrl" : baseUrl,
						"guid": currentRemovedPlugins[guid]["guid"],
						"canRemoved": true,
						"obj": currentRemovedPlugins[guid],
						"removed" : true
					});
				}
			}
		}

		return returnArray;
	};
	/**
    * Removes a plugin with the specified GUID.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {string} guid - The plugin identifier. It must be of the *asc.{UUID}* type.
	 * @param {string} backup - The plugin backup. This parameter is used when working with the desktop editors.
     * @alias RemovePlugin
     * @returns {object} - An object with the result information.
     * @since 7.2.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/RemovePlugin.js
	 */
	Api.prototype["pluginMethod_RemovePlugin"] = function(guid, backup)
	{
		let removedPlugin = window.g_asc_plugins.unregister(guid);
		const isLocal = ( (window["AscDesktopEditor"] !== undefined) && (window.location.protocol.indexOf('file') !== -1) );

		if (isLocal)
		{
			// Вызываем только этот ++код, никаких дополнительных действий типа:
			// window.g_asc_plugins.unregister(guid), window["UpdateInstallPlugins"](), this.sendEvent("asc_onPluginsReset"), window.g_asc_plugins.updateInterface()
			// не требуется, т.к. ++код вызывает UpdateInstallPlugins, в нём идёт перестроение списка плагинов и обновление интерфейса.
			// Просто отдаём менеджеру плагинов ответ.
			// TODO: отслеживать возможные ошибки при +/- плагинов:
			// из ++кода отправлять статус операции и на основе его отправлять в менеджер плагинов корректный ответ.
			// ничего не изменять в менеджере плагинов, если guid пуст

			let result = window["AscDesktopEditor"]["PluginUninstall"](guid, backup);
						
			return {
				"type" : "Removed",
				"guid" : result ? guid : "",
				"backup" : backup
			};
		}
		
		if (removedPlugin)
		{
			let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");
			if (!currentRemovedPlugins)
				currentRemovedPlugins = {};
			currentRemovedPlugins[removedPlugin.guid] = removedPlugin.serialize();
			setLocalStorageItem("asc_plugins_removed", currentRemovedPlugins);

			let currentInstalledPlugins = getLocalStorageItem("asc_plugins_installed");
			if (currentInstalledPlugins && currentInstalledPlugins[removedPlugin.guid])
			{
				delete currentInstalledPlugins[removedPlugin.guid];
				setLocalStorageItem("asc_plugins_installed", currentInstalledPlugins);
			}

			this.disableCheckInstalledPlugins = true;

			this.sendEvent("asc_onPluginsReset");
			window.g_asc_plugins.updateInterface();

			delete this.disableCheckInstalledPlugins;
		}

		return {
			"type" : "Removed",
			"guid" : removedPlugin ? removedPlugin.guid : ""
		};
	};
	/**
    * Installs a plugin using the specified plugin config.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {object} [config] - The plugin {@link https://api.onlyoffice.com/docs/plugin-and-macros/structure/configuration/ config}.
     * @alias InstallPlugin
     * @returns {object} - An object with the result information.
     * @since 7.2.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/InstallPlugin.js
	 */
	Api.prototype["pluginMethod_InstallPlugin"] = function(config)
	{
		return installPlugin(config, "Installed");
	};
	/**
    * Updates a plugin using the specified plugin config.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {object} [config] - The plugin {@link https://api.onlyoffice.com/docs/plugin-and-macros/structure/configuration/ config}.
     * @alias UpdatePlugin
     * @returns {object} - An object with the result information.
     * @since 7.3.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/UpdatePlugin.js
	 */
	Api.prototype["pluginMethod_UpdatePlugin"] = function(config)
	{
		return installPlugin(config, "Updated");
	};

	Api.prototype["installDeveloperPlugin"] = function(configUrl)
	{
		try
		{
			var xhrObj = new XMLHttpRequest();
			if ( xhrObj )
			{
				xhrObj.open('GET', configUrl, false);
				xhrObj.send('');

				var configJson = JSON.parse(xhrObj.responseText);
				configJson["baseUrl"] = configUrl.substr(0, configUrl.lastIndexOf("/") + 1);

				installPlugin(configJson, "Installed");
				return true;
			}
		}
		catch (e) {}
		return false;
	};

	/**
	* Shows or hides buttons in the header.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} id - The button ID.
	 * @param {boolean} bShow - The flag specifies whether the button is shown (**true**) or hidden (**false**).
	 * @param {string} align - The parameter indicates whether the button will be displayed on the right side of the window or on the left. The default value is "left".
	 * @alias ShowButton 
	 * @since 7.2.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ShowButton.js
	 */
	Api.prototype["pluginMethod_ShowButton"] = function(id, bShow, align)
	{
		if (bShow) {
			this.sendEvent("asc_onPluginShowButton", id, (align === 'right'));
		} else {
			this.sendEvent("asc_onPluginHideButton", id);
		}
	};

	Api.prototype["pluginMethod_GetKeychainStorageInfo"] = function(keys)
	{
		if (!this.keychainStorage)
			this.keychainStorage = new AscCrypto.Storage.CStorageLocalStorage();

		window.g_asc_plugins.setPluginMethodReturnAsync();
		this.keychainStorage.command(keys, function(retObj){
			window.g_asc_plugins.onPluginMethodReturn(retObj);
		});
	};

	Api.prototype["pluginMethod_SetKeychainStorageInfo"] = function(items)
	{
		window.g_asc_plugins.setPluginMethodReturnAsync();

		this.keychainStorage.command(items, function(retObj) {
			window.g_asc_plugins.onPluginMethodReturn(retObj);
		});
	};

	Api.prototype["pluginMethod_OnSignWithKeychain"] = function(data)
	{
	};

	/**
	 * Implements the external drag&drop emulation.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {object} obj The drag&drop emulation properties.
     * @param {string} obj.type - The drag&drop event type:
     * <b>onbeforedrop</b> - an event that is fired when the selected text or element is dragged;
     * <b>ondrop</b> - an event that is fired when the selected text or element is dropped on a valid drop target.
     * @param {number} obj.x - The horizontal coordinate (in pixels) at which the mouse was clicked, relative to the left edge of the entire document.
     * @param {number} obj.y - The vertical coordinate (in pixels) at which the mouse was clicked, relative to the top edge of the entire document.
     * @param {string} obj.html - The dragged HTML element.
     * @param {string} obj.text - The dragged text.
	 * @alias OnDropEvent
	 * @since 7.3.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/OnDropEvent.js
	 */
	Api.prototype["pluginMethod_OnDropEvent"] = function(obj)
	{
		if (!obj || !obj["type"])
			return;

		var e = {
			pageX : obj["x"],
			pageY : obj["y"]
		};

		switch (obj.type)
		{
			case "onbeforedrop":
			{
				this.beginInlineDropTarget(e);
				break;
			}
			case "ondrop":
			{
				this.endInlineDropTarget(e);

				if (obj["html"])
					this["pluginMethod_PasteHtml"](obj["html"]);
				else if (obj["text"])
					this["pluginMethod_PasteText"](obj["text"]);

				break;
			}
			default:
				break;
		}
	};

    /**
     * Returns the document language.
     * @memberof Api
     * @typeofeditors ["CDE", "CPE"]
     * @alias GetDocumentLang
     * @returns {string} - Document language.
	 * @since 7.4.0
     * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetDocumentLang.js
	 */
    Api.prototype["pluginMethod_GetDocumentLang"] = function()
    {
        let langCode = 1033; // en-US
        let langName = "en-US";

        if (this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.GetDefaultLanguage)
            langCode = this.WordControl.m_oLogicDocument.GetDefaultLanguage();

        if (window["Common"])
            langName = window["Common"]["util"]['LanguageInfo']['getLocalLanguageName'](langCode)[0];

        return langName;
    };

	function correctItemIcons(item, baseUrl)
	{
		if (item && item["icons"])
		{
			if ((0 === item["icons"].indexOf("http://")) ||
				(0 === item["icons"].indexOf("https://")) ||
				(0 === item["icons"].indexOf("file://")) ||
				(0 === item["icons"].indexOf("www.")))
			{
				// nothing
			}
			else if (0 === item["icons"].indexOf("external://"))
			{
				item["icons"] = item["icons"].substr("external://".length);
			}
			else
			{
				item["icons"] = baseUrl + item["icons"];
			}
		}
	}

	function correctItemsWithData(items, baseUrl)
	{
		for (let i = 0, itemsLen = items.length; i < itemsLen; i++)
		{
			if (undefined !== items[i]["id"] && undefined !== items[i]["data"])
				items[i]["id"] = items[i]["id"] + "_oo_sep_" + items[i]["data"];

			correctItemIcons(items[i], baseUrl);

			if (items[i]["items"])
				correctItemsWithData(items[i]["items"], baseUrl);
		}
	};

	/**
	 * @typedef {Object} ContextMenuItem
	 * The context menu item.
	 * @property {string} id - The item ID.
	 * @property {string} text - The item text.
	 * @property {string} [data] - The item data (this data will be sent to the click event callback).
	 * @property {boolean} [disabled] - Specifies if the current item is disabled or not.
	 * @property {string} [icons] - The item icons (see the plugins {@link /plugin/config config} documentation).
	 * @property {ContextMenuItem[]} items - An array containing the context menu items for the current item.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ContextMenuItem.js
	 */

	/**
	 * Adds an item to the context menu.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias AddContextMenuItem
	 * @param {ContextMenuItem[]} items - An array containing the context menu items.
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/AddContextMenuItem.js
	 */
	Api.prototype["pluginMethod_AddContextMenuItem"] = function(items)
	{
		let baseUrl = this.pluginsManager.pluginsMap[items["guid"]].baseUrl;
		if (items["items"]) correctItemsWithData(items["items"], baseUrl);
		this.onPluginAddContextMenuItem(items);
	};
	/**
	 * Updates an item in the context menu with the specified items.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias UpdateContextMenuItem
	 * @param {ContextMenuItem[]} items - An array containing the context menu items for the current item.
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/UpdateContextMenuItem.js
	 */
	Api.prototype["pluginMethod_UpdateContextMenuItem"] = function(items)
	{
		let baseUrl = this.pluginsManager.pluginsMap[items["guid"]].baseUrl;
		if (items["items"]) correctItemsWithData(items["items"], baseUrl);
		this.onPluginUpdateContextMenuItem([items]);
	};
	
	/**
	 * Adds a button to the specified content controls track.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlButtons
	 * @param buttons
	 * @since 9.0.0
	 */
	Api.prototype["pluginMethod_AddContentControlButtons"] = function(buttons)
	{
		if (AscCommon.c_oEditorId.Word !== this.getEditorId() || !buttons)
			return;
		
		buttons["baseUrl"] = this.pluginsManager.pluginsMap[buttons["guid"]].baseUrl;
		this.WordControl.m_oLogicDocument.DrawingDocument.contentControls.addPluginButtons(buttons);
	};
	
	/**
	 * The possible values of the base which the relative vertical position of the toolbar menu item will be calculated from.
	 * @typedef {("button" | "...")} ToolbarMenuItemType
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ToolbarMenuItemType.js
	 */

	/**
	 * @typedef {Object} ToolbarMenuItem
	 * The toolbar menu item.
	 * @property {string} id - The item ID.
	 * @property {ToolbarMenuItemType} type - The item type.
	 * @property {string} text - The item text.
	 * @property {string} hint - The item hint.
	 * @property {string} [icons] - The item icons (see the plugins {@link /plugin/config config} documentation).
	 * @property {boolean} [disabled] - Specifies if the current item is disabled or not.
	 * @property {boolean} [enableToggle] - Specifies if an item toggle is enabled or not.
	 * @property {boolean} [lockInViewMode] - Specifies if the current item is locked in the view mode or not.
	 * @property {boolean} [separator] - Specifies if a separator is used between the toolbar menu items or not.
	 * @property {boolean} [split] - Specifies if the toolbar menu items are split or not.
	 * @property {ContextMenuItem[]} [items] - An array containing the context menu items for the current item.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ToolbarMenuItem.js
	 */

	/**
	 * @typedef {Object} ToolbarMenuTab
	 * The toolbar menu tab.
	 * @property {string} id - The tab ID.
	 * @property {string} text - The tab text.
	 * @property {ToolbarMenuItem[]} [items] - An array containing the toolbar menu items for the current tab.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ToolbarMenuTab.js
	 */

	/**
	 * @typedef {Object} ToolbarMenuMainItem
	 * The main toolbar menu item.
	 * @property {string} guid - The plugin guid.
	 * @property {ToolbarMenuTab[]} tabs - An array containing the toolbar menu tabs for the current item.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ToolbarMenuMainItem.js
	 */

	function correctToolbarItems(api, items)
	{
		let baseUrl = api.pluginsManager.pluginsMap[items["guid"]].baseUrl;
		for (let i = 0, len = items["tabs"].length; i < len; i++)
		{
			if (items["tabs"][i]["items"])
				correctItemsWithData(items["tabs"][i]["items"], baseUrl);
		}
	}

	/**
	 * Adds an item to the toolbar menu.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias AddToolbarMenuItem
	 * @param {ToolbarMenuMainItem[]} items - An array containing the main toolbar menu items.
	 * @since 8.1.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/AddToolbarMenuItem.js
	 */
	Api.prototype["pluginMethod_AddToolbarMenuItem"] = function(items)
	{
		correctToolbarItems(this, items);
		this.sendEvent("onPluginToolbarMenu", [items]);
	};

	/**
	 * Updates the toolbar menu item.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias UpdateToolbarMenuItem
	 * @param {ToolbarMenuMainItem[]} items - An array containing the main toolbar menu items.
	 * @since 8.1.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/UpdateToolbarMenuItem.js
	 */
	Api.prototype["pluginMethod_UpdateToolbarMenuItem"] = function(items)
	{
		correctToolbarItems(this, items);
		this.sendEvent("onPluginUpdateToolbarMenu", [items]);
	};

	/**
	 * Shows the plugin modal window.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} frameId - The frame ID.
	 * @param {variation} variation - The plugin variation.
	 * @alias ShowWindow 
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ShowWindow.js
	 */
	Api.prototype["pluginMethod_ShowWindow"] = function(frameId, variation)
	{
		let guid = window.g_asc_plugins.getCurrentPluginGuid();
		variation["guid"] = guid;

		let baseUrl = this.pluginsManager.pluginsMap[guid].baseUrl;
		correctItemIcons(variation["icons"], baseUrl);

		this.sendEvent("asc_onPluginWindowShow", frameId, variation);
	};

	/**
	 * Activates (moves forward) the plugin window/panel.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} frameId - The frame ID.
	 * @alias ActivateWindow
	 * @since 8.1.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ActivateWindow.js
	 */
	Api.prototype["pluginMethod_ActivateWindow"] = function(frameId)
	{
		this.sendEvent("asc_onPluginWindowActivate", frameId);
	};

	/**
	 * Closes the plugin modal window.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} frameId - The frame ID.
	 * @alias CloseWindow
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/CloseWindow.js
	 */
	Api.prototype["pluginMethod_CloseWindow"] = function(frameId)
	{
		this.sendEvent("asc_onPluginWindowClose", frameId);
	};

	/**
	 * Sends a message to the plugin modal window.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} windowID - The frame ID.
	 * @param {string} name - The event name.
	 * @param {object} data - The event data.
	 * @alias SendToWindow
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/SendToWindow.js
	 */
	Api.prototype["pluginMethod_SendToWindow"] = function(windowID, name, data)
	{
		window.g_asc_plugins.onPluginEventWindow(windowID, name, data);
	};

	/**
	 * Resizes the plugin modal window.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} frameId - The frame ID.
	 * @param {number} size - The frame size.
	 * @param {number} minSize - The frame minimum size.
	 * @param {number} maxSize - The frame maximum size.
	 * @alias ResizeWindow
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ResizeWindow.js
	 */
	Api.prototype["pluginMethod_ResizeWindow"] = function(frameId, size, minSize, maxSize)
	{
		window.g_asc_plugins.setPluginMethodReturnAsync();
		this.sendEvent("asc_onPluginWindowResize", frameId, size, minSize, maxSize, function(){
			window.g_asc_plugins.onPluginMethodReturn("resize_result");
		});
	};

	/**
	 * Sends an event to the plugin when the mouse button is released inside the plugin iframe.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} frameId - The frame ID.
	 * @param {number} x - The X coordinate.
	 * @param {number} y - The Y coordinate.
	 * @alias MouseUpWindow
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/MouseUpWindow.js
	 */
	Api.prototype["pluginMethod_MouseUpWindow"] = function(frameId, x, y)
	{
		this.sendEvent("asc_onPluginWindowMouseUp", frameId, x, y);
	};

	/**
	 * Sends an event to the plugin when the mouse button is moved inside the plugin iframe.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} frameId - The frame ID.
 	 * @param {number} x - The X coordinate.
	 * @param {number} y - The Y coordinate.
	 * @alias MouseMoveWindow
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/MouseMoveWindow.js
	 */
	Api.prototype["pluginMethod_MouseMoveWindow"] = function(frameId, x, y)
	{
		this.sendEvent("asc_onPluginWindowMouseMove", frameId, x, y);
	};

	/**
	 * Shows an error/warning message.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE", "PDF"]
	 * @param {string} error - The error text.
	 * @param {number} level - -1 or 0 for error or warning.
	 * @alias ShowError
	 * @since 8.3.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ShowError.js
	 */
	Api.prototype["pluginMethod_ShowError"] = function(error, level)
	{
		this.sendEvent("asc_onError", error, level);
	};

	/**
	 * Callback from dockChangedEvents.
	 * @undocumented
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE", "PDF"]
	 * @param {string} windowID - The frame ID.
	 * @alias OnWindowDockChangedCallback
	 * @since 8.2.2
	 */
	Api.prototype["pluginMethod_OnWindowDockChangedCallback"] = function(windowID)
	{
		let key = window.g_asc_plugins.getCurrentPluginGuid() + "_" + windowID;
		if (window.g_asc_plugins.dockCallbacks[key])
		{
			window.g_asc_plugins.dockCallbacks[key]();
			delete window.g_asc_plugins.dockCallbacks[key];
		}
	};

	/**
	 * Catch AI event from plugin.
	 * @memberof Api
	 * @undocumented
	 * @typeofeditors ["CDE", "CSE", "CPE", "PDF"]
	 * @alias onAIRequest
	 * @param {object} data - Data.
	 * @since 9.0.0
	 */
	Api.prototype["pluginMethod_onAIRequest"] = function(data)
	{
		let curItem = this.aiResolvers[0];
		this.aiResolvers.shift();

		if (this.aiResolvers.length > 0)
			this._AI();

		curItem.resolve(data);
	};

	/**
	 * Returns the local path to the image.
	 * @memberof Api
	 * @undocumented
	 * @typeofeditors ["CDE", "CSE", "CPE", "PDF"]
	 * @alias getLocalImagePath
	 * @param {object} data - Data.
	 * @since 9.0.0
	 */
	Api.prototype["pluginMethod_getLocalImagePath"] = function(url)
	{
		window.g_asc_plugins.setPluginMethodReturnAsync();
		AscCommon.sendImgUrls(this, [url], function(data) {
			let ret = {
				"error" : true,
				"url" : "",
				"path" : ""
			};

			if (data[0] && data[0].path != null && data[0].url !== "error")
			{
				ret["error"] = false;
				ret["url"] = AscCommon.g_oDocumentUrls.imagePath2Local(data[0].path);
				ret["path"] = data[0].path;
			}

			window.g_asc_plugins.onPluginMethodReturn(ret);
		});
	};

})(window);


