/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

(function (AscCommon)
{
	
	let c_oSerCustomsPPTY = {
		Custom: 8,
		
		ItemId: 0,
		Uri: 1,
		Content : 2,
		
		Uri2 : 0
	};
	
	/**
	 * @param xmlManager
	 * @param stream
	 * @param isCustomA (document - true, spreadsheets - false)
	 * @constructor
	 */
	function BinaryCustomsTableWriter(xmlManager, stream, isCustomA)
	{
		this.memory = stream;
		this.bs = new AscCommon.BinaryCommonWriter(this.memory);
		this.customXmlManager = xmlManager;
		
		let _t = this;
		this.Write = function()
		{
			this.bs.WriteItemWithLength(function(){
				_t.WriteCustomXmls()
			});
		};
		this.WriteCustomXmls = function()
		{
			for (let i = 0, count = this.customXmlManager.getCount(); i < count; ++i) {
				this.bs.WriteItem(c_oSerCustoms.Custom, function() {
					_t.WriteCustomXml(_t.customXmlManager.getCustomXml(i));
				});
			}
		};
		this.WriteCustomXml = function(customXml) {
			let namespaces = customXml.getAllNamespaces();
			for (let i = 0; i < namespaces.length; ++i) {
				this.bs.WriteItem(c_oSerCustoms.Uri, function () {
					_t.memory.WriteString3(namespaces[i]);
				});
			}
			if (null !== customXml.itemId) {
				this.bs.WriteItem(c_oSerCustoms.ItemId, function() {
					_t.memory.WriteString3(customXml.itemId);
				});
			}
			if (null !== customXml.content) {
				this.bs.WriteItem(isCustomA ? c_oSerCustoms.ContentA : c_oSerCustoms.Content, function() {
					let str  = _t.customXmlManager.getCustomXMLString(customXml);
					let data = AscCommon.Utf8.encode(str);
					_t.memory.WriteULong(data.length);
					_t.memory.WriteBuffer(data, 0, data.length);
				});
			}
		};
	}
	
	/**
	 * @param xmlManager
	 * @param stream
	 * @constructor
	 */
	function BinaryCustomsTableReader(xmlManager, stream) {
		this.stream = stream;
		this.customXmlManager = xmlManager;
		this.bcr = new AscCommon.Binary_CommonReader(this.stream);
		
		let _t = this;
		this.Read = function() {
			return this.bcr.ReadTable(function(t, l) {
				return _t.ReadCustom(t, l);
			});
		};
		this.ReadCustom = function(type, length) {
			let res;
			if (c_oSerCustoms.Custom === type) {
				let customXml = new AscWord.CustomXml();
				res = this.bcr.Read1(length, function(t, l) {
					return _t.ReadCustomContent(t, l, customXml);
				});
				this.customXmlManager.add(customXml);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadCustomContent = function(type, length, custom) {
			if (c_oSerCustoms.Uri === type) {
				custom.setNamespaceUri(this.stream.GetString2LE(length));
			} else if (c_oSerCustoms.ItemId === type) {
				custom.itemId = this.stream.GetString2LE(length);
			} else if (c_oSerCustoms.ContentA === type || c_oSerCustoms.Content === type) {
				custom.addContent(this.stream.GetBuffer(length))
			} else {
				return c_oSerConstants.ReadUnknown;
			}
			return c_oSerConstants.ReadOk;
		};
		this.ReadPPTY = function() {
			let s = this.stream;
			let count = s.GetULong();
			for (let i = 0; i < count; ++i) {
				let t = this.stream.GetUChar();
				let l = this.stream.GetULongLE();
				this.ReadCustomPPTY(t, l);
			}
		};
		this.ReadCustomPPTY = function(type, length) {
			let res;
			if (c_oSerCustomsPPTY.Custom === type) {
				let customXml = new AscWord.CustomXml();
				res = this.bcr.Read1(length, function(t, l) {
					return _t.ReadCustomContentPPTY(t, l, customXml);
				});
				this.customXmlManager.add(customXml);
			} else {
				res = c_oSerConstants.ReadUnknown;
			}
			return res;
		};
		this.ReadCustomContentPPTY = function(type, length, customXml) {
			if (c_oSerCustomsPPTY.Uri === type) {
				let count = this.stream.GetULong();
				for (let i = 0; i < count; ++i) {
					let t = this.stream.GetUChar(); // must be 0
					let l = this.stream.GetULongLE();
					
					this.bcr.Read1(l, function(t, l) {
						return _t.ReadCustomUriPPTY(t, l, customXml);
					});
				}
			} else if (c_oSerCustomsPPTY.ItemId === type) {
				customXml.itemId = this.stream.GetString2(length);
			} else if (c_oSerCustomsPPTY.Content === type) {
				let len = this.stream.GetULong();
				customXml.addContent(this.stream.GetBufferUint8(len));
			} else {
				return c_oSerConstants.ReadUnknown;
			}
			return c_oSerConstants.ReadOk;
		};
		this.ReadCustomUriPPTY = function(type, length, customXml) {
			if (type !== c_oSerCustomsPPTY.Uri2) {
				return c_oSerConstants.ReadUnknown;
			}
			this.stream.GetUChar(); // AscCommon.g_nodeAttributeStart
			if (0 === this.stream.GetUChar()) {
				let ns = this.stream.GetString2();
				customXml.setNamespaceUri(ns);
			}
			this.stream.GetUChar(); // AscCommon.g_nodeAttributeEnd
			return c_oSerConstants.ReadOk;
		};
	}
	//--------------------------------------------------------export----------------------------------------------------
	AscCommon.BinaryCustomsTableWriter = BinaryCustomsTableWriter;
	AscCommon.BinaryCustomsTableReader = BinaryCustomsTableReader;
	
})(window['AscCommon']);
