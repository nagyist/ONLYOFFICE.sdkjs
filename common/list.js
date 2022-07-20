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

"use strict";

(function(window, undefined){

    window["AscCommon"] = window.AscCommon = (window["AscCommon"] || {});

    function CListNode()
    {
        this.value = null;
        this.prev = null;
        this.next = null;
    }
    CListNode.prototype.remove = function()
    {
        if (this.prev)
            this.prev.next = this.next;
        if (this.next)
            this.next.prev = this.prev;
    };

    function CList()
    {
        this.length = 0;
        this.startNode = null;
        this.endNode = null;
    };

    CList.prototype.push = function(element)
    {
        if (0 === this.length)
        {
            this.startNode = new CListNode();
            this.startNode.value = element;

            this.endNode = this.startNode;
            return;
        }

        let oldEnd = this.endNode;
        this.endNode = new CListNode();
        this.endNode.value = element;
        this.endNode.prev = oldEnd;

        oldEnd.next = this.endNode;
    };

    CList.prototype.getNodeAt = function(index)
    {
        if (0 > index || index >= this.length)
            return null;

        if (this.index === 0)
            return this.startNode;
        if (this.index === this.length - 1)
            return this.endNode;

        let node = this.startNode;
        let i = 0;
        while (i < index)
        {
            i++;
            node = node.next;
        }

        return node;
    };

    CList.prototype.getValue = function(index, isRemove)
    {
        let node = this.getNodeAt(index);
        if (!node)
            return null;
        if (isRemove)
            node.remove();
        this.length--;
        if (0 === this.length)
            this.startNode = this.endNode = null;

        return node.value;
    };

    CList.prototype.find = function(element)
    {
        if (0 === this.length)
            return -1;

        let node = this.startNode;
        let index = 0;
        while (index < this.length)
        {
            if (node.value === element)
                return index;

            index++;
            node = node.next;
        }

        return index;
    };

    window["AscCommon"].CList = CList;

})(window);
