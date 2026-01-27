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

QUnit.config.autostart = false;
(function (window)
{
	const {InitEditor} = AscTestShortcut;

	let editor, wb, wbView, ws, wsView, cellEditor;
	InitEditor(function ()
	{
		editor = window["Asc"]["editor"];
		wb = editor.wbModel;
		wbView = editor.wb;
		ws = wb.aWorksheets[0];
		wsView = wbView.getWorksheet();
		cellEditor = wbView.cellEditor;
		QUnit.start();
	});

	QUnit.module("ChartsDraw");
	QUnit.test("GetSelectedShapes", function (assert) {
		let worksheet = editor.GetActiveSheet()

		const fill = editor.CreateSolidFill(editor.CreateRGBColor(51, 51, 51));
		const stroke = editor.CreateStroke(0, editor.CreateNoFill());

		for(let nShape = 0; nShape < 3; nShape++)
		{
			let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
			if (nShape !== 1)
				shape.Select();
		}

		let selectedShapes = worksheet.GetSelectedShapes();
		assert.strictEqual(
			selectedShapes.length,
			2,
			"Count of selected shapes is 2"
		);
	});

})(window);
