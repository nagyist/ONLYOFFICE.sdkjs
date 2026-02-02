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

(function (window)
{
	QUnit.module("ApiWorksheet");
	QUnit.test("GetSelectedShapes", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet()

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		for(let nShape = 0; nShape < 3; nShape++)
		{
			let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
			assert.ok(true, 'Add new ellipse shape');
			if (nShape !== 1) {
				shape.Select();
				assert.ok(true, 'Select added shape');
			}
		}

		let selectedShapes = worksheet.GetSelectedShapes();
		assert.strictEqual(
			selectedShapes.length,
			2,
			"Count of selected shapes is 2"
		);
	});
	
	QUnit.test("GetSelectedDrawings", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet()

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		for(let nShape = 0; nShape < 3; nShape++)
		{
			let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
			assert.ok(true, 'Add new ellipse shape');
			if (nShape !== 1) {
				shape.Select();
				assert.ok(true, 'Select added shape');
			}
		}

		let image = worksheet.AddImage("https://static.onlyoffice.com/assets/docs/samples/img/presentation_sky.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
		assert.ok(true, 'Add new image');
		image.Select();
		assert.ok(true, 'Select added image');

		image = worksheet.AddImage("https://static.onlyoffice.com/assets/docs/samples/img/presentation_sky.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
		assert.ok(true, 'Add new image')

		let selectedShapes = worksheet.GetSelectedDrawings();
		assert.strictEqual(
			selectedShapes.length,
			3,
			"Count of selected shapes is 3 (2 shape, 1 image)"
		);
	});

})(window);
