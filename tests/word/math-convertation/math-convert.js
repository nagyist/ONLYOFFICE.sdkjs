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

$(function () {

	let type = 0;

	const parser 	= 		window.AscMath.ConvertLaTeXToTokensList;
	const unicodeParser = 	window.AscMath.CUnicodeConverter;

	const accent 	= 		window.AscMath.accents;
	const fraction 	= 		window.AscMath.fraction;
	const degree 	= 		window.AscMath.degree;
	const brackets 	= 		window.AscMath.brackets;
	const numerFunc = 		window.AscMath.numericFunctions;
	const sqrt 		=		window.AscMath.sqrt;
	const style 	= 		window.AscMath.style;


	const UAccents = window.AscMath.script;

	const UnicodeBase = window.AscMath.UnicodeBase;

	function GetParser()
	{
		return (type === 1) ? parser : unicodeParser;
	}

    let Root,
		MathContent = new ParaMath(),
		LogicDocument = AscTest.CreateLogicDocument(),
		p1 = new AscWord.CParagraph(editor.WordControl);

	LogicDocument.RemoveFromContent(0, LogicDocument.GetElementsCount(), false);
	LogicDocument.AddToContent(0, p1);

	if (p1.Content.length > 0)
		p1.Content.splice(0, 1);

	p1.AddToContent(0, MathContent);
	Root = MathContent.Root;

    // function Clear() {
    //     Root.Remove_FromContent(0, Root.Content.length);
    //     Root.Correct_Content();
    // }

    function AddText(str)
	{
		let one = str.getUnicodeIterator();

		while (one.isInside()) {
			let oElement = new AscWord.CRunText(one.value());
			MathContent.Add(oElement);
			one.next();
		}
    };
	function SetMathType(intType)
	{
		LogicDocument.SetMathInputType(intType);
		type = intType;
	};
	function test(program, expected, description = "Без описания")
	{
		QUnit.test(description, function (assert)
		{
			let localParser = GetParser();
			const ast = localParser(program, Root, true);
			console.log(ast);
			assert.deepEqual(ast, expected, description);
		});
	}

	SetMathType(0);
	QUnit.module("Check Unicode accents");
	UAccents(test).bind(this);
	UnicodeBase(test).bind(this);

	// SetMathType(1);
	// QUnit.module("Check LaTeX accents");
	// accent(test).bind(this);

 })
