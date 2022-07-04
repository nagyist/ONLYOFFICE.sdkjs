import "./node.js";
import "../NamesOfLiterals.js";
import "../UnicodeParser.js";
import "./UnicodeTestsList/sqrt-tests.js";
import "./UnicodeTestsList/box-tests.js";
import "./UnicodeTestsList/brackets-test.js";
import "./UnicodeTestsList/fraction-test.js";
import "./UnicodeTestsList/literal-tests.js";
import "./UnicodeTestsList/aboveAndBelow-test.js";
import "./UnicodeTestsList/complex-stuff.js";
import "./UnicodeTestsList/hbrack-tests.js";
import "./UnicodeTestsList/script-tests.js";
import "./UnicodeTestsList/special_scripts-tests.js";
import { assert } from "chai";

const parser = window.AscCommonWord.CUnicodeConverter;

const sqrt = window.AscCommonWord.sqrt;
const box = window.AscCommonWord.box;
const bracket = window.AscCommonWord.bracket;
const fraction = window.AscCommonWord.fraction;
const literal = window.AscCommonWord.literal;
const aboveBelow = window.AscCommonWord.aboveBelow;
const complex = window.AscCommonWord.complex;
const hbrack = window.AscCommonWord.hbrack;
const script = window.AscCommonWord.script;
const special = window.AscCommonWord.special;

// const CUnicodeConverter = eval(text)
// const assert = require("chai").assert;
//
// // const arr = [
// // 	"∑",
// // 	"⅀",
// // 	"⨊",
// // 	"∏",
// // 	"∐",
// // 	"⨋",
// // 	"∫",
// // 	"∬",
// // 	"∭",
// // 	"⨌",
// // 	"∮",
// // 	"∯",
// // 	"∰",
// // 	"∱",
// // 	"⨑",
// // 	"∲",
// // 	"∳",
// // 	"⨍",
// // 	"⨎",
// // 	"⨏",
// // 	"⨕",
// // 	"⨖",
// // 	"⨗",
// // 	"⨘",
// // 	"⨙",
// // 	"⨚",
// // 	"⨛",
// // 	"⨜",
// // 	"⨒",
// // 	"⨓",
// // 	"⨔",
// // 	"⋀",
// // 	"⋁",
// // 	"⋂",
// // 	"⋃",
// // 	"⨃",
// // 	"⨄",
// // 	"⨅",
// // 	"⨆",
// // 	"⨀",
// // 	"⨁",
// // 	"⨂",
// // 	"⨉",
// // 	"⫿",
// // 	"(",
// // 	"[",
// // 	"{",
// // 	"〈",
// // 	")",
// // 	"]",
// // 	"}",
// // 	"〉",
// // 	"├",
// // 	"┤",
// // 	"┬",
// // 	"┴",
// // 	"▁",
// // 	"¯",
// // 	"▭",
// // 	"□",
// // 	"&",
// // 	"▒",
// // 	"^",
// // 	"_",
// // 	"¦",
// // 	"√",
// // 	"∛",
// // 	"∜",
// // 	"⊘",
// // 	"/",
// // 	",",
// // 	".",
// // 	"⏜",
// // 	"⏝",
// // 	"⎴",
// // 	"⎵",
// // 	"⏞",
// // 	"⏟",
// // 	"⏠",
// // 	"⏡",
// // 	"&",
// // 	"■",
// // ];
// //
// // describe("Проверка литералов из операторов", function () {
// //   arr.forEach((literal) => li_test("\\" + literal));
// // });
//
describe("Проверка работоспособности простых литералов", function () {
	literal(test);
});
describe("Проверка работоспособности деления", function () {
	fraction(test);
});
describe("Проверка работоспособности радикалов", function () {
	sqrt(test);
});
describe("Проверка работоспособности скриптов", function () {
	script(test);
});
describe("Проверка работоспособности below/above", function () {
	aboveBelow(test);
});
describe("Проверка работоспособности hBrack", function () {
	hbrack(test);
});
describe("Проверка работоспособности скобок", function () {
	bracket(test);
});
describe("Проверка работоспособности комплексных выражений", function () {
	complex(test);
});
describe("Проверка box", function () {
	box(test);
});
describe("Проверка special_scripts", function () {
	special(test);
});

function test(program, expected, description = "Без описания") {
	it(description, function () {
		const ast = parser(program, undefined, true);
		assert.deepEqual(
			ast,
			expected
		);
	});
}