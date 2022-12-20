function accents(test) {

	test(
		"\\dot{a}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": AscMath.MathLiterals.accent.id,
					"base": {
						"type": AscMath.MathLiterals.char.id,
						"value": "a"
					},
					"value": "̇"
				}
			],
		},
		"Check \\dot{a}"
	);
	test(
		"\\ddot{b}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": AscMath.MathLiterals.accent.id,
					"base": {
						"type": AscMath.MathLiterals.char.id,
						"value": "b"
					},
					"value": "̈"
				}
			],
		},
		"Check \\ddot{b}"
	);
	test(
		"\\acute{c}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "c"
				},
				"value": "́"
				}
			]
		},
		"Check \\acute{c}"
	);
	test(
		"\\grave{d}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": AscMath.MathLiterals.accent.id,
					"base": {
						"type": AscMath.MathLiterals.char.id,
						"value": "d"
					},
					"value": "̀"
				}
			]
		},
		"Check \\grave{d}"
	);
	test(
		"\\check{e}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "e"
				},
				"value": "̌"
				}
			]
		},
		"Check \\check{e}"
	);
	test(
		"\\breve{f}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "f"
				},
				"value": "̆"
				}
			]
		},
		"Check \\breve{f}"
	);
	test(
		"\\tilde{g}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "g"
				},
				"value": "̃"
				}
			],
		},
		"Check \\tilde{g}"
	);
	test(
		"\\bar{h}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "h"
				},
				"value": "̄"
				}
			]
		},
		"Check \\bar{h}"
	);
	test(
		"\\widehat{j}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "j"
				},
				"value": "̂"
				}
			]
		},
		"Check \\widehat{j}"
	);
	test(
		"\\vec{k}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": AscMath.MathLiterals.char.id,
					"value": "k"
				},
				"value": "⃗"
				}
			]
		}
		,
		"Check \\vec{k}"
	);
	//doesn't implement in word
	// test(
	// 	"\\not{l}",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": {
	// 			"type": "AccentLiteral",
	// 			"base": {
	// 				"type": "CharLiteral",
	// 				"value": "l"
	// 			},
	// 			"value": 824
	// 		}
	// 	},
	// 	"Check \\not{l}"
	// );

	// test(
	// 	"\\not\\notl2",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": [
	// 			{
	// 				"type": "AccentLiteral",
	// 				"base": {
	// 					"type": "AccentLiteral",
	// 					"base": {
	// 						"type": "CharLiteral",
	// 						"value": "l"
	// 					},
	// 					"value": 824
	// 				},
	// 				"value": 824
	// 			},
	// 			{
	// 				"type": "NumberLiteral",
	// 				"value": "2"
	// 			}
	// 		]
	// 	},
	// 	"Check \\notl"
	// );

	test(
		"\\vec \\frac{k}{2}",
		{
			"type": "LaTeXEquation",
			"body": [
				{
				"type": AscMath.MathLiterals.accent.id,
				"base": {
					"type": "FractionLiteral",
					"up": {
						"type": AscMath.MathLiterals.char.id,
						"value": "k"
					},
					"down": {
						"type": AscMath.MathLiterals.number.id,
						"value": "2"
					}
				},
				"value": "⃗"
			}
			]
		},
		"Check \\vec \\frac{k}{2}"
	);
	test(
		"5''",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": AscMath.MathLiterals.subSup.id,
					"up": {
						"type": AscMath.MathLiterals.char.id,
						"value": "''"
					},
					"value": {
						"type": AscMath.MathLiterals.number.id,
						"value": "5"
					}
				}
				]
		},
		"Check 5''"
	);
	test(
		"\\frac{4}{5}''",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": AscMath.MathLiterals.subSup.id,
					"up": {
						"type": AscMath.MathLiterals.char.id,
						"value": "''"
					},
					"value": {
						"down": {
							"type": AscMath.MathLiterals.number.id,
							"value": "5"
						},
						"type": "FractionLiteral",
						"up": {
							"type": AscMath.MathLiterals.number.id,
							"value": "4"
						}
					}
				}
			],
		},
		"Check \\frac{4}{5}''"
	);
}
window["AscMath"].accents = accents;
