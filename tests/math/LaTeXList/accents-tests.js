function accents(test) {
	test(
		"\\dot{a}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"value": "̇"
			  }
			]
		  },
		"\\dot{a}"
	);
	test(
		"\\ddot{b}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "b"
				},
				"value": "̈"
			  }
			]
		  },
		"\\ddot{b}"
	);
	test(
		"\\acute{c}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "c"
				},
				"value": "́"
			  }
			]
		  },
		"\\acute{c}"
	);
	test(
		"\\grave{d}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "d"
				},
				"value": "̀"
			  }
			]
		  },
		"\\grave{d}"
	);
	test(
		"\\check{e}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "e"
				},
				"value": "̌"
			  }
			]
		  },
		"\\check{e}"
	);
	test(
		"\\breve{f}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "f"
				},
				"value": "̆"
			  }
			]
		  },
		"\\breve{f}"
	);
	test(
		"\\tilde{g}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "g"
				},
				"value": "̃"
			  }
			]
		  },
		"\\tilde{g}"
	);
	test(
		"\\bar{h}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "h"
				},
				"value": "̅"
			  }
			]
		  }
		  ,
		"\\bar{h}"
	);
	test(
		"\\widehat{j}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "j"
				},
				"value": "̂"
			  }
			]
		  },
		"\\widehat{j}"
	);
	test(
		"\\vec{k}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "CharLiteral",
				  "value": "k"
				},
				"value": "⃗"
			  }
			]
		  }
		,
		"\\vec{k}"
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
	test(
		"\\vec \\frac{k}{2}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "FractionLiteral",
				  "up": {
					"type": "CharLiteral",
					"value": "k"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				"value": "⃗"
			  }
			]
		  },
		"\\vec \\frac{k}{2}"
	);
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
		"5''",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "5"
				},
				"up": {
				  "type": "CharLiteral",
				  "value": "''"
				}
			  }
			]
		  },
		"5''"
	);
	test(
		"\\frac{4}{5}''",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "FractionLiteral",
				  "up": {
					"type": "NumberLiteral",
					"value": "4"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "5"
				  }
				},
				"up": {
				  "type": "CharLiteral",
				  "value": "''"
				}
			  }
			]
		  },
		"\\frac{4}{5}''"
	);
}
window["AscMath"].accents = accents;
