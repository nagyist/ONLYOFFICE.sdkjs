function brackets (test) {
	test(
		"(a)[b]\\{c\\}|d|\\|e\\|\\langlef\\rangle\\lfloorg\\rfloor\\lceilh\\rceil\\ulcorneri\\urcorner",
		{
			"type": "LaTeXEquation",
			"body": [
			  [
				{
				  "type": "BracketBlock",
				  "left": "(",
				  "right": ")",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "a"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "[",
				  "right": "]",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "b"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "\\{",
				  "right": "\\}",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "c"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "|",
				  "right": "|",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "d"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "‖",
				  "right": "‖",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "e"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "⟨",
				  "right": "⟩",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "f"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "⌊",
				  "right": "⌋",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "g"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "⌈",
				  "right": "⌉",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "h"
					  }
					]
				  ]
				},
				{
				  "type": "BracketBlock",
				  "left": "┌",
				  "right": "┐",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "i"
					  }
					]
				  ]
				}
			  ],
			]
		  },
		"(a)[b]\\{c\\}|d|\\|e\\|\\langlef\\rangle\\lfloorg\\rfloor\\lceilh\\rceil\\ulcorneri\\urcorner"
	);
	test(
		"(2+1]",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "BracketBlock",
				"left": "(",
				"right": "]",
				"value": [
				  [
					[
					  {
						"type": "NumberLiteral",
						"value": "2"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "NumberLiteral",
						"value": "1"
					  }
					]
				  ]
				]
			  }
			]
		  },
		"(2+1]"
	);
	//TODO doesn't support \backslash bracket
	// test(
	// 	"\\{2+1\\backslash",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": {
	// 			"type": "BracketBlock",
	// 			"left": "\\{",
	// 			"right": "\\",
	// 			"value": [
	// 				{
	// 					"type": "NumberLiteral",
	// 					"value": "2"
	// 				},
	// 				{
	// 					"type": "OperatorLiteral",
	// 					"value": "+"
	// 				},
	// 				{
	// 					"type": "NumberLiteral",
	// 					"value": "1"
	// 				}
	// 			]
	// 		}
	// 	},
	// 	"Check \\{2+1\\backslash"
	// );
	test(
		"\\left.1+2\\right)",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "BracketBlock",
				"left": ".",
				"right": ")",
				"value": [
				  [
					[
					  {
						"type": "NumberLiteral",
						"value": "1"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "NumberLiteral",
						"value": "2"
					  }
					]
				  ]
				]
			  }
			]
		  },
		"\\left.1+2\\right)"
	);
	test(
		"|2|+\\{1\\}+|2|",
		{
			"type": "LaTeXEquation",
			"body": [
			  [
				{
				  "type": "BracketBlock",
				  "left": "|",
				  "right": "|",
				  "value": [
					[
					  {
						"type": "NumberLiteral",
						"value": "2"
					  }
					]
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "BracketBlock",
				  "left": "\\{",
				  "right": "\\}",
				  "value": [
					[
					  {
						"type": "NumberLiteral",
						"value": "1"
					  }
					]
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "BracketBlock",
				  "left": "|",
				  "right": "|",
				  "value": [
					[
					  {
						"type": "NumberLiteral",
						"value": "2"
					  }
					]
				  ]
				}
			  ]
			]
		  },
		"|2|+\\{1\\}+|2|"
	);
}
window["AscMath"].brackets = brackets;
