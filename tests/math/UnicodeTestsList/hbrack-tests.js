function hbrackTests(test) {
	test(
		`⏞(x+⋯+x)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⏞",
				"value": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "x"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "⋯"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "CharLiteral",
						"value": "x"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		" ⏞(x+⋯+x)"
	);

	test(
		`⏞(x+⋯+x)^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⏞",
				"value": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "x"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "⋯"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "CharLiteral",
						"value": "x"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		" ⏞(x+⋯+x)^2"
	);
	test(
		`⏞(x+⋯+x)_2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⏞",
				"value": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "x"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "⋯"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "CharLiteral",
						"value": "x"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		" ⏞(x+⋯+x)_2"
	);
}
window["AscMath"].hbrack = hbrackTests;
