function boxTests(test) {
	test(
		`□(1+2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "BoxLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
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
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"□(1+2)"
	);
	test(
		`□1+2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "BoxLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "1"
				  }
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
		  },
		"□1+2"
	);
	test(
		`□1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "BoxLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "1"
				}
			  }
			]
		  },
		"□1"
	);
	test(
		`□1/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "BoxLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "1"
				  }
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"□1/2"
	);
	test(
		`▭(1+2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "RectLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
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
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"▭(1+2)"
	);
	test(
		`▭1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "RectLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "1"
				}
			  }
			]
		  },
		"▭1"
	);
	test(
		`▭1/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "RectLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "1"
				  }
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"▭1/2"
	);
	test(
		`▁(1+2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "overBarLiteral",
				"overUnder": "▁",
				"value": {
				  "type": "BracketBlock",
				  "value": [
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
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"▁(1+2)"
	);

	test(
		`▁1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "overBarLiteral",
				"overUnder": "▁",
				"value": {
				  "type": "NumberLiteral",
				  "value": "1"
				}
			  }
			]
		  },
		"▁1"
	);
	test(
		`▁1/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "overBarLiteral",
				  "overUnder": "▁",
				  "value": {
					"type": "NumberLiteral",
					"value": "1"
				  }
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"▁1/2"
	);
	test(
		` ̄(1+2)`.trim(),
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "AccentLiteral",
				  "value": "̄"
				},
				{
				  "type": "BracketBlock",
				  "value": [
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
				  ],
				  "left": "(",
				  "right": ")"
				}
			  ]
			]
		  },
		" ̄(1+2)"
	);

	test(
		` ̄1`.trim(),
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "AccentLiteral",
				  "value": "̄"
				},
				{
				  "type": "NumberLiteral",
				  "value": "1"
				}
			  ]
			]
		  },
		" ̄1"
	);
	test(
		`(1+2)̂`.trim(),
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"base": {
				  "type": "BracketBlock",
				  "value": [
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
				  ],
				  "left": "(",
				  "right": ")"
				},
				"value": "̂"
			  }
			]
		  },
		"(1+2)̂"
	);
}

window["AscMath"].box = boxTests;

