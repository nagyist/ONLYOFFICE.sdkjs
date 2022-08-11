function specialTest(test) {
	test(
		`2⁰¹²³⁴⁵⁶⁷⁸⁹`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": [
				  {
					"type": "specialScriptLiteral",
					"value": "0123456789"
				  }
				]
			  }
			]
		  },
		"2⁰¹²³⁴⁵⁶⁷⁸⁹"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": [
				  {
					"type": "specialScriptLiteral",
					"value": "4"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "in"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "5"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "6"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "7"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "8"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "9"
				  }
				]
			  }
			]
		  },
		"2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "4"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "in"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "5"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "6"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "7"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "8"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "9"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "45"
				}
			  ]
			]
		  },
		"2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45"
	);
	test(
		`x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "x"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "4"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "in"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "5"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "6"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "7"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "8"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "9"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "45"
				}
			  ]
			]
		  },
		"x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45"
	);

	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎56`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": [
					{
					  "type": "specialScriptLiteral",
					  "value": "234"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "0"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					}
				  ]
				},
				{
				  "type": "NumberLiteral",
				  "value": "56"
				}
			  ]
			]
		  },
		"2₂₃₄₊₍₆₇₋₀₌₆₇₎56"
	);
	test(
		`z₂₃₄₊₍₆₇₋₀₌₆₇₎56`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "z"
				  },
				  "down": [
					{
					  "type": "specialScriptLiteral",
					  "value": "234"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "0"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					}
				  ]
				},
				{
				  "type": "NumberLiteral",
				  "value": "56"
				}
			  ]
			]
		  },
		"z₂₃₄₊₍₆₇₋₀₌₆₇₎56"
	);

	test(
		`2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": [
				  {
					"type": "specialScriptLiteral",
					"value": "234"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "0"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  }
				],
				"up": [
				  {
					"type": "specialScriptLiteral",
					"value": "0123456789"
				  }
				]
			  }
			]
		  },
		"2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": [
				  {
					"type": "specialScriptLiteral",
					"value": "234"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "0"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  }
				],
				"up": [
				  {
					"type": "specialScriptLiteral",
					"value": "4"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "in"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "5"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "6"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "7"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "8"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "9"
				  }
				]
			  }
			]
		  },
		"2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": [
					{
					  "type": "specialScriptLiteral",
					  "value": "234"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "0"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					}
				  ],
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "4"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "in"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "5"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "6"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "7"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "8"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "9"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "45"
				}
			  ]
			]
		  },
		"2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45"
	);
	test(
		`x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "x"
				  },
				  "down": [
					{
					  "type": "specialScriptLiteral",
					  "value": "234"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "0"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					}
				  ],
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "4"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "in"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "5"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "6"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "7"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "8"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "9"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "45"
				}
			  ]
			]
		  },
		"x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45"
	);

	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": [
				  {
					"type": "specialScriptLiteral",
					"value": "234"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "0"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  }
				],
				"up": [
				  {
					"type": "specialScriptLiteral",
					"value": "0123456789"
				  }
				]
			  }
			]
		  },
		"2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹"
	);
	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": [
				  {
					"type": "specialScriptLiteral",
					"value": "234"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "0"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "67"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  }
				],
				"up": [
				  {
					"type": "specialScriptLiteral",
					"value": "4"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "in"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "("
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "5"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "-"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "6"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "+"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "7"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": "="
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "8"
				  },
				  {
					"type": "specialScriptBracketLiteral",
					"value": ")"
				  },
				  {
					"type": "specialScriptLiteral",
					"value": "9"
				  }
				]
			  }
			]
		  },
		"2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹"
	);
	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": [
					{
					  "type": "specialScriptLiteral",
					  "value": "234"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "0"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					}
				  ],
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "4"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "in"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "5"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "6"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "7"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "8"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "9"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "45"
				}
			  ]
			]
		  },
		"2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45"
	);
	test(
		`x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "x"
				  },
				  "down": [
					{
					  "type": "specialScriptLiteral",
					  "value": "234"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "0"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "67"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					}
				  ],
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "4"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "in"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "("
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "5"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "-"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "6"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "+"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "7"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": "="
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "8"
					},
					{
					  "type": "specialScriptBracketLiteral",
					  "value": ")"
					},
					{
					  "type": "specialScriptLiteral",
					  "value": "9"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "45"
				}
			  ]
			]
		  },
		"x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45"
	);
}
window["AscMath"].special = specialTest;
