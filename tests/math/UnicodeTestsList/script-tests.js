function scriptTests(test) {
	test(
		`2^2 + 2`,
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
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  ]
			]
		  },
		"2^2 + 2"
	);
	test(
		`x^2+2`,
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
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
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
		"x^2+2"
	);
	test(
		`x^(256+34)*y`,
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
				  "up": {
					"type": "BracketBlock",
					"value": [
					  [
						{
						  "type": "NumberLiteral",
						  "value": "256"
						},
						{
						  "type": "OperatorLiteral",
						  "value": "+"
						},
						{
						  "type": "NumberLiteral",
						  "value": "34"
						}
					  ]
					],
					"left": "(",
					"right": ")"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "*"
				},
				{
				  "type": "CharLiteral",
				  "value": "y"
				}
			  ]
			]
		  },
		"x^(256+34)*y"
	);
	test(
		`(x+34)^(256+34)-y/x`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
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
						  "type": "NumberLiteral",
						  "value": "34"
						}
					  ]
					],
					"left": "(",
					"right": ")"
				  },
				  "up": {
					"type": "BracketBlock",
					"value": [
					  [
						{
						  "type": "NumberLiteral",
						  "value": "256"
						},
						{
						  "type": "OperatorLiteral",
						  "value": "+"
						},
						{
						  "type": "NumberLiteral",
						  "value": "34"
						}
					  ]
					],
					"left": "(",
					"right": ")"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "FractionLiteral",
				  "up": {
					"type": "CharLiteral",
					"value": "y"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "x"
				  }
				}
			  ]
			]
		  },
		"(x+34)^(256+34)-y/x"
	);
	test(
		`𝛿_(𝜇 + 𝜈)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "OtherLiteral",
				  "value": "𝛿"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "OtherLiteral",
						"value": "𝜇"
					  },
					  {
						"type": "SpaceLiteral",
						"value": " "
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "SpaceLiteral",
						"value": " "
					  },
					  {
						"type": "OtherLiteral",
						"value": "𝜈"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"𝛿_(𝜇 + 𝜈)"
	);
	test(
		`a_b_c`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "b"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "c"
				  }
				}
			  }
			]
		  },
		"a_b_c"
	);
	test(
		`1_2_3`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "1"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "3"
				  }
				}
			  }
			]
		  },
		"1_2_3"
	);

	test(
		`A^5b^i`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "A"
				},
				"up": {
				  "type": "SubSupLiteral",
				  "value": [
					{
					  "type": "NumberLiteral",
					  "value": "5"
					},
					{
					  "type": "CharLiteral",
					  "value": "b"
					}
				  ],
				  "up": {
					"type": "CharLiteral",
					"value": "i"
				  }
				}
			  }
			]
		  },
		"A^5b^i"
	);
	test(
		`a_b_c^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "b"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "c"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"a_b_c^2"
	);

	test(
		`a_b_c^2^2^2^2^2^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "b"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "c"
				  },
				  "up": {
					"type": "SubSupLiteral",
					"value": {
					  "type": "NumberLiteral",
					  "value": "2"
					},
					"up": {
					  "type": "SubSupLiteral",
					  "value": {
						"type": "NumberLiteral",
						"value": "2"
					  },
					  "up": {
						"type": "SubSupLiteral",
						"value": {
						  "type": "NumberLiteral",
						  "value": "2"
						},
						"up": {
						  "type": "SubSupLiteral",
						  "value": {
							"type": "NumberLiteral",
							"value": "2"
						  },
						  "up": {
							"type": "SubSupLiteral",
							"value": {
							  "type": "NumberLiteral",
							  "value": "2"
							},
							"up": {
							  "type": "NumberLiteral",
							  "value": "2"
							}
						  }
						}
					  }
					}
				  }
				}
			  }
			]
		  },
		"a_b_c^2^2^2^2^2^2"
	);

	test(
		`1_2_3^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "1"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "3"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"1_2_3^2"
	);

	test(
		`a_(b_c)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "SubSupLiteral",
					  "value": {
						"type": "CharLiteral",
						"value": "b"
					  },
					  "down": {
						"type": "CharLiteral",
						"value": "c"
					  }
					}
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"a_(b_c)"
	);

	test(
		`a^b_c`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"down": {
				  "type": "CharLiteral",
				  "value": "c"
				},
				"up": {
				  "type": "CharLiteral",
				  "value": "b"
				}
			  }
			]
		  },
		"a^b_c"
	);
	test(
		`sin^2 x`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "FunctionLiteral",
				  "value": "sin"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"third": {
				  "type": "CharLiteral",
				  "value": "x"
				}
			  }
			]
		  },
		"sin^2 x"
	);
	test(
		`𝑊^3𝛽_𝛿_1𝜌_1𝜎_2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "OtherLiteral",
				  "value": "𝑊"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "OtherLiteral",
					"value": "𝛿"
				  },
				  "down": {
					"type": "SubSupLiteral",
					"value": [
					  {
						"type": "NumberLiteral",
						"value": "1"
					  },
					  {
						"type": "OtherLiteral",
						"value": "𝜌"
					  }
					],
					"down": {
					  "type": "SubSupLiteral",
					  "value": [
						{
						  "type": "NumberLiteral",
						  "value": "1"
						},
						{
						  "type": "OtherLiteral",
						  "value": "𝜎"
						}
					  ],
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  }
				},
				"up": [
				  {
					"type": "NumberLiteral",
					"value": "3"
				  },
				  {
					"type": "OtherLiteral",
					"value": "𝛽"
				  }
				]
			  }
			]
		  },
		"𝑊^3𝛽_𝛿_1𝜌_1𝜎_2"
	);
	test(
		`(_23^4)45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "PreScriptLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "45"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "23"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "4"
				}
			  }
			]
		  },
		"(_23^4)45"
	);
	test(
		`(_x^y)45`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "PreScriptLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "45"
				},
				"down": {
				  "type": "CharLiteral",
				  "value": "x"
				},
				"up": {
				  "type": "CharLiteral",
				  "value": "y"
				}
			  }
			]
		  },
		"(_x^y)45"
	);
	test(
		`(_x^y)zyu`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "PreScriptLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "zyu"
				},
				"down": {
				  "type": "CharLiteral",
				  "value": "x"
				},
				"up": {
				  "type": "CharLiteral",
				  "value": "y"
				}
			  }
			]
		  },
		"(_x^y)zyu"
	);
	test(
		`(_453^56)zyu`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "PreScriptLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "zyu"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "453"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "56"
				}
			  }
			]
		  },
		"(_453^56)zyu"
	);
	test(
		`(_(453+2)^56)zyu`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "PreScriptLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "zyu"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "NumberLiteral",
						"value": "453"
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
				"up": {
				  "type": "NumberLiteral",
				  "value": "56"
				}
			  }
			]
		  },
		"(_(453+2)^56)zyu"
	);
	test(
		`(_(453+2)^(345432+y+x/z))zyu`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "PreScriptLiteral",
				"value": {
				  "type": "CharLiteral",
				  "value": "zyu"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "NumberLiteral",
						"value": "453"
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
				"up": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "NumberLiteral",
						"value": "345432"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "CharLiteral",
						"value": "y"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "FractionLiteral",
						"up": {
						  "type": "CharLiteral",
						  "value": "x"
						},
						"down": {
						  "type": "CharLiteral",
						  "value": "z"
						}
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"(_(453+2)^(345432+y+x/z))zyu"
	);
}
window["AscMath"].script = scriptTests;
