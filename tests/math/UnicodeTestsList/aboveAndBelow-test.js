function testsAboveBelow(test) {
	test(
		`base┴2+2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "base"
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
		"base┴2+2"
	);
	test(
		`base┴2┴x+2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "base"
				  },
				  "up": {
					"type": "functionWithLimitLiteral",
					"value": {
					  "type": "NumberLiteral",
					  "value": "2"
					},
					"up": {
					  "type": "CharLiteral",
					  "value": "x"
					}
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
		"base┴2┴x+2"
	);
	test(
		`base┴2┴(x/y+6)+2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "base"
				  },
				  "up": {
					"type": "functionWithLimitLiteral",
					"value": {
					  "type": "NumberLiteral",
					  "value": "2"
					},
					"up": {
					  "type": "BracketBlock",
					  "value": [
						[
						  {
							"type": "FractionLiteral",
							"up": {
							  "type": "CharLiteral",
							  "value": "x"
							},
							"down": {
							  "type": "CharLiteral",
							  "value": "y"
							}
						  },
						  {
							"type": "OperatorLiteral",
							"value": "+"
						  },
						  {
							"type": "NumberLiteral",
							"value": "6"
						  }
						]
					  ],
					  "left": "(",
					  "right": ")"
					}
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
		"base┴2┴(x/y+6)+2"
	);
	test(
		`x^23┴2/y`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "SubSupLiteral",
					"value": {
					  "type": "CharLiteral",
					  "value": "x"
					},
					"up": {
					  "type": "NumberLiteral",
					  "value": "23"
					}
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				"down": {
				  "type": "CharLiteral",
				  "value": "y"
				}
			  }
			]
		  },
		"x^23┴2/y"
	);
	test(
		`(x^23)┴2/y`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "BracketBlock",
					"value": [
					  {
						"type": "SubSupLiteral",
						"value": {
						  "type": "CharLiteral",
						  "value": "x"
						},
						"up": {
						  "type": "NumberLiteral",
						  "value": "23"
						}
					  }
					],
					"left": "(",
					"right": ")"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				"down": {
				  "type": "CharLiteral",
				  "value": "y"
				}
			  }
			]
		  },
		"(x^23)┴2/y"
	);
	test(
		`4┴2+2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "4"
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
		"4┴2+2"
	);
	test(
		`base┴expre*xz`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "base"
				  },
				  "up": {
					"type": "CharLiteral",
					"value": "expre"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "*"
				},
				{
				  "type": "CharLiteral",
				  "value": "xz"
				}
			  ]
			]
		  },
		"base┴expre*xz"
	);
	test(
		`2┴expre-p`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "up": {
					"type": "CharLiteral",
					"value": "expre"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "CharLiteral",
				  "value": "p"
				}
			  ]
			]
		  },
		"2┴expre-p"
	);
	test(
		`base┬2*x`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "base"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "*"
				},
				{
				  "type": "CharLiteral",
				  "value": "x"
				}
			  ]
			]
		  },
		"base┬2*x"
	);
	test(
		`4┬2+x/y`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "4"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
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
					"value": "y"
				  }
				}
			  ]
			]
		  },
		"4┬2+x/y"
	);
	test(
		`base┬expr*x^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "base"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "expr"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "*"
				},
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
				}
			  ]
			]
		  },
		"base┬expr*x^2"
	);
	test(
		`2┬expr-x_i`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "expr"
				  }
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "x"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "i"
				  }
				}
			  ]
			]
		  },
		"2┬expr-x_i"
	);
	test(
		`2┬(expr+2)+(2+1)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": {
					"type": "BracketBlock",
					"value": [
					  [
						{
						  "type": "CharLiteral",
						  "value": "expr"
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
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "BracketBlock",
				  "value": [
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
				  ],
				  "left": "(",
				  "right": ")"
				}
			  ]
			]
		  },
		"2┬(expr+2)+(2+1)"
	);
	test(
		`2┬(expr+2+x/2)^2 - 1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "2"
				  },
				  "down": {
					"type": "SubSupLiteral",
					"value": {
					  "type": "BracketBlock",
					  "value": [
						[
						  {
							"type": "CharLiteral",
							"value": "expr"
						  },
						  {
							"type": "OperatorLiteral",
							"value": "+"
						  },
						  {
							"type": "NumberLiteral",
							"value": "2"
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
							  "type": "NumberLiteral",
							  "value": "2"
							}
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
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "NumberLiteral",
				  "value": "1"
				}
			  ]
			]
		  },
		"2┬(expr+2+x/2)^2 - 1"
	);
	test(
		`(2+x)┬expr`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
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
						"type": "CharLiteral",
						"value": "x"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				},
				"down": {
				  "type": "CharLiteral",
				  "value": "expr"
				}
			  }
			]
		  },
		"(2+x)┬expr"
	);
	test(
		`(2+y)┬(expr+2+x/2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
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
						"type": "CharLiteral",
						"value": "y"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "expr"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "NumberLiteral",
						"value": "2"
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
						  "type": "NumberLiteral",
						  "value": "2"
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
		"(2+y)┬(expr+2+x/2)"
	);
	test(
		`(2+y^2)┬(expr_3+2+x/2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
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
						"type": "SubSupLiteral",
						"value": {
						  "type": "CharLiteral",
						  "value": "y"
						},
						"up": {
						  "type": "NumberLiteral",
						  "value": "2"
						}
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "SubSupLiteral",
						"value": {
						  "type": "CharLiteral",
						  "value": "expr"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "3"
						}
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "NumberLiteral",
						"value": "2"
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
						  "type": "NumberLiteral",
						  "value": "2"
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
		"(2+y^2)┬(expr_3+2+x/2)"
	);
}
window["AscMath"].aboveBelow = testsAboveBelow;
