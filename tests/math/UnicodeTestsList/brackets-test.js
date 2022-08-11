function bracketsTests(test) {
	test(
		`(1+2)+2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
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
		  }
		  ,
		"(1+2)+2"
	);
	test(
		`{1+2}-X`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
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
				  "left": "{",
				  "right": "}"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "CharLiteral",
				  "value": "X"
				}
			  ]
			]
		  },
		"{1+2}-X"
	);
	test(
		`[1+2]*i`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
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
				  "left": "[",
				  "right": "]"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "*"
				},
				{
				  "type": "CharLiteral",
				  "value": "i"
				}
			  ]
			]
		  },
		"[1+2]*i"
	);
	test(
		`|1+2|-89/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
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
				  "left": "|",
				  "right": "|"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "FractionLiteral",
				  "up": {
					"type": "NumberLiteral",
					"value": "89"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  ]
			]
		  },
		"|1+2|-89/2"
	);
	test(
		`√〖89/2〗`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				}
			  }
			]
		  },
		"√〖89/2〗"
	);
	test(
		`〖89/2〗_2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"〖89/2〗_2"
	);
	test(
		`〖89/2〗^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"〖89/2〗^2"
	);
	test(
		`2_〖89/2〗`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				}
			  }
			]
		  },
		"2_〖89/2〗"
	);
	test(
		`2^〖89/2〗`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				}
			  }
			]
		  },
		"2^〖89/2〗"
	);
	test(
		`2_〖89/2〗_2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "BracketBlock",
					"value": [
					  {
						"type": "FractionLiteral",
						"up": {
						  "type": "NumberLiteral",
						  "value": "89"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "2"
						}
					  }
					],
					"left": "〖",
					"right": "〗"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"2_〖89/2〗_2"
	);
	test(
		`2^〖89/2〗^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "BracketBlock",
					"value": [
					  {
						"type": "FractionLiteral",
						"up": {
						  "type": "NumberLiteral",
						  "value": "89"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "2"
						}
					  }
					],
					"left": "〖",
					"right": "〗"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"2^〖89/2〗^2"
	);
	test(
		`2┴〖89/2〗`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				}
			  }
			]
		  },
		"2┴〖89/2〗"
	);
	test(
		`2┴〖89/2〗┴2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": {
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "BracketBlock",
					"value": [
					  {
						"type": "FractionLiteral",
						"up": {
						  "type": "NumberLiteral",
						  "value": "89"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "2"
						}
					  }
					],
					"left": "〖",
					"right": "〗"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"2┴〖89/2〗┴2"
	);
	test(
		`2┬〖89/2〗`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "NumberLiteral",
						"value": "89"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  }
					}
				  ],
				  "left": "〖",
				  "right": "〗"
				}
			  }
			]
		  },
		"2┬〖89/2〗"
	);
	test(
		`2┬〖89/2〗┬2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"down": {
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "BracketBlock",
					"value": [
					  {
						"type": "FractionLiteral",
						"up": {
						  "type": "NumberLiteral",
						  "value": "89"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "2"
						}
					  }
					],
					"left": "〖",
					"right": "〗"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"2┬〖89/2〗┬2"
	);
	// test(
	// 	"├]a+b┤[",
	// 	{
	// 		type: "UnicodeEquation",
	// 		body: {
	// 			type: "expBracketLiteral",
	// 			open: "]",
	// 			close: "[",
	// 			value: [
	// 				{
	// 					type: "CharLiteral",
	// 					value: "a"
	// 				},
	// 				{
	// 					type: "OperatorLiteral",
	// 					value: "+"
	// 				},
	// 				{
	// 					type: "CharLiteral",
	// 					value: "b"
	// 				}
	// 			]
	// 		}
	// 	},
	// 	"Check: ├]a+b┤["
	// )
}

window["AscMath"].bracket = bracketsTests;
