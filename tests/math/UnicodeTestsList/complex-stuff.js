function complexTest(test) {
	test(
		`(a + b)^n = ∑_(k=0)^n▒(n¦k) a^k b^(n-k),`,
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
						  "value": "a"
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
						  "type": "CharLiteral",
						  "value": "b"
						}
					  ]
					],
					"left": "(",
					"right": ")"
				  },
				  "up": {
					"type": "CharLiteral",
					"value": "n"
				  }
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "OperatorLiteral",
				  "value": "="
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "opNaryLiteral",
					"value": "∑"
				  },
				  "down": {
					"type": "BracketBlock",
					"value": [
					  [
						{
						  "type": "CharLiteral",
						  "value": "k"
						},
						{
						  "type": "OperatorLiteral",
						  "value": "="
						},
						{
						  "type": "NumberLiteral",
						  "value": "0"
						}
					  ]
					],
					"left": "(",
					"right": ")"
				  },
				  "up": {
					"type": "CharLiteral",
					"value": "n"
				  },
				  "third": {
					"type": "BracketBlock",
					"value": [
					  {
						"type": "BinomLiteral",
						"up": {
						  "type": "CharLiteral",
						  "value": "n"
						},
						"down": {
						  "type": "CharLiteral",
						  "value": "k"
						}
					  }
					],
					"left": "(",
					"right": ")"
				  }
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "a"
				  },
				  "up": {
					"type": "CharLiteral",
					"value": "k"
				  }
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "b"
				  },
				  "up": [
					{
					  "type": "BracketBlock",
					  "value": [
						[
						  {
							"type": "CharLiteral",
							"value": "n"
						  },
						  {
							"type": "OperatorLiteral",
							"value": "-"
						  },
						  {
							"type": "CharLiteral",
							"value": "k"
						  }
						]
					  ],
					  "left": "(",
					  "right": ")"
					},
					{
					  "type": "CharLiteral",
					  "value": ","
					}
				  ]
				}
			  ]
			]
		},
		"(a + b)^n = ∑_(k=0)^n▒(n¦k) a^k b^(n-k),"
	);
	test(
		`∑_2^2▒(n/23)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "opNaryLiteral",
				  "value": "∑"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
				"third": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "CharLiteral",
						"value": "n"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "23"
					  }
					}
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"∑_2^2▒(n/23)"
	);
	test(
		`⏞(x+⋯+x)^(k "times")`,
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
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "k"
					  },
					  {
						"type": "SpaceLiteral",
						"value": " "
					  },
					  {
						"type": "CharLiteral",
						"value": "\"times\""
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		" ⏞(x+⋯+x)^(k 'times')"
	);
	test(
		`𝐸 = 𝑚𝑐^2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "OtherLiteral",
				  "value": "𝐸"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "OperatorLiteral",
				  "value": "="
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "OtherLiteral",
					"value": "𝑚𝑐"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  ]
			]
		  },
		"𝐸 = 𝑚𝑐^2"
	);
	test(
		`∫_0^a▒xⅆx/(x^2+a^2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SubSupLiteral",
				"value": {
				  "type": "opNaryLiteral",
				  "value": "∫"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "0"
				},
				"up": {
				  "type": "CharLiteral",
				  "value": "a"
				},
				"third": {
				  "type": "FractionLiteral",
				  "up": {
					"type": "CharLiteral",
					"value": "xⅆx"
				  },
				  "down": {
					"type": "BracketBlock",
					"value": [
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
						  "type": "SubSupLiteral",
						  "value": {
							"type": "CharLiteral",
							"value": "a"
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
				  }
				}
			  }
			]
		  },
		"∫_0^a▒xⅆx/(x^2+a^2)"
	);
	test(
		`lim┬(n→∞) a_n`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "FunctionLiteral",
					"value": "lim"
				  },
				  "down": {
					"type": "BracketBlock",
					"value": [
					  [
						{
						  "type": "CharLiteral",
						  "value": "n"
						},
						{
						  "type": "OperatorLiteral",
						  "value": "→"
						},
						{
						  "type": "CharLiteral",
						  "value": "∞"
						}
					  ]
					],
					"left": "(",
					"right": ")"
				  }
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "a"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "n"
				  }
				}
			  ]
			]
		  },
		"lim┬(n→∞) a_n"
	);
	test(
		`ⅈ²=-1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "ⅈ"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "2"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "="
				},
				{
				  "type": "OperatorLiteral",
				  "value": "-"
				},
				{
				  "type": "NumberLiteral",
				  "value": "1"
				}
			  ]
			]
		  },
		"ⅈ²=-1"
	);
	test(
		`E = m⁢c²`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "CharLiteral",
				  "value": "E"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "OperatorLiteral",
				  "value": "="
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "CharLiteral",
				  "value": "m"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "⁢"
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "c"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "2"
					}
				  ]
				}
			  ]
			]
		  },
		"E = m⁢c²"
	);
	test(
		`a²⋅b²=c²`,
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "a"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "2"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "⋅"
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "b"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "2"
					}
				  ]
				},
				{
				  "type": "OperatorLiteral",
				  "value": "="
				},
				{
				  "type": "SubSupLiteral",
				  "value": {
					"type": "CharLiteral",
					"value": "c"
				  },
				  "up": [
					{
					  "type": "specialScriptLiteral",
					  "value": "2"
					}
				  ]
				}
			  ]
			]
		  },
		"a²⋅b²=c²"
	);
	// test(
	// 	`f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx`,
	// 	{
	// 		type: "UnicodeEquation",
	// 		body: [
	// 			[
	// 				{
	// 					CharLiteral: "f̂",
	// 				},
	// 				{
	// 					type: "expBracketLiteral",
	// 					exp: {
	// 						type: "anOther",
	// 						value: "ξ",
	// 					},
	// 					open: "(",
	// 					close: ")",
	// 				},
	// 			],
	// 			{
	// 				Operator: "=",
	// 			},
	// 			{
	// 				type: "expSubsup",
	// 				base: {
	// 					type: "opNary",
	// 					value: "∫",
	// 				},
	// 				down: {
	// 					type: "soperandLiteral",
	// 					operand: "-∞",
	// 				},
	// 				up: {
	// 					type: "soperandLiteral",
	// 					operand: "∞",
	// 				},
	// 				thirdSoperand: {
	// 					type: "soperandLiteral",
	// 					operand: [
	// 						{
	// 							CharLiteral: "f",
	// 						},
	// 						{
	// 							type: "expBracketLiteral",
	// 							exp: {
	// 								CharLiteral: "x",
	// 							},
	// 							open: "(",
	// 							close: ")",
	// 						},
	// 						{
	// 							type: "expSuperscript",
	// 							base: {
	// 								CharLiteral: "ⅇ",
	// 							},
	// 							up: {
	// 								type: "soperandLiteral",
	// 								operand: [
	// 									{
	// 										NumberLiteral: "2",
	// 									},
	// 									{
	// 										type: "anOther",
	// 										value: "π",
	// 									},
	// 									{
	// 										CharLiteral: "ⅈxξ",
	// 									},
	// 								],
	// 								minus: true,
	// 							},
	// 						},
	// 					],
	// 				},
	// 			},
	// 			{
	// 				type: "SpaceLiteral",
	// 				value: " ",
	// 			},
	// 			{
	// 				CharLiteral: "ⅆx",
	// 			},
	// 		],
	// 	},
	// 	"Проверка простого литерала: f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx"
	// );
	test(
		`(𝑎 + 𝑏)┴→`,
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
						"type": "OtherLiteral",
						"value": "𝑎"
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
						"value": "𝑏"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				},
				"up": {
				  "type": "OperatorLiteral",
				  "value": "→"
				}
			  }
			]
		  },
		"(𝑎 + 𝑏)┴→"
	);
	test(
		`𝑎┴→`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "OtherLiteral",
				  "value": "𝑎"
				},
				"up": {
				  "type": "OperatorLiteral",
				  "value": "→"
				}
			  }
			]
		  },
		"𝑎┴→"
	);
}
window["AscMath"].complex = complexTest;
