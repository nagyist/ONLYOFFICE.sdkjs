function sqrtTests(test) {
	test(
		`√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					value: {
						type: "NumberLiteral",
						value: "5",
					},
				},
			],
		},
		"√5"
	);
	test(
		`√a`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					value: {
						type: "CharLiteral",
						value: "a",
					},
				},
			],
		},
		"√a"
	);
	test(
		`√a/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
				  "type": "FractionLiteral",
				  "up": {
					"type": "CharLiteral",
					"value": "a"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  }
			]
		  },
		"√a/2"
	);
	test(
		`√(2&a-4)`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "NumberLiteral",
						value: "2",
					},
					value: [
						{
							type: "CharLiteral",
							value: "a",
						},
						{
							type: "OperatorLiteral",
							value: "-",
						},
						{
							type: "NumberLiteral",
							value: "4",
						},
					],
				},
			],
		},
		"√(2&a-4)"
	);
	test(
		`∛5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "CharLiteral",
						value: "3",
					},
					value: {
						type: "NumberLiteral",
						value: "5",
					},
				},
			],
		},
		"∛5"
	);
	test(
		`∛a`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "CharLiteral",
						value: "3",
					},
					value: {
						type: "CharLiteral",
						value: "a",
					},
				},
			],
		},
		"∛a"
	);
	test(
		`∛a/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "SqrtLiteral",
				  "index": {
					"type": "CharLiteral",
					"value": "3"
				  },
				  "value": {
					"type": "CharLiteral",
					"value": "a"
				  }
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"∛a/2"
	);
	test(
		`∛(a-4)`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "CharLiteral",
						value: "3",
					},
					value: {
						type: "BracketBlock",
						value: [
							[
								{
									type: "CharLiteral",
									value: "a",
								},
								{
									type: "OperatorLiteral",
									value: "-",
								},
								{
									type: "NumberLiteral",
									value: "4",
								},
							],
						],
						left: "(",
						right: ")",
					},
				},
			],
		},
		"∛(a-4)"
	);
	test(
		`∜5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "CharLiteral",
						value: "4",
					},
					value: {
						type: "NumberLiteral",
						value: "5",
					},
				},
			],
		},
		"∜5"
	);
	test(
		`∜a`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "CharLiteral",
						value: "4",
					},
					value: {
						type: "CharLiteral",
						value: "a",
					},
				},
			],
		},
		"∜a"
	);
	test(
		`∜a/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "SqrtLiteral",
				  "index": {
					"type": "CharLiteral",
					"value": "4"
				  },
				  "value": {
					"type": "CharLiteral",
					"value": "a"
				  }
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"∜a/2"
	);
	test(
		`∜(a-4)`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					index: {
						type: "CharLiteral",
						value: "4",
					},
					value: {
						type: "BracketBlock",
						value: [
							[
								{
									type: "CharLiteral",
									value: "a",
								},
								{
									type: "OperatorLiteral",
									value: "-",
								},
								{
									type: "NumberLiteral",
									value: "4",
								},
							],
						],
						left: "(",
						right: ")",
					},
				},
			],
		},
		"∜(a-4)"
	);
	test(
		`√(10&a/4)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"index": {
				  "type": "NumberLiteral",
				  "value": "10"
				},
				"value": {
				  "type": "FractionLiteral",
				  "up": {
					"type": "CharLiteral",
					"value": "a"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "4"
				  }
				}
			  }
			]
		  },
		"√(10&a/4)"
	);
	test(
		`√(10^2&a/4+2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"index": {
				  "type": "SubSupLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "10"
				  },
				  "up": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				"value": [
				  {
					"type": "FractionLiteral",
					"up": {
					  "type": "CharLiteral",
					  "value": "a"
					},
					"down": {
					  "type": "NumberLiteral",
					  "value": "4"
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
			  }
			]
		  },
		"√(10^2&a/4+2)"
	);
	test(
		`√5^2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					value: {
						type: "SubSupLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
						up: {
							type: "NumberLiteral",
							value: "2",
						},
					},
				},
			],
		},
		"√5^2"
	);
	test(
		`√5_2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					value: {
						type: "SubSupLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
						down: {
							type: "NumberLiteral",
							value: "2",
						},
					},
				},
			],
		},
		"√5_2"
	);
	test(
		`√5^2_x`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					value: {
						type: "SubSupLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
						down: {
							type: "CharLiteral",
							value: "x",
						},
						up: {
							type: "NumberLiteral",
							value: "2",
						},
					},
				},
			],
		},
		"√5^2_x"
	);
	test(
		`√5_2^x`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SqrtLiteral",
					value: {
						type: "SubSupLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
						down: {
							type: "NumberLiteral",
							value: "2",
						},
						up: {
							type: "CharLiteral",
							value: "x",
						},
					},
				},
			],
		},
		"√5_2^x"
	);
	test(
		`(_5^2)√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "PreScriptLiteral",
					value: {
						type: "SqrtLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
					},
					down: {
						type: "NumberLiteral",
						value: "5",
					},
					up: {
						type: "NumberLiteral",
						value: "2",
					},
				},
			],
		},
		"(_5^2)√5"
	);
	test(
		`√5┴exp1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "5"
				  },
				  "up": {
					"type": "FunctionLiteral",
					"value": "exp",
					"third": {
					  "type": "NumberLiteral",
					  "value": "1"
					}
				  }
				}
			  }
			]
		  },
		"√5┴exp1"
	);
	test(
		`√5┬exp1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
				  "type": "functionWithLimitLiteral",
				  "value": {
					"type": "NumberLiteral",
					"value": "5"
				  },
				  "down": {
					"type": "FunctionLiteral",
					"value": "exp",
					"third": {
					  "type": "NumberLiteral",
					  "value": "1"
					}
				  }
				}
			  }
			]
		  },
		"√5┬exp1"
	);
	test(
		`(√5┬exp1]`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "BracketBlock",
				"value": [
				  {
					"type": "SqrtLiteral",
					"value": {
					  "type": "functionWithLimitLiteral",
					  "value": {
						"type": "NumberLiteral",
						"value": "5"
					  },
					  "down": {
						"type": "FunctionLiteral",
						"value": "exp",
						"third": {
						  "type": "NumberLiteral",
						  "value": "1"
						}
					  }
					}
				  }
				],
				"left": "(",
				"right": "]"
			  }
			]
		  },
		"(√5┬exp1]"
	);
	test(
		`□√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "BoxLiteral",
					value: {
						type: "SqrtLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
					},
				},
			],
		},
		"□√5"
	);
	test(
		`▭√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "RectLiteral",
					value: {
						type: "SqrtLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
					},
				},
			],
		},
		"▭√5"
	);
	test(
		`▁√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "overBarLiteral",
					overUnder: "▁",
					value: {
						type: "SqrtLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
					},
				},
			],
		},
		"▁√5"
	);
	test(
		`¯√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "overBarLiteral",
					overUnder: "¯",
					value: {
						type: "SqrtLiteral",
						value: {
							type: "NumberLiteral",
							value: "5",
						},
					},
				},
			],
		},
		"¯√5"
	);
	test(
		`∑_√5^√5`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SubSupLiteral",
					value: {
						type: "opNaryLiteral",
						value: "∑",
					},
					down: {
						type: "SqrtLiteral",
						value: {
							type: "SubSupLiteral",
							value: {
								type: "NumberLiteral",
								value: "5",
							},
							up: {
								type: "SqrtLiteral",
								value: {
									type: "NumberLiteral",
									value: "5",
								},
							},
						},
					},
				},
			],
		},
		"∑_√5^√5"
	);
	// test(
	// 	`\\root n+1\\of(b+c)+x`,
	// 	{},
	// 	"Check \\root n+1\\of(b+c)+x"
	// );
}
window["AscMath"].sqrt = sqrtTests;
