/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

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
