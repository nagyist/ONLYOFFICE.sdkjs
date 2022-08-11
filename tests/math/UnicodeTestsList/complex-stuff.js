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
