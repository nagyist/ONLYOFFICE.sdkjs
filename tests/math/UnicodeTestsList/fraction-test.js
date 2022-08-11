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

function fractionTests(test) {
	test(
		`1/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "NumberLiteral",
				  "value": "1"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"1/2"
	);
	test(
		`x/2`,
		{
			"type": "UnicodeEquation",
			"body": [
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
		  },
		"x/2"
	);
	test(
		`x+5/2`,
		{
			"type": "UnicodeEquation",
			"body": [
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
				  "type": "FractionLiteral",
				  "up": {
					"type": "NumberLiteral",
					"value": "5"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				}
			  ]
			]
		  },
		"x+5/2"
	);
	test(
		`x+5/x+2`,
		{
			"type": "UnicodeEquation",
			"body": [
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
				  "type": "FractionLiteral",
				  "up": {
					"type": "NumberLiteral",
					"value": "5"
				  },
				  "down": {
					"type": "CharLiteral",
					"value": "x"
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
		"x+5/x+2"
	);
	test(
		`1∕2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "FractionLiteral",
					"fracType": 1,
					up: {
						type: "NumberLiteral",
						value: "1",
					},
					down: {
						type: "NumberLiteral",
						value: "2",
					},
				},
			],
		},
		"1∕2"
	);
	test(
		`(x+5)/2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
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
						"value": "5"
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
		"(x+5)/2"
	);
	test(
		`x/(2+1)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "CharLiteral",
				  "value": "x"
				},
				"down": {
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
			  }
			]
		  },
		"x/(2+1)"
	);
	test(
		`(x-5)/(2+1)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "CharLiteral",
						"value": "x"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "-"
					  },
					  {
						"type": "NumberLiteral",
						"value": "5"
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
			  }
			]
		  },
		"(x-5)/(2+1)"
	);
	test(
		`1+3/2/3`,
		{
			"type": "UnicodeEquation",
			"body": [
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
				  "type": "FractionLiteral",
				  "up": {
					"type": "NumberLiteral",
					"value": "3"
				  },
				  "down": {
					"type": "FractionLiteral",
					"up": {
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
			]
		  },
		"1+3/2/3"
	);
	test(
		`(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "SubSupLiteral",
					  "value": {
						"type": "OtherLiteral",
						"value": "𝛼"
					  },
					  "down": {
						"type": "NumberLiteral",
						"value": "2"
					  },
					  "up": {
						"type": "NumberLiteral",
						"value": "3"
					  }
					}
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
						  "type": "OtherLiteral",
						  "value": "𝛽"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "2"
						},
						"up": {
						  "type": "NumberLiteral",
						  "value": "3"
						}
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "SubSupLiteral",
						"value": {
						  "type": "OtherLiteral",
						  "value": "𝛾"
						},
						"down": {
						  "type": "NumberLiteral",
						  "value": "2"
						},
						"up": {
						  "type": "NumberLiteral",
						  "value": "3"
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
		"(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)"
	);

	test(
		`(a/(b+c))/(d/e + f)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "FractionLiteral",
				"up": {
				  "type": "BracketBlock",
				  "value": [
					{
					  "type": "FractionLiteral",
					  "up": {
						"type": "CharLiteral",
						"value": "a"
					  },
					  "down": {
						"type": "BracketBlock",
						"value": [
						  [
							{
							  "type": "CharLiteral",
							  "value": "b"
							},
							{
							  "type": "OperatorLiteral",
							  "value": "+"
							},
							{
							  "type": "CharLiteral",
							  "value": "c"
							}
						  ]
						],
						"left": "(",
						"right": ")"
					  }
					}
				  ],
				  "left": "(",
				  "right": ")"
				},
				"down": {
				  "type": "BracketBlock",
				  "value": [
					[
					  {
						"type": "FractionLiteral",
						"up": {
						  "type": "CharLiteral",
						  "value": "d"
						},
						"down": {
						  "type": "CharLiteral",
						  "value": "e"
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
						"type": "CharLiteral",
						"value": "f"
					  }
					]
				  ],
				  "left": "(",
				  "right": ")"
				}
			  }
			]
		  },
		"(a/(b+c))/(d/e + f)"
	);

	test(
		`(a/(c/(z/x)))`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "BracketBlock",
				"value": [
				  {
					"type": "FractionLiteral",
					"up": {
					  "type": "CharLiteral",
					  "value": "a"
					},
					"down": {
					  "type": "BracketBlock",
					  "value": [
						{
						  "type": "FractionLiteral",
						  "up": {
							"type": "CharLiteral",
							"value": "c"
						  },
						  "down": {
							"type": "BracketBlock",
							"value": [
							  {
								"type": "FractionLiteral",
								"up": {
								  "type": "CharLiteral",
								  "value": "z"
								},
								"down": {
								  "type": "CharLiteral",
								  "value": "x"
								}
							  }
							],
							"left": "(",
							"right": ")"
						  }
						}
					  ],
					  "left": "(",
					  "right": ")"
					}
				  }
				],
				"left": "(",
				"right": ")"
			  }
			]
		  },
		"(a/(c/(z/x)))"
	);

	test(
		`1¦2`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "BinomLiteral",
				"up": {
				  "type": "NumberLiteral",
				  "value": "1"
				},
				"down": {
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  }
			]
		  },
		"1¦2"
	);
	test(
		`(1¦2)`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "BracketBlock",
				"value": [
				  {
					"type": "BinomLiteral",
					"up": {
					  "type": "NumberLiteral",
					  "value": "1"
					},
					"down": {
					  "type": "NumberLiteral",
					  "value": "2"
					}
				  }
				],
				"left": "(",
				"right": ")"
			  }
			]
		  },
		"(1¦2)"
	);
}
window["AscMath"].fraction = fractionTests;
