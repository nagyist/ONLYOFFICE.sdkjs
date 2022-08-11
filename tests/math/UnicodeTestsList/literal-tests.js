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

function literalTests(test) {

	test(
		"×",
		{
			type: "UnicodeEquation",
			body: [{
				type: "OperatorLiteral",
				value: "×"
			}],
		},
		"Check operator: ×"
	);
	test(
		"⋅",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋅"
			}],
		},
		"Check operator: ⋅"
	);
	test(
		"∈",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∈"
			}],
		},
		"Check operator: ∈"
	);
	test(
		"∋",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∋"
			}],
		},
		"Check operator: ∋"
	);
	test(
		"∼",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∼"
			}],
		},
		"Check operator: ∼"
	);
	test(
		"≃",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≃"
			}],
		},
		"Check operator: ≃"
	);
	test(
		"≅",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≅"
			}],
		},
		"Check operator: ≅"
	);
	test(
		"≈",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≈"
			}],
		},
		"Check operator: ≈"
	);
	test(
		"≍",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≍"
			}],
		},
		"Check operator: ≍"
	);
	test(
		"≡",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≡"
			}],
		},
		"Check operator: ≡"
	);
	test(
		"≤",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≤"
			}],
		},
		"Check operator: ≤"
	);
	test(
		"≥",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≥"
			}],
		},
		"Check operator: ≥"
	);
	test(
		"≶",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≶"
			}],
		},
		"Check operator: ≶"
	);
	test(
		"≷",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≷"
			}],
		},
		"Check operator: ≷"
	);
	test(
		"≽",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≽"
			}],
		},
		"Check operator: ≽"
	);
	test(
		"≺",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≺"
			}],
		},
		"Check operator: ≺"
	);
	test(
		"≻",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≻"
			}],
		},
		"Check operator: ≻"
	);
	test(
		"≼",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≼"
			}],
		},
		"Check operator: ≼"
	);
	test(
		"⊂",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊂"
			}],
		},
		"Check operator: ⊂"
	);
	test(
		"⊃",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊃"
			}],
		},
		"Check operator: ⊃"
	);
	test(
		"⊆",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊆"
			}],
		},
		"Check operator: ⊆"
	);
	test(
		"⊇",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊇"
			}],
		},
		"Check operator: ⊇"
	);
	test(
		"⊑",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊑"
			}],
		},
		"Check operator: ⊑"
	);
	test(
		"⊒",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊒"
			}],
		},
		"Check operator: ⊒"
	);
	test(
		"+",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "+"
			}],
		},
		"Check operator: +"
	);
	test(
		"-",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "-"
			}],
		},
		"Check operator: -"
	);
	test(
		"=",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "="
			}],
		},
		"Check operator: ="
	);
	test(
		"*",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "*"
			}],
		},
		"Check operator: *"
	);

	test(
		"∃",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∃"
			}],
		},
		"Check logic operator: ∃"
	);
	test(
		"∀",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∀"
			}],
		},
		"Check logic operator: ∀"
	);
	test(
		"¬",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "¬"
			}],
		},
		"Check logic operator: ¬"
	);
	test(
		"∧",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∧"
			}],
		},
		"Check logic operator: ∧"
	);
	test(
		"∨",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∨"
			}],
		},
		"Check logic operator: ∨"
	);
	test(
		"⇒",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇒"
			}],
		},
		"Check logic operator: ⇒"
	);
	test(
		"⇔",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇔"
			}],
		},
		"Check logic operator: ⇔"
	);
	test(
		"⊕",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊕"
			}],
		},
		"Check logic operator: ⊕"
	);
	test(
		"⊤",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊤"
			}],
		},
		"Check logic operator: ⊤"
	);
	test(
		"⊥",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊥"
			}],
		},
		"Check logic operator: ⊥"
	);
	test(
		"⊢",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊢"
			}],
		},
		"Check logic operator: ⊢"
	);

	test(
		"⨯",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⨯"
			}],
		},
		"Check db operator: ⨯"
	);
	test(
		"⟕",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⟕"
			}],
		},
		"Check db operator: ⟕"
	);
	test(
		"⟖",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⟖"
			}],
		},
		"Check db operator: ⟖"
	);
	test(
		"⟗",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⟗"
			}],
		},
		"Check db operator: ⟗"
	);
	test(
		"⋉",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋉"
			}],
		},
		"Check db operator: ⋉"
	);
	test(
		"⋊",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋊"
			}],
		},
		"Check db operator: ⋊"
	);
	test(
		"▷",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "▷"
			}],
		},
		"Check db operator: ▷"
	);
	test(
		"÷",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "÷"
			}],
		},
		"Check db operator: ÷"
	);

	test(
		"⁡",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⁡"
			}],
		},
		"Check invisible function application operator: ⁡"
	);
	test(
		"⁢",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⁢"
			}],
		},
		"Check invisible times operator: ⁢"
	);
	test(
		"⁣",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⁣"
			}],
		},
		"Check invisible separator operator: ⁣"
	);
	test(
		"⁤",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⁤"
			}],
		},
		"Check invisible plus operator: ⁤"
	);
	test(
		"​",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": "​"
			}],
		},
		"Check zero-width space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check 1/18em space (very very thin math space)"
	);
	test(
		"  ",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				}
			  ]
			]
		  },
		"Check 2/18em space (very thin math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check 3/18em space (thin math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check 5/18em space (thick math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check 6/18em space (very thick math space)"
	);
	test(
		"  ",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				}
			  ]
			]
		  },
		"Check 7/18em space (very very thick math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check 9/18em space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check 1em space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check digit-width space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": [ {
				"type": "SpaceLiteral",
				"value": " ",
			}],
		},
		"Check space-with space (non-breaking space)"
	);

	test(
		`a`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "a"
			  }
			]
		  },
		"Check: a"
	);
	test(
		`abcdef`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "abcdef"
			  }
			]
		  },
		"Check: abcdef"
	);
	test(
		`1`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "NumberLiteral",
				"value": "1"
			  }
			]
		  },
		"Check: 1"
	);
	test(
		`1234`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "NumberLiteral",
				"value": "1234"
			  }
			]
		  },
		"Check: 1234"
	);
	test(
		`1+2`,
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
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  ]
			]
		  },
		"Check: 1+2"
	);
	test(
		`1+2+3`,
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
				  "type": "NumberLiteral",
				  "value": "2"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "3"
				}
			  ]
			]
		  },
		"Check: 1+2+3"
	);

	test(
		`ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω`,
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω"
			  }
			]
		  },
		"Check greek letters: ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω"
	);
	test(
		"abc123def",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "CharLiteral",
				  "value": "abc"
				},
				{
				  "type": "NumberLiteral",
				  "value": "123"
				},
				{
				  "type": "CharLiteral",
				  "value": "def"
				}
			  ]
			]
		  },
		"Check abc123def"
	);
	test(
		"abc+123+def",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "CharLiteral",
				  "value": "abc"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "NumberLiteral",
				  "value": "123"
				},
				{
				  "type": "OperatorLiteral",
				  "value": "+"
				},
				{
				  "type": "CharLiteral",
				  "value": "def"
				}
			  ]
			]
		  },
		"Check abc+123+def"
	);
	test(
		"𝐀𝐁𝐂𝐨𝐹",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "OtherLiteral",
				"value": "𝐀𝐁𝐂𝐨𝐹"
			  }
			]
		  },
		"Check 𝐀𝐁𝐂𝐨𝐹"
	);

	//spaces
	test(
		"   𝐀𝐁𝐂𝐨𝐹   ",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "OtherLiteral",
				  "value": "𝐀𝐁𝐂𝐨𝐹"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				}
			  ]
			]
		},
		"Check '   𝐀𝐁𝐂𝐨𝐹   '"
	);

	//spaces & tabs
	test(
		" 	𝐀𝐁𝐂𝐨𝐹  	 ",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": "\t"
				},
				{
				  "type": "OtherLiteral",
				  "value": "𝐀𝐁𝐂𝐨𝐹"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				},
				{
				  "type": "SpaceLiteral",
				  "value": "\t"
				},
				{
				  "type": "SpaceLiteral",
				  "value": " "
				}
			  ]
			]
		  },
		"Check: ' 	𝐀𝐁𝐂𝐨𝐹  	 '"
	);

	test(
		`1+fbnd+(3+𝐀𝐁𝐂𝐨𝐹)+c+5`,
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
				  "type": "CharLiteral",
				  "value": "fbnd"
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
						"value": "3"
					  },
					  {
						"type": "OperatorLiteral",
						"value": "+"
					  },
					  {
						"type": "OtherLiteral",
						"value": "𝐀𝐁𝐂𝐨𝐹"
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
				  "type": "CharLiteral",
				  "value": "c"
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
			]
		  },
		"Check: '1+fbnd+(3+𝐀𝐁𝐂𝐨𝐹)+c+5'"
	);

	// test(
	// 	`1/3.1416`,
	// 	{
	// 		type: "UnicodeEquation",
	// 		body: {
	// 			type: "expLiteral",
	// 			value: [
	// 				{
	// 					type: "fractionLiteral",
	// 					numerator: {
	// 						type: "numeratorLiteral",
	// 						value: [
	// 							{
	// 								type: "digitsLiteral",
	// 								value: [
	// 									{
	// 										type: "NumericLiteral",
	// 										value: "1",
	// 									},
	// 								],
	// 							},
	// 						],
	// 					},
	// 					opOver: {
	// 						type: "opOver",
	// 						value: "/",
	// 					},
	// 					operand: [
	// 						{
	// 							type: "numberLiteral",
	// 							number: {
	// 								type: "digitsLiteral",
	// 								value: [
	// 									{
	// 										type: "NumericLiteral",
	// 										value: "3",
	// 									},
	// 								],
	// 							},
	// 							decimal: ".",
	// 							after: {
	// 								type: "digitsLiteral",
	// 								value: [
	// 									{
	// 										type: "NumericLiteral",
	// 										value: "1416",
	// 									},
	// 								],
	// 							},
	// 						},
	// 					],
	// 				},
	// 			],
	// 		},
	// 	},
	// 	"Проверка простого литерала - пробелы и табуляция: '1/3.1416'"
	// );


	test(
		"1\\above2",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "functionWithLimitLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "1"
				},
				"up": {
				  "type": "NumberLiteral",
				  "value": "2"
				},
			  }
			]
		  },
		"Check: 1\\above2"
	)
	test(
		"\\above",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "┴"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\above"
	)
	test(
		"1\\acute2",
		{
			"type": "UnicodeEquation",
			"body": [
			  [
				{
				  "type": "AccentLiteral",
				  "base": {
					"type": "NumberLiteral",
					"value": "1"
				  },
				  "value": "́"
				},
				{
				  "type": "NumberLiteral",
				  "value": "2"
				}
			  ]
			]
		  },
		"Check: 1\\acute2"
	)
	test(
		"\\acute",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "́"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\acute"
	)

	test(
		"\\aleph",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℵ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\aleph"
	)
	test(
		"\\alpha",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "α"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\alpha"
	)
	test(
		"\\amalg",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∐"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\amalg"
	);
	test(
		"\\angle",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∠"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\angle"
	)
	test(
		"\\aoint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∳"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\aoint"
	)
	test(
		"\\approx",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≈"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\approx"
	)
	test(
		"\\asmash",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⬆"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\asmash"
	)
	test(
		"\\ast",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∗"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ast"
	)
	test(
		"\\asymp",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≍"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\asymp"
	)
	test(
		"\\atop",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "¦"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\atop"
	)


	test(
		"\\Bar",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "̿"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Bar"
	)
	test(
		"\\bar",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "̅"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bar"
	)
	test(
		"\\because",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∵"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\because"
	)
	test(
		"\\below",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "┬"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\below"
	)
	test(
		"\\beta",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "β"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\beta"
	)
	test(
		"\\beth",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℶ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\beth"
	)
	test(
		"\\bot",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊥"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bot"
	)

	test(
		"\\bigcap",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⋂"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigcap"
	)
	test(
		"\\bigcup",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "opNaryLiteral",
				"value": "⋃"
			  }
			]
		  },
		"Check: \\bigcup"
	)
	test(
		"\\bigodot",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⨀"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigodot"
	)

	test(
		"\\bigoplus",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⨁"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigoplus"
	)
	test(
		"\\bigotimes",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⨂"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigotimes"
	)
	test(
		"\\bigsqcup",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⨆"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigsqcup"
	)
	test(
		"\\biguplus",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⨄"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\biguplus"
	)
	test(
		"\\bigvee",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⋁"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigvee"
	)
	test(
		"\\bigwedge",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⋀"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bigwedge"
	)
	test(
		"\\bowtie",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋈"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bowtie"
	)
	test(
		"\\bra",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "⟨"
			  }
			]
		  },
		"Check: \\bra"
	)
	test(
		"\\breve",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"value": "̆"
			  }
			]
		  },
		"Check: \\breve"
	)
	test(
		"\\bullet",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∙"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\bullet"
	)
	test(
		"\\boxdot",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊡"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\boxdot"
	)
	test(
		"\\boxminus",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊟"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\boxminus"
	)
	test(
		"\\boxplus",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊞"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\boxplus"
	)
	test(
		"\\cap",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∩"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\cap"
	)
	test(
		"\\cbrt",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"index": {
				  "type": "CharLiteral",
				  "value": "3"
				}
			  }
			]
		  },
		"Check: \\cbrt"
	)
	test(
		"\\cdots",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋯"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\cdots"
	)
	test(
		"\\cdot",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋅"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\cdot"
	)
	test(
		"\\check",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "̌"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\check"
	)
	test(
		"\\chi",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "χ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\chi"
	)
	test(
		"\\circ",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∘"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\circ"
	)
	test(
		"\\close",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "┤"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\close"
	)
	test(
		"\\clubsuit",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "♣"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\clubsuit"
	)
	test(
		"\\coint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∲"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\coint"
	)
	test(
		"\\cong",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≅"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\cong"
	)
	test(
		"\\contain",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∋"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\contain"
	)
	test(
		"\\cup",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∪"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\cup"
	)


	test(
		"\\daleth",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℸ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\daleth"
	)
	test(
		"\\dashv",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊣"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\dashv"
	)
	test(
		"\\dd",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ⅆ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\dd"
	)
	test(
		"\\ddddot",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "⃜"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ddddot"
	)
	test(
		"\\dddot",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "⃛"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\dddot"
	)
	test(
		"\\ddot",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "̈"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ddot"
	)
	test(
		"\\ddots",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋱"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ddots"
	)
	test(
		"\\degree",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "°"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\degree"
	)
	test(
		"\\Delta",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Δ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Delta"
	)
	test(
		"\\delta",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "δ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\delta"
	)
	test(
		"\\diamond",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋄"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\diamond"
	)

	test(
		"\\diamondsuit",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "♢"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\diamondsuit"
	)
	test(
		"\\div",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "÷"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\div"
	)
	test(
		"\\dot",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "̇"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\dot"
	)
	test(
		"\\doteq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≐"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\doteq"
	)
	test(
		"\\dots",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "…"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\dots"
	)
	test(
		"\\downarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↓"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\downarrow"
	)
	test(
		"\\dsmash",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⬇"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\dsmash"
	)

	test(
		"\\degc",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "℃"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\degc"
	)
	test(
		"\\degf",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "℉"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\degf"
	)


	test(
		"\\ee",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ⅇ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ee"
	)
	test(
		"\\ell",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℓ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ell"
	)
	test(
		"\\emptyset",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∅"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\emptyset"
	)
	test(
		"\\emsp",
		{
			"body": [ {
				"type": "SpaceLiteral",
				"value": " "
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\emsp"
	)
	test(
		"\\end",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "〗"
			  }
			]
		  },
		"Check: \\end"
	)
	test(
		"\\ensp",
		{
			"body": [ {
				"type": "SpaceLiteral",
				"value": " "
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ensp"
	)
	test(
		"\\epsilon",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ϵ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\epsilon"
	)
	test(
		"\\eqarray",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "█"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\eqarray"
	)
	test(
		"\\eqno",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "#"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\eqno"
	)
	test(
		"\\equiv",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≡"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\equiv"
	)
	test(
		"\\eta",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "η"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\eta"
	)
	test(
		"\\exists",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∃"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\exists"
	)


	test(
		"\\forall",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∀"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\forall"
	)
	test(
		"\\funcapply",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⁡"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\funcapply"
	)
	test(
		"\\frown",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⌑"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\frown"
	)

	test(
		"\\Gamma",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Γ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Gamma"
	)
	test(
		"\\gamma",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "γ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\gamma"
	)
	test(
		"\\ge",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≥"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ge"
	)
	test(
		"\\geq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≥"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\geq"
	)
	test(
		"\\gets",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "←"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\gets"
	)
	test(
		"\\gg",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≫"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\gg"
	)
	test(
		"\\gimel",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℷ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\gimel"
	)
	test(
		"\\grave",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "̀"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\grave"
	)

	test(
		"\\hairsp",
		{
			"body": [ {
				"type": "SpaceLiteral",
				"value": " "
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\hairsp"
	)
	test(
		"\\hat",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "AccentLiteral",
				"value": "̂"
			  }
			]
		  },
		"Check: \\hat"
	)
	test(
		"\\hbar",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "ℏ"
			  }
			]
		  },
		"Check: \\hbar"
	)
	test(
		"\\heartsuit",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "♡"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\heartsuit"
	)
	test(
		"\\hookleftarrow",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "OperatorLiteral",
				"value": "↩"
			  }
			]
		  },
		"Check: \\hookleftarrow"
	)

	test(
		"\\hphantom",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⬄"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\hphantom"
	)


	test(
		"\\hsmash",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⬌"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\hsmash"
	)
	test(
		"\\hvec",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "⃑"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\hvec"
	)


	test(
		"\\Im",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℑ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Im"
	)
	test(
		"\\iiiint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "⨌"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\iiiint"
	)
	test(
		"\\iiint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∭"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\iiint"
	)
	test(
		"\\iint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∬"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\iint"
	)
	test(
		"\\ii",
		{
			"body": [
				{
					"type": "CharLiteral",
					"value": "ⅈ"
				  }
			],
			"type": "UnicodeEquation"
		},
		"Check: \\ii"
	)
	test(
		"\\int",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∫"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\int"
	)
	test(
		"\\imath",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "𝚤"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\imath"
	)
	test(
		"\\inc",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∆"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\inc"
	)
	test(
		"\\infty",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "∞"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\infty"
	)
	test(
		"\\in",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∈"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\in"
	)
	test(
		"\\iota",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ι"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\iota"
	)
	test(
		"\\jj",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ⅉ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\jj"
	)
	test(
		"\\jmath",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "𝚥"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\jmath"
	)
	test(
		"\\kappa",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "κ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\kappa"
	)
	test(
		"\\ket",
		{
			"body": [ {
			    "type": "CharLiteral",
  			    "value": "⟩"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ket"
	)


	test(
		"\\Lambda",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Λ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Lambda"
	)

	test(
		"\\lambda",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "λ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\lambda"
	)
	test(
		"\\langle",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "⟨"
			  }
			]
		  },
		"Check: \\langle"
	)
	test(
		"\\lbrack",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "["
			  }
			]
		  },
		"Check: \\lbrack"
	)

	test(
		"\\ldiv",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "∕"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ldiv"
	)
	test(
		"\\ldots",
		{
			"body": [ {
			    "type": "OperatorLiteral",
    		    "value": "…"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ldots"
	)
	test(
		"\\le",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≤"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\le"
	)
	test(
		"\\Leftarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇐"
		  
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Leftarrow"
	)
	test(
		"\\leftarrow",
		{
			"body": [ {
			    "type": "OperatorLiteral",
      			"value": "←"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\leftarrow"
	)
	test(
		"\\leftharpoondown",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↽"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\leftharpoondown"
	)
	test(
		"\\leftharpoonup",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↼"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\leftharpoonup"
	)
	test(
		"\\Leftrightarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇔"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Leftrightarrow"
	)
	test(
		"\\leftrightarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↔"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\leftrightarrow"
	)
	test(
		"\\leq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≤"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\leq"
	)
	test(
		"\\lfloor",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "⌊"
			  }
			]
		  },
		"Check: \\lfloor"
	)
	test(
		"\\ll",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≪"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ll"
	)
	test(
		"\\Longleftarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⟸"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Longleftarrow"
	)
	test(
		"\\longleftarrow",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "\\longleftarrow"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\longleftarrow"
	)

	test(
		"\\lmoust",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "⎰"
			  }
			]
		  },
		"Check: \\lmoust"
	)

	test(
		"\\mapsto",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↦"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\mapsto"
	)
	test(
		"\\medsp",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SpaceLiteral",
				"value": " "
			  }
			]
		  },
		"Check: \\medsp"
	)
	test(
		"\\mid",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "∣"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\mid"
	)
	test(
		"\\models",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊨"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\models"
	)
	test(
		"\\mp",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∓"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\mp"
	)
	test(
		"\\mu",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "μ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\mu"
	)
	test(
		"\\nabla",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∇"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\nabla"
	)
	test(
		"\\naryand",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "▒"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\naryand"
	)
	// test(
	// 	"\\nbsp",
	// 	{
	// 		"type": "UnicodeEquation",
	// 		"body": [
	// 		  {
	// 			"type": "SpaceLiteral",
	// 			"value": " "
	// 		  }
	// 		]
	// 	  },
	// 	"Check: \\nbsp"
	// )
	test(
		"\\ne",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≠"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ne"
	)
	test(
		"\\nearrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↗"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\nearrow"
	)
	test(
		"\\neg",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "¬"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\neg"
	)
	test(
		"\\neq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≠"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\neq"
	)
	test(
		"\\ni",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∋"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ni"
	)
	test(
		"\\norm",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "‖"
			  }
			]
		  },
		"Check: \\norm"
	)
	test(
		"\\nu",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ν"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\nu"
	)
	test(
		"\\nwarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↖"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\nwarrow"
	)

	test(
		"\\Omega",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Ω"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Omega"
	)
	test(
		"\\odot",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊙"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\odot"
	)
	test(
		"\\of",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "▒"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\of"
	)
	test(
		"\\oiiint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∰"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\oiiint"
	)
	test(
		"\\oiint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∯"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\oiint"
	)
	test(
		"\\oint",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∮"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\oint"
	)
	test(
		"\\omega",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ω"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\omega"
	)
	test(
		"\\ominus",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊖"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ominus"
	)
	// test(
	// 	"\\open",
	// 	{
	// 		"body": [ {
	// 			"type": "OperatorLiteral",
	// 			"value": "̀"
	// 		}],
	// 		"type": "UnicodeEquation"
	// 	},
	// 	"Check: \\open"
	// )
	test(
		"\\oplus",
		{
			"body": [ {
				"type": "OperatorLiteral",
     			"value": "⊕"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\oplus"
	)

	test(
		"\\otimes",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊗"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\otimes"
	)
	test(
		"\\overbar",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "overBarLiteral",
				"overUnder": "¯"
			  }
			]
		  },
		"Check: \\overbar"
	)
	test(
		"\\overbrace",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⏞"
			  }
			]
		  },
		"Check: \\overbrace"
	)
	test(
		"\\overbracket",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⎴"
			  }
			]
		  },
		"Check: \\overbracket"
	)
	test(
		"\\overparen",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⏜"
			  }
			]
		  },
		"Check: \\overparen"
	)
	test(
		"\\overshell",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "hBracketLiteral",
				"hBrack": "⏠"
			  }
			]
		  },
		"Check: \\overshell"
	)
	test(
		"\\Pi",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Π"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Pi"
	)
	test(
		"\\Phi",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Φ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Phi"
	)
	test(
		"\\Psi",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "Ψ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Psi"
	)
	test(
		"\\parallel",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∥"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\parallel"
	)
	test(
		"\\partial",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "∂"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\partial"
	)
	test(
		"\\perp",
		{
			"body": [ {
			    "type": "OperatorLiteral",
    			  "value": "⊥"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\perp"
	)
	test(
		"\\phantom",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "CharLiteral",
				"value": "⟡"
			  }
			]
		  },
		"Check: \\phantom"
	)
	test(
		"\\phi",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ϕ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\phi"
	)
	test(
		"\\pi",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "π"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\pi"
	)
	test(
		"\\pm",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "±"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\pm"
	)
	test(
		"\\pppprime",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "⁗"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\pppprime"
	)
	test(
		"\\ppprime",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "‴"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ppprime"
	)
	test(
		"\\pprime",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "″"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\pprime"
	)
	test(
		"\\prec",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≺"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\prec"
	)
	test(
		"\\prime",
		{
			"body": [ {
				"type": "AccentLiteral",
				"value": "′"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\prime"
	)
	test(
		"\\propto",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "∝"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\propto"
	)
	test(
		"\\psi",
		{
			"body": [ {
				"type": "CharLiteral",
      			"value": "ψ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\psi"
	)
	test(
		"\\qdrt",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"index": {
				  "type": "CharLiteral",
				  "value": "4"
				}
			  }
			]
		  },
		"Check: \\qdrt"
	)
	test(
		"\\Re",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "ℜ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Re"
	)
	test(
		"\\Rightarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇒"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Rightarrow"
	)
	test(
		"\\rangle",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "⟩"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rangle"
	)
	test(
		"\\ratio",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∶"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\ratio"
	)
	test(
		"\\rbrace",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "}"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rbrace"
	)
	test(
		"\\rbrack",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "]"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rbrack"
	)
	test(
		"\\rceil",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "⌉"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rceil"
	)
	test(
		"\\rddots",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋰"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rddots"
	)
	test(
		"\\rect",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "▭"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rect"
	)
	test(
		"\\rfloor",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "⌋"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rfloor"
	)
	test(
		"\\rho",
		{
			"body": [ {
			    "type": "CharLiteral",
   			    "value": "ρ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rho"
	)
	test(
		"\\right",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "┤"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\right"
	)
	test(
		"\\rightarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "→"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rightarrow"
	)
	test(
		"\\rightharpoondown",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇁"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rightharpoondown"
	)
	test(
		"\\rightharpoonup",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⇀"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rightharpoonup"
	)
	test(
		"\\rmoust",
		{
			"body": [ {
			    "type": "CharLiteral",
    			"value": "⎱"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\rmoust"
	)
	// test(
	// 	"\\rrect",
	// 	{
	// 		"body": [ {
	// 			"type": "OperatorLiteral",
	// 			"value": "̀"
	// 		}],
	// 		"type": "UnicodeEquation"
	// 	},
	// 	"Check: \\rrect"
	// )
	test(
		"\\root",
		{
			"type": "UnicodeEquation",
			"body": [
			  {
				"type": "SqrtLiteral"
			  }
			]
		  },
		"Check: \\root"
	)
	test(
		"\\Sigma",
		{
			"body": [ {
				"type": "CharLiteral",
     			"value": "Σ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\Sigma"
	)
	test(
		"\\sdiv",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "⁄"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sdiv"
	)
	test(
		"\\searrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↘"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\searrow"
	)
	test(
		"\\setminus",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "∖"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\setminus"
	)
	test(
		"\\sigma",
		{
			"body": [ {
				"type": "CharLiteral",
				"value": "σ"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sigma"
	)
	test(
		"\\sim",
		{
			"body": [ {
			    "type": "OperatorLiteral",
      			"value": "∼"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sim"
	)
	test(
		"\\simeq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≃"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\simeq"
	)
	test(
		"\\smash",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⬍"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\smash"
	)
	test(
		"\\smile",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⌣"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\smile"
	)
	test(
		"\\spadesuit",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "♠"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\spadesuit"
	)
	test(
		"\\sqcap",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊓"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sqcap"
	)
	test(
		"\\sqcup",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊔"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sqcup"
	)
	test(
		"\\sqsubseteq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊑"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sqsubseteq"
	)
	test(
		"\\sqsuperseteq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊒"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sqsuperseteq"
	)
	test(
		"\\star",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⋆"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\star"
	)
	test(
		"\\subset",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊂"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\subset"
	)
	test(
		"\\subseteq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "⊆"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\subseteq"
	)
	test(
		"\\succeq",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≽"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\succeq"
	)
	test(
		"\\succ",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "≻"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\succ"
	)
	test(
		"\\sum",
		{
			"body": [ {
				"type": "opNaryLiteral",
				"value": "∑"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\sum"
	)
	test(
		"\\superset",
		{
			"body": [ {
			    "type": "OperatorLiteral",
      			"value": "⊃"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\superset"
	)
	test(
		"\\superseteq",
		{
			"body": [ {
				"type": "OperatorLiteral",
     			 "value": "⊇"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\superseteq"
	)
	test(
		"\\swarrow",
		{
			"body": [ {
				"type": "OperatorLiteral",
				"value": "↙"
			}],
			"type": "UnicodeEquation"
		},
		"Check: \\swarrow"
	)
}
window["AscMath"].literal = literalTests;
