function UnicodeBase(test) {
	test(
		"1234567890",
		{
			type: "UnicodeEquation",
			"body": [
					{
						"type": 1,
						"value": "1234567890",
					}
			],
		},
		"Check numbers 1234567890",
	);
	test(
		"abcdefg",
		{
			"type": "UnicodeEquation",
			"body": [
				{
					"type": AscMath.MathLiterals.char.id,
					"value": "abcdefg",
				}
			],
		},
		"Check chars abcdefg"
	);
	test(
		"abc123def",
		{
			"body": [
				[
					{
						"type": AscMath.MathLiterals.char.id,
						"value": "abc"
					},
					{
						"type": AscMath.MathLiterals.number.id,
						"value": "123"
					},
					{
						"type": AscMath.MathLiterals.char.id,
						"value": "def"
					}
				]
			],
			"type": "UnicodeEquation"
		},
		"Check abc123def"
	);
	test(
		"∠",
		{
			"body": [
				{
					"type": AscMath.MathLiterals.operators.id,
					"value": "∠"
				}
			],
			"type": "UnicodeEquation"
		},
		"Check operator ∠"
	);
	test(
		"\\alpha\\beta\\gamma",
		{
			"body": [
				[
					{
					"data": "α",
					"type": AscMath.MathLiterals.operand.id,
				},
				{
					"data": "β",
					"type": AscMath.MathLiterals.operand.id,
				},
				{
					"data": "γ",
					"type": AscMath.MathLiterals.operand.id,
				}
				]
			],
			"type": "UnicodeEquation"
		},
		"Check operator \\alpha\\beta\\gamma"
	);
	test(
		"(",
		{
			"body": [
				{
					"type": AscMath.MathLiterals.char.id,
					"value": "("
				}
			],
			"type": "UnicodeEquation"
		},
		"Check operator ("
	),
	test(
		"(1+\\beta)",
		{
			"body": [
				{
					"left": "(",
					"right": ")",
					"type": AscMath.MathLiterals.bracket.id,
					value: [
						[
							{type: 1, value: '1'},
							{type: 2, value: '+'},
							{type: 3, data: 'β'},
						]
					]
				}
			],
			"type": "UnicodeEquation"
		},
		"Check operator (1+\\beta)"
	)
}
window["AscMath"].UnicodeBase = UnicodeBase;
