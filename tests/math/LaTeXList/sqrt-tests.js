function sqrt(test) {
	test(
		"\\sqrt5",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
				  "type": "NumberLiteral",
				  "value": "5"
				}
			  }
			]
		  },
		"\\sqrt5"
	);
	test(
		"\\sqrt\\frac{1}{2}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
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
			  }
			]
		  },
		"\\sqrt\\frac{1}{2}"
	);
	test(
		"\\sqrt[2^2]\\frac{1}{2}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": {
				  "type": "FractionLiteral",
				  "up": {
					"type": "NumberLiteral",
					"value": "1"
				  },
				  "down": {
					"type": "NumberLiteral",
					"value": "2"
				  }
				},
				"index": {
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
			]
		  },
		"\\sqrt[2^2]\\frac{1}{2}"
	);
	test(
		"\\sqrt[2^2] {\\frac{1}{2}+3}",
		{
			"type": "LaTeXEquation",
			"body": [
			  {
				"type": "SqrtLiteral",
				"value": [
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
				  },
				  {
					"type": "OperatorLiteral",
					"value": "+"
				  },
				  {
					"type": "NumberLiteral",
					"value": "3"
				  }
				],
				"index": {
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
			]
		  },
		"\\sqrt[2^2] {\\frac{1}{2}+3}"
	);
}
window["AscMath"].sqrt = sqrt;
