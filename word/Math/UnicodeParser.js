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

"use strict";

(function (window) {
	const num = 1; //needs for debug, default value: 0

	const oLiteralNames = window.AscCommonWord.oNamesOfLiterals;
	const UnicodeSpecialScript = window.AscCommonWord.UnicodeSpecialScript;
	const ConvertTokens = window.AscCommonWord.ConvertTokens;
	const Tokenizer = window.AscCommonWord.Tokenizer;

	function CUnicodeParser() {
		this.oTokenizer = new Tokenizer(false);
		this.isOneSubSup = false;
		this.isTextLiteral = false;

		//need for group like "|1+2|"
		this._isNotStepInBracketBlock = false;
	}
	CUnicodeParser.prototype.Parse = function (string) {
		this.oTokenizer.Init(string);
		this._lookahead = this.oTokenizer.GetNextToken();
		return this.Program();
	};
	CUnicodeParser.prototype.Program = function () {
		return {
			type: "UnicodeEquation",
			body: this.GetExpLiteral(),
		};
	};
	CUnicodeParser.prototype.GetSpaceLiteral = function () {
		const oSpaceLiteral = this.EatToken(oLiteralNames.spaceLiteral[0]);
		return {
			type: oLiteralNames.spaceLiteral[num],
			value: oSpaceLiteral.data,
		};
	};
	// CUnicodeParser.prototype.GetAASCIILiteral = function () {
	// 	const oASCIILiteral = this.EatToken(oLiteralNames.asciiLiteral[0]);
	// 	return {
	// 		type: oLiteralNames.asciiLiteral[num],
	// 		value: oASCIILiteral.data,
	// 	};
	// };
	// CUnicodeParser.prototype.GetOpArrayLiteral = function () {
	// 	const oOpArrayLiteral = this.EatToken(oLiteralNames.opArrayLiteral[0]);
	// 	return {
	// 		type: oLiteralNames.opArrayLiteral[num],
	// 		value: oOpArrayLiteral.data,
	// 	};
	// };
	CUnicodeParser.prototype.GetOpCloseLiteral = function () {
		let oCloseLiteral;
		if (this._lookahead.class === oLiteralNames.opOpenCloseBracket[0]) {
			oCloseLiteral = this.EatToken(oLiteralNames.opOpenCloseBracket[0]);
			return oCloseLiteral.data;
		}
		oCloseLiteral = this.EatToken(oLiteralNames.opCloseBracket[0]);
		return oCloseLiteral.data;
	};
	CUnicodeParser.prototype.GetOpCloserLiteral = function () {
		switch (this._lookahead.class) {
			case "\\close":
				return {
					type: oLiteralNames.opCloseBracket[num],
					value: this.EatToken("\\close").data,
				};
			case "┤":
				return {
					type: oLiteralNames.opCloseBracket[num],
					value: this.EatToken("┤").data,
				};
			case oLiteralNames.opCloseBracket[0]:
				return this.GetOpCloseLiteral();
			case oLiteralNames.opOpenCloseBracket[0]:
				return this.GetOpCloseLiteral();
		}
	};
	// CUnicodeParser.prototype.IsOpCloserLiteral = function () {
	// 	return this._lookahead.class === oLiteralNames.opCloseBracket[0] || this._lookahead.class === "\\close";
	// };
	// CUnicodeParser.prototype.GetOpDecimalLiteral = function () {
	// 	const oOpDecimal = this.EatToken(oLiteralNames.opDecimal[0]);
	// 	return {
	// 		type: oLiteralNames.opDecimal[num],
	// 		value: oOpDecimal.data,
	// 	};
	// };
	// CUnicodeParser.prototype.GetOpHBracketLiteral = function () {
	// 	const oOpHBracket = this.EatToken(oLiteralNames.hBracketLiteral[0]);
	// 	return {
	// 		type: oLiteralNames.hBracketLiteral[1],
	// 		value: oOpHBracket.data,
	// 	};
	// };
	CUnicodeParser.prototype.IsOpNaryLiteral = function () {
		return this._lookahead.class === oLiteralNames.opNaryLiteral[0];
	}
	CUnicodeParser.prototype.GetOpNaryLiteral = function () {
		const oOpNaryLiteral = this.EatToken(oLiteralNames.opNaryLiteral[0]);
		if (this._lookahead.class === "▒") {
			this.EatToken("▒");
			oOpNaryLiteral.content = this.GetSoOperandLiteral();
		}
		return {
			type: oLiteralNames.opNaryLiteral[num],
			value: oOpNaryLiteral.data,
			content: oOpNaryLiteral.content,
		};
	};
	CUnicodeParser.prototype.GetOpOpenLiteral = function () {
		let oOpLiteral;
		if (this._lookahead.class === oLiteralNames.opOpenCloseBracket[0]) {
			oOpLiteral = this.EatToken(oLiteralNames.opOpenCloseBracket[0]);
			return oOpLiteral.data;
		}
		oOpLiteral = this.EatToken(oLiteralNames.opOpenBracket[0]);
		return oOpLiteral.data;
	};
	// CUnicodeParser.prototype.GetOpOpenerLiteral = function () {
	// 	switch (this._lookahead.class) {
	// 		case "\\open":
	// 			return {
	// 				type: oLiteralNames.opOpenBracket[num],
	// 				value: "\\open",
	// 			};
	// 		case oLiteralNames.opOpenBracket[0]:
	// 			return this.GetOpOpenLiteral();
	// 	}
	// };
	CUnicodeParser.prototype.IsOpOpenerLiteral = function () {
		return this._lookahead.class === oLiteralNames.opOpenBracket[0];
	};
	CUnicodeParser.prototype.GetDigitsLiteral = function () {
		const arrNASCIIList = [this.GetASCIILiteral()];
		while (this._lookahead.class === "nASCII") {
			arrNASCIIList.push(this.GetASCIILiteral());
		}
		return arrNASCIIList;
	};
	CUnicodeParser.prototype.IsDigitsLiteral = function () {
		return this._lookahead.class === oLiteralNames.numberLiteral[0];
	};
	CUnicodeParser.prototype.GetNumberLiteral = function () {
		let strDecimal,
			oRightDigit,
			oLeftDigit = this.GetDigitsLiteral();

		if (this._lookahead.class === "opDecimal") {
			strDecimal = this.EatToken("opDecimal").data;
			if (this.IsDigitsLiteral()) {
				oRightDigit = this.GetDigitsLiteral();
				return {
					type: "numberLiteral",
					number: oLeftDigit,
					decimal: strDecimal,
					after: oRightDigit,
				};
			}
		}
		return oLeftDigit;
	};
	CUnicodeParser.prototype.IsNumberLiteral = function () {
		return this.IsDigitsLiteral();
	};
	CUnicodeParser.prototype.EatCloseOrOpenBracket = function () {
		if (this._lookahead.class === "├") {
			this.EatToken("├");
			const oOpenLiteral = this.EatBracket();
			const oExp = this.GetExpLiteral();
			this.EatToken("┤");
			const oCloseLiteral = this.EatBracket();

			return {
				type: "expBracketLiteral",
				open: oOpenLiteral,
				close: oCloseLiteral,
				value: oExp,
			};
		}
	};
	CUnicodeParser.prototype.EatBracket = function () {
		let oBracket;
		switch (this._lookahead.class) {
			case oLiteralNames.opCloseBracket[0]:
				oBracket = this.GetOpCloseLiteral();
				break;
			case oLiteralNames.opOpenBracket[0]:
				oBracket = this.GetOpOpenLiteral();
				break;
			case "Char":
				oBracket = this.GetCharLiteral();
				break;
			case "nASCII":
				oBracket = this.GetASCIILiteral();
				break;
			case oLiteralNames.spaceLiteral[0]:
				oBracket = this.GetSpaceLiteral();
				break;
		}
		return oBracket;
	};
	CUnicodeParser.prototype.GetWordLiteral = function () {
		const arrWordList = [this.GetASCIILiteral()];
		while (this._lookahead.class === oLiteralNames.asciiLiteral[0]) {
			arrWordList.push(this.GetASCIILiteral());
		}
		return {
			type: "wordLiteral",
			value: this.GetContentOfLiteral(arrWordList),
		};
	};
	CUnicodeParser.prototype.IsWordLiteral = function () {
		return this._lookahead.class === oLiteralNames.asciiLiteral[0];
	};
	CUnicodeParser.prototype.GetSoOperandLiteral = function (isSubSup) {
		if (this.IsOperandLiteral()) {
			return this.GetOperandLiteral(isSubSup);
		}
		switch (this._lookahead.data) {
			case "-":
				this.EatToken(oLiteralNames.operatorLiteral[0]);
				if (this.IsOperandLiteral()) {
					const operand = this.GetOperandLiteral();
					return {
						type: "minusLiteral",
						value: operand,
					};
				}
				break;
			case "-∞":
				const token = this.EatToken(oLiteralNames.operatorLiteral[0]);
				return token.data;
			case "∞":
				const tokens = this.EatToken(oLiteralNames.operatorLiteral[0]);
				return tokens.data;
		}
		if (this._lookahead.class === oLiteralNames.operatorLiteral[0]) {
			return this.GetOperatorLiteral();
		}
	};
	CUnicodeParser.prototype.IsSoOperandLiteral = function () {
		return (
			this.IsOperandLiteral() ||
			this._lookahead.data === "-" ||
			this._lookahead.data === "-∞" ||
			this._lookahead.data === "∞"
		);
	};
	CUnicodeParser.prototype.IsTextLiteral = function () {
		return (this._lookahead.class === "\"" || this._lookahead.class === "\'") && !this.isTextLiteral
	}
	CUnicodeParser.prototype.GetTextLiteral = function () {
		let strSymbol = this.EatToken(this._lookahead.class);
		let strExp = "";
		this.oTokenizer.SaveState();
		this.saveLookahead = this._lookahead;

		while (this._lookahead.class !== "\"" && this._lookahead.class !== "\'" && this._lookahead.class !== undefined) {
			strExp += this.EatToken(this._lookahead.class).data
		}
		if (this._lookahead.class === undefined) {
			this.oTokenizer.RestoreState();
			this._lookahead = this.saveLookahead;
			this.saveLookahead = undefined;
			return {
				type: oLiteralNames.charLiteral[0],
				value: strSymbol,
			}
		}
		else {
			this.EatToken(this._lookahead.class)
			return {
				type: oLiteralNames.textLiteral[num],
				value: strExp,
			}
		}
	}
	CUnicodeParser.prototype.IsBoxLiteral = function () {
		return this._lookahead.data === "□";
	};
	CUnicodeParser.prototype.GetBoxLiteral = function () {
		if (this._lookahead.data === "□") {
			this.EatToken(this._lookahead.class);
			if (this.IsOperandLiteral()) {
				const oToken = this.GetOperandLiteral();
				return {
					type: oLiteralNames.boxLiteral[0],
					value: oToken,
				};
			}
		}
	};
	CUnicodeParser.prototype.GetRectLiteral = function () {
		if (this._lookahead.data === "▭") {
			this.EatToken(this._lookahead.class);
			if (this.IsOperandLiteral()) {
				const oToken = this.GetOperandLiteral();
				return {
					type: oLiteralNames.rectLiteral[num],
					value: oToken,
				};
			}
		}
	};
	CUnicodeParser.prototype.isRectLiteral = function () {
		return this._lookahead.data === "▭";
	};
	CUnicodeParser.prototype.GetOverBarLiteral = function () {
		let oToken;
		this.EatToken(oLiteralNames.diacriticLiteral[0]);
		if (this.IsOperandLiteral()) {
			oToken = this.GetOperandLiteral();
			return {
				type: oLiteralNames.overBarLiteral[num],
				value: oToken,
			};
		}
	};
	CUnicodeParser.prototype.IsOverBarLiteral = function () {
		return this._lookahead.data === "̄";
	};
	CUnicodeParser.prototype.GetUnderBarLiteral = function () {
		if (this._lookahead.class === "▁") {
			this.EatToken("▁");
			if (this.IsOperandLiteral()) {
				const token = this.GetOperandLiteral();
				return {
					type: oLiteralNames.overBarLiteral[num],
					value: token,
				};
			}
		}
	};
	CUnicodeParser.prototype.IsUnderBarLiteral = function () {
		return this._lookahead.class === "▁";
	};
	CUnicodeParser.prototype.GetHBracketLiteral = function () {
		let oUp, oDown, oOperand;
		if (this.IsOperandLiteral()) {
			this.EatToken(oLiteralNames.hBracketLiteral[0]);
			this.SkipSpace();
			oOperand = this.GetOperandLiteral();
			if (this._lookahead.data === "_" || this._lookahead.data === "^") {
				if (this._lookahead.class === "_") {
					oDown = this.GetScriptStandardContentLiteral(undefined, true);
				}
				else {
					oUp = this.GetScriptStandardContentLiteral(undefined, true);
				}
			}
			return {
				type: oLiteralNames.hBracketLiteral[num],
				value: oOperand,
				up: oUp,
				down: oDown,
			};
		}
	};
	CUnicodeParser.prototype.IsHBracketLiteral = function () {
		return this._lookahead.class === oLiteralNames.hBracketLiteral[0];
	};
	CUnicodeParser.prototype.GetSqrtLiteral = function () {
		let oToken;
		this.EatToken(oLiteralNames.opBuildupLiteral[0]);
		this.SkipSpace();
		if (this.IsOperandLiteral()) {
			oToken = this.GetElementLiteral();
		}
		return {
			type: oLiteralNames.sqrtLiteral[num],
			value: oToken ? oToken : undefined,
		};
	};
	CUnicodeParser.prototype.IsSqrtLiteral = function () {
		return this._lookahead.data === "√";
	};
	CUnicodeParser.prototype.GetCubertLiteral = function () {
		let oToken;
		this.EatToken(oLiteralNames.opBuildupLiteral[0]);
		if (this.IsOperandLiteral()) {
			oToken = this.GetOperandLiteral();
			return {
				type: oLiteralNames.cubertLiteral[num],
				value: oToken,
			};
		}
	};
	CUnicodeParser.prototype.IsCubertLiteral = function () {
		return this._lookahead.data === "∛" && this._lookahead.class !== oLiteralNames.operatorLiteral[0];
	};
	CUnicodeParser.prototype.GetFourthrtLiteral = function () {
		let oContent;
		this.EatToken(oLiteralNames.opBuildupLiteral[0]);
		if (this.IsOperandLiteral()) {
			oContent = this.GetOperandLiteral();
			return {
				type: oLiteralNames.fourthrtLiteral[num],
				value: oContent,
			};
		}
	};
	CUnicodeParser.prototype.IsFourthrtLiteral = function () {
		return this._lookahead.data === "∜" && this._lookahead.class !== oLiteralNames.operatorLiteral[0];
	};
	CUnicodeParser.prototype.GetNthrtLiteral = function () {
		let oIndex, oContent;
		this.EatToken(oLiteralNames.opBuildupLiteral[0]);
		if (this.IsOperandLiteral()) {
			oIndex = this.GetExpLiteral();
			if (this._lookahead.data === "&") {
				this.EatToken("opArray");
				if (this.IsOperandLiteral()) {
					oContent = this.GetExpLiteral();
					this.EatToken(oLiteralNames.opCloseBracket[0]);
					return {
						type: oLiteralNames.nthrtLiteral[num],
						index: oIndex,
						value: oContent,
					};
				}
			}
		}
	};
	CUnicodeParser.prototype.IsNthrtLiteral = function () {
		return this._lookahead.data === "√(" && this._lookahead.class !== oLiteralNames.operatorLiteral[0];
	};
	CUnicodeParser.prototype.ProceedSqrt = function () {
		if (this._lookahead.class === "▒") {
			this.EatToken("▒");
			return this.GetEntityLiteral();
		}
	};
	CUnicodeParser.prototype.IsFunctionLiteral = function () {
		return (
			this.IsSqrtLiteral() ||
			this.IsCubertLiteral() ||
			this.IsFourthrtLiteral() ||
			this.IsNthrtLiteral() ||
			this.IsBoxLiteral() ||
			this.isRectLiteral() ||
			this.IsOverBarLiteral() ||
			this.IsUnderBarLiteral() ||
			this.IsHBracketLiteral()
		);
	};
	CUnicodeParser.prototype.GetFunctionLiteral = function () {
		let oFunctionContent;

		if (this.IsSqrtLiteral()) {
			oFunctionContent = this.GetSqrtLiteral();

			let temp = this.ProceedSqrt();
			if (temp) {
				oFunctionContent.index = temp;
			}
		}
		else if (this.IsCubertLiteral()) {
			oFunctionContent = this.GetCubertLiteral();
		}
		else if (this.IsFourthrtLiteral()) {
			oFunctionContent = this.GetFourthrtLiteral();
		}
		else if (this.IsNthrtLiteral()) {
			oFunctionContent = this.GetNthrtLiteral();
		}
		else if (this.IsBoxLiteral()) {
			oFunctionContent = this.GetBoxLiteral();
		}
		else if (this.isRectLiteral()) {
			oFunctionContent = this.GetRectLiteral();
		}
		else if (this.IsOverBarLiteral()) {
			oFunctionContent = this.GetOverBarLiteral();
		}
		else if (this.IsUnderBarLiteral()) {
			oFunctionContent = this.GetUnderBarLiteral();
		}
		else if (this.IsHBracketLiteral()) {
			oFunctionContent = this.GetHBracketLiteral();
		}
		return oFunctionContent;
	};
	CUnicodeParser.prototype.IsExpBracketLiteral = function () {
		return (
			this.IsOpOpenerLiteral() ||
			(this._lookahead.class === oLiteralNames.opOpenCloseBracket[0] &&
				this._isNotStepInBracketBlock === false) ||
			this._lookahead.class === "├"
		);
	};
	CUnicodeParser.prototype.GetExpBracketLiteral = function () {
		let open,
			close,
			exp;

		if (this._lookahead.class === oLiteralNames.opOpenBracket[0] || this._lookahead.class === oLiteralNames.opOpenCloseBracket[0]) {
			open = this.GetOpOpenLiteral();
			this._isNotStepInBracketBlock = true;
			if (this.IsPreScriptLiteral() && open === "(") {
				return this.GetPreScriptLiteral();
			}
			exp = this.GetExpLiteral();
			if (this._lookahead.class)
			close = this.GetOpCloserLiteral();
			this._isNotStepInBracketBlock = false;
		}
		else if (this._lookahead.class === "├") {
			return this.EatCloseOrOpenBracket();
		}

		return {
			type: "expBracketLiteral",
			exp,
			open,
			close,
		};
	};
	CUnicodeParser.prototype.GetPreScriptLiteral = function () {
		let oFirstSoOperand,
			oSecondSoOperand,
			oBase;
		let strTypeOfPreScript = this._lookahead.data;

		this.EatToken(this._lookahead.class);
		if (strTypeOfPreScript === "_") {
			oFirstSoOperand = this.GetSoOperandLiteral("preScript");
		}
		else {
			oSecondSoOperand = this.GetSoOperandLiteral("preScript");
		}

		if (this._lookahead.data !== strTypeOfPreScript && this.IsPreScriptLiteral()) {
			this.EatToken(this._lookahead.class);
			if (strTypeOfPreScript === "_") {
				oSecondSoOperand = this.GetSoOperandLiteral("preScript");
			}
			else {
				oFirstSoOperand = this.GetSoOperandLiteral("preScript");
			}
		}

		if (this._lookahead.class === oLiteralNames.opOpenCloseBracket[0]) {
			this.EatToken(oLiteralNames.opOpenCloseBracket[0]);
		} else if (this._lookahead.class === oLiteralNames.opCloseBracket[0]) {
			this.EatToken(oLiteralNames.opCloseBracket[0]);
		}

		oBase = this.GetElementLiteral();
		return {
			type: oLiteralNames.preScriptLiteral[1],
			base: oBase,
			down: oFirstSoOperand,
			up: oSecondSoOperand,
		}
	};
	CUnicodeParser.prototype.IsPreScriptLiteral = function () {
		return (this._lookahead.data === "_" || this._lookahead.data === "^")
	};
	CUnicodeParser.prototype.GetScriptBaseLiteral = function () {
		if (this.IsWordLiteral()) {
			let token = this.GetWordLiteral();
			if (this._lookahead.class === oLiteralNames.numberLiteral[0]) {
				token.nASCII = this.GetASCIILiteral();
			}
			return token;
		}
		else if (this._lookahead.class === oLiteralNames.anMathLiteral[0]) {
			return this.GetAnMathLiteral();
		}
		else if (this.IsNumberLiteral()) {
			return this.GetNumberLiteral();
		}
		else if (this.isOtherLiteral()) {
			return this.otherLiteral();
		}
		else if (this.IsExpBracketLiteral()) {
			return this.GetExpBracketLiteral();
		}
		else if (this._lookahead.class === oLiteralNames.opBuildupLiteral[0]) {
			return this.GetOpNaryLiteral();
		}
		else if (this.IsAnOtherLiteral()) {
			return this.GetAnOtherLiteral();
		}
		else if (this._lookahead.class === oLiteralNames.charLiteral[0]) {
			return this.GetCharLiteral();
		}
		else if (this.IsSqrtLiteral()) {
			return this.GetSqrtLiteral();
		}
		else if (this.IsCubertLiteral()) {
			return this.GetCubertLiteral();
		}
		else if (this.IsFourthrtLiteral()) {
			return this.GetFourthrtLiteral();
		}
		else if (this.IsNthrtLiteral()) {
			return this.GetNthrtLiteral();
		}
	};
	// CUnicodeParser.prototype.IsScriptBaseLiteral = function () {
	// 	return (
	// 		this.IsWordLiteral() ||
	// 		this.IsNumberLiteral() ||
	// 		this.isOtherLiteral() ||
	// 		this.IsExpBracketLiteral() ||
	// 		this._lookahead.class === "anOther" ||
	// 		this._lookahead.class === oLiteralNames.opNaryLiteral[0] ||
	// 		this._lookahead.class === "┬" ||
	// 		this._lookahead.class === "┴" ||
	// 		this._lookahead.class === oLiteralNames.charLiteral[0] ||
	// 		this._lookahead.class === oLiteralNames.anMathLiteral[0]
	// 	);
	// };
	CUnicodeParser.prototype.GetScriptSpecialContent = function (base) {
		let oFirstSoOperand = [],
			oSecondSoOperand = [];

		const ProceedScript = function (context) {
			while (context.IsScriptSpecialContent()) {
				if (context._lookahead.class === oLiteralNames.specialScriptNumberLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptNumberLiteral, true);
					oFirstSoOperand.push(oSpecial);
				}
				if (context._lookahead.class === oLiteralNames.specialScriptCharLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptCharLiteral, true);
					oFirstSoOperand.push(oSpecial);
				}
				if (context._lookahead.class === oLiteralNames.specialScriptBracketLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptBracketLiteral, true)
					oFirstSoOperand.push(oSpecial);
				}
				if (context._lookahead.class === oLiteralNames.specialScriptOperatorLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialScriptOperatorLiteral, true)
					oFirstSoOperand.push(oSpecial);
				}
			}
		};
		const ProceedIndex = function (context) {
			while (context.IsIndexSpecialContent()) {
				if (context._lookahead.class === oLiteralNames.specialIndexNumberLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexNumberLiteral, true);
					oSecondSoOperand.push(oSpecial);
				}
				if (context._lookahead.class === oLiteralNames.specialIndexCharLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexCharLiteral, true);
					oSecondSoOperand.push(oSpecial);
				}
				if (context._lookahead.class === oLiteralNames.specialIndexBracketLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexBracketLiteral, true)
					oSecondSoOperand.push(oSpecial);
				}
				if (context._lookahead.class === oLiteralNames.specialIndexOperatorLiteral[0]) {
					let oSpecial = context.ReadTokensWhileEnd(oLiteralNames.specialIndexOperatorLiteral, true)
					oSecondSoOperand.push(oSpecial);
				}
			}
		};

		if (this.IsScriptSpecialContent()) {
			ProceedScript(this);
			if (this.IsIndexSpecialContent()) {
				ProceedIndex(this);
			}
		}
		else if (this.IsIndexSpecialContent()) {
			ProceedIndex(this);
			if (this.IsScriptSpecialContent()) {
				ProceedScript(this);
			}
		}

		return {
			type: oLiteralNames.subSupLiteral[num],
			base,
			down: oSecondSoOperand,
			up: oFirstSoOperand,
		};
	}
	CUnicodeParser.prototype.IsSpecialContent = function () {
		return (
			this._lookahead.class === oLiteralNames.specialScriptNumberLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialScriptCharLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialScriptBracketLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialScriptOperatorLiteral[0] ||

			this._lookahead.class === oLiteralNames.specialIndexNumberLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialIndexCharLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialIndexBracketLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialIndexOperatorLiteral[0]

		);
	};
	CUnicodeParser.prototype.IsScriptSpecialContent = function () {
		return (
			this._lookahead.class === oLiteralNames.specialScriptNumberLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialScriptCharLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialScriptBracketLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialScriptOperatorLiteral[0]
		);
	};
	CUnicodeParser.prototype.IsIndexSpecialContent = function () {
		return (
			this._lookahead.class === oLiteralNames.specialIndexNumberLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialIndexCharLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialIndexBracketLiteral[0] ||
			this._lookahead.class === oLiteralNames.specialIndexOperatorLiteral[0]
		);
	};
	CUnicodeParser.prototype.IsExpSubSupLiteral = function () {
		return (
			this.IsScriptStandardContentLiteral() ||
			this.IsScriptBelowOrAboveContent() ||
			this.IsSpecialContent()
		);
	};
	CUnicodeParser.prototype.GetExpSubSupLiteral = function (oBase) {
		if (undefined === oBase) {
			oBase = this.GetScriptBaseLiteral();
		}
		let oContent;

		// if (this.isPreScriptLiteral()) {
		// 	return this.preScriptLiteral();
		// }

		if (this.IsScriptStandardContentLiteral()) {
			oContent = this.GetScriptStandardContentLiteral(oBase);
		}
		else if (this.IsScriptBelowOrAboveContent()) {
			oContent = this.GetScriptBelowOrAboveContent(oBase);
		}
		else if (this.IsSpecialContent()) {
			oContent = this.GetScriptSpecialContent(oBase);
		}

		if (this._lookahead.class === "▒") {
			if (oBase.type === oLiteralNames.opBuildupLiteral[1] || oBase.type === oLiteralNames.opNaryLiteral[1]) {
				this.EatToken("▒");
				let oThirdSoOperand = this.GetSoOperandLiteral();
				return {
					type: oLiteralNames.subSupLiteral[num],
					base: oBase,
					down: oContent.down,
					up: oContent.up,
					content: oThirdSoOperand,
				};
			}
		}
		else {
			return oContent;
		}
	};
	//TODO remove code repeat
	CUnicodeParser.prototype.GetScriptStandardContentLiteral = function (oBase) {
		let oFirstElement;
		let oSecondElement;

		if (this._lookahead.data === "_") {
			this.EatToken(oLiteralNames.opBuildupLiteral[0]);

			if (this.IsSoOperandLiteral()) {
				oFirstElement = (oBase && oBase.type === oLiteralNames.opNaryLiteral[1])
					? this.GetSoOperandLiteral("custom")
					: this.GetSoOperandLiteral("_");

				// Get second element
				if (this._lookahead.data === "^" && !this.isOneSubSup) {
					this.EatToken(oLiteralNames.opBuildupLiteral[0]);

					if (this.IsSoOperandLiteral()) {
						oSecondElement = this.GetSoOperandLiteral("^");
						return {
							type: oLiteralNames.subSupLiteral[0],
							base: oBase,
							down: oFirstElement,
							up: oSecondElement,
						};
					}
				}
				return {
					type: "expSubscript",
					base: oBase,
					down: oFirstElement,
				};
			}
		}
		else if (this._lookahead.data === "^") {
			this.EatToken(oLiteralNames.opBuildupLiteral[0]);

			if (this.IsSoOperandLiteral()) {
				oSecondElement = (oBase && oBase.type === oLiteralNames.opNaryLiteral[1])
					? this.GetSoOperandLiteral("custom")
					: this.GetSoOperandLiteral("^");

				if (this._lookahead.data === "_") {
					this.EatToken(oLiteralNames.opBuildupLiteral[0]);
					if (this.IsSoOperandLiteral()) {
						oFirstElement = this.GetSoOperandLiteral("_");
						return {
							type: oLiteralNames.subSupLiteral[num],
							oBase,
							down: oFirstElement,
							up: oSecondElement,
						};
					}
				}
				return {
					type: "expSuperscript",
					base: oBase,
					up: oSecondElement,
				};
			}
		}
	};
	CUnicodeParser.prototype.IsScriptStandardContentLiteral = function () {
		return this._lookahead.data === "_" || this._lookahead.data === "^";
	};
	CUnicodeParser.prototype.GetScriptBelowOrAboveContent = function (base) {
		let oBelowLiteral,
			oAboveLiteral,
			strType;

		if (this._lookahead.class === "┬" || this._lookahead.class === "┴") {
			strType = this.EatToken(this._lookahead.class).data;
			if (strType === "┬") {
				oBelowLiteral = this.GetSoOperandLiteral();
			}
			else if (strType === "┴") {
				oAboveLiteral = this.GetSoOperandLiteral();
			}
		}

		return {
			type: "expAbove",
			base,
			up: oAboveLiteral,
			down: oBelowLiteral
		};
	};
	CUnicodeParser.prototype.IsScriptBelowOrAboveContent = function () {
		return this._lookahead.class === "┬" || this._lookahead.class === "┴";
	};
	CUnicodeParser.prototype.GetFractionLiteral = function (oNumerator) {
		if (undefined === oNumerator) {
			oNumerator = this.GetOperandLiteral();
		}
		if (this._lookahead.class === oLiteralNames.overLiteral[0]) {
			const strOpOver = this.EatToken(oLiteralNames.overLiteral[0]).data;
			let strLiteralType = (strOpOver === "¦" || strOpOver === "⒞")
				? oLiteralNames.binomLiteral[num]
				: oLiteralNames.fractionLiteral[num];

			if (this.IsOperandLiteral()) {
				let oOperand = this.GetFractionLiteral();
				return {
					type: strLiteralType,
					up: oNumerator,
					opOver: strOpOver,
					down: oOperand,
				};
			}
		}
		else {
			return oNumerator;
		}
	};
	CUnicodeParser.prototype.IsFractionLiteral = function () {
		return this.IsOperandLiteral();
	};
	CUnicodeParser.prototype.otherLiteral = function () {
		return this.GetCharLiteral();
	};
	CUnicodeParser.prototype.isOtherLiteral = function () {
		return this._lookahead.class === oLiteralNames.charLiteral[0];
	};
	CUnicodeParser.prototype.GetOperatorLiteral = function () {
		const oOperator = this.EatToken(oLiteralNames.operatorLiteral[0]);
		return {
			type: oLiteralNames.operatorLiteral[1],
			value: oOperator.data,
		};
	};
	CUnicodeParser.prototype.GetASCIILiteral = function () {
		return this.ReadTokensWhileEnd(oLiteralNames.numberLiteral, false)
	};
	CUnicodeParser.prototype.GetCharLiteral = function () {
		return this.ReadTokensWhileEnd(oLiteralNames.charLiteral, false)
	};
	CUnicodeParser.prototype.GetAnMathLiteral = function () {
		const oAnMathLiteral = this.EatToken(oLiteralNames.anMathLiteral[0]);
		return {
			type: oLiteralNames.anMathLiteral[1],
			value: oAnMathLiteral.data,
		};
	};
	CUnicodeParser.prototype.IsAnMathLiteral = function () {
		return this._lookahead.class === oLiteralNames.anMathLiteral[0];
	};
	CUnicodeParser.prototype.GetAnOtherLiteral = function () {
		if (this._lookahead.class === oLiteralNames.otherLiteral[0]) {
			return this.ReadTokensWhileEnd(oLiteralNames.otherLiteral, false)
		}
		else if (this._lookahead.class === oLiteralNames.charLiteral[0]) {
			return this.GetCharLiteral();
		}
		else if (this._lookahead.class === oLiteralNames.numberLiteral[0]) {
			return this.GetNumberLiteral();
		}
	};
	CUnicodeParser.prototype.IsAnOtherLiteral = function () {
		return (
			this._lookahead.class === oLiteralNames.otherLiteral[0] ||
			this._lookahead.class === oLiteralNames.charLiteral[0] ||
			this._lookahead.class === oLiteralNames.numberLiteral[0]
		);
	};
	CUnicodeParser.prototype.GetAnLiteral = function () {
		if (this.IsAnOtherLiteral()) {
			return this.GetAnOtherLiteral();
		}
		return this.GetAnMathLiteral();
	};
	CUnicodeParser.prototype.IsAnLiteral = function () {
		return this.IsAnOtherLiteral() || this.IsAnMathLiteral();
	};
	CUnicodeParser.prototype.GetDiacriticLiteral = function () {
		const oDiacritic = this.EatToken(oLiteralNames.diacriticLiteral[0]);
		return {
			type: oLiteralNames.diacriticLiteral[1],
			value: oDiacritic.data,
		};
	};
	CUnicodeParser.prototype.GetDiacriticBaseLiteral = function () {
		let oDiacriticBase;
		const strDiacriticBaseLiteral = oLiteralNames.diacriticBaseLiteral[1];

		if (this.IsAnLiteral()) {
			oDiacriticBase = this.GetAnLiteral();
			return {
				type: strDiacriticBaseLiteral,
				data: oDiacriticBase,
				isAn: true,
			};
		}
		else if (this._lookahead.class === "nASCII") {
			oDiacriticBase = this.GetASCIILiteral();
			return {
				type: strDiacriticBaseLiteral,
				data: oDiacriticBase,
			};
		}
		else if (this._lookahead.class === "(") {
			this.EatToken("(");
			oDiacriticBase = this.GetExpLiteral();
			this.EatToken(")");
			return {
				type: strDiacriticBaseLiteral,
				data: oDiacriticBase,
			};
		}
	};
	CUnicodeParser.prototype.IsDiacriticBaseLiteral = function () {
		return (
			this.IsAnLiteral() ||
			this._lookahead.class === oLiteralNames.numberLiteral[0] ||
			this._lookahead.class === "("
		);
	};
	CUnicodeParser.prototype.GetDiacriticsLiteral = function () {
		const arrDiacriticList = [];
		while (this._lookahead.class === oLiteralNames.diacriticLiteral[0]) {
			arrDiacriticList.push(this.GetDiacriticLiteral());
		}
		return this.GetContentOfLiteral(arrDiacriticList);
	};
	CUnicodeParser.prototype.IsDiacriticsLiteral = function () {
		return this._lookahead.class === oLiteralNames.diacriticLiteral[0];
	};
	CUnicodeParser.prototype.GetAtomLiteral = function () {
		const oAtom = this.GetDiacriticBaseLiteral();
		if (oAtom.isAn) {
			return oAtom.data
		}
		return oAtom;
	};
	CUnicodeParser.prototype.IsAtomLiteral = function () {
		return this.IsAnLiteral() || this.IsDiacriticBaseLiteral();
	};
	CUnicodeParser.prototype.GetAtomsLiteral = function () {
		const arrAtomsList = [];
		while (this.IsAtomLiteral()) {
			arrAtomsList.push(this.GetAtomLiteral());
		}
		return this.GetContentOfLiteral(arrAtomsList)
	};
	CUnicodeParser.prototype.IsAtomsLiteral = function () {
		return this.IsAtomLiteral();
	};
	CUnicodeParser.prototype.GetEntityLiteral = function () {
		if (this.IsAtomsLiteral()) {
			return this.GetAtomsLiteral();
		}
		else if (this.IsExpBracketLiteral()) {
			return this.GetExpBracketLiteral();
		}
		else if (this._lookahead.class === oLiteralNames.operatorLiteral[0]) {
			return this.GetOperatorLiteral();
		}
		else if (this.IsNumberLiteral()) {
			return this.GetNumberLiteral();
		}
		else if (this.IsOpNaryLiteral()) {
			return this.GetOpNaryLiteral();
		}
		else if (this.IsTextLiteral()) {
			return this.GetTextLiteral()
		}
	};
	CUnicodeParser.prototype.IsEntityLiteral = function () {
		return (
			this.IsAtomsLiteral() ||
			this.IsExpBracketLiteral() ||
			this.IsNumberLiteral() ||
			this.IsOpNaryLiteral() ||
			this.IsTextLiteral()
		);
	};
	//TODO preScript get only one element (_2^3)45 <=> prescript4, 5
	CUnicodeParser.prototype.GetFactorLiteral = function () {
		if (this.IsEntityLiteral() && !this.IsFunctionLiteral()) {
			let oEntity = this.GetEntityLiteral();
			if (this._lookahead.class === "!") {
				this.EatToken("!");
				return {
					type: "factorialLiteral",
					exp: oEntity,
				};
			}
			else if (this._lookahead.class === "!!") {
				this.EatToken("!!");
				return {
					type: "factorialLiteral",
					exp: oEntity,
				};
			}
			else if (this.IsDiacriticsLiteral()) {
				const oDiacritic = this.GetDiacriticsLiteral().value;
				return {
					type: oLiteralNames.diacriticLiteral[num],
					base: oEntity,
					value: oDiacritic,
				};
			}

			return oEntity;
		}
		else if (this.IsFunctionLiteral()) {
			return this.GetFunctionLiteral();
		}
		else if (this.IsExpSubSupLiteral()) {
			return this.GetExpSubSupLiteral();
		}
	};
	CUnicodeParser.prototype.IsFactorLiteral = function () {
		return this.IsEntityLiteral() || this.IsFunctionLiteral()
	};
	CUnicodeParser.prototype.GetOperandLiteral = function (isNoSubSup) {
		const arrFactorList = [];
		if (undefined === isNoSubSup) {
			isNoSubSup = false;
		}

		while (this.IsFactorLiteral() && !this.IsExpSubSupLiteral()) {
			arrFactorList.push(this.GetFactorLiteral());
		}

		//if next token "_" or "^" proceed as index/degree
		if (this._lookahead.data === isNoSubSup || !isNoSubSup && this.IsScriptStandardContentLiteral()) {
			return this.GetExpSubSupLiteral(arrFactorList[0]);
		}
		//if next token "┬" or "┴" proceed as below/above
		else if (this._lookahead.data === isNoSubSup || !isNoSubSup && this.IsScriptBelowOrAboveContent()) {
			return this.GetScriptBelowOrAboveContent(arrFactorList[0]);
		}
		//if next token like ⁶⁷⁸⁹ or ₂₃₄ proceed as special degree/index
		else if (this._lookahead.data === isNoSubSup || !isNoSubSup && this.IsSpecialContent()) {
			return this.GetScriptSpecialContent(arrFactorList[0]);
		}

		return this.GetContentOfLiteral(arrFactorList);
	};
	CUnicodeParser.prototype.IsOperandLiteral = function () {
		return this.IsFactorLiteral();
	};
	CUnicodeParser.prototype.IsRowLiteral = function () {
		return this.IsExpLiteral();
	};
	CUnicodeParser.prototype.GetRowLiteral = function () {
		let arrRow = [];
		while (this.IsExpLiteral() || this._lookahead.class === "&") {
			if (this._lookahead.class === "&") {
				this.EatToken("&");
			}
			else {
				arrRow.push(this.GetExpLiteral());
			}
		}
		return {
			type: "rowLiteral",
			value: arrRow,
		};
	};
	CUnicodeParser.prototype.GetRowsLiteral = function () {
		let arrRows = [];
		while (this.IsRowLiteral() || this._lookahead.class === "@") {
			if (this._lookahead.class === "@") {
				this.EatToken("@");
			}
			else {
				arrRows.push(this.GetRowLiteral());
			}
		}
		return {
			type: "rowsLiteral",
			value: arrRows,
		};
	};
	CUnicodeParser.prototype.GetArrayLiteral = function () {
		this.EatToken("\\array(");
		const oExp = this.GetRowsLiteral();
		if (this._lookahead.data === ")") {
			this.EatToken(oLiteralNames.opCloseBracket[0]);
			return {
				type: "arrayLiteral",
				value: oExp,
			};
		}
	};
	CUnicodeParser.prototype.IsArrayLiteral = function () {
		return this._lookahead.class === "\\array(";
	};
	CUnicodeParser.prototype.IsElementLiteral = function () {
		return (
			this.IsFractionLiteral() ||
			this.IsOperandLiteral() ||
			this.IsArrayLiteral() ||
			this._lookahead.class === oLiteralNames.spaceLiteral[0]
		);
	};
	CUnicodeParser.prototype.GetElementLiteral = function () {
		if (this.IsArrayLiteral()) {
			return this.GetArrayLiteral();
		}

		if (this._lookahead.class === oLiteralNames.spaceLiteral[0]) {
			let oSpace = this.EatToken(oLiteralNames.spaceLiteral[0]);
			return {
				type: oLiteralNames.spaceLiteral[num],
				value: oSpace.data,
			}
		}

		const oOperandLiteral = this.GetOperandLiteral();
		if (this._lookahead.class === oLiteralNames.overLiteral[0]) {
			return this.GetFractionLiteral(oOperandLiteral);
		}
		else {
			return oOperandLiteral;
		}
	};
	CUnicodeParser.prototype.IsExpLiteral = function () {
		return this.IsElementLiteral();
	};
	CUnicodeParser.prototype.GetExpLiteral = function () {
		const oExpLiteral = [];
		while (this.IsElementLiteral() || this.isOtherLiteral() || this._lookahead.class === oLiteralNames.operatorLiteral[0]) {
			if (this.IsElementLiteral()) {
				let oElement = this.GetElementLiteral();
				if (oElement !== null) {
					oExpLiteral.push(oElement);
				}
			}
			else if (this._lookahead.class === oLiteralNames.operatorLiteral[0]) {
				oExpLiteral.push(this.GetOperatorLiteral())
			}
			else if (this.isOtherLiteral()) {
				oExpLiteral.push(this.otherLiteral());
			}
		}

		if (oExpLiteral.length === 1) {
			return oExpLiteral[0];
		}
		return oExpLiteral;
	};

	/**
	 * Метод позволяет обрабатывать токены одного типа, пока они не прервутся другим типом токенов
	 *
	 * @param arrTypeOfLiteral {LiteralType}
	 * @param isSpecial {boolean}
	 * @return {array} Обработанные токены
	 * @constructor
	 */
	CUnicodeParser.prototype.ReadTokensWhileEnd = function (arrTypeOfLiteral, isSpecial) {
		let arrLiterals = [];
		//todo let isOne = this._

		// if (isOne) {
		// 	let oLiteral = {
		// 		type: arrTypeOfLiteral[1],
		// 		data: this.EatToken(arrTypeOfLiteral[0]).data,
		// 	};
		// 	arrLiterals.push(oLiteral);
		// }
		// else {

		let strLiteral = "";
		while (this._lookahead.class === arrTypeOfLiteral[0]) {
			if (isSpecial) {
				strLiteral += UnicodeSpecialScript[this.EatToken(arrTypeOfLiteral[0]).data];
			}
			else {
				strLiteral += this.EatToken(arrTypeOfLiteral[0]).data;
			}
		}
		arrLiterals.push({
			type: arrTypeOfLiteral[num],
			value: strLiteral,
		})
		//	}

		if (arrLiterals.length === 1) {
			return arrLiterals[0];
		}
		return arrLiterals
	};
	CUnicodeParser.prototype.EatToken = function (tokenType) {
		const token = this._lookahead;

		if (token === null) {
			throw new SyntaxError(
				`Unexpected end of input, expected: "${tokenType}"`
			);
		}

		if (token.class !== tokenType) {
			throw new SyntaxError(
				`Unexpected token: "${token.class}", expected: "${tokenType}"`
			);
		}
		this._lookahead = this.oTokenizer.GetNextToken();
		return token;
	};
	CUnicodeParser.prototype.GetContentOfLiteral = function (oContent) {
		if (Array.isArray(oContent)) {
			if (oContent.length === 1) {
				return oContent[0];
			}
			return oContent;
		}
		return oContent;
	}
	CUnicodeParser.prototype.SkipSpace = function () {
		while (this._lookahead.class === oLiteralNames.spaceLiteral[0]) {
			this.EatToken(oLiteralNames.spaceLiteral[0]);
		}
	}

	function CUnicodeConverter(str, oContext, isGetOnlyTokens) {
		if (undefined === str || null === str) {
			return
		}
		let oParser = new CUnicodeParser();
		const oTokens = oParser.Parse(str);

		if (!isGetOnlyTokens) {
			ConvertTokens(oTokens, oContext);
		} else {
			return oTokens;
		}
		return true
	}
	//--------------------------------------------------------export----------------------------------------------------
	window["AscCommonWord"] = window["AscCommonWord"] || {};
	window["AscCommonWord"].CUnicodeConverter = CUnicodeConverter;
})(window);