/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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

(/**
	 * @param {jQuery} $
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function ($, window, undefined) {

		/*
		 * Import
		 * -----------------------------------------------------------------------------
		 */
		var prot;
		var c_oAscBorderStyles = Asc.c_oAscBorderStyles;
		var c_oAscMaxCellOrCommentLength = Asc.c_oAscMaxCellOrCommentLength;
		var doc = window.document;
		var copyPasteUseBinary = true;
		var CopyPasteCorrectString = AscCommon.CopyPasteCorrectString;

		var c_oSpecialPasteProps = Asc.c_oSpecialPasteProps;

		var c_MaxStringLength = 536870888;
		var notSupportExternalReferenceFileFormat = {"csv": 1};

		function number2color(n) {
			if (typeof (n) === "string" && n.indexOf("rgb") > -1)
				return n;
			return "rgb(" + (n >> 16 & 0xFF) + "," + (n >> 8 & 0xFF) + "," + (n & 0xFF) + ")";
		}

		function CSpecialPasteProps() {
			this.cellStyle = null;
			this.val = null;
			this.numFormat = null;
			this.formula = null;
			this.font = null;
			this.alignVertical = null;
			this.alignHorizontal = null;
			this.fontSize = null;
			this.fontName = null;
			this.merge = null;
			this.borders = null;
			this.wrap = null;
			this.fill = null;
			this.angle = null;
			this.hyperlink = null;

			this.format = null;
			this.formatTable = null;

			this.images = null;

			this.width = null;
			this.transpose = null;

			this.comment = null;

			//for paste text as csv
			this.advancedOptions = null;

			this.operation = null;
			this.skipBlanks = null;

			this.init();
		}

		CSpecialPasteProps.prototype = {

			constructor: CSpecialPasteProps,

			init: function () {
				this.setBaseOptions();
			},

			setBaseOptions: function (onlyMain) {
				this.cellStyle = true;
				this.val = true;
				this.numFormat = true;
				this.formula = true;
				this.font = true;
				this.alignVertical = true;
				this.alignHorizontal = true;
				this.fontSize = true;
				this.fontName = true;
				this.merge = true;
				this.borders = true;
				this.wrap = true;
				this.fill = true;
				this.angle = true;
				this.hyperlink = true;

				this.format = true;
				this.formatTable = true;

				this.images = true;

				this.width = null;
				this.transpose = null;

				this.comment = true;
				this.property = null;

				this.advancedOptions = null;

				if (!onlyMain) {
					this.operation = null;
					this.skipBlanks = null;
				}
			},
			_clean: function () {
				this.cellStyle = null;
				this.val = null;
				this.numFormat = null;
				this.formula = null;
				this.font = null;
				this.alignVertical = null;
				this.alignHorizontal = null;
				this.fontSize = null;
				this.fontName = null;
				this.merge = null;
				this.borders = null;
				this.wrap = null;
				this.fill = null;
				this.angle = null;
				this.hyperlink = null;

				this.format = null;
				this.formatTable = null;

				this.images = null;

				this.width = null;
				this.transpose = null;

				this.comment = null;

				//for paste text as csv
				this.advancedOptions = null;

				//this.operation = null;
				this.skipBlanks = null;
			},
			asc_setProps: function (props) {
				this.setBaseOptions(true);
				this.property = props;
				switch (props) {
					case c_oSpecialPasteProps.paste: {
						break;
					}
					case c_oSpecialPasteProps.pasteOnlyFormula: {
						//только формулы(или значения)
						this._clean();
						this.formula = true;
						this.val = true;

						break;
					}
					case c_oSpecialPasteProps.formulaNumberFormat: {
						//только формулы(или значения) и числовой формат
						this._clean();
						this.formula = true;
						this.numFormat = true;
						this.val = true;

						break;
					}
					case c_oSpecialPasteProps.formulaAllFormatting: {
						//формулы и формат
						break;
					}
					case c_oSpecialPasteProps.formulaWithoutBorders: {
						//всё кроме бордеров
						this.borders = null;
						break;
					}
					case c_oSpecialPasteProps.formulaColumnWidth: {
						this.width = true;
						break;
					}
					case c_oSpecialPasteProps.mergeConditionalFormating: {
						break;
					}
					case c_oSpecialPasteProps.pasteOnlyValues: {
						//только значения(вместо формул также вставляются значения)
						this._clean();
						this.val = true;
						break;
					}
					case c_oSpecialPasteProps.valueNumberFormat: {
						this._clean();
						this.val = true;
						this.numFormat = true;
						break;
					}
					case c_oSpecialPasteProps.valueAllFormating: {
						//все кроме формул
						this.formula = null;
						this.formatTable = null;
						break;
					}
					case c_oSpecialPasteProps.pasteOnlyFormating: {
						this.formula = null;
						this.val = null;
						this.formatTable = null;
						break;
					}
					case c_oSpecialPasteProps.transpose: {
						this.transpose = true;
						break;
					}
					case c_oSpecialPasteProps.link: {
						this._clean();
						break;
					}
					case c_oSpecialPasteProps.picture: {
						break;
					}
					case c_oSpecialPasteProps.linkedPicture: {
						break;
					}
					case c_oSpecialPasteProps.sourceformatting: {
						break;
					}
					case c_oSpecialPasteProps.destinationFormatting: {
						//только значения(вместо формул также вставляются значения)
						this._clean();
						this.val = true;
						//картинки из word сохраняем в данной ситуации
						if (window['AscCommon'].g_specialPasteHelper.specialPasteData.pasteFromWord) {
							this.images = true;
						}

						break;
					}
					case c_oSpecialPasteProps.comments: {
						this._clean();
						this.comment = true;
						break;
					}
					case c_oSpecialPasteProps.columnWidth: {
						this._clean();
						this.width = true;
						break;
					}
				}
			},
			asc_getProps: function () {
				return this.property;
			},
			asc_setAdvancedOptions: function (props) {
				this.advancedOptions = props;
			},
			asc_getAdvancedOptions: function () {
				return this.advancedOptions;
			},
			asc_setOperation: function (val) {
				this.operation = val;
			},
			asc_getOperation: function () {
				return this.operation;
			},
			asc_setTranspose: function (val) {
				this.transpose = val;
			},
			asc_setSkipBlanks: function (val) {
				this.skipBlanks = val;
			},
			asc_getTableAllowed: function () {
				return this.formatTable;
			}
		};

		/** @constructor */
		function Clipboard() {
			this.copyProcessor = new CopyProcessorExcel();
			this.pasteProcessor = new PasteProcessorExcel();

			return this;
		}

		Clipboard.prototype.checkCopyToClipboard = function (ws, _clipboard, _formats) {
			var _data = null;
			var wb = window["Asc"]["editor"].wb;

			window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
			ws.handlers.trigger("cleanCutData", true);
			wb.handlers.trigger("cleanCopyData", true);

			if (wb.getCellEditMode())//text in cell
			{
				//only TEXT
				var fragments = wb.cellEditor.copySelection();

				if (null !== fragments) {
					_data = AscCommonExcel.getFragmentsText(fragments);
				}

				if (null !== _data) {
					_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _data);
				}
			} else {
				if (!this.copyProcessor.canCopy(ws)) {
					var selectedDrawings = ws.objectRender.getSelectedGraphicObjects();
					if (0 === selectedDrawings.length) {
						ws.handlers.trigger("onErrorEvent", Asc.c_oAscError.ID.CopyMultiselectAreaError,
							Asc.c_oAscError.Level.NoCritical);
						return;
					}
				}

				//ignore hidden rows
				var activeCell = ws.model.selectionRange.activeCell.clone();
				var activeRange, selectionRange;
				if (1 !== ws.model.selectionRange.ranges.length) {
					selectionRange = new AscCommonExcel.MultiplyRange(ws.model.selectionRange.ranges).getUnionRange();
					activeRange = ws.model.getRange3(selectionRange.r1, selectionRange.c1, selectionRange.r2, selectionRange.c2);
				} else {
					activeRange = ws.getSelectedRange();
					selectionRange = ws.model.selectionRange.getLast();
				}


				//TODO игнорировать нужно и формулы и скрытые строчки в случае, если селект их задевает + стандартные условия в bIsExcludeHiddenRows
				if (ws.model.autoFilters.bIsExcludeHiddenRows(selectionRange, activeCell, true)) {
					ws.model.excludeHiddenRows(true);
					ws.model.ignoreWriteFormulas(true);
				}


				// IMAGE
				if (_formats & AscCommon.c_oAscClipboardDataFormat.Image) {
					let imageData = this.drawSelectedArea(ws, true);
					if (imageData) {
						_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Image, imageData);
					}
				}

				//TEXT
				if (AscCommon.c_oAscClipboardDataFormat.Text & _formats) {
					_data = this.copyProcessor.getText(activeRange, ws);

					if (null !== _data) {
						_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _data);
					}
				}
				//HTML
				if (AscCommon.c_oAscClipboardDataFormat.Html & _formats) {
					_data = this.copyProcessor.getHtml(activeRange, ws);

					if (null !== _data && "" !== _data.html) {
						_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _data.html)
					}
				}
				//INTERNAL
				if (AscCommon.c_oAscClipboardDataFormat.Internal & _formats) {
					if (window["NATIVE_EDITOR_ENJINE"]) {
						_data = this.copyProcessor.getBinaryForMobile();
					} else {
						if (_data && _data.base64) {
							_data = _data.base64;
						} else {
							_data = this.copyProcessor.getBinaryForCopy(ws.model, ws.objectRender);
						}
					}

					if (null !== _data) {
						_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Internal, _data);
					}
				}

				ws.model.excludeHiddenRows(false);
				ws.model.ignoreWriteFormulas(false);
			}

			if (ws && ws.workbook && !ws.workbook.getCellEditMode()) {
				if (AscCommon.g_clipboardBase.bCut) {
					//в данном случае не вырезаем, а записываем
					if (!ws.isNeedSelectionCut() && false === ws.isMultiSelect()) {
						ws.workbook.cutIdSheet = ws.model.Id;
						ws.setCutRange([ws.model.selectionRange.getLast()]);
					}
				} else if (!ws.objectRender.selectedGraphicObjectsExists()) {
					ws.setCutRange(ws.model.selectionRange.ranges);
				}
			}
		};

		Clipboard.prototype.pasteData = function (ws, _format, data1, data2, text_data, bIsSpecialPaste, doNotShowButton, isPasteAll, callback) {
			var t = this;
			var wb = window["Asc"]["editor"].wb;
			//if open range dialog - return
			if (ws.getSelectionDialogMode() && (!wb.getCellEditMode() || wb.isWizardMode)) {
				return;
			}
			var cellEditor = wb.cellEditor;
			t.pasteProcessor.clean();

			if (!window['AscCommon'].g_specialPasteHelper.specialPasteStart) {
				window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
			}
			window['AscCommon'].g_specialPasteHelper.Paste_Process_Start(doNotShowButton);

			if (!bIsSpecialPaste && !wb.getCellEditMode()) {
				window['AscCommon'].g_specialPasteHelper.specialPasteData.activeRange = ws.model.selectionRange.clone(ws.model);
				window['AscCommon'].g_specialPasteHelper.specialPasteData.pasteFromWord = false;
			}

			switch (_format) {
				case AscCommon.c_oAscClipboardDataFormat.HtmlElement: {
					if (wb.getCellEditMode()) {
						//fragments = пока только для плагина вставка символов
						var fragments;
						if (window['AscCommon'].g_clipboardBase.bSaveFormat) {
							//проверяем иероглифы внутри
							fragments = this.pasteProcessor._getFragmentsFromHtml(data1);
						}
						if (fragments) {
							var pasteFragments = fragments.fragments;
							var newFonts = fragments.fonts;
							ws._loadFonts(newFonts, function () {
								cellEditor.paste(pasteFragments);
								window['AscCommon'].g_specialPasteHelper.Paste_Process_End();

								//TODO пересмотреть! по возможности вызывать из меню!
								//при использовании рекдактора ячейки в качестве редактора HF
								//в функции onLongActionEnd(ф-я меню) не вызывается asc_enableKeyEvents(true)
								//из-за этого enableKeyEvents остаётся выставленным в false
								//поэтому приходится вызывать здесь, после того, как пройдет загрузка шрифтов

								if (cellEditor.getMenuEditorMode()) {
									window["Asc"]["editor"].asc_enableKeyEvents(true);
								}
							});

						} else {
							this._pasteTextInCellEditor(text_data || data1.innerText);
						}
					} else {
						t.pasteProcessor.editorPasteExec(ws, data1, null, callback);
					}

					break;
				}
				case AscCommon.c_oAscClipboardDataFormat.Internal: {
					if (wb.getCellEditMode()) {
						if (null !== text_data && undefined !== text_data) {
							this._pasteTextInCellEditor(text_data);
						} else {
							this._pasteTextInCellEditor(this.pasteProcessor.pasteFromBinary(ws, data1, true));
						}
					} else {
						t.pasteProcessor.pasteFromBinary(ws, data1, null, isPasteAll);
					}

					break;
				}
				case AscCommon.c_oAscClipboardDataFormat.Text: {
					if (wb.getCellEditMode()) {
						this._pasteTextInCellEditor(data1);
					} else {
						//не показываем иконку с/в если вставляется только текст
						//window['AscCommon'].g_specialPasteHelper.Special_Paste_Hide_Button();
						t.pasteProcessor.pasteTextOnSheet(ws, data1);
					}
					wb.handlers.trigger("cleanCopyData", true);
					break;
				}
			}

			//ws.handlers.trigger("cleanCutData", true);

			if (!bIsSpecialPaste) {
				window['AscCommon'].g_specialPasteHelper.specialPasteData._format = _format;
				window['AscCommon'].g_specialPasteHelper.specialPasteData.data1 = data1;
				window['AscCommon'].g_specialPasteHelper.specialPasteData.data2 = data2;
				window['AscCommon'].g_specialPasteHelper.specialPasteData.text_data = text_data;
			}
		};

		Clipboard.prototype._pasteTextInCellEditor = function (text) {
			if (!text) {
				return;
			}
			var editor = window["Asc"]["editor"];
			AscFonts.FontPickerByCharacter.getFontsByString(text);
			editor._loadFonts([], function () {
				editor.wb.skipHelpSelector = true;
				editor.wb.cellEditor.pasteText(text);
				AscCommon.g_specialPasteHelper.Paste_Process_End();
				editor.wb.skipHelpSelector = false;

				if (editor.wb.cellEditor.getMenuEditorMode()) {
					editor.asc_enableKeyEvents(true);
				}
			});
		};

		Clipboard.prototype.drawSelectedArea = function (ws, opt_get_bytes) {
			let activeRange = ws.model.selectionRange.getLast();

			let base64;
			//ws.workbook._executeWithoutZoom(function () {
				let ctx = ws.workbook.printForCopyPaste(ws, activeRange);
				if (ctx && ctx.canvas) {
					base64 = ctx.canvas.toDataURL("image/png");
				}
			//})

			if (opt_get_bytes && base64) {
				// Get base64 data without header
				let base64Data = base64.split(',')[1];

				// Convert base64 to binary
				let byteCharacters = atob(base64Data);
				let byteArrays = [];

				for (let offset = 0; offset < byteCharacters.length; offset += 512) {
					let slice = byteCharacters.slice(offset, offset + 512);
					let byteNumbers = new Array(slice.length);

					for (let i = 0; i < slice.length; i++) {
						byteNumbers[i] = slice.charCodeAt(i);
					}

					let byteArray = new Uint8Array(byteNumbers);
					byteArrays.push(byteArray);
				}

				return byteArrays;
			}

			// Get image data
			return base64;
		};



		function CopyProcessorExcel() {

		}

		CopyProcessorExcel.prototype = {

			constructor: CopyProcessorExcel,

			getHtml: function (range, worksheet) {
				var t = this;
				var sBase64 = null;

				var objectRender = worksheet.objectRender;
				var isIntoShape = objectRender.controller.getTargetDocContent();

				History.TurnOff();
				//use binary strings
				if (copyPasteUseBinary) {
					sBase64 = this.getBinaryForCopy(worksheet.model, worksheet.objectRender);
				}
				History.TurnOn();

				var innerHtml = t._generateHtml(range, worksheet, isIntoShape, sBase64);
				if (innerHtml === false) {
					return null;
				}

				return {base64: sBase64, html: innerHtml};
			},

			getBinaryForCopy: function (wsModel, objectRender, activeRange, selectAll, ignoreCopyPaste, allSelectedTabs) {
				let isIntoShape = !selectAll && objectRender && objectRender.controller ? objectRender.controller.getTargetDocContent() : null;
				let sBase64 = null;

				if (isIntoShape) {
					//в данному случае пишем бинарник с меткой pptData - с префиксом xlsData отдельно параграфы записать не получится
					sBase64 = this._getBinaryShapeContent(isIntoShape);
				} else {

					AscFormat.ExecuteNoHistory(function () {
						pptx_content_writer.Start_UseFullUrl();

						let unselectedIndexes = [];
						if (selectAll) {
							activeRange = new Asc.Range(0, 0, AscCommon.gc_nMaxCol - 1, AscCommon.gc_nMaxRow - 1);

							for (let i = 0; i < wsModel.Drawings.length; i++) {
								if (!wsModel.Drawings[i].graphicObject.selected) {
									unselectedIndexes[i] = true;
									wsModel.Drawings[i].graphicObject.selected = true;
								}
							}
						}

						let selectionRange;
						if (activeRange) {
							selectionRange = activeRange;
						} else {
							if (1 !== wsModel.selectionRange.ranges.length) {
								selectionRange = new AscCommonExcel.MultiplyRange(wsModel.selectionRange.ranges).getUnionRange();
							} else {
								selectionRange = wsModel.selectionRange.getLast();
							}
						}
						let maxRowCol = this._getRangeMaxRowCol(wsModel, selectionRange);
						if (null !== maxRowCol) {
							if (maxRowCol.col < selectionRange.c1) {
								maxRowCol.col = selectionRange.c1;
							}
							if (maxRowCol.row < selectionRange.r1) {
								maxRowCol.row = selectionRange.r1;
							}
							selectionRange = new Asc.Range(selectionRange.c1, selectionRange.r1, maxRowCol.col, maxRowCol.row);
						}

						let wb = wsModel.workbook;
						let isNullCore = false;
						if (!wb.Core) {
							isNullCore = true;
							wb.Core = new window['AscCommon'].CCore();
						}

						//подменяем identifier/creator для того, чтобы протащить id
						let oldCreator = wb.Core.creator;
						let oldIdentifier = wb.Core.identifier;
						let oldLanguage = wb.Core.language;
						let oldTitle = wb.Core.title;
						let oldCategory = wb.Core.category;
						let oldContentStatus = wb.Core.contentStatus;

						wb.Core.creator = wb.oApi && wb.oApi.CoAuthoringApi ? wb.oApi.CoAuthoringApi.getUserConnectionId() : null;
						wb.Core.identifier = wb.oApi && wb.oApi.DocInfo ? wb.oApi.DocInfo.Id : null;

						//для внешних данных необходимо протащить docInfo->ReferenceData
						//пока беру данные поля, поскольку для копипаста они не используются. по названию не особо совпадают - пересмотреть
						let isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
						if (isLocalDesktop && window["AscDesktopEditor"]["LocalFileGetSaved"]()) {
							wb.Core.contentStatus = window["AscDesktopEditor"]["LocalFileGetSourcePath"]();
							wb.Core.category = null;
						} else if (wb.oApi && wb.oApi.DocInfo && !notSupportExternalReferenceFileFormat[wb.oApi.DocInfo.Format]) {
							wb.Core.contentStatus = wb.oApi.DocInfo.ReferenceData ? wb.oApi.DocInfo.ReferenceData["fileKey"] : null;
							wb.Core.category = wb.oApi.DocInfo.ReferenceData ? wb.oApi.DocInfo.ReferenceData["instanceId"] : null;
						}

						wb.Core.title = wb.oApi && wb.oApi.DocInfo && wb.oApi.DocInfo.Title ? wb.oApi.DocInfo.Title : null;

						//записываю изображение выделенного фрагмента. пока только для изоюражений
						//выбрал для этого поле subject
						let oldSubject = wb.Core.subject;
						let objectRender = Asc.editor.wb.getWorksheet().objectRender;
						let _imgProperty = objectRender && objectRender.controller && objectRender.controller.getSelectionImage();
						if (_imgProperty) {
							let _base64 = _imgProperty.asc_getImageUrl();
							let _width = _imgProperty.Width;
							let _height = _imgProperty.Height;
							wb.Core.subject = _width + ";" + _height + ";" + _base64;
						}

						//так же необходимо протащить локаль, для этого использую поля language
						//и записываю туда номер локали, предварительно конвертируя его в строку
						//пока буду использовать его только при вставке в документы, а в документах устанавливать -> AscCommon.setCurrentCultureInfo(val)
						let locale = wb.oApi ? wb.oApi.asc_getLocale() : null;
						wb.Core.language = (undefined !== locale && null !== locale) ? locale.toString() : null;


						//WRITE
						let oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(wb, !ignoreCopyPaste ? selectionRange : false);
						sBase64 = "xslData;" + oBinaryFileWriter.Write();
						pptx_content_writer.End_UseFullUrl();


						if (selectAll) {
							for (let i in unselectedIndexes) {
								wsModel.Drawings[i].graphicObject.selected = false;
							}
						}

						if (isNullCore) {
							wb.Core = null;
						} else {
							wb.Core.creator = oldCreator;
							wb.Core.identifier = oldIdentifier;
							wb.Core.language = oldLanguage;
							wb.Core.subject = oldSubject;
							wb.Core.title = oldTitle;
							wb.Core.category = oldCategory;
							wb.Core.contentStatus = oldContentStatus;
						}
					}, this, []);
				}

				return sBase64;
			},

			_getRangeMaxRowCol: function (worksheet, selectionRange, range3) {
				var res = null;

				//при выделенных столбцах ищем последнюю непустую ячейку
				//ограничиваемся последней пустой строкой/столбцом
				var type = selectionRange.getType();
				var oType = Asc.c_oAscSelectionType;
				if (type === oType.RangeCol || type === oType.RangeRow || type === oType.RangeMax) {
					if (!range3) {
						range3 = worksheet.getRange3(selectionRange.r1, selectionRange.c1, selectionRange.r2, selectionRange.c2);
					}

					//нужно вычислить последнюю ячейку в столбце, где есть данные
					var maxRow = 0;
					var maxCol = 0;
					range3._foreachNoEmpty(function (cell) {
						if (cell.nCol > maxCol) {
							maxCol = cell.nCol;
						}
						if (cell.nRow > maxRow) {
							maxRow = cell.nRow;
						}
					});

					res = {col: maxCol, row: maxRow};
				}

				return res;
			},

			_getBinaryShapeContent: function (isIntoShape) {
				var sBase64;

				var selectedContent = new AscCommonWord.CSelectedContent();
				AscFormat.ExecuteNoHistory(function () {
					isIntoShape.GetSelectedContent(selectedContent);
				}, this, []);


				if (!selectedContent || !selectedContent.Elements || !selectedContent.Elements.length) {
					return null;
				}

				var oPresentationWriter = new AscCommon.CBinaryFileWriter();

				//начало записи
				oPresentationWriter.WriteString2("");
				oPresentationWriter.WriteString2("");
				oPresentationWriter.WriteDouble(1);
				oPresentationWriter.WriteDouble(1);

				oPresentationWriter.WriteBool(true);
				oPresentationWriter.WriteULong(1);

				if (selectedContent)//пишем контент
				{
					var docContent = selectedContent;

					if (docContent.Elements) {
						var elements = docContent.Elements;

						//пишем метку и длины
						oPresentationWriter.WriteString2("SelectedContent");
						oPresentationWriter.WriteULong(1);
						oPresentationWriter.WriteULong(1);
						oPresentationWriter.WriteULong(1);
						oPresentationWriter.WriteString2("DocContent");
						oPresentationWriter.WriteDouble(elements.length);

						//пишем контент
						for (var Index = 0; Index < elements.length; Index++) {
							var Item;
							if (elements[Index].Element) {
								Item = elements[Index].Element;
							} else {
								Item = elements[Index];
							}

							if (type_Paragraph === Item.GetType()) {
								oPresentationWriter.StartRecord(0);
								oPresentationWriter.WriteParagraph(Item);
								oPresentationWriter.EndRecord();
							}
						}

						sBase64 = oPresentationWriter.GetBase64Memory();
					}

					if (null !== sBase64) {
						sBase64 = "pptData;" + oPresentationWriter.pos + ";" + sBase64;
					}
				}

				return sBase64;
			},

			getText: function (range, worksheet) {
				var t = this;
				var res = null;

				var objectRender = worksheet.objectRender;
				var isIntoShape = objectRender.controller.getTargetDocContent();

				if (isIntoShape) {
					res = t._getTextFromShape(isIntoShape);
				} else {
					res = t._getTextFromSheet(range, worksheet);
				}

				return res;
			},

			getBinaryForMobile: function () {
				var api = window["Asc"]["editor"];
				if (!api || !api.wb) {
					return false;
				}

				var worksheetView = api.wb.getWorksheet();

				var objectRender = worksheetView.objectRender;
				var isIntoShape = objectRender.controller.getTargetDocContent();

				History.TurnOff();
				var sBase64 = null;
				if (!isIntoShape) {
					sBase64 = this.getBinaryForCopy(worksheetView.model, worksheetView.objectRender);
				}
				History.TurnOn();

				var selectedImages = objectRender.getSelectedGraphicObjects();

				var drawingUrls = [];
				if (selectedImages && selectedImages.length) {
					var correctUrl, graphicObj;
					for (var i = 0; i < selectedImages.length; i++) {
						graphicObj = selectedImages[i];
						if (graphicObj.isImage()) {
							if (window["NativeCorrectImageUrlOnCopy"]) {
								correctUrl = window["NativeCorrectImageUrlOnCopy"](graphicObj.getImageUrl());
								drawingUrls[i] = correctUrl;
							} else {
								drawingUrls[i] = graphicObj.getBase64Img();
							}
						}
					}
				}

				return {sBase64: sBase64, drawingUrls: drawingUrls};
			},

			canCopy: function (ws) {
				var res = true;
				if (1 !== ws.model.selectionRange.ranges.length) {
					if (AscCommon.g_clipboardBase.bCut) {
						return false;
					}
					var firstRange = ws.model.selectionRange.ranges[0];
					var byCol = null;
					for (var i = 1; i < ws.model.selectionRange.ranges.length; i++) {
						var nextRange = ws.model.selectionRange.ranges[i];

						if (nextRange.isIntersectWithRanges(ws.model.selectionRange.ranges, i)) {
							res = false;
							break;
						}

						if (byCol === null && firstRange.r1 === nextRange.r1 && firstRange.r2 === nextRange.r2) {
							byCol = false;
						} else if (byCol === null && firstRange.c1 === nextRange.c1 && firstRange.c2 === nextRange.c2) {
							byCol = true;
						} else if ((byCol && firstRange.c1 === nextRange.c1 && firstRange.c2 === nextRange.c2) || (byCol === false && firstRange.r1 === nextRange.r1 && firstRange.r2 === nextRange.r2)) {
							continue;
						} else {
							res = false;
							break;
						}
					}
				}
				return res;
			},

			//TODO пересмотреть функцию
			_generateHtml: function (range, worksheet, isIntoShape, sBase64) {
				var isSelectedImages = this._getSelectedDrawingIndex(worksheet);
				var htmlObj, container;
				if (isIntoShape)//если курсор находится внутри шейпа
				{
					var selectedContent = new AscCommonWord.CSelectedContent();
					AscFormat.ExecuteNoHistory(function () {
						isIntoShape.GetSelectedContent(selectedContent);
					}, this, []);

					var oCopyProcessor = new AscCommon.CopyProcessor({WordControl: {m_oLogicDocument: isIntoShape}});
					oCopyProcessor.CopyDocument2(oCopyProcessor.oRoot, isIntoShape, selectedContent.Elements, true);

					if (sBase64) {
						if (oCopyProcessor.oRoot && oCopyProcessor.oRoot.aChildren && oCopyProcessor.oRoot.aChildren.length > 0) {
							oCopyProcessor.oRoot.aChildren[0].oAttributes["class"] = sBase64;
						}
					}
					return oCopyProcessor.getInnerHtml();
				} else if (isSelectedImages && isSelectedImages !== -1) {//графические объекты
					htmlObj = this._generateHtmlImg(isSelectedImages, worksheet);
					if (!htmlObj) {
						return false;
					}
					container = doc.createElement("DIV");
					container.appendChild(htmlObj);

					if (sBase64 && container.children[0]) {
						container.children[0].setAttribute("class", sBase64);
					}

					return container.innerHTML;
				} else {
					return this._generateHtmlDocStr(range, worksheet, sBase64);
				}
			},

			_generateHtmlImg: function (isSelectedImages, worksheet) {
				if (window["Asc"]["editor"] && window["Asc"]["editor"].isChartEditor) {
					return false;
				}
				let objectRender = worksheet.objectRender;

				objectRender.preCopy();
				let res = document.createElement('span');
				let drawings = worksheet.model.Drawings;

				for (let j = 0; j < isSelectedImages.length; ++j) {
					let oDrawing = drawings[isSelectedImages[j]];
					if (oDrawing) {
						let oGraphicObj = oDrawing.graphicObject;
						if (oGraphicObj) {
							let oHtmlImage = new Image();
							oHtmlImage.src = oGraphicObj.getBase64Img();
							if (oGraphicObj.cachedPixW && oGraphicObj.cachedPixH) {
								oHtmlImage.width = oGraphicObj.cachedPixW;
								oHtmlImage.height = oGraphicObj.cachedPixH;
							} else {
								oHtmlImage.width = oDrawing.getWidthFromTo();
								oHtmlImage.height = oDrawing.getHeightFromTo();
							}
							res.appendChild(oHtmlImage);
						}
					}
				}
				return res;
			},

			_generateHtmlDoc: function (range, worksheet) {
				function skipMerged() {
					var m = merged.filter(function (e) {
						return row >= e.r1 && row <= e.r2 && col >= e.c1 && col <= e.c2
					});
					if (m.length > 0) {
						col = m[0].c2;
						return true;
					}
					return false;
				}

				function makeBorder(border) {
					if (!border || border.s === c_oAscBorderStyles.None) {
						return "";
					}

					var style = "";
					switch (border.s) {
						case c_oAscBorderStyles.Thin:
							style = "solid";
							break;
						case c_oAscBorderStyles.Medium:
							style = "solid";
							break;
						case c_oAscBorderStyles.Thick:
							style = "solid";
							break;
						case c_oAscBorderStyles.DashDot:
						case c_oAscBorderStyles.DashDotDot:
						case c_oAscBorderStyles.Dashed:
							style = "dashed";
							break;
						case c_oAscBorderStyles.Double:
							style = "double";
							break;
						case c_oAscBorderStyles.Hair:
						case c_oAscBorderStyles.Dotted:
							style = "dotted";
							break;
						case c_oAscBorderStyles.MediumDashDot:
						case c_oAscBorderStyles.MediumDashDotDot:
						case c_oAscBorderStyles.MediumDashed:
						case c_oAscBorderStyles.SlantDashDot:
							style = "dashed";
							break;
					}
					return border.w + "px " + style + " " + number2color(border.getRgbOrNull());
				}

				var fn = worksheet.model.workbook.getDefaultFont();
				var fs = worksheet.model.workbook.getDefaultSize();
				var bbox = range.getBBox0();
				var merged = [];
				var t = this;
				var table, tr, td, cell, j, row, col, mbbox, h, w, b;

				table = document.createElement("TABLE");
				table.style.display = 'none';
				table.cellPadding = "0";
				table.cellSpacing = "0";
				table.style.borderCollapse = "collapse";
				table.style.fontFamily = fn;
				table.style.fontSize = fs + "pt";
				table.style.color = "#000";
				table.style.backgroundColor = "transparent";

				var maxRow = bbox.r2;
				var maxCol = bbox.c2;
				var maxRowCol = this._getRangeMaxRowCol(worksheet, bbox, range);
				if (null !== maxRowCol) {
					maxRow = maxRowCol.row;
					maxCol = maxRowCol.col;
				}

				for (row = bbox.r1; row <= maxRow; ++row) {
					if (worksheet.model.bExcludeHiddenRows && worksheet.model.getRowHidden(row)) {
						continue;
					}

					tr = document.createElement("TR");
					h = worksheet.model.getRowHeight(row);
					if (h > 0) {
						tr.style.height = h + "pt";
					}

					for (col = bbox.c1; col <= maxCol; ++col) {
						if (skipMerged()) {
							continue;
						}

						cell = worksheet.model.getCell3(row, col);
						td = document.createElement("TD");

						var style = "";
						mbbox = cell.hasMerged();
						if (mbbox) {
							merged.push(mbbox);
							td.colSpan = mbbox.c2 - mbbox.c1 + 1;
							td.rowSpan = mbbox.r2 - mbbox.r1 + 1;
							for (w = 0, j = mbbox.c1; j <= mbbox.c2; ++j) {
								w += worksheet.getColumnWidth(j, 1/*pt*/);
							}
							style += "width:" + w + "pt" + ";";
						} else {
							style += "width:" + worksheet.getColumnWidth(col, 1/*pt*/) + "pt" + ";";
						}

						if (cell.getType() !== null) {
							var align = cell.getAlign();
							if (!align.getWrap()) {
								style += "white-space:" + "nowrap" + ";";
							} else {
								style += "white-space:" + "normal" + ";";
							}

							switch (align.getAlignHorizontal()) {
								case AscCommon.align_Left:
									style += "text-align:" + "left" + ";";
									break;
								case AscCommon.align_Right:
									style += "text-align:" + "right" + ";";
									break;
								case AscCommon.align_Center:
									style += "text-align:" + "center" + ";";
									break;
								case AscCommon.align_Justify:
									style += "text-align:" + "justify" + ";";
									break;
							}
							switch (align.getAlignVertical()) {
								case Asc.c_oAscVAlign.Bottom:
									style += "vertical-align:" + "bottom" + ";";
									break;
								case Asc.c_oAscVAlign.Center:
								case Asc.c_oAscVAlign.Dist:
								case Asc.c_oAscVAlign.Just:
									style += "vertical-align:" + "middle" + ";";
									break;
								case Asc.c_oAscVAlign.Top:
									style += "vertical-align:" + "top" + ";";
									break;
							}

							//borders
							b = cell.getBorderFull();
							var _border;
							if (mbbox) {
								var cellMergeFinish = worksheet.model.getCell3(mbbox.r2, mbbox.c2);
								var borderMergeCell = cellMergeFinish.getBorderFull();
								_border = makeBorder(borderMergeCell.r);
								style += _border ? "border-right:" + _border + ";" : "";
								_border = makeBorder(borderMergeCell.b);
								style += _border ? "border-bottom:" + makeBorder(borderMergeCell.b) + ";" : "";
							} else {
								_border = makeBorder(b.r);
								style += _border ? "border-right:" + makeBorder(b.r) + ";" : "";
								_border = makeBorder(b.b);
								style += _border ? "border-bottom:" + makeBorder(b.b) + ";" : "";
							}
							_border = makeBorder(b.l);
							style += _border ? "border-left:" + makeBorder(b.l) + ";" : "";
							_border = makeBorder(b.t);
							style += _border ? "border-top:" + makeBorder(b.t) + ";" : "";

							b = cell.getFillColor();
							// если b==0 мы не зайдем в if, хотя b==0 это ни что иное, как черный цвет заливки.
							if (b != null) {
								style += "background-color:" + number2color(b.getRgb()) + ";";
							}

							td.setAttribute("style", style);

							this._makeNodesFromCellValue(cell.getValue2(), fn, fs, cell).forEach(function (node) {
								td.appendChild(node);
							});
						}

						tr.appendChild(td);
					}
					table.appendChild(tr);
				}

				return table;
			},

			_generateHtmlDocStr: function (range, worksheet, sBase64) {
				function skipMerged() {
					var m = merged.filter(function (e) {
						return row >= e.r1 && row <= e.r2 && col >= e.c1 && col <= e.c2
					});
					if (m.length > 0) {
						col = m[0].c2;
						return true;
					}
					return false;
				}

				function makeBorder(border) {
					if (!border || border.s === c_oAscBorderStyles.None) {
						return "";
					}

					var _style = "";
					switch (border.s) {
						case c_oAscBorderStyles.Thin:
							_style = "solid";
							break;
						case c_oAscBorderStyles.Medium:
							_style = "solid";
							break;
						case c_oAscBorderStyles.Thick:
							_style = "solid";
							break;
						case c_oAscBorderStyles.DashDot:
						case c_oAscBorderStyles.DashDotDot:
						case c_oAscBorderStyles.Dashed:
							_style = "dashed";
							break;
						case c_oAscBorderStyles.Double:
							_style = "double";
							break;
						case c_oAscBorderStyles.Hair:
						case c_oAscBorderStyles.Dotted:
							_style = "dotted";
							break;
						case c_oAscBorderStyles.MediumDashDot:
						case c_oAscBorderStyles.MediumDashDotDot:
						case c_oAscBorderStyles.MediumDashed:
						case c_oAscBorderStyles.SlantDashDot:
							_style = "dashed";
							break;
					}
					return border.w + "px " + _style + " " + number2color(border.getRgbOrNull());
				}

				var fn = worksheet.model.workbook.getDefaultFont();
				var fs = worksheet.model.workbook.getDefaultSize();
				var bbox = range.getBBox0();
				var merged = [];
				var t = this;
				var table, tr, td, cell, j, row, col, mbbox, h, w, b;

				var str = "";

				var addByStr = function (_str) {
					//исследования по максимальной длине строки:
					//Edge/chrome 96.0.4664 -> 536870888 символов(1073741776 байта)
					//FF 95.0.2 -> 1073741822 символов(2147483644 байта)
					if (str.length + _str.length > c_MaxStringLength) {
						return false;
					}
					str += _str;
					return true;
				};

				if (!addByStr('<table cellpadding="0" cellSpacing="0"')) {
					return false;
				}

				if (sBase64) {
					if (!addByStr(' class=' + '"' + sBase64 + '"')) {
						return false;
					}
				}
				if (!addByStr(' style="border-collapse: collapse;font-family:' + fn + ';font-size:' + fs + 'pt' + ';color:#000;background-color:transparent;' + '">')) {
					return false;
				}

				var maxRow = bbox.r2;
				var maxCol = bbox.c2;
				var maxRowCol = this._getRangeMaxRowCol(worksheet, bbox, range);
				if (null !== maxRowCol) {
					maxRow = maxRowCol.row;
					maxCol = maxRowCol.col;
				}

				for (row = bbox.r1; row <= maxRow; ++row) {
					if (worksheet.model.bExcludeHiddenRows && worksheet.model.getRowHidden(row)) {
						continue;
					}
					if (!addByStr('<tr style="height:')) {
						return false;
					}

					h = worksheet.model.getRowHeight(row);

					if (!addByStr(h + "pt" + '"' + ">")) {
						return false;
					}

					for (col = bbox.c1; col <= maxCol; ++col) {
						if (skipMerged()) {
							continue;
						}

						if (!addByStr('<td')) {
							return false;
						}

						var style = "";
						cell = worksheet.model.getCell3(row, col);
						mbbox = cell.hasMerged();
						if (mbbox) {
							merged.push(mbbox);
							if (!addByStr(" colspan=" + (mbbox.c2 - mbbox.c1 + 1))) {
								return false;
							}
							if (!addByStr(" rowSpan=" + (mbbox.r2 - mbbox.r1 + 1))) {
								return false;
							}
							for (w = 0, j = mbbox.c1; j <= mbbox.c2; ++j) {
								w += worksheet.getColumnWidth(j, 1/*pt*/);
							}
							style += "width:" + w + "pt" + ";";
						} else {
							style += "width:" + worksheet.getColumnWidth(col, 1/*pt*/) + "pt" + ";";
						}

						if (cell.getType() !== null) {
							var align = cell.getAlign();
							if (!align.getWrap()) {
								style += "white-space:" + "nowrap" + ";";
							} else {
								style += "white-space:" + "normal" + ";";
							}

							switch (align.getAlignHorizontal()) {
								case AscCommon.align_Left:
									style += "text-align:" + "left" + ";";
									break;
								case AscCommon.align_Right:
									style += "text-align:" + "right" + ";";
									break;
								case AscCommon.align_Center:
									style += "text-align:" + "center" + ";";
									break;
								case AscCommon.align_Justify:
									style += "text-align:" + "justify" + ";";
									break;
							}
							switch (align.getAlignVertical()) {
								case Asc.c_oAscVAlign.Bottom:
									style += "vertical-align:" + "bottom" + ";";
									break;
								case Asc.c_oAscVAlign.Center:
								case Asc.c_oAscVAlign.Dist:
								case Asc.c_oAscVAlign.Just:
									style += "vertical-align:" + "middle" + ";";
									break;
								case Asc.c_oAscVAlign.Top:
									style += "vertical-align:" + "top" + ";";
									break;
							}

							//borders
							b = cell.getBorderFull();
							var _border;
							if (mbbox) {
								var cellMergeFinish = worksheet.model.getCell3(mbbox.r2, mbbox.c2);
								var borderMergeCell = cellMergeFinish.getBorderFull();
								_border = makeBorder(borderMergeCell.r);
								style += _border ? "border-right:" + _border + ";" : "";
								_border = makeBorder(borderMergeCell.b);
								style += _border ? "border-bottom:" + makeBorder(borderMergeCell.b) + ";" : "";
							} else {
								_border = makeBorder(b.r);
								style += _border ? "border-right:" + makeBorder(b.r) + ";" : "";
								_border = makeBorder(b.b);
								style += _border ? "border-bottom:" + makeBorder(b.b) + ";" : "";
							}
							_border = makeBorder(b.l);
							style += _border ? "border-left:" + makeBorder(b.l) + ";" : "";
							_border = makeBorder(b.t);
							style += _border ? "border-top:" + makeBorder(b.t) + ";" : "";

							b = cell.getFillColor();
							// если b==0 мы не зайдем в if, хотя b==0 это ни что иное, как черный цвет заливки.
							if (b != null) {
								style += "background-color:" + number2color(b.getRgb()) + ";";
							}

							if (!addByStr(' style=' + '"' + style + '">')) {
								return false;
							}
							if (!addByStr(this._makeNodesFromCellValueStr(cell.getValue2(), fn, fs, cell))) {
								return false;
							}
						} else {
							if (!addByStr('>')) {
								return false;
							}
						}
						if (!addByStr('</td>')) {
							return false;
						}
					}
					if (!addByStr("</tr>")) {
						return false;
					}
				}
				if (!addByStr("</table>")) {
					return false;
				}

				return str;
			},

			_getSelectedDrawingIndex: function (worksheet) {
				if (!worksheet) {
					return false;
				}
				var images = worksheet.model.Drawings;
				var n = 0;
				var arrImages = [];
				if (images) {
					for (var i = 0; i < images.length; i++) {
						if ((images[i].graphicObject && images[i].graphicObject.selected === true)) {
							arrImages[n] = i;
							n++;
						}
					}
				}
				if (n === 0) {
					return -1;
				} else {
					return arrImages;
				}
			},

			_makeNodesFromCellValue: function (val, defFN, defFS, cell) {
				var i, res, span, f;

				function getTextDecoration(format) {
					var res = [];
					if (Asc.EUnderline.underlineNone !== format.getUnderline()) {
						res.push("underline");
					}
					if (format.getStrikeout()) {
						res.push("line-through");
					}
					return res.length > 0 ? res.join(",") : "";
				}

				var hyperlink;
				if (cell) {
					hyperlink = cell.getHyperlink();
				}

				for (res = [], i = 0; i < val.length; ++i) {
					if (val[i] && val[i].format && val[i].format.getSkip()) {
						continue;
					}
					if (cell == undefined || (cell != undefined && (hyperlink == null || (hyperlink != null && hyperlink.getLocation() != null)))) {
						span = doc.createElement("SPAN");
					} else {
						span = doc.createElement("A");
						if (hyperlink.Hyperlink != null) {
							span.href = hyperlink.Hyperlink;
						} else if (hyperlink.getLocation() != null) {
							span.href = "#" + hyperlink.getLocation();
						}
						if (hyperlink.Tooltip != null) {
							span.title = hyperlink.Tooltip;
						}
					}

					var text = val[i].getFragmentText();
					var isBr = val[i].getFragmentText().indexOf("\n");
					span.textContent = text;

					f = val[i].format;
					var fn = f.getName();
					var fs = f.getSize();
					var fc = f.getColor();
					var va = f.getVerticalAlign();
					var style = "";
					if (fc && fc.rgb !== 0) {
						style += "color:" + number2color(fc.getRgb()) + ";";
					}

					if (fn !== defFN) {
						style += "font-family:" + fn + ";";
					}
					if (fs !== defFS) {
						style += "font-size:" + fs + 'pt' + ";";
					}
					if (f.getBold()) {
						style += "font-weight:" + "bold" + ";";
					}
					if (f.getItalic()) {
						style += "font-style:" + "italic" + ";";
					}
					var textDecoration = getTextDecoration(f);
					if (textDecoration) {
						style += "text-decoration:" + textDecoration + ";";
					}
					var verticalAlign = va === AscCommon.vertalign_SubScript ? 'sub' : va === AscCommon.vertalign_SuperScript ? 'super' : 'baseline';
					style += "vertical-align:" + verticalAlign + ";";
					span.setAttribute('style', style);

					if (isBr !== -1) {
						span.innerHTML = span.innerHTML.replace(/\n/g, '<br>');
					}

					res.push(span);
				}
				return res;
			},

			_makeNodesFromCellValueStr: function (val, defFN, defFS, cell) {
				var i, res, span, f;

				function getTextDecoration(format) {
					var res = [];
					if (Asc.EUnderline.underlineNone !== format.getUnderline()) {
						res.push("underline");
					}
					if (format.getStrikeout()) {
						res.push("line-through");
					}
					return res.length > 0 ? res.join(",") : "";
				}

				var str = "";
				var hyperlink;
				if (cell) {
					hyperlink = cell.getHyperlink();
				}

				for (res = [], i = 0; i < val.length; ++i) {
					if (val[i] && val[i].format && val[i].format.getSkip()) {
						continue;
					}
					if (cell == undefined || (cell != undefined && (hyperlink == null || (hyperlink != null && hyperlink.getLocation() != null)))) {
						str += '<span ';
					} else {
						str += '<a';
						if (hyperlink.Hyperlink != null) {
							str += ' href=' + '"' + hyperlink.Hyperlink + '"';
						} else if (hyperlink.getLocation() != null) {
							str += ' href=' + '"' + "#" + hyperlink.getLocation() + '"';
						}
						if (hyperlink.Tooltip) {
							str += ' title=' + '"' + hyperlink.Tooltip + '" ';
						}
					}

					var text = CopyPasteCorrectString(val[i].getFragmentText());
					text = text.replace(/\n/g, '<br>');

					f = val[i].format;
					var fn = f.getName();
					var fs = f.getSize();
					var fc = f.getColor();
					var va = f.getVerticalAlign();
					var style = "";
					if (fc && fc.rgb !== 0) {
						style += "color:" + number2color(fc.getRgb()) + ";";
					}

					if (fn !== defFN) {
						style += "font-family:" + fn + ";";
					}
					if (fs !== defFS) {
						style += "font-size:" + fs + 'pt' + ";";
					}
					if (f.getBold()) {
						style += "font-weight:" + "bold" + ";";
					}
					if (f.getItalic()) {
						style += "font-style:" + "italic" + ";";
					}
					var textDecoration = getTextDecoration(f);
					if (textDecoration) {
						style += "text-decoration:" + textDecoration + ";";
					}
					var verticalAlign = va === AscCommon.vertalign_SubScript ? 'sub' : va === AscCommon.vertalign_SuperScript ? 'super' : 'baseline';
					style += "vertical-align:" + verticalAlign + ";";

					str += ' style="' + style + '">';
					str += text + '</span>';
				}
				return str;
			},

			_getTextFromShape: function (documentContent) {
				var res = null;

				var addByStr = function (_str) {
					if (res && res.length + _str.length > c_MaxStringLength) {
						return false;
					}
					if (!res) {
						res = "";
					}
					res += _str;
					return true;
				};

				if (documentContent && documentContent.Content && documentContent.Content.length) {
					for (var i = 0; i < documentContent.Content.length; i++) {
						if (documentContent.Content[i]) {
							var paraText = documentContent.Content[i].GetSelectedText();
							if (paraText) {
								if (i !== 0) {
									if (!addByStr('\r\n')) {
										return null;
									}
								}
								if (!addByStr(paraText)) {
									return null;
								}
							}
						}
					}
				}

				return res;
			},

			_getTextFromSheet: function (range, worksheet) {
				var res = null;

				var addByStr = function (_str) {
					if (res && res.length + _str.length > c_MaxStringLength) {
						return false;
					}
					if (!res) {
						res = "";
					}
					res += _str;
					return true;
				};

				if (range) {
					var bbox = range.bbox;

					var maxRow = bbox.r2;
					var maxCol = bbox.c2;
					var maxRowCol = this._getRangeMaxRowCol(worksheet.model, bbox, range);
					if (null !== maxRowCol) {
						maxRow = maxRowCol.row;
						maxCol = maxRowCol.col;
					}

					res = '';
					for (var row = bbox.r1; row <= maxRow; ++row) {
						if (worksheet.model.bExcludeHiddenRows && worksheet.model.getRowHidden(row)) {
							continue;
						}

						if (row !== bbox.r1) {
							if (!addByStr('\r\n')) {
								return null;
							}
						}

						for (var col = bbox.c1; col <= maxCol; ++col) {
							if (col !== bbox.c1) {
								if (!addByStr('\t')) {
									return null;
								}
							}

							var currentRange = worksheet.model.getCell3(row, col);
							var textRange = currentRange.getValueWithFormat();
							if (textRange !== '' && textRange !== undefined) {
								if (!addByStr(textRange)) {
									return null;
								}
							}
						}
					}
				}

				return res;
			}
		};

		function PasteProcessorExcel() {
			this.activeRange = null;
			this.alreadyLoadImagesOnServer = false;

			this.fontsNew = {};
			this.oImages = {};
			this.multipleSettings = null;
			this.pasteCallBack = null;
		}

		PasteProcessorExcel.prototype = {

			constructor: PasteProcessorExcel,

			clean: function () {
				this.activeRange = null;
				this.alreadyLoadImagesOnServer = false;

				this.fontsNew = {};
				this.oImages = {};
				this.multipleSettings = null;
			},

			pasteFromBinary: function (worksheet, binary, isCellEditMode, isPasteAll) {
				var base64 = null, base64FromWord = null, base64FromPresentation = null, base64FromPDF = null, t = this;

				if (binary.indexOf("xslData;") > -1) {
					base64 = binary.split('xslData;')[1];
				} else if (binary.indexOf("docData;") > -1) {
					base64FromWord = binary.split('docData;')[1];
				} else if (binary.indexOf("pptData;") > -1) {
					base64FromPresentation = binary.split('pptData;')[1];
				} else if (binary.indexOf("pdfData;") > -1) {
					base64FromPDF = binary.split('pdfData;')[1];
				}

				var result = false;
				var objectRender = worksheet.objectRender;
				var isIntoShape = objectRender && objectRender.controller && objectRender.controller.getTargetDocContent();
				if (base64 != null)//from excel
				{
					result = this._pasteFromBinaryExcel(worksheet, base64, isIntoShape, isCellEditMode, isPasteAll);
				} else if (base64FromWord)//from word
				{
					this.activeRange = worksheet.model.selectionRange.getLast().clone(true);
					result = this._pasteFromBinaryWord(worksheet, base64FromWord, isIntoShape, isCellEditMode);
					window['AscCommon'].g_specialPasteHelper.specialPasteData.pasteFromWord = true;
					worksheet.workbook.handlers.trigger("cleanCopyData", true);
				} else if (base64FromPresentation) {
					result = this._pasteFromBinaryPresentation(worksheet, base64FromPresentation, isIntoShape, isCellEditMode);
					worksheet.workbook.handlers.trigger("cleanCopyData", true);
				} else if (base64FromPDF) {
					result = this._pasteFromBinaryPDF(worksheet, base64FromPDF, isIntoShape, isCellEditMode);
					worksheet.workbook.handlers.trigger("cleanCopyData", true);
				}

				return result;
			},

			_pasteFromBinaryExcel: function (worksheet, base64, isIntoShape, isCellEditMode, isPasteAll) {
				var t = this;
				var newFonts;
				var tempWorkbook = new AscCommonExcel.Workbook(undefined, undefined, false);
				tempWorkbook.DrawingDocument = Asc.editor.wbModel.DrawingDocument;
				var aPastedImages = this._readExcelBinary(base64, tempWorkbook);

				if (!isIntoShape && this._checkCutBefore(worksheet, tempWorkbook)) {
					return true;
				}

				var pasteData = null;
				if (tempWorkbook) {
					pasteData = tempWorkbook.aWorksheets[0];
				}

				if (isPasteAll) {
					History.Create_NewPoint();
					History.StartTransaction();
				}

				var doPasteData = function () {
					newFonts = {};
					newFonts = tempWorkbook.generateFontMap2();
					newFonts = t._convertFonts(newFonts);

					if (isPasteAll) {
						window['AscCommon'].g_specialPasteHelper.Special_Paste_Hide_Button();
					}

					//закрываем общую транзакцию _loadDataBeforePaste после загрузки шрифтов
					worksheet.setSelectionInfo('paste', {
						data: pasteData,
						fromBinary: true,
						fontsNew: newFonts,
						pasteAllSheet: isPasteAll,
						wb: tempWorkbook
					});
				};

				var doPasteIntoShape = function () {
					History.TurnOff();
					var docContent = t._convertTableFromExcelToDocument(worksheet, pasteData, isIntoShape);
					History.TurnOn();

					var callback = function (isSuccess) {
						if (isSuccess) {
							t._insertBinaryIntoShapeContent(worksheet, [docContent]);
						}
						window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					};

					worksheet.objectRender.controller.checkPasteInText(callback);
				};

				let getFontsFromDrawings = function (_drawings) {
					let _fonts = {};
					if (_drawings) {
						for (var i = 0; i < _drawings.length; i++) {
							_drawings[i].graphicObject.getAllFonts(_fonts);
						}
					}
					return _fonts;
				};

				var res = false;
				if (isCellEditMode) {
					res = this._getTextFromWorksheet(pasteData);
				} else if (pasteData) {
					if (pasteData.Drawings && pasteData.Drawings.length) {

						if (window["IS_NATIVE_EDITOR"]) {
							t._insertImagesFromBinary(worksheet, pasteData, isIntoShape, null, isPasteAll, tempWorkbook);
							if (isPasteAll) {
								doPasteData();
							}
						} else if (window["NativeCorrectImageUrlOnPaste"]) {
							newFonts = getFontsFromDrawings(pasteData.Drawings);
							worksheet._loadFonts(newFonts, function () {
								var url;
								for (var i = 0, length = aPastedImages.length; i < length; ++i) {
									url = window["NativeCorrectImageUrlOnPaste"](aPastedImages[i].Url);
									aPastedImages[i].Url = url;

									var imageElem = aPastedImages[i];
									if (null != imageElem) {
										imageElem.SetUrl(url);
									}
								}

								t._insertImagesFromBinary(worksheet, pasteData, isIntoShape, null, isPasteAll, tempWorkbook);
								if (isPasteAll) {
									doPasteData();
								}
							});
						} else if (!(window["Asc"]["editor"] && window["Asc"]["editor"].isChartEditor)) {

							newFonts = getFontsFromDrawings(pasteData.Drawings);

							worksheet._loadFonts(newFonts, function () {
								if (aPastedImages && aPastedImages.length) {
									t._loadImagesOnServer(aPastedImages, function () {
										t._insertImagesFromBinary(worksheet, pasteData, isIntoShape, null, isPasteAll, tempWorkbook);
										if (isPasteAll) {
											doPasteData();
										}
									});
								} else {
									t._insertImagesFromBinary(worksheet, pasteData, isIntoShape, null, isPasteAll, tempWorkbook);
									if (isPasteAll) {
										doPasteData();
									}
								}
							});
						}
					} else {
						if (isIntoShape) {
							doPasteIntoShape();
						} else if (this._checkPasteFromBinaryExcel(worksheet, true, pasteData)) {
							doPasteData();
						}
					}

					res = true;
				}

				return res;
			},

			_readExcelBinary: function (base64, tempWorkbook, selectAllSheet) {
				let oBinaryFileReader = new AscCommonExcel.BinaryFileReader(true);
				oBinaryFileReader.InitOpenManager.copyPasteObj.selectAllSheet = selectAllSheet;
				let t = this;
				let aPastedImages;

				AscFormat.ExecuteNoHistory(function () {
					pptx_content_loader.Start_UseFullUrl();
					pptx_content_loader.Reader.ClearConnectedObjects();
					oBinaryFileReader.Read(base64, tempWorkbook);

					//paste multiselect
					//TODO if paste multiselect into multiselect - we must show error
					let pastedWorksheet = tempWorkbook.aWorksheets && tempWorkbook.aWorksheets[0];
					let _ranges = pastedWorksheet && pastedWorksheet.selectionRange && pastedWorksheet.selectionRange.ranges;
					if (_ranges && _ranges.length > 1) {
						//paste - should only be with an equal number of rows/columns
						//there should be only rows/columns between the ranges, then -> delete the extra rows/columns
						let byCol = null;
						if (_ranges[0].r1 === _ranges[1].r1 && _ranges[0].r2 === _ranges[1].r2) {
							byCol = false;
						} else if (_ranges[0].c1 === _ranges[1].c1 && _ranges[0].c2 === _ranges[1].c2) {
							byCol = true;
						}

						_ranges.sort(function sortArr(a, b) {
							if (!byCol) {
								return a.c1 > b.c1 ? 1 : -1;
							} else {
								return a.r1 > b.r1 ? 1 : -1;
							}
						});


						let diff = 0;
						for (let i = 1; i < _ranges.length; i++) {
							if (byCol) {
								if (_ranges[i - 1].r2 + 1 !== _ranges[i].r1) {
									pastedWorksheet.removeRows(_ranges[i - 1].r2 + 1 - diff, _ranges[i].r1 - 1 - diff);
									diff += _ranges[i].r1 - _ranges[i - 1].r2 - 1;
								}
							} else {
								if (_ranges[i - 1].c2 + 1 !== _ranges[i].c1) {
									pastedWorksheet.removeCols(_ranges[i - 1].c2 + 1 - diff, _ranges[i].c1 - 1 - diff);
									diff += _ranges[i].c1 - _ranges[i - 1].c2 - 1;
								}
							}
						}

						if (diff !== 0) {
							AscCommonExcel.executeInR1C1Mode(false, function () {
								let pasteRange = AscCommonExcel.g_oRangeCache.getAscRange(oBinaryFileReader.InitOpenManager.copyPasteObj.activeRange);
								if (pasteRange) {
									pasteRange = pasteRange.clone();
									if (byCol) {
										pasteRange.r2 -= diff;
									} else {
										pasteRange.c2 -= diff;
									}
									oBinaryFileReader.InitOpenManager.copyPasteObj.activeRange = pasteRange.getName();
								}
							});
						}
					}

					t.activeRange = oBinaryFileReader.InitOpenManager.copyPasteObj.activeRange;

					aPastedImages = pptx_content_loader.End_UseFullUrl();
					pptx_content_loader.Reader.AssignConnectedObjects();
				}, this, []);

				return aPastedImages;
			},

			_checkCutBefore: function (ws, pastedWb) {
				var res = false;

				//***MOVE***
				//проверяем, может это вырезанный фрагмент пытаемся вставить в пределах одного документа
				//чтобы не передавать изменения на сервер, даже в случае одного пользователя в разных вкладках
				//вырезать и вставить будут работать независимо, поэтому при вставке сравнивем ещё и id юзера
				var pasteInOriginalDoc = this._checkPastedInOriginalDoc(pastedWb);
				if (pasteInOriginalDoc && null !== window["Asc"]["editor"].wb.cutIdSheet) {
					var wsFrom = window["Asc"]["editor"].wb.getWorksheetById(window["Asc"]["editor"].wb.cutIdSheet);
					var fromRange = wsFrom ? wsFrom.getCutRange() : null;
					if (fromRange) {
						fromRange = fromRange[0];
						var aRange = ws.model.selectionRange.getLast();
						var toRange = new Asc.Range(aRange.c1, aRange.r1, aRange.c1 + (fromRange.c2 - fromRange.c1), aRange.r1 + (fromRange.r2 - fromRange.r1));

						if (ws.model.getSheetProtection() && ws.model.isIntersectLockedRanges([toRange])) {
							ws.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.ChangeOnProtectedSheet, c_oAscError.Level.NoCritical);
							ws.handlers.trigger("cleanCutData", true);
							return true;
						}

						if (ws.model.isUserProtectedRangesIntersection(toRange)) {
							ws.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.ProtectedRangeByOtherUser, c_oAscError.Level.NoCritical);
							ws.handlers.trigger("cleanCutData", true);
							return;
						}

						var wsTo = ws.model.Id !== wsFrom.model.Id ? ws : null;
						wsFrom.applyCutRange(fromRange, toRange, wsTo);
						ws.handlers.trigger("cleanCutData", true);
						res = true;

						AscCommon.g_clipboardBase.needClearBuffer = true;
					}
				} else if (!pasteInOriginalDoc) {
					ws.workbook.handlers.trigger("cleanCopyData", true);
				}

				return res;
			},

			_checkPastedInOriginalDoc: function (pastedWb, notCheckUser) {
				var res = false;

				var api = window["Asc"]["editor"];
				var curDocId = api.DocInfo && api.DocInfo.Id;
				var curUserId = api.CoAuthoringApi && api.CoAuthoringApi.getUserConnectionId();

				if (pastedWb && pastedWb.Core && pastedWb.Core.identifier === curDocId) {
					if (notCheckUser) {
						res = true;
					} else if (pastedWb.Core.creator === curUserId) {
						res = true;
					}
				}

				return res;
			},

			_pasteFromBinaryWord: function (worksheet, base64, isIntoShape, isCellEditMode) {
				var res = true;
				var t = this;
				var pasteData = this.ReadFromBinaryWord(base64, worksheet);

				if (isCellEditMode) {
					res = this._getTextFromWord(pasteData.content);
				}
				//insert binary from word into SHAPE
				else if (isIntoShape) {
					var callback = function (isSuccess) {
						if (isSuccess) {
							t._insertBinaryIntoShapeContent(worksheet, pasteData.content, true);
						}
						window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					};

					worksheet.objectRender.controller.checkPasteInText(callback);
				} else {
					var oPasteFromBinaryWord = new pasteFromBinaryWord(this, worksheet);
					oPasteFromBinaryWord._paste(worksheet, pasteData);
				}

				return res;
			},

			_pasteFromBinaryPresentation: function (worksheet, base64, isIntoShape, isCellEditMode) {
				pptx_content_loader.Clear();

				var _stream = AscFormat.CreateBinaryReader(base64, 0, base64.length);
				var stream = new AscCommon.FileStream(_stream.data, _stream.size);
				var p_url = stream.GetString2();
				var p_theme = stream.GetString2();
				var p_width = stream.GetULong() / 100000;
				var p_height = stream.GetULong() / 100000;
				var t = this;

				var bIsMultipleContent = stream.GetBool();

				var selectedContent2 = [], i;
				if (true === bIsMultipleContent) {
					var multipleParamsCount = stream.GetULong();
					for (i = 0; i < multipleParamsCount; i++) {
						selectedContent2.push(this._readPresentationSelectedContent(stream, worksheet));
					}
				}

				var specialOptionsArr = [];
				var specialProps = Asc.c_oSpecialPasteProps;
				if (2 === multipleParamsCount) {
					specialOptionsArr = [specialProps.sourceformatting];
				} else if (3 === multipleParamsCount) {
					specialOptionsArr = [specialProps.sourceformatting, specialProps.picture];
				}

				var defaultSelectedContent = selectedContent2[0] ? selectedContent2[0] : selectedContent2[1];
				var bSlideObjects = defaultSelectedContent && defaultSelectedContent.content.SlideObjects &&
					defaultSelectedContent.content.SlideObjects.length > 0;
				var pasteObj = bSlideObjects ? selectedContent2[2] : defaultSelectedContent;

				if (window['AscCommon'].g_specialPasteHelper.specialPasteStart) {
					var props = window['AscCommon'].g_specialPasteHelper.specialPasteProps.property;
					switch (props) {
						case Asc.c_oSpecialPasteProps.picture: {
							if (selectedContent2[2]) {
								pasteObj = selectedContent2[2];
							}
							break;
						}
					}
				}

				var arr_Images, fonts, content = null;
				if (pasteObj) {
					arr_Images = pasteObj.images;
					fonts = pasteObj.fonts;
					content = pasteObj.content;
				}

				if (null === content) {
					window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					return;
				}

				if (content.DocContent) {
					var docContent = content.DocContent.Elements;

					if (isCellEditMode) {
						var text = this._getTextFromWord(docContent);
						window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
						return text;
					} else if (isIntoShape) {
						var callback = function (isSuccess) {
							if (isSuccess) {
								t._insertBinaryIntoShapeContent(worksheet, docContent);
							}
							window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
						};

						worksheet.objectRender.controller.checkPasteInText(callback);
						return true;
					} else {
						History.TurnOff();
						var oPasteFromBinaryWord = new pasteFromBinaryWord(this, worksheet, true);

						var oTempDrawingDocument = window["Asc"]["editor"].wbModel.getDrawingDocument();
						var newCDocument = new CDocument(oTempDrawingDocument, false);
						newCDocument.bFromDocument = true;
						newCDocument.theme = window["Asc"]["editor"].wbModel.theme;

						var newContent = [];
						for (i = 0; i < docContent.length; i++) {
							if (type_Paragraph === docContent[i].GetType())//paragraph
							{
								docContent[i] = AscFormat.ConvertParagraphToWord(docContent[i], newCDocument);
								docContent[i].bFromDocument = true;
								newContent.push(docContent[i]);
							} else if (type_Table === docContent[i].GetType())//table
							{
								//TODO вырезать из таблицы параграфы
							}
						}
						docContent = newContent;

						History.TurnOn();

						oPasteFromBinaryWord._paste(worksheet, {content: docContent});
					}
				} else if (content.Drawings) {

					if (isCellEditMode) {
						return "";
					}

					if (!bSlideObjects) {
						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
						specialProps = window['AscCommon'].g_specialPasteHelper.buttonInfo;
						if (specialOptionsArr.length > 1) {
							specialProps.asc_setOptions(specialOptionsArr);
						}
					}

					var arr_shapes = content.Drawings;
					if (arr_shapes && arr_shapes.length && !(window["Asc"]["editor"] && window["Asc"]["editor"].isChartEditor)) {
						if (!bSlideObjects && content.Drawings.length === selectedContent2[1].content.Drawings.length) {
							var oEndContent = {
								Drawings: []
							};
							var oSourceContent = {
								Drawings: []
							};
							for (i = 0; i < content.Drawings.length; ++i) {
								oEndContent.Drawings.push({Drawing: content.Drawings[i].graphicObject});
								oSourceContent.Drawings.push({Drawing: selectedContent2[1].content.Drawings[i].graphicObject});
							}
							AscFormat.checkDrawingsTransformBeforePaste(oEndContent, oSourceContent, null);
						}
						var newFonts = {};
						for (i = 0; i < arr_shapes.length; i++) {
							arr_shapes[i].graphicObject.getAllFonts(newFonts);
						}

						var aPastedImages = arr_Images;
						worksheet._loadFonts(newFonts, function () {
							if (aPastedImages && aPastedImages.length) {
								t._loadImagesOnServer(aPastedImages, function () {
									t._insertImagesFromBinary(worksheet, {Drawings: arr_shapes}, isIntoShape, true);
								});
							} else {
								t._insertImagesFromBinary(worksheet, {Drawings: arr_shapes}, isIntoShape, true);
							}
						});
					} else {
						window['AscCommon'].g_specialPasteHelper.buttonInfo.clean();
					}

					return true;
				}
			},

			_readPresentationSelectedContent: function (stream, worksheet) {

				var presentationSelectedContent = null;
				var fonts = [];
				var arr_Images = {};
				var oThis = this;

				var loader = new AscCommon.BinaryPPTYLoader();
				loader.presentation = worksheet.model;
				loader.DrawingDocument = worksheet.getDrawingDocument();
				loader.stream = stream;

				var readContent = function () {
					History.TurnOff();
					var docContent = AscCommon.PasteProcessor.prototype.ReadPresentationText.call(this, stream, worksheet);
					History.TurnOn();

					if (docContent.length === 0) {
						return;
					}
					presentationSelectedContent.DocContent = new AscCommonWord.CSelectedContent();
					presentationSelectedContent.DocContent.Elements = docContent;

					//перебираем шрифты
					for (var i in oThis.oFonts) {
						fonts.push(new AscFonts.CFont(i));
					}
				};

				var readDrawings = function () {

					var objects = oThis.ReadPresentationShapes(stream, worksheet);
					presentationSelectedContent.Drawings = objects.arrShapes;

					var arr_shapes = objects.arrShapes;
					var font_map = {}, i;
					for (i = 0; i < arr_shapes.length; ++i) {
						if (arr_shapes[i].graphicObject.getAllFonts) {
							arr_shapes[i].graphicObject.getAllFonts(font_map);
						}
					}

					for (i in font_map) {
						fonts.push(new AscFonts.CFont(i));
					}

					arr_Images = objects.arrImages;
				};

				var skipIndexes = function () {
					var count = stream.GetULong();

					var array = [];
					for (var i = 0; i < count; ++i) {
						stream.GetULong();
					}
				};

				var skip1 = function () {
					var selected_objs = loader.stream.GetULong();

					for (var i = 0; i < selected_objs; ++i) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
					}
				};

				var first_content = stream.GetString2();
				if (first_content === "SelectedContent") {
					var PresentationWidth = stream.GetULong() / 100000.0;
					var PresentationHeight = stream.GetULong() / 100000.0;
					var countContent = stream.GetULong();
					for (var i = 0; i < countContent; i++) {
						if (null === presentationSelectedContent) {
							presentationSelectedContent = typeof PresentationSelectedContent !== "undefined" ? new PresentationSelectedContent() : {};
						}
						var first_string = stream.GetString2();
						switch (first_string) {
							case "DocContent": {
								readContent(stream, worksheet);
								break;
							}
							case "Drawings": {
								readDrawings();
								break;
							}
							case "SlideObjects": {
								presentationSelectedContent.SlideObjects = [null];
								skip1();
								break;
							}
							case "LayoutsIndexes": {
								skipIndexes();
								break;
							}
							case "MastersIndexes": {
								skipIndexes();
								break;
							}
							case "NotesMastersIndexes": {
								skipIndexes();
								break;
							}
							case "ThemeIndexes": {
								skipIndexes();
								break;
							}
							default: {
								skip1();
							}
						}
					}
				}

				return {content: presentationSelectedContent, fonts: fonts, images: arr_Images};
			},

			_pasteFromBinaryPDF: function (worksheet, base64, isIntoShape, isCellEditMode) {
				pptx_content_loader.Clear();

				let _stream	= AscFormat.CreateBinaryReader(base64, 0, base64.length);
				let stream	= new AscCommon.FileStream(_stream.data, _stream.size);
				let p_url	= stream.GetString2();
				let t = this;

				let bIsMultipleContent = stream.GetBool();

				let selectedContent2 = [], i;
				let multipleParamsCount = stream.GetULong();
				for (i = 0; i < multipleParamsCount; i++) {
					selectedContent2.push(this._readPDFSelectedContent(stream, worksheet));
				}

				let specialOptionsArr = [];
				let specialProps = Asc.c_oSpecialPasteProps;
				if (2 === multipleParamsCount) {
					specialOptionsArr = [specialProps.sourceformatting];
				} else if (3 === multipleParamsCount) {
					specialOptionsArr = [specialProps.sourceformatting, specialProps.picture];
				}

				let pasteObj = selectedContent2[0] ? selectedContent2[0] : selectedContent2[1];

				if (window['AscCommon'].g_specialPasteHelper.specialPasteStart) {
					let props = window['AscCommon'].g_specialPasteHelper.specialPasteProps.property;
					switch (props) {
						case Asc.c_oSpecialPasteProps.picture: {
							if (selectedContent2[2]) {
								pasteObj = selectedContent2[2];
							}
							break;
						}
					}
				}

				let arr_Images, fonts, content = null;
				if (pasteObj) {
					arr_Images = pasteObj.images;
					fonts = pasteObj.fonts;
					content = pasteObj.content;
				}

				if (null === content) {
					window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					return;
				}

				if (content.DocContent) {
					let docContent = content.DocContent.Elements;

					if (isCellEditMode) {
						let text = this._getTextFromWord(docContent);
						window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
						return text;
					} else if (isIntoShape) {
						let callback = function (isSuccess) {
							if (isSuccess) {
								t._insertBinaryIntoShapeContent(worksheet, docContent);
							}
							window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
						};

						worksheet.objectRender.controller.checkPasteInText(callback);
						return true;
					} else {
						History.TurnOff();
						let oPasteFromBinaryWord = new pasteFromBinaryWord(this, worksheet, true);

						let oTempDrawingDocument = window["Asc"]["editor"].wbModel.getDrawingDocument();
						let newCDocument = new CDocument(oTempDrawingDocument, false);
						newCDocument.bFromDocument = true;
						newCDocument.theme = window["Asc"]["editor"].wbModel.theme;

						let newContent = [];
						for (i = 0; i < docContent.length; i++) {
							if (type_Paragraph === docContent[i].GetType())//paragraph
							{
								docContent[i] = AscFormat.ConvertParagraphToWord(docContent[i], newCDocument);
								docContent[i].bFromDocument = true;
								newContent.push(docContent[i]);
							} else if (type_Table === docContent[i].GetType())//table
							{
								//TODO вырезать из таблицы параграфы
							}
						}
						docContent = newContent;

						History.TurnOn();

						oPasteFromBinaryWord._paste(worksheet, {content: docContent});
					}
				} else if (content.Drawings) {

					if (isCellEditMode) {
						return "";
					}

					window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					specialProps = window['AscCommon'].g_specialPasteHelper.buttonInfo;
					if (specialOptionsArr.length > 1) {
						specialProps.asc_setOptions(specialOptionsArr);
					}

					let arr_shapes = content.Drawings;
					if (arr_shapes && arr_shapes.length && !(window["Asc"]["editor"] && window["Asc"]["editor"].isChartEditor)) {
						if (content.Drawings.length === selectedContent2[1].content.Drawings.length) {
							let oEndContent = {
								Drawings: []
							};
							let oSourceContent = {
								Drawings: []
							};
							for (i = 0; i < content.Drawings.length; ++i) {
								oEndContent.Drawings.push({Drawing: content.Drawings[i].graphicObject});
								oSourceContent.Drawings.push({Drawing: selectedContent2[1].content.Drawings[i].graphicObject});
							}
							AscFormat.checkDrawingsTransformBeforePaste(oEndContent, oSourceContent, null);
						}
						let newFonts = {};
						for (i = 0; i < arr_shapes.length; i++) {
							arr_shapes[i].graphicObject.getAllFonts(newFonts);
						}

						let aPastedImages = arr_Images;
						worksheet._loadFonts(newFonts, function () {
							if (aPastedImages && aPastedImages.length) {
								t._loadImagesOnServer(aPastedImages, function () {
									t._insertImagesFromBinary(worksheet, {Drawings: arr_shapes}, isIntoShape, true);
								});
							} else {
								t._insertImagesFromBinary(worksheet, {Drawings: arr_shapes}, isIntoShape, true);
							}
						});
					} else {
						window['AscCommon'].g_specialPasteHelper.buttonInfo.clean();
					}

					return true;
				}
			},

			_readPDFSelectedContent: function (stream, worksheet) {

				var oPdfSelContent = null;
				var fonts = [];
				var arr_Images = {};
				var oThis = this;

				var loader = new AscCommon.BinaryPPTYLoader();
				loader.presentation = worksheet.model;
				loader.DrawingDocument = worksheet.getDrawingDocument();
				loader.stream = stream;

				var readContent = function () {
					History.TurnOff();
					var docContent = AscCommon.PasteProcessor.prototype.ReadPresentationText.call(this, stream, worksheet);
					History.TurnOn();

					if (docContent.length === 0) {
						return;
					}
					oPdfSelContent.DocContent = new AscCommonWord.CSelectedContent();
					oPdfSelContent.DocContent.Elements = docContent;

					//перебираем шрифты
					for (var i in oThis.oFonts) {
						fonts.push(new AscFonts.CFont(i));
					}
				};

				var readDrawings = function () {

					var objects = oThis.ReadPresentationShapes(stream, worksheet);
					oPdfSelContent.Drawings = objects.arrShapes;

					var arr_shapes = objects.arrShapes;
					var font_map = {}, i;
					for (i = 0; i < arr_shapes.length; ++i) {
						if (arr_shapes[i].graphicObject.getAllFonts) {
							arr_shapes[i].graphicObject.getAllFonts(font_map);
						}
					}

					for (i in font_map) {
						fonts.push(new AscFonts.CFont(i));
					}

					arr_Images = objects.arrImages;
				};

				var skip1 = function () {
					var selected_objs = loader.stream.GetULong();

					for (var i = 0; i < selected_objs; ++i) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
					}
				};

				var first_content = stream.GetString2();
				if (first_content === "SelectedContent") {
					var countContent = stream.GetULong();
					for (var i = 0; i < countContent; i++) {
						if (null === oPdfSelContent) {
							oPdfSelContent = window["AscPDF"] && window["AscPDF"].PDFSelectedContent ? new window["AscPDF"].PDFSelectedContent() : {};
						}
						var first_string = stream.GetString2();
						switch (first_string) {
							case "DocContent": {
								readContent(stream, worksheet);
								break;
							}
							case "Drawings": {
								readDrawings();
								break;
							}
							default: {
								skip1();
							}
						}
					}
				}

				return {content: oPdfSelContent, fonts: fonts, images: arr_Images};
			},

			_insertBinaryIntoShapeContent: function (worksheet, content, isConvertToPPTX) {
				if (!content || !content.length) {
					return;
				}

				History.Create_NewPoint();
				History.StartTransaction();

				//ещё раз вызваем getTargetDocContent с флагом true после создания точки в истории(getTargetDocContent добавляет данные в историю)
				var isIntoShape = worksheet.objectRender.controller.getTargetDocContent(true);
				isIntoShape.Remove(1, true, true);

				var insertContent = new AscCommonWord.CSelectedContent();
				var target_doc_content = isIntoShape;

				insertContent.Elements =
					this._convertBeforeInsertIntoShapeContent(worksheet, content, isConvertToPPTX, target_doc_content);

				//TODO конвертирую в текст без форматирую. пеесмотреть!
				var specialPasteProps = window['AscCommon'].g_specialPasteHelper.specialPasteProps;
				if (specialPasteProps && !specialPasteProps.format) {
					for (var i = 0; i < insertContent.Elements.length; i++) {
						insertContent.Elements[i].Element.Clear_TextFormatting();
					}
				}

				this._insertSelectedContentIntoShapeContent(worksheet, insertContent, target_doc_content);

				History.EndTransaction();

				worksheet.objectRender.controller.startRecalculate();
				//for special paste
				this._setShapeSpecialPasteProperties(worksheet, isIntoShape);
			},

			_setShapeSpecialPasteProperties: function (worksheet, isIntoShape) {
				//TODO пока выключаю специальную ставку внутри math, позже доработать и включить
				var oInfo = new CSelectedElementsInfo();
				//var selectedElementsInfo = isIntoShape.GetSelectedElementsInfo(oInfo);
				var mathObj = oInfo.GetMath();

				if (/*!window['AscCommon'].g_specialPasteHelper.specialPasteStart && */null === mathObj) {
					var sProps = Asc.c_oSpecialPasteProps;
					var curShape = isIntoShape.Parent.parent;

					var asc_getcvt = Asc.getCvtRatio;
					var mmToPx = asc_getcvt(3/*px*/, 0/*pt*/, worksheet._getPPIX());

					var cellsLeft = worksheet.cellsLeft;
					var cellsTop = worksheet.cellsTop;

					var cursorPos = isIntoShape.GetCursorPosXY();
					var offsetX = worksheet._getColLeft(worksheet.visibleRange.c1) - cellsLeft;
					var offsetY = worksheet._getRowTop(worksheet.visibleRange.r1) - cellsTop;
					var posX = curShape.transformText.TransformPointX(cursorPos.X, cursorPos.Y) * mmToPx - offsetX +
						cellsLeft;
					var posY = curShape.transformText.TransformPointY(cursorPos.X, cursorPos.Y) * mmToPx - offsetY +
						cellsTop;

					posX = AscCommon.AscBrowser.convertToRetinaValue(posX);
					posY = AscCommon.AscBrowser.convertToRetinaValue(posY);

					var position = {x: posX, y: posY};

					var allowedSpecialPasteProps = [sProps.sourceformatting, sProps.destinationFormatting];

					window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					window['AscCommon'].g_specialPasteHelper.buttonInfo.asc_setOptions(allowedSpecialPasteProps);
					window['AscCommon'].g_specialPasteHelper.buttonInfo.setPosition(position);
					window['AscCommon'].g_specialPasteHelper.buttonInfo.setRange(cursorPos);
					window['AscCommon'].g_specialPasteHelper.buttonInfo.setShapeId(isIntoShape.Id);
				}
			},

			_convertBeforeInsertIntoShapeContent: function (worksheet, content, isConvertToPPTX, target_doc_content) {
				var elements = [], selectedElement, element;
				var bRemoveHyperlink = false;
				if (target_doc_content && target_doc_content.Is_ChartTitleContent &&
					target_doc_content.Is_ChartTitleContent()) {
					bRemoveHyperlink = true;
				}
				for (var i = 0; i < content.length; i++) {
					selectedElement = new AscCommonWord.CSelectedElement();
					element = content[i];

					if (type_Paragraph === element.GetType())//paragraph
					{
						selectedElement.Element =
							AscFormat.ConvertParagraphToPPTX(element, worksheet.getDrawingDocument(),
								target_doc_content, true, bRemoveHyperlink);
						elements.push(selectedElement);
					} else if (type_Table === element.GetType())//table
					{
						//excel (извне и из word) вставляет в шейпы аналогично, а вот из excel в excel вставляет в одну строку(?). мы сделаем для всех случаев одинаково. 
						var paragraphs = [];
						element.GetAllParagraphs({All: true}, paragraphs);
						for (var j = 0; j < paragraphs.length; j++) {
							selectedElement = new AscCommonWord.CSelectedElement();
							selectedElement.Element =
								AscFormat.ConvertParagraphToPPTX(paragraphs[j], worksheet.getDrawingDocument(),
									target_doc_content, true, bRemoveHyperlink);
							elements.push(selectedElement);
						}
					}
				}

				return elements;
			},

			_insertSelectedContentIntoShapeContent: function (worksheet, selectedContent, target_doc_content) {
				var paragraph = target_doc_content.Content[target_doc_content.CurPos.ContentPos];
				if (null != paragraph && type_Paragraph === paragraph.GetType() && selectedContent.Elements && selectedContent.Elements.length) {

					selectedContent.EndCollect(target_doc_content);
					let NearPos = {Paragraph: paragraph, ContentPos: paragraph.Get_ParaContentPos(false, false)};
					paragraph.Check_NearestPos(NearPos);
					selectedContent.Insert(NearPos, false);
					var oTargetTextObject = AscFormat.getTargetTextObject(worksheet.objectRender.controller);
					oTargetTextObject && oTargetTextObject.checkExtentsByDocContent &&
					oTargetTextObject.checkExtentsByDocContent();
					worksheet.objectRender.controller.startRecalculate();
				}
			},

			_insertImagesFromBinary: function (ws, data, isIntoShape, needShowSpecialProps, savePosition, pastedWb) {
				var activeCell = ws.model.selectionRange.activeCell;
				var curCol, drawingObject, curRow, startCol, startRow, xfrm, aImagesSync = [], activeRow, activeCol,
					offX, offY, rot;

				//отдельная обработка для вставки одной таблички из презентаций
				drawingObject = data.Drawings[0];
				if (data.Drawings.length === 1 && typeof AscFormat.CGraphicFrame !== "undefined" &&
					drawingObject.graphicObject instanceof AscFormat.CGraphicFrame) {
					this._insertTableFromPresentation(ws, data.Drawings[0]);
					return;
				}

				if (window["Asc"]["editor"].collaborativeEditing.getGlobalLock() || !window["Asc"]["editor"].canEdit()) {
					return;
				}

				History.Create_NewPoint();
				History.StartTransaction();

				var _abortPaste = function (_error) {
					window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					if (_error) {
						ws.handlers.trigger("onErrorEvent", Asc.c_oAscError.ID.PasteSlicerError, Asc.c_oAscError.Level.NoCritical);
					}
					History.EndTransaction();
				};

				var callback = function (success, slicersNames) {
					if (!success) {
						_abortPaste();
						return;
					}

					//определяем стартовую позицию, если изображений несколько вставляется
					var graphicObject, i;
					for (i = 0; i < data.Drawings.length; i++) {
						drawingObject = data.Drawings[i];
						graphicObject = drawingObject.graphicObject;

						xfrm = graphicObject.spPr.xfrm;
						if (xfrm) {
							offX = 0;
							offY = 0;
							rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
							rot = AscFormat.normalizeRotate(rot);
							if (AscFormat.checkNormalRotate(rot)) {
								if (AscFormat.isRealNumber(xfrm.offX) && AscFormat.isRealNumber(xfrm.offY)) {
									offX = xfrm.offX;
									offY = xfrm.offY;
								}
							} else {
								if (AscFormat.isRealNumber(xfrm.offX) && AscFormat.isRealNumber(xfrm.offY) &&
									AscFormat.isRealNumber(xfrm.extX) && AscFormat.isRealNumber(xfrm.extY)) {
									offX = xfrm.offX + xfrm.extX / 2 - xfrm.extY / 2;
									offY = xfrm.offY + xfrm.extY / 2 - xfrm.extX / 2;
								}
							}

							if (i === 0) {
								startCol = offX;
								startRow = offY;
							} else {
								if (startCol > offX) {
									startCol = offX;
								}
								if (startRow > offY) {
									startRow = offY;
								}
							}
						} else {
							if (i === 0) {
								startCol = drawingObject.from.col;
								startRow = drawingObject.from.row;
							} else {
								if (startCol > drawingObject.from.col) {
									startCol = drawingObject.from.col;
								}
								if (startRow > drawingObject.from.row) {
									startRow = drawingObject.from.row;
								}
							}
						}
					}

					var aCopies = [];
					var oIdMap = {};
					var _slicerCounter = 0;
					var oCopyPr = new AscFormat.CCopyObjectProperties();
					oCopyPr.idMap = oIdMap;
					ws.objectRender.controller.resetSelection();
					let aDrawings = [];
					for (i = 0; i < data.Drawings.length; i++) {

						if (slicersNames && data.Drawings[i].graphicObject.getObjectType() === AscDFH.historyitem_type_SlicerView) {
							if (slicersNames[_slicerCounter]) {
								data.Drawings[i].graphicObject.setName(slicersNames[_slicerCounter]);
							} else {
								continue;
							}
							_slicerCounter++;
						}

						var _copy = data.Drawings[i].graphicObject.copy(oCopyPr);
						if (_copy.convertFromSmartArt) {
							_copy.convertFromSmartArt(true);
						}
						_copy.clearChartDataCache();
						oIdMap[data.Drawings[i].graphicObject.Id] = _copy.Id;
						data.Drawings[i].graphicObject = _copy;
						aCopies.push(data.Drawings[i].graphicObject);

						drawingObject = data.Drawings[i];
						AscFormat.CheckSpPrXfrm2(drawingObject.graphicObject);
						xfrm = drawingObject.graphicObject.spPr.xfrm;

						activeRow = activeCell.row;
						activeCol = activeCell.col;
						if (isIntoShape && isIntoShape.Parent && isIntoShape.Parent.parent &&
							isIntoShape.Parent.parent.drawingBase && isIntoShape.Parent.parent.drawingBase.from) {
							activeRow = isIntoShape.Parent.parent.drawingBase.from.row;
							activeCol = isIntoShape.Parent.parent.drawingBase.from.col;
						}

						if (savePosition) {
							curCol = xfrm.offX;
							curRow = xfrm.offY;
						} else {
							curCol = xfrm.offX - startCol + ws.objectRender.convertMetric(ws._getColLeft(activeCol + 1*ws.getRightToLeft()) - ws._getColLeft(0), 0, 3);
							curRow = xfrm.offY - startRow + ws.objectRender.convertMetric(ws._getRowTop(activeRow) - ws._getRowTop(0), 0, 3);
						}


						drawingObject = ws.objectRender.cloneDrawingObject(drawingObject);
						drawingObject.graphicObject.setDrawingBase(drawingObject);

						drawingObject.graphicObject.setDrawingObjects(ws.objectRender);
						drawingObject.graphicObject.setWorksheet(ws.model);
						let _left = ws.checkRtl(curCol, null, 3);
						if (ws.getRightToLeft()) {
							let mmToPx = Asc.getCvtRatio(3/*px*/, 0/*pt*/, ws._getPPIX());
							let _widthDrawing = drawingObject.getWidthFromTo() / mmToPx;
							let _widthCtx = ws.getCtxWidth() / mmToPx;
							if (_left + _widthDrawing > _widthCtx) {
								_left -= _left + _widthDrawing - _widthCtx;
							}
						}
						xfrm.setOffX( _left);
						xfrm.setOffY(curRow);
						aDrawings.push(drawingObject.graphicObject);
						drawingObject.graphicObject.getAllRasterImages(aImagesSync);
					}

					let dShift = ws.objectRender.controller.getDrawingsPasteShift(aDrawings);
					for (let nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing) {
						let oDrawing = aDrawings[nDrawing];
						if (dShift > 0) {
							oDrawing.shiftXfrm(dShift, dShift);
						}
						oDrawing.addToDrawingObjects();
						oDrawing.checkDrawingBaseCoords();
						oDrawing.generateSmartArtDrawingPart();
						oDrawing.recalculate();
						oDrawing.select(ws.objectRender.controller, 0);
					}

					AscFormat.fResetConnectorsIds(aCopies, oIdMap);
					if (aImagesSync.length > 0) {
						window["Asc"]["editor"].ImageLoader.LoadDocumentImages(aImagesSync);
					}
					ws.objectRender.controller.updateSelectionState();
					ws.objectRender.controller.updatePlaceholders();
					ws.objectRender.showDrawingObjects();

					if (needShowSpecialProps) {
						if (!window['AscCommon'].g_specialPasteHelper.buttonInfo.options) {
							window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
						} else {
							var lastAddedImg = ws.model.Drawings[ws.model.Drawings.length - 1];
							if (drawingObject && lastAddedImg) {
								window['AscCommon'].g_specialPasteHelper.buttonInfo.setRange({
									r1: lastAddedImg.from.row,
									c1: lastAddedImg.from.col,
									r2: lastAddedImg.to.row,
									c2: lastAddedImg.to.col
								});
							}
						}
					}

					ws.objectRender.controller.updateOverlay();
					ws.setSelectionShape(true);
					History.EndTransaction();


					if (!needShowSpecialProps) {
						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					}
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				};

				//вставляем срезы следующим образом: проверяем что вставлем в оригинальный документ
				//далее смотрим, есть ли такой внутренний объект этого кэша - для таблиц проверяем
				//соответствие имени таблицы и имени колонки
				//если все объекты внутренние имеются, проверяем локи, если нет - ошибка
				var pastedInOriginalDoc = this._checkPastedInOriginalDoc(pastedWb)
				var pastedSlicers = [];
				for (var i = 0; i < data.Drawings.length; i++) {
					if (data.Drawings[i].graphicObject.getObjectType() === AscDFH.historyitem_type_SlicerView) {
						if (pastedInOriginalDoc) {
							var pastedSlicer = data.getSlicerByName(data.Drawings[i].graphicObject.name);
							if (pastedSlicer) {
								if (pastedSlicer.checkModelContent(ws.model)) {
									pastedSlicers.push(pastedSlicer);
								} else {
									//error
									_abortPaste(true);
									return;
								}
							}
						} else {
							//error
							_abortPaste(true);
							return;
						}
					}
				}

				if (pastedSlicers && pastedSlicers.length) {
					ws.tryPasteSlicers(pastedSlicers, callback);
				} else {
					callback(true);
				}
			},

			_insertImagesFromBinaryWord: function (ws, data, aImagesSync) {
				var activeRange = ws.model.selectionRange.getLast().clone();
				var curCol, drawingObject, curRow, startCol = 0, startRow = 0, xfrm, drawingBase, graphicObject, offX,
					offY, rot;

				History.Create_NewPoint();
				History.StartTransaction();
				var api = window["Asc"]["editor"];
				var addImagesFromWord = data.props.addImagesFromWord;
				//определяем стартовую позицию, если изображений несколько вставляется
				let aDrawings = [];
				for (var i = 0; i < addImagesFromWord.length; i++) {
					if (para_Math === addImagesFromWord[i].image.Type) {
						graphicObject = ws.objectRender.createShapeAndInsertContent(addImagesFromWord[i].image);
					} else {
						graphicObject = addImagesFromWord[i].image.GraphicObj;

						//convert from word
						if (graphicObject.setBDeleted2) {
							graphicObject.setBDeleted2(true);
						} else {
							graphicObject.bDeleted = true;
						}
						graphicObject = graphicObject.convertToPPTX(ws.getDrawingDocument(), ws.model, true);
					}


					//create new drawingBase
					drawingObject = ws.objectRender.createDrawingObject();
					drawingObject.graphicObject = graphicObject;

					if (graphicObject.spPr && graphicObject.spPr.xfrm) {
						xfrm = graphicObject.spPr.xfrm;
						offX = 0;
						offY = 0;
						rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
						rot = AscFormat.normalizeRotate(rot);
						if (AscFormat.checkNormalRotate(rot)) {
							if (AscFormat.isRealNumber(xfrm.offX) && AscFormat.isRealNumber(xfrm.offY)) {
								offX = xfrm.offX;
								offY = xfrm.offY;
							}
						} else {
							if (AscFormat.isRealNumber(xfrm.offX) && AscFormat.isRealNumber(xfrm.offY) &&
								AscFormat.isRealNumber(xfrm.extX) && AscFormat.isRealNumber(xfrm.extY)) {
								offX = xfrm.offX + xfrm.extX / 2 - xfrm.extY / 2;
								offY = xfrm.offY + xfrm.extY / 2 - xfrm.extX / 2;
							}
						}

						if (i === 0) {
							startCol = offX;
							startRow = offY;
						} else {
							if (startCol > offX) {
								startCol = offX;
							}
							if (startRow > offY) {
								startRow = offY;
							}
						}
					} else {
						if (i === 0) {
							startCol = drawingObject.from.col;
							startRow = drawingObject.from.row;
						} else {
							if (startCol > drawingObject.from.col) {
								startCol = drawingObject.from.col;
							}
							if (startRow > drawingObject.from.row) {
								startRow = drawingObject.from.row;
							}
						}
					}


					AscFormat.CheckSpPrXfrm2(graphicObject);
					xfrm = graphicObject.spPr.xfrm;

					curCol = xfrm.offX - startCol + ws.objectRender.convertMetric(ws._getColLeft(addImagesFromWord[i].col + activeRange.c1) - ws._getColLeft(0), 0, 3);
					curRow = xfrm.offY - startRow + ws.objectRender.convertMetric(ws._getRowTop(addImagesFromWord[i].row + activeRange.r1) - ws._getRowTop(0), 0, 3);

					xfrm.setOffX(curCol);
					xfrm.setOffY(curRow);

					drawingObject = ws.objectRender.cloneDrawingObject(drawingObject);
					if (graphicObject.convertFromSmartArt) {
						graphicObject.convertFromSmartArt(true);
					}
					graphicObject.clearChartDataCache();
					graphicObject.setDrawingBase(drawingObject);
					graphicObject.setDrawingObjects(ws.objectRender);
					graphicObject.setWorksheet(ws.model);
					var nAnchorType = AscCommon.c_oAscCellAnchorType.cellanchorTwoCell;
					if (graphicObject.getObjectType() === AscDFH.historyitem_type_ImageShape) {
						nAnchorType = AscCommon.c_oAscCellAnchorType.cellanchorOneCell;
					}
					graphicObject.setDrawingBaseType(nAnchorType);
					if (graphicObject.checkExtentsByDocContent) {
						graphicObject.checkDrawingBaseCoords();
						graphicObject.generateSmartArtDrawingPart();
						graphicObject.checkExtentsByDocContent();
					}
					aDrawings.push(graphicObject);
				}
				let dShift = ws.objectRender.controller.getDrawingsPasteShift(aDrawings);
				ws.objectRender.controller.resetSelection();
				for (let nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing) {
					let oDrawing = aDrawings[nDrawing];
					if (dShift > 0) {
						oDrawing.shiftXfrm(dShift, dShift);
					}
					oDrawing.addToDrawingObjects();
					oDrawing.checkDrawingBaseCoords();
					oDrawing.recalculate();
					oDrawing.select(ws.objectRender.controller, 0);
				}

				var old_val = api.ImageLoader.bIsAsyncLoadDocumentImages;
				api.ImageLoader.bIsAsyncLoadDocumentImages = true;
				api.ImageLoader.LoadDocumentImages(aImagesSync);
				api.ImageLoader.bIsAsyncLoadDocumentImages = old_val;

				ws.objectRender.showDrawingObjects();
				ws.setSelectionShape(true);
				ws.objectRender.controller.updatePlaceholders();
				ws.objectRender.controller.updateOverlay();
				History.EndTransaction();
			},


			_loadImagesOnServer: function (aPastedImages, callback) {
				var api = Asc["editor"];

				var oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aPastedImages);

				AscCommon.sendImgUrls(api, oObjectsForDownload.aUrls, function (data) {
					var oImageMap = {};

					History.TurnOff();
					AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
					History.TurnOn();

					callback();
				});
			},

			_insertTableFromPresentation: function (ws, graphicFrame) {
				History.TurnOff();
				//из-за данной строки падение при вставке таблицы pe->se
				//graphicFrame.graphicObject = graphicFrame.graphicObject.copy();
				var drawingObject = graphicFrame;

				//вставляем табличку из презентаций
				var oPasteFromBinaryWord = new pasteFromBinaryWord(this, ws, true);
				var oTempDrawingDocument = ws.getDrawingDocument();

				var newCDocument = new CDocument(oTempDrawingDocument, false);
				newCDocument.bFromDocument = true;
				newCDocument.theme = window["Asc"]["editor"].wbModel.theme;

				drawingObject.graphicObject.setBDeleted(true);
				drawingObject.graphicObject.setWordFlag(false, newCDocument);

				var oldLogicDocument = oTempDrawingDocument.m_oLogicDocument;
				oTempDrawingDocument.m_oLogicDocument = newCDocument;
				drawingObject.graphicObject.graphicObject.Set_Parent(newCDocument);
				oTempDrawingDocument.m_oLogicDocument = oldLogicDocument;

				History.TurnOn();

				oPasteFromBinaryWord._paste(ws, {content: [drawingObject.graphicObject.graphicObject]});
			},

			editorPasteExec: function (worksheet, node, isText, cb) {
				if (node === undefined) {
					return;
				}

				var binaryResult, t = this;
				t.alreadyLoadImagesOnServer = false;

				//если находимся внутри шейпа
				var isIntoShape = worksheet.objectRender.controller.getTargetDocContent();
				if (isIntoShape) {
					var callback = function (isSuccess) {
						if (isSuccess) {
							t._pasteInShape(worksheet, node, isIntoShape);
						} else {
							window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
						}
					};

					worksheet.objectRender.controller.checkPasteInText(callback);
					return;
				}

				//****binary****
				if (copyPasteUseBinary) {
					this.activeRange = worksheet.model.selectionRange.getLast().clone();
					binaryResult = this._pasteFromBinaryClassHtml(worksheet, node, isIntoShape);

					if (binaryResult === true) {
						return;
					} else if (binaryResult !== false && binaryResult !== undefined) {
						node = binaryResult;
					}
				}

				this.pasteCallBack = cb;

				this.activeRange = worksheet.model.selectionRange.getLast().clone();
				worksheet.workbook.handlers.trigger("cleanCopyData", true);

				var callBackAfterLoadImages = function () {
					History.TurnOff();

					var oTempDrawingDocument = worksheet.getDrawingDocument();
					var newCDocument = new CDocument(oTempDrawingDocument, false);
					newCDocument.bFromDocument = true;

					//TODo!!!!!!
					newCDocument.Content[0].bFromDocument = true;
					newCDocument.theme = window["Asc"]["editor"].wbModel.theme;

					var oOldLogicDocument = oTempDrawingDocument.m_oLogicDocument;
					oTempDrawingDocument.m_oLogicDocument = newCDocument;
					var oOldEditor = undefined;
					if ("undefined" !== typeof editor) {
						oOldEditor = editor;
					}
					editor = {WordControl: oTempDrawingDocument, isDocumentEditor: true};
					var oPasteProcessor = new AscCommon.PasteProcessor({
						WordControl: {m_oLogicDocument: newCDocument},
						FontLoader: {}
					}, false, false, null, true);
					oPasteProcessor._Prepeare_recursive(node, true, true);

					//при специальной вставке в firefox _getComputedStyle возвращает null
					//TODO пересмотреть функцию _getComputedStyle
					let specialPasteHelper = window['AscCommon'].g_specialPasteHelper;
					let specialPasteProps = specialPasteHelper.specialPasteProps;
					let props = specialPasteProps ? specialPasteProps.property : null;

					if (window['AscCommon'].g_specialPasteHelper.specialPasteStart && window['AscCommon'].g_specialPasteHelper.specialPasteData.aContent && props !== Asc.c_oSpecialPasteProps.picture) {
						oPasteProcessor.aContent = window['AscCommon'].g_specialPasteHelper.specialPasteData.aContent;
					} else {
						oPasteProcessor._Execute(node, {}, true, true, false);
						if (props !== Asc.c_oSpecialPasteProps.picture) {
							window['AscCommon'].g_specialPasteHelper.specialPasteData.aContent = oPasteProcessor.aContent;
						}
					}

					editor = oOldEditor;
					oTempDrawingDocument.m_oLogicDocument = oOldLogicDocument;
					History.TurnOn();

					var oPasteFromBinaryWord = new pasteFromBinaryWord(t, worksheet);
					oPasteFromBinaryWord._paste(worksheet, {content: oPasteProcessor.aContent});
				};

				var aImagesToDownload = this._getImageFromHtml(node, true);
				var specialPasteProps = window['AscCommon'].g_specialPasteHelper.specialPasteProps;
				var api = Asc["editor"];
				if (!api.isChartEditor && aImagesToDownload !== null &&
					(!specialPasteProps || (specialPasteProps && specialPasteProps.images)))//load to server
				{
					AscCommon.sendImgUrls(api, aImagesToDownload, function (data) {
						for (var i = 0, length = Math.min(data.length, aImagesToDownload.length); i < length; ++i) {
							var elem = data[i];
							var sFrom = aImagesToDownload[i];
							if (null != elem.url) {
								t.oImages[sFrom] = AscCommon.g_oDocumentUrls.imagePath2Local(elem.path);
							} else {
								t.oImages[sFrom] = sFrom;
							}
						}

						t.alreadyLoadImagesOnServer = true;
						callBackAfterLoadImages();

					});

				} else {
					callBackAfterLoadImages();
				}
			},

			_pasteFromBinaryClassHtml: function (worksheet, node, isIntoShape) {
				//find class xslData or docData
				var returnBinary = this._getClassBinaryFromHtml(node);
				var base64 = returnBinary.base64;
				var base64FromWord = returnBinary.base64FromWord;
				var base64FromPresentation = returnBinary.base64FromPresentation;

				var result = false;
				if (base64 != null)//from excel
				{
					result = this._pasteFromBinaryExcel(worksheet, base64, isIntoShape);
				} else if (base64FromWord)//from word
				{
					result = this._pasteFromBinaryWord(worksheet, base64FromWord, isIntoShape);
					worksheet.workbook.handlers.trigger("cleanCopyData", true);
				} else if (base64FromPresentation) {
					result = this._pasteFromBinaryPresentation(worksheet, base64FromPresentation, isIntoShape);
					worksheet.workbook.handlers.trigger("cleanCopyData", true);
				}

				return result;
			},

			_getClassBinaryFromHtml: function (node) {
				var base64 = null, base64FromWord = null, base64FromPresentation = null, base64FromPDF;
				var classNode = AscCommon.searchBinaryClass(node);

				if (classNode != null) {
					var cL = classNode.split(" ");
					for (var i = 0; i < cL.length; i++) {
						if (cL[i].indexOf("xslData;") > -1) {
							base64 = cL[i].split('xslData;')[1];
						} else if (cL[i].indexOf("docData;") > -1) {
							base64FromWord = cL[i].split('docData;')[1];
						} else if (cL[i].indexOf("pptData;") > -1) {
							base64FromPresentation = cL[i].split('pptData;')[1];
						} else if (cL[i].indexOf("pdfData;") > -1) {
							base64FromPDF = cL[i].split('pdfData;')[1];
						}
					}
				}

				return {base64: base64, base64FromWord: base64FromWord, base64FromPresentation: base64FromPresentation};
			},

			_getImageFromHtml: function (html, isGetUrlsArray) {
				var res = null;

				if (html) {
					var allImages = html.getElementsByTagName('img');

					if (allImages && allImages.length) {
						if (isGetUrlsArray) {
							var arrayImages = [];
							for (var i = 0; i < allImages.length; i++) {
								arrayImages[i] = allImages[i].src;
							}

							res = arrayImages;
						} else {
							res = allImages;
						}
					}
				}

				return res;
			},

			_getBinaryColor: function (c) {
				var bin, m, x, type, r, g, b, a, s;
				var reColor = /^\s*(?:#?([0-9a-f]{6})|#?([0-9a-f]{3})|rgba?\s*\(\s*((?:\d*\.?\d+)(?:\s*,\s*(?:\d*\.?\d+)){2,3})\s*\))\s*$/i;
				if (typeof c === "number") {
					bin = c;
				} else {

					m = reColor.exec(c);
					if (!m) {
						return null;
					}

					if (m[1]) {
						x = [m[1].slice(0, 2), m[1].slice(2, 4), m[1].slice(4)];
						type = 1;
					} else if (m[2]) {
						x = [m[2].slice(0, 1), m[2].slice(1, 2), m[2].slice(2)];
						type = 0;
					} else {
						x = m[3].split(/\s*,\s*/i);
						type = x.length === 3 ? 2 : 3;
					}

					r = parseInt(type !== 0 ? x[0] : x[0] + x[0], type < 2 ? 16 : 10);
					g = parseInt(type !== 0 ? x[1] : x[1] + x[1], type < 2 ? 16 : 10);
					b = parseInt(type !== 0 ? x[2] : x[2] + x[2], type < 2 ? 16 : 10);
					a = type === 3 ? (Math.round(parseFloat(x[3]) * 100) * 0.01) : 1;
					bin = (r << 16) | (g << 8) | b;

				}
				return bin;
			},

			_checkMaxTextLength: function (aResult, r, c) {
				var result = false;
				var isChange = false;

				var currentCellData = aResult.content[r][c];
				if (currentCellData && currentCellData.content) {
					currentCellData = currentCellData.content;
					for (var i = 0; i < currentCellData.length; i++) {
						if (currentCellData[i] && currentCellData[i].text && currentCellData[i].text.length > c_oAscMaxCellOrCommentLength) {
							isChange = true;

							var text = currentCellData[i].text;
							var format = currentCellData[i].format;
							var lengthOfText = text.length;
							var iterCount = Math.ceil(lengthOfText / c_oAscMaxCellOrCommentLength);
							var splitText;
							for (var j = 0; j < iterCount; j++) {
								splitText = text.substr(c_oAscMaxCellOrCommentLength * j, c_oAscMaxCellOrCommentLength * (j + 1));
								if (!aResult.content[r]) {
									aResult.content[r] = [];
								}
								if (!aResult.content[r][c]) {
									aResult.content[r][c] = [];
								}
								if (!aResult.content[r][c].content) {
									aResult.content[r][c].content = [];
								}

								aResult.content[r][c].content[0] = {format: format, text: splitText};

								if (iterCount !== j + 1) {
									r++;
								}
							}
						}
					}
					if (isChange) {
						result = {aResult: aResult, r: r, c: c};
					}
				}

				return result;
			},

			_getBorderStyleName: function (borderStyle, borderWidth) {
				var res = null;
				var nBorderWidth = parseFloat(borderWidth);
				if (isNaN(nBorderWidth)) {
					return res;
				}
				if (typeof borderWidth === "string" && -1 !== borderWidth.indexOf("pt")) {
					nBorderWidth = nBorderWidth * 96 / 72;
				}

				switch (borderStyle) {
					case "solid":
						if (0 < nBorderWidth && nBorderWidth <= 1) {
							res = c_oAscBorderStyles.Thin;
						} else if (1 <
							nBorderWidth && nBorderWidth <= 2) {
							res = c_oAscBorderStyles.Medium;
						} else if (2 <
							nBorderWidth && nBorderWidth <= 6) {
							res = c_oAscBorderStyles.Thick;
						} else {
							res = c_oAscBorderStyles.None;
						}
						break;
					case "dashed":
						if (0 < nBorderWidth && nBorderWidth <= 1) {
							res = c_oAscBorderStyles.DashDot;
						} else {
							res = c_oAscBorderStyles.MediumDashDot;
						}
						break;
					case "double":
						res = c_oAscBorderStyles.Double;
						break;
					case "dotted":
						res = c_oAscBorderStyles.Hair;
						break;
				}
				return res;
			},

			ReadPresentationShapes: function (stream, worksheet) {
				History.TurnOff();

				var loader = new AscCommon.BinaryPPTYLoader();
				loader.presentation = worksheet.model;
				loader.DrawingDocument = worksheet.getDrawingDocument();
				loader.Start_UseFullUrl();

				loader.ClearConnectedObjects();
				loader.stream = stream;

				var count = stream.GetULong();
				var arr_shapes = [];
				var arr_transforms = [];
				var arrBase64Img = [];
				var cStyle;

				for (var i = 0; i < count; ++i) {
					var style_index = null;
					//читаем флаг о наличии табличного стиля
					if (!loader.stream.GetBool()) {
						if (loader.stream.GetBool()) {
							loader.stream.Skip2(1);
							cStyle = loader.ReadTableStyle();

							loader.stream.GetBool();
							style_index = stream.GetString2();
						}
					}

					var drawing = loader.ReadGraphicObject();
					var isGraphicFrame = typeof AscFormat.CGraphicFrame !== "undefined" && drawing instanceof AscFormat.CGraphicFrame;

					//для случая, когда копируем 1 таблицу из презентаций, в бинарник заносим ещё одну такую же табличку, но со скомпиоированными стилями(для вставки в word / excel)
					if (count === 1 && isGraphicFrame) {
						drawing = loader.ReadGraphicObject();
					}

					var x = stream.GetULong() / 100000;
					var y = stream.GetULong() / 100000;
					var extX = stream.GetULong() / 100000;
					var extY = stream.GetULong() / 100000;
					var base64 = stream.GetString2();

					if (count !== 1 && isGraphicFrame) {
						drawing = AscFormat.DrawingObjectsController.prototype.createImage(base64, x, y, extX, extY);
						arrBase64Img.push(
							new AscCommon.CBuilderImages(drawing.blipFill, base64, drawing, drawing.spPr, null));
					}

					var oDrawingBase = worksheet.objectRender.createDrawingObject();
					oDrawingBase.graphicObject = drawing;
					var nAnchorType = AscCommon.c_oAscCellAnchorType.cellanchorTwoCell;
					if (drawing.getObjectType() === AscDFH.historyitem_type_ImageShape) {
						nAnchorType = AscCommon.c_oAscCellAnchorType.cellanchorOneCell;
					}
					oDrawingBase.Type = nAnchorType;
					arr_shapes[i] = oDrawingBase;
				}
				loader.AssignConnectedObjects();
				History.TurnOn();

				var arrImages = arrBase64Img.concat(loader.End_UseFullUrl());
				return {arrShapes: arr_shapes, arrImages: arrImages, arrTransforms: arr_transforms};
			},

			ReadFromBinaryWord: function (sBase64, worksheet) {
				History.TurnOff();
				AscCommon.g_oIdCounter.m_bRead = true;

				//передавать CDrawingDocument текущего worksheet
				var oTempDrawingDocument = worksheet.getDrawingDocument();

				var newCDocument = new CDocument(oTempDrawingDocument, false);
				newCDocument.bFromDocument = true;
				newCDocument.theme = window["Asc"]["editor"].wbModel.theme;

				var old_m_oLogicDocument = oTempDrawingDocument.m_oLogicDocument;
				oTempDrawingDocument.m_oLogicDocument = newCDocument;

				var oOldEditor = undefined;
				if ("undefined" !== typeof editor) {
					oOldEditor = editor;
				}
				//создается глобальная переменная
				editor = {isDocumentEditor: true, WordControl: {m_oLogicDocument: newCDocument}};

				pptx_content_loader.Clear();
				pptx_content_loader.Start_UseFullUrl();

				var openParams = {checkFileSize: false, charCount: 0, parCount: 0, bCopyPaste: true};
				var oBinaryFileReader = new AscCommonWord.BinaryFileReader(newCDocument, openParams);
				var oRes = oBinaryFileReader.ReadFromString(sBase64, {excelCopyPaste: true});

				//check parents
				if (oRes.content) {
					for (let i = 0; i < oRes.content.length; i++) {
						if (oRes.content[i] && oRes.content[i].Parent == null) {
							oRes.content[i].Parent = newCDocument;
						}
					}
				}

				var defrPr = oBinaryFileReader.oReadResult && oBinaryFileReader.oReadResult.DefrPr;
				if (defrPr && newCDocument.Styles && newCDocument.Styles.Default && newCDocument.Styles.Default.TextPr) {
					newCDocument.Styles.Default.TextPr.FontSize = defrPr.FontSize;
					if (defrPr.RFonts && defrPr.RFonts.Ascii !== undefined) {
						newCDocument.Styles.Default.TextPr.RFonts = defrPr.RFonts;
					}
				}

				pptx_content_loader.End_UseFullUrl();

				oTempDrawingDocument.m_oLogicDocument = old_m_oLogicDocument;
				History.TurnOn();
				AscCommon.g_oIdCounter.m_bRead = false;
				editor = oOldEditor;

				return oRes;
			},

			_checkPasteFromBinaryExcel: function (worksheet, isWriteError, insertWorksheet) {
				var t = this;
				var activeCellsPasteFragment = false;
				AscCommonExcel.executeInR1C1Mode(false, function () {
					activeCellsPasteFragment = AscCommonExcel.g_oRangeCache.getAscRange(t.activeRange);
				});

				var lastRange = worksheet.model.selectionRange.getLast();
				var rMax = (activeCellsPasteFragment.r2 - activeCellsPasteFragment.r1) + lastRange.r1;
				var cMax = (activeCellsPasteFragment.c2 - activeCellsPasteFragment.c1) + lastRange.c1;
				var res = true;

				//если область вставки выходит за пределы доступной области
				if (cMax > AscCommon.gc_nMaxCol0 || rMax > AscCommon.gc_nMaxRow0) {
					if (isWriteError) {
						worksheet.handlers.trigger("onErrorEvent", Asc.c_oAscError.ID.PasteMaxRangeError, Asc.c_oAscError.Level.NoCritical);
					}

					res = false;
				} else if (worksheet.handlers && !worksheet.handlers.trigger("getLockDefNameManagerStatus") && insertWorksheet && insertWorksheet.TableParts && insertWorksheet.TableParts.length) {
					//если пытаемся вставить вторым пользователем форматированную таблицу, когда первый уже добавил другую форматированную таблицу
					worksheet.handlers.trigger("onErrorEvent", c_oAscError.ID.LockCreateDefName,
						c_oAscError.Level.NoCritical);

					res = false;
				}

				return res;
			},

			_pasteInShape: function (worksheet, node, targetDocContent) {
				var t = this;
				targetDocContent.DrawingDocument.m_oLogicDocument = null;

				var oPasteProcessor = new AscCommon.PasteProcessor({
					WordControl: {m_oLogicDocument: targetDocContent},
					FontLoader: {}
				}, false, false, true, true);

				oPasteProcessor.map_font_index = window["Asc"]["editor"].FontLoader.map_font_index;
				oPasteProcessor.bIsDoublePx = false;
				oPasteProcessor.pasteInPresentationShape = true;

				var newFonts;
				//если находимся внутри диаграммы убираем ссылки
				if (targetDocContent && targetDocContent.Parent && targetDocContent.Parent.parent && targetDocContent.Parent.parent.chart) {
					var changeTag = $(node).find("a");
					this._changeHtmlTag(changeTag);
				}

				oPasteProcessor._Prepeare_recursive(node, true, true);

				oPasteProcessor.aContent = [];

				newFonts = this._convertFonts(oPasteProcessor.oFonts);


				History.StartTransaction();
				oPasteProcessor._Execute(node, {}, true, true, false);
				if (!oPasteProcessor.aContent || !oPasteProcessor.aContent.length) {
					History.EndTransaction();
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					return false;
				}

				var targetContent = worksheet.objectRender.controller.getTargetDocContent(true);//нужно для заголовков диаграмм
				targetContent.Remove(1, true, true);

				worksheet._loadFonts(newFonts, function () {

					//TODO конвертирую в текст без форматирую. пеесмотреть!
					var specialPasteProps = window['AscCommon'].g_specialPasteHelper.specialPasteProps;
					if (specialPasteProps && !specialPasteProps.format) {
						for (var i = 0; i < oPasteProcessor.aContent.length; i++) {
							oPasteProcessor.aContent[i].Clear_TextFormatting();
						}
					}

					oPasteProcessor.InsertInPlace(targetContent, oPasteProcessor.aContent);
					var oTargetTextObject = AscFormat.getTargetTextObject(worksheet.objectRender.controller);
					oTargetTextObject && oTargetTextObject.checkExtentsByDocContent && oTargetTextObject.checkExtentsByDocContent();
					worksheet.objectRender.controller.startRecalculate();
					History.EndTransaction();

					t._setShapeSpecialPasteProperties(worksheet, targetContent);
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				});

				return true;
			},

			_convertFonts: function (oFonts) {
				var newFonts = {};
				var fontName;
				for (var i in oFonts) {
					fontName = oFonts[i] ? (oFonts[i].Name || oFonts[i].name) : undefined;
					if (undefined !== fontName) {
						newFonts[fontName] = 1;
					}
				}
				return newFonts;
			},

			_convertTableFromExcelToDocument: function (worksheet, aContentExcel, documentContent) {
				var oCurPar = new AscWord.Paragraph(documentContent);

				var getElem = function (text, format, isAddSpace, isHyperLink) {
					var result = null;
					var value = text;
					if (isAddSpace) {
						value += " ";
					}
					if ("" === value) {
						return result;
					}


					if (isHyperLink) {
						var oCurHyperlink = new ParaHyperlink();
						oCurHyperlink.SetParagraph(oCurPar);
						var sHValue = "";
						if (typeof isHyperLink.Hyperlink === "string") {
							sHValue = isHyperLink.Hyperlink;
						}
						oCurHyperlink.Set_Value(sHValue);
						if (isHyperLink.Tooltip) {
							oCurHyperlink.SetToolTip(isHyperLink.Tooltip);
						}
					}

					var oCurRun = new ParaRun(oCurPar);
					//***text property***
					if (format) {
						oCurRun.Pr.Bold = format.getBold();
						oCurRun.Pr.Italic = format.getItalic();
						oCurRun.Pr.Strikeout = format.getStrikeout();
						oCurRun.Pr.Underline = format.getUnderline() !== 2;
					}

					for (var k = 0, length = value.length; k < length; k++) {
						var nUnicode = null;
						var nCharCode = value.charCodeAt(k);
						if (AscCommon.isLeadingSurrogateChar(nCharCode)) {
							if (k + 1 < length) {
								k++;
								var nTrailingChar = value.charCodeAt(k);
								nUnicode = AscCommon.decodeSurrogateChar(nCharCode, nTrailingChar);
							}
						} else {
							nUnicode = nCharCode;
						}

						if (null != nUnicode) {
							var Item;
							if (0x20 !== nUnicode && 0xA0 !== nUnicode && 0x2009 !== nUnicode) {
								Item = new AscWord.CRunText(nUnicode);
							} else {
								Item = new AscWord.CRunSpace();
							}

							//add text
							oCurRun.Add_ToContent(k, Item, false);
						}
					}

					//add run
					if (oCurHyperlink) {
						oCurHyperlink.Add_ToContent(0, oCurRun, false);
						result = oCurHyperlink;
					} else {
						result = oCurRun;
					}

					return result;
				};

				var n = 0;
				aContentExcel._forEachCell(function (cell) {
					var isHyperlink = aContentExcel.getCell3(cell.nRow, cell.nCol).getHyperlink();

					var multiText = cell.getValueMultiText(), format, elem;
					if (multiText) {
						for (var m = 0; m < multiText.length; m++) {
							var curMultiText = multiText[m];
							format = curMultiText.format;

							elem = getElem(curMultiText.text, format);
							if (null !== elem) {
								oCurPar.Internal_Content_Add(n, elem, false);
								n++;
							}
						}
					} else {
						format = cell.xfs && cell.xfs.font ? cell.xfs.font : null;

						elem = getElem(cell.getValue(), format, null, isHyperlink);
						if (null !== elem) {
							oCurPar.Internal_Content_Add(n, elem, false);
							n++;
						}
					}
					//add space
					elem = getElem("", null, true);
					oCurPar.Internal_Content_Add(n, elem, false);
					n++;
				});

				return oCurPar;
			},

			//TODO сделать при получаении структуры. игнорировать размер шрифта и тп
			_changeHtmlTag: function (arr) {
				var oldElem, value, style, bold, underline, italic;
				for (var i = 0; i < arr.length; i++) {
					oldElem = arr[i];
					value = oldElem.innerText;

					underline = "none";
					if (oldElem.style.textDecoration && oldElem.style.textDecoration != "") {
						underline = oldElem.style.textDecoration;
					}

					italic = "normal";
					if (oldElem.style.textDecoration && oldElem.style.textDecoration != "") {
						italic = oldElem.style.fontStyle;
					}

					bold = "normal";
					if (oldElem.style.fontWeight && oldElem.style.fontWeight != "") {
						bold = oldElem.style.fontWeight;
					}

					style = ' style = "text-decoration:' + underline + ";" + "font-style:" + italic + ";" + "font-weight:" + bold + ";" + '"';
					$(oldElem).replaceWith("<span" + style + ">" + value + "</span>");
				}
			},

			pasteTextOnSheet: function (worksheet, text) {
				var t = this;
				if (!text || (text && !text.length)) {
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					return;
				}

				//TEXT IMPORT WITH OPTIONS
				var specialPasteHelper = window['AscCommon'].g_specialPasteHelper;
				var specialPasteProps = specialPasteHelper.specialPasteProps;
				var textImport;

				var props = specialPasteProps ? specialPasteProps.property : null;
				textImport = props === c_oSpecialPasteProps.useTextImport;


				//TODO сделать вставку текста всегда через эту функцию
				this.activeRange = worksheet.model.selectionRange.getLast().clone(true);

				//если находимся внутри шейпа
				var isIntoShape = worksheet.objectRender.controller.getTargetDocContent();
				if (isIntoShape) {
					var callback = function (isSuccess) {
						if (isSuccess === false) {
							return false;
						}

						isIntoShape = worksheet.objectRender.controller.getTargetDocContent(true);
						if (!isIntoShape) {
							return false;
						}
						isIntoShape.Remove(1, true, true);
						var Count = text.length;

						var newParagraph = new AscWord.Paragraph(isIntoShape);
						var selectedElements = new AscCommonWord.CSelectedContent();
						var insertText = "";
						for (var Index = 0; Index < Count; Index++) {
							var _char = text.charAt(Index);
							var _charCode = text.charCodeAt(Index);

							if (_charCode !== 0x0A) {
								insertText += _char;
							}

							if (_charCode === 0x0A || Index === Count - 1) {
								var newParaRun = new ParaRun();
								window['AscCommon'].addTextIntoRun(newParaRun, insertText);
								newParagraph.Internal_Content_Add(newParagraph.Content.length - 1, newParaRun, false);
								var selectedElement = new AscCommonWord.CSelectedElement();
								selectedElement.Element = newParagraph;
								selectedElements.Elements.push(selectedElement);

								insertText = "";
								newParagraph = new AscWord.Paragraph(isIntoShape);
							}
						}

						t._insertSelectedContentIntoShapeContent(worksheet, selectedElements, isIntoShape);

						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
						window['AscCommon'].g_specialPasteHelper.Paste_Process_End();

						//for special paste
						if (!window['AscCommon'].g_specialPasteHelper.specialPasteStart) {
							var sProps = Asc.c_oSpecialPasteProps;
							var allowedSpecialPasteProps = [sProps.sourceformatting, sProps.destinationFormatting];
						}
					};
					//check text
					AscFonts.FontPickerByCharacter.getFontsByString(text);
					worksheet._loadFonts([], function () {
						worksheet.objectRender.controller.checkPasteInText(callback);
					});
					return;
				}

				if (textImport) {
					var advancedOptions = specialPasteProps.asc_getAdvancedOptions();
					if (Asc.typeOf(text) !== "array") {
						text = AscCommon.parseText(text, advancedOptions, true);
					}
				}
				var aResult = this._getTableFromText(text, textImport);
				if (aResult && !(aResult.onlyImages && window["Asc"]["editor"] && window["Asc"]["editor"].isChartEditor)) {
					if (textImport) {
						var arn = worksheet.model.selectionRange.getLast().clone();
						var width = aResult.content && aResult.content[0] ? aResult.content[0].length - 1 : 0;
						var height = aResult.content ? aResult.content.length - 1 : 0;
						var arnTo = new Asc.Range(arn.c1, arn.r1, arn.c1 + width, arn.r1 + height);

						var resmove = worksheet.model._prepareMoveRange(arn, arnTo);
						if (resmove === -2) {
							window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
							worksheet.handlers.trigger("onErrorEvent", c_oAscError.ID.CannotMoveRange, c_oAscError.Level.NoCritical);
						} else if (resmove === -1) {
							worksheet.model.workbook.handlers.trigger("asc_onConfirmAction", Asc.c_oAscConfirm.ConfirmReplaceRange,
								function (can) {
									if (can) {
										worksheet.setSelectionInfo('paste', {data: aResult, bText: true});
									} else {
										window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
									}
								});
						} else {
							worksheet.setSelectionInfo('paste', {data: aResult, bText: true});
						}
					} else {
						worksheet.setSelectionInfo('paste', {data: aResult, bText: true});
					}
				}
			},

			_getTextFromWorksheet: function (worksheet) {
				var res = "";
				var curRow = -1;
				worksheet._forEachCell(function (cell) {
					if (curRow !== cell.nRow) {
						if (-1 !== curRow) {
							res += '\r\n';
						}
						curRow = cell.nRow;
					} else {
						res += " ";
					}
					res += cell.getValue();
				});
				return res;
			},

			_getTextFromWord: function (data) {
				var res = "";

				var getTextFromCell = function (cell) {
					if (cell.Content && cell.Content.Content) {
						getTextFromDocumentContent(cell.Content.Content, true);
					}
				};

				var getTextFromTable = function (table) {
					for (var i = 0; i < table.Content.length; i++) {
						var row = table.Content[i];
						for (var j = 0; j < row.Content.length; j++) {
							res += " ";

							var cell = row.Content[j];
							getTextFromCell(cell);
						}

						res += '\r\n';
					}
				};

				var getTextFromDocumentContent = function (documentContent, isNAddNewLine) {
					for (var i = 0; i < documentContent.length; i++) {
						var item = documentContent[i];
						if (type_Paragraph === item.GetType()) {
							if (!isNAddNewLine) {
								res += '\r\n';
							}

							getTextFromParagraph(item);
						} else if (type_Table === item.GetType()) {
							res += '\r\n';

							getTextFromTable(item);
						}
					}
				};

				var getTextFromParagraph = function (paragraph) {
					for (var j = 0; j < paragraph.Content.length; j++) {
						if (para_Math === paragraph.Content[j].GetType() && paragraph.Content[j].Root) {
							var mathTextContent = paragraph.Content[j].Root.GetTextContent();
							if (mathTextContent) {
								res += mathTextContent.str;
							}
						} else {
							res += paragraph.Content[j].GetSelectedText(true);
						}
					}
				};

				getTextFromDocumentContent(data);

				return res;
			},

			_getTableFromText: function (text, bPastedArray) {

				var addTextIntoCell = function (row, col, sText) {
					var cell = aResult.getCell(rowCounter, colCounter);

					//check text
					AscFonts.FontPickerByCharacter.getFontsByString(sText);
					cell.content[0] = {text: sText, format: new AscCommonExcel.Font()};

					return cell;
				};

				var aResult = new excelPasteContent();
				var width = 0;
				var colCounter = 0;
				var rowCounter = 0;
				var sCurChar = "";
				var _parseText = function (sText, ignoreTab) {
					if (bPastedArray && sText === "") {
						if (colCounter > width) {
							width = colCounter;
						}
						addTextIntoCell(rowCounter, colCounter, sText);
						sCurChar = "";
					}
					for (var i = 0, length = sText.length; i < length; i++) {
						var Char = sText.charAt(i);
						var Code = sText.charCodeAt(i);

						if (colCounter > width) {
							width = colCounter;
						}

						if (13 === Code || '\n' === Char || sCurChar.length >= Asc.c_oAscMaxCellOrCommentLength) {
							if ("" === sCurChar) {
								addTextIntoCell(rowCounter, colCounter, sCurChar);
								colCounter = 0;
								rowCounter++;
							} else {
								addTextIntoCell(rowCounter, colCounter, sCurChar);
								colCounter = 0;

								rowCounter++;
								sCurChar = "";
							}
						} else {
							if (32 === Code || 160 === Code) {//160 - nbsp
								sCurChar += " ";
							} else if (9 === Code) {//tab
								if (!ignoreTab) {
									addTextIntoCell(rowCounter, colCounter, sCurChar);
									colCounter++;
									sCurChar = "";
								}
							} else {
								sCurChar += Char;
							}

							if (i === length - 1) {
								addTextIntoCell(rowCounter, colCounter, sCurChar);
								sCurChar = "";
							}
						}
					}
				};

				if (!bPastedArray) {
					text = text.replace(/\r\n/g, '\n');
					_parseText(text);
				} else {
					for (var i = 0; i < text.length; i++) {
						colCounter = 0;
						if (text[i]) {
							for (var j = 0; j < text[i].length; j++) {
								_parseText(text[i][j], true);
								colCounter++;
							}
						}
						rowCounter++;
					}
				}

				aResult.props.cellCount = width + 1;
				aResult.props.rowSpanSpCount = 0;

				return aResult;
			},

			_getFragmentsFromHtml: function (html) {
				//даная функция рассчитана только на вставку символа из плагина
				//TODO для вставки полноценной html нужно писать обработку

				var res = null;
				if (html && html.children) {
					for (var i = 0; i < html.children.length; i++) {
						if (html.children[i] && html.children[i].nodeName) {
							var sChildNodeName = html.children[i].nodeName.toLowerCase();
							if (sChildNodeName === "style" || sChildNodeName === "#comment" || sChildNodeName === "script") {
								continue;
							}
						}

						if (!res) {
							res = {fragments: [], fonts: {}};
						}

						var children = html.children[i];
						var computedStyle = this._getComputedStyle(children);
						var fragment = new AscCommonExcel.Fragment();
						var format = new AscCommonExcel.Font();

						if (computedStyle) {
							var bold = computedStyle.getPropertyValue("font-weight");
							if ("bold" === bold) {
								format.setBold(true);
							}
							/*var color = computedStyle.getPropertyValue("color");
							 if(color){
							 format.setColor(color);
							 }*/
							var italic = computedStyle.getPropertyValue("font-style");
							if ("italic" === italic) {
								format.setItalic(true);
							}
							var fontFamily = window["AscCommon"].CheckDefaultFontFamily(computedStyle.getPropertyValue("font-family"), window["Asc"]["editor"]);
							fontFamily = g_fontApplication.GetFontNameDictionary(fontFamily, true);
							if (fontFamily) {
								format.setName(fontFamily);
								res.fonts[fontFamily] = 1;
							}
							/*var fontSize = computedStyle.getPropertyValue("font-size");
							 if(fontSize){
							 format.setSize(fontSize);
							 }*/
							var text_decoration = computedStyle.getPropertyValue("text-decoration");
							if (text_decoration) {
								var underline, Strikeout;
								if (-1 !== text_decoration.indexOf("underline")) {
									underline = true;
								} else if (-1 !== text_decoration.indexOf("none")) {
									underline = false;
								}
								if (-1 !== text_decoration.indexOf("line-through")) {
									Strikeout = true;
								}
								if (underline) {
									format.setUnderline(true);
								}
								if (Strikeout) {
									format.setUnderline(true);
								}
							}
						}

						fragment.setFragmentText(children.innerText);
						AscFonts.FontPickerByCharacter.getFontsByString(fragment.getFragmentText());
						fragment.format = format;

						res.fragments.push(fragment);
					}
				}
				return res;
			},

			_getComputedStyle: function (node) {
				var computedStyle = null;
				if (null != node && Node.ELEMENT_NODE === node.nodeType) {
					var defaultView = node.ownerDocument.defaultView;
					computedStyle = defaultView.getComputedStyle(node, null);
				}
				return computedStyle;
			}
		};

		/** @constructor */
		function excelPasteContent() {
			this.content = [];
			this.props = {};

			return this;
		}

		excelPasteContent.prototype = {

			constructor: excelPasteContent,

			setCellContent: function (row, col, data) {
				if (!this.content[row]) {
					this.content[row] = [];
				}
				if (!this.content[row][col]) {
					this.content[row][col] = [];
				}

				this.content[row][col] = data;
			},

			getCell: function (row, col) {
				if (!this.content[row]) {
					this.content[row] = [];
				}
				if (!this.content[row][col]) {
					this.content[row][col] = new pasteCell();
				}

				return this.content[row][col];
			},

			deleteCell: function (row, col) {
				delete this.content[row][col];
			}
		};

		/** @constructor */
		function pasteCell() {
			this.content = [];

			this.rowSpan = null;
			this.colSpan = null;
			this.bc = null;
			this.borders = null;
			this.toolTip = null;
			this.hyperLink = null;
			this.location = null;

			this.indent = null;
			this.props = null;

			return this;
		}

		pasteCell.prototype = {

			constructor: pasteCell,

			addContentItem: function (item) {
				this.content.push(item);
			},

			clone: function () {
				var result = new pasteCell();

				for (var item = 0; item < this.content.length; item++) {
					result.content[item] = {text: this.content[item].text, format: this.content[item].format};
				}

				result.borders = this.borders;
				result.rowSpan = this.rowSpan;
				result.colSpan = this.colSpan;
				result.toolTip = this.toolTip;
				result.bc = this.bc;
				result.hyperLink = this.hyperLink;
				result.location = this.location;

				result.alignVertical = this.alignVertical;
				result.alignHorizontal = this.alignHorizontal;

				return result;
			}
		};


		/** @constructor */
		function pasteFromBinaryWord(clipboard, ws, bFromPresentation) {
			this.aResult = new excelPasteContent();

			this.fontsNew = {};
			this.clipboard = clipboard;
			this.ws = ws;
			this.isUsuallyPutImages = null;
			this.maxLengthRowCount = 0;
			//для этого сделал функцию _getParagraphMeasure у DocumentContentBoundsElement
			//this.rowDiff = 0;//для обработки данных в ране, разделенных shift+enter

			this.paragraphText = "";
			this.bFromPresentation = bFromPresentation;
			this.prevTextPr = null;

			//работает когда есть внутренние табы
			//TODO  в дальнейшем учитывать внутреннии табы в DocumentContentBounds
			this.maxCellCount = 0;

			this.footnotesCount = 0;

			return this;
		}

		pasteFromBinaryWord.prototype = {

			constructor: pasteFromBinaryWord,

			_paste: function (worksheet, pasteData) {
				var documentContent = pasteData.content;
				var t = this;

				//у родителя(CDocument) проставляю контент. нужно для вставки извне нумерованного списка. ф-ия CalculateNumberingValues требует наличие этих параграфов в родителе.
				var cDocument = documentContent && documentContent[0] &&
				documentContent[0].Parent instanceof CDocument ? documentContent[0].Parent : null;
				if (cDocument && cDocument.Content && 1 === cDocument.Content.length) {
					cDocument.Content = documentContent;
				}

				if (pasteData.images && pasteData.images.length) {
					this.isUsuallyPutImages = true;
				}

				if (!documentContent || (documentContent && !documentContent.length)) {
					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
					return;
				}

				var documentContentBounds = new DocumentContentBounds();
				var coverDocument = documentContentBounds.getBounds(0, 0, documentContent);
				AscFormat.ExecuteNoHistory(this._parseChildren, this, [coverDocument]);

				//не вставляем графику в редактор диаграмм
				//если кроме графики есть ещё данные, то убираем только графику
				if (window["Asc"]["editor"] && window["Asc"]["editor"].isChartEditor) {
					if (this.aResult.props && this.aResult.props.addImagesFromWord && this.aResult.props.addImagesFromWord.length === 1 && this.aResult.content) {
						if (1 === this.aResult.content.length && 1 === this.aResult.content[0].length && this.aResult.content[0][0].content && this.aResult.content[0][0].content.length === 0) {
							window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
							return;
						} else {
							this.aResult.props.addImagesFromWord = [];
							this.clipboard.alreadyLoadImagesOnServer = true;
						}
					}
				}

				var newFonts = this.fontsNew;
				if (pasteData.fonts && pasteData.fonts.length) {
					newFonts = {};
					for (var i = 0; i < pasteData.fonts.length; i++) {
						newFonts[pasteData.fonts[i].name] = 1;
					}
				}

				this.aResult.props.fontsNew = newFonts;
				this.aResult.props.rowSpanSpCount = 0;
				this.aResult.props.cellCount = this.maxCellCount + 1 > coverDocument.width ? this.maxCellCount + 1 : coverDocument.width;
				this.aResult.props._images = pasteData.images && pasteData.images.length ? pasteData.images : this.aResult.props._images;
				this.aResult.props._aPastedImages = pasteData.aPastedImages && pasteData.aPastedImages.length ? pasteData.aPastedImages : this.aResult.props._aPastedImages;


				//TODO alreadyLoadImagesOnServer - пересмотреть
				//alreadyLoadImagesOnServer - флаг используется для загрузки изображений из html
				//грузим картинки для вствки из документов(если это необходимо)
				//в данный момент в worksheetView не грузятся изображения
				var specialPasteProps = window['AscCommon'].g_specialPasteHelper.specialPasteProps;
				var aImagesToDownload = this.aResult.props._images;
				if (!this.clipboard.alreadyLoadImagesOnServer && aImagesToDownload && (!specialPasteProps || (specialPasteProps && specialPasteProps.images)))//load to server
				{
					var oObjectsForDownload = AscCommon.GetObjectsForImageDownload(t.aResult.props._aPastedImages);
					var api = window["Asc"]["editor"];
					var oImageMap = {};
					AscCommon.sendImgUrls(api, oObjectsForDownload.aUrls, function (data) {
						History.TurnOff();
						AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl,
							oImageMap);
						History.TurnOn();
						t.aResult.props.oImageMap = oImageMap;
						t.aResult.props.data = data;
						worksheet.setSelectionInfo('paste', {data: t.aResult});
					});
				} else {
					worksheet.setSelectionInfo('paste', {data: t.aResult});
				}
			},

			_parseChildren: function (children) {
				var childrens = children.children;
				var colSpan, rowSpan;
				for (var i = 0; i < childrens.length; i++) {
					if (childrens[i].type === c_oAscBoundsElementType.Cell) {
						for (var row = childrens[i].top; row < childrens[i].top + childrens[i].height; row++) {
							for (var col = childrens[i].left; col < childrens[i].left + childrens[i].width; col++) {
								var isCtable = false;
								var tempChildren = childrens[i].children[0].children;
								colSpan = null;
								rowSpan = null;
								for (var temp = 0; temp < tempChildren.length; temp++) {
									if (tempChildren[temp].type === c_oAscBoundsElementType.Table) {
										isCtable = true;
									}
								}
								if (childrens[i].width > 1 && isCtable && col === childrens[i].left) {
									colSpan = childrens[i].width;
									rowSpan = 1;
								} else if (!isCtable && tempChildren.length === 1) {
									rowSpan = childrens[i].height;
									colSpan = childrens[i].width;
								}


								var newCell = this.aResult.getCell(row /*+ this.rowDiff*/, col);
								newCell.rowSpan = rowSpan;
								newCell.colSpan = colSpan;

								//backgroundColor
								var backgroundColor = this.getBackgroundColorTCell(childrens[i]);
								if (backgroundColor) {
									newCell.bc = backgroundColor;
								}

								newCell.borders = this._getBorders(childrens[i], row, col, newCell.borders);

								var alignVertical = childrens[i].elem.Get_VAlign();
								switch (alignVertical) {
									case vertalignjc_Bottom:
										newCell.alignVertical = Asc.c_oAscVAlign.Bottom;
										break;
									case vertalignjc_Center:
										newCell.alignVertical = Asc.c_oAscVAlign.Center;
										break;
									case vertalignjc_Top:
										newCell.alignVertical = Asc.c_oAscVAlign.Top;
										break;
								}
							}
						}
					}

					if (childrens[i].children.length === 0) {
						//if parent - cell of table
						colSpan = null;
						rowSpan = null;

						this._parseParagraph(childrens[i], childrens[i].top /*+ this.rowDiff*/, childrens[i].left);
					} else {
						this._parseChildren(childrens[i]);
					}
				}
			},

			_parseParagraph: function (paragraph, row, col) {
				this.paragraphText = "";

				var t = this;
				var content = paragraph.elem.Content;
				var fontFamily = "Arial";
				var text = null;
				var oNewItem = new pasteCell();

				var aResult = this.aResult;
				if (row === undefined) {
					if (aResult.content.length === 0) {
						row = 0;
					} else {
						row = aResult.length;
					}
				}

				var cell = this.aResult.getCell(row + this.maxLengthRowCount, col);
				if (cell && cell.content && cell.content.length === 0 && (cell.borders || cell.rowSpan != null)) {
					if (cell.borders) {
						oNewItem.borders = cell.borders;
					}

					if (undefined !== cell.alignVertical) {
						oNewItem.alignVertical = cell.alignVertical;
					}

					if (cell.rowSpan != null) {
						oNewItem.rowSpan = cell.rowSpan;
						oNewItem.colSpan = cell.colSpan;
					}
					this.aResult.deleteCell(row + this.maxLengthRowCount, col)
				}

				var startRow = row;
				var innerCol = 0;
				if (col === undefined) {
					col = 0;
				}

				//backgroundColor
				var backgroundColor = this.getBackgroundColorTCell(paragraph);
				if (backgroundColor) {
					oNewItem.bc = backgroundColor;
				}

				//настройки параграфа
				paragraph.elem.CompiledPr.NeedRecalc = true;
				var paraPr = paragraph.elem.Get_CompiledPr();

				var firstLine = paraPr && paraPr.ParaPr && paraPr.ParaPr.Ind && paraPr.ParaPr.Ind.Left;
				if (firstLine) {
					oNewItem.indent = parseInt(firstLine / AscCommon.koef_mm_to_indent);
					if (oNewItem.indent < 0) {
						oNewItem.indent = null;
					}
				}

				//горизонтальное выравнивание
				var horisontalAlign = this._getAlignHorisontal(paraPr);
				if (horisontalAlign == null) {
					oNewItem.wrap = true;
				}

				oNewItem.alignHorizontal = horisontalAlign;

				//вертикальное выравнивание
				//oNewItem.alignVertical = Asc.c_oAscVAlign.Center;

				//так же wrap выставляем у параграфа, чьим родителем является ячейка таблицы
				var cellParent = this._getParentByTag(paragraph, c_oAscBoundsElementType.Cell);
				if (cellParent) {
					oNewItem.wrap = !(cellParent.elem && cellParent.elem.Pr && true === cellParent.elem.Pr.NoWrap);
				}

				//Numbering
				var LvlPr = null;
				var Lvl = null;
				var oNumPr = paragraph.elem.GetNumPr ? paragraph.elem.GetNumPr() : null;
				var numberingText = null;
				var formatText;
				if (oNumPr != null) {
					var oNum = paragraph.elem.Parent.GetNumbering().GetNum(oNumPr.NumId);
					if (oNum) {
						LvlPr = oNum.GetLvl(oNumPr.Lvl);
						Lvl = oNumPr.Lvl;
					}

					numberingText = this._parseNumbering(paragraph.elem);

					if (text == null) {
						text = "";
					}

					text += this._getAllNumberingText(Lvl, numberingText);

					formatText = this._getPrParaRun(paraPr, LvlPr.GetTextPr());
					fontFamily = formatText.format.getName();
					this.fontsNew[fontFamily] = 1;

					oNewItem.content.push(formatText);

					if (text !== null) {
						oNewItem.content[oNewItem.content.length - 1].text = text;
						this.paragraphText += text;
					}
					text = "";
				}


				//проходимся по контенту paragraph
				var paraRunObj;
				//получае общий шрифт для ячейки для случая когда вставляем нумерованный список
				//общего может не быть в том случае, если шрифты внутри ячейки разные
				var allParaFont, textPr;
				for (var n = 0; n < content.length; n++) {
					this.aResult.getCell(row + this.maxLengthRowCount, innerCol + col);

					if (text == null) {
						text = "";
					}

					switch (content[n].Type) {
						case para_Run://*paraRun*
						{
							paraRunObj = this._parseParaRun(content[n], oNewItem, paraPr, innerCol, row, col, text);

							if (null !== allParaFont) {
								textPr = content[n].Get_CompiledPr();
								if (textPr && textPr.FontFamily && textPr.FontFamily.Name) {
									if (undefined === allParaFont) {
										allParaFont = textPr.FontFamily.Name;
									}
								} else if (textPr.FontFamily.Name !== allParaFont) {
									allParaFont = null;
								}
							}

							innerCol = paraRunObj.col;
							row = paraRunObj.row;
							oNewItem = paraRunObj.oNewItem;
							break;
						}
						case para_Hyperlink://*hyperLink*
						{
							//если несколько ссылок в одном параграфе, то отменяем ссылки
							if (!oNewItem.doNotApplyHyperlink) {
								if (!oNewItem.hyperLink) {
									oNewItem.hyperLink = content[n].Value;
									oNewItem.toolTip = content[n].ToolTip;
									oNewItem.location = content[n].Anchor;
								} else {
									oNewItem.hyperLink = null;
									oNewItem.toolTip = null;
									oNewItem.doNotApplyHyperlink = true;
								}
							}

							var lastTab;
							for (var h = 0; h < content[n].Content.length; h++) {
								switch (content[n].Content[h].Type) {
									case para_Run://*paraRun*
									{
										paraRunObj = this._parseParaRun(content[n].Content[h], oNewItem, paraPr, innerCol, row, col, text);
										oNewItem = paraRunObj.oNewItem;
										innerCol = paraRunObj.col;
										row = paraRunObj.row;
										if (lastTab) {
											oNewItem.hyperLink = content[n].Value;
											oNewItem.toolTip = content[n].ToolTip;
											oNewItem.location = content[n].Anchor;
										}

										lastTab = paraRunObj.lastTab;

										break;
									}
								}
							}
							allParaFont = null;
							break;
						}
						case para_Math://*para_Math*
						{

							if (this.bFromPresentation) {
								var mathTextContent = content[n].Root.GetTextContent();
								if (mathTextContent) {
									var newParaRun = new ParaRun();
									window['AscCommon'].addTextIntoRun(newParaRun, mathTextContent.str);
									var objTemp = t._parseParaRun(newParaRun, oNewItem, paraPr, innerCol, row, col, text);
									innerCol = objTemp.col;
									row = objTemp.row;
								}
							} else {
								var tempFonts = [];
								content[n].Get_AllFontNames(tempFonts);

								for (var i in tempFonts) {
									this.fontsNew[i] = 1;
								}

								if (!aResult.props.addImagesFromWord) {
									aResult.props.addImagesFromWord = [];
								}
								aResult.props.addImagesFromWord.push(
									{image: content[n], col: innerCol + col, row: row});

								if (null === this.isUsuallyPutImages) {
									this._addImageToMap(content[n]);
								}
							}
							allParaFont = null;
							break;
						}
						case para_InlineLevelSdt://*InlineLevelSdt*
						{
							//TODO allParaFont ? 
							for (h = 0; h < content[n].Content.length; h++) {
								switch (content[n].Content[h].Type) {
									case para_Run://*paraRun*
									{
										paraRunObj = this._parseParaRun(content[n].Content[h], oNewItem, paraPr, innerCol, row, col, text);

										oNewItem = paraRunObj.oNewItem;
										innerCol = paraRunObj.col;
										row = paraRunObj.row;
										break;
									}
								}
							}
							break;
						}
					}
				}

				if (null !== numberingText && allParaFont) {
					oNewItem.props = {fontName: allParaFont};
				}

				oNewItem.textVal = this.paragraphText;
				//this.rowDiff += row - startRow;
			},

			_parseParaRun: function (paraRun, oNewItem, paraPr, innerCol, row, col, text, prevTextPr) {
				var t = this;
				var paraRunContent = paraRun.Content;
				var aResult = this.aResult;
				var paragraphFontFamily = paraPr.TextPr.FontFamily.Name;
				var cloneNewItem, formatText;
				var newLine = "\n";

				var cTextPr = prevTextPr ? prevTextPr : paraRun.Get_CompiledPr();
				if (cTextPr && !(paraRunContent.length === 1 && paraRunContent[0] instanceof AscWord.CRunParagraphMark))//settings for text
				{
					formatText = this._getPrParaRun(paraPr, cTextPr);
				} else if (!formatText) {
					formatText = this._getPrParaRun(paraPr, cTextPr);
				}
				this.prevTextPr = cTextPr;

				var pushData = function () {
					t.fontsNew[paragraphFontFamily] = 1;
					oNewItem.content.push(formatText);

					if (text !== null) {
						oNewItem.content[oNewItem.content.length - 1].text = text;
					}

					cloneNewItem = oNewItem.clone();
					cloneNewItem.rowSpan = null;
					cloneNewItem.colSpan = null;

					//переходим в следующую ячейку
					cell = aResult.getCell(row + t.maxLengthRowCount, innerCol + col);
					aResult.setCellContent(row + t.maxLengthRowCount, innerCol + col, cloneNewItem);

					text = "";
					oNewItem = new pasteCell();
				};

				//проходимся по контенту paraRun
				var cell;
				var lastTab;
				for (var pR = 0; pR < paraRunContent.length; pR++) {
					lastTab = false;
					switch (paraRunContent[pR].Type) {
						case para_Math_BreakOperator:
						case para_Math_Text: {
							text += AscCommon.encodeSurrogateChar(paraRunContent[pR].value);
							break;
						}
						case para_Text: {
							text += AscCommon.encodeSurrogateChar(paraRunContent[pR].Value);
							break;
						}
						case para_NewLine: {
							//on feature: read ms style comments and use only sameCell flag
							if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Excel ||
								(paraRunContent[pR].Flags && paraRunContent[pR].Flags.sameCell)) {
								text += newLine;
							} else {
								pushData();
								row++;
								innerCol = 0;
							}

							break;
						}
						case para_Space: {
							text += " ";
							break;
						}
						case para_Tab: {
							pushData();
							lastTab = true;
							innerCol++;

							if (innerCol > this.maxCellCount) {
								this.maxCellCount = innerCol;
							}

							break;
						}
						case para_Drawing: {
							if (!aResult.props.addImagesFromWord) {
								aResult.props.addImagesFromWord = [];
							}

							var drawing = this._checkDrawingList(paraRunContent[pR]);
							aResult.props.addImagesFromWord.push({image: drawing, col: innerCol + col, row: row});

							if (null === this.isUsuallyPutImages) {
								this._addImageToMap(drawing);
							}

							break;
						}
						case para_FootnoteReference: {
							if (1 === paraRunContent.length) {
								var footnotesNumber = this.footnotesCount + 1;
								text += "[" + footnotesNumber + "]";
								this.footnotesCount++;
							}
							break;
						}
						case para_End: {
							cell = aResult.getCell(row, innerCol + col);
							aResult.setCellContent(row, innerCol + col, oNewItem);

							var checkMaxTextLength = this.clipboard._checkMaxTextLength(this.aResult,
								row + this.maxLengthRowCount, innerCol + col);
							if (checkMaxTextLength) {
								aResult.content = checkMaxTextLength.aResult.content;
								this.maxLengthRowCount += checkMaxTextLength.r - row;
							}

							break;
						}
					}
				}

				if (text !== "") {
					this.fontsNew[paragraphFontFamily] = 1;
					oNewItem.content.push(formatText);
					if (text !== null) {
						oNewItem.content[oNewItem.content.length - 1].text = text;
						this.paragraphText += text;
					}
				}

				return {col: innerCol, row: row, prevTextPr: cTextPr, oNewItem: oNewItem, lastTab: lastTab};
			},

			_addImageToMap: function (paraDrawing) {
				var aResult = this.aResult;
				if (!aResult.props._aPastedImages) {
					aResult.props._aPastedImages = [];
				}
				if (!aResult.props._images) {
					aResult.props._images = [];
				}

				var oGraphicObj = paraDrawing.GraphicObj;
				if (!oGraphicObj || (oGraphicObj && !oGraphicObj.blipFill) ||
					(oGraphicObj && oGraphicObj.blipFill && !oGraphicObj.blipFill.RasterImageId)) {
					return;
				}

				var sImageUrl = oGraphicObj.blipFill.RasterImageId;
				aResult.props._aPastedImages.push(new AscCommon.CBuilderImages(oGraphicObj.blipFill, sImageUrl, oGraphicObj, oGraphicObj.spPr, null));
				aResult.props._images.push(sImageUrl);
			},


			_checkDrawingList: function (drawing) {
				var t = this;
				var checkParagraph = function (paragraph) {
					//TODO не могу получить в данном случае GetNumPr
					/*var LvlPr = null;
					 var Lvl = null;
					 var oNumPr = paragraph.GetNumPr ? paragraph.GetNumPr() : null;
					 var numberingText = null;
					 var formatText;
					 if (oNumPr != null) {
					 var oNum = paragraph.Parent.GetNumbering().GetNum(oNumPr.NumId);
					 if (oNum) {
					 LvlPr = oNum.GetLvl(oNumPr.Lvl);
					 Lvl = oNumPr.Lvl;
					 }

					 numberingText = t._parseNumbering(paragraph);


					 var text = t._getAllNumberingText(Lvl, numberingText);
					 formatText = this._getPrParaRun(paraPr, LvlPr.GetTextPr());
					 }*/
				};

				if (drawing.GraphicObj && drawing.GraphicObj.textBoxContent &&
					drawing.GraphicObj.textBoxContent.Content) {
					var content = drawing.GraphicObj.textBoxContent.Content;
					for (var i = 0; i < content.length; i++) {
						checkParagraph(content[i]);
					}
				}

				return drawing;
			},

			_getBorders: function (cellTable, top, left, oldBorders) {
				var borders = cellTable.elem.Get_Borders();
				var widthCell = cellTable.width;
				var heightCell = cellTable.height;
				var defaultStyle = "solid";
				var borderStyleName;
				var t = this;

				var getBorderColor = function (curBorder) {
					var color = null;
					var backgroundColor = null;

					var unifill = curBorder.Unifill && curBorder.Unifill.fill && curBorder.Unifill.fill.color ? curBorder.Unifill.fill.color : null;
					if (unifill && unifill.color && unifill.color.type !== Asc.c_oAscColor.COLOR_TYPE_SCHEME && unifill.color.RGBA) {
						color = unifill.color.RGBA;
						backgroundColor = new AscCommonExcel.RgbColor(t.clipboard._getBinaryColor("rgb(" + color.R + "," + color.G + "," + color.B + ")"));
					} else {
						color = curBorder.Color;
						backgroundColor = new AscCommonExcel.RgbColor(t.clipboard._getBinaryColor("rgb(" + color.r + "," + color.g + "," + color.b + ")"));
					}

					return backgroundColor;
				};

				var formatBorders;
				if (oldBorders) {
					formatBorders = oldBorders;
				} else {
					formatBorders = new AscCommonExcel.Border();
					formatBorders.initDefault();
				}
				//top border for cell
				if (top === cellTable.top && (!formatBorders.t || formatBorders.t.isEmpty()) && borders.Top.Value !== AscWord.BorderType.none) {
					borderStyleName = this.clipboard._getBorderStyleName(defaultStyle, this.ws.objectRender.convertMetric(borders.Top.Size, 3, 1));
					if (null !== borderStyleName) {
						formatBorders.t = new AscCommonExcel.BorderProp();
						formatBorders.t.setStyle(borderStyleName);
						formatBorders.t.c = getBorderColor(borders.Top);
					}
				}
				//left border for cell
				if (left === cellTable.left && (!formatBorders.l || formatBorders.l.isEmpty()) && borders.Left.Value !== AscWord.BorderType.none) {
					borderStyleName = this.clipboard._getBorderStyleName(defaultStyle, this.ws.objectRender.convertMetric(borders.Left.Size, 3, 1));
					if (null !== borderStyleName) {
						formatBorders.l = new AscCommonExcel.BorderProp();
						formatBorders.l.setStyle(borderStyleName);
						formatBorders.l.c = getBorderColor(borders.Left);
					}
				}
				//bottom border for cell
				if (top === cellTable.top + heightCell - 1 && (!formatBorders.b || formatBorders.b.isEmpty()) &&
					borders.Bottom.Value !== AscWord.BorderType.none) {
					borderStyleName = this.clipboard._getBorderStyleName(defaultStyle, this.ws.objectRender.convertMetric(borders.Bottom.Size, 3, 1));
					if (null !== borderStyleName) {
						formatBorders.b = new AscCommonExcel.BorderProp();
						formatBorders.b.setStyle(borderStyleName);
						formatBorders.b.c = getBorderColor(borders.Bottom);
					}
				}
				//right border for cell
				if (left === cellTable.left + widthCell - 1 && (!formatBorders.r || formatBorders.r.isEmpty()) &&
					borders.Right.Value !== AscWord.BorderType.none) {
					borderStyleName = this.clipboard._getBorderStyleName(defaultStyle, this.ws.objectRender.convertMetric(borders.Right.Size, 3, 1));
					if (null !== borderStyleName) {
						formatBorders.r = new AscCommonExcel.BorderProp();
						formatBorders.r.setStyle(borderStyleName);
						formatBorders.r.c = getBorderColor(borders.Right);
					}
				}

				return formatBorders;
			},

			_getAlignHorisontal: function (paraPr) {
				var result;
				var settings = paraPr.ParaPr;

				if (!settings) {
					return;
				}

				switch (settings.Jc) {
					case 0: {
						result = AscCommon.align_Right;
						break;
					}
					case 1: {
						result = AscCommon.align_Left;
						break;
					}
					case 2: {
						result = AscCommon.align_Center;
						break;
					}
					case 3: {
						result = null;
						break;
					}
				}

				return result;
			},

			getBackgroundColorTCell: function (elem) {
				var compiledPrCell, color;
				var backgroundColor = null;

				//TODO внутреннии таблицы без стиля - цвет фона белый
				var tableCell = this._getParentByTag(elem, c_oAscBoundsElementType.Cell);

				if (tableCell) {
					compiledPrCell = tableCell.elem.Get_CompiledPr();

					if (compiledPrCell && compiledPrCell.Shd.Value !== 1/*shd_Nil*/) {
						var shd = compiledPrCell.Shd;
						if (shd.Unifill && shd.Unifill.fill && shd.Unifill.fill.color && shd.Unifill.fill.color.color && shd.Unifill.fill.color.color.RGBA) {
							color = shd.Unifill.fill.color.color.RGBA;
							backgroundColor = new AscCommonExcel.RgbColor(this.clipboard._getBinaryColor("rgb(" + color.R + "," + color.G + "," + color.B + ")"));
						} else {
							color = shd.Color;
							backgroundColor = new AscCommonExcel.RgbColor(this.clipboard._getBinaryColor("rgb(" + color.r + "," + color.g + "," + color.b + ")"));
						}
					}
				}

				return backgroundColor;
			},

			_getParentByTag: function (elem, tag) {
				var result;
				if (!elem) {
					return null;
				}

				if (elem.type === tag) {
					result = elem;
				} else if (elem.parent) {
					result = this._getParentByTag(elem.parent, tag);
				} else if (!elem.parent) {
					result = null;
				}

				return result;
			},

			_getAllNumberingText: function (Lvl, numberingText) {
				var preSpace, result;
				if (Lvl === 0) {
					preSpace = "     ";
				} else if (Lvl === 1) {
					preSpace = "          ";
				} else if (Lvl >= 2) {
					preSpace = "               ";
				}

				var betweenSpace = "  ";
				result = preSpace + numberingText + betweenSpace;

				return result;
			},

			_parseNumbering: function (paragraph) {
				var Pr = paragraph.Get_CompiledPr();

				if (paragraph.Numbering) {
					var NumberingItem = paragraph.Numbering;
					if (para_Numbering === NumberingItem.Type) {
						var NumPr = Pr.ParaPr.NumPr;
						if (undefined === NumPr || undefined === NumPr.NumId || 0 === NumPr.NumId ||
							"0" === NumPr.NumId) {
							// Ничего не делаем
						} else {
							var Numbering = paragraph.Parent.GetNumbering();
							var NumLvl = Numbering.GetNum(NumPr.NumId).GetLvl(NumPr.Lvl);
							//var NumSuff   = NumLvl.GetSuff();
							//var NumJc     = NumLvl.GetJc();
							var NumTextPr = paragraph.Get_CompiledPr2(false).TextPr.Copy();

							// Word не рисует подчеркивание у символа списка, если оно пришло из настроек для
							// символа параграфа.

							var TextPr_temp = paragraph.TextPr.Value.Copy();
							TextPr_temp.Underline = undefined;

							NumTextPr.Merge(TextPr_temp);
							NumTextPr.Merge(NumLvl.GetTextPr());

							var text = null;
							var oNumPr = paragraph.GetNumPr();
							var Lvl;
							if (oNumPr != null) {
								var oNum = paragraph.Parent.GetNumbering().GetNum(oNumPr.NumId);
								if (null != oNum) {
									Lvl = oNumPr.Lvl;
									text = oNum.GetText(Lvl, paragraph.Parent.CalculateNumberingValues(paragraph, NumPr));
								}
							}

							return text;
						}
					}
				}
			},

			_getPrParaRun: function (paraPr, cTextPr) {
				var formatText, fontFamily, colorText;

				var paragraphFontSize = paraPr.TextPr.FontSize;
				var paragraphFontFamily = paraPr.TextPr && paraPr.TextPr.FontFamily ? paraPr.TextPr.FontFamily.Name : "Arial";
				var paragraphBold = paraPr.TextPr.Bold;
				var paragraphItalic = paraPr.TextPr.Italic;
				var paragraphStrikeout = paraPr.TextPr.Strikeout;
				var paragraphUnderline = paraPr.TextPr.Underline ? Asc.EUnderline.underlineSingle :
					Asc.EUnderline.underlineNone;
				var paragraphVertAlign = null;
				if (paraPr.TextPr.VertAlign === 1) {
					paragraphVertAlign = AscCommon.vertalign_SuperScript;
				} else if (paraPr.TextPr.VertAlign === 2) {
					paragraphVertAlign = AscCommon.vertalign_SubScript;
				}

				var colorParagraph = new AscCommonExcel.RgbColor(this.clipboard._getBinaryColor("rgb(" + paraPr.TextPr.Color.r + "," + paraPr.TextPr.Color.g + "," + paraPr.TextPr.Color.b + ")"));

				if (cTextPr.Color) {
					colorText = new AscCommonExcel.RgbColor(this.clipboard._getBinaryColor("rgb(" + cTextPr.Color.r + "," + cTextPr.Color.g + "," + cTextPr.Color.b + ")"));
				} else {
					colorText = null;
				}

				fontFamily = cTextPr.FontFamily ? cTextPr.FontFamily.Name : cTextPr.RFonts.CS ? cTextPr.RFonts.CS.Name : paragraphFontFamily;
				this.fontsNew[fontFamily] = 1;

				var verticalAlign;
				if (cTextPr.VertAlign === 2) {
					verticalAlign = AscCommon.vertalign_SubScript;
				} else if (cTextPr.VertAlign === 1) {
					verticalAlign = AscCommon.vertalign_SuperScript;
				}

				var font = new AscCommonExcel.Font();
				font.assignFromObject({
					fn: fontFamily,
					fs: cTextPr.FontSize ? cTextPr.FontSize : paragraphFontSize,
					c: colorText ? colorText : colorParagraph,
					b: cTextPr.Bold ? cTextPr.Bold : paragraphBold,
					i: cTextPr.Italic ? cTextPr.Italic : paragraphItalic,
					u: cTextPr.Underline ? Asc.EUnderline.underlineSingle : paragraphUnderline,
					s: cTextPr.Strikeout ? cTextPr.Strikeout :
						cTextPr.DStrikeout ? cTextPr.DStrikeout : paragraphStrikeout,
					va: verticalAlign ? verticalAlign : paragraphVertAlign
				});
				formatText = {
					format: font
				};

				return formatText;
			}
		};

		var c_oAscBoundsElementType = {
			Content: 0,
			Paragraph: 1,
			Table: 2,
			Row: 3,
			Cell: 4
		};

		function DocumentContentBoundsElement(elem, type, parent) {
			this.elem = elem;
			this.type = type;
			this.parent = parent;
			this.children = [];

			this.left = 0;
			this.top = 0;
			this.width = 0;
			this.height = 0;
		}

		function DocumentContentBounds() {
		}

		DocumentContentBounds.prototype = {

			constructor: DocumentContentBounds,

			getBounds: function (nLeft, nTop, aDocumentContent) {
				//в первный проход заполняем размеры
				//и могут заноситься относительные сдвиги при небходимости
				var oRes = this._getMeasure(aDocumentContent, null);
				//заполняем абсолютные сдвиги
				this._getOffset(nLeft, nTop, oRes);
				return oRes;
			},
			_getOffset: function (nLeft, nTop, elem) {
				elem.left += nLeft;
				elem.top += nTop;
				var nCurLeft = elem.left;
				var nCurTop = elem.top;
				var bIsRow = elem.elem instanceof CTableRow;
				for (var i = 0, length = elem.children.length; i < length; i++) {
					var child = elem.children[i];
					this._getOffset(nCurLeft, nCurTop, child);
					if (bIsRow)
						nCurLeft += child.width;
					else
						nCurTop += child.height;
				}
			},
			_getMeasure: function (aDocumentContent, oParent) {
				var oRes = new DocumentContentBoundsElement(aDocumentContent, c_oAscBoundsElementType.Content, oParent);
				for (var i = 0, length = aDocumentContent.length; i < length; i++) {
					var elem = aDocumentContent[i];
					var oNewElem = null;
					if (type_Paragraph === elem.GetType()) {
						oNewElem = this._getParagraphMeasure(elem, oRes);
					} else if (type_Table === elem.GetType()) {
						elem.ReIndexing(0);
						oNewElem = this._getTableMeasure(elem, oRes);
					} else if (type_BlockLevelSdt === elem.GetType()) {
						oNewElem = this._getMeasure(elem.Content.Content, oRes);
					}

					if (null != oNewElem) {
						oRes.children.push(oNewElem);
						if (oNewElem.width && oNewElem.width > oRes.width)
							oRes.width = oNewElem.width;
						oRes.height += oNewElem.height;
					}
				}
				return oRes;
			},
			_getParagraphMeasure: function (elem, oParent) {
				var oNewElem = new DocumentContentBoundsElement(elem, c_oAscBoundsElementType.Paragraph, oParent);
				oNewElem.width = 1;
				oNewElem.height = 1;

				for (var i = 0, length = elem.Content.length; i < length; i++) {
					if (elem.Content[i] && elem.Content[i].Content) {
						for (var j = 0; j < elem.Content[i].Content.length; j++) {
							let innerElem = elem.Content[i].Content[j];
							if (innerElem && para_NewLine === innerElem.GetType()) {
								//on feature: read ms style comments and use only sameCell flag 
								if (AscCommon.g_clipboardBase.pastedFrom !== AscCommon.c_oClipboardPastedFrom.Excel && !(innerElem.Flags && innerElem.Flags.sameCell === true)) {
									oNewElem.height++;
								}
							}
						}
					}
				}

				return oNewElem;
			},
			_getTableMeasure: function (table, oParent) {
				var oRes = new DocumentContentBoundsElement(table, c_oAscBoundsElementType.Table, oParent);
				//надо рассчитать сколько ячеек приходится на tableGrid, по умолчанию по одной
				//todo надо оптимизировать размер 
				var aGridWidth = [], i, length;
				for (i = 0, length = table.TableGrid.length; i < length; i++)
					aGridWidth.push(1);
				//заполняем aGridWidth
				for (i = 0, length = table.Content.length; i < length; i++) {
					var row = table.Content[i];
					var oNewElem = this._setRowGridWidth(row, oRes, aGridWidth);
					if (null != oNewElem)
						oRes.children.push(oNewElem);
				}
				var aSumGridWidth = [];
				var nTempSum = 0;
				for (i = 0, length = aGridWidth.length; i < length + 1; i++) {
					aSumGridWidth[i] = nTempSum;
					var nCurValue = aGridWidth[i];
					if (nCurValue)
						nTempSum += nCurValue;
				}
				//заполняем размеры
				for (i = 0, length = oRes.children.length; i < length; i++) {
					var rowWrapped = oRes.children[i];
					this._getRowMeasure(rowWrapped, aSumGridWidth, oRes.children, i);
					oRes.height += rowWrapped.height;
					//в left временно занесен относительный сдвиг
					if (rowWrapped.width + rowWrapped.left > oRes.width)
						oRes.width = rowWrapped.width + rowWrapped.left;
				}
				return oRes;
			},
			_setRowGridWidth: function (row, oParent, aGridWidth) {
				var oRes = new DocumentContentBoundsElement(row, c_oAscBoundsElementType.Row, oParent);
				var nSumGrid = 0;
				var BeforeInfo = row.Get_Before();
				if (BeforeInfo && BeforeInfo.GridBefore)
					nSumGrid += BeforeInfo.GridBefore;
				for (var i = 0, length = row.Content.length; i < length; i++) {
					var cell = row.Content[i];
					var oNewCell = new DocumentContentBoundsElement(cell, c_oAscBoundsElementType.Cell, oRes);
					oRes.children.push(oNewCell);
					var oNewElem = this._getMeasure(cell.Content.Content, oNewCell);
					oNewCell.children.push(oNewElem);
					oNewCell.width = oNewElem.width;
					oNewCell.height = oNewElem.height;
					if (oNewCell.height > oRes.height)
						oRes.height = oNewCell.height;
					var nCellGrid = cell.Get_GridSpan();
					if (oNewElem.width > nCellGrid) {
						var nFirstGridWidth = oNewElem.width - nCellGrid + 1;
						var nCurValue = aGridWidth[nSumGrid];
						if (null != nCurValue && nCurValue < nFirstGridWidth)
							aGridWidth[nSumGrid] = nFirstGridWidth;
					}
					nSumGrid += nCellGrid;
				}
				return oRes;
			},
			_getRowMeasure: function (rowWrapped, aSumGridWidth, rows, index) {

				var getNotMergedPreviousCell = function (i) {
					var res = null;
					if (rows && index > 0) {
						for (var j = index; j >= 0; j--) {
							if (rows[j] && rows[j].children && rows[j].children[i]) {
								var vMerge = rows[j].children[i].elem.GetVMerge();
								if (1 === vMerge) {
									res = rows[j].children[i];
									break;
								}
							}
						}
					}
					return res;
				};

				var nSumGrid = 0;
				var BeforeInfo = rowWrapped.elem.Get_Before();
				if (BeforeInfo && BeforeInfo.GridBefore) {
					//временно заносим относительный сдвиг
					rowWrapped.left = aSumGridWidth[nSumGrid + BeforeInfo.GridBefore] - aSumGridWidth[nSumGrid];
					nSumGrid += BeforeInfo.GridBefore;
				}

				for (var i = 0, length = rowWrapped.children.length; i < length; i++) {
					var cellWrapped = rowWrapped.children[i];
					var nCellGrid = cellWrapped.elem.Get_GridSpan();
					cellWrapped.width = aSumGridWidth[nSumGrid + nCellGrid] - aSumGridWidth[nSumGrid];
					//выравниваем высоту ячеек по-максимому
					cellWrapped.height = rowWrapped.height;
					rowWrapped.width += cellWrapped.width;
					nSumGrid += nCellGrid;

					//vertical merge
					var previousCell = getNotMergedPreviousCell(i);
					if (previousCell) {
						var vMerge = cellWrapped.elem.GetVMerge();
						if (vMerge > 1) {
							previousCell.height += vMerge - 1;
						}
					}
				}
			}
		};


		//---------------------------------------------------------export---------------------------------------------------
		var g_clipboardExcel = new Clipboard();
		window['AscCommonExcel'] = window['AscCommonExcel'] || {};
		window['AscCommonExcel'].g_clipboardExcel = g_clipboardExcel;

		window["Asc"]["SpecialPasteProps"] = window["Asc"].SpecialPasteProps = CSpecialPasteProps;
		prot = CSpecialPasteProps.prototype;
		prot["asc_setProps"] = prot.asc_setProps;
		prot["asc_getProps"] = prot.asc_getProps;
		prot["asc_setAdvancedOptions"] = prot.asc_setAdvancedOptions;
		prot["asc_setOperation"] = prot.asc_setOperation;
		prot["asc_getOperation"] = prot.asc_getOperation;
		prot["asc_setTranspose"] = prot.asc_setTranspose;
		prot["asc_setSkipBlanks"] = prot.asc_setSkipBlanks;
		prot["asc_getTableAllowed"] = prot.asc_getTableAllowed;

	}
)(jQuery, window);
