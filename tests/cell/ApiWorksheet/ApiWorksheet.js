QUnit.config.autostart = false;
$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._canResize = function () {
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getPPIY = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getPPIX = function () {
	};

	AscCommonExcel.asc_CEventsController.prototype.init = function () {
	};

	AscCommon.InitBrowserInputContext = function () {

	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
		this.ImageLoader = AscCommon.g_image_loader;
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
			setTimeout(startTests.bind(this), 0)
		}
	};

	window["Asc"]["editor"] = api;
	var wb, ws, wsData, wsView;

	function openDocument() {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		AscFormat.initStyleManager();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
		api.initCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		api.wb._init();
		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function () {
			return null;
		});
		ws = api.wbModel.aWorksheets[0];
		api.asc_insertWorksheet(["Data"]);
		wsData = wb.getWorksheetByName(["Data"], 0);

		wsView = api.wb.getWorksheet(0);
		wsView.handlers = api.handlers;
		wsView.objectRender = new AscFormat.DrawingObjects();
		wsView.objectRender.init(wsView);
		wsView.objectRender.drawingDocument.TargetHtmlElement = document.getElementById('editor_sdk');
		//wsView.objectRender.controller = new AscFormat.DrawingObjectsController(wsView.objectRender);
	}

	QUnit.module("ChartsDraw");

	function startTests()
    {
        QUnit.start();
        QUnit.test("GetSelectedShapes", function (assert) {
            debugger

			let wss = api.GetActiveSheet();

            for(let nShape = 0; nShape < 3; nShape++)
            {
                let shape = wss.AddShape("ellipse", 50 * 36000, 50 * 36000, api.CreateNoFill(), api.CreateStroke(0, api.CreateNoFill()), 0, 0, 0, 0);
                if (nShape !== 1)
                    shape.Select();
            }

            let selectedShapes = ws.GetSelectedShapes();
            console.log(selectedShapes);

            assert.strictEqual(
                true,
                0,
                "All CFs cleared by initializeTest"
            );
        });
    }
});
