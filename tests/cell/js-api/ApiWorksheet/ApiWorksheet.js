$(function () {

	// ====== REQUIRED ENVIRONMENT SETUP (preserve these stubs/settings) ======
    Asc.spreadsheet_api.prototype._init = function () { };
    Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
        callback();
    };
    AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () { };
    AscCommonExcel.WorkbookView.prototype._init = function () { };
    AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged =
        function () { };
    AscCommonExcel.WorkbookView.prototype.showWorksheet = function () { };
    AscCommonExcel.WorksheetView.prototype._init = function () { };
    AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable =
        function () { };
    AscCommonExcel.WorksheetView.prototype.setSelection = function () { };
    AscCommonExcel.WorksheetView.prototype.draw = function () { };
    AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects =
        function () { };
    AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () { };
    AscCommonExcel.WorksheetView.prototype.getZoom = function () { };
    AscCommonExcel.WorksheetView.prototype._getPPIY = function () { };
    AscCommonExcel.WorksheetView.prototype._getPPIX = function () { };
    AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () { };
    Asc.ReadDefTableStyles = function () { };

    var api = new Asc.spreadsheet_api({ "id-view": "editor_sdk" });

    api.FontLoader = { LoadDocumentFonts: function () { } };
    window["Asc"]["editor"] = api;
    AscCommon.g_oTableId.init();
    api._onEndLoadSdk();
    api.isOpenOOXInBrowser = false;
    api.OpenDocumentFromBin(null, AscCommon.getEmpty());
    api.initCollaborativeEditing({});
    api.wb = new AscCommonExcel.WorkbookView(
        api.wbModel,
        api.controller,
        api.handlers,
        api.HtmlElement,
        api.topLineEditorElement,
        api,
        api.collaborativeEditing,
        api.fontRenderingMode
    );

    var wsView = api.wb.getWorksheet(0);
    wsView.handlers = api.handlers;
    wsView.objectRender = new AscFormat.DrawingObjects();
    wsView.objectRender.controller = new AscFormat.DrawingObjectsController(
        wsView.objectRender
    );
    var ws = api.GetActiveSheet();

    // ====== TEST UTILITIES ======

    // MUST-HAVE helper: clear all conditional formats in A1:Z100 before each test
    window.initializeTest = function () {
        var r = ws.GetRange("A1:Z100");
        ws.worksheet.AutoFilter = null;
        // r.Clear();
    };

    theRange = function (address) {
        return ws.GetRange(address);
    };


	QUnit.module("ChartsDraw");
	QUnit.test("GetSelectedShapes", function (assert) {
		debugger


		for(let nShape = 0; nShape < 3; nShape++)
		{
			let shape = ws.AddShape("ellipse", 50 * 36000, 50 * 36000, api.CreateNoFill(), api.CreateStroke(0, api.CreateNoFill()), 0, 0, 0, 0);
			if (nShape !== 1)
				shape.Select();
		}

		let selectedShapes = ws.GetSelectedShapes();
		console.log(selectedShapes);

		assert.strictEqual(
			true,
			0,
			""
		);
	});

});
