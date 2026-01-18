/**
 * 模块4：拆分源工作簿
 * 功能：按拆分标识将源工作簿拆分为多个独立文件
 */

function _getConstants() {
    return {
        COL_SHEET_NAME: 0, COL_DEAL_TYPE: 1, COL_KEY_COLUMN: 2,
        COL_HEADER_ROWS: 3, COL_FOOTER_ROWS: 4, COL_SHEET_PWD: 5,
        COL_KEYWORD: 0, COL_SAVE_NAME: 1, COL_OPEN_PWD: 2,
        xlManual: -4135, xlAutomatic: -4105, xlFormulas: -4123,
        xlByRows: 1, xlPrevious: 2, xlSheetVisible: -1,
        xlExcelLinks: 1, xlOpenXMLWorkbook: 51
    };
}

function 拆分源工作簿() {
    var startTime = new Date().getTime();
    var C = _getConstants();
    
    setApplicationState(false);
    
    try {
        var config = loadConfiguration();
        if (!config) { setApplicationState(true); return; }
        
        var saveDir = createOutputDirectory(config.saveBaseDir);
        if (!saveDir) { setApplicationState(true); return; }
        
        var sheetConfigs = loadNamedRangeAsArray("拆表参数");
        var keywords = loadNamedRangeAsArray("拆分标识");
        
        var splitSheetCount = countSheetsByType(sheetConfigs, "拆", C);
        showProgressStart(keywords.length, splitSheetCount);
        
        var sourceWorkbook = openAndPrepareSource(config.filePath, sheetConfigs, C);
        if (!sourceWorkbook) { finishProgress(); setApplicationState(true); return; }
        
        executeSplitLoop(keywords, sheetConfigs, saveDir, splitSheetCount, C);
        
        finishProgress();
        closeWorkbookSafely(ActiveWorkbook, false);
        
        var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
        MsgBox("已完成拆分，耗时 " + elapsed + " 秒", 64);
        
    } catch (e) {
        MsgBox("拆分过程中发生错误：\n" + e.message, 16, "错误");
    } finally {
        setApplicationState(true);
    }
}

function loadConfiguration() {
    var filePath = readCellValue(Range("A2"));
    var saveBaseDir = readCellValue(Range("H2"));
    
    if (!filePath) { MsgBox("请先在 A2 单元格指定源文件路径", 16, "错误"); return null; }
    if (!saveBaseDir) { MsgBox("请先在 H2 单元格指定保存目录", 16, "错误"); return null; }
    
    return { filePath: filePath, saveBaseDir: saveBaseDir };
}

function loadNamedRangeAsArray(rangeName) {
    var range = Range(rangeName);
    var rows = range.Rows.Count;
    var cols = range.Columns.Count;
    var result = [];
    
    for (var r = 1; r <= rows; r++) {
        var rowData = [];
        for (var c = 1; c <= cols; c++) {
            rowData.push(readCellValue(range.Cells(r, c)));
        }
        result.push(rowData);
    }
    return result;
}

function createOutputDirectory(baseDir) {
    var now = new Date();
    var dateStr = (now.getMonth() + 1) + "月" + now.getDate() + "日 " +
                  now.getHours() + "时" + now.getMinutes() + "分" + now.getSeconds() + "秒";
    var saveDir = baseDir + "/拆分结果【" + dateStr + "】/";
    
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(saveDir)) { fso.CreateFolder(saveDir); }
        return saveDir;
    } catch (e) {
        try {
            MkDir(saveDir);
            return saveDir;
        } catch (e2) {
            MsgBox("无法创建目录：" + saveDir + "\n错误：" + e2.message, 16, "错误");
            return null;
        }
    }
}

function openAndPrepareSource(filePath, sheetConfigs, C) {
    try {
        Workbooks.Open(filePath);
    } catch (e) {
        MsgBox("无法打开源文件：" + filePath + "\n错误：" + e.message, 16, "错误");
        return null;
    }
    
    var wb = ActiveWorkbook;
    Calculate();
    Application.Calculation = C.xlManual;
    
    breakExternalLinks(wb, C);
    unfreezeAndShowAllSheets(wb, C);
    deleteMarkedSheets(sheetConfigs, C);
    
    return wb;
}

function breakExternalLinks(workbook, C) {
    try {
        var links = workbook.LinkSources(C.xlExcelLinks);
        if (links) {
            for (var i = 1; i <= links.length; i++) {
                workbook.BreakLink(links[i - 1], C.xlExcelLinks);
            }
        }
    } catch (e) {}
}

function unfreezeAndShowAllSheets(workbook, C) {
    for (var i = 1; i <= workbook.Worksheets.Count; i++) {
        var ws = workbook.Worksheets(i);
        ws.Activate();
        ActiveWindow.FreezePanes = false;
        ws.Visible = C.xlSheetVisible;
    }
    try { workbook.Worksheets(1).Activate(); } catch (e) {}
}

function deleteMarkedSheets(sheetConfigs, C) {
    for (var i = 0; i < sheetConfigs.length; i++) {
        var dealType = String(sheetConfigs[i][C.COL_DEAL_TYPE] || "").trim();
        if (dealType === "删") {
            try { Sheets(sheetConfigs[i][C.COL_SHEET_NAME]).Delete(); } catch (e) {}
        }
    }
}

function countSheetsByType(sheetConfigs, dealType, C) {
    var count = 0;
    for (var i = 0; i < sheetConfigs.length; i++) {
        if (String(sheetConfigs[i][C.COL_DEAL_TYPE] || "").trim() === dealType) { count++; }
    }
    return count;
}

function executeSplitLoop(keywords, sheetConfigs, saveDir, splitSheetCount, C) {
    for (var i = 0; i < keywords.length; i++) {
        var currentKeyword = keywords[i][C.COL_KEYWORD];
        if (!currentKeyword || currentKeyword === "") { continue; }
        
        updateKeywordProgress(i + 1, currentKeyword);
        
        Sheets.Copy();
        Application.Calculation = C.xlManual;
        
        processAllSheets(sheetConfigs, currentKeyword, splitSheetCount, C);
        removeEmptySheet1();
        
        Application.Calculation = C.xlAutomatic;
        saveAndCloseWorkbook(saveDir, keywords[i], C);
    }
}

function processAllSheets(sheetConfigs, currentKeyword, splitSheetCount, C) {
    var sheetProcessed = 0;
    
    for (var j = 0; j < sheetConfigs.length; j++) {
        var config = sheetConfigs[j];
        var sheetName = config[C.COL_SHEET_NAME];
        var dealType = String(config[C.COL_DEAL_TYPE] || "").trim();
        var deleted = false;
        
        if (dealType === "拆") {
            sheetProcessed++;
            deleted = processSplitSheet(config, currentKeyword, sheetProcessed, splitSheetCount, C);
        }
        
        if (dealType !== "删" && !deleted) {
            applySheetProtection(sheetName, config[C.COL_SHEET_PWD]);
        }
    }
}

function processSplitSheet(config, currentKeyword, sheetIndex, splitSheetCount, C) {
    var sheetName = config[C.COL_SHEET_NAME];
    var keyColumn = String(config[C.COL_KEY_COLUMN] || "").trim();
    var headerRows = parseInt(config[C.COL_HEADER_ROWS]) || 0;
    var footerRows = parseInt(config[C.COL_FOOTER_ROWS]) || 0;
    
    var sheet = Sheets(sheetName);
    var totalRows = getLastRow(sheet, C);
    var bodyRows = totalRows - headerRows - footerRows;
    
    updateSheetProgress(sheetIndex, splitSheetCount, bodyRows);
    sheet.Select();
    convertFormulasToValues(sheet);
    
    var deletedRows = deleteNonMatchingRows(sheet, keyColumn, headerRows, bodyRows, currentKeyword);
    
    if (deletedRows === bodyRows) {
        try { sheet.Delete(); return true; } catch (e) {}
    }
    return false;
}

function convertFormulasToValues(sheet) {
    try { sheet.UsedRange.Value = sheet.UsedRange.Value; } catch (e) {}
}

function deleteNonMatchingRows(sheet, keyColumn, headerRows, bodyRows, keyword) {
    var deletedCount = 0;
    for (var k = bodyRows; k >= 1; k--) {
        var rowIndex = headerRows + k;
        var cellValue = readCellValue(sheet.Range(keyColumn + rowIndex));
        if (String(cellValue) !== String(keyword)) {
            sheet.Range(keyColumn + rowIndex).EntireRow.Delete();
            deletedCount++;
        }
        updateRowProgress(bodyRows - k + 1);
    }
    return deletedCount;
}

function getLastRow(sheet, C) {
    try {
        var found = sheet.Cells.Find("*", undefined, C.xlFormulas, undefined, C.xlByRows, C.xlPrevious);
        return found ? found.Row : 1;
    } catch (e) { return 1; }
}

function applySheetProtection(sheetName, password) {
    var pwdStr = (password != null && password !== undefined) ? String(password).trim() : "";
    if (pwdStr === "") return;
    try { ActiveWorkbook.Sheets(sheetName).Protect(pwdStr); } catch (e) {}
}

function removeEmptySheet1() {
    try {
        var sheet1 = null;
        try { sheet1 = Sheets("Sheet1"); } catch (e) { return; }
        if (sheet1 && Sheets.Count > 1) {
            if (Application.WorksheetFunction.CountA(sheet1.Cells) === 0) { sheet1.Delete(); }
        }
    } catch (e) {}
}

function saveAndCloseWorkbook(saveDir, keywordRow, C) {
    var savePath = saveDir + keywordRow[C.COL_SAVE_NAME] + ".xlsx";
    var openPwd = keywordRow[C.COL_OPEN_PWD];
    
    try {
        if (openPwd) { ActiveWorkbook.SaveAs(savePath, C.xlOpenXMLWorkbook, openPwd); }
        else { ActiveWorkbook.SaveAs(savePath, C.xlOpenXMLWorkbook); }
    } catch (e) {
        MsgBox("保存文件失败：" + savePath + "\n错误：" + e.message, 16, "错误");
    }
    closeWorkbookSafely(ActiveWorkbook, false);
}

function readCellValue(cell) {
    var val = cell.Value2;
    if (val === undefined) { val = cell.Value; if (val === undefined) { val = cell.Formula; } }
    if (typeof val === "function") { try { val = val(); } catch (e) { val = ""; } }
    return (val != null && val !== undefined) ? String(val) : "";
}

function closeWorkbookSafely(workbook, saveChanges) {
    try { workbook.Close(saveChanges); } catch (e) {}
}

function setApplicationState(enabled) {
    Application.ScreenUpdating = enabled;
    Application.DisplayAlerts = enabled;
}

function _getProgressState() {
    if (typeof ThisWorkbook._progressState === "undefined") {
        ThisWorkbook._progressState = { currentKeyword: "", keywordCurrent: 0, keywordTotal: 0, sheetCurrent: 0, sheetTotal: 0, startTime: 0 };
    }
    return ThisWorkbook._progressState;
}

function showProgressStart(totalKeywords, totalSheets) {
    var state = _getProgressState();
    state.keywordTotal = totalKeywords || 0;
    state.sheetTotal = totalSheets || 0;
    state.startTime = new Date().getTime();
    Application.ScreenUpdating = true;
    Application.StatusBar = "正在准备拆分，共 " + totalKeywords + " 个拆分标识...";
    Application.ScreenUpdating = false;
}

function updateKeywordProgress(current, keyword) {
    var state = _getProgressState();
    state.keywordCurrent = current;
    state.currentKeyword = keyword || "";
    state.sheetCurrent = 0;
    try { ActiveWorkbook.Worksheets(1).Activate(); } catch (e) {}
    Application.ScreenUpdating = true;
    Application.StatusBar = "正在拆分【" + keyword + "】 (" + current + "/" + state.keywordTotal + ")";
    Application.ScreenUpdating = false;
}

function updateSheetProgress(current, sheetTotal, totalRows) {
    var state = _getProgressState();
    state.sheetCurrent = current;
    state.sheetTotal = sheetTotal || state.sheetTotal;
}

function updateRowProgress(current) {}

function finishProgress() {
    Application.StatusBar = false;
    ThisWorkbook._progressState = undefined;
}
