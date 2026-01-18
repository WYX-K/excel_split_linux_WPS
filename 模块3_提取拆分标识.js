/**
 * 模块3：提取拆分标识
 * 功能：从工作簿中提取所有拆分标识并去重
 */

function 提取拆分标识() {
    Application.ScreenUpdating = false;
    var startTime = new Date().getTime();
    
    var ShtName = 0, Deal = 1, KeyPos = 2, HeaderPos = 3, FootPos = 4;
    
    var ShtRange = Range("拆表参数");
    var Sht = _getArrayFromRange_M3(ShtRange);
    
    var FilePath = String(_getCellValue(Range("A2")) || "");
    var NamePre = String(_getCellValue(Range("I3")) || "");
    var NameSuf = String(_getCellValue(Range("I4")) || "");
    
    if (FilePath == "") {
        MsgBox("请先运行【模块1_读取源工作簿】选择要拆分的文件！", 48, "提示");
        Application.ScreenUpdating = true;
        return;
    }
    
    var wb;
    try {
        wb = Workbooks.Open(FilePath);
    } catch (e) {
        MsgBox("无法打开工作簿：" + FilePath + "\n错误：" + e.message, 16, "错误");
        Application.ScreenUpdating = true;
        return;
    }
    
    var arr = [];
    
    for (var k = 0; k < Sht.length; k++) {
        var dealType = String(Sht[k][Deal] || "").trim();
        
        if (dealType == "拆") {
            var SheetName = Sht[k][ShtName];
            var KeyCol = Sht[k][KeyPos];
            var FirstRows = parseInt(Sht[k][HeaderPos]) || 0;
            var EndRows = parseInt(Sht[k][FootPos]) || 0;
            
            var ws = wb.Sheets(SheetName);
            var findResult = ws.Cells.Find("*", undefined, -4123, undefined, 1, 2);
            var TotalRows = findResult ? findResult.Row : 0;
            var BodyRows = TotalRows - EndRows - FirstRows;
            
            for (var j = 1; j <= BodyRows; j++) {
                var cellValue = _getCellValue(ws.Range(KeyCol + (FirstRows + j)));
                arr.push(cellValue);
            }
        }
    }
    
    var uniqueObj = {};
    for (var i = 0; i < arr.length; i++) {
        var key = String(arr[i]);
        if (!uniqueObj.hasOwnProperty(key)) {
            uniqueObj[key] = arr[i];
        }
    }
    
    var drr = [];
    for (var key in uniqueObj) {
        if (uniqueObj.hasOwnProperty(key)) {
            drr.push(uniqueObj[key]);
        }
    }
    
    wb.Close(false);
    
    var tbl = ActiveSheet.ListObjects("拆分标识");
    var rowCount = tbl.ListRows.Count;
    for (var r = rowCount; r >= 1; r--) {
        try {
            tbl.ListRows(r).Delete();
        } catch (e) {}
    }
    
    var headerRange = tbl.HeaderRowRange;
    var FirstRow = headerRange.Row + 1;
    var FirstCol = headerRange.Column;
    
    for (var k = 0; k < drr.length; k++) {
        if (k > 0) {
            try {
                tbl.ListRows.Add();
            } catch (e) {}
        }
        Cells(FirstRow + k, FirstCol).Formula = drr[k];
        Cells(FirstRow + k, FirstCol + 1).Formula = (NamePre || "") + drr[k] + (NameSuf || "");
    }
    
    var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
    MsgBox("已完成，耗时" + elapsed + "秒", 64);
    
    Application.ScreenUpdating = true;
}

function _getCellValue(cell) {
    var val = cell.Value2;
    if (val === undefined) {
        val = cell.Value;
        if (val === undefined) {
            val = cell.Formula;
        }
    }
    if (typeof val === 'function') {
        try { val = val(); } catch (e) { val = ""; }
    }
    return val;
}

function _getArrayFromRange_M3(range) {
    var rows = range.Rows.Count;
    var cols = range.Columns.Count;
    var arr = [];
    
    for (var r = 1; r <= rows; r++) {
        var rowArr = [];
        for (var c = 1; c <= cols; c++) {
            rowArr.push(_getCellValue(range.Cells(r, c)));
        }
        arr.push(rowArr);
    }
    return arr;
}
