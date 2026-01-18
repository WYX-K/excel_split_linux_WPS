/**
 * 模块1：读取源工作簿
 * 功能：选择要拆分的Excel工作簿，读取工作表信息
 */
function 读取源工作簿() {
    Application.ScreenUpdating = false;
    
    var FileDialogObject = Application.FileDialog(1);
    FileDialogObject.Title = "请选择要拆分的工作簿";
    FileDialogObject.InitialFileName = ThisWorkbook.Path;
    FileDialogObject.AllowMultiSelect = false;
    
    FileDialogObject.Filters.Clear();
    FileDialogObject.Filters.Add("Excel 文件", "*.xlsx;*.xls;*.xlsm;*.et");
    FileDialogObject.Filters.Add("所有文件", "*.*");
    FileDialogObject.FilterIndex = 1;
    
    var WorkbookPaths = null;
    if (FileDialogObject.Show()) {
        WorkbookPaths = FileDialogObject.SelectedItems;
    }
    
    if (WorkbookPaths == null || WorkbookPaths.Count == 0) {
        MsgBox("你没有选中任何文件，程序退出！");
        Application.ScreenUpdating = true;
        return;
    }
    
    var Workbookpath = WorkbookPaths.Item(1);
    
    Range("A2:F2").ClearContents();
    Range("A2").Formula = Workbookpath;
    
    var lastSlashPos = Workbookpath.lastIndexOf("/");
    if (lastSlashPos == -1) {
        lastSlashPos = Workbookpath.lastIndexOf("\\");
    }
    
    if (lastSlashPos > 0) {
        Range("B3").Formula = Workbookpath.substring(0, lastSlashPos);
        Range("B4").Formula = Workbookpath.substring(lastSlashPos + 1);
    }
    
    var wb = Workbooks.Open(Workbookpath);
    
    var arr = [];
    var hiddenCount = 0;
    
    for (var i = 1; i <= wb.Sheets.Count; i++) {
        var visible = wb.Sheets.Item(i).Visible;
        if (visible === 0 || visible === 2 || visible === false) {
            hiddenCount++;
        } else {
            arr.push(wb.Sheets.Item(i).Name);
        }
    }
    
    wb.Close(false);
    
    var tbl = ActiveSheet.ListObjects("拆表参数");
    
    var rowCount = tbl.ListRows.Count;
    for (var r = rowCount; r >= 1; r--) {
        try {
            tbl.ListRows(r).Delete();
        } catch (e) {}
    }
    
    var headerRange = tbl.HeaderRowRange;
    var FirstRow = headerRange.Row + 1;
    var FirstCol = headerRange.Column;
    
    for (var i = 0; i < arr.length; i++) {
        if (i > 0) {
            try {
                tbl.ListRows.Add();
            } catch (e) {}
        }
        Cells(FirstRow + i, FirstCol).Formula = arr[i];
    }
    
    if (hiddenCount > 0) {
        MsgBox("注意！待拆分工作簿中存在【" + hiddenCount + "】个隐藏工作表\n" +
               "请检查确认，是否需要拆分或删除，以免造成敏感信息泄露！", 48, "提示");
    }
    
    Application.ScreenUpdating = true;
}
