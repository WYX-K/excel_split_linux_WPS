/**
 * 模块2：选择存放目录
 * 功能：选择拆分结果的存放目录
 */

function 选择存放目录() {
    Application.ScreenUpdating = false;
    
    var FolderDialogObject = Application.FileDialog(4);
    FolderDialogObject.Title = "请选择存放目录";
    FolderDialogObject.InitialFileName = ThisWorkbook.Path;
    FolderDialogObject.AllowMultiSelect = false;
    
    FolderDialogObject.Show();
    
    var FolderPaths = FolderDialogObject.SelectedItems;
    
    if (FolderPaths.Count == 0) {
        MsgBox("你没有选中任何文件夹，程序退出！");
        Application.ScreenUpdating = true;
        return;
    }
    
    var FolderPath = FolderPaths.Item(1);
    
    Range("H2:J2").ClearContents();
    Range("H2").Formula = FolderPath;
    
    Application.ScreenUpdating = true;
}
