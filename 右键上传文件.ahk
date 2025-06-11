; 右键上传文件脚本
#NoEnv
#SingleInstance force
SendMode Input
SetWorkingDir %A_ScriptDir%

; 添加右键菜单项
RegWrite, REG_SZ, HKEY_CLASSES_ROOT\*\shell\上传文件到指定地址, , 上传到指定地址
RegWrite, REG_SZ, HKEY_CLASSES_ROOT\*\shell\上传文件到指定地址\command, , "%A_AhkPath%" "%A_ScriptFullPath%" "`%1"

; 主处理函数
if (A_Args.Length() > 0) {
    filePath := A_Args[1]
    
    ; 读取INI文件获取上次保存的路径
    IniRead, defaultDir, %A_ScriptDir%\upload_config.ini, Settings, DefaultDir, %A_MyDocuments%
    
    ; 显示文件夹选择对话框
    FileSelectFolder, selectedDir, *%defaultDir%, 3, 请选择目标文件夹
    if (selectedDir = "") {
        return  ; 用户取消选择
    }
    
    ; 保存选择的路径到INI文件
    IniWrite, %selectedDir%, %A_ScriptDir%\upload_config.ini, Settings, DefaultDir
    
    ; 获取文件名
    SplitPath, filePath, fileName
    
    ; 复制文件
    destPath := selectedDir "\" fileName
    FileCopy, %filePath%, %destPath%, 1
    
    ; 显示成功提示
    if (ErrorLevel = 0) {
        MsgBox, 64, 上传成功, 文件已成功复制到：%destPath%
    } else {
        MsgBox, 16, 上传失败, 文件复制失败，请重试
    }
}

return