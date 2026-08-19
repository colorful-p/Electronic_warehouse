Attribute VB_Name = "MaterialExport"
' ============================================================================
'  物料库导出 + 自动推送
' ============================================================================
'  使用方法:
'    1. 打开元器件管理表.xlsm（启用宏）
'    2. 按 Alt+F8 运行 ExportAndOpen
'    3. 宏会自动完成:
'       a. 清理 Excel 文档属性中的个人信息
'       b. 读取当前工作表生成 data.js
'       c. 用 git 推送到 GitHub
'       d. 打开本地网站预览
'
'  注意: GitHub Pages 会自动更新线上网站，约 1 分钟后生效
'  网页显示完全跟随 Excel 列，表头是什么就显示什么
'
'  隐私说明
'    - 自动清理文档属性，避免泄露姓名等个人信息
'    - git 提交身份使用中性昵称
' ============================================================================

' ----------------------------------------------------------------------------
'  主流程: 清理属性 -> 生成数据 -> 推送 -> 打开网站
' ----------------------------------------------------------------------------
Sub ExportAndOpen()
    Dim pushResult As String

    ' 1. 清理文档属性并保存
    Call CleanDocProperties
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    ' 2. 生成 data.js
    Call ExportDataJS

    ' 3. git 推送到 GitHub
    pushResult = GitPush()

    ' 4. 打开本地网站
    Call OpenWebsite

    ' 5. 显示推送结果
    If pushResult <> "" Then
        MsgBox pushResult, vbInformation, "物料库更新"
    End If
End Sub

' ----------------------------------------------------------------------------
'  读取当前工作表，动态生成 data.js（列完全跟随表头）
' ----------------------------------------------------------------------------
Sub ExportDataJS()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 获取数据范围
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "没有数据可导出：第 1 行必须是表头，第 2 行起才是数据。", vbExclamation
        Exit Sub
    End If

    ' ---- 读取表头（空列自动跳过）----
    Dim headers() As String
    ReDim headers(1 To lastCol)
    Dim colIndex As Long
    Dim headerText As String
    Dim hasHeader As Boolean
    hasHeader = False

    For colIndex = 1 To lastCol
        headerText = Trim(ws.Cells(1, colIndex).Value)
        If headerText <> "" Then
            headers(colIndex) = headerText
            hasHeader = True
        Else
            headers(colIndex) = ""
        End If
    Next colIndex

    If Not hasHeader Then
        MsgBox "第 1 行没有找到表头。", vbExclamation
        Exit Sub
    End If

    ' ---- 拼接 data.js 内容 ----
    Dim s As String
    s = "// 物料库数据由 VBA 宏自动生成，请勿手动修改" & vbCrLf
    s = s & "// 生成时间: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    s = s & "// 表头: "
    Dim firstCol As Boolean
    firstCol = True
    For colIndex = 1 To lastCol
        If headers(colIndex) <> "" Then
            If Not firstCol Then s = s & " | "
            s = s & headers(colIndex)
            firstCol = False
        End If
    Next colIndex
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "const materials = [" & vbCrLf

    Dim i As Long
    Dim val As String
    Dim firstField As Boolean

    For i = 2 To lastRow
        s = s & "  {"

        firstField = True
        For colIndex = 1 To lastCol
            If headers(colIndex) <> "" Then
                val = Trim(ws.Cells(i, colIndex).Value)

                ' 只导出非空字段
                If val <> "" Then
                    If Not firstField Then
                        s = s & ","
                    End If
                    firstField = False
                    s = s & escapeJS(headers(colIndex)) & ":" & escapeJS(val)
                End If
            End If
        Next colIndex

        s = s & "}"

        If i <> lastRow Then
            s = s & ","
        End If
        s = s & vbCrLf
    Next i

    s = s & "];"

    ' ---- 以 UTF-8 编码写入 data.js ----
    Dim exportPath As String
    exportPath = ThisWorkbook.Path & "\data.js"

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.SaveToFile exportPath, 2
    stm.Close

    Exit Sub

ErrorHandler:
    MsgBox "导出失败" & vbCrLf & vbCrLf & _
           "错误编号: " & Err.Number & vbCrLf & _
           "错误描述: " & Err.Description, vbCritical, "物料库导出"
End Sub

' ----------------------------------------------------------------------------
'  清理文档属性，避免个人信息泄露
' ----------------------------------------------------------------------------
Sub CleanDocProperties()
    On Error Resume Next
    ThisWorkbook.BuiltinDocumentProperties("Author") = "warehouse"
    ThisWorkbook.BuiltinDocumentProperties("Last Author") = "warehouse"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
'  在 Excel 所在目录执行 git 命令
' ----------------------------------------------------------------------------
Function RunGit(args As String) As Long
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = ThisWorkbook.Path
    RunGit = wsh.Run("cmd /c git " & args, 0, True)
End Function

' ----------------------------------------------------------------------------
'  git 提交并推送到 GitHub
' ----------------------------------------------------------------------------
Function GitPush() As String
    ' 检查是否在 git 仓库内
    If Dir(ThisWorkbook.Path & "\.git", vbDirectory) = "" Then
        GitPush = "未找到 git 仓库（.git 文件夹），已跳过推送。" & vbCrLf & _
                  "data.js 已生成，网站本地可用。"
        Exit Function
    End If

    Dim code As Long
    Dim msg As String

    ' 暂存所有改动（.gitignore 已排除无关文件）
    code = RunGit("add -A")
    If code <> 0 Then
        GitPush = "git add 失败（代码 " & code & "），文件未推送。"
        Exit Function
    End If

    ' 提交（无改动时 commit 返回 0 不影响后续）
    Dim commitMsg As String
    commitMsg = "update " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    RunGit("commit -m """ & commitMsg & """")

    ' 推送，凭据在本地 .git/config 中
    code = RunGit("push")
    If code = 0 Then
        GitPush = "已推送到 GitHub，约 1 分钟后线上自动更新。"
    Else
        GitPush = "推送失败（代码 " & code & "）。" & vbCrLf & _
                  "请检查网络或 GitHub 令牌。" & vbCrLf & _
                  "data.js 已生成并保存在本地。"
    End If
End Function

' ----------------------------------------------------------------------------
'  打开本地网站预览
' ----------------------------------------------------------------------------
Sub OpenWebsite()
    Dim htmlPath As String
    htmlPath = ThisWorkbook.Path & "\index.html"

    If Dir(htmlPath) = "" Then
        MsgBox "未找到 index.html，请确认网站文件完整。", vbExclamation
        Exit Sub
    End If

    ActiveWorkbook.FollowHyperlink htmlPath
End Sub

' ----------------------------------------------------------------------------
'  将字符串转义为 JavaScript 字符串字面量
' ----------------------------------------------------------------------------
Function escapeJS(ByVal s As String) As String
    Dim result As String
    result = s

    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    escapeJS = """" & result & """"
End Function
