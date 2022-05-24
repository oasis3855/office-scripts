Attribute VB_Name = "Module1"
Option Explicit

' 検索ボタンが押された時の処理
Sub ButtonFindAllComment()

    FindCommentCell (False)
End Sub

' 削除ボタンが押された時の処理
Sub ButtonDeletedAllComment()

    FindCommentCell (True)

End Sub

' ユーザが指定するワークブックより、コメントを検索・削除する
Function FindCommentCell(ByVal bClearComment As Boolean)
    
    ' 戻り値（文字列）の初期化
    FindCommentCell = ""
    
    ' デスクトップのフルパス名を得る
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = wsh.SpecialFolders("Desktop")
    
    ' ファイルダイアログを表示
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "コメントを検索するファイルの選択"
        With .Filters
            .Clear
            .Add "Excelワークブック", "*.xlsx; *.xls; *.xlsm", 1
        End With
        ' フォルダ名が \ で終了していないと、ファイル名として扱われる場合がある対処
        If Right(DesktopFolderName, 1) <> "\" Then
            DesktopFolderName = DesktopFolderName + "\"
        End If
        
        .InitialFileName = DesktopFolderName
    End With
    If fd.Show <> True Then
        MsgBox ("キャンセルされました")
        Exit Function
    End If
    MsgBox ("対象ファイル " & fd.SelectedItems(1) & " が選択されました")

    ' 結果表示を格納する文字列
    Dim strMsg As String
    strMsg = "コメントのあるセル一覧" & vbCrLf & vbCrLf

    ' ステータスバーの書き換えを許可する（現在のモードは保存しておく）
    Dim boolModeDispStatusbar As Boolean
    boolModeDispStatusbar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Application.StatusBar = fd.SelectedItems(1) & " を開いています ..."
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    ' 検索対象ワークブック
    Dim wb As Workbook
    Set wb = xl.Workbooks.Open(fd.SelectedItems(1))
    Set xl = Nothing
    ' 検索対象ワークシート
    Dim ws As Worksheet
    ' Set ws = wb.Worksheets(1)  ' デバッグ用 １個目のワークシートのみ
    
    For Each ws In wb.Worksheets
        Application.StatusBar = ws.Name & " 検索中 ..."
        strMsg = strMsg & ws.Name & ": " & FindCommentCellOnWorksheet(ws, bClearComment) & vbCrLf
    Next ws
    Application.StatusBar = False
    Application.DisplayStatusBar = boolModeDispStatusbar
    
    If bClearComment = True Then
        MsgBox (strMsg & vbCrLf & "次のダイアログでこれらのコメントを削除し、新規ファイルに保存できます")
        ' コメント削除したワークブックを、名前を付けて保存する
        Dim result As Boolean
        result = SaveAsNewfile(wb, fd.SelectedItems(1))
        If result = True Then
            MsgBox ("新規ファイルに保存しました")
        Else
            MsgBox ("キャンセルしました")
        End If
        
    Else
        MsgBox (strMsg)
    End If

    ' 検索対象ワークブックを閉じる
    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing

End Function

' 指定したワークシート内のコメントを検索しセルアドレスを返す。またコメントを削除する
Function FindCommentCellOnWorksheet(ByVal ws As Worksheet, ByVal bClearComment As Boolean) As String

    ' 戻り値（文字列）の初期化
    FindCommentCellOnWorksheet = ""

    Dim CommentCells As Range

    On Error Resume Next    ' シート内に一つも該当セルが無い場合、「該当するセルが見つかりません」エラーを回避
    Set CommentCells = ws.Cells.SpecialCells(xlCellTypeComments)
    On Error GoTo 0
    
    Dim str As String
    If CommentCells Is Nothing Then
        str = ""
    Else
        str = CommentCells.Address
        ' 絶対アドレスから、見やすいように「$」を除去する
        str = Replace(str, "$", "")
        
        ' コメント1個ずつにアクセスする方法
        'Dim c As Range
        'Dim temp As String
        'For Each c In CommentCells
        '    temp = c.Address & ":" & c.Comment.Text
        'Next c
        
        'コメントの削除
        If bClearComment = True Then
            CommentCells.ClearComments
        End If
    End If
    
    FindCommentCellOnWorksheet = str

End Function

' 名前を付けてワークブックを保存する
Function SaveAsNewfile(ByVal wb As Workbook, ByVal filepath As String)
On Error GoTo catch
    ' 名前を付けて保存する ダイアログを表示する
    filepath = Application.GetSaveAsFilename(InitialFileName:=filepath, Title:="コメント除去後のファイルを名前を付けて保存する")
    If filepath = "False" Then
        SaveAsNewfile = False
        Exit Function
    End If
    ' ファイルに保存する
    Call wb.SaveAs(filepath)
    SaveAsNewfile = True
    Exit Function
    
catch:
    SaveAsNewfile = False
    Exit Function
End Function
