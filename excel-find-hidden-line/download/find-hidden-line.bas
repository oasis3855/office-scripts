Attribute VB_Name = "Module1"
Option Explicit

Sub FindHiddenLine()
    
    ' デスクトップのフルパス名を得る
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = wsh.SpecialFolders("Desktop")
    
    ' ファイルダイアログを表示
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "非表示行・列を検索するファイルの選択"
        With .Filters
            .Clear
            .Add "Excelワークブック", "*.xlsx; *.xls; *.xlsm", 1
        End With
        .InitialFileName = DesktopFolderName
    End With
    If fd.Show <> True Then
        MsgBox ("キャンセルされました")
        Exit Sub
    End If
    MsgBox ("対象ファイル " & fd.SelectedItems(1) & " が選択されました")

    ' 結果表示を格納する文字列
    Dim strMsg As String
    strMsg = ""


    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    ' 検索対象ワークブック
    Dim wb As Workbook
    Set wb = xl.Workbooks.Open(fd.SelectedItems(1))
    Set xl = Nothing
    ' 検索対象ワークシート
    Dim ws As Worksheet
    ' Set ws = wb.Worksheets(1)  ' デバッグ用 １個目のワークシートのみ
    
    Dim boolModeDispStatusbar As Boolean
    boolModeDispStatusbar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    For Each ws In wb.Worksheets
        Application.StatusBar = ws.Name & " 検索中 ..."
        strMsg = strMsg & ws.Name & ": " & FindHiddenLineOnWorksheet(ws) & vbCrLf
    Next ws
    Application.StatusBar = False
    Application.DisplayStatusBar = boolModeDispStatusbar
    
    ' 検索対象ワークブックを閉じる
    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing

    MsgBox (strMsg)

End Sub


Function FindHiddenLineOnWorksheet(ByVal ws As Worksheet) As String

    ' 戻り値（文字列）の初期化
    FindHiddenLineOnWorksheet = ""
    Dim i As Integer
    
    Dim maxrow As Integer
    ' ws.UsedRange.Row : データが始まる最初の行
    ' ws.UsedRange.Rows.Count : データが入力されている行範囲の行数
    maxrow = ws.UsedRange.Rows.Count + ws.UsedRange.Row
    
    For i = 1 To maxrow
        If ws.Rows(i).Hidden = True Then
            ' MsgBox ("Row = " & i & " hidden")
            FindHiddenLineOnWorksheet = FindHiddenLineOnWorksheet & i & " 行目,"
        End If
    Next i

    Dim maxcol As Integer
    maxcol = ws.UsedRange.Columns.Count + ws.UsedRange.Column
    
    For i = 1 To maxrow
        If ws.Columns(i).Hidden = True Then
            ' MsgBox ("Col = " & i & " hidden")
            ' FindHiddenLineOnWorksheet = FindHiddenLineOnWorksheet & i & " 列目,"
            FindHiddenLineOnWorksheet = FindHiddenLineOnWorksheet & Split(Cells(1, i).Address(True, False), "$")(0) & " 列目,"
        End If
    Next i
End Function
