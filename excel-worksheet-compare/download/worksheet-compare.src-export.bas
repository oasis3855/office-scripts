Attribute VB_Name = "Module1"
Public wbThisWorkbook As Workbook


Sub Button_CompareWorksheet()
    Set wbThisWorkbook = ActiveWorkbook
    
    ' 添付ファイルのフルパス名を作成する
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = ShellObject.SpecialFolders("Desktop")

    ' 出力ワークブック（新規作成）
    Dim wbOutput As Workbook
    Dim wsOutput As Worksheet

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "「比較元」ファイルの選択"
        With .Filters
            .Clear
            .Add "Excelワークブック", "*.xlsx; *.xls; *.xlsm", 1
        End With
        .InitialFileName = DesktopFolderName
    End With

    If fd.Show = True Then
        MsgBox ("比較元ファイル " & fd.SelectedItems(1) & " が選択されました")

        ' エクスポートワークブック（wb）を新規作成
        Set wbOutput = Workbooks.Add
        ' エクスポートワークブック（wb）に新しいワークシート「各事業の光熱水費」を追加
        Set wsOutput = wbOutput.Worksheets.Add
        wsOutput.Name = "比較元"
        
        If CopyWorksheet(wbOutput, wsOutput, fd.SelectedItems(1)) < 0 Then
            ' 異常終了またはキャンセルボタンが押された場合
            MsgBox ("キャンセルが押された")
            ' 出力ワークブックを閉じる（保存せず、ウィンドウ消去）
            wbOutput.Close SaveChanges:=False
            Set wbOutput = Nothing
            Exit Sub
        End If

        ' 出力ワークシートを閉じる
        Set wsOutput = Nothing
    Else
        MsgBox ("キャンセルが押された")
        ' 出力ワークブックを閉じる（まだウィンドウは開かれていないが、明示的に閉じる）
        Set wbOutput = Nothing
        Exit Sub
    End If
    
    Set fd = Nothing
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "「比較先」ファイルの選択"
        With .Filters
            .Clear
            .Add "Excelワークブック", "*.xlsx; *.xls; *.xlsm", 1
        End With
        .InitialFileName = DesktopFolderName
    End With

    If fd.Show = True Then
        MsgBox ("比較先ファイル " & fd.SelectedItems(1) & " が選択されました")

        ' エクスポートワークブック（wb）に新しいワークシート「各事業の光熱水費」を追加
        Set wsOutput = wbOutput.Worksheets.Add
        wsOutput.Name = "比較先"

        If CopyWorksheet(wbOutput, wsOutput, fd.SelectedItems(1)) < 0 Then
            ' 異常終了またはキャンセルボタンが押された場合
            MsgBox ("キャンセルが押された")
            ' 出力ワークブックを閉じる（保存せず、ウィンドウ消去）
            wbOutput.Close SaveChanges:=False
            Set wbOutput = Nothing
            Exit Sub
        End If

        ' 出力ワークシートを閉じる
        Set wsOutput = Nothing
    Else
        MsgBox ("キャンセルが押された")
        ' 出力ワークブックを閉じる（保存せず、ウィンドウ消去）
        wbOutput.Close SaveChanges:=False
        Set wbOutput = Nothing
        Exit Sub
    End If

    ' ファイル選択ダイアログを開放する
    Set fd = Nothing

    ' ワークシートを比較し着色する
    Call CompareWorksheet(wbOutput)

    ' 出力ワークブックを閉じる（ウィンドウ自体は閉じないで結果表示として残す）
    Set wbOutput = Nothing

    MsgBox ("比較処理が終了しました")
End Sub

'  ワークブックの中の1枚のワークシートを、引数として指定されたワークブックにコピー（値コピー）する
'  （ワークブック内の全ワークシートをリスト表示し、ユーザに選択させるUserForm機能あり）
'
'  戻り値：1=正常, -1=異常（キャンセル）
'
Function CopyWorksheet(ByRef wbOutput As Workbook, ByRef wsOutput As Worksheet, wbSourceFilename As String) As Integer
    CopyWorksheet = 1   ' 戻り値 = 1 （正常終了）
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    ' データ抽出元ワークブック（このワークブック）
    Dim wbSource As Workbook
    Set wbSource = xl.Workbooks.Open(wbSourceFilename)
    Set xl = Nothing
    Dim wsSource As Worksheet

    ' ワークシート一覧のリストボックスより、対象とするワークシートを選択する
    Dim ws As Worksheet
    For Each ws In wbSource.Worksheets
        ' リストボックスに一つづつワークシート名を追加
        UserFormWs.ListBoxWs.AddItem (ws.Name)
    Next ws
    UserFormWs.ListBoxWs.ListIndex = 0
    Dim selectedIndex As Integer
    selectedIndex = UserFormWs.doModal()
    If selectedIndex < 0 Then
        CopyWorksheet = -1   ' 戻り値 = -1 （異常終了）
        ' コピー元のワークブックを閉じる
        wbSource.Close SaveChanges:=False
        Set wbSource = Nothing
        Exit Function
    End If

    Set wsSource = wbSource.Worksheets(selectedIndex + 1)   ' ワークシートNo.は1～のため、リストボックスNo.に1加える

    ' データが入力されている最大カラム数、行数を得る
    Dim MaxRange As Range
    Set MaxRange = wsSource.UsedRange
    Dim maxRow As Integer
    Dim maxCol As Integer
    maxRow = MaxRange.Rows.Count
    maxCol = MaxRange.Columns.Count
    
    Dim i As Integer
    Dim j As Integer
    
    ' 自動計算を一時的に無効化
    Application.Calculation = xlCalculationManual
    
    
    ' 逐次コピーは遅いので、Range.Copy and  Range.PasteSpecial に変更した
'    For i = 1 To maxRow
'        For j = 1 To maxCol
'            wsOutput.Cells(i, j) = wsSource.Cells(i, j).Text
'        Next j
'    Next i
    
    ' コピー ＆ ペースト
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(maxRow, maxCol)).Copy
    wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(maxRow, maxCol)).PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True, operation:=xlPasteSpecialOperationNone
    ' 貼り付けた範囲内の書式をクリア
    Dim wsThisWorksheet As Worksheet
    Set wsThisWorksheet = wbThisWorkbook.Worksheets(1)
    ' 書式をコピーする
    If wsThisWorksheet.CheckBoxes(1).Value = xlOn Then
        ' 背景色のみクリア
        wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(maxRow, maxCol)).Interior.ColorIndex = xlNone
    Else
        ' 全ての書式をクリア
        wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(maxRow, maxCol)).ClearFormats
    End If
    ' 列幅をコピーする
    If wsThisWorksheet.CheckBoxes(2).Value = xlOn Then
        For i = 1 To maxCol
            wsOutput.Columns(i).ColumnWidth = wsSource.Columns(i).ColumnWidth
        Next i
    End If
    ' 行高さをコピーする
    If wsThisWorksheet.CheckBoxes(2).Value = xlOn Then
        For i = 1 To maxRow
            wsOutput.Rows(i).RowHeight = wsSource.Rows(i).RowHeight
        Next i
    End If
    
    
    ' クリップボードの削除処理
    Application.CutCopyMode = False     ' クリップボードをクリアする
    wsSource.Range("A1").Copy           ' ファイルを閉じるときに「クリップボードに大きな情報があります」確認dlgを抑止するためのダミーコピー
    wsOutput.Range("A1").Select     ' コピー範囲が全選択になっていたのを解除する（A1選択に戻る）

    ' 自動計算を有効化
    Application.Calculation = xlCalculationAutomatic

    ' コピー元のワークブックを閉じる
    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing

End Function

'  引数で指定されたワークブックの左から2枚のワークシートを比較し、異なる値のセルを赤で着色する
'
Sub CompareWorksheet(ByRef wbSource As Workbook)
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")

    Dim wsSource01 As Worksheet
    Dim wsSource02 As Worksheet
    Set wsSource01 = wbSource.Worksheets(1)
    Set wsSource02 = wbSource.Worksheets(2)


    Dim MaxRange As Range
    Set MaxRange = wsSource01.UsedRange
    
    Dim maxRow As Integer
    Dim maxCol As Integer
    maxRow = MaxRange.Rows.Count
    maxCol = MaxRange.Columns.Count
    
    Dim i As Integer
    Dim j As Integer
    
    ' 自動計算を一時的に無効化
    Application.Calculation = xlCalculationManual
    
    ' 差がある部分を着色
    For i = 1 To maxRow
        For j = 1 To maxCol
            If wsSource01.Cells(i, j).Text <> wsSource02.Cells(i, j).Text Then
                wsSource01.Cells(i, j).Interior.Color = RGB(255, 216, 216)
                wsSource02.Cells(i, j).Interior.Color = RGB(255, 216, 216)
            End If
        Next j
    Next i
    
    ' 自動計算を有効化
    Application.Calculation = xlCalculationAutomatic


End Sub
