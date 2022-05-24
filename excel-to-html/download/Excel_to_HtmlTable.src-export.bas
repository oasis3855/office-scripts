Attribute VB_Name = "Module1"
Option Explicit

Sub Convert2HtmlTable_EntryPoint()
    Dim wb As Workbook
    Dim workbookName() As String
    Dim strMsg As String: strMsg = ""
    
    ' Uboundが要素0でエラーとなる対策として、ダミーで最初の１つを作成
    ReDim workbookName(0)
    
    ' 起動中のExcelワークブック名を動的配列に格納
    For Each wb In Workbooks
        ReDim Preserve workbookName(UBound(workbookName) + 1)
        workbookName(UBound(workbookName)) = wb.Name
        ' InputBoxメッセージ表示用文字列
        strMsg = strMsg & UBound(workbookName) & " : " & wb.Name & vbCrLf
    Next wb
    
    ' ユーザにワークブック名を選択させる
    Dim userSelect As String
    userSelect = InputBox("No : ワークブック名" & vbCrLf & _
        "------------------" & vbCrLf & _
        strMsg, "対象のワークブックNoを選択", "1")

    ' Inputダイアログでキャンセルまたは長さゼロの文字列が入力された場合
    If StrPtr(userSelect) = 0 Or Len(userSelect) <= 0 Then
        MsgBox ("キャンセルしました")
        Exit Sub
    End If

    ' 全角英数字を半角に変換
    userSelect = StrConv(userSelect, vbNarrow)

    ' 入力値は、数値、かつワークブックのNo（1,2,3,...）範囲内
    If IsNumeric(userSelect) = False Or Val(userSelect) <= 0 Or Val(userSelect) > UBound(workbookName) Then
        MsgBox ("範囲外が入力されました")
        Erase workbookName
        Exit Sub
    End If

    ' 指定されたワークブックを変数として渡して、変換サブルーチンを呼び出す
    Set wb = Workbooks(workbookName(userSelect))
    Convert2HtmlTable wb

    Erase workbookName

End Sub

' Convert2HtmlTable : ワークシート表をHTML tableファイルに変換・エクスポートする
'
' 引数 wb : 対象ワークブック（変換されるのは、このワークブックの表示中シート）
Sub Convert2HtmlTable(wb As Workbook)
    
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet
    
    Dim MaxRange As Range
    Set MaxRange = ws.UsedRange
    Dim maxRow As Integer
    Dim maxCol As Integer
    maxRow = MaxRange.Rows.Count
    maxCol = MaxRange.Columns.Count

    Dim col As Integer
    Dim row As Integer
    ' 縦・横が結合されている行・列数（縦結合の途中行の場合は、mergeRows=-1）
    Dim mergeRows As Integer
    Dim mergeCols As Integer
    
    ' HTMLファイルに書き込む内容を一時保存する
    Dim HtmlString As String
    
    ' HTMLヘッダ部分
    HtmlString = "<html>" & vbCrLf & _
        "<head>" & vbCrLf & _
        "    <style>" & vbCrLf & _
        "    table { border-collapse: collapse; }" & vbCrLf & _
        "    td { border: 1px solid black; }" & vbCrLf & _
        "    </style>" & vbCrLf & _
        "</head>" & vbCrLf & _
        "<body>" & vbCrLf & _
        "<table>" & vbCrLf
    
    
    For row = MaxRange.row To MaxRange.row + maxRow - 1
        HtmlString = HtmlString & "    <tr>" & vbCrLf
        For col = MaxRange.Column To MaxRange.Column + maxCol - 1
            mergeRows = 1
            mergeCols = 1
            ' 指定セルが結合内にある場合、そのRangeを得る
            Dim mergeArea As Range
            Set mergeArea = ws.Cells(row, col).mergeArea
            
            If ws.Cells(row, col).MergeCells = True Then
                ' 横セル結合されている場合、結合数をmergeRowsに代入（縦のみ結合の場合は横結合無しの場合と同じ1となる）
                mergeCols = mergeArea.Item(mergeArea.Count).Column - col + 1
            End If
            If ws.Cells(row, col).MergeCells = True Then
                ' 縦セル結合されている場合、結合数をmergeRowsに代入（横のみ結合の場合は縦結合無しの場合と同じ1となる）
                mergeRows = mergeArea.Item(mergeArea.Count).row - row + 1
                If mergeArea.Item(1).row <> row Then
                    ' 縦結合で最初の行でない場合
                    mergeRows = -1
                End If
            End If
            
            ' tdタグのrowspan,colspanを一時文字列extraTdTagに格納
            Dim extraTdTag As String
            extraTdTag = ""
            If mergeCols > 1 Then
                ' 横結合されている場合
                extraTdTag = " colspan = " & mergeCols
            End If
            If mergeRows > 1 Then
                ' 縦結合されている場合
                extraTdTag = extraTdTag & " rowspan = " & mergeRows
            End If
            
            If mergeRows > 0 Then
                ' 通常のセル（縦結合されていない）
                HtmlString = HtmlString & "        <td " & extraTdTag & ">"
                HtmlString = HtmlString & ws.Cells(row, col)
                HtmlString = HtmlString & "</td>" & vbCrLf
            Else
                ' 縦結合されている場合で、2行目以降
                HtmlString = HtmlString & "        <!-- rowspan -->" & vbCrLf
            End If
            
            
            If mergeCols > 1 Then
                ' 横結合されている場合、Forループ(col)を飛ばす
                col = col + mergeCols - 1
            End If
            
        Next col
        HtmlString = HtmlString & "    </tr>" & vbCrLf
    Next row
    
    ' HTMLフッタ
    HtmlString = HtmlString & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf

    ' 情報出力
    MsgBox ("ワークブック : " & wb.Name & vbCrLf & _
            "変換範囲 : " & ConvR1c1ToA1("R" & MaxRange.row & "C" & MaxRange.Column) & " : " & _
            ConvR1c1ToA1("R" & MaxRange.row + maxRow - 1 & "C" & MaxRange.Column + maxCol - 1) & vbCrLf & vbCrLf)
        

    ' デスクトップ ディレクトリを得る
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
            
    ' SavAsダイアログで表示する初期ディレクトリに移動する
    ChDir (ShellObject.SpecialFolders("Desktop"))
    Set ShellObject = Nothing
    
    ' SaveAsダイアログ表示
    Dim OutputFilename As Variant
    OutputFilename = Application.GetSaveAsFilename("output.html", _
            "HTMLファイル,*.html,テキストファイル,*.txt")
    If OutputFilename = False Then
        MsgBox ("キャンセルしました")
        Exit Sub
    End If
    
    ' ファイルが既に存在する場合は、上書き警告する
    If Dir(OutputFilename) <> "" Then
        If Not (MsgBox("ファイル " & Dir(OutputFilename) & " に上書きします", vbYesNo) = vbYes) Then
            MsgBox ("キャンセルしました")
            Exit Sub
        End If
    End If
    
    ' ファイルに書き込む
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(OutputFilename, ForWriting, True, TristateTrue)
    ts.write (HtmlString)
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing

    MsgBox (OutputFilename & " に書き込みました")


End Sub

' 「R1C1形式の文字列」を「A1形式の文字列」に変換
Function ConvR1c1ToA1(StrR1C1 As String)
    ConvR1c1ToA1 = Application.ConvertFormula( _
        Formula:=StrR1C1, _
        fromReferenceStyle:=xlR1C1, _
        toreferencestyle:=xlA1, _
        toabsolute:=xlRelative)
End Function


