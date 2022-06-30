Attribute VB_Name = "Module1"
Option Explicit

' 発送名簿ワークシートのデータスタート行（先頭行は文字列置換のキーワード項目名）
Public Const START_ROW As Integer = 2

' デバッグ：先頭5行の実際の実行する行数
Public Const DEBUG5LINE_LINE = 5

' エラートラップを有効化する
Public Const ERROR_TRAP_ENABLE = True

' 添付ファイル編集保護パスワード
Dim strAttachWordPassword As String


Sub Button_SendMail()

    ' エラーをトラップし、VBEデバッグダイアログを表示させずに、スクリプトを強制終了する
    If ERROR_TRAP_ENABLE = True Then On Error GoTo ERRTRAP_Button_SendMail
    Err.Clear

    Dim wbCurrent As Workbook
    Set wbCurrent = ActiveWorkbook
    Dim wsAddrbook As Worksheet
    Set wsAddrbook = wbCurrent.Worksheets("名簿")
    Dim wsControl As Worksheet
    Set wsControl = wbCurrent.Worksheets("制御画面")
    
    ' 発送名簿のカラム名、カラムNoを格納した配列
    Dim arrKey() As String, arrColNo() As Integer
    Call SetArrayColName(arrKey(), arrColNo(), wsAddrbook)
    
    Dim i As Integer, j As Integer

    ' デバッグ（Outlook出力ではなく、デスクトップ上にログファイルを出力）
    Dim DEBUG_NO_OUTLOOK As Boolean: DEBUG_NO_OUTLOOK = False
    If wsControl.CheckBoxes("chkbox_debugtxt").Value = xlOn Then DEBUG_NO_OUTLOOK = True
    ' デバッグファイルのためのファイルハンドル
    Dim fn As Integer

    ' デバッグ出力ファイルのフルパス名を作成する
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim filepathDebugOutput As String
    filepathDebugOutput = ShellObject.SpecialFolders("Desktop") & "\" & "debug.txt"
    

    ' データが入力されている最大カラム数、行数を得る
    Dim MaxRange As Range
    Set MaxRange = wsAddrbook.UsedRange
    Dim rowsMax As Integer
    Dim colsMax As Integer
    rowsMax = MaxRange.row + MaxRange.Rows.Count - 1
    colsMax = MaxRange.Column + MaxRange.Columns.Count - 1
    

    If DEBUG_NO_OUTLOOK Then
        ' デスクトップにログファイルを作成（既存ファイルがある場合はゼロに切り捨てる）
        fn = FreeFile
        Open filepathDebugOutput For Output As #fn
            Print #fn, "rowsMax = " & rowsMax
            Print #fn, "colsMax = " & colsMax
            Print #fn, "Ubound(arrKey) = " & UBound(arrKey)
            Print #fn, "======"
        Close #fn
    End If


    Dim objOutlook As Outlook.Application
    Set objOutlook = New Outlook.Application
    Dim objMail As Outlook.MailItem


    For i = START_ROW To rowsMax
        If wsControl.CheckBoxes("chkbox_debug5line").Value = xlOn And i >= START_ROW + DEBUG5LINE_LINE Then Exit For
        
        ' Excel画面下のステータスバーに、処理中状況表示を行う
        Application.StatusBar = (i - START_ROW + 1) & "/" & (rowsMax - START_ROW + 1) & "処理中..."
        
        '制御画面のL4セルで「発送制御する」（列名文字列が入力されている）設定となっている場合
        '(N4セルに「*」が入力されている場合は、全ての値に一致する設定のため、発送制御処理は読み飛ばす)
        Dim colControl As Integer: colControl = -1
        If wsControl.Range("L4").Value <> "" And wsControl.Range("N4").Value <> "*" Then
            ' L4セルの文字列が、名簿シートの何列目か検索
            For colControl = 0 To UBound(arrKey)
                If arrKey(colControl) = wsControl.Range("L4").Value Then Exit For
            Next colControl
            ' 名簿シートの「発送制御」セルが制御シート「N4の値」でないなら、現在処理行の残りの処理（メール送信）をスキップする
            If colControl >= 0 And colControl <= UBound(arrColNo) Then
                colControl = arrColNo(colControl)
                If wsAddrbook.Cells(i, colControl) <> wsControl.Range("N4").Value Then GoTo SKIP_TO_NEXT_ROW
            Else
                colControl = -1
            End If
        End If
        
        
        ' メールアドレス「〇〇〇〇 <xxxx@example.com>」、メール題名、メール本文を文字列置換する
        Dim strEmailFull As String
        Dim strSubject As String
        Dim strMailText As String
    
        strEmailFull = wsControl.Range("C4").Value & " <" & wsControl.Range("I4").Value & ">;"
        strEmailFull = ReplaceString(strEmailFull, arrKey, arrColNo, wsAddrbook, i)
    
        strSubject = wsControl.Range("C6").Value
        strSubject = ReplaceString(strSubject, arrKey, arrColNo, wsAddrbook, i)
    
        strMailText = wsControl.Range("C8").Value
        strMailText = ReplaceString(strMailText, arrKey, arrColNo, wsAddrbook, i)
        
    
        ' デスクトップにログファイルを作成
        If DEBUG_NO_OUTLOOK Then
            fn = FreeFile
            Open filepathDebugOutput For Append As #fn
                Print #fn, strEmailFull & ",  " & strSubject
            Close #fn
        
            ' Outlookの処理をしない（デバッグ用）
            GoTo SKIP_OUTLOOK_PROCESS
        End If
    
        ' Outlookのメールアイテム（1通分）のオブジェクトを作成
        Set objMail = objOutlook.CreateItem(olMailItem)
        With objMail
            .To = strEmailFull              ' 送信先アドレス
            .Subject = strSubject           ' 題名
            .Body = strMailText             ' メール本文
            .BodyFormat = olFormatPlain     'メールの形式
        End With
        
        ' 添付ファイル用パスワードをグローバル変数に格納
        strAttachWordPassword = wsControl.Range("D35").Value
        
        ' 添付ファイルの処理（文字列置換し、TEMPフォルダに一時格納し、添付する）
        Dim filepath As String
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim arr() As Variant, rowCell As Variant
        arr = Array(26, 28, 30, 32)     ' ファイル名が格納された C26, C28, C30, C32 セルを順に処理
        For Each rowCell In arr
            filepath = wsControl.Cells(rowCell, "C")
            If wsControl.CheckBoxes("chkbox_replace_" & rowCell).Value = xlOn Then
            'If wsControl.Cells(rowCell, "L") = "置換あり" Then
                Dim tempfilepath As String
                tempfilepath = ReplaceString_WordExcel(filepath, arrKey, arrColNo, wsAddrbook, i)
                If tempfilepath <> "" Then
                    objMail.Attachments.Add (tempfilepath)
                    fso.DeleteFile (tempfilepath)
                End If
            Else
                If filepath <> "" And Dir(filepath) <> "" Then objMail.Attachments.Add (filepath)
            End If
        Next rowCell
        Set fso = Nothing

        ' メールをOutlookの「下書き」フォルダに保存する（実際の送信はしない）
        'objMail.Send
        'objMail.Display
        objMail.Save
        Set objMail = Nothing
  
    
SKIP_OUTLOOK_PROCESS:
SKIP_TO_NEXT_ROW:
    Next i
    
    Application.StatusBar = False
    MsgBox ((rowsMax - START_ROW + 1) & "件のメールをOutlook「下書き」フォルダに格納しました")

ERRTRAP_Button_SendMail:
    If Err.Number Then
        MsgBox ("エラー(Button_SendMail) : " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "===== スクリプトを終了します =====")
        Err.Clear
        End
    End If


End Sub


'***********************************************************
' 文字列置換用に、「キーワード」と「そのキーワードが入っている列番号」を配列に格納する
'***********************************************************
Sub SetArrayColName(ByRef arrKey() As String, ByRef arrColNo() As Integer, wsAddrbook As Worksheet)
    Dim i As Integer
    
    ' データが入力されている最大カラム数を得る
    Dim MaxRange As Range
    Set MaxRange = wsAddrbook.UsedRange
    Dim colsMax As Integer
    colsMax = MaxRange.Column + MaxRange.Columns.Count - 1
    
    Erase arrKey
    Erase arrColNo
    ' 1行目のカラム名を配列に格納
    For i = 1 To colsMax
        If wsAddrbook.Cells(1, i) <> "" And Left(wsAddrbook.Cells(1, i), 1) <> "▲" Then
            If IsArrayEx(arrKey) = 0 Then
                ' 配列が空の場合は、最初のエレメントを追加
                ReDim arrKey(0)
                ReDim arrColNo(0)
            Else
                ' 配列が空でない場合は、エレメントを1個追加
                ReDim Preserve arrKey(UBound(arrKey) + 1)
                ReDim Preserve arrColNo(UBound(arrColNo) + 1)
            End If
            arrKey(UBound(arrKey)) = wsAddrbook.Cells(1, i)
            arrColNo(UBound(arrColNo)) = i
        End If
    Next i

End Sub


'***********************************************************
' 機能   : 引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数   : varArray  配列
' 戻り値 : 判定結果（1:配列/0:空の配列/-1:配列じゃない）
'***********************************************************
Private Function IsArrayEx(arr As Variant) As Long
On Error GoTo ERROR_ARRAY_EMPTY

    If IsArray(arr) Then
        IsArrayEx = IIf(UBound(arr) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_ARRAY_EMPTY:
    ' 配列が空の場合、UBound(arr)がエラーとなるのをトラップ
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function


'***********************************************************
' 文字列置換
' 引数   : str 置換元の文字列、arrKey() 置換元文字列、arrColNo() 置換先文字列の列番号、wsAddrbook アドレス帳ワークシート、row アドレス帳の対象行
' 戻り値 : 置換後の文字列
'***********************************************************
Private Function ReplaceString(str As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString = str    ' 戻り値（初期値）
    Dim i As Integer
    
    ' データが入力されている最大カラム数、行数を得る
    Dim MaxRange As Range
    Set MaxRange = wsAddrbook.UsedRange
    Dim rowsMax As Integer
    rowsMax = MaxRange.row + MaxRange.Rows.Count - 1
    
    For i = 0 To UBound(arrKey)
        ' 文字列置換
        str = Replace(str, "■" & arrKey(i) & "■", wsAddrbook.Cells(row, arrColNo(i)))
    Next i
    
    ReplaceString = str    ' 戻り値

End Function

 
'***********************************************************
' 文字列置換（Word, Excel 自動判別）
' 引数   : filepath 置換元の対象ファイル、arrKey() 置換元文字列、arrColNo() 置換先文字列の列番号、wsAddrbook アドレス帳ワークシート、row アドレス帳の対象行
' 戻り値 : 置換後のファイル名（TEMPフォルダ内のファイル フルパス）
'***********************************************************
Private Function ReplaceString_WordExcel(filepath As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString_WordExcel = ""    ' 戻り値
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 指定されたファイルが存在しない、ファイル名が "" の場合
    If filepath = "" Or Dir(filepath) = "" Then
        Set fso = Nothing
        Exit Function
    End If

    ' ユーザ一時ディレクトリ & ファイル名
    ReplaceString_WordExcel = fso.GetSpecialFolder(2) & "\" & fso.GetFileName(filepath)
    ' 添付ファイル名の文字列置換（宛先ごとに添付ファイル名を変更する場合）
    ReplaceString_WordExcel = ReplaceString(ReplaceString_WordExcel, arrKey, arrColNo, wsAddrbook, row)
    
    If fso.GetExtensionName(filepath) = "doc" Or fso.GetExtensionName(filepath) = "docx" Then
        ' Word文書内の文字列置換
        ReplaceString_WordExcel = ReplaceString_Word(filepath, ReplaceString_WordExcel, arrKey, arrColNo, wsAddrbook, row)
    ElseIf fso.GetExtensionName(filepath) = "xls" Or fso.GetExtensionName(filepath) = "xlsx" Then
        ' Excel文書内の文字列置換
        ReplaceString_WordExcel = ReplaceString_Excel(filepath, ReplaceString_WordExcel, arrKey, arrColNo, wsAddrbook, row)
    End If
    
    Set fso = Nothing
End Function

Private Function ReplaceString_Word(filepath As String, outputfilepath As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString_Word = outputfilepath     ' 戻り値

    ' エラーをトラップし、VBEデバッグダイアログを表示させずに、スクリプトを強制終了する
    If ERROR_TRAP_ENABLE = True Then On Error GoTo ERRTRAP_ReplaceString_Word
    Err.Clear

    Dim objWord As Word.Application
    Dim objWordDoc As Word.Document
    ' デスクトップに保存されているWordテンプレートファイルを開く
    Set objWord = New Word.Application
    Set objWordDoc = objWord.Documents.Open(filepath)
    ' システムが処理できるまで2秒待つ（余裕を見ている）
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' 「校閲‐編集の制限」の保護を解除する
    Dim flagProtected As Boolean: flagProtected = False
    If objWordDoc.ProtectionType <> wdNoProtection Then
        objWordDoc.Unprotect (strAttachWordPassword)
        flagProtected = True    ' 「名前を付けて保存」時に保護有効化するかどうかのために参照
    End If
    
    
    Dim i As Integer
    For i = 0 To UBound(arrKey)
        ' 本文中の文字列置換
        With objWordDoc.Content.Find
            .Text = "■" & arrKey(i) & "■"
            .Forward = True
            .Replacement.Text = wsAddrbook.Cells(row, arrColNo(i))
            .Wrap = wdFindContinue
            .MatchFuzzy = True
            .Execute Replace:=wdReplaceAll
        End With
        ' 本文中のハイパーリンク（メール）のSubject内文字列置換
        Dim hyp As Object
        For Each hyp In objWordDoc.Hyperlinks
            If hyp.Address Like "mailto*" Then
                hyp.EmailSubject = Replace(hyp.EmailSubject, "■" & arrKey(i) & "■", wsAddrbook.Cells(row, arrColNo(i)))
            End If
        Next hyp
        ' 図形テキストボックス内の文字列置換
        Dim shp As Word.Shape
        For Each shp In objWordDoc.Shapes
            If shp.Type = msoTextBox Then
                With shp.TextFrame.TextRange.Find
                    .ClearFormatting
                    .Forward = True
                    .Text = "■" & arrKey(i) & "■"
                    .Replacement.Text = wsAddrbook.Cells(row, arrColNo(i))
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        Next shp
        Set shp = Nothing
    Next i
    
    ' 元ファイルが保護有効だった場合、名前を付けて保存する場合も保護有効化する
    If flagProtected = True Then
        objWordDoc.Protect Type:=wdAllowOnlyReading, Password:=strAttachWordPassword
    End If
    
    ' Wordドキュメントを名前を付けて保存する
    objWordDoc.SaveAs (outputfilepath)
    objWordDoc.Close SaveChanges:=False
    objWord.Quit
    
ERRTRAP_ReplaceString_Word:
    If Err.Number Then
        MsgBox ("エラー(ReplaceString_Word) : " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "===== スクリプトを終了します =====")
        Err.Clear
        End
    End If
    
    Set objWordDoc = Nothing
    Set objWord = Nothing

End Function


Private Function ReplaceString_Excel(filepath As String, outputfilepath As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString_Excel = outputfilepath     ' 戻り値

    ' エラーをトラップし、VBEデバッグダイアログを表示させずに、スクリプトを強制終了する
    If ERROR_TRAP_ENABLE = True Then On Error GoTo ERRTRAP_ReplaceString_Excel
    Err.Clear
    
    ' 画面に表示させないため、新しいExcelのオブジェクトを作成し、その中でExcelブックを開く
    Dim objExcel As Excel.Application
    Set objExcel = New Excel.Application
    ' 検索対象文字列が存在しなかった場合の、Warningダイアログの出現を抑止する
    objExcel.DisplayAlerts = False
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = objExcel.Workbooks.Open(filepath, ReadOnly:=True)
    ' システムが処理できるまで2秒待つ（余裕を見ている）
    Application.Wait (Now + TimeValue("0:00:02"))


    For Each ws In wb.Worksheets
    
        Dim i As Integer
        For i = 0 To UBound(arrKey)
            ' 本文中の文字列置換
'            ws.Cells.Replace What:="■" & arrKey(i) & "■", Replacement:=wsAddrbook.Cells(row, arrColNo(i)), LookAt:=xlPart, _
'                    SearchOrder:=xlByRows, MatchCase:=True, MatchByte:=True, SearchFormat:=False, ReplaceFormat:=False
            ws.UsedRange.Replace What:="■" & arrKey(i) & "■", Replacement:=wsAddrbook.Cells(row, arrColNo(i)), LookAt:=xlPart
'            Dim EachCell As Range
'            For Each EachCell In ws.UsedRange
'                EachCell = Replace(EachCell, "■" & arrKey(i) & "■", wsAddrbook.Cells(row, arrColNo(i)))
'            Next EachCell
            
        Next i
    
    Next ws
    On Error GoTo 0

    wb.SaveAs (outputfilepath)
    wb.Close SaveChanges:=False
    objExcel.Quit
    
ERRTRAP_ReplaceString_Excel:
    If Err.Number Then
        MsgBox ("エラー(ReplaceString_Excel) : " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "===== スクリプトを終了します =====")
        Err.Clear
        End
    End If

    Set wb = Nothing
    objExcel.DisplayAlerts = True   ' Warningダイアログ抑止設定を元に戻しておく
    Set objExcel = Nothing

End Function

