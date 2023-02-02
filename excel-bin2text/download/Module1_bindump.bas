Attribute VB_Name = "Module1"
Option Explicit
' *****
' バイナリダンプ関連：ボタン押し下げハンドラサブルーチン
' *****

Sub BtnBinToText()

    Dim inputFilename As String
    Dim outputFilename As Variant

    Dim arrFilter() As Variant
    arrFilter = Array("全てのファイル", "*.*", "テキストファイル", "*.txt", "zipファイル", "*.zip")
    inputFilename = FileOpenDlg("入力ファイル名の指定", arrFilter)
    If inputFilename = "" Then
        Exit Sub
    End If

    outputFilename = FileSaveDlg("出力（保存）ファイル名の指定", "テキストファイル,*.txt,全てのファイル,*.*", True)
    If outputFilename = "" Then
        Exit Sub
    End If
    
    
    ' *****
    ' ソースファイルを全て読み込み、バイナリ配列に格納する
    ' *****
    Dim fn
    fn = FreeFile
    Dim numFileSize As Long
    Dim byteArray() As Byte
    Open inputFilename For Binary Access Read As #fn
    
    numFileSize = LOF(fn)
    ReDim byteArray(numFileSize)
 
    byteArray = InputB(numFileSize, fn)
    
    Close (fn)
    

    ' *****
    ' バイナリ配列を、出力ファイル（テキストファイル）に書き込む
    ' *****
    Dim i As Long
    Dim ch As Byte
    Dim strBuffer As String
    fn = FreeFile   ' 新しいファイル番号を得る
    Open CStr(outputFilename) For Output As #fn
    
    For i = 0 To numFileSize - 1
        If i Mod 16 = 0 Then
            ' 出力行頭は、ファイル先頭からのアドレス 6桁
            strBuffer = Right("000000" & Hex(i), 6)
            Print #fn, strBuffer & " | ";       ' スクリプト行末 セミコロンで改行しない出力
        End If
        ' バイト配列から取り出されたデータを16進文字列化してテキストファイルに書き込む
        ch = byteArray(i)
        strBuffer = Right("0" & Hex(ch), 2)     ' ゼロ埋め2桁にする
        Print #fn, strBuffer & " ";             ' スクリプト行末 セミコロンで改行しない出力
        If (i + 1) Mod 16 = 0 Then
            Print #fn, vbCrLf;
        End If
    Next i
    
    Close (fn)
    Set fn = Nothing
    
    MsgBox ("処理終了" & vbCrLf & "変換元 : " & inputFilename & vbCrLf & "変換先 : " & outputFilename)
    
    
End Sub


Sub BtnTextToBin()
    Dim inputFilename As String
    Dim outputFilename As Variant

    Dim arrFilter() As Variant
    arrFilter = Array("テキストファイル", "*.txt", "全てのファイル", "*.*")
    inputFilename = FileOpenDlg("入力ファイル名の指定", arrFilter)
    If inputFilename = "" Then
        Exit Sub
    End If

    outputFilename = FileSaveDlg("出力（保存）ファイル名の指定", "全てのファイル,*.*,テキストファイル,*.txt,zipファイル,*.zip", True)
    If outputFilename = "" Then
        Exit Sub
    End If



    ' *****
    ' 入力・出力ファイルを開く
    ' *****
    Dim fnIn
    Dim fnOut
    fnIn = FreeFile
    Open inputFilename For Input As #fnIn
    fnOut = FreeFile
    Open outputFilename For Binary Access Write As #fnOut
    
    ' *****
    ' ソースファイルを1行ずつ読み込みながら、バイナリに戻し出力ファイルに書き込む
    ' *****
    Dim strLine As String
    Dim chStr As Variant
    Dim chByte As Byte
    Dim arrayData As Variant
    Do Until EOF(fnIn)
        Line Input #fnIn, strLine           ' ソースファイルのテキスト1行読み込み
        arrayData = Split(strLine, "|")     ' アドレス部分とデータ部分を切り分け
        arrayData = Split(CStr(arrayData(1)), " ")    ' データを1バイト分ずつ切り分け
        For Each chStr In arrayData
            If Len(chStr) >= 1 Then         ' 1文字以上であれば（NULL文字列をスキップ）
                chByte = Val("&H" & chStr)
                Put #fnOut, , chByte
            End If
        Next
    Loop
    
    Close (fnOut)
    Close (fnIn)
    
    MsgBox ("処理終了" & vbCrLf & "変換元 : " & inputFilename & vbCrLf & "変換先 : " & outputFilename)

End Sub



