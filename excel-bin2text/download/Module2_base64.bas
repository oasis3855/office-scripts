Attribute VB_Name = "Module2"
Option Explicit
' *****
' Base64関連：ボタン押し下げハンドラサブルーチン、変換関数
' *****


Sub BtnBase64Encode()
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
    ' 指定されたファイルを読み込み、Base64にエンコードし、string変数として返す
    ' *****
    Dim strBase64 As String
    strBase64 = EncodeBase64(inputFilename)
    
    
    ' *****
    ' string変数（Base64にエンコードした文字列）をテキストファイルに書き込む
    ' *****
    Dim fn As Variant
    fn = FreeFile
    Open CStr(outputFilename) For Output As #fn
    
    Print #fn, strBase64
    
    Close (fn)

    MsgBox ("処理終了" & vbCrLf & "変換元 : " & inputFilename & vbCrLf & "変換先 : " & outputFilename)


End Sub

Sub BtnBase64Decode()
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
    ' ソースファイル（Base64でエンコードされたファイル）を、一気にString変数に読み込む
    ' *****
    Dim strBase64 As String
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(inputFilename).OpenAsTextStream
            strBase64 = .ReadAll
            .Close
        End With
    End With
    
    ' *****
    ' Base64でエンコードされた文字列を復号化し、ファイルに書き込む
    ' *****
    Dim retCode As Long
    retCode = DecodeBase64(strBase64, outputFilename)

    MsgBox ("処理終了" & vbCrLf & "変換元 : " & inputFilename & vbCrLf & "変換先 : " & outputFilename)

End Sub


' *****
' 指定したファイルから読み込んだデータを、Base64変換した文字列を返す
' 引数 :
'   inputFilename : 読み込み対象のファイル名
' 戻り値 :
'   Base64変換したテキスト文字列
' 著作権
'   この関数は、https://www.ka-net.org/blog/?p=4479 より転載した
' *****
Private Function EncodeBase64(ByVal inputFilename As String) As String
    'ファイルをBase64エンコード
    Dim elm As Object
    Dim ret As String
    Const adTypeBinary = 1
    Const adReadAll = -1
    
    ret = ""    ' 戻り文字列の初期化
    On Error Resume Next
    Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .LoadFromFile inputFilename
        elm.DataType = "bin.base64"
        elm.nodeTypedValue = .read(adReadAll)
        ret = elm.Text
        .Close
    End With
    On Error GoTo 0
    
    EncodeBase64 = ret
End Function


' *****
' 与えられた文字列をBase64から復号化し、指定したファイルに書き込む
' 引数 :
'   strBase64       : Base64でエンコードされた文字列（これが復号化対象となる）
'   outputFilename  : 復号化したデータを書き込むファイル名
' 戻り値 :
'   Long値
' 著作権
'   この関数は、https://www.ka-net.org/blog/?p=4479 より転載した
' *****
Private Function DecodeBase64(ByVal strBase64 As String, ByVal outputFilename As String) As Long
    Dim elm As Object
    Dim ret As Long
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    
    Dim strVar As String
    
    ret = -1    ' 戻り値初期化
    On Error Resume Next
    Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
    elm.DataType = "bin.base64"
    elm.Text = strBase64
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        ' 変換結果は elm.nodeTypedValue 配列(Variant)に格納されている
        .Write elm.nodeTypedValue
        .SaveToFile outputFilename, adSaveCreateOverWrite
        .Close
    End With
    If Err.Number <> 0 Then ret = 0
    On Error GoTo 0
    DecodeBase64 = ret
End Function
