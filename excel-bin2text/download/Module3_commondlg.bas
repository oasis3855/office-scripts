Attribute VB_Name = "Module3"
Option Explicit
' *****
' 共通関数（ファイル選択ダイアログ）
'
'  Function FileOpenDlg(ByVal strTitleMsg As String, ByRef arrFilter() As Variant) As String
'  Function FileSaveDlg(ByVal strTitleMsg As String, ByVal strFilter As String, ByVal boolExistDelete As Boolean) As String
' *****


' *****
' 読み込みファイル名を選択するダイアログを表示。ファイル存在チエック付き
' 引数 :
'   strTitileMsg : ダイアログのウインドウタイトル文字列
'   arrFilter()  : ファイル種別選択フィルター
' 戻り値 :
'   正常選択 : ファイルのフルパス文字列
'   キャンセルボタン or ファイル不存在 : "" 文字列
' *****
Function FileOpenDlg(ByVal strTitleMsg As String, ByRef arrFilter() As Variant) As String
    FileOpenDlg = ""        ' 戻り値の初期値設定

    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    ' デスクトップフォルダを初期値とする。末尾を \ で終わることで、ディレクトリ明示
    DesktopFolderName = ShellObject.SpecialFolders("Desktop") & "\"
    Set ShellObject = Nothing

    ' *****
    ' 読み込みファイル名を選択・入力するGUIオブジェクトを構築
    ' *****
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = strTitleMsg
        With .Filters
            ' **
            ' フィルターの設定。書式は次のようなもので、functionの引数より設定する
            ' .Add "テキストファイル", "*.txt", 1
            ' **
            .Clear
            Dim numItems As Integer
            numItems = (UBound(arrFilter) + 1) / 2
            Dim i As Integer
            For i = 0 To numItems - 1
                .Add arrFilter(i * 2), arrFilter(i * 2 + 1), i + 1
            Next i
        End With
        .InitialFileName = DesktopFolderName
    End With

    Set ShellObject = Nothing

    ' *****
    ' 読み込みファイル名を選択・入力するGUIダイアログを表示
    ' *****
    If fd.Show <> True Then
        MsgBox ("キャンセルボタンが押されました")
        Set fd = Nothing
        Exit Function
    End If
    
    FileOpenDlg = fd.SelectedItems(1)       ' 選択されたファイルのフルパス名
    Set fd = Nothing
    
    ' *****
    ' 入力ファイルの存在を確認
    ' *****
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FileOpenDlg) <> True Then
        MsgBox ("入力ファイルが見つかりません" & vbCrLf & FileOpenDlg)
        Set fso = Nothing
        Exit Function
    End If
    Set fso = Nothing
 

End Function


' *****
' 書き込みファイル名を選択するダイアログを表示。ファイル存在チエック付き
' 引数 :
'   strTitileMsg : ダイアログのウインドウタイトル文字列
'   strFilter    : ファイル種別選択フィルター（例 : "テキストファイル,*.txt,全てのファイル,*.*"）
'   boolExistDelete  : Trueを指定するとファイルが存在する場合は、削除する
' 戻り値 :
'   正常選択 : ファイルのフルパス文字列
'   キャンセルボタン : "" 文字列
' *****

Function FileSaveDlg(ByVal strTitleMsg As String, ByVal strFilter As String, ByVal boolExistDelete As Boolean) As String
    FileSaveDlg = ""        ' 戻り値の初期値設定
    
    Dim outputFilename As Variant

    ' *****
    ' 書き込みファイル名を選択・入力するGUI表示
    ' *****
    outputFilename = Application.GetSaveAsFilename(FileFilter:=strFilter)
    ' OKが押された場合はファイル名のStringが、キャンセルの場合は FALSE が返る
    If VarType(outputFilename) = vbBoolean Then
        MsgBox ("キャンセルボタンが押されました")
        Exit Function
    End If
    
    ' *****
    ' 出力ファイルが存在する場合は、削除する
    ' *****
    If boolExistDelete = True Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(outputFilename) Then
            If MsgBox("出力ファイルが存在しますので削除します" & vbCrLf & outputFilename, vbYesNo) = vbYes Then
                Kill (outputFilename)
            Else
                MsgBox ("既存ファイルを削除せずに、この先の処理は出来ません")
                Set fso = Nothing
                Exit Function
            End If
        End If
        Set fso = Nothing
    End If
    
    ' 戻り値（選択されたファイル名）
    FileSaveDlg = outputFilename

End Function

