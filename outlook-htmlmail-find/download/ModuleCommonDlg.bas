Attribute VB_Name = "ModuleCommonDlg"
' ***********************
'   OutlookEmailText : ModuleCommonDialog.bas ver 1.1
'   OutlookHtmlFind : ModuleCommonDialog.bas ver 1.2
'
'   （上記2つのプログラムで使用されています）
'
' SDK 関数
' Windows のコモンダイアログ 「ファイルを開く」、「ファイルを保存する」
' ***********************
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
(pOpenfilename As OPENFILENAME) As Long

Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
(pOpenfilename As OPENFILENAME) As Long

' ***********************
' OPENFILENAME 構造体
' ***********************
Public Type OPENFILENAME
    lStructSize As Long             'この構造体の長さ
    hwndOwner As Long               '呼び出し元ウインドウハンドル
    hInstance As Long
    lpstrFilter As String           'フィルタ文字列
    lpstrCustomFilter As String
    nMaxCustrFilter As Long
    nFilterIndex As Long
    lpstrFile As String             '選択されたファイル名（フルパス）
    nMaxFile As Long                'lpstrFileのバッファサイズ
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String       '初期フォルダ名
    lpstrTitle As String            'コモンダイアログのタイトル名
    flags As Long                   'フラグ
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String           'ファイル名の入力時、拡張子が省略された時の拡張子
    lCustrData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' ファイル終了 EOF
' ***********************
