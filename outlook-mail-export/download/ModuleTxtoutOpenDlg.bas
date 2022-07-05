Attribute VB_Name = "ModuleTxtoutOpenDlg"
' *******************
'   Outlook メール テキスト化 VBA ( マクロ呼び出し ModuleTxtoutOpenDlgn )
'   Version 1.3
'   (C) 2001-2022 INOUE. Hirokazu
'
'   このVBAスクリプトは GNU General Public License v3ライセンスで公開する フリーソフトウエア
' *******************
Option Explicit

Sub TxtoutOpenDlg()
    
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim tmpStr As String        ' テンポラリ文字列
    ' Outlook 自体
    Dim myNamespace As NameSpace
    ' フォルダオブジェクト
    Dim OlkEmailFolder As MAPIFolder
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    
    Set myNamespace = Application.GetNamespace("MAPI")
    
    ' ******************************
    ' フォルダ（最上階）の選択肢をコンボボックスに設定
    ' ******************************
    For i = 1 To DlgTxtOutMain.CmbboxFolder.ListCount
        ' 全てのアイテムを消去する
        DlgTxtOutMain.CmbboxFolder.RemoveItem (0)
    Next i
    DlgTxtOutMain.CmbboxFolder.AddItem (">> 未選択")
    For i = 1 To myNamespace.Folders.count
        tmpStr = myNamespace.Folders.Item(i)
        DlgTxtOutMain.CmbboxFolder.AddItem (tmpStr)
    Next i
    DlgTxtOutMain.CmbboxFolder.ListIndex = 0   ' 一つ目の項目を表示
    
    ' ******************************
    ' トレイ１、トレイ２ の選択肢をコンボボックスに設定
    ' ******************************
    For i = 1 To DlgTxtOutMain.CmbboxTray1.ListCount
        ' 全てのアイテムを消去する
        DlgTxtOutMain.CmbboxTray1.RemoveItem (0)
    Next i
    DlgTxtOutMain.CmbboxTray1.AddItem (">> 未選択")
    DlgTxtOutMain.CmbboxTray1.ListIndex = 0   ' 一つ目の項目を表示
    For i = 1 To DlgTxtOutMain.CmbboxTray2.ListCount
        ' 全てのアイテムを消去する
        DlgTxtOutMain.CmbboxTray2.RemoveItem (0)
    Next i
    DlgTxtOutMain.CmbboxTray2.AddItem (">> 未選択")
    DlgTxtOutMain.CmbboxTray2.ListIndex = 0   ' 一つ目の項目を表示
    
    ' ******************************
    ' チェックボックスの設定
    ' ******************************
    DlgTxtOutMain.ChkSort.Value = True
    DlgTxtOutMain.ChkSortRev.Value = True
    DlgTxtOutMain.ChkIndxSentMail.Value = False
    DlgTxtOutMain.ChkUnicode.Value = True
    
    
    ' フォームの表示
    DlgTxtOutMain.Show
    
    ' ダイアログ項目の表示・入力値などをクリアするために初期化する
    Set DlgTxtOutMain = Nothing
        
End Sub



' ファイル終了 EOF
' ***********************
