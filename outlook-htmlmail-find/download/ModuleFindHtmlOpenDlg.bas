Attribute VB_Name = "ModuleFindHtmlOpenDlg"
' *******************
'   OutlookHtmlFind : ModuleFindHtmlOpen.bas ver 1.2
'
'   Outlook HTMLメール発見ツール VBA メインダイアログのコード
'
'   メニューダイアログを表示する
'   version 1.2 for Microsoft Outlook 2000 Japanese Edition
'
'   (C) 2001-2003 INOUE. Hirokazu , All rights reserved
'   http://inoue-h.connect.to/
'  このプログラム／スクリプトはフリーウエアーです
'  このプログラム／スクリプトに対する動作・非動作の保証、実行結果の保証はありません
' *******************
Option Explicit

Sub FindHtmlOpenDlg()
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
    For i = 1 To DlgFindHtml.CmbboxFolder.ListCount
        ' 全てのアイテムを消去する
        DlgFindHtml.CmbboxFolder.RemoveItem (0)
    Next i
    DlgFindHtml.CmbboxFolder.AddItem (">> 未選択")
    For i = 1 To myNamespace.Folders.count
        tmpStr = myNamespace.Folders.Item(i)
        DlgFindHtml.CmbboxFolder.AddItem (tmpStr)
    Next i
    DlgFindHtml.CmbboxFolder.ListIndex = 0   ' 一つ目の項目を表示
    
    ' ******************************
    ' トレイ１、トレイ２ の選択肢をコンボボックスに設定
    ' ******************************
    For i = 1 To DlgFindHtml.CmbboxTray1.ListCount
        ' 全てのアイテムを消去する
        DlgFindHtml.CmbboxTray1.RemoveItem (0)
    Next i
    DlgFindHtml.CmbboxTray1.AddItem (">> 未選択")
    DlgFindHtml.CmbboxTray1.ListIndex = 0   ' 一つ目の項目を表示
    For i = 1 To DlgFindHtml.CmbboxTray2.ListCount
        ' 全てのアイテムを消去する
        DlgFindHtml.CmbboxTray2.RemoveItem (0)
    Next i
    DlgFindHtml.CmbboxTray2.AddItem (">> 未選択")
    DlgFindHtml.CmbboxTray2.ListIndex = 0   ' 一つ目の項目を表示
    
    
    
    ' フォームの表示
    DlgFindHtml.Show
        
End Sub

' ファイル終了 EOF
' ***********************
