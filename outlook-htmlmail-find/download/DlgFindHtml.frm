VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgFindHtml 
   Caption         =   "HTMLメール 発見ツール"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   OleObjectBlob   =   "DlgFindHtml.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DlgFindHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************
'   OutlookHtmlFind : DlgFindHtml.frm ver 1.2
'
'   Outlook HTMLメール発見ツール VBA メインダイアログのコード
'   version 1.2 for Microsoft Outlook 2000 Japanese Edition
'
'   (C) 2001-2003 INOUE. Hirokazu , All rights reserved
'   http://inoue-h.connect.to/
'  このプログラム／スクリプトはフリーウエアーです
'  このプログラム／スクリプトに対する動作・非動作の保証、実行結果の保証はありません
'
'
' ● 重要 ● Outlookの「ツール｣-「マクロ｣-「セキュリティ｣メニューの設定が、「中｣以下で無いと実行できない。
'
' *******************
Option Explicit

' ******************************
' ソート数の最大を指定します。大きくすると、メモリを食います
' ******************************
Const max_a = 2000  ' 日付ソート配列の最大値

Private Sub BtnExec_Click()
' ******************************
' 実行ボタンを押したとき
' 「テキスト化ツール｣を流用
' ******************************
    On Error GoTo BtnExec_ErrHandler
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim j As Integer            ' カウンタ用変数
    Dim tmpStr As String        ' テンポラリ文字列
    ' Outlook 自体
    Dim myNamespace As NameSpace
    ' フォルダオブジェクト
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    Dim OlkEmailEnt As MAPIFolder
    Dim OlkEmailItem As MailItem    ' MailItem で明示的に宣言
    
    ' 日付ソート用配列
    ReDim a_indx(max_a) As Integer ' インデックスの配列
    ReDim a_date(max_a) As Date    ' 日付データの配列
    
    ' 出力テキスト
    Dim OutputStr As String
    
    
    Set myNamespace = Application.GetNamespace("MAPI")

    If (CmbboxFolder.Value = ">> 未選択") Or (CmbboxTray1.Value = ">> 未選択") Then
        i = MsgBox("フォルダ および トレイ１ を選択する必要があります", vbOKOnly + vbExclamation, "Outlook メール テキスト化 VBA エラー")
        Exit Sub
    End If
    
    
    Set OlkEmailTray1 = myNamespace.Folders(CmbboxFolder.Value)
    Set OlkEmailTray2 = OlkEmailTray1.Folders(CmbboxTray1.Value)
    If CmbboxTray2.Value = ">> 未選択" Then
        Set OlkEmailEnt = OlkEmailTray2
    Else
        Set OlkEmailEnt = OlkEmailTray2.Folders(CmbboxTray2.Value)
    End If
    
    ' 日付データによるソーティングを行う
    If (OlkEmailEnt.Items.count < max_a) Then
        For i = 1 To OlkEmailEnt.Items.count
            a_indx(i) = i
            Set OlkEmailItem = OlkEmailEnt.Items(i)
            a_date(i) = OlkEmailItem.SentOn
        Next i
        i = Sort_By_Date(a_indx, a_date, max_a, OlkEmailEnt.Items.count)
    End If
                
    OutputStr = "発見されたHTMLメールの「題名｣、発信者名、発信日は次のとおりです" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA)
    j = 0       ' HTMLメールの数をカウントする
                
    For i = 1 To OlkEmailEnt.Items.count
        ' 以後のメンバ変数の参照のために、明示的なオブジェクトに代入
        If (OlkEmailEnt.Items.count < max_a) Then
            Set OlkEmailItem = OlkEmailEnt.Items(a_indx(i))
        Else
            Set OlkEmailItem = OlkEmailEnt.Items(i)
        End If

        
    If OlkEmailItem.HTMLBody <> "" Then
        j = j + 1
        OutputStr = OutputStr + "「" + OlkEmailItem.Subject + " 」" + OlkEmailItem.SenderName + "  on " + Format(OlkEmailItem.SentOn, "yy/mm/dd hh:mm:ss") + Chr(&HD) + Chr(&HA)
    End If
        
    Next i
    
    ' 結果ダイアログを出す
    If j = 0 Then
        i = MsgBox("HTMLメールは発見されませんでした", vbOKOnly + vbInformation, "HTMLメール 発見ツール VBA")
    Else
        OutputStr = OutputStr + "以上、合計 " + Str(j) + " 通のメールが発見されました。" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + "これらのメールをテキストファイルに出力しますか？"
        
        i = MsgBox(OutputStr, vbYesNo + vbExclamation, "HTMLメール 発見ツール VBA")
        If i = vbYes Then
            ' テキストファイル化するサブルーチンへ
            Call OutputTextFile(CmbboxFolder.Text, CmbboxTray1.Text, CmbboxTray2.Text)
        End If
    End If
    
    Exit Sub
BtnExec_ErrHandler:
    i = MsgBox("エラーが発生しました。処理を中止します。", vbOKOnly + vbExclamation, "Outlook メール テキスト化 VBA 致命的エラー")
    Exit Sub

End Sub

Private Sub CmbboxFolder_Change()
' ******************************
' フォルダ項目が新たに選択されたとき
' ******************************
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim tmpStr As String        ' テンポラリ文字列
    ' Outlook 自体
    Dim myNamespace As NameSpace
    ' フォルダオブジェクト
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    
    Set myNamespace = Application.GetNamespace("MAPI")

    ' 未選択を選択した場合
    If CmbboxFolder.Value = ">> 未選択" Then
        For i = 1 To CmbboxTray1.ListCount
            ' 全てのアイテムを消去する
            CmbboxTray1.RemoveItem (0)
        Next i
        CmbboxTray1.AddItem (">> 未選択")
        CmbboxTray1.ListIndex = 0   ' 一つ目の項目を表示
        Exit Sub
    End If
    
    Set OlkEmailTray1 = myNamespace.Folders(CmbboxFolder.Value)
    
    For i = 1 To CmbboxTray1.ListCount
        ' 全てのアイテムを消去する
        CmbboxTray1.RemoveItem (0)
    Next i
    CmbboxTray1.AddItem (">> 未選択")
    For i = 1 To OlkEmailTray1.Folders.count
        tmpStr = OlkEmailTray1.Folders.Item(i)
        CmbboxTray1.AddItem (tmpStr)
    Next i
    CmbboxTray1.ListIndex = 0   ' 一つ目の項目を表示
End Sub

Private Sub CmbboxTray1_Change()
' ******************************
' トレイ１項目が新たに選択されたとき
' ******************************
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim tmpStr As String        ' テンポラリ文字列
    ' Outlook 自体
    Dim myNamespace As NameSpace
    ' フォルダオブジェクト
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    
    Set myNamespace = Application.GetNamespace("MAPI")

    ' 未選択を選択した場合
    If CmbboxTray1.Value = ">> 未選択" Then
        For i = 1 To CmbboxTray2.ListCount
            ' 全てのアイテムを消去する
            CmbboxTray2.RemoveItem (0)
        Next i
        CmbboxTray2.AddItem (">> 未選択")
        CmbboxTray2.ListIndex = 0   ' 一つ目の項目を表示
        Exit Sub
    End If
    
    Set OlkEmailTray1 = myNamespace.Folders(CmbboxFolder.Value)
    Set OlkEmailTray2 = OlkEmailTray1.Folders(CmbboxTray1.Value)
    
    For i = 1 To CmbboxTray2.ListCount
        ' 全てのアイテムを消去する
        CmbboxTray2.RemoveItem (0)
    Next i
    CmbboxTray2.AddItem (">> 未選択")
    For i = 1 To OlkEmailTray2.Folders.count
        tmpStr = OlkEmailTray2.Folders.Item(i)
        CmbboxTray2.AddItem (tmpStr)
    Next i
    CmbboxTray2.ListIndex = 0   ' 一つ目の項目を表示

End Sub

Private Sub OutputTextFile(tray0 As String, tray1 As String, tray2 As String)
' ******************************
' HTMLメールのみをテキストファイルに書き出すサブルーチン
' 「テキスト化ツール｣を流用
' ******************************
    
    Dim strFname As String      ' 出力ファイル名
    Dim tmpStr As String        ' テンポラリ文字列
    ' Windows のコモンダイアログ 用の構造体
    Dim strOfn As OPENFILENAME
    Dim nLRet As Long, nNULLPos As Long
    
    With strOfn
        .lStructSize = Len(strOfn)
        .lpstrInitialDir = ""      '（最初に表示するディレクトリ）
                                            '（フィルターでファイル種類を絞る）
        .lpstrFilter = "テキストファイル(*.txt)" & vbNullChar & "*.txt" _
        & vbNullChar & "全てのファイル(*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .nMaxFile = 256                        '（ファイル名の最大長（パス含む））
        .lpstrFile = String(256, vbNullChar)   '（ファイル名を格納する文字列
                                                ' NULLで埋めておく）
        .lpstrTitle = "データを書き込むテキストファイル名の指定"
    End With
    
    ' 「ファイルを保存する」ダイアログ
    nLRet = GetSaveFileName(strOfn)
    
    nNULLPos = InStr(strOfn.lpstrFile, vbNullChar)  'ファイル名の終り(NULLの位置)を調べる
    tmpStr = Left(strOfn.lpstrFile, nNULLPos - 1) 'ファイル名の有効部分を取り出す
    
    ' キャンセルボタンが押された
    If nLRet = False Then
        MsgBox ("出力を中止します")
        Exit Sub
    End If
    ' 拡張子 .txt が指定されなかった場合 .txt を付ける
    If InStr(tmpStr, ".txt") < 1 Then
        ' 拡張子フィルタが .txt のとき
        If strOfn.nFilterIndex = 1 Then
            tmpStr = tmpStr + ".txt"
        End If
    End If
    
    ' 出力ファイル名を確定
    strFname = tmpStr
    
    On Error GoTo BtnExec_ErrHandler
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim j As Integer            ' カウンタ用変数
    ' Outlook 自体
    Dim myNamespace As NameSpace
    ' フォルダオブジェクト
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    Dim OlkEmailEnt As MAPIFolder
    Dim OlkEmailItem As MailItem    ' MailItem で明示的に宣言
    
    ' 日付ソート用配列
    ReDim a_indx(max_a) As Integer ' インデックスの配列
    ReDim a_date(max_a) As Date    ' 日付データの配列
    
    ' 出力テキストファイルオブジェクト
    Dim fs                      ' FileSystemObject
    Dim fi_out                  ' TextStream
    Dim FileName As String      ' ファイル名
    
    Set myNamespace = Application.GetNamespace("MAPI")

    Set OlkEmailTray1 = myNamespace.Folders(tray0)
    Set OlkEmailTray2 = OlkEmailTray1.Folders(tray1)
    If tray2 = ">> 未選択" Then
        Set OlkEmailEnt = OlkEmailTray2
    Else
        Set OlkEmailEnt = OlkEmailTray2.Folders(tray2)
    End If
    
    ' 日付データによるソーティングを行う
    If OlkEmailEnt.Items.count < max_a Then
        For i = 1 To OlkEmailEnt.Items.count
            a_indx(i) = i
            Set OlkEmailItem = OlkEmailEnt.Items(i)
            a_date(i) = OlkEmailItem.SentOn
        Next i
        i = Sort_By_Date(a_indx, a_date, max_a, OlkEmailEnt.Items.count)
    End If
            

    ' ファイルシステムのオブジェクトを得る
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strFname) = True Then
        If vbNo = MsgBox("指定されたファイルはすでに存在します。上書きしますか ？" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + "  「いいえ｣を選択すると、既存ファイルに追加するモードとなります", vbYesNo + vbQuestion, "Outlook メール テキスト化 VBA 確認") Then
            ' 既存ファイルに追加
            Set fi_out = fs.OpenTextFile(strFname, 8, True)
        Else
            ' 上書き
            Set fi_out = fs.CreateTextFile(strFname, True)
        End If
    Else
        ' テキストファイルを新規作成しオープン
        Set fi_out = fs.CreateTextFile(strFname, True)
    End If
    ' ヘッダ情報を書く
    tmpStr = "Outlook メール テキスト化 VBA " + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    tmpStr = "X#########################################################################" + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    
    j = 0
    For i = 1 To OlkEmailEnt.Items.count
        ' 以後のメンバ変数の参照のために、明示的なオブジェクトに代入
        If OlkEmailEnt.Items.count < max_a Then
            Set OlkEmailItem = OlkEmailEnt.Items(a_indx(i))
        Else
            Set OlkEmailItem = OlkEmailEnt.Items(i)
        End If

        If OlkEmailItem.HTMLBody <> "" Then
            j = j + 1
            tmpStr = "題名 ： 「" + OlkEmailItem.Subject + " 」" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "発信者 ： " + OlkEmailItem.SentOnBehalfOfName + " / " + OlkEmailItem.SenderName + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "ReplyTo ： " + OlkEmailItem.ReplyRecipientNames + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "宛先 ： " + OlkEmailItem.To + "  CC ： " + OlkEmailItem.CC + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "BCC ： " + OlkEmailItem.BCC + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "送信日 ： " + Format(OlkEmailItem.SentOn, "yy/mm/dd hh:mm:ss") + "  受信日 ： " + Format(OlkEmailItem.ReceivedTime, "yy/mm/dd hh/mm/ss") + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "本文 ： " + Chr(&HD) + Chr(&HA) + OlkEmailItem.Body + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "X#########################################################################" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "HTMLの本文 ： " + Chr(&HD) + Chr(&HA) + OlkEmailItem.HTMLBody + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "X#########################################################################" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "##########################################################################" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            
        End If
        
    Next i
    
    ' 処理件数を書き込む
    tmpStr = Chr(&HD) + Chr(&HA) + "件数 : " + Str(j) + "     処理終了" + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    tmpStr = "E#########################################################################" + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    'ファイルをクローズ
    fi_out.Close
    
    tmpStr = "テキスト化したメールデータを " + Str(j) + " 件書き込みました"
    i = MsgBox(tmpStr, vbOKOnly + vbInformation, "Outlook メール テキスト化 VBA")
    
    Exit Sub
BtnExec_ErrHandler:
    i = MsgBox("テキストファイル書き込み中にエラーが発生しました。処理を中止します。", vbOKOnly + vbExclamation, "Outlook メール テキスト化 VBA 致命的エラー")
    Exit Sub

End Sub

Function Sort_By_Date(a_indx() As Integer, a_date() As Date, max_a As Integer, count As Integer)
' ******************************
' ソーティング （最遅モード）
' もっとスマートにしたいなら、書き換えてください。（他のアルゴリズムに）
' でも、Pentium III とか、速いCPUではほとんど体感できないはずですが…
' ******************************
    Dim i As Integer
    Dim j As Integer
    Dim tmp_indx As Integer
    Dim tmp_date As Date
    
    ' 昇順か降順かによって切り替える
'    If DlgTxtOutMain.ChkSortRev.Value = True Then
        For i = 1 To count - 1
            For j = i + 1 To count
                If a_date(i) > a_date(j) Then
                    tmp_indx = a_indx(i)
                    tmp_date = a_date(i)
                    a_indx(i) = a_indx(j)
                    a_date(i) = a_date(j)
                    a_indx(j) = tmp_indx
                    a_date(j) = tmp_date
                End If
            Next j
        Next i
'    Else
'        For i = 1 To count - 1
'            For j = i + 1 To count
'                If a_date(i) < a_date(j) Then
'                    tmp_indx = a_indx(i)
'                    tmp_date = a_date(i)
'                    a_indx(i) = a_indx(j)
'                    a_date(i) = a_date(j)
'                    a_indx(j) = tmp_indx
'                    a_date(j) = tmp_date
'                End If
'            Next j
'        Next i
'    End If
    
    Sort_By_Date = 0
        
End Function

Private Sub BtnAbout_Click()
' ******************************
' 著作権表示
' ******************************
    Dim i As Integer
    i = MsgBox("HTMLメール 発見ツール VBA" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + "(C) 2001-2003 INOUE. Hirokazu" + Chr(&HD) + Chr(&HA) + "version 1.2 / フリーウエア" + Chr(&HD) + Chr(&HA) + "http://inoue-h.connect.to/", vbOKOnly + vbInformation, "HTMLメール 発見ツール VBA")
End Sub

Private Sub BtnCansel_Click()
' ******************************
' キャンセルボタンを押したとき、ダイアログを閉じる
' ******************************
    DlgFindHtml.Hide
End Sub

' ファイル終了 EOF
' ***********************
