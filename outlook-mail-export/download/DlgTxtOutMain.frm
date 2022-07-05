VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgTxtOutMain 
   Caption         =   "Outlook メール テキスト化 VisualBasic for Application"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   OleObjectBlob   =   "DlgTxtOutMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DlgTxtOutMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' *******************
'   Outlook メール テキスト化 VBA ( フォーム DlgTxtOutMain )
'   Version 1.3
'   (C) 2001-2022 INOUE. Hirokazu
'
'   このVBAスクリプトは GNU General Public License v3ライセンスで公開する フリーソフトウエア
'
'   このソフトウエアの機能拡張に貢献してくださった方
'   Mr. Hamada : ver 1.2 UNICODE update
' *******************
'
' FileSystemObjectを利用するため、VBEのツール->参照設定で Microsoft Scripting Runtime を有効化する
'
Option Explicit

Const MAX_MAILS = 5000  ' 日付ソート配列の最大値

' ******************************
' キャンセルボタンを押したとき、ダイアログを閉じる
' ******************************
Private Sub BtnCansel_Click()
    DlgTxtOutMain.Hide
End Sub


' ******************************
' チェックボックス「ソート｣を変更したときの処理
' 「古い順｣のチェックボックスをグレーにするかどうか反映
' ******************************
Private Sub ChkSort_Click()
    If ChkSort.Value = False Then
        ChkSortRev.Enabled = False
    Else
        ChkSortRev.Enabled = True
    End If
End Sub

' ******************************
' フォルダ項目が新たに選択されたとき
' ******************************
Private Sub CmbboxFolder_Change()
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim strTemp As String        ' テンポラリ文字列
    ' Outlook 自体
    Dim olkMAPI As NameSpace
    ' フォルダオブジェクト
    Dim olkFolder1 As MAPIFolder
    Dim olkFolder2 As MAPIFolder
    
    Set olkMAPI = Application.GetNamespace("MAPI")

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
    
    Set olkFolder1 = olkMAPI.Folders(CmbboxFolder.Value)
    
    For i = 1 To CmbboxTray1.ListCount
        ' 全てのアイテムを消去する
        CmbboxTray1.RemoveItem (0)
    Next i
    CmbboxTray1.AddItem (">> 未選択")
    For i = 1 To olkFolder1.Folders.count
        strTemp = olkFolder1.Folders.Item(i)
        CmbboxTray1.AddItem (strTemp)
    Next i
    CmbboxTray1.ListIndex = 0   ' 一つ目の項目を表示
    
End Sub

' ******************************
' トレイ１項目が新たに選択されたとき
' ******************************
Private Sub CmbboxTray1_Change()
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim strTemp As String        ' テンポラリ文字列
    ' Outlook 自体
    Dim olkMAPI As NameSpace
    ' フォルダオブジェクト
    Dim olkFolder1 As MAPIFolder
    Dim olkFolder2 As MAPIFolder
    
    Set olkMAPI = Application.GetNamespace("MAPI")

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
    
    Set olkFolder1 = olkMAPI.Folders(CmbboxFolder.Value)
    Set olkFolder2 = olkFolder1.Folders(CmbboxTray1.Value)
    
    For i = 1 To CmbboxTray2.ListCount
        ' 全てのアイテムを消去する
        CmbboxTray2.RemoveItem (0)
    Next i
    CmbboxTray2.AddItem (">> 未選択")
    For i = 1 To olkFolder2.Folders.count
        strTemp = olkFolder2.Folders.Item(i)
        CmbboxTray2.AddItem (strTemp)
    Next i
    CmbboxTray2.ListIndex = 0   ' 一つ目の項目を表示

End Sub

' ******************************
' 実行ボタンを押したとき
' ******************************
Private Sub BtnExec_Click()
'    On Error GoTo ERROR_TRAP
    ' 一般変数
    Dim i As Integer            ' カウンタ用変数
    Dim strTemp As String       ' テンポラリ文字列
    ' Outlook 自体
    Dim olkMAPI As NameSpace
    ' フォルダオブジェクト
    Dim olkFolder1 As MAPIFolder
    Dim olkFolder2 As MAPIFolder
    Dim olkEmailEnt As MAPIFolder
'    Dim olkMailItem As AppointmentItem     ' VBEでプロパティ等の自動補完入力するときに適宜変更
    Dim olkMailItem As Variant      ' MailItem, ReportItem, AppointmentItem, MeetingItem の複数オブジェクトに対応するため
    
    ' 日付ソート用配列
    ReDim arrIndex(MAX_MAILS) As Integer    ' インデックスの配列
    ReDim arrDate(MAX_MAILS) As Date        ' 日付データの配列
    
    ' 出力テキストファイルオブジェクト
    ' *** [ユーザー定義型は定義されていません] エラーが表示される場合は、
    ' FileSystemObjectを利用するため、VBEのツール->参照設定で Microsoft Scripting Runtime を有効化する
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim strExportFilepath As String ' エクスポート ファイル名
    
    ' unicode 変換制御  ***** 2005/11/18 追加 ver 1.2
    Dim flagUnicodeFile As Boolean  ' TRUE:unicode, FALSE:Shift JIS
    If ChkUnicode.Value = True Then
        flagUnicodeFile = True
    Else
        flagUnicodeFile = False
    End If
    ' ***** 2005/11/18 追加 ver1.2 ここまで
        
        
    strExportFilepath = InputBox("デスクトップ上に作成するメール書き出しファイル名の入力", "メール書き出しファイル名の入力", "outlook_export.txt")
    If strExportFilepath = "" Then
        MsgBox ("キャンセルしました")
        Exit Sub
    End If
    strExportFilepath = MakeDesktopFilepath(strExportFilepath)
    
    MsgBox (strExportFilepath + vbCrLf + " にメールをテキスト出力します")
    
    
    Set olkMAPI = Application.GetNamespace("MAPI")

    If (CmbboxFolder.Value = ">> 未選択") Or (CmbboxTray1.Value = ">> 未選択") Then
        i = MsgBox("フォルダ および トレイ１ を選択する必要があります", vbOKOnly + vbExclamation, "Outlook メール テキスト化 VBA エラー")
            Set olkMAPI = Nothing
        Exit Sub
    End If
    

    Set olkFolder1 = olkMAPI.Folders(CmbboxFolder.Value)
    Set olkFolder2 = olkFolder1.Folders(CmbboxTray1.Value)
    If CmbboxTray2.Value = ">> 未選択" Then
        Set olkEmailEnt = olkFolder2
    Else
        Set olkEmailEnt = olkFolder2.Folders(CmbboxTray2.Value)
    End If
    
    If olkEmailEnt.Items.count <= 0 Then
        MsgBox ("指定されたフォルダにはメールが存在しませんでした")
        Set olkEmailEnt = Nothing
        Set olkFolder2 = Nothing
        Set olkFolder1 = Nothing
        Set olkMAPI = Nothing
        Exit Sub
    End If
    
    ' 日付データによるソーティングを行う
    If (olkEmailEnt.Items.count < MAX_MAILS) And (ChkSort.Value = True) Then
        For i = 1 To olkEmailEnt.Items.count
            arrIndex(i) = i
            Set olkMailItem = olkEmailEnt.Items(i)
            If TypeName(olkMailItem) = "MailItem" Or TypeName(olkMailItem) = "MeetingItem" Then
                ' MailItem, MeetingItemの場合は送信日時
                arrDate(i) = olkMailItem.SentOn
            ElseIf TypeName(olkMailItem) = "AppointmentItem" Then
                ' 予定表は開始時間
                arrDate(i) = olkMailItem.Start
            Else
                ' それ以外(ReportItem, AppointmentItem等)は現在日時
                arrDate(i) = Now()
            End If
        Next i
        i = Sort_By_Date(arrIndex, arrDate, MAX_MAILS, olkEmailEnt.Items.count)
    End If


    ' ファイルシステムのオブジェクトを得る
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strExportFilepath) = True Then
        If vbNo = MsgBox("指定されたファイルはすでに存在します。上書きしますか ？", vbYesNo) Then
            ' 上書きをキャンセルし、処理終了
            Set fs = Nothing
            Set olkMailItem = Nothing
            Set olkEmailEnt = Nothing
            Set olkFolder2 = Nothing
            Set olkFolder1 = Nothing
            Set olkMAPI = Nothing
            Exit Sub
        Else
            ' 上書き
            Set ts = fs.CreateTextFile(strExportFilepath, True, flagUnicodeFile)     ' ***** unicode対応化 2005/11/19
        End If
    Else
        ' テキストファイルを新規作成しオープン
        Set ts = fs.CreateTextFile(strExportFilepath, True, flagUnicodeFile)         ' ***** unicode対応化 2005/11/19
    End If
    ' ヘッダ情報を書く
    strTemp = "メールテキスト化対象フォルダ : " + olkFolder1.Name + " -> " + olkFolder2.Name + " -> " + olkEmailEnt.Name + vbCrLf
    Call WriteTextStream(ts, strTemp, flagUnicodeFile)


    ' 目次（インデックス）を書く
    strTemp = "--------------------------------------------------------------------------" + vbCrLf
    For i = 1 To olkEmailEnt.Items.count
        ' 以後のメンバ変数の参照のために、明示的なオブジェクトに代入
        If (olkEmailEnt.Items.count < MAX_MAILS) And (ChkSort.Value = True) Then
            Set olkMailItem = olkEmailEnt.Items(arrIndex(i))
        Else
            Set olkMailItem = olkEmailEnt.Items(i)
        End If
        
        strTemp = strTemp + Format(i, "0000") + ", "
        ' MailItemとMeetingItemのみ、送信日時・送信者名を表示
        If TypeName(olkMailItem) = "MailItem" Or TypeName(olkMailItem) = "MeetingItem" Then
            strTemp = strTemp + Format(olkMailItem.SentOn, "yy/mm/dd hh:mm:ss") + ", "
            If ChkIndxSentMail.Value = False Then
                ' 発信者名
                strTemp = strTemp + olkMailItem.SenderName + ", "
            Else
                If TypeOf olkMailItem Is MailItem Then
                    ' あて先
                    strTemp = strTemp + olkMailItem.To + ", "
                End If
            End If
        ElseIf TypeName(olkMailItem) = "AppointmentItem" Then
            strTemp = strTemp + Format(olkMailItem.Start, "yy/mm/dd hh:mm:ss") + ", "
        End If
        ' メールタイトル
        strTemp = strTemp + olkMailItem.Subject + vbCrLf
        
    Next i
    strTemp = strTemp + "--------------------------------------------------------------------------" + vbCrLf + vbCrLf + vbCrLf
    Call WriteTextStream(ts, strTemp, flagUnicodeFile)
    
    ' 本文を書く
    For i = 1 To olkEmailEnt.Items.count
        ' 以後のメンバ変数の参照のために、明示的なオブジェクトに代入
        If (olkEmailEnt.Items.count < MAX_MAILS) And (ChkSort.Value = True) Then
            Set olkMailItem = olkEmailEnt.Items(arrIndex(i))
        Else
            Set olkMailItem = olkEmailEnt.Items(i)
        End If

        strTemp = "Message-Id: " + Format(i, "0000") + "  " + TypeName(olkMailItem) + vbCrLf
        strTemp = strTemp + "Subject: " + olkMailItem.Subject + vbCrLf
        
        If TypeOf olkMailItem Is MailItem Then
            ' noop
        ElseIf TypeOf olkMailItem Is ReportItem Then
            ' noop
        ElseIf TypeOf olkMailItem Is MeetingItem Then
            ' 予定表, Teams
            ' Subject, SentOn, SenderName, SenderEmailAddress, ConversationTopic, Body
        Else
            ' 予定表
            ' AppointmentItem
            ' Subject
            GoTo continue
        End If
        
        ' メールヘッダ
        If TypeOf olkMailItem Is MailItem Or TypeOf olkMailItem Is MeetingItem Then
            strTemp = strTemp + "From: " + olkMailItem.SenderName + " <" + olkMailItem.SenderEmailAddress + ">" + vbCrLf
        End If
        If TypeOf olkMailItem Is MailItem Then
            strTemp = strTemp + "ReplyTo ： " + olkMailItem.ReplyRecipientNames + vbCrLf
            strTemp = strTemp + "To: " + olkMailItem.To + vbCrLf
            strTemp = strTemp + "CC： " + olkMailItem.CC + vbCrLf
        End If
        If TypeOf olkMailItem Is MailItem Or TypeOf olkMailItem Is MeetingItem Then
            strTemp = strTemp + "Date: " + Format(olkMailItem.SentOn, "yyyy/mm/dd hh:mm:ss") + vbCrLf
        End If
        Call WriteTextStream(ts, strTemp, flagUnicodeFile)
        
        strTemp = "--------------" + vbCrLf
        ' メール本文
        strTemp = strTemp + olkMailItem.Body + vbCrLf + vbCrLf + vbCrLf
        ' メール1通ごとの区切り線
        strTemp = strTemp + "==========================================================================" + vbCrLf + vbCrLf + vbCrLf
        ' ファイル書き込み
        Call WriteTextStream(ts, strTemp, flagUnicodeFile)
        
continue:
    Next i
    
    ' 処理件数を書き込む
    strTemp = "出力件数 : " + str(olkEmailEnt.Items.count) + vbCrLf
    strTemp = strTemp + "==========================================================================" + vbCrLf
    Call WriteTextStream(ts, strTemp, flagUnicodeFile)
    'ファイルをクローズ
    
    ts.Close
    
    strTemp = "メールデータを " + str(olkEmailEnt.Items.count) + " 件書き込みました"
    i = MsgBox(strTemp, vbOKOnly + vbInformation, "Outlook メール テキスト化 VBA")
    
    Set ts = Nothing
    Set fs = Nothing
    Set olkMailItem = Nothing
    Set olkEmailEnt = Nothing
    Set olkFolder2 = Nothing
    Set olkFolder1 = Nothing
    Set olkMAPI = Nothing
    
    Exit Sub
ERROR_TRAP:
    MsgBox ("ファイル作成・メール抽出時のエラー" & vbCrLf & "LineNo : " & CStr(Erl) & vbCrLf & "ErrNumber : " & Err.Number & vbCrLf & "Description : " & Err.Description & vbCrLf & Err.Source)
    Set ts = Nothing
    Set fs = Nothing
    Set olkMailItem = Nothing
    Set olkEmailEnt = Nothing
    Set olkFolder2 = Nothing
    Set olkFolder1 = Nothing
    Set olkMAPI = Nothing
    Exit Sub
End Sub

' ******************************
' ソーティング （もっとも単純な逐次ソート）
' ******************************
Function Sort_By_Date(ByRef arrIndex() As Integer, ByRef arrDate() As Date, max_a As Integer, count As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim tmp_indx As Integer
    Dim tmp_date As Date
    
    ' 昇順か降順かによって切り替える
    If DlgTxtOutMain.ChkSortRev.Value = True Then
        For i = 1 To count - 1
            For j = i + 1 To count
                If arrDate(i) > arrDate(j) Then
                    tmp_indx = arrIndex(i)
                    tmp_date = arrDate(i)
                    arrIndex(i) = arrIndex(j)
                    arrDate(i) = arrDate(j)
                    arrIndex(j) = tmp_indx
                    arrDate(j) = tmp_date
                End If
            Next j
        Next i
    Else
        For i = 1 To count - 1
            For j = i + 1 To count
                If arrDate(i) < arrDate(j) Then
                    tmp_indx = arrIndex(i)
                    tmp_date = arrDate(i)
                    arrIndex(i) = arrIndex(j)
                    arrDate(i) = arrDate(j)
                    arrIndex(j) = tmp_indx
                    arrDate(j) = tmp_date
                End If
            Next j
        Next i
    End If
    
    Sort_By_Date = 0
        
End Function


' ******************************
' ファイルへの書き込み
' ******************************
Sub WriteTextStream(ByRef ts As TextStream, str As String, flagUtf8 As Boolean)
'    On Error GoTo ERROR_TRAP
    
    ' 行末記号変換（外国語の場合のエラーに対応） ***** 2005/11/19 追加
    ' テキスト中のバイナリコード混入でts.Writeがエラーを出すので、それの対策も含む
    ' VBA内部処理のUnicode(UTF16)をいったんShiftJISに変換し、もう一度Unicodeに戻すことでShifJIS表現できないバイナリ文字などを排除する
    If flagUtf8 = False Then
        str = StrConv(str, vbFromUnicode)   ' UFT8 -> 8bit(SJIS...)
        str = StrConv(str, vbUnicode)       ' 8bit(SJIS...) -> UTF8
    End If
    
    ' TextStreamに書き込み
    ' ファイルオープン時（CreateTextFile）に指定したエンコード方法（Unicode / Shift JIS）に自動変換されて書き込まれる
    ts.Write (str)
    
    Exit Sub
ERROR_TRAP:
    MsgBox ("ファイル書き込み時のエラー" & vbCrLf & "LineNo : " & CStr(Erl) & vbCrLf & "ErrNumber : " & Err.Number & vbCrLf & "Description : " & Err.Description & vbCrLf & Err.Source)
End Sub


' ******************************
' ファイル名にデスクトップディレクトリを付加して、フルパス文字列に変換する
' ******************************
Function MakeDesktopFilepath(strFnameCore As String)
    
    ' パス名で入力されている場合に、「\」で区切り、最後のもののみをファイル名として抜き出すための一時配列
    Dim arrTemp As Variant
    arrTemp = Split(strFnameCore, "\")
    
    ' デスクトップのディレクトリ名を得る
    Dim objWsh As Object
    Set objWsh = CreateObject("Wscript.Shell")
    
    Dim strDesktopFolder As String
    strDesktopFolder = objWsh.SpecialFolders("Desktop")

    ' フルパスを組み立てる
    MakeDesktopFilepath = strDesktopFolder + "\" + arrTemp(UBound(arrTemp))

    Set objWsh = Nothing

End Function


' ファイル終了 EOF
' ***********************
