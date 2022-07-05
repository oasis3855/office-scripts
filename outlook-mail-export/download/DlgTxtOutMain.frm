VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgTxtOutMain 
   Caption         =   "Outlook ���[�� �e�L�X�g�� VisualBasic for Application"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   OleObjectBlob   =   "DlgTxtOutMain.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DlgTxtOutMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' *******************
'   Outlook ���[�� �e�L�X�g�� VBA ( �t�H�[�� DlgTxtOutMain )
'   Version 1.3
'   (C) 2001-2022 INOUE. Hirokazu
'
'   ����VBA�X�N���v�g�� GNU General Public License v3���C�Z���X�Ō��J���� �t���[�\�t�g�E�G�A
'
'   ���̃\�t�g�E�G�A�̋@�\�g���ɍv�����Ă�����������
'   Mr. Hamada : ver 1.2 UNICODE update
' *******************
'
' FileSystemObject�𗘗p���邽�߁AVBE�̃c�[��->�Q�Ɛݒ�� Microsoft Scripting Runtime ��L��������
'
Option Explicit

Const MAX_MAILS = 5000  ' ���t�\�[�g�z��̍ő�l

' ******************************
' �L�����Z���{�^�����������Ƃ��A�_�C�A���O�����
' ******************************
Private Sub BtnCansel_Click()
    DlgTxtOutMain.Hide
End Sub


' ******************************
' �`�F�b�N�{�b�N�X�u�\�[�g���ύX�����Ƃ��̏���
' �u�Â�����̃`�F�b�N�{�b�N�X���O���[�ɂ��邩�ǂ������f
' ******************************
Private Sub ChkSort_Click()
    If ChkSort.Value = False Then
        ChkSortRev.Enabled = False
    Else
        ChkSortRev.Enabled = True
    End If
End Sub

' ******************************
' �t�H���_���ڂ��V���ɑI�����ꂽ�Ƃ�
' ******************************
Private Sub CmbboxFolder_Change()
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim strTemp As String        ' �e���|����������
    ' Outlook ����
    Dim olkMAPI As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim olkFolder1 As MAPIFolder
    Dim olkFolder2 As MAPIFolder
    
    Set olkMAPI = Application.GetNamespace("MAPI")

    ' ���I����I�������ꍇ
    If CmbboxFolder.Value = ">> ���I��" Then
        For i = 1 To CmbboxTray1.ListCount
            ' �S�ẴA�C�e������������
            CmbboxTray1.RemoveItem (0)
        Next i
        CmbboxTray1.AddItem (">> ���I��")
        CmbboxTray1.ListIndex = 0   ' ��ڂ̍��ڂ�\��
        Exit Sub
    End If
    
    Set olkFolder1 = olkMAPI.Folders(CmbboxFolder.Value)
    
    For i = 1 To CmbboxTray1.ListCount
        ' �S�ẴA�C�e������������
        CmbboxTray1.RemoveItem (0)
    Next i
    CmbboxTray1.AddItem (">> ���I��")
    For i = 1 To olkFolder1.Folders.count
        strTemp = olkFolder1.Folders.Item(i)
        CmbboxTray1.AddItem (strTemp)
    Next i
    CmbboxTray1.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    
End Sub

' ******************************
' �g���C�P���ڂ��V���ɑI�����ꂽ�Ƃ�
' ******************************
Private Sub CmbboxTray1_Change()
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim strTemp As String        ' �e���|����������
    ' Outlook ����
    Dim olkMAPI As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim olkFolder1 As MAPIFolder
    Dim olkFolder2 As MAPIFolder
    
    Set olkMAPI = Application.GetNamespace("MAPI")

    ' ���I����I�������ꍇ
    If CmbboxTray1.Value = ">> ���I��" Then
        For i = 1 To CmbboxTray2.ListCount
            ' �S�ẴA�C�e������������
            CmbboxTray2.RemoveItem (0)
        Next i
        CmbboxTray2.AddItem (">> ���I��")
        CmbboxTray2.ListIndex = 0   ' ��ڂ̍��ڂ�\��
        Exit Sub
    End If
    
    Set olkFolder1 = olkMAPI.Folders(CmbboxFolder.Value)
    Set olkFolder2 = olkFolder1.Folders(CmbboxTray1.Value)
    
    For i = 1 To CmbboxTray2.ListCount
        ' �S�ẴA�C�e������������
        CmbboxTray2.RemoveItem (0)
    Next i
    CmbboxTray2.AddItem (">> ���I��")
    For i = 1 To olkFolder2.Folders.count
        strTemp = olkFolder2.Folders.Item(i)
        CmbboxTray2.AddItem (strTemp)
    Next i
    CmbboxTray2.ListIndex = 0   ' ��ڂ̍��ڂ�\��

End Sub

' ******************************
' ���s�{�^�����������Ƃ�
' ******************************
Private Sub BtnExec_Click()
'    On Error GoTo ERROR_TRAP
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim strTemp As String       ' �e���|����������
    ' Outlook ����
    Dim olkMAPI As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim olkFolder1 As MAPIFolder
    Dim olkFolder2 As MAPIFolder
    Dim olkEmailEnt As MAPIFolder
'    Dim olkMailItem As AppointmentItem     ' VBE�Ńv���p�e�B���̎����⊮���͂���Ƃ��ɓK�X�ύX
    Dim olkMailItem As Variant      ' MailItem, ReportItem, AppointmentItem, MeetingItem �̕����I�u�W�F�N�g�ɑΉ����邽��
    
    ' ���t�\�[�g�p�z��
    ReDim arrIndex(MAX_MAILS) As Integer    ' �C���f�b�N�X�̔z��
    ReDim arrDate(MAX_MAILS) As Date        ' ���t�f�[�^�̔z��
    
    ' �o�̓e�L�X�g�t�@�C���I�u�W�F�N�g
    ' *** [���[�U�[��`�^�͒�`����Ă��܂���] �G���[���\�������ꍇ�́A
    ' FileSystemObject�𗘗p���邽�߁AVBE�̃c�[��->�Q�Ɛݒ�� Microsoft Scripting Runtime ��L��������
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim strExportFilepath As String ' �G�N�X�|�[�g �t�@�C����
    
    ' unicode �ϊ�����  ***** 2005/11/18 �ǉ� ver 1.2
    Dim flagUnicodeFile As Boolean  ' TRUE:unicode, FALSE:Shift JIS
    If ChkUnicode.Value = True Then
        flagUnicodeFile = True
    Else
        flagUnicodeFile = False
    End If
    ' ***** 2005/11/18 �ǉ� ver1.2 �����܂�
        
        
    strExportFilepath = InputBox("�f�X�N�g�b�v��ɍ쐬���郁�[�������o���t�@�C�����̓���", "���[�������o���t�@�C�����̓���", "outlook_export.txt")
    If strExportFilepath = "" Then
        MsgBox ("�L�����Z�����܂���")
        Exit Sub
    End If
    strExportFilepath = MakeDesktopFilepath(strExportFilepath)
    
    MsgBox (strExportFilepath + vbCrLf + " �Ƀ��[�����e�L�X�g�o�͂��܂�")
    
    
    Set olkMAPI = Application.GetNamespace("MAPI")

    If (CmbboxFolder.Value = ">> ���I��") Or (CmbboxTray1.Value = ">> ���I��") Then
        i = MsgBox("�t�H���_ ����� �g���C�P ��I������K�v������܂�", vbOKOnly + vbExclamation, "Outlook ���[�� �e�L�X�g�� VBA �G���[")
            Set olkMAPI = Nothing
        Exit Sub
    End If
    

    Set olkFolder1 = olkMAPI.Folders(CmbboxFolder.Value)
    Set olkFolder2 = olkFolder1.Folders(CmbboxTray1.Value)
    If CmbboxTray2.Value = ">> ���I��" Then
        Set olkEmailEnt = olkFolder2
    Else
        Set olkEmailEnt = olkFolder2.Folders(CmbboxTray2.Value)
    End If
    
    If olkEmailEnt.Items.count <= 0 Then
        MsgBox ("�w�肳�ꂽ�t�H���_�ɂ̓��[�������݂��܂���ł���")
        Set olkEmailEnt = Nothing
        Set olkFolder2 = Nothing
        Set olkFolder1 = Nothing
        Set olkMAPI = Nothing
        Exit Sub
    End If
    
    ' ���t�f�[�^�ɂ��\�[�e�B���O���s��
    If (olkEmailEnt.Items.count < MAX_MAILS) And (ChkSort.Value = True) Then
        For i = 1 To olkEmailEnt.Items.count
            arrIndex(i) = i
            Set olkMailItem = olkEmailEnt.Items(i)
            If TypeName(olkMailItem) = "MailItem" Or TypeName(olkMailItem) = "MeetingItem" Then
                ' MailItem, MeetingItem�̏ꍇ�͑��M����
                arrDate(i) = olkMailItem.SentOn
            ElseIf TypeName(olkMailItem) = "AppointmentItem" Then
                ' �\��\�͊J�n����
                arrDate(i) = olkMailItem.Start
            Else
                ' ����ȊO(ReportItem, AppointmentItem��)�͌��ݓ���
                arrDate(i) = Now()
            End If
        Next i
        i = Sort_By_Date(arrIndex, arrDate, MAX_MAILS, olkEmailEnt.Items.count)
    End If


    ' �t�@�C���V�X�e���̃I�u�W�F�N�g�𓾂�
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strExportFilepath) = True Then
        If vbNo = MsgBox("�w�肳�ꂽ�t�@�C���͂��łɑ��݂��܂��B�㏑�����܂��� �H", vbYesNo) Then
            ' �㏑�����L�����Z�����A�����I��
            Set fs = Nothing
            Set olkMailItem = Nothing
            Set olkEmailEnt = Nothing
            Set olkFolder2 = Nothing
            Set olkFolder1 = Nothing
            Set olkMAPI = Nothing
            Exit Sub
        Else
            ' �㏑��
            Set ts = fs.CreateTextFile(strExportFilepath, True, flagUnicodeFile)     ' ***** unicode�Ή��� 2005/11/19
        End If
    Else
        ' �e�L�X�g�t�@�C����V�K�쐬���I�[�v��
        Set ts = fs.CreateTextFile(strExportFilepath, True, flagUnicodeFile)         ' ***** unicode�Ή��� 2005/11/19
    End If
    ' �w�b�_��������
    strTemp = "���[���e�L�X�g���Ώۃt�H���_ : " + olkFolder1.Name + " -> " + olkFolder2.Name + " -> " + olkEmailEnt.Name + vbCrLf
    Call WriteTextStream(ts, strTemp, flagUnicodeFile)


    ' �ڎ��i�C���f�b�N�X�j������
    strTemp = "--------------------------------------------------------------------------" + vbCrLf
    For i = 1 To olkEmailEnt.Items.count
        ' �Ȍ�̃����o�ϐ��̎Q�Ƃ̂��߂ɁA�����I�ȃI�u�W�F�N�g�ɑ��
        If (olkEmailEnt.Items.count < MAX_MAILS) And (ChkSort.Value = True) Then
            Set olkMailItem = olkEmailEnt.Items(arrIndex(i))
        Else
            Set olkMailItem = olkEmailEnt.Items(i)
        End If
        
        strTemp = strTemp + Format(i, "0000") + ", "
        ' MailItem��MeetingItem�̂݁A���M�����E���M�Җ���\��
        If TypeName(olkMailItem) = "MailItem" Or TypeName(olkMailItem) = "MeetingItem" Then
            strTemp = strTemp + Format(olkMailItem.SentOn, "yy/mm/dd hh:mm:ss") + ", "
            If ChkIndxSentMail.Value = False Then
                ' ���M�Җ�
                strTemp = strTemp + olkMailItem.SenderName + ", "
            Else
                If TypeOf olkMailItem Is MailItem Then
                    ' ���Đ�
                    strTemp = strTemp + olkMailItem.To + ", "
                End If
            End If
        ElseIf TypeName(olkMailItem) = "AppointmentItem" Then
            strTemp = strTemp + Format(olkMailItem.Start, "yy/mm/dd hh:mm:ss") + ", "
        End If
        ' ���[���^�C�g��
        strTemp = strTemp + olkMailItem.Subject + vbCrLf
        
    Next i
    strTemp = strTemp + "--------------------------------------------------------------------------" + vbCrLf + vbCrLf + vbCrLf
    Call WriteTextStream(ts, strTemp, flagUnicodeFile)
    
    ' �{��������
    For i = 1 To olkEmailEnt.Items.count
        ' �Ȍ�̃����o�ϐ��̎Q�Ƃ̂��߂ɁA�����I�ȃI�u�W�F�N�g�ɑ��
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
            ' �\��\, Teams
            ' Subject, SentOn, SenderName, SenderEmailAddress, ConversationTopic, Body
        Else
            ' �\��\
            ' AppointmentItem
            ' Subject
            GoTo continue
        End If
        
        ' ���[���w�b�_
        If TypeOf olkMailItem Is MailItem Or TypeOf olkMailItem Is MeetingItem Then
            strTemp = strTemp + "From: " + olkMailItem.SenderName + " <" + olkMailItem.SenderEmailAddress + ">" + vbCrLf
        End If
        If TypeOf olkMailItem Is MailItem Then
            strTemp = strTemp + "ReplyTo �F " + olkMailItem.ReplyRecipientNames + vbCrLf
            strTemp = strTemp + "To: " + olkMailItem.To + vbCrLf
            strTemp = strTemp + "CC�F " + olkMailItem.CC + vbCrLf
        End If
        If TypeOf olkMailItem Is MailItem Or TypeOf olkMailItem Is MeetingItem Then
            strTemp = strTemp + "Date: " + Format(olkMailItem.SentOn, "yyyy/mm/dd hh:mm:ss") + vbCrLf
        End If
        Call WriteTextStream(ts, strTemp, flagUnicodeFile)
        
        strTemp = "--------------" + vbCrLf
        ' ���[���{��
        strTemp = strTemp + olkMailItem.Body + vbCrLf + vbCrLf + vbCrLf
        ' ���[��1�ʂ��Ƃ̋�؂��
        strTemp = strTemp + "==========================================================================" + vbCrLf + vbCrLf + vbCrLf
        ' �t�@�C����������
        Call WriteTextStream(ts, strTemp, flagUnicodeFile)
        
continue:
    Next i
    
    ' ������������������
    strTemp = "�o�͌��� : " + str(olkEmailEnt.Items.count) + vbCrLf
    strTemp = strTemp + "==========================================================================" + vbCrLf
    Call WriteTextStream(ts, strTemp, flagUnicodeFile)
    '�t�@�C�����N���[�Y
    
    ts.Close
    
    strTemp = "���[���f�[�^�� " + str(olkEmailEnt.Items.count) + " ���������݂܂���"
    i = MsgBox(strTemp, vbOKOnly + vbInformation, "Outlook ���[�� �e�L�X�g�� VBA")
    
    Set ts = Nothing
    Set fs = Nothing
    Set olkMailItem = Nothing
    Set olkEmailEnt = Nothing
    Set olkFolder2 = Nothing
    Set olkFolder1 = Nothing
    Set olkMAPI = Nothing
    
    Exit Sub
ERROR_TRAP:
    MsgBox ("�t�@�C���쐬�E���[�����o���̃G���[" & vbCrLf & "LineNo : " & CStr(Erl) & vbCrLf & "ErrNumber : " & Err.Number & vbCrLf & "Description : " & Err.Description & vbCrLf & Err.Source)
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
' �\�[�e�B���O �i�����Ƃ��P���Ȓ����\�[�g�j
' ******************************
Function Sort_By_Date(ByRef arrIndex() As Integer, ByRef arrDate() As Date, max_a As Integer, count As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim tmp_indx As Integer
    Dim tmp_date As Date
    
    ' �������~�����ɂ���Đ؂�ւ���
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
' �t�@�C���ւ̏�������
' ******************************
Sub WriteTextStream(ByRef ts As TextStream, str As String, flagUtf8 As Boolean)
'    On Error GoTo ERROR_TRAP
    
    ' �s���L���ϊ��i�O����̏ꍇ�̃G���[�ɑΉ��j ***** 2005/11/19 �ǉ�
    ' �e�L�X�g���̃o�C�i���R�[�h������ts.Write���G���[���o���̂ŁA����̑΍���܂�
    ' VBA����������Unicode(UTF16)����������ShiftJIS�ɕϊ����A������xUnicode�ɖ߂����Ƃ�ShifJIS�\���ł��Ȃ��o�C�i�������Ȃǂ�r������
    If flagUtf8 = False Then
        str = StrConv(str, vbFromUnicode)   ' UFT8 -> 8bit(SJIS...)
        str = StrConv(str, vbUnicode)       ' 8bit(SJIS...) -> UTF8
    End If
    
    ' TextStream�ɏ�������
    ' �t�@�C���I�[�v�����iCreateTextFile�j�Ɏw�肵���G���R�[�h���@�iUnicode / Shift JIS�j�Ɏ����ϊ�����ď������܂��
    ts.Write (str)
    
    Exit Sub
ERROR_TRAP:
    MsgBox ("�t�@�C���������ݎ��̃G���[" & vbCrLf & "LineNo : " & CStr(Erl) & vbCrLf & "ErrNumber : " & Err.Number & vbCrLf & "Description : " & Err.Description & vbCrLf & Err.Source)
End Sub


' ******************************
' �t�@�C�����Ƀf�X�N�g�b�v�f�B���N�g����t�����āA�t���p�X������ɕϊ�����
' ******************************
Function MakeDesktopFilepath(strFnameCore As String)
    
    ' �p�X���œ��͂���Ă���ꍇ�ɁA�u\�v�ŋ�؂�A�Ō�̂��݂̂̂��t�@�C�����Ƃ��Ĕ����o�����߂̈ꎞ�z��
    Dim arrTemp As Variant
    arrTemp = Split(strFnameCore, "\")
    
    ' �f�X�N�g�b�v�̃f�B���N�g�����𓾂�
    Dim objWsh As Object
    Set objWsh = CreateObject("Wscript.Shell")
    
    Dim strDesktopFolder As String
    strDesktopFolder = objWsh.SpecialFolders("Desktop")

    ' �t���p�X��g�ݗ��Ă�
    MakeDesktopFilepath = strDesktopFolder + "\" + arrTemp(UBound(arrTemp))

    Set objWsh = Nothing

End Function


' �t�@�C���I�� EOF
' ***********************
