VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgFindHtml 
   Caption         =   "HTML���[�� �����c�[��"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   OleObjectBlob   =   "DlgFindHtml.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DlgFindHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************
'   OutlookHtmlFind : DlgFindHtml.frm ver 1.2
'
'   Outlook HTML���[�������c�[�� VBA ���C���_�C�A���O�̃R�[�h
'   version 1.2 for Microsoft Outlook 2000 Japanese Edition
'
'   (C) 2001-2003 INOUE. Hirokazu , All rights reserved
'   http://inoue-h.connect.to/
'  ���̃v���O�����^�X�N���v�g�̓t���[�E�G�A�[�ł�
'  ���̃v���O�����^�X�N���v�g�ɑ΂��铮��E�񓮍�̕ۏ؁A���s���ʂ̕ۏ؂͂���܂���
'
'
' �� �d�v �� Outlook�́u�c�[���-�u�}�N���-�u�Z�L�����e�B����j���[�̐ݒ肪�A�u����ȉ��Ŗ����Ǝ��s�ł��Ȃ��B
'
' *******************
Option Explicit

' ******************************
' �\�[�g���̍ő���w�肵�܂��B�傫������ƁA��������H���܂�
' ******************************
Const max_a = 2000  ' ���t�\�[�g�z��̍ő�l

Private Sub BtnExec_Click()
' ******************************
' ���s�{�^�����������Ƃ�
' �u�e�L�X�g���c�[����𗬗p
' ******************************
    On Error GoTo BtnExec_ErrHandler
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim j As Integer            ' �J�E���^�p�ϐ�
    Dim tmpStr As String        ' �e���|����������
    ' Outlook ����
    Dim myNamespace As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    Dim OlkEmailEnt As MAPIFolder
    Dim OlkEmailItem As MailItem    ' MailItem �Ŗ����I�ɐ錾
    
    ' ���t�\�[�g�p�z��
    ReDim a_indx(max_a) As Integer ' �C���f�b�N�X�̔z��
    ReDim a_date(max_a) As Date    ' ���t�f�[�^�̔z��
    
    ' �o�̓e�L�X�g
    Dim OutputStr As String
    
    
    Set myNamespace = Application.GetNamespace("MAPI")

    If (CmbboxFolder.Value = ">> ���I��") Or (CmbboxTray1.Value = ">> ���I��") Then
        i = MsgBox("�t�H���_ ����� �g���C�P ��I������K�v������܂�", vbOKOnly + vbExclamation, "Outlook ���[�� �e�L�X�g�� VBA �G���[")
        Exit Sub
    End If
    
    
    Set OlkEmailTray1 = myNamespace.Folders(CmbboxFolder.Value)
    Set OlkEmailTray2 = OlkEmailTray1.Folders(CmbboxTray1.Value)
    If CmbboxTray2.Value = ">> ���I��" Then
        Set OlkEmailEnt = OlkEmailTray2
    Else
        Set OlkEmailEnt = OlkEmailTray2.Folders(CmbboxTray2.Value)
    End If
    
    ' ���t�f�[�^�ɂ��\�[�e�B���O���s��
    If (OlkEmailEnt.Items.count < max_a) Then
        For i = 1 To OlkEmailEnt.Items.count
            a_indx(i) = i
            Set OlkEmailItem = OlkEmailEnt.Items(i)
            a_date(i) = OlkEmailItem.SentOn
        Next i
        i = Sort_By_Date(a_indx, a_date, max_a, OlkEmailEnt.Items.count)
    End If
                
    OutputStr = "�������ꂽHTML���[���́u�薼��A���M�Җ��A���M���͎��̂Ƃ���ł�" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA)
    j = 0       ' HTML���[���̐����J�E���g����
                
    For i = 1 To OlkEmailEnt.Items.count
        ' �Ȍ�̃����o�ϐ��̎Q�Ƃ̂��߂ɁA�����I�ȃI�u�W�F�N�g�ɑ��
        If (OlkEmailEnt.Items.count < max_a) Then
            Set OlkEmailItem = OlkEmailEnt.Items(a_indx(i))
        Else
            Set OlkEmailItem = OlkEmailEnt.Items(i)
        End If

        
    If OlkEmailItem.HTMLBody <> "" Then
        j = j + 1
        OutputStr = OutputStr + "�u" + OlkEmailItem.Subject + " �v" + OlkEmailItem.SenderName + "  on " + Format(OlkEmailItem.SentOn, "yy/mm/dd hh:mm:ss") + Chr(&HD) + Chr(&HA)
    End If
        
    Next i
    
    ' ���ʃ_�C�A���O���o��
    If j = 0 Then
        i = MsgBox("HTML���[���͔�������܂���ł���", vbOKOnly + vbInformation, "HTML���[�� �����c�[�� VBA")
    Else
        OutputStr = OutputStr + "�ȏ�A���v " + Str(j) + " �ʂ̃��[������������܂����B" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + "�����̃��[�����e�L�X�g�t�@�C���ɏo�͂��܂����H"
        
        i = MsgBox(OutputStr, vbYesNo + vbExclamation, "HTML���[�� �����c�[�� VBA")
        If i = vbYes Then
            ' �e�L�X�g�t�@�C��������T�u���[�`����
            Call OutputTextFile(CmbboxFolder.Text, CmbboxTray1.Text, CmbboxTray2.Text)
        End If
    End If
    
    Exit Sub
BtnExec_ErrHandler:
    i = MsgBox("�G���[���������܂����B�����𒆎~���܂��B", vbOKOnly + vbExclamation, "Outlook ���[�� �e�L�X�g�� VBA �v���I�G���[")
    Exit Sub

End Sub

Private Sub CmbboxFolder_Change()
' ******************************
' �t�H���_���ڂ��V���ɑI�����ꂽ�Ƃ�
' ******************************
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim tmpStr As String        ' �e���|����������
    ' Outlook ����
    Dim myNamespace As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    
    Set myNamespace = Application.GetNamespace("MAPI")

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
    
    Set OlkEmailTray1 = myNamespace.Folders(CmbboxFolder.Value)
    
    For i = 1 To CmbboxTray1.ListCount
        ' �S�ẴA�C�e������������
        CmbboxTray1.RemoveItem (0)
    Next i
    CmbboxTray1.AddItem (">> ���I��")
    For i = 1 To OlkEmailTray1.Folders.count
        tmpStr = OlkEmailTray1.Folders.Item(i)
        CmbboxTray1.AddItem (tmpStr)
    Next i
    CmbboxTray1.ListIndex = 0   ' ��ڂ̍��ڂ�\��
End Sub

Private Sub CmbboxTray1_Change()
' ******************************
' �g���C�P���ڂ��V���ɑI�����ꂽ�Ƃ�
' ******************************
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim tmpStr As String        ' �e���|����������
    ' Outlook ����
    Dim myNamespace As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    
    Set myNamespace = Application.GetNamespace("MAPI")

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
    
    Set OlkEmailTray1 = myNamespace.Folders(CmbboxFolder.Value)
    Set OlkEmailTray2 = OlkEmailTray1.Folders(CmbboxTray1.Value)
    
    For i = 1 To CmbboxTray2.ListCount
        ' �S�ẴA�C�e������������
        CmbboxTray2.RemoveItem (0)
    Next i
    CmbboxTray2.AddItem (">> ���I��")
    For i = 1 To OlkEmailTray2.Folders.count
        tmpStr = OlkEmailTray2.Folders.Item(i)
        CmbboxTray2.AddItem (tmpStr)
    Next i
    CmbboxTray2.ListIndex = 0   ' ��ڂ̍��ڂ�\��

End Sub

Private Sub OutputTextFile(tray0 As String, tray1 As String, tray2 As String)
' ******************************
' HTML���[���݂̂��e�L�X�g�t�@�C���ɏ����o���T�u���[�`��
' �u�e�L�X�g���c�[����𗬗p
' ******************************
    
    Dim strFname As String      ' �o�̓t�@�C����
    Dim tmpStr As String        ' �e���|����������
    ' Windows �̃R�����_�C�A���O �p�̍\����
    Dim strOfn As OPENFILENAME
    Dim nLRet As Long, nNULLPos As Long
    
    With strOfn
        .lStructSize = Len(strOfn)
        .lpstrInitialDir = ""      '�i�ŏ��ɕ\������f�B���N�g���j
                                            '�i�t�B���^�[�Ńt�@�C����ނ��i��j
        .lpstrFilter = "�e�L�X�g�t�@�C��(*.txt)" & vbNullChar & "*.txt" _
        & vbNullChar & "�S�Ẵt�@�C��(*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .nMaxFile = 256                        '�i�t�@�C�����̍ő咷�i�p�X�܂ށj�j
        .lpstrFile = String(256, vbNullChar)   '�i�t�@�C�������i�[���镶����
                                                ' NULL�Ŗ��߂Ă����j
        .lpstrTitle = "�f�[�^���������ރe�L�X�g�t�@�C�����̎w��"
    End With
    
    ' �u�t�@�C����ۑ�����v�_�C�A���O
    nLRet = GetSaveFileName(strOfn)
    
    nNULLPos = InStr(strOfn.lpstrFile, vbNullChar)  '�t�@�C�����̏I��(NULL�̈ʒu)�𒲂ׂ�
    tmpStr = Left(strOfn.lpstrFile, nNULLPos - 1) '�t�@�C�����̗L�����������o��
    
    ' �L�����Z���{�^���������ꂽ
    If nLRet = False Then
        MsgBox ("�o�͂𒆎~���܂�")
        Exit Sub
    End If
    ' �g���q .txt ���w�肳��Ȃ������ꍇ .txt ��t����
    If InStr(tmpStr, ".txt") < 1 Then
        ' �g���q�t�B���^�� .txt �̂Ƃ�
        If strOfn.nFilterIndex = 1 Then
            tmpStr = tmpStr + ".txt"
        End If
    End If
    
    ' �o�̓t�@�C�������m��
    strFname = tmpStr
    
    On Error GoTo BtnExec_ErrHandler
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim j As Integer            ' �J�E���^�p�ϐ�
    ' Outlook ����
    Dim myNamespace As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    Dim OlkEmailEnt As MAPIFolder
    Dim OlkEmailItem As MailItem    ' MailItem �Ŗ����I�ɐ錾
    
    ' ���t�\�[�g�p�z��
    ReDim a_indx(max_a) As Integer ' �C���f�b�N�X�̔z��
    ReDim a_date(max_a) As Date    ' ���t�f�[�^�̔z��
    
    ' �o�̓e�L�X�g�t�@�C���I�u�W�F�N�g
    Dim fs                      ' FileSystemObject
    Dim fi_out                  ' TextStream
    Dim FileName As String      ' �t�@�C����
    
    Set myNamespace = Application.GetNamespace("MAPI")

    Set OlkEmailTray1 = myNamespace.Folders(tray0)
    Set OlkEmailTray2 = OlkEmailTray1.Folders(tray1)
    If tray2 = ">> ���I��" Then
        Set OlkEmailEnt = OlkEmailTray2
    Else
        Set OlkEmailEnt = OlkEmailTray2.Folders(tray2)
    End If
    
    ' ���t�f�[�^�ɂ��\�[�e�B���O���s��
    If OlkEmailEnt.Items.count < max_a Then
        For i = 1 To OlkEmailEnt.Items.count
            a_indx(i) = i
            Set OlkEmailItem = OlkEmailEnt.Items(i)
            a_date(i) = OlkEmailItem.SentOn
        Next i
        i = Sort_By_Date(a_indx, a_date, max_a, OlkEmailEnt.Items.count)
    End If
            

    ' �t�@�C���V�X�e���̃I�u�W�F�N�g�𓾂�
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strFname) = True Then
        If vbNo = MsgBox("�w�肳�ꂽ�t�@�C���͂��łɑ��݂��܂��B�㏑�����܂��� �H" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + "  �u���������I������ƁA�����t�@�C���ɒǉ����郂�[�h�ƂȂ�܂�", vbYesNo + vbQuestion, "Outlook ���[�� �e�L�X�g�� VBA �m�F") Then
            ' �����t�@�C���ɒǉ�
            Set fi_out = fs.OpenTextFile(strFname, 8, True)
        Else
            ' �㏑��
            Set fi_out = fs.CreateTextFile(strFname, True)
        End If
    Else
        ' �e�L�X�g�t�@�C����V�K�쐬���I�[�v��
        Set fi_out = fs.CreateTextFile(strFname, True)
    End If
    ' �w�b�_��������
    tmpStr = "Outlook ���[�� �e�L�X�g�� VBA " + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    tmpStr = "X#########################################################################" + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    
    j = 0
    For i = 1 To OlkEmailEnt.Items.count
        ' �Ȍ�̃����o�ϐ��̎Q�Ƃ̂��߂ɁA�����I�ȃI�u�W�F�N�g�ɑ��
        If OlkEmailEnt.Items.count < max_a Then
            Set OlkEmailItem = OlkEmailEnt.Items(a_indx(i))
        Else
            Set OlkEmailItem = OlkEmailEnt.Items(i)
        End If

        If OlkEmailItem.HTMLBody <> "" Then
            j = j + 1
            tmpStr = "�薼 �F �u" + OlkEmailItem.Subject + " �v" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "���M�� �F " + OlkEmailItem.SentOnBehalfOfName + " / " + OlkEmailItem.SenderName + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "ReplyTo �F " + OlkEmailItem.ReplyRecipientNames + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "���� �F " + OlkEmailItem.To + "  CC �F " + OlkEmailItem.CC + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "BCC �F " + OlkEmailItem.BCC + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "���M�� �F " + Format(OlkEmailItem.SentOn, "yy/mm/dd hh:mm:ss") + "  ��M�� �F " + Format(OlkEmailItem.ReceivedTime, "yy/mm/dd hh/mm/ss") + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "�{�� �F " + Chr(&HD) + Chr(&HA) + OlkEmailItem.Body + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "X#########################################################################" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "HTML�̖{�� �F " + Chr(&HD) + Chr(&HA) + OlkEmailItem.HTMLBody + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "X#########################################################################" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "��������������������������������������������������������������������������" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            tmpStr = "##########################################################################" + Chr(&HD) + Chr(&HA)
            fi_out.Write tmpStr
            
        End If
        
    Next i
    
    ' ������������������
    tmpStr = Chr(&HD) + Chr(&HA) + "���� : " + Str(j) + "     �����I��" + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    tmpStr = "E#########################################################################" + Chr(&HD) + Chr(&HA)
    fi_out.Write tmpStr
    '�t�@�C�����N���[�Y
    fi_out.Close
    
    tmpStr = "�e�L�X�g���������[���f�[�^�� " + Str(j) + " ���������݂܂���"
    i = MsgBox(tmpStr, vbOKOnly + vbInformation, "Outlook ���[�� �e�L�X�g�� VBA")
    
    Exit Sub
BtnExec_ErrHandler:
    i = MsgBox("�e�L�X�g�t�@�C���������ݒ��ɃG���[���������܂����B�����𒆎~���܂��B", vbOKOnly + vbExclamation, "Outlook ���[�� �e�L�X�g�� VBA �v���I�G���[")
    Exit Sub

End Sub

Function Sort_By_Date(a_indx() As Integer, a_date() As Date, max_a As Integer, count As Integer)
' ******************************
' �\�[�e�B���O �i�Œx���[�h�j
' �����ƃX�}�[�g�ɂ������Ȃ�A���������Ă��������B�i���̃A���S���Y���Ɂj
' �ł��APentium III �Ƃ��A����CPU�ł͂قƂ�Ǒ̊��ł��Ȃ��͂��ł����c
' ******************************
    Dim i As Integer
    Dim j As Integer
    Dim tmp_indx As Integer
    Dim tmp_date As Date
    
    ' �������~�����ɂ���Đ؂�ւ���
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
' ���쌠�\��
' ******************************
    Dim i As Integer
    i = MsgBox("HTML���[�� �����c�[�� VBA" + Chr(&HD) + Chr(&HA) + Chr(&HD) + Chr(&HA) + "(C) 2001-2003 INOUE. Hirokazu" + Chr(&HD) + Chr(&HA) + "version 1.2 / �t���[�E�G�A" + Chr(&HD) + Chr(&HA) + "http://inoue-h.connect.to/", vbOKOnly + vbInformation, "HTML���[�� �����c�[�� VBA")
End Sub

Private Sub BtnCansel_Click()
' ******************************
' �L�����Z���{�^�����������Ƃ��A�_�C�A���O�����
' ******************************
    DlgFindHtml.Hide
End Sub

' �t�@�C���I�� EOF
' ***********************
