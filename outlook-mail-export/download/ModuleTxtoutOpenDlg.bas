Attribute VB_Name = "ModuleTxtoutOpenDlg"
' *******************
'   Outlook ���[�� �e�L�X�g�� VBA ( �}�N���Ăяo�� ModuleTxtoutOpenDlgn )
'   Version 1.3
'   (C) 2001-2022 INOUE. Hirokazu
'
'   ����VBA�X�N���v�g�� GNU General Public License v3���C�Z���X�Ō��J���� �t���[�\�t�g�E�G�A
' *******************
Option Explicit

Sub TxtoutOpenDlg()
    
    ' ��ʕϐ�
    Dim i As Integer            ' �J�E���^�p�ϐ�
    Dim tmpStr As String        ' �e���|����������
    ' Outlook ����
    Dim myNamespace As NameSpace
    ' �t�H���_�I�u�W�F�N�g
    Dim OlkEmailFolder As MAPIFolder
    Dim OlkEmailTray1 As MAPIFolder
    Dim OlkEmailTray2 As MAPIFolder
    
    Set myNamespace = Application.GetNamespace("MAPI")
    
    ' ******************************
    ' �t�H���_�i�ŏ�K�j�̑I�������R���{�{�b�N�X�ɐݒ�
    ' ******************************
    For i = 1 To DlgTxtOutMain.CmbboxFolder.ListCount
        ' �S�ẴA�C�e������������
        DlgTxtOutMain.CmbboxFolder.RemoveItem (0)
    Next i
    DlgTxtOutMain.CmbboxFolder.AddItem (">> ���I��")
    For i = 1 To myNamespace.Folders.count
        tmpStr = myNamespace.Folders.Item(i)
        DlgTxtOutMain.CmbboxFolder.AddItem (tmpStr)
    Next i
    DlgTxtOutMain.CmbboxFolder.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    
    ' ******************************
    ' �g���C�P�A�g���C�Q �̑I�������R���{�{�b�N�X�ɐݒ�
    ' ******************************
    For i = 1 To DlgTxtOutMain.CmbboxTray1.ListCount
        ' �S�ẴA�C�e������������
        DlgTxtOutMain.CmbboxTray1.RemoveItem (0)
    Next i
    DlgTxtOutMain.CmbboxTray1.AddItem (">> ���I��")
    DlgTxtOutMain.CmbboxTray1.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    For i = 1 To DlgTxtOutMain.CmbboxTray2.ListCount
        ' �S�ẴA�C�e������������
        DlgTxtOutMain.CmbboxTray2.RemoveItem (0)
    Next i
    DlgTxtOutMain.CmbboxTray2.AddItem (">> ���I��")
    DlgTxtOutMain.CmbboxTray2.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    
    ' ******************************
    ' �`�F�b�N�{�b�N�X�̐ݒ�
    ' ******************************
    DlgTxtOutMain.ChkSort.Value = True
    DlgTxtOutMain.ChkSortRev.Value = True
    DlgTxtOutMain.ChkIndxSentMail.Value = False
    DlgTxtOutMain.ChkUnicode.Value = True
    
    
    ' �t�H�[���̕\��
    DlgTxtOutMain.Show
    
    ' �_�C�A���O���ڂ̕\���E���͒l�Ȃǂ��N���A���邽�߂ɏ���������
    Set DlgTxtOutMain = Nothing
        
End Sub



' �t�@�C���I�� EOF
' ***********************
