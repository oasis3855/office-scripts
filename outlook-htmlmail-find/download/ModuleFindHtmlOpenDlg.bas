Attribute VB_Name = "ModuleFindHtmlOpenDlg"
Option Explicit

Sub FindHtmlOpenDlg()
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
    For i = 1 To DlgFindHtml.CmbboxFolder.ListCount
        ' �S�ẴA�C�e������������
        DlgFindHtml.CmbboxFolder.RemoveItem (0)
    Next i
    DlgFindHtml.CmbboxFolder.AddItem (">> ���I��")
    For i = 1 To myNamespace.Folders.count
        tmpStr = myNamespace.Folders.Item(i)
        DlgFindHtml.CmbboxFolder.AddItem (tmpStr)
    Next i
    DlgFindHtml.CmbboxFolder.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    
    ' ******************************
    ' �g���C�P�A�g���C�Q �̑I�������R���{�{�b�N�X�ɐݒ�
    ' ******************************
    For i = 1 To DlgFindHtml.CmbboxTray1.ListCount
        ' �S�ẴA�C�e������������
        DlgFindHtml.CmbboxTray1.RemoveItem (0)
    Next i
    DlgFindHtml.CmbboxTray1.AddItem (">> ���I��")
    DlgFindHtml.CmbboxTray1.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    For i = 1 To DlgFindHtml.CmbboxTray2.ListCount
        ' �S�ẴA�C�e������������
        DlgFindHtml.CmbboxTray2.RemoveItem (0)
    Next i
    DlgFindHtml.CmbboxTray2.AddItem (">> ���I��")
    DlgFindHtml.CmbboxTray2.ListIndex = 0   ' ��ڂ̍��ڂ�\��
    
    
    
    ' �t�H�[���̕\��
    DlgFindHtml.Show
        
End Sub
    

