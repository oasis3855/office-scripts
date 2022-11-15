Attribute VB_Name = "ModuleCommonDlg"
' ***********************
'   OutlookEmailText : ModuleCommonDialog.bas ver 1.1
'   OutlookHtmlFind : ModuleCommonDialog.bas ver 1.2
'
'   �i��L2�̃v���O�����Ŏg�p����Ă��܂��j
'
' SDK �֐�
' Windows �̃R�����_�C�A���O �u�t�@�C�����J���v�A�u�t�@�C����ۑ�����v
' ***********************
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
(pOpenfilename As OPENFILENAME) As Long

Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
(pOpenfilename As OPENFILENAME) As Long

' ***********************
' OPENFILENAME �\����
' ***********************
Public Type OPENFILENAME
    lStructSize As Long             '���̍\���̂̒���
    hwndOwner As Long               '�Ăяo�����E�C���h�E�n���h��
    hInstance As Long
    lpstrFilter As String           '�t�B���^������
    lpstrCustomFilter As String
    nMaxCustrFilter As Long
    nFilterIndex As Long
    lpstrFile As String             '�I�����ꂽ�t�@�C�����i�t���p�X�j
    nMaxFile As Long                'lpstrFile�̃o�b�t�@�T�C�Y
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String       '�����t�H���_��
    lpstrTitle As String            '�R�����_�C�A���O�̃^�C�g����
    flags As Long                   '�t���O
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String           '�t�@�C�����̓��͎��A�g���q���ȗ����ꂽ���̊g���q
    lCustrData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' �t�@�C���I�� EOF
' ***********************
