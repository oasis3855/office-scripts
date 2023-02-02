Attribute VB_Name = "Module3"
Option Explicit
' *****
' ���ʊ֐��i�t�@�C���I���_�C�A���O�j
'
'  Function FileOpenDlg(ByVal strTitleMsg As String, ByRef arrFilter() As Variant) As String
'  Function FileSaveDlg(ByVal strTitleMsg As String, ByVal strFilter As String, ByVal boolExistDelete As Boolean) As String
' *****


' *****
' �ǂݍ��݃t�@�C������I������_�C�A���O��\���B�t�@�C�����݃`�G�b�N�t��
' ���� :
'   strTitileMsg : �_�C�A���O�̃E�C���h�E�^�C�g��������
'   arrFilter()  : �t�@�C����ʑI���t�B���^�[
' �߂�l :
'   ����I�� : �t�@�C���̃t���p�X������
'   �L�����Z���{�^�� or �t�@�C���s���� : "" ������
' *****
Function FileOpenDlg(ByVal strTitleMsg As String, ByRef arrFilter() As Variant) As String
    FileOpenDlg = ""        ' �߂�l�̏����l�ݒ�

    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    ' �f�X�N�g�b�v�t�H���_�������l�Ƃ���B������ \ �ŏI��邱�ƂŁA�f�B���N�g������
    DesktopFolderName = ShellObject.SpecialFolders("Desktop") & "\"
    Set ShellObject = Nothing

    ' *****
    ' �ǂݍ��݃t�@�C������I���E���͂���GUI�I�u�W�F�N�g���\�z
    ' *****
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = strTitleMsg
        With .Filters
            ' **
            ' �t�B���^�[�̐ݒ�B�����͎��̂悤�Ȃ��̂ŁAfunction�̈������ݒ肷��
            ' .Add "�e�L�X�g�t�@�C��", "*.txt", 1
            ' **
            .Clear
            Dim numItems As Integer
            numItems = (UBound(arrFilter) + 1) / 2
            Dim i As Integer
            For i = 0 To numItems - 1
                .Add arrFilter(i * 2), arrFilter(i * 2 + 1), i + 1
            Next i
        End With
        .InitialFileName = DesktopFolderName
    End With

    Set ShellObject = Nothing

    ' *****
    ' �ǂݍ��݃t�@�C������I���E���͂���GUI�_�C�A���O��\��
    ' *****
    If fd.Show <> True Then
        MsgBox ("�L�����Z���{�^����������܂���")
        Set fd = Nothing
        Exit Function
    End If
    
    FileOpenDlg = fd.SelectedItems(1)       ' �I�����ꂽ�t�@�C���̃t���p�X��
    Set fd = Nothing
    
    ' *****
    ' ���̓t�@�C���̑��݂��m�F
    ' *****
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FileOpenDlg) <> True Then
        MsgBox ("���̓t�@�C����������܂���" & vbCrLf & FileOpenDlg)
        Set fso = Nothing
        Exit Function
    End If
    Set fso = Nothing
 

End Function


' *****
' �������݃t�@�C������I������_�C�A���O��\���B�t�@�C�����݃`�G�b�N�t��
' ���� :
'   strTitileMsg : �_�C�A���O�̃E�C���h�E�^�C�g��������
'   strFilter    : �t�@�C����ʑI���t�B���^�[�i�� : "�e�L�X�g�t�@�C��,*.txt,�S�Ẵt�@�C��,*.*"�j
'   boolExistDelete  : True���w�肷��ƃt�@�C�������݂���ꍇ�́A�폜����
' �߂�l :
'   ����I�� : �t�@�C���̃t���p�X������
'   �L�����Z���{�^�� : "" ������
' *****

Function FileSaveDlg(ByVal strTitleMsg As String, ByVal strFilter As String, ByVal boolExistDelete As Boolean) As String
    FileSaveDlg = ""        ' �߂�l�̏����l�ݒ�
    
    Dim outputFilename As Variant

    ' *****
    ' �������݃t�@�C������I���E���͂���GUI�\��
    ' *****
    outputFilename = Application.GetSaveAsFilename(FileFilter:=strFilter)
    ' OK�������ꂽ�ꍇ�̓t�@�C������String���A�L�����Z���̏ꍇ�� FALSE ���Ԃ�
    If VarType(outputFilename) = vbBoolean Then
        MsgBox ("�L�����Z���{�^����������܂���")
        Exit Function
    End If
    
    ' *****
    ' �o�̓t�@�C�������݂���ꍇ�́A�폜����
    ' *****
    If boolExistDelete = True Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(outputFilename) Then
            If MsgBox("�o�̓t�@�C�������݂��܂��̂ō폜���܂�" & vbCrLf & outputFilename, vbYesNo) = vbYes Then
                Kill (outputFilename)
            Else
                MsgBox ("�����t�@�C�����폜�����ɁA���̐�̏����͏o���܂���")
                Set fso = Nothing
                Exit Function
            End If
        End If
        Set fso = Nothing
    End If
    
    ' �߂�l�i�I�����ꂽ�t�@�C�����j
    FileSaveDlg = outputFilename

End Function

