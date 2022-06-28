Attribute VB_Name = "Module1"
Option Explicit

' �����{�^���������ꂽ���̏���
Sub ButtonFindAllComment()

    FindCommentCell (False)
End Sub

' �폜�{�^���������ꂽ���̏���
Sub ButtonDeletedAllComment()

    FindCommentCell (True)

End Sub

' ���[�U���w�肷�郏�[�N�u�b�N���A�R�����g�������E�폜����
Function FindCommentCell(ByVal bClearComment As Boolean)
    
    ' �߂�l�i������j�̏�����
    FindCommentCell = ""
    
    ' �f�X�N�g�b�v�̃t���p�X���𓾂�
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = wsh.SpecialFolders("Desktop")
    
    ' �t�@�C���_�C�A���O��\��
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "�R�����g����������t�@�C���̑I��"
        With .Filters
            .Clear
            .Add "Excel���[�N�u�b�N", "*.xlsx; *.xls; *.xlsm", 1
        End With
        ' �t�H���_���� \ �ŏI�����Ă��Ȃ��ƁA�t�@�C�����Ƃ��Ĉ�����ꍇ������Ώ�
        If Right(DesktopFolderName, 1) <> "\" Then
            DesktopFolderName = DesktopFolderName + "\"
        End If
        
        .InitialFileName = DesktopFolderName
    End With
    If fd.Show <> True Then
        MsgBox ("�L�����Z������܂���")
        Exit Function
    End If
    MsgBox ("�Ώۃt�@�C�� " & fd.SelectedItems(1) & " ���I������܂���")

    ' ���ʕ\�����i�[���镶����
    Dim strMsg As String
    strMsg = "�R�����g�̂���Z���ꗗ" & vbCrLf & vbCrLf

    ' �X�e�[�^�X�o�[�̏���������������i���݂̃��[�h�͕ۑ����Ă����j
    Dim boolModeDispStatusbar As Boolean
    boolModeDispStatusbar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Application.StatusBar = fd.SelectedItems(1) & " ���J���Ă��܂� ..."
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    ' �����Ώۃ��[�N�u�b�N
    Dim wb As Workbook
    Set wb = xl.Workbooks.Open(fd.SelectedItems(1))
    Set xl = Nothing
    ' �����Ώۃ��[�N�V�[�g
    Dim ws As Worksheet
    ' Set ws = wb.Worksheets(1)  ' �f�o�b�O�p �P�ڂ̃��[�N�V�[�g�̂�
    
    For Each ws In wb.Worksheets
        Application.StatusBar = ws.Name & " ������ ..."
        strMsg = strMsg & ws.Name & ": " & FindCommentCellOnWorksheet(ws, bClearComment) & vbCrLf
    Next ws
    Application.StatusBar = False
    Application.DisplayStatusBar = boolModeDispStatusbar
    
    If bClearComment = True Then
        MsgBox (strMsg & vbCrLf & "���̃_�C�A���O�ł����̃R�����g���폜���A�V�K�t�@�C���ɕۑ��ł��܂�")
        ' �R�����g�폜�������[�N�u�b�N���A���O��t���ĕۑ�����
        Dim result As Boolean
        result = SaveAsNewfile(wb, fd.SelectedItems(1))
        If result = True Then
            MsgBox ("�V�K�t�@�C���ɕۑ����܂���")
        Else
            MsgBox ("�L�����Z�����܂���")
        End If
        
    Else
        MsgBox (strMsg)
    End If

    ' �����Ώۃ��[�N�u�b�N�����
    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing

End Function

' �w�肵�����[�N�V�[�g���̃R�����g���������Z���A�h���X��Ԃ��B�܂��R�����g���폜����
Function FindCommentCellOnWorksheet(ByVal ws As Worksheet, ByVal bClearComment As Boolean) As String

    ' �߂�l�i������j�̏�����
    FindCommentCellOnWorksheet = ""

    Dim CommentCells As Range

    On Error Resume Next    ' �V�[�g���Ɉ���Y���Z���������ꍇ�A�u�Y������Z����������܂���v�G���[�����
    Set CommentCells = ws.Cells.SpecialCells(xlCellTypeComments)
    On Error GoTo 0
    
    Dim str As String
    If CommentCells Is Nothing Then
        str = ""
    Else
        str = CommentCells.Address
        ' ��΃A�h���X����A���₷���悤�Ɂu$�v����������
        str = Replace(str, "$", "")
        
        ' �R�����g1���ɃA�N�Z�X������@
        'Dim c As Range
        'Dim temp As String
        'For Each c In CommentCells
        '    temp = c.Address & ":" & c.Comment.Text
        'Next c
        
        '�R�����g�̍폜
        If bClearComment = True Then
            CommentCells.ClearComments
        End If
    End If
    
    FindCommentCellOnWorksheet = str

End Function

' ���O��t���ă��[�N�u�b�N��ۑ�����
Function SaveAsNewfile(ByVal wb As Workbook, ByVal filepath As String)
On Error GoTo catch
    ' ���O��t���ĕۑ����� �_�C�A���O��\������
    filepath = Application.GetSaveAsFilename(InitialFileName:=filepath, Title:="�R�����g������̃t�@�C���𖼑O��t���ĕۑ�����")
    If filepath = "False" Then
        SaveAsNewfile = False
        Exit Function
    End If
    ' �t�@�C���ɕۑ�����
    Call wb.SaveAs(filepath)
    SaveAsNewfile = True
    Exit Function
    
catch:
    SaveAsNewfile = False
    Exit Function
End Function
