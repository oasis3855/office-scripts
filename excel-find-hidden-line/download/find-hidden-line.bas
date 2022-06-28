Attribute VB_Name = "Module1"
Option Explicit

Sub FindHiddenLine()
    
    ' �f�X�N�g�b�v�̃t���p�X���𓾂�
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = wsh.SpecialFolders("Desktop")
    
    ' �t�@�C���_�C�A���O��\��
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "��\���s�E�����������t�@�C���̑I��"
        With .Filters
            .Clear
            .Add "Excel���[�N�u�b�N", "*.xlsx; *.xls; *.xlsm", 1
        End With
        .InitialFileName = DesktopFolderName
    End With
    If fd.Show <> True Then
        MsgBox ("�L�����Z������܂���")
        Exit Sub
    End If
    MsgBox ("�Ώۃt�@�C�� " & fd.SelectedItems(1) & " ���I������܂���")

    ' ���ʕ\�����i�[���镶����
    Dim strMsg As String
    strMsg = ""


    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    ' �����Ώۃ��[�N�u�b�N
    Dim wb As Workbook
    Set wb = xl.Workbooks.Open(fd.SelectedItems(1))
    Set xl = Nothing
    ' �����Ώۃ��[�N�V�[�g
    Dim ws As Worksheet
    ' Set ws = wb.Worksheets(1)  ' �f�o�b�O�p �P�ڂ̃��[�N�V�[�g�̂�
    
    Dim boolModeDispStatusbar As Boolean
    boolModeDispStatusbar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    For Each ws In wb.Worksheets
        Application.StatusBar = ws.Name & " ������ ..."
        strMsg = strMsg & ws.Name & ": " & FindHiddenLineOnWorksheet(ws) & vbCrLf
    Next ws
    Application.StatusBar = False
    Application.DisplayStatusBar = boolModeDispStatusbar
    
    ' �����Ώۃ��[�N�u�b�N�����
    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing

    MsgBox (strMsg)

End Sub


Function FindHiddenLineOnWorksheet(ByVal ws As Worksheet) As String

    ' �߂�l�i������j�̏�����
    FindHiddenLineOnWorksheet = ""
    Dim i As Integer
    
    Dim maxrow As Integer
    ' ws.UsedRange.Row : �f�[�^���n�܂�ŏ��̍s
    ' ws.UsedRange.Rows.Count : �f�[�^�����͂���Ă���s�͈͂̍s��
    maxrow = ws.UsedRange.Rows.Count + ws.UsedRange.Row
    
    For i = 1 To maxrow
        If ws.Rows(i).Hidden = True Then
            ' MsgBox ("Row = " & i & " hidden")
            FindHiddenLineOnWorksheet = FindHiddenLineOnWorksheet & i & " �s��,"
        End If
    Next i

    Dim maxcol As Integer
    maxcol = ws.UsedRange.Columns.Count + ws.UsedRange.Column
    
    For i = 1 To maxrow
        If ws.Columns(i).Hidden = True Then
            ' MsgBox ("Col = " & i & " hidden")
            ' FindHiddenLineOnWorksheet = FindHiddenLineOnWorksheet & i & " ���,"
            FindHiddenLineOnWorksheet = FindHiddenLineOnWorksheet & Split(Cells(1, i).Address(True, False), "$")(0) & " ���,"
        End If
    Next i
End Function
