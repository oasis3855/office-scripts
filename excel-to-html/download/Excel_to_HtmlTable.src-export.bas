Attribute VB_Name = "Module1"
Option Explicit

Sub Convert2HtmlTable_EntryPoint()
    Dim wb As Workbook
    Dim workbookName() As String
    Dim strMsg As String: strMsg = ""
    
    ' Ubound���v�f0�ŃG���[�ƂȂ�΍�Ƃ��āA�_�~�[�ōŏ��̂P���쐬
    ReDim workbookName(0)
    
    ' �N������Excel���[�N�u�b�N���𓮓I�z��Ɋi�[
    For Each wb In Workbooks
        ReDim Preserve workbookName(UBound(workbookName) + 1)
        workbookName(UBound(workbookName)) = wb.Name
        ' InputBox���b�Z�[�W�\���p������
        strMsg = strMsg & UBound(workbookName) & " : " & wb.Name & vbCrLf
    Next wb
    
    ' ���[�U�Ƀ��[�N�u�b�N����I��������
    Dim userSelect As String
    userSelect = InputBox("No : ���[�N�u�b�N��" & vbCrLf & _
        "------------------" & vbCrLf & _
        strMsg, "�Ώۂ̃��[�N�u�b�NNo��I��", "1")

    ' Input�_�C�A���O�ŃL�����Z���܂��͒����[���̕����񂪓��͂��ꂽ�ꍇ
    If StrPtr(userSelect) = 0 Or Len(userSelect) <= 0 Then
        MsgBox ("�L�����Z�����܂���")
        Exit Sub
    End If

    ' �S�p�p�����𔼊p�ɕϊ�
    userSelect = StrConv(userSelect, vbNarrow)

    ' ���͒l�́A���l�A�����[�N�u�b�N��No�i1,2,3,...�j�͈͓�
    If IsNumeric(userSelect) = False Or Val(userSelect) <= 0 Or Val(userSelect) > UBound(workbookName) Then
        MsgBox ("�͈͊O�����͂���܂���")
        Erase workbookName
        Exit Sub
    End If

    ' �w�肳�ꂽ���[�N�u�b�N��ϐ��Ƃ��ēn���āA�ϊ��T�u���[�`�����Ăяo��
    Set wb = Workbooks(workbookName(userSelect))
    Convert2HtmlTable wb

    Erase workbookName

End Sub

' Convert2HtmlTable : ���[�N�V�[�g�\��HTML table�t�@�C���ɕϊ��E�G�N�X�|�[�g����
'
' ���� wb : �Ώۃ��[�N�u�b�N�i�ϊ������̂́A���̃��[�N�u�b�N�̕\�����V�[�g�j
Sub Convert2HtmlTable(wb As Workbook)
    
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet
    
    Dim MaxRange As Range
    Set MaxRange = ws.UsedRange
    Dim maxRow As Integer
    Dim maxCol As Integer
    maxRow = MaxRange.Rows.Count
    maxCol = MaxRange.Columns.Count

    Dim col As Integer
    Dim row As Integer
    ' �c�E������������Ă���s�E�񐔁i�c�����̓r���s�̏ꍇ�́AmergeRows=-1�j
    Dim mergeRows As Integer
    Dim mergeCols As Integer
    
    ' HTML�t�@�C���ɏ������ޓ��e���ꎞ�ۑ�����
    Dim HtmlString As String
    
    ' HTML�w�b�_����
    HtmlString = "<html>" & vbCrLf & _
        "<head>" & vbCrLf & _
        "    <style>" & vbCrLf & _
        "    table { border-collapse: collapse; }" & vbCrLf & _
        "    td { border: 1px solid black; }" & vbCrLf & _
        "    </style>" & vbCrLf & _
        "</head>" & vbCrLf & _
        "<body>" & vbCrLf & _
        "<table>" & vbCrLf
    
    
    For row = MaxRange.row To MaxRange.row + maxRow - 1
        HtmlString = HtmlString & "    <tr>" & vbCrLf
        For col = MaxRange.Column To MaxRange.Column + maxCol - 1
            mergeRows = 1
            mergeCols = 1
            ' �w��Z�����������ɂ���ꍇ�A����Range�𓾂�
            Dim mergeArea As Range
            Set mergeArea = ws.Cells(row, col).mergeArea
            
            If ws.Cells(row, col).MergeCells = True Then
                ' ���Z����������Ă���ꍇ�A��������mergeRows�ɑ���i�c�̂݌����̏ꍇ�͉����������̏ꍇ�Ɠ���1�ƂȂ�j
                mergeCols = mergeArea.Item(mergeArea.Count).Column - col + 1
            End If
            If ws.Cells(row, col).MergeCells = True Then
                ' �c�Z����������Ă���ꍇ�A��������mergeRows�ɑ���i���̂݌����̏ꍇ�͏c���������̏ꍇ�Ɠ���1�ƂȂ�j
                mergeRows = mergeArea.Item(mergeArea.Count).row - row + 1
                If mergeArea.Item(1).row <> row Then
                    ' �c�����ōŏ��̍s�łȂ��ꍇ
                    mergeRows = -1
                End If
            End If
            
            ' td�^�O��rowspan,colspan���ꎞ������extraTdTag�Ɋi�[
            Dim extraTdTag As String
            extraTdTag = ""
            If mergeCols > 1 Then
                ' ����������Ă���ꍇ
                extraTdTag = " colspan = " & mergeCols
            End If
            If mergeRows > 1 Then
                ' �c��������Ă���ꍇ
                extraTdTag = extraTdTag & " rowspan = " & mergeRows
            End If
            
            If mergeRows > 0 Then
                ' �ʏ�̃Z���i�c��������Ă��Ȃ��j
                HtmlString = HtmlString & "        <td " & extraTdTag & ">"
                HtmlString = HtmlString & ws.Cells(row, col)
                HtmlString = HtmlString & "</td>" & vbCrLf
            Else
                ' �c��������Ă���ꍇ�ŁA2�s�ڈȍ~
                HtmlString = HtmlString & "        <!-- rowspan -->" & vbCrLf
            End If
            
            
            If mergeCols > 1 Then
                ' ����������Ă���ꍇ�AFor���[�v(col)���΂�
                col = col + mergeCols - 1
            End If
            
        Next col
        HtmlString = HtmlString & "    </tr>" & vbCrLf
    Next row
    
    ' HTML�t�b�^
    HtmlString = HtmlString & "</table>" & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf

    ' ���o��
    MsgBox ("���[�N�u�b�N : " & wb.Name & vbCrLf & _
            "�ϊ��͈� : " & ConvR1c1ToA1("R" & MaxRange.row & "C" & MaxRange.Column) & " : " & _
            ConvR1c1ToA1("R" & MaxRange.row + maxRow - 1 & "C" & MaxRange.Column + maxCol - 1) & vbCrLf & vbCrLf)
        

    ' �f�X�N�g�b�v �f�B���N�g���𓾂�
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
            
    ' SavAs�_�C�A���O�ŕ\�����鏉���f�B���N�g���Ɉړ�����
    ChDir (ShellObject.SpecialFolders("Desktop"))
    Set ShellObject = Nothing
    
    ' SaveAs�_�C�A���O�\��
    Dim OutputFilename As Variant
    OutputFilename = Application.GetSaveAsFilename("output.html", _
            "HTML�t�@�C��,*.html,�e�L�X�g�t�@�C��,*.txt")
    If OutputFilename = False Then
        MsgBox ("�L�����Z�����܂���")
        Exit Sub
    End If
    
    ' �t�@�C�������ɑ��݂���ꍇ�́A�㏑���x������
    If Dir(OutputFilename) <> "" Then
        If Not (MsgBox("�t�@�C�� " & Dir(OutputFilename) & " �ɏ㏑�����܂�", vbYesNo) = vbYes) Then
            MsgBox ("�L�����Z�����܂���")
            Exit Sub
        End If
    End If
    
    ' �t�@�C���ɏ�������
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(OutputFilename, ForWriting, True, TristateTrue)
    ts.write (HtmlString)
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing

    MsgBox (OutputFilename & " �ɏ������݂܂���")


End Sub

' �uR1C1�`���̕�����v���uA1�`���̕�����v�ɕϊ�
Function ConvR1c1ToA1(StrR1C1 As String)
    ConvR1c1ToA1 = Application.ConvertFormula( _
        Formula:=StrR1C1, _
        fromReferenceStyle:=xlR1C1, _
        toreferencestyle:=xlA1, _
        toabsolute:=xlRelative)
End Function


