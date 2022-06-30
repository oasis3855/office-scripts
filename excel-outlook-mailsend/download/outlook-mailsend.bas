Attribute VB_Name = "Module1"
Option Explicit

' �������냏�[�N�V�[�g�̃f�[�^�X�^�[�g�s�i�擪�s�͕�����u���̃L�[���[�h���ږ��j
Public Const START_ROW As Integer = 2

' �f�o�b�O�F�擪5�s�̎��ۂ̎��s����s��
Public Const DEBUG5LINE_LINE = 5

' �G���[�g���b�v��L��������
Public Const ERROR_TRAP_ENABLE = True

' �Y�t�t�@�C���ҏW�ی�p�X���[�h
Dim strAttachWordPassword As String


Sub Button_SendMail()

    ' �G���[���g���b�v���AVBE�f�o�b�O�_�C�A���O��\���������ɁA�X�N���v�g�������I������
    If ERROR_TRAP_ENABLE = True Then On Error GoTo ERRTRAP_Button_SendMail
    Err.Clear

    Dim wbCurrent As Workbook
    Set wbCurrent = ActiveWorkbook
    Dim wsAddrbook As Worksheet
    Set wsAddrbook = wbCurrent.Worksheets("����")
    Dim wsControl As Worksheet
    Set wsControl = wbCurrent.Worksheets("������")
    
    ' ��������̃J�������A�J����No���i�[�����z��
    Dim arrKey() As String, arrColNo() As Integer
    Call SetArrayColName(arrKey(), arrColNo(), wsAddrbook)
    
    Dim i As Integer, j As Integer

    ' �f�o�b�O�iOutlook�o�͂ł͂Ȃ��A�f�X�N�g�b�v��Ƀ��O�t�@�C�����o�́j
    Dim DEBUG_NO_OUTLOOK As Boolean: DEBUG_NO_OUTLOOK = False
    If wsControl.CheckBoxes("chkbox_debugtxt").Value = xlOn Then DEBUG_NO_OUTLOOK = True
    ' �f�o�b�O�t�@�C���̂��߂̃t�@�C���n���h��
    Dim fn As Integer

    ' �f�o�b�O�o�̓t�@�C���̃t���p�X�����쐬����
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim filepathDebugOutput As String
    filepathDebugOutput = ShellObject.SpecialFolders("Desktop") & "\" & "debug.txt"
    

    ' �f�[�^�����͂���Ă���ő�J�������A�s���𓾂�
    Dim MaxRange As Range
    Set MaxRange = wsAddrbook.UsedRange
    Dim rowsMax As Integer
    Dim colsMax As Integer
    rowsMax = MaxRange.row + MaxRange.Rows.Count - 1
    colsMax = MaxRange.Column + MaxRange.Columns.Count - 1
    

    If DEBUG_NO_OUTLOOK Then
        ' �f�X�N�g�b�v�Ƀ��O�t�@�C�����쐬�i�����t�@�C��������ꍇ�̓[���ɐ؂�̂Ă�j
        fn = FreeFile
        Open filepathDebugOutput For Output As #fn
            Print #fn, "rowsMax = " & rowsMax
            Print #fn, "colsMax = " & colsMax
            Print #fn, "Ubound(arrKey) = " & UBound(arrKey)
            Print #fn, "======"
        Close #fn
    End If


    Dim objOutlook As Outlook.Application
    Set objOutlook = New Outlook.Application
    Dim objMail As Outlook.MailItem


    For i = START_ROW To rowsMax
        If wsControl.CheckBoxes("chkbox_debug5line").Value = xlOn And i >= START_ROW + DEBUG5LINE_LINE Then Exit For
        
        ' Excel��ʉ��̃X�e�[�^�X�o�[�ɁA�������󋵕\�����s��
        Application.StatusBar = (i - START_ROW + 1) & "/" & (rowsMax - START_ROW + 1) & "������..."
        
        '�����ʂ�L4�Z���Łu�������䂷��v�i�񖼕����񂪓��͂���Ă���j�ݒ�ƂȂ��Ă���ꍇ
        '(N4�Z���Ɂu*�v�����͂���Ă���ꍇ�́A�S�Ă̒l�Ɉ�v����ݒ�̂��߁A�������䏈���͓ǂݔ�΂�)
        Dim colControl As Integer: colControl = -1
        If wsControl.Range("L4").Value <> "" And wsControl.Range("N4").Value <> "*" Then
            ' L4�Z���̕����񂪁A����V�[�g�̉���ڂ�����
            For colControl = 0 To UBound(arrKey)
                If arrKey(colControl) = wsControl.Range("L4").Value Then Exit For
            Next colControl
            ' ����V�[�g�́u��������v�Z��������V�[�g�uN4�̒l�v�łȂ��Ȃ�A���ݏ����s�̎c��̏����i���[�����M�j���X�L�b�v����
            If colControl >= 0 And colControl <= UBound(arrColNo) Then
                colControl = arrColNo(colControl)
                If wsAddrbook.Cells(i, colControl) <> wsControl.Range("N4").Value Then GoTo SKIP_TO_NEXT_ROW
            Else
                colControl = -1
            End If
        End If
        
        
        ' ���[���A�h���X�u�Z�Z�Z�Z <xxxx@example.com>�v�A���[���薼�A���[���{���𕶎���u������
        Dim strEmailFull As String
        Dim strSubject As String
        Dim strMailText As String
    
        strEmailFull = wsControl.Range("C4").Value & " <" & wsControl.Range("I4").Value & ">;"
        strEmailFull = ReplaceString(strEmailFull, arrKey, arrColNo, wsAddrbook, i)
    
        strSubject = wsControl.Range("C6").Value
        strSubject = ReplaceString(strSubject, arrKey, arrColNo, wsAddrbook, i)
    
        strMailText = wsControl.Range("C8").Value
        strMailText = ReplaceString(strMailText, arrKey, arrColNo, wsAddrbook, i)
        
    
        ' �f�X�N�g�b�v�Ƀ��O�t�@�C�����쐬
        If DEBUG_NO_OUTLOOK Then
            fn = FreeFile
            Open filepathDebugOutput For Append As #fn
                Print #fn, strEmailFull & ",  " & strSubject
            Close #fn
        
            ' Outlook�̏��������Ȃ��i�f�o�b�O�p�j
            GoTo SKIP_OUTLOOK_PROCESS
        End If
    
        ' Outlook�̃��[���A�C�e���i1�ʕ��j�̃I�u�W�F�N�g���쐬
        Set objMail = objOutlook.CreateItem(olMailItem)
        With objMail
            .To = strEmailFull              ' ���M��A�h���X
            .Subject = strSubject           ' �薼
            .Body = strMailText             ' ���[���{��
            .BodyFormat = olFormatPlain     '���[���̌`��
        End With
        
        ' �Y�t�t�@�C���p�p�X���[�h���O���[�o���ϐ��Ɋi�[
        strAttachWordPassword = wsControl.Range("D35").Value
        
        ' �Y�t�t�@�C���̏����i������u�����ATEMP�t�H���_�Ɉꎞ�i�[���A�Y�t����j
        Dim filepath As String
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim arr() As Variant, rowCell As Variant
        arr = Array(26, 28, 30, 32)     ' �t�@�C�������i�[���ꂽ C26, C28, C30, C32 �Z�������ɏ���
        For Each rowCell In arr
            filepath = wsControl.Cells(rowCell, "C")
            If wsControl.CheckBoxes("chkbox_replace_" & rowCell).Value = xlOn Then
            'If wsControl.Cells(rowCell, "L") = "�u������" Then
                Dim tempfilepath As String
                tempfilepath = ReplaceString_WordExcel(filepath, arrKey, arrColNo, wsAddrbook, i)
                If tempfilepath <> "" Then
                    objMail.Attachments.Add (tempfilepath)
                    fso.DeleteFile (tempfilepath)
                End If
            Else
                If filepath <> "" And Dir(filepath) <> "" Then objMail.Attachments.Add (filepath)
            End If
        Next rowCell
        Set fso = Nothing

        ' ���[����Outlook�́u�������v�t�H���_�ɕۑ�����i���ۂ̑��M�͂��Ȃ��j
        'objMail.Send
        'objMail.Display
        objMail.Save
        Set objMail = Nothing
  
    
SKIP_OUTLOOK_PROCESS:
SKIP_TO_NEXT_ROW:
    Next i
    
    Application.StatusBar = False
    MsgBox ((rowsMax - START_ROW + 1) & "���̃��[����Outlook�u�������v�t�H���_�Ɋi�[���܂���")

ERRTRAP_Button_SendMail:
    If Err.Number Then
        MsgBox ("�G���[(Button_SendMail) : " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "===== �X�N���v�g���I�����܂� =====")
        Err.Clear
        End
    End If


End Sub


'***********************************************************
' ������u���p�ɁA�u�L�[���[�h�v�Ɓu���̃L�[���[�h�������Ă����ԍ��v��z��Ɋi�[����
'***********************************************************
Sub SetArrayColName(ByRef arrKey() As String, ByRef arrColNo() As Integer, wsAddrbook As Worksheet)
    Dim i As Integer
    
    ' �f�[�^�����͂���Ă���ő�J�������𓾂�
    Dim MaxRange As Range
    Set MaxRange = wsAddrbook.UsedRange
    Dim colsMax As Integer
    colsMax = MaxRange.Column + MaxRange.Columns.Count - 1
    
    Erase arrKey
    Erase arrColNo
    ' 1�s�ڂ̃J��������z��Ɋi�[
    For i = 1 To colsMax
        If wsAddrbook.Cells(1, i) <> "" And Left(wsAddrbook.Cells(1, i), 1) <> "��" Then
            If IsArrayEx(arrKey) = 0 Then
                ' �z�񂪋�̏ꍇ�́A�ŏ��̃G�������g��ǉ�
                ReDim arrKey(0)
                ReDim arrColNo(0)
            Else
                ' �z�񂪋�łȂ��ꍇ�́A�G�������g��1�ǉ�
                ReDim Preserve arrKey(UBound(arrKey) + 1)
                ReDim Preserve arrColNo(UBound(arrColNo) + 1)
            End If
            arrKey(UBound(arrKey)) = wsAddrbook.Cells(1, i)
            arrColNo(UBound(arrColNo)) = i
        End If
    Next i

End Sub


'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************
Private Function IsArrayEx(arr As Variant) As Long
On Error GoTo ERROR_ARRAY_EMPTY

    If IsArray(arr) Then
        IsArrayEx = IIf(UBound(arr) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_ARRAY_EMPTY:
    ' �z�񂪋�̏ꍇ�AUBound(arr)���G���[�ƂȂ�̂��g���b�v
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function


'***********************************************************
' ������u��
' ����   : str �u�����̕�����AarrKey() �u����������AarrColNo() �u���敶����̗�ԍ��AwsAddrbook �A�h���X�����[�N�V�[�g�Arow �A�h���X���̑Ώۍs
' �߂�l : �u����̕�����
'***********************************************************
Private Function ReplaceString(str As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString = str    ' �߂�l�i�����l�j
    Dim i As Integer
    
    ' �f�[�^�����͂���Ă���ő�J�������A�s���𓾂�
    Dim MaxRange As Range
    Set MaxRange = wsAddrbook.UsedRange
    Dim rowsMax As Integer
    rowsMax = MaxRange.row + MaxRange.Rows.Count - 1
    
    For i = 0 To UBound(arrKey)
        ' ������u��
        str = Replace(str, "��" & arrKey(i) & "��", wsAddrbook.Cells(row, arrColNo(i)))
    Next i
    
    ReplaceString = str    ' �߂�l

End Function

 
'***********************************************************
' ������u���iWord, Excel �������ʁj
' ����   : filepath �u�����̑Ώۃt�@�C���AarrKey() �u����������AarrColNo() �u���敶����̗�ԍ��AwsAddrbook �A�h���X�����[�N�V�[�g�Arow �A�h���X���̑Ώۍs
' �߂�l : �u����̃t�@�C�����iTEMP�t�H���_���̃t�@�C�� �t���p�X�j
'***********************************************************
Private Function ReplaceString_WordExcel(filepath As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString_WordExcel = ""    ' �߂�l
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �w�肳�ꂽ�t�@�C�������݂��Ȃ��A�t�@�C������ "" �̏ꍇ
    If filepath = "" Or Dir(filepath) = "" Then
        Set fso = Nothing
        Exit Function
    End If

    ' ���[�U�ꎞ�f�B���N�g�� & �t�@�C����
    ReplaceString_WordExcel = fso.GetSpecialFolder(2) & "\" & fso.GetFileName(filepath)
    ' �Y�t�t�@�C�����̕�����u���i���悲�ƂɓY�t�t�@�C������ύX����ꍇ�j
    ReplaceString_WordExcel = ReplaceString(ReplaceString_WordExcel, arrKey, arrColNo, wsAddrbook, row)
    
    If fso.GetExtensionName(filepath) = "doc" Or fso.GetExtensionName(filepath) = "docx" Then
        ' Word�������̕�����u��
        ReplaceString_WordExcel = ReplaceString_Word(filepath, ReplaceString_WordExcel, arrKey, arrColNo, wsAddrbook, row)
    ElseIf fso.GetExtensionName(filepath) = "xls" Or fso.GetExtensionName(filepath) = "xlsx" Then
        ' Excel�������̕�����u��
        ReplaceString_WordExcel = ReplaceString_Excel(filepath, ReplaceString_WordExcel, arrKey, arrColNo, wsAddrbook, row)
    End If
    
    Set fso = Nothing
End Function

Private Function ReplaceString_Word(filepath As String, outputfilepath As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString_Word = outputfilepath     ' �߂�l

    ' �G���[���g���b�v���AVBE�f�o�b�O�_�C�A���O��\���������ɁA�X�N���v�g�������I������
    If ERROR_TRAP_ENABLE = True Then On Error GoTo ERRTRAP_ReplaceString_Word
    Err.Clear

    Dim objWord As Word.Application
    Dim objWordDoc As Word.Document
    ' �f�X�N�g�b�v�ɕۑ�����Ă���Word�e���v���[�g�t�@�C�����J��
    Set objWord = New Word.Application
    Set objWordDoc = objWord.Documents.Open(filepath)
    ' �V�X�e���������ł���܂�2�b�҂i�]�T�����Ă���j
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' �u�Z�{�]�ҏW�̐����v�̕ی����������
    Dim flagProtected As Boolean: flagProtected = False
    If objWordDoc.ProtectionType <> wdNoProtection Then
        objWordDoc.Unprotect (strAttachWordPassword)
        flagProtected = True    ' �u���O��t���ĕۑ��v���ɕی�L�������邩�ǂ����̂��߂ɎQ��
    End If
    
    
    Dim i As Integer
    For i = 0 To UBound(arrKey)
        ' �{�����̕�����u��
        With objWordDoc.Content.Find
            .Text = "��" & arrKey(i) & "��"
            .Forward = True
            .Replacement.Text = wsAddrbook.Cells(row, arrColNo(i))
            .Wrap = wdFindContinue
            .MatchFuzzy = True
            .Execute Replace:=wdReplaceAll
        End With
        ' �{�����̃n�C�p�[�����N�i���[���j��Subject��������u��
        Dim hyp As Object
        For Each hyp In objWordDoc.Hyperlinks
            If hyp.Address Like "mailto*" Then
                hyp.EmailSubject = Replace(hyp.EmailSubject, "��" & arrKey(i) & "��", wsAddrbook.Cells(row, arrColNo(i)))
            End If
        Next hyp
        ' �}�`�e�L�X�g�{�b�N�X���̕�����u��
        Dim shp As Word.Shape
        For Each shp In objWordDoc.Shapes
            If shp.Type = msoTextBox Then
                With shp.TextFrame.TextRange.Find
                    .ClearFormatting
                    .Forward = True
                    .Text = "��" & arrKey(i) & "��"
                    .Replacement.Text = wsAddrbook.Cells(row, arrColNo(i))
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        Next shp
        Set shp = Nothing
    Next i
    
    ' ���t�@�C�����ی�L���������ꍇ�A���O��t���ĕۑ�����ꍇ���ی�L��������
    If flagProtected = True Then
        objWordDoc.Protect Type:=wdAllowOnlyReading, Password:=strAttachWordPassword
    End If
    
    ' Word�h�L�������g�𖼑O��t���ĕۑ�����
    objWordDoc.SaveAs (outputfilepath)
    objWordDoc.Close SaveChanges:=False
    objWord.Quit
    
ERRTRAP_ReplaceString_Word:
    If Err.Number Then
        MsgBox ("�G���[(ReplaceString_Word) : " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "===== �X�N���v�g���I�����܂� =====")
        Err.Clear
        End
    End If
    
    Set objWordDoc = Nothing
    Set objWord = Nothing

End Function


Private Function ReplaceString_Excel(filepath As String, outputfilepath As String, arrKey() As String, arrColNo() As Integer, wsAddrbook As Worksheet, row As Integer) As String
    ReplaceString_Excel = outputfilepath     ' �߂�l

    ' �G���[���g���b�v���AVBE�f�o�b�O�_�C�A���O��\���������ɁA�X�N���v�g�������I������
    If ERROR_TRAP_ENABLE = True Then On Error GoTo ERRTRAP_ReplaceString_Excel
    Err.Clear
    
    ' ��ʂɕ\�������Ȃ����߁A�V����Excel�̃I�u�W�F�N�g���쐬���A���̒���Excel�u�b�N���J��
    Dim objExcel As Excel.Application
    Set objExcel = New Excel.Application
    ' �����Ώە����񂪑��݂��Ȃ������ꍇ�́AWarning�_�C�A���O�̏o����}�~����
    objExcel.DisplayAlerts = False
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = objExcel.Workbooks.Open(filepath, ReadOnly:=True)
    ' �V�X�e���������ł���܂�2�b�҂i�]�T�����Ă���j
    Application.Wait (Now + TimeValue("0:00:02"))


    For Each ws In wb.Worksheets
    
        Dim i As Integer
        For i = 0 To UBound(arrKey)
            ' �{�����̕�����u��
'            ws.Cells.Replace What:="��" & arrKey(i) & "��", Replacement:=wsAddrbook.Cells(row, arrColNo(i)), LookAt:=xlPart, _
'                    SearchOrder:=xlByRows, MatchCase:=True, MatchByte:=True, SearchFormat:=False, ReplaceFormat:=False
            ws.UsedRange.Replace What:="��" & arrKey(i) & "��", Replacement:=wsAddrbook.Cells(row, arrColNo(i)), LookAt:=xlPart
'            Dim EachCell As Range
'            For Each EachCell In ws.UsedRange
'                EachCell = Replace(EachCell, "��" & arrKey(i) & "��", wsAddrbook.Cells(row, arrColNo(i)))
'            Next EachCell
            
        Next i
    
    Next ws
    On Error GoTo 0

    wb.SaveAs (outputfilepath)
    wb.Close SaveChanges:=False
    objExcel.Quit
    
ERRTRAP_ReplaceString_Excel:
    If Err.Number Then
        MsgBox ("�G���[(ReplaceString_Excel) : " & vbCrLf & Err.Description & vbCrLf & vbCrLf & "===== �X�N���v�g���I�����܂� =====")
        Err.Clear
        End
    End If

    Set wb = Nothing
    objExcel.DisplayAlerts = True   ' Warning�_�C�A���O�}�~�ݒ�����ɖ߂��Ă���
    Set objExcel = Nothing

End Function

