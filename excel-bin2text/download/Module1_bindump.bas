Attribute VB_Name = "Module1"
Option Explicit
' *****
' �o�C�i���_���v�֘A�F�{�^�����������n���h���T�u���[�`��
' *****

Sub BtnBinToText()

    Dim inputFilename As String
    Dim outputFilename As Variant

    Dim arrFilter() As Variant
    arrFilter = Array("�S�Ẵt�@�C��", "*.*", "�e�L�X�g�t�@�C��", "*.txt", "zip�t�@�C��", "*.zip")
    inputFilename = FileOpenDlg("���̓t�@�C�����̎w��", arrFilter)
    If inputFilename = "" Then
        Exit Sub
    End If

    outputFilename = FileSaveDlg("�o�́i�ۑ��j�t�@�C�����̎w��", "�e�L�X�g�t�@�C��,*.txt,�S�Ẵt�@�C��,*.*", True)
    If outputFilename = "" Then
        Exit Sub
    End If
    
    
    ' *****
    ' �\�[�X�t�@�C����S�ēǂݍ��݁A�o�C�i���z��Ɋi�[����
    ' *****
    Dim fn
    fn = FreeFile
    Dim numFileSize As Long
    Dim byteArray() As Byte
    Open inputFilename For Binary Access Read As #fn
    
    numFileSize = LOF(fn)
    ReDim byteArray(numFileSize)
 
    byteArray = InputB(numFileSize, fn)
    
    Close (fn)
    

    ' *****
    ' �o�C�i���z����A�o�̓t�@�C���i�e�L�X�g�t�@�C���j�ɏ�������
    ' *****
    Dim i As Long
    Dim ch As Byte
    Dim strBuffer As String
    fn = FreeFile   ' �V�����t�@�C���ԍ��𓾂�
    Open CStr(outputFilename) For Output As #fn
    
    For i = 0 To numFileSize - 1
        If i Mod 16 = 0 Then
            ' �o�͍s���́A�t�@�C���擪����̃A�h���X 6��
            strBuffer = Right("000000" & Hex(i), 6)
            Print #fn, strBuffer & " | ";       ' �X�N���v�g�s�� �Z�~�R�����ŉ��s���Ȃ��o��
        End If
        ' �o�C�g�z�񂩂���o���ꂽ�f�[�^��16�i�����񉻂��ăe�L�X�g�t�@�C���ɏ�������
        ch = byteArray(i)
        strBuffer = Right("0" & Hex(ch), 2)     ' �[������2���ɂ���
        Print #fn, strBuffer & " ";             ' �X�N���v�g�s�� �Z�~�R�����ŉ��s���Ȃ��o��
        If (i + 1) Mod 16 = 0 Then
            Print #fn, vbCrLf;
        End If
    Next i
    
    Close (fn)
    Set fn = Nothing
    
    MsgBox ("�����I��" & vbCrLf & "�ϊ��� : " & inputFilename & vbCrLf & "�ϊ��� : " & outputFilename)
    
    
End Sub


Sub BtnTextToBin()
    Dim inputFilename As String
    Dim outputFilename As Variant

    Dim arrFilter() As Variant
    arrFilter = Array("�e�L�X�g�t�@�C��", "*.txt", "�S�Ẵt�@�C��", "*.*")
    inputFilename = FileOpenDlg("���̓t�@�C�����̎w��", arrFilter)
    If inputFilename = "" Then
        Exit Sub
    End If

    outputFilename = FileSaveDlg("�o�́i�ۑ��j�t�@�C�����̎w��", "�S�Ẵt�@�C��,*.*,�e�L�X�g�t�@�C��,*.txt,zip�t�@�C��,*.zip", True)
    If outputFilename = "" Then
        Exit Sub
    End If



    ' *****
    ' ���́E�o�̓t�@�C�����J��
    ' *****
    Dim fnIn
    Dim fnOut
    fnIn = FreeFile
    Open inputFilename For Input As #fnIn
    fnOut = FreeFile
    Open outputFilename For Binary Access Write As #fnOut
    
    ' *****
    ' �\�[�X�t�@�C����1�s���ǂݍ��݂Ȃ���A�o�C�i���ɖ߂��o�̓t�@�C���ɏ�������
    ' *****
    Dim strLine As String
    Dim chStr As Variant
    Dim chByte As Byte
    Dim arrayData As Variant
    Do Until EOF(fnIn)
        Line Input #fnIn, strLine           ' �\�[�X�t�@�C���̃e�L�X�g1�s�ǂݍ���
        arrayData = Split(strLine, "|")     ' �A�h���X�����ƃf�[�^������؂蕪��
        arrayData = Split(CStr(arrayData(1)), " ")    ' �f�[�^��1�o�C�g�����؂蕪��
        For Each chStr In arrayData
            If Len(chStr) >= 1 Then         ' 1�����ȏ�ł���΁iNULL��������X�L�b�v�j
                chByte = Val("&H" & chStr)
                Put #fnOut, , chByte
            End If
        Next
    Loop
    
    Close (fnOut)
    Close (fnIn)
    
    MsgBox ("�����I��" & vbCrLf & "�ϊ��� : " & inputFilename & vbCrLf & "�ϊ��� : " & outputFilename)

End Sub



