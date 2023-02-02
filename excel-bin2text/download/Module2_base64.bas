Attribute VB_Name = "Module2"
Option Explicit
' *****
' Base64�֘A�F�{�^�����������n���h���T�u���[�`���A�ϊ��֐�
' *****


Sub BtnBase64Encode()
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
    ' �w�肳�ꂽ�t�@�C����ǂݍ��݁ABase64�ɃG���R�[�h���Astring�ϐ��Ƃ��ĕԂ�
    ' *****
    Dim strBase64 As String
    strBase64 = EncodeBase64(inputFilename)
    
    
    ' *****
    ' string�ϐ��iBase64�ɃG���R�[�h����������j���e�L�X�g�t�@�C���ɏ�������
    ' *****
    Dim fn As Variant
    fn = FreeFile
    Open CStr(outputFilename) For Output As #fn
    
    Print #fn, strBase64
    
    Close (fn)

    MsgBox ("�����I��" & vbCrLf & "�ϊ��� : " & inputFilename & vbCrLf & "�ϊ��� : " & outputFilename)


End Sub

Sub BtnBase64Decode()
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
    ' �\�[�X�t�@�C���iBase64�ŃG���R�[�h���ꂽ�t�@�C���j���A��C��String�ϐ��ɓǂݍ���
    ' *****
    Dim strBase64 As String
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(inputFilename).OpenAsTextStream
            strBase64 = .ReadAll
            .Close
        End With
    End With
    
    ' *****
    ' Base64�ŃG���R�[�h���ꂽ������𕜍������A�t�@�C���ɏ�������
    ' *****
    Dim retCode As Long
    retCode = DecodeBase64(strBase64, outputFilename)

    MsgBox ("�����I��" & vbCrLf & "�ϊ��� : " & inputFilename & vbCrLf & "�ϊ��� : " & outputFilename)

End Sub


' *****
' �w�肵���t�@�C������ǂݍ��񂾃f�[�^���ABase64�ϊ������������Ԃ�
' ���� :
'   inputFilename : �ǂݍ��ݑΏۂ̃t�@�C����
' �߂�l :
'   Base64�ϊ������e�L�X�g������
' ���쌠
'   ���̊֐��́Ahttps://www.ka-net.org/blog/?p=4479 ���]�ڂ���
' *****
Private Function EncodeBase64(ByVal inputFilename As String) As String
    '�t�@�C����Base64�G���R�[�h
    Dim elm As Object
    Dim ret As String
    Const adTypeBinary = 1
    Const adReadAll = -1
    
    ret = ""    ' �߂蕶����̏�����
    On Error Resume Next
    Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .LoadFromFile inputFilename
        elm.DataType = "bin.base64"
        elm.nodeTypedValue = .read(adReadAll)
        ret = elm.Text
        .Close
    End With
    On Error GoTo 0
    
    EncodeBase64 = ret
End Function


' *****
' �^����ꂽ�������Base64���畜�������A�w�肵���t�@�C���ɏ�������
' ���� :
'   strBase64       : Base64�ŃG���R�[�h���ꂽ������i���ꂪ�������ΏۂƂȂ�j
'   outputFilename  : �����������f�[�^���������ރt�@�C����
' �߂�l :
'   Long�l
' ���쌠
'   ���̊֐��́Ahttps://www.ka-net.org/blog/?p=4479 ���]�ڂ���
' *****
Private Function DecodeBase64(ByVal strBase64 As String, ByVal outputFilename As String) As Long
    Dim elm As Object
    Dim ret As Long
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    
    Dim strVar As String
    
    ret = -1    ' �߂�l������
    On Error Resume Next
    Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
    elm.DataType = "bin.base64"
    elm.Text = strBase64
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        ' �ϊ����ʂ� elm.nodeTypedValue �z��(Variant)�Ɋi�[����Ă���
        .Write elm.nodeTypedValue
        .SaveToFile outputFilename, adSaveCreateOverWrite
        .Close
    End With
    If Err.Number <> 0 Then ret = 0
    On Error GoTo 0
    DecodeBase64 = ret
End Function
