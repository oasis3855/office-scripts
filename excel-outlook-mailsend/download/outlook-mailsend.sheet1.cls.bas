VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Sub Button_BrowseAttachFile_01()
    Dim of As Variant
    
    ' �Y�t�t�@�C���A�f�o�b�O�o�̓t�@�C���̃t���p�X�����쐬����
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = ShellObject.SpecialFolders("Desktop")
    
    If Mid(DesktopFolderName, 2, 1) = ":" Then ChDrive Left(DesktopFolderName, 1)
    ChDir DesktopFolderName
    of = Application.GetOpenFilename(FileFilter:="Excel�t�@�C��,*.xls;*.xlsx,Word�t�@�C��,*.doc;*.docx,�S�Ẵt�@�C��(*.*),*.*")

    If of <> False Then
        Me.Range("C26") = of
    End If
    

End Sub

Sub Button_BrowseAttachFile_02()
    Dim of As Variant
    
    ' �Y�t�t�@�C���A�f�o�b�O�o�̓t�@�C���̃t���p�X�����쐬����
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = ShellObject.SpecialFolders("Desktop")
    
    If Mid(DesktopFolderName, 2, 1) = ":" Then ChDrive Left(DesktopFolderName, 1)
    ChDir DesktopFolderName
    of = Application.GetOpenFilename(FileFilter:="Excel�t�@�C��,*.xls;*.xlsx,Word�t�@�C��,*.doc;*.docx,�S�Ẵt�@�C��(*.*),*.*")

    If of <> False Then
        Me.Range("C28") = of
    End If
    

End Sub

Sub Button_BrowseAttachFile_03()
    Dim of As Variant
    
    ' �Y�t�t�@�C���A�f�o�b�O�o�̓t�@�C���̃t���p�X�����쐬����
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = ShellObject.SpecialFolders("Desktop")
    
    If Mid(DesktopFolderName, 2, 1) = ":" Then ChDrive Left(DesktopFolderName, 1)
    ChDir DesktopFolderName
    of = Application.GetOpenFilename(FileFilter:="Excel�t�@�C��,*.xls;*.xlsx,Word�t�@�C��,*.doc;*.docx,�S�Ẵt�@�C��(*.*),*.*")

    If of <> False Then
        Me.Range("C30") = of
    End If
    

End Sub

Sub Button_BrowseAttachFile_04()
    Dim of As Variant
    
    ' �Y�t�t�@�C���A�f�o�b�O�o�̓t�@�C���̃t���p�X�����쐬����
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim DesktopFolderName As String
    DesktopFolderName = ShellObject.SpecialFolders("Desktop")
    
    If Mid(DesktopFolderName, 2, 1) = ":" Then ChDrive Left(DesktopFolderName, 1)
    ChDir DesktopFolderName
    of = Application.GetOpenFilename(FileFilter:="Excel�t�@�C��,*.xls;*.xlsx,Word�t�@�C��,*.doc;*.docx,�S�Ẵt�@�C��(*.*),*.*")

    If of <> False Then
        Me.Range("C32") = of
    End If
    

End Sub

