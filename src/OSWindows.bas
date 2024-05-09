Attribute VB_Name = "OSWindows"
Option Explicit

''' OS�̖��̂�Ԃ�
Public Function osCanBeExecuted() As String
    If Application.OperatingSystem Like "*Mac*" Then
        osCanBeExecuted = "MacOS�p���g���ĉ������B"
    Else
        osCanBeExecuted = ""
    End If
End Function

''' �t�@�C���p�X�̌���
Public Function osBuildPath(directory As String, filename As String) As String
    osBuildPath = directory & "\" & filename
End Function 

''' �f�B���N�g���̍쐬
Public Function osMakeDirectory(directory As String) As String
    ' �G���[�����̓o�^
    osMakeDirectory = ""
    On Error GoTo ErrorHandler

    MkDir directory

    On Error GoTo 0
    Exit Function
ErrorHandler:
    osMakeDirectory = Err.Number & ":" & Err.Description
    Resume Next
End Function

''' �N���b�v�{�[�h���摜�t�@�C���ɒ��o����
Public Function osClipboardToImageFile(filename As String) As String
    ' �G���[�����̓o�^
    osClipboardToImageFile = ""
    On Error GoTo ErrorHandler

    ' �N���b�v�{�[�h�ɉ摜�������
    'Dim formats() As XlClipboardFormat
    Dim formats As Variant
    formats = Application.ClipboardFormats
    if formats(1) >= 0 Then

        ' �摜�t�@�C�������쐬
        Dim stype As String
        stype = "png"
        Dim imagefile As String
        imagefile = Replace(filename & "." & stype, " ", "_")

        ' �摜�t�@�C���������o��
        Dim wscript As Object
        Set wscript = CreateObject("WScript.Shell")
        Dim scommand As String
        scommand = "powershell Add-Type -AssemblyName System.Windows.Forms;$ImagePath = '" & imagefile & "';  [Windows.Forms.Clipboard]::GetImage().Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::" & stype & ")"
        wscript.Run Command:=scommand, WindowStyle:=0, WaitOnReturn:=True
        Set wscript = Nothing
    End If
    Set formats = Nothing

    On Error GoTo 0
    Exit Function
ErrorHandler:
    osClipboardToImageFile = Err.Number & ":" & Err.Description
    Resume Next
End Function
