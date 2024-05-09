Attribute VB_Name = "OSMac"
Option Explicit

''' OS�̖��̂�Ԃ�
Public Function osCanBeExecuted() As String
    If Application.OperatingSystem Like "*Mac*" Then
        osCanBeExecuted = ""
    Else
        osCanBeExecuted = "Windows�p���g���ĉ������B"
    End If
End Function

''' �t�@�C���p�X�̌���
Public Function osBuildPath(directory As String, filename As String) As String
    osBuildPath = directory & "/" & filename
End Function 

''' �f�B���N�g���̍쐬
Public Function osMakeDirectory(directory As String) As String
    ' �G���[�����̓o�^
    osMakeDirectory = ""
    On Error GoTo ErrorHandler

    osMakeDirectory = AppleScriptTask("excel-extructor.applescript", "MakeDirectory", directory)

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
        osClipboardToImageFile = AppleScriptTask("excel-extructor.applescript", "ClipboardToImageFile", imagefile)
    End If
    Set formats = Nothing

    On Error GoTo 0
    Exit Function
ErrorHandler:
    osClipboardToImageFile = Err.Number & ":" & Err.Description
    Resume Next
End Function
