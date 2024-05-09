Attribute VB_Name = "ExtructItem"
Option Explicit

''' �Z���������Ԃ�
Public Function CellText(ocell As range) As String
    On Error Resume Next
    If ocell.HasFormula Then
        CellText = Chr(&H22) & ocell.Value2 & Chr(&H22) & "(" & ocell.Formula & ")"
    Else
        CellText = Chr(&H22) & FrameText(ocell) & Chr(&H22)
    End If
    On Error GoTo 0
End Function

''' �R�����g�X���b�h��Ԃ�
Public Function CommentThreadedText(othread As CommentThreaded) As String
    On Error Resume Next

    Dim sthread As String
    sthread = sthread & othread.Date & " " & othread.Author.Name & ":" & Chr(&H22) & othread.Text & Chr(&H22)
    If othread.Replies.Count > 0 Then
        Dim oreply As CommentThreaded
        For Each oreply In othread.Replies
            sthread = sthread & vbLf & CommentThreadedText(oreply)
        Next oreply
    End If

    CommentThreadedText = sthread
    On Error GoTo 0
End Function

''' �}�`���摜�t�@�C���ɏ�������
Public Function ShapeToImageFile(oshape As Shape, imagefile As String) As String
    On Error Resume Next
    ' �摜���N���b�v�{�[�h�ɃR�s�[
    oshape.Copy
    ' �N���b�v�{�[�h���摜�t�@�C���ɏ�������
    ShapeToImageFile = osClipboardToImageFile(imagefile)
    On Error GoTo 0
End Function

''' �}�`�̖��̂�Ԃ�
Public Function ShapeName(oshape As Shape) As String
    On Error Resume Next
    Dim ssharp As String
    ssharp = Chr(&H22) & oshape.Name & Chr(&H22)
    If Len(oshape.AlternativeText) > 0 Then
        ssharp = ssharp & "(" & oshape.AlternativeText & ")"
    End If
    ShapeName = ssharp
    On Error GoTo 0
End Function

' �e�L�X�g�g�𕶎���ŕԂ�
Public Function FrameText(vframe As Variant) As String
    On Error Resume Next
    FrameText = ""
    Dim ncolor As Long
    Dim fBold As Boolean
    Dim fItalic As Boolean
    Dim fStrikethrough As Boolean
    Dim sUnderline As XlUnderlineStyle
    ncolor = &H0
    fBold = False
    fItalic = False
    fStrikethrough = False
    sUnderline = xlUnderlineStyleNone
    Dim svalue As String
    Dim nchara As Integer
    For nchara = 1 To Len(vframe.Characters().Text)

        ' ����̏I��
        If sUnderline <> vframe.Characters(nchara, 1).Font.Underline Then
            If sUnderline <> xlUnderlineStyleNone Then svalue = svalue & "</����>"
        End If
        If fStrikethrough <> vframe.Characters(nchara, 1).Font.Strikethrough Then
            If fStrikethrough Then svalue = svalue & "</��������>"
        End If
        If fItalic <> vframe.Characters(nchara, 1).Font.Italic Then
            If fItalic Then svalue = svalue & "</�Α�>"
        End If
        If fBold <> vframe.Characters(nchara, 1).Font.Bold Then
            If fBold Then svalue = svalue & "</����>"
        End If
        If ncolor <> vframe.Characters(nchara, 1).Font.Color Then
            If ncolor <> &H0 Then svalue = svalue & "</�F>"
        End If

        ' ����̊J�n
        If ncolor <> vframe.Characters(nchara, 1).Font.Color Then
            ncolor = vframe.Characters(nchara, 1).Font.Color
            If ncolor <> &H0 Then svalue = svalue & "<�F:0x" & Hex(ncolor) & ">"
        End If
        If fBold <> vframe.Characters(nchara, 1).Font.Bold Then
            fBold = vframe.Characters(nchara, 1).Font.Bold
            If fBold Then svalue = svalue & "<����>"
        End If
        If fItalic <> vframe.Characters(nchara, 1).Font.Italic Then
            fItalic = vframe.Characters(nchara, 1).Font.Italic
            If fItalic Then svalue = svalue & "<�Α�>"
        End If
        If fStrikethrough <> vframe.Characters(nchara, 1).Font.Strikethrough Then
            fStrikethrough = vframe.Characters(nchara, 1).Font.Strikethrough
            If fStrikethrough Then svalue = svalue & "<��������>"
        End If
        If sUnderline <> vframe.Characters(nchara, 1).Font.Underline Then
            sUnderline = vframe.Characters(nchara, 1).Font.Underline
            Select Case sUnderline
                Case xlUnderlineStyleDouble:            svalue = svalue & "<������d����>"
                Case xlUnderlineStyleDoubleAccounting:  svalue = svalue & "<����2�{�ׂ̍���>"
                Case xlUnderlineStyleNone:              svalue = svalue & ""
                Case xlUnderlineStyleSingle:            svalue = svalue & "<��d����>"
                Case xlUnderlineStyleSingleAccounting:  svalue = svalue & "<��T�|�[�g����>"
                Case Else:                              svalue = svalue & "<�s���ȉ���>"
            End Select
        End If
        svalue = svalue & vframe.Characters(nchara, 1).Text
    Next nchara

    ' ����̏I��
    If sUnderline <> xlUnderlineStyleNone   Then svalue = svalue & "</����>"
    If fStrikethrough                       Then svalue = svalue & "</��������>"
    If fItalic                              Then svalue = svalue & "</�Α�>"
    If fBold                                Then svalue = svalue & "</����>"
    If ncolor <> &H0                        Then svalue = svalue & "</�F>" 
    FrameText = svalue
    On Error GoTo 0
End Function
