Attribute VB_Name = "ExtructItem"
Option Explicit

''' セル文字列を返す
Public Function CellText(ocell As range) As String
    On Error Resume Next
    If ocell.HasFormula Then
        CellText = Chr(&H22) & ocell.Value2 & Chr(&H22) & "(" & ocell.Formula & ")"
    Else
        CellText = Chr(&H22) & FrameText(ocell) & Chr(&H22)
    End If
    On Error GoTo 0
End Function

''' コメントスレッドを返す
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

''' 図形を画像ファイルに書き込む
Public Function ShapeToImageFile(oshape As Shape, imagefile As String) As String
    On Error Resume Next
    ' 画像をクリップボードにコピー
    oshape.Copy
    ' クリップボードを画像ファイルに書き込む
    ShapeToImageFile = osClipboardToImageFile(imagefile)
    On Error GoTo 0
End Function

''' 図形の名称を返す
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

' テキスト枠を文字列で返す
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

        ' 飾りの終了
        If sUnderline <> vframe.Characters(nchara, 1).Font.Underline Then
            If sUnderline <> xlUnderlineStyleNone Then svalue = svalue & "</下線>"
        End If
        If fStrikethrough <> vframe.Characters(nchara, 1).Font.Strikethrough Then
            If fStrikethrough Then svalue = svalue & "</取り消し線>"
        End If
        If fItalic <> vframe.Characters(nchara, 1).Font.Italic Then
            If fItalic Then svalue = svalue & "</斜体>"
        End If
        If fBold <> vframe.Characters(nchara, 1).Font.Bold Then
            If fBold Then svalue = svalue & "</太字>"
        End If
        If ncolor <> vframe.Characters(nchara, 1).Font.Color Then
            If ncolor <> &H0 Then svalue = svalue & "</色>"
        End If

        ' 飾りの開始
        If ncolor <> vframe.Characters(nchara, 1).Font.Color Then
            ncolor = vframe.Characters(nchara, 1).Font.Color
            If ncolor <> &H0 Then svalue = svalue & "<色:0x" & Hex(ncolor) & ">"
        End If
        If fBold <> vframe.Characters(nchara, 1).Font.Bold Then
            fBold = vframe.Characters(nchara, 1).Font.Bold
            If fBold Then svalue = svalue & "<太字>"
        End If
        If fItalic <> vframe.Characters(nchara, 1).Font.Italic Then
            fItalic = vframe.Characters(nchara, 1).Font.Italic
            If fItalic Then svalue = svalue & "<斜体>"
        End If
        If fStrikethrough <> vframe.Characters(nchara, 1).Font.Strikethrough Then
            fStrikethrough = vframe.Characters(nchara, 1).Font.Strikethrough
            If fStrikethrough Then svalue = svalue & "<取り消し線>"
        End If
        If sUnderline <> vframe.Characters(nchara, 1).Font.Underline Then
            sUnderline = vframe.Characters(nchara, 1).Font.Underline
            Select Case sUnderline
                Case xlUnderlineStyleDouble:            svalue = svalue & "<太い二重下線>"
                Case xlUnderlineStyleDoubleAccounting:  svalue = svalue & "<並んだ2本の細い線>"
                Case xlUnderlineStyleNone:              svalue = svalue & ""
                Case xlUnderlineStyleSingle:            svalue = svalue & "<一重下線>"
                Case xlUnderlineStyleSingleAccounting:  svalue = svalue & "<非サポート下線>"
                Case Else:                              svalue = svalue & "<不明な下線>"
            End Select
        End If
        svalue = svalue & vframe.Characters(nchara, 1).Text
    Next nchara

    ' 飾りの終了
    If sUnderline <> xlUnderlineStyleNone   Then svalue = svalue & "</下線>"
    If fStrikethrough                       Then svalue = svalue & "</取り消し線>"
    If fItalic                              Then svalue = svalue & "</斜体>"
    If fBold                                Then svalue = svalue & "</太字>"
    If ncolor <> &H0                        Then svalue = svalue & "</色>" 
    FrameText = svalue
    On Error GoTo 0
End Function
