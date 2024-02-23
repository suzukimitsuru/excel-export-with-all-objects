VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtExtruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' 抽出ボタンのクリックイベント
Sub cmdExtruct_Click()
    ' エラー処理の登録
    Dim errNumber As Integer
    On Error GoTo ErrorHandler

    ' ファイル拡張子が在ったら
    Dim filename As String
    filename = Me.Cells(2, 3).Value
    Dim oprogress As Progress
    Set oprogress = New Progress
    oprogress.SetView Me.Cells(3, 3)
    
    ' このOSで実行出来なかったら、中断する
    Dim serror As String
    serror = osCanBeExecuted()
    If Len(serror) > 0 Then
        oprogress.ErrorView serror
    Else
        ' Excelブックが開けたら
        errNumber = 0
        Dim in_book As Workbook
        Set in_book = Workbooks.Open(filename:=filename, ReadOnly:=True)
        If errNumber <= 0 Then
            Me.Activate

            ' シートの内部名称が取得出来なかったら
            If in_book.Sheets(1).CodeName = "" Then
                oprogress.ErrorView "開発リボンの Visual Basic を開いて下さい。"
            Else
                ' ファイル拡張子の位置を調べる
                Dim ext_position As Long
                ext_position = InStrRev(filename, ".")
                If Len(ext_position) <= 0 Then
                    ext_position = Len(filename)
                End If
                Debug.Print "filename:" & filename & ":" & ext_position

                ' 出力ディレクトリを作成する
                Dim directory As String
                directory = Left(filename, ext_position - 1)
                Dim errMessage As String
                errMessage = osMakeDirectory(directory)
                Debug.Print "osMakeDirectory(" & directory & "):" & errMessage

                ' 出力するテキストファイルが開けたら
                Dim textfile As String
                textfile = osBuildPath(directory, "export.txt")
                Debug.Print "textfile:" & textfile
                serror = oprogress.FileOpen(textfile)
                If Len(serror) > 0 Then
                    oprogress.ErrorView("出力するテキストファイルが開けませんでした。" & serror)
                Else
                    ' 全てのシートの経過の数を集計する
                    Dim osheet As Worksheet
                    For Each osheet In in_book.Sheets
                        oprogress.FilePrint osheet.CodeName & ":" & osheet.Name & " Cells:" & osheet.UsedRange.Count & " Shapes:" & osheet.UsedRange.Count
                        oprogress.ProgressSumCount osheet.UsedRange.Count, osheet.Shapes.Count
                    Next osheet

                    ' 全てのシートを抽出する
                    For Each osheet In in_book.Sheets
                        errNumber = ExtructSheet(oprogress, osheet, directory)
                        oprogress.ProgressCountupSheets
                    Next osheet
                    oprogress.ProgressView

                    ' 出力したテキストファイルを閉じる
                    serror = oprogress.FileClose()
                End If
            End If

            ' Excelブックを閉じる
            in_book.Close(False)
            Set in_book = Nothing
        Else
            oprogress.ErrorView("Excel ブックが開けませんでした。")
        End If
    End If
    
    Set oprogress = Nothing
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    errNumber = Err.Number
    If oprogress.FileIsOpend Then
        oprogress.FilePrint("cmdExtruct_Click() ERROR " & Err.Number & ":" & Err.Description)
    Else
        Debug.Print("cmdExtruct_Click() ERROR " & Err.Number & ":" & Err.Description)
    End If
    Resume Next
End Sub

''' シートの抽出
Function ExtructSheet(oprogress As Progress, osheet As Worksheet, directory As String) As Integer
    Dim errNumber As Integer
    On Error GoTo ErrorHandler

    ' 使用領域が在ったら
    If osheet.UsedRange.Count > 0 Then
        oprogress.FilePrint "--- " & osheet.CodeName & ":" & osheet.Name & " セル ---"

        ' 全ての使用領域で
        Dim nrow As Integer
        For nrow = osheet.UsedRange.Row To osheet.UsedRange.Rows.Count
            Dim ncol As Integer
            For ncol = osheet.UsedRange.Column To osheet.UsedRange.Columns.Count

                ' セルを抽出
                Dim ocell As range
                Set ocell = osheet.Cells(nrow, ncol)
                errNumber = ExtructCell(oprogress, ocell)
                Set ocell = Nothing
                oprogress.ProgressCountupCells
            Next ncol
            oprogress.ProgressView
            DoEvents
        Next nrow
    End If
    
    ' 図形が在ったら
    If osheet.Shapes.Count > 0 Then
        oprogress.FilePrint "--- " & osheet.CodeName & ":" & osheet.Name & " 図形 ---"

        ' 全ての図形の抽出
        Dim oshape As Shape
        For Each oshape In osheet.Shapes
            Dim ssaddress As String
            ssaddress = oshape.Parent.CodeName & "!" & ExtructShapeName(oshape)

            ' 図形のテキストが無かったら
            Dim sshape As String
            sshape = ExtructShapeText(oshape, "")
            If Len(sshape) <= 0 Then
                ' 画像をクリップボードにコピー
                oshape.Copy
                ' クリップボードを画像ファイルに書き込む
                Dim imagefile As String
                imagefile = osBuildPath(directory, oshape.Parent.CodeName & "!" & oshape.Name)
                Dim errMessage As String
                errMessage = osClipboardToImageFile(imagefile)
                if Len(errMessage) > 0 Then
                    oprogress.FilePrint ssaddress & " ERROR " & errMessage
                End If
                oprogress.FilePrint ssaddress
            Else
                oprogress.FilePrint oshape.Parent.CodeName & "!" & sshape
            End If
            oprogress.ProgressCountupShapes
            oprogress.ProgressView
            DoEvents
        Next oshape
    End If
    
    ExtructSheet = errNumber
    On Error GoTo 0
    Exit Function
ErrorHandler:
    errNumber = Err.Number
    oprogress.FilePrint "ExtructSheet() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function

''' セルの抽出
Function ExtructCell(oprogress As Progress, ocell As range) As Integer
    Dim errNumber As Integer
    On Error GoTo ErrorHandler

    ' セルの内容を抽出する
    Dim saddress As String
    saddress = ocell.Parent.CodeName & "!" & ocell.Address(False, False)
    If ocell.HasFormula Then
        oprogress.FilePrint saddress & " " & Chr(&H22) & ocell.Value2 & Chr(&H22) & "(" & ocell.Formula & ")"
    Else
        oprogress.FilePrint saddress & " " & Chr(&H22) & ExtructFrameText(ocell) & Chr(&H22)
    End If

    ' コメントを抽出する
    Dim ocomment As Comment
    Set ocomment = ocell.Comment
    If Not ocomment Is Nothing Then
        oprogress.FilePrint saddress & " " & Chr(&H22) & ocomment.Shape.Name & Chr(&H22) & _
            " " & ocomment.Author & ":" & Chr(&H22) & ocomment.Text & Chr(&H22)
    End If
    Set ocomment = Nothing

    ' スレッドを抽出する
    Dim othread As CommentThreaded
    Set othread = ocell.CommentThreaded
    If Not othread Is Nothing Then
        Dim sthread As String
        sthread = ExtructCommentThreaded(othread)
        oprogress.FilePrint saddress & " スレッド " & sthread
    End If
    Set othread = Nothing

    ExtructCell = errNumber
    On Error GoTo 0
    Exit Function
ErrorHandler:
    errNumber = Err.Number
    oprogress.FilePrint "ExtructCell() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function

''' コメントスレッドの抽出
Function ExtructCommentThreaded(othread As CommentThreaded) As String
    On Error Resume Next

    Dim sthread As String
    sthread = sthread & othread.Date & " " & othread.Author.Name & ":" & Chr(&H22) & othread.Text & Chr(&H22)
    If othread.Replies.Count > 0 Then
        Dim oreply As CommentThreaded
        For Each oreply In othread.Replies
            sthread = sthread & vbLf & ExtructCommentThreaded(oreply)
        Next oreply
    End If

    ExtructCommentThreaded = sthread
    On Error GoTo 0
End Function

''' 図形のテキストを返す
Function ExtructShapeText(oshape As Shape, sindent As String) As String
    On Error Resume Next

    ' グループ化していたら
    Dim sgroups As String
    If oshape.Type = msoGroup Then
        If oshape.GroupItems.Count > 0 Then

            ' 全てのグループ化した図形のテキストを抽出
            Dim oitem As Shape
            For Each oitem In oshape.GroupItems
                sgroups = sgroups & vbLf & sindent  & Space(4) & ExtructShapeText(oitem, sindent & Space(4))
            Next oitem
        End If
    End If

    ' テキスト枠とグループのテキストが在ったら
    Dim sshape As String
    Dim sframe As String
    sframe = ExtructFrameText(oshape.TextFrame)
    If Len(sframe & sgroups) > 0 Then

        ' 図形の名称
        sshape = ExtructShapeName(oshape)
        ' 図形のテキスト枠
        If Len(sframe) > 0 Then
            sshape = sshape & " " & Chr(&H22) & sframe & Chr(&H22)
        End If
        ' 図形のグループ
        sshape = sshape & sgroups
    End If

    ExtructShapeText = sshape
    On Error GoTo 0
End Function

''' 図形の名称を返す
Function ExtructShapeName(oshape As Shape) As String
    On Error Resume Next
    Dim ssharp As String
    ssharp = Chr(&H22) & oshape.Name & Chr(&H22)
    If Len(oshape.AlternativeText) > 0 Then
        ssharp = ssharp & "(" & oshape.AlternativeText & ")"
    End If
    ExtructShapeName = ssharp
    On Error GoTo 0
End Function

' テキスト枠を文字列で返す
Function ExtructFrameText(vframe As Variant) As String
    On Error Resume Next
    ExtructFrameText = ""
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
    ExtructFrameText = svalue
    On Error GoTo 0
End Function
