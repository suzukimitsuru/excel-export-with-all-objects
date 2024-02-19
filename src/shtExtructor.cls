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
    Dim textnumber As Integer
    
    ' OSの種類を出力
    Debug.Print ("osName(): " & osName())
    
    ' ファイル拡張子が在ったら
    Dim filename As String
    filename = Me.Cells(2, 3).Value
    Dim ext_position As Long
    ext_position = InStrRev(filename, ".")
    If Len(ext_position) > 0 Then
        Debug.Print ("filename:" & filename & ":" & ext_position)

        Dim directory As String
        directory = Left(filename, ext_position - 1)
        Dim errMessage As String
        errMessage = osMakeDirectory(directory)
        Debug.Print ("osMakeDirectory(" & directory & "):" & errMessage)

        ' 出力するテキストファイルが開けたら
        Dim textfile As String
        textfile = osBuildPath(directory, "export.txt")
        Debug.Print ("textfile:" & textfile)

        textnumber = FreeFile()
        Open textfile For Output As #textnumber
        If errNumber <= 0 Then

            ' Excelブックが開けたら
            Dim in_book As Workbook
            Set in_book = Workbooks.Open(filename:=filename, ReadOnly:=True)
            If errNumber <= 0 Then
                Me.Activate
                Dim categories() As ShapeCategory
                categories = ShapeCategoriesInitialize()

                ' 全てのシートを抽出する
                Dim osheet As Worksheet
                For Each osheet In in_book.Sheets
                    errNumber = ExtructSheet(textnumber, osheet, directory, categories)
                Next osheet

                ' Excelブックを閉じる
                in_book.Close (False)
            End If
            Set in_book = Nothing

            ' 出力したテキストファイルを閉じる
            Close #textnumber
        End If
    End If
    
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    errNumber = Err.Number
    If textnumber > 0 Then
        Print #textnumber, "cmdExtruct_Click()  ERROR " & Err.Number & ":" & Err.Description
    Else
        Debug.Print("cmdExtruct_Click()  ERROR " & Err.Number & ":" & Err.Description)
    End If
    Resume Next
End Sub

''' シートの抽出
Function ExtructSheet(textnumber As Integer, sheet As Worksheet, directory As String, ByRef categories() As ShapeCategory) As Integer
    Dim errNumber As Integer
    On Error GoTo ErrorHandler
    Print #textnumber, sheet.CodeName & ":" & sheet.Name

    ' 使用領域が在ったら
    If sheet.UsedRange.Rows.Count > 0 Then

        ' 全ての使用領域で
        Dim nrow As Integer
        For nrow = sheet.UsedRange.Row To sheet.UsedRange.Rows.Count
            Dim ncol As Integer
            For ncol = sheet.UsedRange.Column To sheet.UsedRange.Columns.Count

                ' セルを抽出
                Dim ocell As range
                Set ocell = sheet.Cells(nrow, ncol)
                errNumber = ExtructCell(textnumber, ocell, categories)
                Set ocell = Nothing
            Next ncol
        Next nrow
    End If
    
    ' 図形が在ったら
    If sheet.Shapes.Count > 0 Then

        ' 全ての図形の抽出
        Dim oshape As Shape
        For Each oshape In sheet.Shapes
            Dim category As ShapeCategory
            category = ShapeCategoriesFind(categories, oshape.Type)
            Dim sshape As String
            sshape = oshape.Parent.CodeName & "!" & Chr(&H22) & oshape.Name & Chr(&H22)
            Dim salternative As String
            If Len(oshape.AlternativeText) > 0 Then
                salternative = ":" & Chr(&H22) & oshape.AlternativeText & Chr(&H22)
            End If
            Dim sframe As String
            If category.hasText Then
                sframe = " " & Chr(&H22) & TextFrame(oshape.TextFrame) & Chr(&H22)
            End If
            If category.hasImage Then
                oshape.Copy
                Dim imagefile As String
                imagefile = osBuildPath(directory, oshape.Parent.CodeName & "!" & oshape.Name)
                Dim errMessage As String
                errMessage = osClipboardToImageFile(imagefile)
                if Len(errMessage) > 0 Then
                    Print #textnumber, sshape & " ERROR " & errMessage
                End If
            End If
            Print #textnumber, sshape & " " & category.name & salternative & sframe
        Next oshape
    End If
    
    ExtructSheet = errNumber
    On Error GoTo 0
    Exit Function
ErrorHandler:
    errNumber = Err.Number
    Print #textnumber, "ExtructSheet() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function

''' セルの抽出
Function ExtructCell(textnumber As Integer, ocell As range, categories() As ShapeCategory) As Integer
    Dim errNumber As Integer
    On Error GoTo ErrorHandler

    ' セルの内容を抽出する
    Dim saddress As String
    saddress = ocell.Parent.CodeName & "!" & ocell.Address(False, False)
    If ocell.HasFormula Then
        Print #textnumber, saddress & " " & Chr(&H22) & ocell.Value2 & Chr(&H22) & "("; ocell.Formula & ")"
    Else
        Dim sframe As String
        sframe = TextFrame(ocell)
        Print #textnumber, saddress & " " & Chr(&H22) & sframe & Chr(&H22)
    End If

    ' コメントを抽出する
    Dim ocomment As Comment
    Set ocomment = ocell.Comment
    If Not ocomment Is Nothing Then
        Dim sshape As String
        sshape = saddress & " " & Chr(&H22) & ocomment.Shape.Name & Chr(&H22)
        Dim category As ShapeCategory
        category = ShapeCategoriesFind(categories, ocomment.Shape)
        Print #textnumber, sshape & " " & category.name & " " & ocomment.Author & ":" & Chr(&H22) & ocomment.Text & Chr(&H22)
    End If
    Set ocomment = Nothing

    ' スレッドを抽出する
    Dim othread As CommentThreaded
    Set othread = ocell.CommentThreaded
    If Not othread Is Nothing Then
        Dim sreplies As String
        sreplies = sreplies & othread.Date & " " & othread.Author.Name & ":" & othread.Text
        If othread.Replies.Count > 0 Then
            Dim oreply As CommentThreaded
            For Each oreply In othread.Replies
                sreplies = sreplies  & vbLf & oreply.Date & " " & oreply.Author & ":" & oreply.Text
            Next oreply
        End If
        Print #textnumber, saddress & " スレッド " & Chr(&H22) & sreplies & Chr(&H22)
    End If
    Set othread = Nothing

    ExtructCell = errNumber
    On Error GoTo 0
    Exit Function
ErrorHandler:
    errNumber = Err.Number
    Print #textnumber, "ExtructCell() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function

' テキスト枠を文字列で返す
Function TextFrame(vframe As Variant) As String
    TextFrame = ""
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
    TextFrame = svalue
End Function
