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
    
    ' OSの種類を出力
    If Application.OperatingSystem Like "*Mac*" Then
        Debug.Print (" OS: MacOS")
    Else
        Debug.Print (" OS: Windows")
    End If
    
    ' ファイル拡張子が在ったら
    Dim filename As String
    filename = Me.Cells(2, 3).Value
    Dim ext_position As Long
    ext_position = InStrRev(filename, ".")
    Debug.Print ("filename:" & filename & ":" & ext_position)
    Dim textfile As String
    If Len(ext_position) > 0 Then

        ' 出力するテキストファイルが開けたら
        textfile = Left(filename, ext_position) & "txt"
        Debug.Print ("textfile:" & textfile)
        Dim textnumber As Integer
        textnumber = FreeFile()
        Open textfile For Output As #textnumber
        If errNumber <= 0 Then

            ' Excelブックが開けたら
            Dim in_book As Workbook
            Set in_book = Workbooks.Open(filename:=filename, ReadOnly:=True)
            If errNumber <= 0 Then
                Me.Activate

                ' 全てのシートを抽出する
                Dim osheet As Worksheet
                For Each osheet In in_book.Sheets
                    errNumber = ExtructSheet(textnumber, osheet)
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
    Print #textnumber, "cmdExtruct_Click()  ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Sub

''' シートの抽出
Function ExtructSheet(textnumber As Integer, sheet As Worksheet) As Integer
    Dim errNumber As Integer
    On Error GoTo ErrorHandler
    Print #textnumber, sheet.CodeName & ":" & sheet.Name

    ' 全ての使用領域で
    Dim nrow As Integer
    For nrow = sheet.UsedRange.Row To sheet.UsedRange.Rows.Count
        Dim ncol As Integer
        For ncol = sheet.UsedRange.Column To sheet.UsedRange.Columns.Count

            ' セルを抽出
            Dim ocell As range
            Set ocell = sheet.Cells(nrow, ncol)
            errNumber = ExtructCell(textnumber, ocell)
            Set ocell = Nothing
        Next ncol
    Next nrow
    
    ' 図形が在ったら
    If sheet.Shapes.Count > 0 Then

        ' 全ての図形の抽出
        Dim oshape As Shape
        For Each oshape In sheet.Shapes
            Dim sshape As String
            Dim stype As String
            Dim sframe As String
            sshape = oshape.Parent.CodeName & "!" & Chr(&H22) & oshape.Name & Chr(&H22)
            stype = ShapeType(oshape)
            sframe = TextFrame(oshape.TextFrame)
            Print #textnumber, sshape & " " & stype & " " & Chr(&H22) & sframe & Chr(&H22)
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
Function ExtructCell(textnumber As Integer, ocell As range) As Integer
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
        Dim stype As String
        sshape = saddress & " " & Chr(&H22) & ocomment.Shape.Name & Chr(&H22)
        stype = ShapeType(ocomment.Shape)
        Print #textnumber, sshape & " " & stype & " " & ocomment.Author & ":" & Chr(&H22) & ocomment.Text & Chr(&H22)
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

' 図形の型を文字列で返す
Function ShapeType(oshape As Shape) As String
    Dim stype As String
    Select Case oshape.Type
    Case mso3DModel:           stype = "3D モデル"
    Case msoAutoShape:         stype = "オートシェイプ"
    Case msoCallout:           stype = "吹き出し"
    Case msoCanvas:            stype = "キャンバス"
    Case msoChart:             stype = "グラフ"
    Case msoComment:           stype = "コメント"
    Case msoContentApp:        stype = "コンテンツ Office アドイン"
    Case msoDiagram:           stype = "図"
    Case msoEmbeddedOLEObject: stype = "埋め込み OLE オブジェクト"
    Case msoFormControl:       stype = "フォーム コントロール"
    Case msoFreeform:          stype = "フリーフォーム"
    Case msoGraphic:           stype = "グラフィック"
    Case msoGroup:             stype = "Group"
    'Case msoIgxGraphic:        stype = "SmartArt グラフィック"
    Case msoInk:               stype = "インク"
    Case msoInkComment:        stype = "インク コメント"
    Case msoLine:              stype = "Line"
    Case msoLinked3DModel:     stype = "リンクされた 3D モデル"
    Case msoLinkedGraphic:     stype = "リンクされたグラフィック"
    Case msoLinkedOLEObject:   stype = "リンク OLE オブジェクト"
    Case msoLinkedPicture:     stype = "リンク画像"
    Case msoMedia:             stype = "メディア"
    Case msoOLEControlObject:  stype = "OLE コントロール オブジェクト"
    Case msoPicture:           stype = "画像"
    Case msoPlaceholder:       stype = "プレースホルダー"
    Case msoScriptAnchor:      stype = "スクリプト アンカー"
    Case msoShapeTypeMixed:    stype = "図形の種類の組み合わせ"
    Case msoSlicer:            stype = "Slicer"
    Case msoTable:             stype = "テーブル"
    Case msoTextBox:           stype = "テキスト ボックス"
    Case msoTextEffect:        stype = "テキスト効果"
    Case msoWebVideo:          stype = "Web ビデオ"
    Case Else:                 stype = "[不明]"
    End Select
    ShapeType = stype
End Function

' テキスト枠を文字列で返す
Function TextFrame(vframe As Variant) As String
    TextFrame = ""
    Dim ncolor As Double
    Dim fBold As Boolean
    Dim fItalic As Boolean
    Dim fStrikethrough As Boolean
    Dim sUnderline As XlUnderlineStyle
    ncolor = &H808080
    fBold = False
    fItalic = False
    fStrikethrough = False
    sUnderline = xlUnderlineStyleNone
    Dim svalue As String
    Dim nchara As Integer
    For nchara = 1 To Len(vframe.Characters().Text)
        If ncolor <> vframe.Characters(nchara, 1).Font.Color Then
            If ncolor <> &H808080 Then svalue = svalue & "</色>"
            ncolor = vframe.Characters(nchara, 1).Font.Color
            svalue = svalue & "<色:0x" & Hex(ncolor) & ">"
        End If
        If fBold <> vframe.Characters(nchara, 1).Font.Bold Then
            fBold = vframe.Characters(nchara, 1).Font.Bold
            svalue = svalue & IIf(fBold, "<太字>", "</太字>")
        End If
        If fItalic <> vframe.Characters(nchara, 1).Font.Italic Then
            fItalic = vframe.Characters(nchara, 1).Font.Italic
            svalue = svalue & IIf(fItalic, "<斜体>", "</斜体>")
        End If
        If fStrikethrough <> vframe.Characters(nchara, 1).Font.Strikethrough Then
            fStrikethrough = vframe.Characters(nchara, 1).Font.Strikethrough
            svalue = svalue & IIf(fStrikethrough, "<取り消し線>", "</取り消し線>")
        End If
        If sUnderline <> vframe.Characters(nchara, 1).Font.Underline Then
            sUnderline = vframe.Characters(nchara, 1).Font.Underline
            Select Case sUnderline
                Case xlUnderlineStyleDouble:            svalue = svalue & "<太い二重下線>"
                Case xlUnderlineStyleDoubleAccounting:  svalue = svalue & "<並んだ2本の細い線>"
                Case xlUnderlineStyleNone:              svalue = svalue & "</下線>"
                Case xlUnderlineStyleSingle:            svalue = svalue & "<一重下線>"
                Case xlUnderlineStyleSingleAccounting:  svalue = svalue & "<非サポート下線>"
                Case Else:                              svalue = svalue & "<不明な下線>"
            End Select
        End If
        svalue = svalue & vframe.Characters(nchara, 1).Text
    Next nchara
    If ncolor <> &H808080                   Then svalue = svalue & "</色>" 
    If fBold <> False                       Then svalue = svalue & "</太字>"
    If fItalic <> False                     Then svalue = svalue & "</斜体>"
    If fStrikethrough <> False              Then svalue = svalue & "</取り消し線>"
    If sUnderline <> xlUnderlineStyleNone   Then svalue = svalue & "</下線>"
    TextFrame = svalue
End Function
