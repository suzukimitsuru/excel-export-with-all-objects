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

'''  実行ボタンのクリックイベント
Public Sub cmdExecute_Click()
    ' エラー処理の登録
    Dim errNumber As Integer
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' 経過表示セルを取得
    Dim oprogress As Progress
    Set oprogress = New Progress
    oprogress.SetView Me.Cells(4, 3)
    oprogress.ErrorView ""
    
   ' ファイル名を取得
    Dim original_file As String
    original_file = Me.Cells(2, 3).Value
    Dim compare_file As String
    compare_file = IIf(Me.Cells(2, 1).Value = "Compare", Me.Cells(3, 3).Value, "")
    Me.Cells(3, 3).Value = compare_file

    ' このOSで実行出来なかったら、中断する
    Dim serror As String
    serror = osCanBeExecuted()
    If Len(serror) > 0 Then
        Application.ScreenUpdating = True
        oprogress.ErrorView serror
    Else
        ' 原始ブックが開けたら
        errNumber = 0
        Dim original_book As Workbook
        Set original_book = Workbooks.Open(filename:=original_file, ReadOnly:=True)
        If errNumber <= 0 Then
            Application.ScreenUpdating = True
            Me.Activate

            ' シートの内部名称が取得出来なかったら
            If original_book.Sheets(1).CodeName = "" Then
                oprogress.ErrorView "開発リボンの Visual Basic を開いて下さい。"
            Else
                ' 比較なら
                If Len(compare_file) > 0 Then

                    ' 比較ブックが開けたら
                    Dim compare_book As Workbook
                    Set compare_book = Workbooks.Open(filename:=compare_file, ReadOnly:=True)
                    If errNumber <= 0 Then

                        ' ブックを比較する
                        errNumber = CompareBook(oprogress, original_book, compare_book)

                        ' 原始ブックを閉じる
                        compare_book.Close(False)
                        Set compare_book = Nothing
                    Else
                        Application.ScreenUpdating = True
                        oprogress.ErrorView "比較ブックが開けませんでした。"
                    End If
                Else
                    ' ブックを抽出する
                    errNumber = ExtructBook(oprogress, original_book, original_file)
                End If
            End If

            ' 原始ブックを閉じる
            original_book.Close(False)
            Set original_book = Nothing
        Else
            Application.ScreenUpdating = True
            oprogress.ErrorView "Excel ブックが開けませんでした。"
        End If
    End If
    
    Set oprogress = Nothing
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    errNumber = Err.Number
    If oprogress.FileIsOpend Then
        oprogress.FilePrint("cmdExecute_Click() ERROR " & Err.Number & ":" & Err.Description)
    Else
        Debug.Print("cmdExecute_Click() ERROR " & Err.Number & ":" & Err.Description)
    End If
    Resume Next
End Sub

''' ブックの抽出
Private Function ExtructBook(oprogress As Progress, obook As Workbook, filename As String) As Integer
    On Error GoTo ErrorHandler

    ' ファイル拡張子の位置を調べる
    Dim ext_position As Long
    ext_position = InStrRev(filename, ".")
    If Len(ext_position) <= 0 Then
        ext_position = Len(filename)
    End If

    ' 出力ディレクトリを作成する
    Dim directory As String
    directory = Left(filename, ext_position - 1)
    Dim serror As String
    serror = osMakeDirectory(directory)
    If Len(serror) > 0 Then
        oprogress.ErrorView "出力するディレクトリが作れませんでした。" & serror
    End If

    ' 出力するテキストファイルが開けたら
    Dim textfile As String
    textfile = osBuildPath(directory, "export.txt")
    serror = oprogress.FileOpen(textfile)
    If Len(serror) > 0 Then
        oprogress.ErrorView "出力するテキストファイルが開けませんでした。" & serror
    Else
        ' 全てのシートの経過の数を集計する
        Dim osheet As Worksheet
        For Each osheet In obook.Sheets
            oprogress.FilePrint osheet.CodeName & ":" & osheet.Name & " Cells:" & osheet.UsedRange.Count & " Shapes:" & osheet.UsedRange.Count
            oprogress.ProgressSumCount osheet.UsedRange.Count, osheet.Shapes.Count
        Next osheet

        ' 全てのシートを抽出する
        For Each osheet In obook.Sheets
            ExtructBook = ExtructSheet(oprogress, osheet, directory)
            oprogress.ProgressCountupSheets
        Next osheet
        oprogress.ProgressView

        ' 出力したテキストファイルを閉じる
        serror = oprogress.FileClose()
    End If

    On Error GoTo 0
    Exit Function
ErrorHandler:
    ExtructBook = Err.Number
    oprogress.FilePrint "ExtructBook() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function

''' シートの抽出
Private Function ExtructSheet(oprogress As Progress, osheet As Worksheet, directory As String) As Integer
    On Error GoTo ErrorHandler

    ' 使用領域が在ったら
    If osheet.UsedRange.Count > 0 Then
        oprogress.FilePrint "--- " & osheet.CodeName & ":" & osheet.Name & " - セル ---"

        ' 全ての使用領域で
        Dim nrow As Integer
        For nrow = osheet.UsedRange.Row To osheet.UsedRange.Rows.Count
            Dim ncol As Integer
            For ncol = osheet.UsedRange.Column To osheet.UsedRange.Columns.Count

                ' セルを抽出
                Dim ocell As range
                Set ocell = osheet.Cells(nrow, ncol)
                If ocell.Address = ocell.MergeArea(1, 1).Address Then

                    ' セルのアドレスを取得する
                    Dim saddress As String
                    saddress = ocell.Parent.CodeName & "!" & ocell.MergeArea.Address(False, False)
                    ' セル文字列を抽出
                    oprogress.FilePrint saddress & " " & CellText(ocell)

                    ' コメントが在ったら、抽出する
                    If Not ocell.Comment Is Nothing Then
                        oprogress.FilePrint saddress & " " & Chr(&H22) & ocell.Comment.Shape.Name & Chr(&H22) & _
                            " " & ocell.Comment.Author & ":" & Chr(&H22) & ocell.Comment.Text & Chr(&H22)
                    End If

                    ' スレッドが在ったら、抽出する
                    If Not ocell.CommentThreaded Is Nothing Then
                        Dim sthread As String
                        sthread = CommentThreadedText(ocell.CommentThreaded)
                        oprogress.FilePrint saddress & " スレッド " & sthread
                    End If
                End If
                Set ocell = Nothing
                oprogress.ProgressCountupCells
            Next ncol
            oprogress.ProgressView
            DoEvents
        Next nrow
    End If
    
    ' 図形が在ったら
    If osheet.Shapes.Count > 0 Then
        oprogress.FilePrint "--- " & osheet.CodeName & ":" & osheet.Name & " - 図形 ---"

        ' 全ての図形の抽出
        Dim oshape As Shape
        For Each oshape In osheet.Shapes
            oprogress.FilePrint ExtructShape(oshape, "", oshape.Parent.CodeName, directory)
            oprogress.ProgressCountupShapes
            oprogress.ProgressView
            DoEvents
        Next oshape
    End If
    
    On Error GoTo 0
    Exit Function
ErrorHandler:
    ExtructSheet = Err.Number
    oprogress.FilePrint "ExtructSheet() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function

''' 図形の抽出
Public Function ExtructShape(oshape As Shape, sindent As String, sheetName As String, directory As String) As String
    On Error Resume Next
    Dim sshape As String

    ' 図形のテキストを求める
    Dim saddress As String
    saddress = sheetName & "!" & ShapeName(oshape)
    Dim sframe As String
    sframe = FrameText(oshape.TextFrame)
    sshape = saddress & " " & Chr(&H22) & sframe & Chr(&H22)

    ' クリップボードを画像ファイルに書き込む
    Dim errMessage As String
    errMessage = ShapeToImageFile(oshape, osBuildPath(directory, sheetName & "!" & oshape.Name))
    if Len(errMessage) > 0 Then
        Debug.Print(saddress & " ERROR " & errMessage)
    End If

    ' グループ化していたら
    If oshape.Type = msoGroup Then
        If oshape.GroupItems.Count > 0 Then

            ' 全てのグループ化した図形を抽出
            Dim oitem As Shape
            For Each oitem In oshape.GroupItems
                sshape = sshape & vbLf & sindent & vbTab & ExtructShape(oitem, sindent & vbTab, sheetName, directory)
            Next oitem
        End If
    End If

    ExtructShape = sshape
    On Error GoTo 0
End Function

''' ブックの比較
Private Function CompareBook(oprogress As Progress, original As Workbook, compare As Workbook) As Integer
    On Error GoTo ErrorHandler

    ' シートの変更を調査する
    Dim doriginal() As CompareEntry
    Dim dcompare() As CompareEntry
    EnumSheets doriginal, original
    EnumSheets dcompare, compare
    Dim dremoves() As CompareEntry
    Dim daddes() As CompareEntry
    Dim dchanges() As CompareEntry
    EnumSheets dremoves, original
    EnumSheets daddes, compare
    EnumSheets dchanges, original
    ExtructChanges daddes, dremoves, dchanges, doriginal, dcompare

    Dim ocompare As Worksheet
    Set ocompare = Application.ThisWorkbook.Sheets("比較結果")

    ' 全てのシートの経過の数を集計する
    Dim nconpare_row As Integer
    nconpare_row = 1
    Dim entry As CompareEntry
    Dim osheet As Worksheet
    Dim index As Integer
    For index = LBound(daddes) To UBound(daddes)
        entry = daddes(index)
        Set osheet = compare.Sheets(entry.Index)
        ocompare.Cells(nconpare_row, 1).Value = "Add"
        ocompare.Cells(nconpare_row, 2).Value = entry.Name & ":" & osheet.Name & " Cells:" & osheet.UsedRange.Count & " Shapes:" & osheet.UsedRange.Count
        nconpare_row = nconpare_row + 1
    Next index
    For index = LBound(dremoves) To UBound(dremoves)
        entry = dremoves(index)
        Set osheet = original.Sheets(entry.Index)
        ocompare.Cells(nconpare_row, 1).Value = "Remove"
        ocompare.Cells(nconpare_row, 2).Value = entry.Name & ":" & osheet.Name & " Cells:" & osheet.UsedRange.Count & " Shapes:" & osheet.UsedRange.Count
        nconpare_row = nconpare_row + 1
    Next index
    For index = LBound(dchanges) To UBound(dchanges)
        entry = dchanges(index)
        Set osheet = original.Sheets(entry.Index)
        oprogress.ProgressSumCount osheet.UsedRange.Count, osheet.Shapes.Count
        nconpare_row = nconpare_row + 1
    Next index

    ' 変更したシートを比較する
    Dim dchange As CompareEntry
    For index = LBound(dchanges) To UBound(dchanges)
        entry = dchanges(index)
        Set osheet = original.Sheets(entry.Index)
        oprogress.ProgressView
        'CompareBook = ExtructSheet(oprogress, osheet, directory)
        nconpare_row = nconpare_row + 1
        oprogress.ProgressCountupSheets
    Next index
    oprogress.ProgressView

    On Error GoTo 0
    Exit Function
ErrorHandler:
    CompareBook = Err.Number
    oprogress.ErrorView "CompareBook() ERROR " & Err.Number & ":" & Err.Description
    Resume Next
End Function
