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

Sub cmdExtruct_Click()
    Dim errNumber As Integer
    Dim errDescription As String
    On Error GoTo ErrorHandler
    
    If Application.OperatingSystem Like "*Mac*" Then
        Debug.Print (" OS: MacOS")
    Else
        Debug.Print (" OS: Windows")
    End If
    
    Dim filename As String
    filename = Me.Cells(2, 3).Value
    Dim ext_position As Long
    ext_position = InStrRev(filename, ".")
    Debug.Print ("filename:" & filename & ":" & ext_position)
    Dim textfile As String
    If Len(ext_position) > 0 Then
        textfile = Left(filename, ext_position) & "txt"
        Debug.Print ("textfile:" & textfile)
        Dim textnumber As Integer
        textnumber = FreeFile()
        Open textfile For Output As #textnumber
        If errNumber <= 0 Then
            Dim in_book As Workbook
            Set in_book = Workbooks.Open(filename:=filename, ReadOnly:=True)
            Me.Activate
            Dim sheet_number As Long
            For sheet_number = 1 To in_book.Sheets.Count
                errNumber = ExtructSheet(textnumber, in_book.Sheets(sheet_number))
            Next sheet_number
        End If
        in_book.Close (False)
        Set in_book = Nothing
        Close #textnumber
    End If
    
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print ("cmdExtruct_Click() " & Err.Number & ":" & Err.Description)
    Resume Next
End Sub

Function ExtructSheet(textnumber As Integer, sheet As Worksheet)
    Dim errNumber As Integer
    Dim errDescription As String
    On Error GoTo ErrorHandler

    Print #textnumber, sheet.CodeName & ":" & sheet.Name
    Dim min_row As Integer
    Dim min_col As Integer
    Dim max_row As Integer
    Dim max_col As Integer
    min_row = sheet.UsedRange.Row
    min_col = sheet.UsedRange.Column
    max_row = sheet.UsedRange.Rows.Count
    max_col = sheet.UsedRange.Columns.Count
    Dim nrow As Integer
    For nrow = sheet.UsedRange.Row To sheet.UsedRange.Rows.Count
        Dim ncol As Integer
        For ncol = sheet.UsedRange.Column To sheet.UsedRange.Columns.Count
            Dim ocell As range
            Dim saddress As String
            Set ocell = sheet.Cells(nrow, ncol)
            saddress = sheet.CodeName & "!" & ocell.Address(False, False)
            If ocell.HasFormula Then
                Print #textnumber, saddress & " " & Chr(&H22) & ocell.Value2 & Chr(&H22) & "("; ocell.Formula & ")"
            Else
                Print #textnumber, saddress & " " & Chr(&H22) & ocell.Value & Chr(&H22)
            End If
            Dim ocomment As Comment
            Set ocomment = ocell.Comment
            If Not ocomment Is Nothing Then
            End If
            Set ocomment = Nothing
            Set ocell = Nothing
        Next ncol
    Next nrow
    
    If sheet.Shapes.Count > 0 Then
        Dim nshape As Integer
        For nshape = 1 To sheet.Shapes.Count
            Dim oshape As Shape
            Set oshape = sheet.Shapes(nshape)
            Dim sshape As String
            sshape = sheet.CodeName & "!" & oshape.Name
            Print #textnumber, sshape & " " & Chr(&H22) & oshape.TextEffect.Text & Chr(&H22)
            Set oshape = Nothing
        Next nshape
    End If
    
    ExtructSheet = errNumber
    On Error GoTo 0
    Exit Function
ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print ("ExtructSheet() " & Err.Number & ":" & Err.Description)
    Resume Next
End Function
