Attribute VB_Name = "CompareItem"
Option Explicit

''' ��r�G���g��
Public Type CompareEntry
    Name As String
    Index As Long
    Compare As Long
End Type

''' �V�[�g���
Public Sub EnumSheets(ByRef entries() As CompareEntry, obook As Workbook)
    On Error Resume Next
    Dim nstack As Integer
    nstack = 0
    Dim osheet As Worksheet
    For Each osheet In obook.Sheets
        nstack = nstack + 1
        ReDim Preserve entries(1 To nstack)
        entries(nstack).Index = nstack
        entries(nstack).Name = osheet.CodeName
    Next osheet
    On Error GoTo 0
End Sub

''' �}�`���
Public Sub EnumShapes(ByRef entries() As CompareEntry, osheet As Worksheet)
    On Error Resume Next
    Dim nstack As Integer
    nstack = 0
    Dim oshape As Shape
    For Each oshape In osheet.Shapes
        nstack = nstack + 1
        ReDim Preserve entries(1 To nstack)
        entries(nstack).Index = nstack
        entries(nstack).Name = oshape.Name
    Next oshape
    On Error GoTo 0
End Sub

''' �ύX�𒊏o
Public Sub ExtructChanges(ByRef addes() As CompareEntry, ByRef removes() As CompareEntry, ByRef changes() As CompareEntry, _
    original() As CompareEntry, compare() As CompareEntry)
    On Error Resume Next

    ' �ǉ����X�g�𒊏o
    removeEqualEntries addes, original
    reductionEntries addes

    ' �폜���X�g�𒊏o
    removeEqualEntries removes, compare
    reductionEntries removes

    ' �ύX���X�g�𒊏o
    removeEqualEntries changes, removes
    reductionEntries changes
    On Error GoTo 0
End Sub

''' �����V�[�g�����O
Private Sub removeEqualEntries(ByRef removes() As CompareEntry, refrences() As CompareEntry)
    On Error Resume Next
    Dim nremove As Integer
    For nremove = LBound(removes) To UBound(removes)
        Dim nrefrence As Integer
        For nrefrence = LBound(refrences) To UBound(refrences)
            If removes(nremove).Name = refrences(nrefrence).Name Then
                removes(nremove).Index = 0
                Exit For
            End If
        Next nrefrence
    Next nremove
    On Error GoTo 0
End Sub

''' ������z����k������
Private Sub reductionEntries(ByRef entries() As CompareEntry)
    On Error Resume Next
    Dim nlow As Integer
    nlow = LBound(entries)
    Dim npack As Integer
    npack = nlow - 1
    Dim nindex As Integer
    For nindex = LBound(entries) To UBound(entries)
        If entries(nindex).Index > 0 Then
            npack = npack + 1
            entries(npack) = entries(nindex)
        End If
    Next nindex
    If nlow <= npack Then
        ReDim Preserve entries(nlow To npack)
    Else
        ReDim entries(0)
    End If
    On Error GoTo 0
End Sub
