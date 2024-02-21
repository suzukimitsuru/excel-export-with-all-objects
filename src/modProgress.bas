Attribute VB_Name = "modProgress"
Option Explicit

''' 経過型
Public Type Progressing
    Current As Long
    Count As Long
End Type

''' ワークブックの経過
Public Type ProgressOfWorkbook
    FileNumber As Integer
    Progress As Range
    Sheets As Progressing
    Cells As Progressing
    Shapes As Progressing
End Type