VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''' 経過型
Private Type Progressing
    Current As Long
    Count As Long
End Type

Private mFileNumber As Integer
Private mFileIsOpend As Boolean
Private mView As Range
Private mSheets As Progressing
Private mCells As Progressing
Private mShapes As Progressing

''' クラス生成
Private Sub Class_Initialize()
    mFileNumber = FreeFile()
    mFileIsOpend = False
End Sub

''' クラス破棄
Private Sub Class_Terminate()
    Set mView = Nothing
End Sub

''' 表示セルの設定
Public Sub SetView(ocell As Range)
    Set mView = ocell
End Sub

''' エラー表示
Public Sub ErrorView(smessage As String)
    mView.Font.Color = RGB(255, 0, 0)
    mView.Value = smessage
End Sub

''' ファイルを開く
Public Function FileOpen(filename As String)
    FileOpen = ""
    On Error GoTo ErrorHandler

    Open filename For Output As #mFileNumber
    mFileIsOpend = True

    On Error GoTo 0
    Exit Function
ErrorHandler:
    FileOpen = Err.Number & ":" & Err.Description
    Resume Next
End Function

''' ファイルを閉じる
Public Function FileClose()
    FileClose = ""
    On Error GoTo ErrorHandler

    Close #mFileNumber
    mFileIsOpend = False

    On Error GoTo 0
    Exit Function
ErrorHandler:
    FileClose = Err.Number & ":" & Err.Description
    Resume Next
End Function

''' ファイルを開いているか返す
Public Property Get FileIsOpend() As Boolean
    FileIsOpend = mFileIsOpend
End Property

''' ファイルに1行出力する
Public Sub FilePrint(sline As String)
    Print #mFileNumber, sline
End Sub

''' 経過の数を集計する
Public Sub ProgressSumCount(ncells As Integer, nsharps As Integer)
    mSheets.Count = mSheets.Count + 1
    mCells.Count = mCells.Count + ncells
    mShapes.Count = mShapes.Count + nsharps
End Sub

''' セルをカウントする
Public Sub ProgressCountupSheets()
    mSheets.Current = mSheets.Current +1
End Sub

''' セルをカウントする
Public Sub ProgressCountupCells()
    mCells.Current = mCells.Current +1
End Sub

''' セルをカウントする
Public Sub ProgressCountupShapes()
    mShapes.Current = mShapes.Current +1
End Sub

''' エラー表示
Public Sub ProgressView()
    mView.Font.Color = RGB(0, 0, 0)
    mView.Value = mSheets.Current & "/" & mSheets.Count & "シート " & _
        mCells.Current & "/" & mCells.Count & "セル " & _
        mShapes.Current & "/" & mShapes.Count & "図形"
End Sub
