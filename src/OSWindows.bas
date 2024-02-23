Attribute VB_Name = "OSWindows"
Option Explicit

''' 経過型
Public Type Progressing
    Current As Long
    Count As Long
End Type

''' OSの名称を返す
Public Function osCanBeExecuted() As String
    If Application.OperatingSystem Like "*Mac*" Then
        osCanBeExecuted = "MacOS用を使って下さい。"
    Else
        osCanBeExecuted = ""
    End If
End Function

''' ファイルパスの結合
Public Function osBuildPath(directory As String, filename As String) As String
    osBuildPath = directory & "¥" & filename
End Function 

''' ディレクトリの作成
Public Function osMakeDirectory(directory As String) As String
    ' エラー処理の登録
    osMakeDirectory = ""
    On Error GoTo ErrorHandler

    MkDir directory

    On Error GoTo 0
    Exit Function
ErrorHandler:
    osMakeDirectory = Err.Number & ":" & Err.Description
    Resume Next
End Function

''' クリップボードを画像ファイルに抽出する
Public Function osClipboardToImageFile(filename As String) As String
    ' エラー処理の登録
    osClipboardToImageFile = ""
    On Error GoTo ErrorHandler

    ' クリップボードに画像があれば
    'Dim formats() As XlClipboardFormat
    Dim formats As Variant
    formats = Application.ClipboardFormats
    if formats(1) >= 0 Then

        ' 画像ファイル名を作成
        Dim stype As String
        stype = "png"
        Dim imagefile As String
        imagefile = Replace(filename & "." & stype, " ", "_")

        ' 画像ファイルを書き出し
        Dim wscript As Object
        Set wscript = CreateObject("WScript.Shell")
        Dim scommand As String
        scommand = "powershell Add-Type -AssemblyName System.Windows.Forms;$ImagePath = '" & imagefile & "';  [Windows.Forms.Clipboard]::GetImage().Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::" & stype & ")"
        wscript.Run Command:=scommand, WindowStyle:=0, WaitOnReturn:=True
        Set wscript = Nothing
    End If
    Set formats = Nothing

    On Error GoTo 0
    Exit Function
ErrorHandler:
    osClipboardToImageFile = Err.Number & ":" & Err.Description
    Resume Next
End Function
