Attribute VB_Name = "modOS"
Option Explicit

''' OSの名称を返す
Public Function osName() As String
    If Application.OperatingSystem Like "*Mac*" Then
        osName = "MacOS"
    Else
        osName = "Windows"
    End If
End Function

''' ファイルパスの結合
Public Function osBuildPath(directory As String, filename As String) As String
    Dim sepalater As String
    sepalater = IIf(Application.OperatingSystem Like "*Mac*", "/", "¥")
    osBuildPath = directory & sepalater & filename
End Function 

''' ディレクトリの作成
Public Function osMakeDirectory(directory As String) As String
    ' エラー処理の登録
    osMakeDirectory = ""
    On Error GoTo ErrorHandler

    If Application.OperatingSystem Like "*Mac*" Then
        osMakeDirectory = AppleScriptTask("excel-extructor.applescript", "MakeDirectory", directory)
    Else
        MkDir directory
    End If

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
        If Application.OperatingSystem Like "*Mac*" Then
            osClipboardToImageFile = AppleScriptTask("excel-extructor.applescript", "ClipboardToImageFile", imagefile)
        Else
            Dim wscript As Object
            Set wscript = CreateObject("WScript.Shell")
            Dim scommand As String
            scommand = "powershell Add-Type -AssemblyName System.Windows.Forms;$ImagePath = '" & imagefile & "';  [Windows.Forms.Clipboard]::GetImage().Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::" & stype & ")"
            wscript.Run Command:=scommand, WindowStyle:=0, WaitOnReturn:=True
            Set wscript = Nothing
        End If
    End If
    Set formats = Nothing

    On Error GoTo 0
    Exit Function
ErrorHandler:
    osClipboardToImageFile = Err.Number & ":" & Err.Description
    Resume Next
End Function
