Attribute VB_Name = "Module1"
Option Explicit

Public passox(0 To 24) As Single
Public passoY(0 To 24) As Single

Public Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "GDI32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Public Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Const STRETCHMODE As Long = vbPaletteModeNone
Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Pt As POINTAPI
Public capture As Integer
Public XX As Long
Public YY As Long
Public Nome_do_gif As String
Public f() As New Obj
Public Qt_obj As Integer
Public RepeatTimes As Long
Public TotalFrames As Long
Public qt_frames() As Integer
Public Diretorio As String
Public x() As Byte

Public Function LoadGif(sFile As String, aImg As Variant) As Boolean
    
  Dim n As Integer
  Dim fNum As Integer
  Dim imgCount As Integer
  
  Dim i As Long
  Dim j As Long
  Dim xOff As Long
  Dim yOff As Long
  Dim TimeWait As Long
  
  Dim imgHeader As String
  Dim fileHeader As String
  Dim buf As String
  Dim picbuf As String
  Dim GifEnd As String
  
    On Error GoTo ErrHandler
 
    LoadGif = False
    
    If Dir$(sFile) = "" Or sFile = "" Then
        MsgBox "File " & sFile & " not found", vbCritical
        Exit Function '>---> Bottom
    End If
  
    GifEnd = Chr$(0) & Chr$(33) & Chr$(249)
    
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
    buf = String$(LOF(fNum), Chr$(0))
    Get #fNum, , buf 'Get GIF File into buffer
    Close fNum
    
    i = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = Left$(buf, j)
    
    If Left$(fileHeader, 3) <> "GIF" Then
        MsgBox "This file is not a *.gif file", vbCritical
        Exit Function '>---> Bottom
    End If
    
    LoadGif = True
    i = j + 2
    
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid$(fileHeader, 126, 1)) + (Asc(Mid$(fileHeader, 127, 1)) * 256&)
      Else
        RepeatTimes = 0
    End If

    Do ' Split GIF Files at separate pictures
        ' and load them into Image Array
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + 3
        If j > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
            picbuf = String$(Len(fileHeader) + j - i, Chr$(0))
            picbuf = fileHeader & Mid$(buf, i - 1, j - i)
            Put #fNum, 1, picbuf
            imgHeader = Left$(Mid$(buf, i - 1, j - i), 16)
            Close fNum
            TimeWait = ((Asc(Mid$(imgHeader, 4, 1))) + (Asc(Mid$(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(Mid$(imgHeader, 9, 1)) + (Asc(Mid$(imgHeader, 10, 1)) * 256&)
                yOff = Asc(Mid$(imgHeader, 11, 1)) + (Asc(Mid$(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            End If
            ' Use .Tag Property to save TimeWait interval for separate Image
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            DoEvents
            Kill ("temp.gif")
            i = j
        End If
        DoEvents
    Loop Until j = 3
    
    ' If there are one more Image - Load it
    If i < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
        picbuf = String$(Len(fileHeader) + Len(buf) - i, Chr$(0))
        picbuf = fileHeader & Mid$(buf, i - 1, Len(buf) - i)
        Put #fNum, 1, picbuf
        imgHeader = Left$(Mid$(buf, i - 1, Len(buf) - i), 16)
        Close fNum
        TimeWait = ((Asc(Mid$(imgHeader, 4, 1))) + (Asc(Mid$(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(Mid$(imgHeader, 9, 1)) + (Asc(Mid$(imgHeader, 10, 1)) * 256)
            yOff = Asc(Mid$(imgHeader, 11, 1)) + (Asc(Mid$(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    qt_frames(Qt_obj) = TotalFrames

Exit Function

ErrHandler:
    MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
    LoadGif = False
    On Error GoTo 0

End Function


