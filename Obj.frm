VERSION 5.00
Begin VB.Form Obj 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   107
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1095
      Top             =   3015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Index           =   0
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox Trabalho 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   3810
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   2940
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Obj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Cur_frame As Integer

Private Sub Form_DblClick()

    Timer1.Interval = 1
    Timer1.Enabled = Timer1.Enabled Xor -1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Cur_frame = Cur_frame + 1
    
    Show_frame

End Sub

Private Sub Form_Load()

  Dim Ret As Long
   
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_COLORKEY Or LWA_ALPHA
    
    ReDim Preserve qt_frames(Qt_obj)
    
    If LoadGif(Nome_do_gif, Picture1) Then
        Width = Picture1(0).Width * Screen.TwipsPerPixelX
        Height = Picture1(0).Height * Screen.TwipsPerPixelY
        Trabalho.Width = Picture1(0).Width
        Trabalho.Height = Picture1(0).Height
        GdiTransparentBlt hDC, 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture1(0).hDC, 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, vbWhite
        Refresh
    End If
  
    Tag = Qt_obj

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 2 Then
        PopupMenu Form1.Menu
        Exit Sub
    End If

    XX = x * Screen.TwipsPerPixelX
    YY = Y * Screen.TwipsPerPixelY
    capture = True
    ReleaseCapture
    SetCapture Me.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If capture Then
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - XX, Pt.Y * Screen.TwipsPerPixelY - YY
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    capture = False

End Sub

Private Sub Timer1_Timer()
  Static Rodando As Integer
    
    Move Left + passox(Cur_frame), Top + passoY(Cur_frame)

    If Left < 0 Or Left + Width > Screen.Width Or (Top + Height) > Screen.Height Or Top < 0 Then
        Rodando = Rodando + 1
        Cur_frame = Cur_frame + 1
        Show_frame
        If Rodando = 24 Then
            Rodando = 0
            If Left < Screen.Width / 2 Then
                Left = Left + 100
            Else
                Left = Left - 100
            End If
            If Top < Screen.Height / 2 Then
                Top = Top + 100
            Else
                Top = Top - 100
            End If
        End If
        Exit Sub
    End If

    Rodando = False
    
   
    Select Case Int(Rnd * 100)
    
      Case 0
        Cur_frame = Cur_frame + 1
        Show_frame
      Case 1
        Cur_frame = Cur_frame - 1
        Show_frame
    End Select
    
    Timer1.Interval = Picture1(Cur_frame).Tag
    
End Sub

Private Sub Show_frame()
    
    If Cur_frame > qt_frames(Tag) Then
        Cur_frame = 0
    End If
        
    If Cur_frame < 0 Then
        Cur_frame = qt_frames(Tag) - 1
    End If
    
    Trabalho.Cls
    GdiTransparentBlt Trabalho.hDC, Picture1(Cur_frame).Left / Screen.TwipsPerPixelX, Picture1(Cur_frame).Top / Screen.TwipsPerPixelY, Picture1(Cur_frame).ScaleWidth, Picture1(Cur_frame).ScaleHeight, Picture1(Cur_frame).hDC, 0, 0, Picture1(Cur_frame).ScaleWidth, Picture1(Cur_frame).ScaleHeight, vbWhite
    Trabalho.Refresh
    StretchBlt hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, vbSrcCopy
    Refresh

End Sub


