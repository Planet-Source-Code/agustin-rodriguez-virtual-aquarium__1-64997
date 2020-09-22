VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Hide_icon 
         Caption         =   "Hide Desktop Icons"
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Begin VB.Menu Menu_about 
            Caption         =   "Autor: Agustin Rodriguez"
            Index           =   0
         End
         Begin VB.Menu Menu_about 
            Caption         =   "E-Mail: virtual_guitar_1@hotmail.com"
            Index           =   1
         End
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Oculte_icones_do_desktop As Integer
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long



Private Sub Exit_Click()
    Static i As Integer

    For i = 0 To Qt_obj - 1
        Unload f(i)
    Next
    Unload Me
    End

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 24
        passox(i) = (-Sin((i + 1) * 15 * 3.14 / 180)) * 100
        passoY(i) = (Cos((i + 1) * 15 * 3.14 / 180)) * 100
    Next i

    For i = 101 To 110

        x = LoadResData(i, "CUSTOM")

        Open "c:\temp.gif" For Binary As 1
        Put #1, 1, x
        Close 1
        Nome_do_gif = "c:\temp.gif"

        ReDim Preserve f(Qt_obj)

        f(Qt_obj).Show
        f(Qt_obj).Move 1000 + Int(Rnd * Screen.Width / 3), 1000 + Int(Rnd * Screen.Height / 3)
        Kill Nome_do_gif
        f(Qt_obj).Timer1.Enabled = True
        Qt_obj = Qt_obj + 1
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim hwnd As Long
    hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hwnd, 5
End Sub

Private Sub Hide_icon_Click()

  Dim hwnd As Long, i As Integer

    Oculte_icones_do_desktop = Oculte_icones_do_desktop Xor 1

    Select Case Oculte_icones_do_desktop
      Case 0
        hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
        ShowWindow hwnd, 5
        Hide_icon.Caption = "Hide Desktop Icons"
      Case 1
        hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
        ShowWindow hwnd, 0
        Hide_icon.Caption = "Show Desktop Icons"
    End Select

End Sub

