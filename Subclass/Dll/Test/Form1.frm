VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Thorough Testing Tool"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CopyRect Lib "user32.dll" (lpDestRect As Any, lpSourceRect As Any) As Long

Implements iSubclass

Private Sub Command1_Click()
    fTest.Show
End Sub

Private Sub Form_Load()
    With Subclasses(Me).Add(hWnd)
        .AddMsg WM_SIZING, MSG_BEFORE
        '.AddMsg WM_*
        '.AddMsg ....
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Subclasses(Me).Clear
End Sub

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hWnd As Long, iMsg As Subclass.eMsg, wParam As Long, lParam As Long)
    
    Dim ltRect As RECT
    With ltRect
        .Left = ScaleX(Me.Left, vbTwips, vbPixels)
        .Top = ScaleY(Me.Top, vbTwips, vbPixels)
        .Right = .Left + ScaleX(Me.Width, vbTwips, vbPixels)
        .Bottom = .Top + ScaleY(Me.Height, vbTwips, vbPixels)
    End With
    
    CopyRect ByVal lParam, ltRect
    
    lReturn = 1
    bHandled = True
End Sub
