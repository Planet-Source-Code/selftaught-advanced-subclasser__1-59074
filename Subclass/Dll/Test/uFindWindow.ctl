VERSION 5.00
Begin VB.UserControl ucFindWindow 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   ScaleHeight     =   1695
   ScaleWidth      =   5340
   Begin VB.Label lbl 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1080
      Picture         =   "uFindWindow.ctx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "Finder Tool:"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "ucFindWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long




Event WindowChanged(ByVal hWndNew As Long)
Event WindowSelected(ByVal hWndNew As Long)

Private mhWnd As Long

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        With Screen
            Set .MouseIcon = Image1.Picture
            .MousePointer = vbCustom
        End With
        Image1.Visible = False
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button And vbLeftButton Then
        Dim ltPoint As POINTAPI
        With ltPoint
            GetCursorPos ltPoint
            ShowWindowInfo WindowFromPoint(.x, .y)
        End With
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        With Screen
            Set .MouseIcon = Nothing
            .MousePointer = vbDefault
        End With
        Image1.Visible = True
        RaiseEvent WindowSelected(mhWnd)
    End If
End Sub

Private Sub ShowWindowInfo(ByVal hWndShow As Long)
    lbl(0).Caption = "Window: 0x" & FmtHex(hWndShow) & vbCrLf & _
                     "Caption:  " & WindowText(hWndShow) & vbCrLf & _
                     "Class:    " & ClassName(hWndShow) & vbCrLf & _
                     "EXE:      " & ExeFileName(hWndShow) & vbCrLf

    If hWndShow <> mhWnd Then
        RaiseEvent WindowChanged(hWndShow)
        mhWnd = hWndShow
    End If
End Sub


Public Property Get ShowDescription() As Boolean
    ShowDescription = lbl(0).Visible
End Property
Public Property Let ShowDescription(ByVal bVal As Boolean)
    lbl(0).Visible = bVal
End Property

Public Property Let Enabled(ByVal bVal As Boolean)
    UserControl.Enabled = bVal
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

