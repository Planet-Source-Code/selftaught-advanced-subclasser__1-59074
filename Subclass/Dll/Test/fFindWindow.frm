VERSION 5.00
Begin VB.Form fFindWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Window"
   ClientHeight    =   4125
   ClientLeft      =   2445
   ClientTop       =   2040
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucFindWindow uFindWindow1 
      Height          =   1695
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3473
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   833
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Frame fra 
      Caption         =   "Subclass"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Label lbl 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   5175
      End
      Begin VB.Label lbl 
         Caption         =   $"fFindWindow.frx":0000
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5175
      End
   End
End
Attribute VB_Name = "fFindWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mhWnd As Long

Private moParent As Form
Attribute moParent.VB_VarHelpID = -1

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
    If (Index = 0&) Then moParent.AddSubclass mhWnd
    Hide
End Sub

Public Sub GetWindow(oForm As fTest)
    On Error Resume Next
    Set moParent = oForm
    Show vbModeless
End Sub

Private Sub uFindWindow1_WindowChanged(ByVal hWndNew As Long)
    
    If IsWindowLocal(hWndNew) Then
        mhWnd = hWndNew
        Command1(0).Enabled = True
        lbl(1).Caption = "This window is local, so you can subclass it"
    Else
        mhWnd = 0
        Command1(0).Enabled = False
        lbl(1).Caption = "This window is not local, so you can't subclass it"
    End If
    
End Sub

