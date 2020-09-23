VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Testing..."
   ClientHeight    =   8160
   ClientLeft      =   540
   ClientTop       =   1380
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   2
      Top             =   3720
      Width           =   9135
   End
   Begin VB.PictureBox picDisplayHeader 
      Align           =   1  'Align Top
      HasDC           =   0   'False
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   2760
      Width           =   9195
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   45
         TabIndex        =   1
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   8940
      End
   End
   Begin VB.PictureBox picSubclass 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2760
      Left            =   0
      ScaleHeight     =   2760
      ScaleWidth      =   9195
      TabIndex        =   3
      Top             =   0
      Width           =   9195
      Begin VB.CommandButton cmd 
         Caption         =   "Remove"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1575
      End
      Begin VB.ListBox lstSubclass 
         Height          =   1635
         ItemData        =   "fTestSubclass.frx":0000
         Left            =   120
         List            =   "fTestSubclass.frx":0002
         TabIndex        =   12
         Top             =   990
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Caption         =   "After original WndProc"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1815
         TabIndex        =   11
         Top             =   1920
         Width           =   1950
      End
      Begin VB.PictureBox picContainer 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   465
         Index           =   1
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   1665
         TabIndex        =   8
         Top             =   2205
         Width           =   1665
         Begin VB.OptionButton opt 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   0
            TabIndex        =   9
            Top             =   255
            Width           =   1695
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "Before original WndProc"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1770
         TabIndex        =   7
         Top             =   990
         Width           =   2025
      End
      Begin VB.PictureBox picContainer 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   465
         Index           =   0
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   1650
         TabIndex        =   4
         Top             =   1245
         Width           =   1650
         Begin VB.OptionButton opt 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   270
            Width           =   1680
         End
         Begin VB.OptionButton opt 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lvwMessages 
         Height          =   2400
         Index           =   0
         Left            =   3825
         TabIndex        =   15
         Top             =   270
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4233
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4075
         EndProperty
      End
      Begin MSComctlLib.ListView lvwMessages 
         Height          =   2400
         Index           =   1
         Left            =   6465
         TabIndex        =   16
         Top             =   255
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4233
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4075
         EndProperty
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         Caption         =   "Before                   After"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   17
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   4380
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuItm 
         Caption         =   "&New Instance"
         Index           =   0
      End
      Begin VB.Menu mnuItm 
         Caption         =   "&Close"
         Index           =   1
      End
      Begin VB.Menu mnuItm 
         Caption         =   "&End"
         Index           =   2
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iSubclass

Private Enum eChk
    chkSubclassBefore
    chkSubclassAfter
End Enum

Private Enum eCmd
    cmdAddSubclass
    cmdDelSubclass
End Enum

Private Enum eLvw
    lvwSubclassBefore
    lvwSubclassAfter
End Enum

Private Enum eOpt
    optAllBefore
    optSelBefore
    optAllAfter
    optSelAfter
End Enum

Private miTextHeight        As Long
Private mbLoaded            As Boolean

Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, ByVal lprcScroll As Long, ByVal lprcClip As Long, ByVal hrgnUpdate As Long, ByVal lprcUpdate As Long) As Long

'This flag is set when changing item checked states in the subclass listviews.
'It prevents adding or removing messages when the ItemCheck events are fired.
Private mbFreezeMessages    As Boolean

'############################
'##    Event Procedures    ##
'############################

Private Sub chk_Click(Index As Integer)
On Error GoTo handler
    Dim lbVal As Boolean
    Dim lhWnd As Long
    Dim lsClass As String
    
    lbVal = (chk(Index) = vbChecked)
    Select Case CLng(Index)
    Case chkSubclassBefore
        pDeselect lvwMessages(lvwSubclassBefore)
        lvwMessages(lvwSubclassBefore).Enabled = False
        
        opt(optAllBefore).Enabled = lbVal
        opt(optSelBefore).Enabled = lbVal
        opt(optAllBefore).Value = False
        opt(optSelBefore).Value = False
        
        If lstSubclass.ListIndex > -1& Then
            Subclasses(Me).Item(HexVal(lstSubclass.Text)).DelMsg ALL_MESSAGES, MSG_BEFORE
        End If
        
    Case chkSubclassAfter
    
        pDeselect lvwMessages(lvwSubclassAfter)
        lvwMessages(lvwSubclassAfter).Enabled = False
        
        opt(optAllAfter).Enabled = lbVal
        opt(optSelAfter).Enabled = lbVal
        opt(optAllAfter).Value = False
        opt(optSelAfter).Value = False
    
        If lstSubclass.ListIndex > -1& Then
            Subclasses(Me).Item(HexVal(lstSubclass.Text)).DelMsg ALL_MESSAGES, MSG_AFTER
        End If
    
    End Select

    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error GoTo handler
    Dim lhWnd As Long
    Dim liId As Long
    Dim liInterval As Long
    Dim lsMsg As String
    Dim lsClass As String
    
    
    Select Case CLng(Index)
    Case cmdAddSubclass
        fFindWindow.GetWindow Me
    Case cmdDelSubclass
        With lstSubclass
            lhWnd = HexVal(.Text)
            If lhWnd Then
                If MsgBox("Delete this subclass?   " & FmtHex(lhWnd), vbYesNo + vbQuestion) = vbYes Then
                    Subclasses(Me).Remove lhWnd
                    .RemoveItem .ListIndex
                    lstSubclass_Click
                End If
            End If
        End With
    End Select
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub Form_Load()
  giTestForms = giTestForms + 1&
    mbLoaded = True
  Dim i As eMsg
  Dim s As String

    lblHeader.Caption = "######## When.. lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... "

    Dim loLV As ListView

    miTextHeight = picDisplay.TextHeight("M")

  'Adjust the height of the window... like the IntegralHeight property in a listbox
  'Height = Height - (((Me.ScaleHeight Mod nTxtHeight) - 2) * Screen.TwipsPerPixelY)

  For i = 0 To &H400
    s = GetMsgName(i)
    If Asc(s) <> vbKey0 Then
      For Each loLV In lvwMessages
        loLV.ListItems.Add , "k" & i, s
      Next
    End If
  Next i

  For Each loLV In lvwMessages
    loLV.Sorted = True
  Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mbLoaded = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    Dim liOffset As Long
    liOffset = picDisplayHeader.Top + picDisplayHeader.Height
    picDisplay.Move 0, liOffset, ScaleWidth, ScaleHeight - liOffset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mbLoaded = False
    Subclasses(Me).Clear
        
    giTestForms = giTestForms - 1&
    
    If giTestForms = 0& Then
        Unload fFindWindow
    End If

    'FadeOut hwnd
End Sub


Private Sub lstSubclass_Click()
On Error GoTo handler
    pShowMessages
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub lvwMessages_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
On Error GoTo handler
    
    If Not mbFreezeMessages Then
        Dim lhWnd As Long
        Dim liMsg As eMsg
        Dim liWhen As eMsgWhen
        Dim lbVal As Boolean
        Dim lsClass As String
        
        lbVal = Item.Checked
        liMsg = Val(Right$(Item.Key, Len(Item.Key) - 1))
        
        Select Case CLng(Index)
        Case lvwSubclassBefore, lvwSubclassAfter
            liWhen = IIf(CLng(Index) = lvwSubclassBefore, MSG_BEFORE, MSG_AFTER)
            With Subclasses(Me).Item(HexVal(lstSubclass.Text))
                If lbVal _
                    Then .AddMsg liMsg, liWhen _
                    Else .DelMsg liMsg, liWhen
            End With
        End Select
    End If
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub mnuItm_Click(Index As Integer)
    If Index = 0 Then
        Dim temp As New fTest
        temp.Show
    ElseIf Index = 1 Then
        Unload Me
    Else
        If MsgBox("Execute an ""End"" Statement to return to design mode?", vbYesNo) = vbYes Then End
    End If
End Sub

Private Sub opt_Click(Index As Integer)
On Error GoTo handler
    
    pMessageOption Choose(Index + 1, _
                      MSG_BEFORE, MSG_BEFORE, _
                      MSG_AFTER, MSG_AFTER), _
               Index = optAllBefore Or Index = optAllAfter
    
    Exit Sub
handler:
pDisplayErr
End Sub

'############################
'##  Private Procedures    ##
'############################

Private Sub pDisplay(ByRef sString As String)

  If pTextIsBelow Then
    With picDisplay
        ScrollDC .hdc, 0, -miTextHeight, 0&, 0&, 0&, 0&
        Do Until Not pTextIsBelow
            .CurrentY = .CurrentY - miTextHeight
        Loop
    End With
   End If
   
   picDisplay.Print sString

End Sub

Private Function pTextIsBelow() As Boolean
    pTextIsBelow = (picDisplay.CurrentY + miTextHeight + miTextHeight) >= picDisplay.ScaleHeight
End Function

Private Sub pShowMessagesSub(ByRef iArray() As Long, ByVal iCount As Long, ByVal oChk As CheckBox, ByVal oLvw As ListView, Optional ByVal oOptAll As OptionButton, Optional ByVal oOptSel As OptionButton)
    If oOptAll Is Nothing Or oOptSel Is Nothing Then
        If iCount = -1& Then oChk.Value = vbChecked Else oChk.Value = vbUnchecked
    Else
        If iCount > 0& Or iCount = -1& Then oChk.Value = vbChecked Else oChk.Value = vbUnchecked
        oOptAll.Value = (iCount = -1&)
        oOptSel.Value = (iCount > 0&)
    End If
    pSelectItems oLvw, iArray, iCount
End Sub

Private Sub pEnableMessages(ByVal bVal As Boolean)
    pEnableMessagesSub bVal, chk(chkSubclassAfter), lvwMessages(lvwSubclassAfter), opt(optAllAfter), opt(optSelAfter)
    pEnableMessagesSub bVal, chk(chkSubclassBefore), lvwMessages(lvwSubclassBefore), opt(optAllBefore), opt(optSelBefore)
End Sub

Private Sub pEnableMessagesSub(ByVal bVal As Boolean, ByVal oChk As CheckBox, ByVal oLvw As ListView, Optional ByVal oOptAll As OptionButton, Optional ByVal oOptSel As OptionButton)
    Dim lbVal As Boolean
    
    oChk.Enabled = bVal
    If Not bVal Then oChk.Value = vbUnchecked
    If oOptAll Is Nothing Or oOptSel Is Nothing Then
        lbVal = (oChk.Value = vbUnchecked)
    Else
        lbVal = (oChk.Value = vbChecked)
        'If lbVal Then oOptAll.Value = False: oOptSel.Value = False
        oOptAll.Enabled = lbVal: oOptSel.Enabled = lbVal
        lbVal = lbVal And oOptSel.Value
    End If
    oLvw.Enabled = bVal And lbVal
End Sub

Private Sub pDeselect(ByVal lv As ListView)
    Dim itm As MSComctlLib.ListItem
    Dim bWasFroze As Boolean
    
    bWasFroze = mbFreezeMessages
    mbFreezeMessages = True
    
    For Each itm In lv.ListItems
        itm.Checked = False
    Next
    
    mbFreezeMessages = bWasFroze
End Sub

Private Sub pDisplayErr()
    MsgBox "Error #: " & Err.Number & vbNewLine & "Source: " & Err.Source & vbNewLine & vbNewLine & Err.Description
End Sub

Private Sub pSelectItems(ByVal lv As ListView, ByRef iArray() As Long, ByVal iCount As Long)
    On Error Resume Next
    pDeselect lv
    With lv.ListItems
        For iCount = 0 To iCount - 1&
            .Item("k" & iArray(iCount)).Checked = True
        Next
    End With
End Sub

Private Sub pMessageOption(ByVal iWhen As eMsgWhen, ByVal bAll As Boolean)
    Dim lv As ListView
    Set lv = lvwMessages(iWhen And Not MSG_BEFORE)
    pDeselect lv
    lv.Enabled = Not bAll
    
    If Not mbFreezeMessages Then
        With Subclasses(Me).Item(HexVal(lstSubclass.Text))
            If bAll Then
                .AddMsg ALL_MESSAGES, iWhen
            Else
                .DelMsg ALL_MESSAGES, iWhen
            End If
        End With
    End If
End Sub

Private Sub pShowMessages()
    Dim lhWnd As Long
    Dim iCount As Long
    Dim iMessages() As Long
    Dim bEnabled As Boolean
    
    On Error Resume Next
    
    mbFreezeMessages = True

    lhWnd = HexVal(lstSubclass.Text)
    If lhWnd Then
        With Subclasses(Me).Item(lhWnd)
            iCount = .GetMessages(iMessages, MSG_BEFORE)
            pShowMessagesSub iMessages, iCount, chk(chkSubclassBefore), lvwMessages(lvwSubclassBefore), opt(optAllBefore), opt(optSelBefore)

            iCount = .GetMessages(iMessages, MSG_AFTER)
            pShowMessagesSub iMessages, iCount, chk(chkSubclassAfter), lvwMessages(lvwSubclassAfter), opt(optAllAfter), opt(optSelAfter)
        End With
        bEnabled = True
    Else
        bEnabled = False
    End If
    pEnableMessages bEnabled
    mbFreezeMessages = False
End Sub

'############################
'## Implemented Interfaces ##
'############################

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef hWnd As Long, ByRef uMsg As eMsg, ByRef wParam As Long, ByRef lParam As Long)
Static nMsgNo As Long
    If mbLoaded Then
        'If we try to Display the paint message we'll just cause another paint message... vicious circle.
        If Not ((uMsg = WM_PAINT Or uMsg = WM_ERASEBKGND) And (hWnd = Me.hWnd Or hWnd = picDisplay.hWnd)) Then
            nMsgNo = nMsgNo + 1
            pDisplay FmtHex(nMsgNo) & _
                    IIf(bBefore, "Before ", "After  ") & _
                    FmtHex(lReturn) & _
                    FmtHex(hWnd) & _
                    FmtHex(uMsg) & _
                    FmtHex(wParam) & _
                    FmtHex(lParam) & _
                    GetMsgName(uMsg)

        End If
    End If
End Sub


'############################
'##   Public Procedures    ##
'############################

Public Sub AddSubclass(ByVal ihWnd As Long)
    On Error GoTo handler
    If ihWnd Then
        With Subclasses(Me)
            If Not .Exists(ihWnd) Then
                .Add ihWnd
                With lstSubclass
                    .AddItem "0x" & FmtHex(ihWnd)
                    .ListIndex = .NewIndex
                End With
            Else
                If MsgBox("This window is already being subclassed!" & vbCrLf & vbCrLf & _
                          "Do you want to try another one?", _
                          vbYesNo + vbDefaultButton1 + vbQuestion, _
                          "Window already Subclassed") _
                    = vbYes Then fFindWindow.GetWindow Me
            End If
        End With
    End If
    Exit Sub
handler:
    pDisplayErr
End Sub
