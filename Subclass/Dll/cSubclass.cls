VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cSubclass.cls                          7/5/04
'
'           PURPOSE:
'               Represents a subclass on a single hWnd for a single client object, allowing the dll
'               client to specify which messages are to be received.
'
'           MODULES CALLED FROM THIS MODULE:
'               mSubclass.bas
'
'==================================================================================================

'1.  Friendly Interface       - Used internally to initialize the modular variables.
'2.  Message Interface        - Methods which manage the messages called back by this subclass.
'3.  CallOldWndProc Interface - Calls the original WndProc at any old time you please.

Private miPtr As Long   'Stores a pointer to the iSubclass object to be called back for messages
Private mhWnd As Long   'stores the mhWnd subclassed by this object

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit( _
        ByVal iPtr As Long, _
        ByVal ihWnd As Long)
    mhWnd = ihWnd
    miPtr = iPtr
End Sub
'</Friendly Interface>

'<Public Interface>
'<Message Interface>
Public Function AddMsg( _
            ByVal iMsg As eMsg, _
            ByVal iWhen As eMsgWhen) _
                As Boolean
    AddMsg = mSubclass.MsgHubObject(mhWnd).AddMsg(iMsg, iWhen, miPtr)
End Function

Public Function DelMsg( _
            ByVal iMsg As eMsg, _
            ByVal iWhen As eMsgWhen) _
                As Boolean
    DelMsg = mSubclass.MsgHubObject(mhWnd).DelMsg(iMsg, iWhen, miPtr)
End Function

Public Function MsgExists( _
            ByVal iMsg As eMsg, _
            ByVal iWhen As eMsgWhen) _
                As Boolean
    MsgExists = mSubclass.MsgHubObject(mhWnd).MsgExists(iMsg, iWhen, miPtr)
End Function

Public Function MsgCount( _
            ByVal iWhen As eMsgWhen) _
                As Long
    MsgCount = mSubclass.MsgHubObject(mhWnd).MsgCount(miPtr, iWhen)
End Function

Public Function GetMessages( _
            ByRef iArray() As Long, _
            ByVal iWhen As eMsgWhen) _
                As Long
    GetMessages = mSubclass.MsgHubObject(mhWnd).GetMessages(miPtr, iWhen, iArray)
End Function
'</Message Interface>

'<CallOldWndProc Interface>
Public Function CallOldWndProc( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    CallOldWndProc = mSubclass.MsgHubObject(mhWnd).CallOldWndProc(iMsg, wParam, lParam)
End Function
'</CallOldWndProc Interface>

Public Property Get hWnd() As Long
    hWnd = mhWnd
End Property
'</Public Interface>
