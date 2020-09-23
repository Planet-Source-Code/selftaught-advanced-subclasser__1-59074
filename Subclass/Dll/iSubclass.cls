VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'iSubclass.cls                          8/19/04
'
'           LINEAGE:
'               Paul Caton's WinSubHook.tlb at www.planet-source-code.com
'
'           PURPOSE:
'               Callback procedure for subclasses.
'
'==================================================================================================


Public Sub Proc( _
            ByVal bBefore As Boolean, _
            ByRef bHandled As Boolean, _
            ByRef lReturn As Long, _
            ByRef hWnd As Long, _
            ByRef iMsg As eMsg, _
            ByRef wParam As Long, _
            ByRef lParam As Long)

End Sub
