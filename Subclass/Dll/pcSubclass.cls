VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcSubclass - module-less, IDE safe, machine code subclassing thunk
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.00 20030107 First cut..........................................................................
'v1.01 20031118 Allow control over callback gating
'               Use global memory for the machine code buffer
'               Reform the assembler...............................................................
'v1.02 20040118 Use EbMode for breakpoint/stop detection rather than callback gating
'               Further reform the assembler for greater speed and smaller size
'               Made InIDE public..................................................................
'
'==================================================================================================

'<Utility Api's>
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC As Long = -4
'</Utility Api's>

Private Const PATCH_05  As Long = 93                        'Table B (before) entry count
Private Const PATCH_09  As Long = 137                       'Table A (after) entry count

Private hWndSub         As Long                             'Handle of the window being subclassed
Private nAddrSubclass   As Long                             'The address of our WndProc
Private nAddrOriginal   As Long                             'The address of the existing WndProc

'============================================
'Class creation/destruction
'============================================

'Build the subclass thunk into allocated memory
Private Sub Class_Initialize()
Const PATCH_01 As Long = 18                                 'Code buffer offset to the location of the relative address to EbMode
Const PATCH_03 As Long = 78                                 'Relative address of SetWindowsLong
Const PATCH_07 As Long = 121                                'Relative address of CallWindowProc
Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL As String = "SetWindowLongA"                 'SetWindowLong allows the pcSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const FUNC_CWP As String = "CallWindowProcA"                'We use CallWindowProc to call the original WndProc
Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
Const MOD_USER As String = "user32"                         'Location of the SetWindowLong & CallWindowProc functions
  
  Dim i As Long
  
  nAddrSubclass = AllocASM(asmSubclass)

  If InIDE And Not gbDisableIDEProtection Then
    'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
    Call CopyMemory(ByVal nAddrSubclass + 15, &H9090, 2)
    
    i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll
    If i = 0 Then                                           'Found?
      i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
    End If

    Debug.Assert i                                          'Ensure the EbMode function was found
    Call PatchValueRelative(nAddrSubclass, PATCH_01, i)     'Patch the relative address to the EbMode api function
  End If
  

  Call PatchValueRelative(nAddrSubclass, PATCH_03, _
                          AddrFunc(MOD_USER, FUNC_SWL))     'Address of the SetWindowLong api function
  Call PatchValueRelative(nAddrSubclass, PATCH_07, _
                          AddrFunc(MOD_USER, FUNC_CWP))     'Address of the CallWindowProc api function
  
  Call PatchValue(nAddrSubclass, PATCH_05, 0)               'Initial before table entry count
  Call PatchValue(nAddrSubclass, PATCH_09, 0)               'Initial after table entry count
End Sub

'UnSubclass and release the allocated memory
Private Sub Class_Terminate()
  Call Me.UnSubclass                                        'UnSubclass if the Subclass thunk is active
  Call GlobalFree(nAddrSubclass)                            'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

'Arbitarily call the original WndProc
Public Function CallOrigWndProc(ByVal iMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long
  If hWndSub <> 0 Then
    CallOrigWndProc = CallWindowProc(nAddrOriginal, hWndSub, iMsg, wParam, lParam) 'Call the original WndProc
  End If
End Function

'Set the window subclass
Public Function Subclass(ByVal hWnd As Long, _
                         ByVal Owner As iSubclass) As Boolean
Const PATCH_02 As Long = 68                                 'Address of the previous WndProc
Const PATCH_06 As Long = 116                                'Address of the previous WndProc
Const PATCH_0A As Long = 186                                'Address of the owner object
  
  If hWndSub = 0 Then
    Debug.Assert IsWindow(hWnd)                             'Invalid window handle
    hWndSub = hWnd                                          'Store the window handle
    
    'Get the original window proc
    nAddrOriginal = GetWindowLong(hWnd, GWL_WNDPROC)
    Call PatchValue(nAddrSubclass, PATCH_02, nAddrOriginal) 'Original WndProc address for CallWindowProc, call the original WndProc
    Call PatchValue(nAddrSubclass, PATCH_06, nAddrOriginal) 'Original WndProc address for SetWindowLong, unsubclass on IDE stop
    Call PatchValue(nAddrSubclass, PATCH_0A, ObjPtr(Owner)) 'Owner object address for iSubclass_Proc
    
    'Set our WndProc in place of the original
    nAddrOriginal = SetWindowLong(hWnd, GWL_WNDPROC, nAddrSubclass)
    Subclass = nAddrOriginal <> 0                           'Success
  End If
  
  Debug.Assert Subclass
End Function

'Stop subclassing the window
Public Function UnSubclass() As Boolean
  If hWndSub <> 0 Then
    Call PatchValue(nAddrSubclass, PATCH_05, 0&)            'Patch the Table B entry count to ensure no further iSubclass_Proc callbacks
    Call PatchValue(nAddrSubclass, PATCH_09, 0&)            'Patch the Table A entry count to ensure no further iSubclass_Proc callbacks
    
    'Restore the original WndProc
    Call SetWindowLong(hWndSub, GWL_WNDPROC, nAddrOriginal)
    
    hWndSub = 0                                             'Indicate the subclasser is inactive
    UnSubclass = True                                       'Success
  End If
End Function







'Added 4/10/04 for utility purposes
Public Property Get Active() As Boolean
    Active = hWndSub <> 0
End Property

'4/10/04 His code, rearranged
Public Sub PatchTable(aTable() As Long, iCount As Long, iWhen As eMsgWhen)
    
    Const PATCH_04 As Long = 88                     'Table B (before) address
    Const PATCH_08 As Long = 132                    'Table A (after) address
    
    On Error Resume Next
    Dim nOff1 As Long
    Dim nOff2 As Long
    Dim liAddr As Long
    
    If iWhen = MSG_BEFORE Then
        nOff1 = PATCH_04                            'Offset to the Before table address
        nOff2 = PATCH_05                            'Offset to the Before table entry count
    ElseIf iWhen = MSG_AFTER Then
        nOff1 = PATCH_08                            'Offset to the After table address
        nOff2 = PATCH_09                            'Offset to the After table entry count
    Else
        Exit Sub
    End If

    liAddr = VarPtr(aTable(0))
    
    Call PatchValue(nAddrSubclass, nOff1, liAddr)   'Patch the appropriate table address. We need
                                                    'do this because there's no guarantee that the
                                                    'table existed before this call, the table may
                                                    'only get created if a message is added.
    Call PatchValue(nAddrSubclass, nOff2, iCount)   'Patch the appropriate table entry count

End Sub

Public Property Get hWnd() As Long
    hWnd = hWndSub
End Property

