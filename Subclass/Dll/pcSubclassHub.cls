VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcSubclassHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcSubclassHub.cls                      7/4/04
'
'           PURPOSE:
'               Maintain a subclass and a callback rountine, and deliver notifications to requesting
'               objects.
'
'           MODULES CALLED FROM THIS MODULE:
'               mVbBaseGeneral.bas
'
'           CLASSES CREATED BY THIS MODULE:
'               pcSubclass.cls
'
'==================================================================================================

'1.  Private Interface           - Utility procedures for maintaining private variables
'2.  Public Interface            - procedures to add and remove messages and clients and to start/stop subclassing
'3.  iSubclass Interface         - callback from the asm thunk

Implements iSubclass

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

#Const bVBVMTypeLib = False                              'Toggle between use of the VBVM Type Library for typecasting and more

Private Const DWordBits                 As Long = 32    'don't adjust this one - 32 bits in a long

Private Const MsgMaskCount              As Long = 4     'Adjust this constant up/down to change the limit on messages
                                                        'normally it should be an even power of 2 and not larger than 16
                                                        
Private Const MaxMsgCount               As Long _
                     = MsgMaskCount * DWordBits         'no more than this number of messages will be delivered collectively to
                                                        'all objects in the before or the after table.
                                                        
Private Const ALL_MESSAGES_BEFORE       As Long = 1
Private Const ALL_MESSAGES_AFTER        As Long = 2

Private Type tSubclassClient
    Ptr                                 As Long         'ObjPtr for the iSubclass Interface
    Flags                               As Long         'bits 1 & 2 are for all messages before/after the wndproc
    MsgMaskB(0 To MsgMaskCount - 1&)    As Long         'masks indicating which messages are received before the wndproc
    MsgMaskA(0 To MsgMaskCount - 1&)    As Long         'masks indicating which messages are received after the wndproc
End Type

Private miMsgTableB()                   As Long         'these arrays will never grow beyond MaxMsgCount # of elements
Private miMsgCountB                     As Long

Private miMsgTableA()                   As Long
Private miMsgCountA                     As Long

Private miAllClientsB                   As Long         'count the number of clients who have requested all messages
Private miAllClientsA                   As Long         'before and after the wndproc

Private mtClients()                     As tSubclassClient  'client data store
Private miClientCount                   As Long         'number of clients defined in the array

Private moSubclass                      As pcSubclass   'worker - asm thunk

'<Private Interface>
Private Sub Class_Initialize()
    Set moSubclass = New pcSubclass                      'Initialize the worker
End Sub

Private Sub Class_Terminate()
    Debug.Assert Active = False                          'Just checking!
    Set moSubclass = Nothing                             'term the worker
End Sub

Private Function Client( _
            ByVal iIndex As Long) _
                 As iSubclass
    Dim loTemp As iSubclass                     'Temp object variable
    Dim liPtr As Long                           'Pointer to the client object
    liPtr = mtClients(iIndex).Ptr               'Get the objptr
    If liPtr <> 0& And liPtr <> Undefined Then  'If the pointer is valid
        #If bVBVMTypeLib Then
            ObjectPtr(loTemp) = liPtr           'If we're using the VBVM type lib then set the objptr directly
            Set Client = loTemp
            ObjectPtr(loTemp) = 0&
        #Else
            CopyMemory loTemp, liPtr, 4&        'If we're not using the type lib then use the std copymem
            Set Client = loTemp
            CopyMemory loTemp, 0&, 4&
        #End If
    End If
End Function

Private Function pAddMsg( _
            ByVal iMsg As eMsg, _
            ByRef iMsgTbl() As Long, _
            ByRef iMsgCnt As Long, _
            ByVal iWhen As eMsgWhen, _
            ByVal iClientPtr As Long) _
                As Boolean
    
    Dim liClientIndex       As Long                                     'Store the index of the clients in miClients
    Dim liMsgIndex          As Long                                     'store the index of the msg in the main msg table
    Dim liClientMsgIndex    As Long                                     'index in tSubclassClient's array
    Dim liClientMsgMask     As Long                                     'a bit identifying a message
    Dim liTempMask          As Long                                     'store the client's current message mask at the index that we need
    Dim liFirst             As Long                                     'first available index in iMsgTbl
    
    liClientIndex = pFindClient(iClientPtr)                             'Look for the client in our client table
    
    If liClientIndex <> Undefined Then                                  'If the client is in the table
        
        If iMsg <> ALL_MESSAGES Then                                    'If we're adding a single message
            liMsgIndex = ArrFindInt(iMsgTbl, iMsgCnt, iMsg, liFirst)    'Find the index of the message in the table
    
            If liMsgIndex = Undefined Then                              'If the message is not already in the table
                If liFirst > Undefined Then                             'If an open slot exists in the table
                    liMsgIndex = liFirst                                'Fill that slot with this message
                Else
                    If iMsgCnt = MaxMsgCount Then Exit Function         'don't go past the max!
                    liMsgIndex = iMsgCnt                                'Next index will be equal to the current count
                    iMsgCnt = iMsgCnt + 1&                              'Increase the msg count
                    ArrRedim iMsgTbl, iMsgCnt, True                     'Make sure we have enough room in the array
                End If
                iMsgTbl(liMsgIndex) = iMsg                              'Store this message in the table
            End If
            
            pTranslateMsgIndex liMsgIndex, liClientMsgIndex, liClientMsgMask 'get the index and bit that identifies this message
            
            With mtClients(liClientIndex)
                If iWhen = MSG_BEFORE Then                              'if adding the msg to the before table
                    liTempMask = .MsgMaskB(liClientMsgIndex)            'store the current mask at this index
                    pAddMsg = Not CBool(liTempMask And liClientMsgMask) 'indicate success if the message is not already being received
                    .MsgMaskB(liClientMsgIndex) = _
                        liTempMask Or liClientMsgMask                   'store that this client has requested this message
                ElseIf iWhen = MSG_AFTER Then                           'else - we're adding the msg to the after table
                    liTempMask = .MsgMaskA(liClientMsgIndex)            'store the current mask at this index
                    pAddMsg = Not CBool(liTempMask And liClientMsgMask)
                    .MsgMaskA(liClientMsgIndex) = _
                        liTempMask Or liClientMsgMask
                End If
            End With
            
            'Message was already there!
            Debug.Assert pAddMsg
            
        Else
            With mtClients(liClientIndex)
                If iWhen = MSG_BEFORE Then                                  'We are adding ALL_MESSAGES to the before table
                    If Not CBool(.Flags And ALL_MESSAGES_BEFORE) Then       'If the client has not already requested all msgs
                        pAddMsg = True                                      'Indicate success
                        .Flags = .Flags Or ALL_MESSAGES_BEFORE              'store that this client has requested all message before
                        miAllClientsB = miAllClientsB + 1&                  'inc the count of objects that've requested all messages
                    Else
                        'This client is already getting all messages before
                        Debug.Assert False
                    End If
                ElseIf iWhen = MSG_AFTER Then                               'We are adding ALL_MESSAGES to the after table
                    If Not CBool(.Flags And ALL_MESSAGES_AFTER) Then        'If the client has not already requested all msgs
                        pAddMsg = True                                      'Indicate success
                        .Flags = .Flags Or ALL_MESSAGES_AFTER               'store that this client has requested all message After
                        miAllClientsA = miAllClientsA + 1&                  'inc the count of objects that've requested all messages
                    Else
                        'This client is already getting all messages After
                        Debug.Assert False
                    End If
                End If
            End With
        End If
        
        PatchTable iMsgTbl, iMsgCnt, iWhen                              'Patch the new table info
    
    Else
        'Adding a message for a client that wasn't found!
        Debug.Assert False
    End If

End Function

Private Sub PatchTable( _
                ByRef aTbl() As Long, _
                ByVal iCnt As Long, _
                ByVal iWhen As eMsgWhen)
    
    If iWhen = MSG_BEFORE Then
        If miAllClientsB > 0& Then iCnt = -1& 'If one or more clients have requested all messages, then count = -1&
    ElseIf iWhen = MSG_AFTER Then
        If miAllClientsA > 0& Then iCnt = -1& 'If one or more clients have requested all messages, then count = -1&
    Else
        Exit Sub
    End If
    
    moSubclass.PatchTable aTbl, iCnt, iWhen    'Ensure that the ASM thunk has the correct table address and count
End Sub

Private Function pDelMsg( _
            ByVal iMsg As eMsg, _
            ByRef iMsgTbl() As Long, _
            ByRef iMsgCnt As Long, _
            ByVal iWhen As eMsgWhen, _
            ByVal iClientPtr As Long) _
                 As Boolean
    
    Dim liClientIndex      As Long                              'The client index in mtClients
    Dim liMsgIndex         As Long                              'index of the message in iMsgTbl
    Dim liClientMsgIndex As Long                                'index in mtClients(liClientIndex).MsgTable (A or B)
    Dim liClientMsgMask  As Long                                'mask identifying this message in the table
    Dim liTempMask         As Long
    
    liClientIndex = pFindClient(iClientPtr)                     'Find the index of the client
    
    If liClientIndex <> Undefined Then                          'If the client's index was found
    
        If iMsg <> ALL_MESSAGES Then                            'If we're adding a single message
            liMsgIndex = ArrFindInt(iMsgTbl, iMsgCnt, iMsg)     'See if the message exists in the table
            
            If liMsgIndex <> Undefined Then                     'if the message exists
                                                                'get the array index and bit mask for our MsgMask
                pTranslateMsgIndex liMsgIndex, liClientMsgIndex, liClientMsgMask
                
                With mtClients(liClientIndex)
                    If iWhen = MSG_BEFORE Then                   'If we're working with the before table
                        liTempMask = .MsgMaskB(liClientMsgIndex) 'Store the current mask
                        pDelMsg = CBool(liTempMask And liClientMsgMask) 'succeed only if subscribed to this msg
                        .MsgMaskB(liClientMsgIndex) = _
                            liTempMask And Not liClientMsgMask   'remove the bit from the MsgMask
                    ElseIf iWhen = MSG_AFTER Then
                        liTempMask = .MsgMaskA(liClientMsgIndex) 'Store the current mask
                        pDelMsg = CBool(liTempMask And liClientMsgMask) 'succeed only if subscribed to this msg
                        .MsgMaskA(liClientMsgIndex) = _
                            liTempMask And Not liClientMsgMask   'remove the bit from the MsgMask
                    End If
                End With
                
                If pDelMsg Then                                  'if we succeeded then we may have a message that's not delivered to any object
                    If pMsgIsDead(liClientMsgIndex, liClientMsgMask, iWhen) Then _
                        iMsgTbl(liMsgIndex) = Undefined
                End If
                
            End If
            
        Else
            If iWhen = MSG_BEFORE Then
                pDelMsg = CBool(mtClients(liClientIndex).Flags And ALL_MESSAGES_BEFORE)
            ElseIf iWhen = MSG_AFTER Then
                pDelMsg = CBool(mtClients(liClientIndex).Flags And ALL_MESSAGES_AFTER)
            End If
            pDelMsg = pDelMsg Or CBool(pDelAllMessages(liClientIndex, iWhen))
        End If
        
        For iMsgCnt = iMsgCnt - 1& To 0& Step -1&               'Loop backwards to find the last defined msg
            If iMsgTbl(iMsgCnt) <> Undefined Then Exit For
        Next
        
        iMsgCnt = iMsgCnt + 1&                                  'Store the 1-based count, not 0-based index
        
        PatchTable iMsgTbl, iMsgCnt, iWhen                      'Patch the new table info
    Else
        'Deleting a message for a client that wasn't found!
        Debug.Assert False
    End If
End Function

Private Function pDelAllMessages( _
                ByVal iClientIndex As Long, _
                ByVal iWhen As eMsgWhen) _
                    As Long

    Dim liMask     As Long      'a mask containing each bit one at a time
    Dim liTempMask As Long      'the current msgmask
    Dim iEachBit   As Long      'counter
    Dim iEachIndex As Long      'counter
    
    With mtClients(iClientIndex)
        
        If iWhen = MSG_BEFORE Then                                  'if deleting from the before table
            
            If CBool(.Flags And ALL_MESSAGES_BEFORE) Then
                miAllClientsB = miAllClientsB - 1&                      'maintain the count of clients requesting ALL_MESSAGES
                .Flags = .Flags And Not ALL_MESSAGES_BEFORE
            End If
            
            For iEachBit = 0 To BitMax                              'loop through each bit
                liMask = BitMask(iEachBit)                          'store the bitmask for this bit
                For iEachIndex = 0 To MsgMaskCount - 1&             'for each bit, loop through each index
                    liTempMask = .MsgMaskB(iEachIndex)              'store the msgmask at this index
                    If CBool(liTempMask And liMask) Then            'if the client has requested this message
                        pDelAllMessages = pDelAllMessages + 1&      'inc the count of messages removed
                        .MsgMaskB(iEachIndex) = liTempMask And Not liMask 'remove the message from the client's table
                        If pMsgIsDead(iEachIndex, liMask, iWhen) Then _
                            miMsgTableB(iEachIndex * DWordBits + iEachBit) = Undefined 'If nobody wants the message anymore, remove it from the table
                    End If
                Next
            Next
        ElseIf iWhen = MSG_AFTER Then                               'if deleting from the before table
            
            If CBool(.Flags And ALL_MESSAGES_AFTER) Then
                miAllClientsA = miAllClientsA - 1&
                .Flags = .Flags And Not ALL_MESSAGES_AFTER
            End If
            
            For iEachBit = 0 To BitMax                              'loop through each bit
                liMask = BitMask(iEachBit)                          'store the bitmask for this bit
                For iEachIndex = 0 To MsgMaskCount - 1&             'for each bit, loop through each index
                    liTempMask = .MsgMaskA(iEachIndex)              'store the msgmask at this index
                    If CBool(liTempMask And liMask) Then            'if the client has requested this message
                        pDelAllMessages = pDelAllMessages + 1&      'inc the count of messages removed
                        .MsgMaskA(iEachIndex) = liTempMask And Not liMask 'remove the message from the client's table
                        If pMsgIsDead(iEachIndex, liMask, iWhen) Then _
                            miMsgTableA(iEachIndex * DWordBits + iEachBit) = Undefined 'If nobody wants the message anymore, remove it from the table
                    End If
                Next
            Next
        End If
    End With

End Function

Private Function pMsgIsDead( _
                ByVal iTableIndex As Long, _
                ByVal iMask As Long, _
                ByVal iWhen As eMsgWhen) _
                    As Boolean

    Dim iEachClient As Long
    
    If iWhen = MSG_BEFORE Then                                                  'If checking the before table
        For iEachClient = 0 To miClientCount - 1&                               'loop through each client
            If mtClients(iEachClient).Ptr <> Undefined Then                     'if the client is defined
                If CBool(mtClients(iEachClient).MsgMaskB(iTableIndex) And iMask) Then _
                    Exit Function                                               'if the client wants this message, bail
            End If
        Next
    ElseIf iWhen = MSG_AFTER Then                                               'if checking the after table
        For iEachClient = 0 To miClientCount - 1&                               'loop through each client
            If mtClients(iEachClient).Ptr <> Undefined Then                     'if the client is defined
                If CBool(mtClients(iEachClient).MsgMaskA(iTableIndex) And iMask) Then _
                    Exit Function                                               'if the client wants this message then bail
            End If
        Next
    End If
    
    pMsgIsDead = True                                                           'if we made it here, nobody cares about this message
End Function
                    
Private Function pFindClient( _
             ByVal iPtr As Long, _
    Optional ByRef iFirstAvailable As Long) _
                 As Long
    
    Dim liPtr As Long
    
    iFirstAvailable = Undefined                                          'Initialize the first available slot to nothing
    
    For pFindClient = 0& To miClientCount - 1&                           'Loop through each client
        liPtr = mtClients(pFindClient).Ptr                               'Store this client's ptr
        If liPtr <> 0& And liPtr <> Undefined Then                       'If the ptr is defined
            If liPtr = iPtr Then Exit Function                           'If the ptr matches then bail
        Else
            If iFirstAvailable = Undefined Then _
                        iFirstAvailable = pFindClient                    'If the ptr is undefined and the first available has not already been set, then this index is the first available
        End If
    Next
    
    pFindClient = Undefined                                              'If we made it here, then the client was not found
End Function

Private Sub pTranslateMsgIndex( _
            ByVal iMsgTableIndex As Long, _
            ByRef iClientTableIndex As Long, _
            ByRef iMask As Long)

    iClientTableIndex = iMsgTableIndex \ DWordBits                      'get the index in the tSubclassClient msg array
    iMask = BitMask(iMsgTableIndex Mod DWordBits)                       'get the mask that identifies this message
    
    'internal error
    Debug.Assert iClientTableIndex < MsgMaskCount                       'alert developer in case we've exceeded the max
End Sub

'private implementation of ArrRedim to allow strong typing
Private Sub ArrRedimT( _
            ByRef tArray() As tSubclassClient, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundT(tArray) Then
        If bPreserve Then _
            ReDim Preserve tArray(0 To liNewUbound) _
        Else _
            ReDim tArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundT( _
            ByRef tArray() As tSubclassClient) _
                As Long
    On Error Resume Next
    ArrUboundT = UBound(tArray)
    If Err.Number <> 0& Then ArrUboundT = Undefined
End Function
'</Private Interface>

'<Public Interface>
Public Function AddClient( _
            ByVal iWho As Long) _
                As Boolean

    Dim liFirst As Long
    Dim liIndex As Long
    
    liIndex = pFindClient(iWho, liFirst)    'Must call this sub to search through all the clients
                                            'In case it's already there
    
    If liIndex = Undefined Then             'If this client was not found
        AddClient = True                    'then we can add it
        If liFirst = Undefined Then         'if there is no open slot
            liFirst = miClientCount         'next index is current count
            miClientCount = miClientCount + 1& 'inc the count
            ArrRedimT mtClients, miClientCount, True 'resize the array
        End If
        
        If liFirst > 0& Then                'if we're not already adding to the zeroth index
            CopyMemory mtClients(1), _
                       mtClients(0), _
                       LenB(mtClients(0)) * liFirst 'push all the clients up one notch
            ZeroMemory mtClients(0), LenB(mtClients(0)) 'zero out the zeroth index
        End If
        
        mtClients(0).Ptr = iWho             'init the client on the zeroth element
    End If

End Function

Public Function CallOldWndProc( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    CallOldWndProc = moSubclass.CallOrigWndProc(iMsg, wParam, lParam)   'delegate
End Function

Public Function ClientExists( _
            ByVal iWho As Long) _
                As Boolean
    ClientExists = pFindClient(iWho) <> Undefined
End Function

Public Function DelClient( _
            ByVal iWho As Long) _
                As Boolean

    Dim liIndex As Long
    
    liIndex = pFindClient(iWho)         'try to find the client index
    
    If liIndex <> Undefined Then        'if the client index was found
        
        DelClient = True                'then we can delete it
        
        pDelAllMessages liIndex, MSG_AFTER  'remove all messages that this client has asked for
        pDelAllMessages liIndex, MSG_BEFORE
        
        mtClients(liIndex).Ptr = Undefined  'uninit the client on this element
        
        If liIndex = miClientCount - 1& Then 'loop backwards to find the lowest possible value for miClienCount
            For liIndex = liIndex - 1& To 0& Step -1&
                If mtClients(liIndex).Ptr <> Undefined Then Exit For
            Next
            miClientCount = liIndex + 1&
        End If

    End If
    
End Function

Public Function AddMsg( _
            ByVal iMsg As eMsg, _
            ByVal iWhen As eMsgWhen, _
            ByVal iWho As Long) _
                As Boolean
    
    
    If iWhen And MSG_BEFORE Then _
        AddMsg = pAddMsg(iMsg, miMsgTableB, miMsgCountB, MSG_BEFORE, iWho)          'Add Msg to the before table
    
    If iWhen And MSG_AFTER Then _
        AddMsg = AddMsg Or pAddMsg(iMsg, miMsgTableA, miMsgCountA, MSG_AFTER, iWho) 'Add Msg to the after table

End Function

Public Function DelMsg( _
            ByVal iMsg As eMsg, _
            ByVal iWhen As eMsgWhen, _
            ByVal iWho As Long) _
                As Boolean
    
    If iWhen And MSG_BEFORE Then _
        DelMsg = pDelMsg(iMsg, miMsgTableB, _
                         miMsgCountB, MSG_BEFORE, iWho)          'Del Msg from the before table

    If iWhen And MSG_AFTER Then _
        DelMsg = DelMsg Or pDelMsg(iMsg, miMsgTableA, _
                                   miMsgCountA, MSG_AFTER, iWho) 'Del Msg from the after table

End Function
                  
Public Function MsgCount( _
            ByVal iWho As Long, _
            ByVal iWhen As eMsgWhen) _
                As Long
    
    Dim liClientIndex As Long
    Dim liEachBit As Long
    Dim liEachIndex As Long
    
    
    liClientIndex = pFindClient(iWho)                                           'try to find the client index
    
    If liClientIndex <> Undefined Then                                          'if the client index was found
        With mtClients(liClientIndex)
            If iWhen = MSG_BEFORE Then                                          'if counting the before table
                For liEachIndex = 0 To MsgMaskCount - 1&                        'loop through each message
                    For liEachBit = 0 To BitMax                                 'loop though all bits
                        If liEachIndex * DWordBits + liEachBit = miMsgCountB Then Exit Function 'if we're past the last message, bail
                        If CBool(.MsgMaskB(liEachIndex) And BitMask(liEachBit)) Then _
                            MsgCount = MsgCount + 1&                            'if the client wants this message, inc the count
                    Next
                Next
            ElseIf iWhen = MSG_AFTER Then                                       'if counting the after table
                For liEachIndex = 0 To MsgMaskCount - 1&                        'loop through each message
                    For liEachBit = 0 To BitMax
                        If liEachIndex * DWordBits + liEachBit = miMsgCountA Then Exit Function 'if we're past the last message, bail
                        If CBool(.MsgMaskA(liEachIndex) And BitMask(liEachBit)) Then _
                            MsgCount = MsgCount + 1&                            'if the client wants this message, inc the count
                    Next
                Next
            End If
        End With
    End If
End Function

Public Function MsgExists( _
            ByVal iMsg As eMsg, _
            ByVal iWhen As eMsgWhen, _
            ByVal iWho As Long) _
                As Boolean
    
    Dim liClientIndex As Long
    Dim liMsgIndex As Long
    Dim liClientMsgIndex As Long
    Dim liClientMsgMask As Long
    
    liClientIndex = pFindClient(iWho)                                   'Try to find the client
    
    If liClientIndex <> Undefined Then                                  'if the client was found
        If iWhen = MSG_BEFORE Then                                      'if checking the before table
            If iMsg <> ALL_MESSAGES Then                                'if checking a specific message
                liMsgIndex = ArrFindInt(miMsgTableB, miMsgCountB, iMsg) 'locate the message in the table
                If liMsgIndex <> Undefined Then                         'if the message was found
                    pTranslateMsgIndex liMsgIndex, liClientMsgIndex, liClientMsgMask 'get the MsgMask Index
                    MsgExists = _
                        CBool(mtClients(liClientIndex).MsgMaskB(liClientMsgIndex) _
                                 And _
                              liClientMsgMask)                          'exists if the bitmask says so
                End If
            Else                                                        'if we're looking for ALL_MESSAGES
                MsgExists = _
                    CBool(mtClients(liClientIndex).Flags _
                            And _
                          ALL_MESSAGES_BEFORE)                          'exists if the flag is set
            End If
            If Not MsgExists Then Exit Function
        End If
        
        If iWhen = MSG_AFTER Then                                       'if checking the before table
            If iMsg <> ALL_MESSAGES Then                                'if checking a specific message
                liMsgIndex = ArrFindInt(miMsgTableA, miMsgCountA, iMsg) 'locate the message in the table
                If liMsgIndex <> Undefined Then                         'if the message was found
                    pTranslateMsgIndex liMsgIndex, liClientMsgIndex, liClientMsgMask 'get the MsgMask index
                    MsgExists = _
                        CBool(mtClients(liClientIndex).MsgMaskB(liClientMsgIndex) _
                                And _
                             liClientMsgMask)                           'exists if the bitmask says so
                End If
            Else                                                        'if we're looking for ALL_MESSAGES
                MsgExists = _
                    CBool(mtClients(liClientIndex).Flags _
                            And _
                          ALL_MESSAGES_AFTER)                           'exists if the flag is set
            End If
        End If

    End If

End Function

Public Function Subclass( _
            ByVal hWnd As Long) _
                As Boolean
    
    Debug.Assert moSubclass.Active = False      'Shouldn't happen, but just in case
    Subclass = moSubclass.Subclass(hWnd, Me)    'start the subclass
    
End Function

Public Function GetMessages( _
            ByVal iWho As Long, _
            ByVal iWhen As eMsgWhen, _
            ByRef iArray() As Long) _
                As Long
    
    Dim liClientIndex As Long
    Dim liEachBit As Long
    Dim liEachIndex As Long
    
    liClientIndex = pFindClient(iWho)                                                   'try to find the client

    If liClientIndex <> Undefined Then                                                  'if the client was found
        With mtClients(liClientIndex)
            If iWhen = MSG_BEFORE Then                                                  'if we're looking in the before table
                If Not CBool(.Flags And ALL_MESSAGES_BEFORE) Then                       'if the client does not get ALL_MESSAGES
                    For liEachIndex = 0& To MsgMaskCount - 1&                           'loop through all MsgMask Indexes
                        For liEachBit = 0& To BitMax                                    'loop through each bit
                            If liEachIndex * DWordBits + liEachBit = miMsgCountB Then GoTo GetOuttaHere 'if we're past the last message, then bail
                            If CBool(.MsgMaskB(liEachIndex) And BitMask(liEachBit)) Then 'if the client gets this message
                                GetMessages = GetMessages + 1&                           'inc the count
                                ArrRedim iArray, GetMessages, True                       'resize the array if necessary
                                iArray(GetMessages - 1&) = miMsgTableB(liEachIndex * DWordBits + liEachBit) 'store the message
                            End If
                        Next
                    Next
                Else
                    GetMessages = -1&                                                   'client gets all messages, return -1&
                End If
            ElseIf iWhen = MSG_AFTER Then                                               'if we're looking in the after table
                If Not CBool(.Flags And ALL_MESSAGES_AFTER) Then                        'if the client does not get ALL_MESSAGES
                    For liEachIndex = 0& To MsgMaskCount - 1&                           'loop through all MsgMask Indexes
                        For liEachBit = 0& To BitMax                                    'loop through each bit
                            If liEachIndex * DWordBits + liEachBit = miMsgCountA Then GoTo GetOuttaHere 'if we're past the last message, then bail
                            If CBool(.MsgMaskA(liEachIndex) And BitMask(liEachBit)) Then 'if the client gets this message
                                GetMessages = GetMessages + 1&                           'inc the count
                                ArrRedim iArray, GetMessages, True                       'resize the array if necessary
                                iArray(GetMessages - 1&) = miMsgTableA(liEachIndex * DWordBits + liEachBit) 'store the message
                            End If
                        Next
                    Next
                Else
                    GetMessages = -1&
                End If
            End If
        End With
    End If
    
GetOuttaHere:
    If GetMessages > 0& Then
        ReDim Preserve iArray(0 To GetMessages - 1)     'trim the excess
    Else
        Erase iArray
    End If
    
End Function

Public Property Get hWnd() As Long
    hWnd = moSubclass.hWnd
End Property

Public Property Get Active() As Boolean
    Active = miClientCount > 0&
End Property
'</Public Interface>


'<iSubclass Interface>
Private Sub iSubclass_Proc( _
                ByVal bBefore As Boolean, _
                ByRef bHandled As Boolean, _
                ByRef lReturn As Long, _
                ByRef hWnd As Long, _
                ByRef iMsg As eMsg, _
                ByRef wParam As Long, _
                ByRef lParam As Long)
    On Error Resume Next

    Dim liMsgIndex As Long
    Dim liMsgMask As Long
    Dim liEachClient As Long
    Dim liIndex As Long
    
    Dim lbCall   As Boolean

    Dim loClient    As iSubclass                            'reference the client to call the subclassproc
    
    If bBefore Then                                          'We're handling the message before the orig. wndproc
        liIndex = ArrFindInt(miMsgTableB, miMsgCountB, iMsg) 'Find the message in the table
    
        If liIndex <> Undefined Then pTranslateMsgIndex liIndex, liMsgIndex, liMsgMask 'get the msgmask index
        
        For liEachClient = 0 To miClientCount - 1&          'loop through each client
            lbCall = CBool(mtClients(liEachClient).Flags And ALL_MESSAGES_BEFORE) 'call if the client wants all messages
            If Not lbCall Then lbCall = CBool(mtClients(liEachClient).MsgMaskB(liMsgIndex) And liMsgMask) 'or if the client wants this message
            
            If lbCall Then
                Set loClient = Client(liEachClient)         'get the client reference
                If Not loClient Is Nothing Then             'if the reference is valid
                    loClient.Proc bBefore, bHandled, lReturn, hWnd, iMsg, wParam, lParam 'make the call
                Else
                    Debug.Assert False
                End If
                
                If bHandled Then Exit Sub                   'if handled by this client, bail
            End If
        Next
    Else
        liIndex = ArrFindInt(miMsgTableA, miMsgCountA, iMsg) 'Find the message in the table
    
        If liIndex <> Undefined Then pTranslateMsgIndex liIndex, liMsgIndex, liMsgMask 'get the msgmask index
        
        For liEachClient = 0 To miClientCount - 1&                                      'loop through each client
            lbCall = CBool(mtClients(liEachClient).Flags And ALL_MESSAGES_AFTER)        'call if the client wants all messages
            If Not lbCall Then lbCall = CBool(mtClients(liEachClient).MsgMaskA(liMsgIndex) And liMsgMask) 'of the client wants this message
            
            If lbCall Then
                Set loClient = Client(liEachClient)         'get the client reference
                If Not loClient Is Nothing Then             'if the reference is valid
                    loClient.Proc bBefore, bHandled, lReturn, hWnd, iMsg, wParam, lParam 'make the call
                Else
                    Debug.Assert False
                End If
                
                If bHandled Then Exit Sub                   'if handled by this client, bail
            End If
        Next
    End If
End Sub
'</iSubclass Interface>
