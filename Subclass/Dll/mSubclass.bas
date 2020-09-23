Attribute VB_Name = "mSubclass"
Option Explicit

'==================================================================================================
'mSubclass.bas                          7/5/04
'
'           PURPOSE:
'               Uses a separate cSubclassHub object for each window that is subclassed.  The
'               cSubclassHub object is responsible for delivering the messages to the requesting objects.
'
'           CLASSES CREATED BY THIS MODULE:
'               pcSubclassHub
'
'               cSubclass
'
'==================================================================================================

'1.  Private Interface      - Utility procedures
'2.  cSubclass Interface    - Procedures called by cSubclass
'3.  cSubclasses Interface  - Procedures called by cSubclasses

#Const bVBVMTypeLib = False      'Toggles between using the MSVBVM type library

Private Type tSubclassClient    'store one record for each object requesting subclasses
    hWnd() As Long              'store the hWnds being subclassed
    hWndCount As Long           'store the count of the hWnds
    Ptr As Long                 'store the pointer to the object
    iControl As Long            'keep track of changes for enumeration
End Type

Private mtClients() As tSubclassClient  'one record for each object requesting subclasses
Private miClientCount As Long           'current record count

Public mCollSubclasses As Collection    'collection of pcSubclass objects to relay the messages

'<Private Interface>
Public Property Get MsgHubObject( _
                ByVal hWnd As Long, _
       Optional ByVal bForce As Boolean) _
                    As pcSubclassHub
    
    On Error GoTo NotThere
    Set MsgHubObject = mCollSubclasses("h" & hWnd)  'return the collection item for this hWnd
    Exit Property
NotThere:
    If bForce Then
        If mCollSubclasses Is Nothing Then Set mCollSubclasses = New Collection
        Set MsgHubObject = New pcSubclassHub
        mCollSubclasses.Add MsgHubObject, "h" & hWnd
        If Not MsgHubObject.Subclass(hWnd) Then gErr vbbApiError, "cSubclasses.Add"
    Else
        gErr vbbItemDetached, "cSubclass"
    End If
End Property

Private Sub RemoveHub( _
            ByVal hWnd As Long)
    On Error Resume Next
    mCollSubclasses.Remove "h" & hWnd        'Remove the object from the collection
End Sub

Private Function FindClient( _
            ByVal iPtr As Long, _
   Optional ByRef iFirstAvailable As Long) _
                As Long
    Dim liTemp As Long
    iFirstAvailable = Undefined                         'Initialize the first available slot to nothing
    For FindClient = 0& To miClientCount - 1&           'loop through each client
        liTemp = mtClients(FindClient).Ptr              'store the client's pointer
        If Not (liTemp = 0& Or liTemp = Undefined) Then 'if this pointer is valid
            If iPtr = liTemp Then Exit Function         'if the pointer matches then bail
        Else                                            'if the pointer is invalid
            If iFirstAvailable = Undefined Then _
                iFirstAvailable = FindClient            'it may be the first available slot
        End If
    Next
    FindClient = Undefined                              'if we made it out here, then the client was not found
End Function

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
'<cSubclasses Interface>
Public Function Subclasses_Add( _
            ByVal iWho As Long, _
            ByVal hWnd As Long _
    ) As cSubclass
    
    If Not MsgHubObject(hWnd, True).AddClient(iWho) Then gErr vbbKeyAlreadyExists, "cSubclasses.Add"

    Dim liIndex As Long
    Dim liFirst As Long
    liIndex = FindClient(iWho, liFirst) 'get the index of the client if it exists, and get the first available slot

    If liIndex = Undefined Then         'if the client was not already there then
        If liFirst = Undefined Then     'if there was an open slot then
            liFirst = miClientCount     'next index is current count
            miClientCount = miClientCount + 1&  'inc the count
            ArrRedimT mtClients, miClientCount, True    'resize the array
        End If
        With mtClients(liFirst)         'with this array index
            .Ptr = iWho                 'set the pointer to this object
            .hWndCount = 0              'initialize the hWnd count
        End With
        liIndex = liFirst               'store this index
    End If
    Incr mtClients(liIndex).iControl
                                        'add the hWnd to the table
    ArrAddInt mtClients(liIndex).hWnd, mtClients(liIndex).hWndCount, hWnd
    
    Set Subclasses_Add = New cSubclass
    Subclasses_Add.fInit iWho, hWnd

End Function

Public Sub Subclasses_Remove( _
                    ByVal iWho As Long, _
                    ByVal hWnd As Long)

    On Error GoTo NotThere
        
    Dim loHub As pcSubclassHub
    Set loHub = MsgHubObject(hWnd)
    If Not loHub.DelClient(iWho) Then
NotThere:
        gErr vbbKeyNotFound, "cSubclasses.Remove"
    End If
    

    If Not loHub.Active Then            'If there's nobody left to notify then destroy the object
        
        Dim liIndex As Long
        liIndex = FindClient(iWho)      'get the index of the client in our array
        If liIndex <> Undefined Then    'if the index was found
            With mtClients(liIndex)     'remove the hWnd from the table
                ArrDelInt .hWnd, .hWndCount, hWnd
                If .hWndCount = 0& Then 'if that was the last hWnd
                    .Ptr = 0&           'remove the client from the table
                End If
                Incr .iControl
            End With
        Else
            'This would be bad!
            'client wasn't found!!
            Debug.Assert False
        End If

        Set loHub = Nothing             'destroy our reference to the object
        RemoveHub hWnd                  'remove the object from the collection
    End If

End Sub

Public Function Subclasses_Item( _
            ByVal hWnd As Long, _
            ByVal iWho As Long) _
                As cSubclass
    On Error GoTo NoHub
    If MsgHubObject(hWnd).ClientExists(iWho) Then
        Set Subclasses_Item = New cSubclass
        Subclasses_Item.fInit iWho, hWnd
    Else
NoHub:
        gErr vbbKeyNotFound, "cSubclasses.Item"
    End If

End Function

Public Function Subclasses_Exists( _
            ByVal hWnd As Long, _
            ByVal iWho As Long _
        ) As Boolean
        
    On Error GoTo NoHub
    Subclasses_Exists = MsgHubObject(hWnd).ClientExists(iWho)
    Exit Function
NoHub:
    Err.Clear
End Function

Public Function Subclasses_Count( _
            ByVal iWho As Long) _
                As Long
        
    Dim liIndex As Long
    liIndex = FindClient(iWho)      'find the index of the client
                                    'if the index was found, return the hWnd count
                                    
    If liIndex > Undefined Then
        With mtClients(liIndex)
            Subclasses_Count = .hWndCount
            For liIndex = 0& To .hWndCount - 1&
                If .hWnd(liIndex) = 0& Or .hWnd(liIndex) = Undefined _
                Then Subclasses_Count = Subclasses_Count - 1&
            Next
        End With
    End If
    
End Function

Public Function Subclasses_Clear( _
            ByVal iWho As Long) _
                As Long
    
    Dim liIndex As Long
    Dim i       As Long
    Dim loHub   As pcSubclassHub
    
    liIndex = FindClient(iWho)                          'Find the index of the client
    
    If liIndex <> Undefined Then                        'If the index was found
        With mtClients(liIndex)
            For i = 0 To .hWndCount - 1&                'loop through each hWnd
                If .hWnd(i) <> Undefined Then
                    Set loHub = MsgHubObject(.hWnd(i))  'Retrieve the subclasser associated with this hwnd
                    If Not loHub Is Nothing Then
                                                        'Tell the object to cease notifications for this client
                        If loHub.DelClient(iWho) Then _
                            Subclasses_Clear = Subclasses_Clear + 1&
                        
                        If Not loHub.Active Then        'If there's nobody left to notify then destroy the object
                            Set loHub = Nothing
                            RemoveHub .hWnd(i)
                        End If
                    Else
                        'hWnd is not subclassed???
                        Debug.Assert False 'this would be bad!
                    End If
                End If
            Next
            .hWndCount = 0&                             'remove the hWnd
            .Ptr = 0&                                   'remove the ptr
            Incr .iControl
        End With
    End If
    
End Function

'Public Sub Subclasses_NextItem( _
'            ByVal iWho As Long, _
'            ByRef tEnum As tEnum, _
'            ByRef vNextItem As Variant, _
'            ByRef bNoMore As Boolean)
'
'    Dim loSub As cSubclass
'    Dim liIndex As Long
'    Dim liClient As Long
'    Dim i As Long
'
'    liClient = FindClient(iWho)
'    If liClient <> Undefined Then
'
'        liIndex = tEnum.iIndex
'        liIndex = liIndex + 1&
'
'        With mtClients(liClient)
'
'            If .iControl <> tEnum.iControl Then gErr vbbCollChangedDuringEnum, "cSubclasses.NewEnum"
'
'            For i = liIndex To .hWndCount - 1&
'                If .hWnd(i) <> Undefined Then       'if the hWnd exists
'                    Set loSub = New cSubclass       'create a new subclass object
'                    loSub.fInit iWho, .hWnd(i)      'initialize it
'                    Set vNextItem = loSub
'                    Exit For
'                End If
'            Next
'            If i = .hWndCount Then bNoMore = True
'        End With
'
'        tEnum.iIndex = i
'    Else
'        bNoMore = True
'    End If
'End Sub
'
'Public Function Subclasses_Skip( _
'            ByVal iWho As Long, _
'            ByRef tEnum As tEnum, _
'            ByVal iSkipCount As Long, _
'            ByRef bSkippedAll As Boolean)
'
'    Dim liSkipped As Long
'    Dim liClient As Long
'
'    liClient = FindClient(iWho)
'    If liClient <> Undefined Then
'
'
'        With mtClients(liClient)
'
'            If .iControl <> tEnum.iControl Then gErr vbbCollChangedDuringEnum, "cSubclasses.NewEnum"
'
'            For tEnum.iIndex = tEnum.iIndex + 1& To .hWndCount - 1&
'                If .hWnd(tEnum.iIndex) <> Undefined Then liSkipped = liSkipped + 1&
'                If liSkipped = iSkipCount Then Exit For
'            Next
'            bSkippedAll = CBool(liSkipped = iSkipCount)
'
'        End With
'    Else
'        bSkippedAll = False
'
'    End If
'
'End Function
'
'Public Function Subclasses_Control(ByVal iWho As Long) As Long
'    Subclasses_Control = FindClient(iWho)
'    If Subclasses_Control <> Undefined Then
'        Subclasses_Control = mtClients(Subclasses_Control).iControl
'    Else
'        'client not there!
'        'Debug.Assert False
'    End If
'End Function
'</cSubclasses Interface>
'</Public Interface>

