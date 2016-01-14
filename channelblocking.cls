VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannelBLK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

' This extension class module supports these extensibility interfaces
Implements CHATSVCLib.IChatExtensionCallBack
Implements CHATSVCLib.IChatServerCallBack

' Hold a reference to the chat registrar to allow event binding
Dim gRegistrar As CHATSVCLib.IChatRegistrar

' Path to the file that stores our bad word list
Const WORDLIST = "C:\BLOCKCH.TXT"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      The base class ChannelBlockExtension.ChannelBlocking
''      supports interfaces to manipulate the blocked name list
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Holds the collection of blocked channel names
Private gBlockedNames As Collection

' Method to support adding an addition blocked name
Public Function Add(ByVal Name As String) As Boolean
    On Error GoTo AddFailed
    gBlockedNames.Add UCase(Name), UCase(Name)
    Add = True
    Exit Function
AddFailed:
    Add = False
End Function

' Method to remove an existing blocked name by index or name
Public Sub Remove(vntIndexKey As Variant)
    gBlockedNames.Remove vntIndexKey
End Sub

' Property to support the 'for each' syntax to list the blocked names
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = gBlockedNames.[_NewEnum]
End Property

' The default property to allow indexing into the collection of names
Public Property Get Item(vntIndexKey As Variant) As String
Attribute Item.VB_UserMemId = 0
    Item = gBlockedNames(vntIndexKey)
End Property

' Property that returns the count of blocked names
Public Property Get Count() As Long
    Count = gBlockedNames.Count
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      Private Module Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IsNameBlocked(ByVal Name As String) As Boolean
    ' trap the error if name isn't in collection
    On Error GoTo notfound
    IsNameBlocked = (gBlockedNames(Name) <> "")
    On Error GoTo 0
    Exit Function
    
notfound:
    IsNameBlocked = False
End Function

Private Sub LoadWordList()
    Dim NewName As String
    Dim fn As Integer
    
    On Error GoTo CloseFile
    fn = FreeFile
    Open WORDLIST For Input As #fn
    
GetNextLine:
    Input #fn, NewName
    Add NewName
    GoTo GetNextLine
    
CloseFile:
    On Error GoTo 0
    Close #fn
End Sub

Private Sub SaveWordList()
    Dim fn As Integer
    Dim i As Integer
    
    fn = FreeFile
    Open WORDLIST For Output As #fn
    For i = 1 To gBlockedNames.Count
        Print #fn, gBlockedNames(i)
    Next i
    Close #fn
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      Implement necessary extensibility interface methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' All extension should define a short-name
Private Property Get IChatExtensionCallBack_Name() As String
    IChatExtensionCallBack_Name = "ChannelBlocking"
End Property

' This method is invoked whenever this extension is started
Private Function IChatExtensionCallBack_Init(ByVal Server As CHATSVCLib.IChatServer, ByVal Registrar As CHATSVCLib.IChatRegistrar) As Long
    Set gBlockedNames = New Collection
    Set gRegistrar = Registrar
    gRegistrar.AddServerEvent "OnNewChannel"
    
    ' Load the default blocked channel names
    Call LoadWordList
    
    IChatExtensionCallBack_Init = 0
End Function

' This method is called just before the extension is stopped
Private Function IChatExtensionCallBack_Term() As Long
    Call SaveWordList
    Set gBlockedNames = Nothing
End Function


Private Function IChatExtensionCallBack_OnGetPropertyPageClass(ByVal Class As CHATSVCLib.PROPERTY_PAGE) As String
    If Class = PPAGE_SERVER Then
        IChatExtensionCallBack_OnGetPropertyPageClass = "ChannelBlockPP.ServerPages"
    End If
End Function

Private Function IChatServerCallBack_OnNewChannel(ByVal Channel As CHATSVCLib.IChatChannel, ByVal PostUpdate As Boolean) As Long
    IChatServerCallBack_OnNewChannel = 0
    
    If Not PostUpdate Then
        If IsNameBlocked(Channel.Name) Then
            ' return IRC's ERR_NOSUCHCHANNEL error message
            IChatServerCallBack_OnNewChannel = 403
        End If
    End If
End Function


Private Sub IChatExtensionCallBack_OnInstall()
    Dim fn As Integer
    
    ' When installed, create a persistent file of blocked channel names
    fn = FreeFile
    Open WORDLIST For Output As #fn
    Print #fn, "#reserved"
    Print #fn, "#blocked"
    Close #fn
End Sub

Private Sub IChatExtensionCallBack_OnUninstall()
    Kill WORDLIST
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Declare, but don't implement the unused extensibility interface methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IChatServerCallBack_OnAddAccess(ByVal Server As CHATSVCLib.IChatServer, ByVal AccessEntry As CHATSVCLib.IChatAccessEntry, ByVal PostUpdate As Boolean) As Long
End Function

Private Sub IChatServerCallBack_OnCloseChannel(ByVal Channel As CHATSVCLib.IChatChannel)
End Sub

Private Sub IChatServerCallBack_OnCloseUser(ByVal User As CHATSVCLib.IChatUser)
End Sub

Private Function IChatServerCallBack_OnNewUser(ByVal User As CHATSVCLib.IChatUser, ByVal PostUpdate As Boolean) As Long
End Function

Private Sub IChatExtensionCallBack_OnConfigureExtension(ByVal User As CHATSVCLib.IChatUser, ByVal Command As String)
End Sub

Private Sub IChatServerCallBack_OnRemoveAccess(ByVal Server As CHATSVCLib.IChatServer, ByVal AccessEntry As CHATSVCLib.IChatAccessEntry)
End Sub

