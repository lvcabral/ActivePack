VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Instances"
Attribute VB_Ext_KEY = "Member0" ,"Instances"
'local variable to hold collection
Private mCol As Collection

Public Function Add(nName As String, oType As String, Description As String, _
                    dType As String, Arguments As String, MoreInfo As String, _
                    Value As String, Example As String, CLSID As String) As clsInstances
    'create a new object
    
    On Error Resume Next
    Dim objNewMember As clsInstances
    Set objNewMember = New clsInstances
 
    Key = nName
    sKey = nName
    'set the properties passed into the method
    objNewMember.Example = Example
    objNewMember.CLSID = CLSID
    objNewMember.Value = Value
    
    objNewMember.Key = Key
    objNewMember.oType = oType
    objNewMember.MoreInfo = MoreInfo
    objNewMember.Description = Description
    objNewMember.nName = nName
    objNewMember.dType = dType
    objNewMember.Arguments = Arguments
    objNewMember.Key = Key
    objNewMember.oType = oType
    
    objNewMember.Description = Description
    objNewMember.nName = nName
    objNewMember.dType = dType
    objNewMember.Arguments = Arguments
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
 
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    On Error GoTo 0
 
End Function

Public Property Get Item(vntIndexKey As Variant) As clsInstances
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    
    'collection. Syntax: Debug.Print x.Count
    
    Count = mCol.Count
    
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    
 
 
 
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
    
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    
    Set mCol = New Collection
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    
    Set mCol = Nothing
    
End Sub

