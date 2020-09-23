VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MenuItem"
Attribute VB_Ext_KEY = "Member0" ,"MenuItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===========================================================================
'
' Class Name:   cMenu
' Author:       Slider
' Date:         29/05/2001
' Version:      01.00.00
' Description:  Collection of MenuItem data
' Edit History: 01.00.00 29/05/01 Initial Release
'
'===========================================================================

Option Explicit

Private mCol As Collection

Public Function Add(Desc As String, _
                    Tip As String, _
                    Tag As Variant, _
           Optional Enabled As Boolean = True) As cMenuItem

    Dim oItem As cMenuItem
    Set oItem = New cMenuItem

    oItem.Desc = Desc
    oItem.Tip = Tip
    If TypeName(Tag) = "Error" Then Tag = Desc
    oItem.Tag = Tag
    oItem.Enabled = Enabled
    mCol.Add oItem, Desc

    Set Add = oItem
    Set oItem = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cMenuItem
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
