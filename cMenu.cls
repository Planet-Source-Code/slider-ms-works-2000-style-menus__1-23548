VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===========================================================================
'
' Class Name:   cMenu
' Author:       Slider
' Date:         29/05/2001
' Version:      01.00.00
' Description:  Holds Menu data
' Edit History: 01.00.00 29/05/01 Initial Release
'
'===========================================================================

Option Explicit

#Const DebugMode = 1

#If DebugMode Then
    Private mlClassDebugID As Long
#End If

Private oItem     As colMenuItems
Private sDesc     As String
Private bEnabled  As Boolean
Private lSelected As Long

Public Property Let Selected(ByVal vData As Long)
    lSelected = vData
End Property

Public Property Get Selected() As Long
    Selected = lSelected
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    bEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Public Property Let Desc(ByVal vData As String)
    sDesc = vData
End Property

Public Property Get Desc() As String
    Desc = sDesc
End Property

Public Property Set MenuItem(ByVal vData As colMenuItems)
    Set oItem = vData
End Property

Public Property Get MenuItem() As colMenuItems
    Set MenuItem = oItem
End Property

Private Sub Class_Initialize()
    #If DebugMode Then
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
    Set oItem = New colMenuItems
End Sub

Private Sub Class_Terminate()
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
    Set oItem = Nothing
End Sub

#If DebugMode Then
    Public Property Get ClassDebugID()
        ClassDebugID = mlClassDebugID
    End Property
#End If
