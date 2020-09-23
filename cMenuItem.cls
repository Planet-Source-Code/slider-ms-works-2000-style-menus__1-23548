VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===========================================================================
'
' Class Name:   cMenuItem
' Author:       Slider
' Date:         29/05/2001
' Version:      01.00.00
' Description:  Holds MenuItem data
' Edit History: 01.00.00 29/05/01 Initial Release
'
'===========================================================================

Option Explicit

#Const DebugMode = 1

#If DebugMode Then
    Private mlClassDebugID As Long
#End If

Private sDesc    As String
Private sTip     As String
Private bEnabled As Boolean
Private vTag     As Variant
Private lNdx     As Long

Friend Property Let pDisplayNDX(ByVal vData As Long)
    lNdx = vData
End Property

Friend Property Get pDisplayNDX() As Long
    pDisplayNDX = lNdx
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    bEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Public Property Let Tip(ByVal vData As String)
    sTip = vData
End Property

Public Property Get Tip() As String
    Tip = sTip
End Property

Public Property Let Desc(ByVal vData As String)
    sDesc = vData
End Property

Public Property Get Desc() As String
    Desc = sDesc
End Property

Public Property Let Tag(ByVal vData As Variant)
    vTag = vData
End Property

Public Property Get Tag() As Variant
    Tag = vTag
End Property

Private Sub Class_Initialize()
    #If DebugMode Then
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub

Private Sub Class_Terminate()
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    Public Property Get ClassDebugID()
        ClassDebugID = mlClassDebugID
    End Property
#End If
