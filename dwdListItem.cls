VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dwdListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************************************
'DWD Multi-Column ComboBox - List Item Class
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'***********************************************************************************************
'***********************************************************************************************
'Dependencies:
'***********************************************************************************************
'Revision History:
'[Matthew Hood]
'   08/13/01 - New
'***********************************************************************************************

'***********************************************************************************************
'Private Member Variables
'***********************************************************************************************
Private mText() As String
'***********************************************************************************************
'Public Properties
'***********************************************************************************************
Public Property Get ColCount() As Long

    ColCount = UBound(mText)

End Property
Public Property Let ColCount(ByVal Value As Long)

    If (Value = 0) Then Value = 1

    ReDim Preserve mText(1 To Value) As String
    
End Property

Public Property Get Text(ByVal Index As Long) As String

    Text = mText(Index)

End Property
Public Property Let Text(ByVal Index As Long, ByVal Value As String)

    mText(Index) = Value

End Property
'***********************************************************************************************
'Class Initialization/Termination Events
'***********************************************************************************************
Private Sub Class_Initialize()
On Error Resume Next

    Me.ColCount = 1

End Sub
