VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dwdSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'***********************************************************************************************
'DWD Multi-Column ComboBox - Subclassing Implements Class
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'Techniques copied from vbAccelerator SSubTimer6.dll
'***********************************************************************************************
'***********************************************************************************************
'Dependencies: None
'***********************************************************************************************
'Revision History:
'[Matthew Hood]
'   08/13/01 - New
'***********************************************************************************************

'***********************************************************************************************
'Public API Enumerations
'***********************************************************************************************
Public Enum EMsgResponse
    emrConsume      ' Process instead of original WindowProc
    emrPostProcess  ' Process after original WindowProc
    emrPreprocess   ' Process before original WindowProc
End Enum
'***********************************************************************************************
'Public Member Variables
'***********************************************************************************************
Public MsgResponse As EMsgResponse
Attribute MsgResponse.VB_VarDescription = "Get/set how you will respond to a message.  Raised whenever a message is encountered.  Call the CurrentMessage property to determine what message you are currently responding to."
'***********************************************************************************************
'Public Methods
'***********************************************************************************************
Function WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
'Needed for Implements
End Function
