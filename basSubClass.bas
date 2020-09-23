Attribute VB_Name = "basSubClass"
Option Explicit
'***********************************************************************************************
'DWD Multi-Column ComboBox - Subclassing Routines
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
'Private API Constants
'***********************************************************************************************
Private Const GWL_WNDPROC As Long = (-4)
'***********************************************************************************************
'Private API Declarations
'***********************************************************************************************
Private Declare Function CallWindowProcAPI Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemoryAPI Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetCurrentProcessIdAPI Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function GetPropAPI Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowLongAPI Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowThreadProcessIdAPI Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindowAPI Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
Private Declare Function RemovePropAPI Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetPropAPI Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLongAPI Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'***********************************************************************************************
'Public API Enumerations
'***********************************************************************************************
Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum
'***********************************************************************************************
'Private Member Variables
'***********************************************************************************************
Private mCurrentMessage As Long
Private mOldProcID As Long
'***********************************************************************************************
'Private Methods
'***********************************************************************************************
Private Sub pErrRaise(ErrNumber As Long)
On Error Resume Next
    Dim sText As String
    Dim sSource As String
    
    sSource = App.EXEName & ".WindowProc"
    If ErrNumber > 1000 Then
        Select Case ErrNumber
            Case eeCantSubclass
                sText = "Can't subclass window"
            Case eeAlreadyAttached
                sText = "Message already handled by another class"
            Case eeInvalidWindow
                sText = "Invalid window"
            Case eeNoExternalWindow
                sText = "Can't modify external window"
        End Select
        Err.Raise ErrNumber Or vbObjectError, sSource, sText
    Else
        Err.Raise ErrNumber, sSource
    End If

End Sub

Private Function pIsWindowLocal(ByVal hWnd As Long) As Boolean
On Error Resume Next
    Dim idWnd As Long

    Call GetWindowThreadProcessIdAPI(hWnd, idWnd)
    
    pIsWindowLocal = (idWnd = GetCurrentProcessIdAPI())

End Function

Private Function pWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    Dim procOld As Long
    Dim pSubclass As Long
    Dim f As Long
    Dim iwp As dwdSubClass
    Dim iwpT As dwdSubClass
    Dim iPC As Long
    Dim iP As Long
    Dim bNoProcess As Long
    Dim bCalled As Boolean

    procOld = GetPropAPI(hWnd, hWnd)
    Debug.Assert procOld <> 0
    
    bCalled = False
    iPC = GetPropAPI(hWnd, hWnd & "#" & wMsg & "C")
    If (iPC > 0) Then
        For iP = 1 To iPC
            bNoProcess = False
            pSubclass = GetPropAPI(hWnd, hWnd & "#" & wMsg & "#" & iP)
            If pSubclass = 0 Then
                pWindowProc = CallWindowProcAPI(procOld, hWnd, wMsg, wParam, ByVal lParam)
                bNoProcess = True
            End If

            If Not (bNoProcess) Then
                Call CopyMemoryAPI(iwpT, pSubclass, 4)
                Set iwp = iwpT
                Call CopyMemoryAPI(iwpT, 0&, 4)
                
                mCurrentMessage = wMsg
                mOldProcID = procOld
                
                With iwp
                    If (iP = 1) Then
                        If .MsgResponse = emrPreprocess Then
                           If Not (bCalled) Then
                              pWindowProc = CallWindowProcAPI(procOld, hWnd, wMsg, wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                    
                    pWindowProc = .WindowProc(hWnd, wMsg, wParam, ByVal lParam)
                    
                    If (iP = iPC) Then
                        If .MsgResponse = emrPostProcess Then
                           If Not (bCalled) Then
                              pWindowProc = CallWindowProcAPI(procOld, hWnd, wMsg, wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                End With
            End If
        Next iP
    Else
        pWindowProc = CallWindowProcAPI(procOld, hWnd, wMsg, wParam, ByVal lParam)
    End If

End Function
'***********************************************************************************************
'Public Methods
'***********************************************************************************************
Public Sub AttachMessage(iwp As dwdSubClass, ByVal hWnd As Long, ByVal wMsg As Long)
On Error Resume Next
    Dim procOld As Long
    Dim f As Long
    Dim c As Long
    Dim iC As Long
    Dim bFail As Boolean

    If IsWindowAPI(hWnd) = False Then Call pErrRaise(eeInvalidWindow)
    If pIsWindowLocal(hWnd) = False Then Call pErrRaise(eeNoExternalWindow)

    c = GetPropAPI(hWnd, "C" & hWnd)
    If c = 0 Then
        procOld = SetWindowLongAPI(hWnd, GWL_WNDPROC, AddressOf pWindowProc)
        If procOld = 0 Then Call pErrRaise(eeCantSubclass)
        f = SetPropAPI(hWnd, hWnd, procOld)
        Debug.Assert f <> 0
        c = 1
        f = SetPropAPI(hWnd, "C" & hWnd, c)
    Else
        c = c + 1
        f = SetPropAPI(hWnd, "C" & hWnd, c)
    End If
    Debug.Assert f <> 0
    
    c = GetPropAPI(hWnd, hWnd & "#" & wMsg & "C")
    If (c > 0) Then
        For iC = 1 To c
            If (GetPropAPI(hWnd, hWnd & "#" & wMsg & "#" & iC) = ObjPtr(iwp)) Then
                Call pErrRaise(eeAlreadyAttached)
                bFail = True
                Exit For
            End If
        Next iC
    End If
                
    If Not (bFail) Then
        c = c + 1
        f = SetPropAPI(hWnd, hWnd & "#" & wMsg & "C", c)
        Debug.Assert f <> 0
        
        f = SetPropAPI(hWnd, hWnd & "#" & wMsg & "#" & c, ObjPtr(iwp))
        Debug.Assert f <> 0
    End If

End Sub

Public Function CallOldWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

   CallOldWindowProc = CallWindowProcAPI(mOldProcID, hWnd, wMsg, wParam, lParam)

End Function

Public Function CurrentMessage() As Long
On Error Resume Next
   
   CurrentMessage = mCurrentMessage

End Function

Public Sub DetachMessage(iwp As dwdSubClass, ByVal hWnd As Long, ByVal wMsg As Long)
On Error Resume Next
    Dim procOld As Long
    Dim f As Long
    Dim c As Long
    Dim iC As Long
    Dim iP As Long
    Dim lPtr As Long

    c = GetPropAPI(hWnd, "C" & hWnd)
    If c = 1 Then
        procOld = GetPropAPI(hWnd, hWnd)
        Debug.Assert procOld <> 0
        
        Call SetWindowLongAPI(hWnd, GWL_WNDPROC, procOld)
        Call RemovePropAPI(hWnd, hWnd)
        Call RemovePropAPI(hWnd, "C" & hWnd)
    Else
        c = GetPropAPI(hWnd, "C" & hWnd)
        c = c - 1
        f = SetPropAPI(hWnd, "C" & hWnd, c)
    End If
    
    c = GetPropAPI(hWnd, hWnd & "#" & wMsg & "C")
    If (c > 0) Then
        For iC = 1 To c
            If (GetPropAPI(hWnd, hWnd & "#" & wMsg & "#" & iC) = ObjPtr(iwp)) Then
                iP = iC
                Exit For
            End If
        Next iC
    
        If (iP <> 0) Then
             For iC = iP + 1 To c
                lPtr = GetPropAPI(hWnd, hWnd & "#" & wMsg & "#" & iC)
                Call SetPropAPI(hWnd, hWnd & "#" & wMsg & "#" & (iC - 1), lPtr)
             Next iC
        End If

        Call RemovePropAPI(hWnd, hWnd & "#" & wMsg & "#" & c)
        c = c - 1
        Call SetPropAPI(hWnd, hWnd & "#" & wMsg & "C", c)
    End If

End Sub
