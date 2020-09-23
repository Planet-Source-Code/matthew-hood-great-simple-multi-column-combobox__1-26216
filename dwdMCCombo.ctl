VERSION 5.00
Begin VB.UserControl dwdMCCombo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ScaleHeight     =   705
   ScaleWidth      =   2010
   ToolboxBitmap   =   "dwdMCCombo.ctx":0000
   Begin VB.ListBox lstList 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "dwdMCCombo.ctx":0312
      Left            =   60
      List            =   "dwdMCCombo.ctx":0314
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtText 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
End
Attribute VB_Name = "dwdMCCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements dwdSubClass
'***********************************************************************************************
'DWD Multi-Column ComboBox - ActiveX Control
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'***********************************************************************************************
'***********************************************************************************************
'Dependencies:
'***********************************************************************************************
'Revision History:
'[Matthew Hood]
'   08/13/01 - New
'   02/07/02 - Bug fix.
'              The application will now remain activated when the user shows the drop-down.
'              Added support for MDI & MDI Child forms.
'   02/12/02 - Bug fix.
'              Moved "Me.ButtonDown = False" to the lstList_Click event from the lstList_MouseDown
'              event to fix a clicking bug.  Removed the lstList_MouseDown event.
'   04/07/02 - Bug fix.
'              Fixed scrollbar display for WindowsXP.
'***********************************************************************************************
'***********************************************************************************************
'Private API Types
'***********************************************************************************************
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type
'***********************************************************************************************
'Private API Constants
'***********************************************************************************************
Private Const SM_CXHTHUMB As Long = 10

Private Const DFC_SCROLL As Long = 3
Private Const DFCS_SCROLLDOWN As Long = &H1
Private Const DFCS_PUSHED As Long = &H200
Private Const DFCS_FLAT As Long = &H4000
Private Const DFCS_INACTIVE As Long = &H100

Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)

Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5

Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_HIDEWINDOW As Long = &H80

Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_BORDER As Long = &H800000
Private Const WS_CHILD As Long = &H40000000

Private Const SPI_GETWORKAREA As Long = 48
Private Const HWND_TOPMOST As Long = -1

Private Const LB_FINDSTRING As Long = &H18F
Private Const LB_SETTABSTOPS As Long = &H192

Private Const WM_ACTIVATE As Long = &H6
Private Const WM_NCACTIVATE As Long = &H86

Private Const GA_ROOT As Long = 2
'***********************************************************************************************
'Private API Declarations
'***********************************************************************************************
Private Declare Function DrawFrameControlAPI Lib "user32" Alias "DrawFrameControl" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetDesktopWindowAPI Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function GetAncestorAPI Lib "user32.dll" Alias "GetAncestor" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRectAPI Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindowAPI Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessageAPI Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParentAPI Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLongAPI Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPosAPI Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindowAPI Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SystemParametersInfoAPI Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function LBItemFromPtAPI Lib "comctl32.dll" Alias "LBItemFromPt" (ByVal hWnd As Long, ByVal ptx As Long, ByVal pty As Long, ByVal bAutoScroll As Long) As Long
Private Declare Function ClientToScreenAPI Lib "user32" Alias "ClientToScreen" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLongAPI Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowThemeAPI Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, pszSubAppName As String, pszSubIdList As String) As Long 'Rev. (04/07/02)
'***********************************************************************************************
'Public Enumerations
'***********************************************************************************************
Public Enum dwdComboBoxStyle
    vbComboDropdown
    vbComboDropdownList
End Enum
'***********************************************************************************************
'Public Events
'***********************************************************************************************
Public Event ButtonClick(ByVal ButtonDown As Boolean)
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***********************************************************************************************
'Private Member Variables
'***********************************************************************************************
Private mButtonDown As Boolean
Private mColumnCount As Long
Private mColumnWidth() As Long
Private mDropDownHeight As Long
Private mDropDownWidth As Long
Private mDropDownWidthAutoSize As Boolean
Private mHotTracking As Boolean
Private mList As Collection
Private mListIndex As Long
Private mLocked As Boolean
Private mNoClick As Boolean
Private mOldValue As String
Private mStyle As dwdComboBoxStyle
'***********************************************************************************************
'Public Properties
'***********************************************************************************************
Public ParentHwnd As Long
Public Property Get Alignment() As AlignmentConstants
On Error Resume Next

    Alignment = txtText.Alignment
    
End Property
Public Property Let Alignment(ByVal Value As AlignmentConstants)

    txtText.Alignment = Value
    
    Call PropertyChanged("Alignment")

End Property

Public Property Get BackColor() As OLE_COLOR
On Error Resume Next

    BackColor = txtText.BackColor
    
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)

    txtText.BackColor = Value
    lstList.BackColor = Value

    Call PropertyChanged("BackColor")

End Property

Public Property Get ButtonDown() As Boolean
On Error Resume Next

    ButtonDown = mButtonDown
    
End Property
Public Property Let ButtonDown(ByVal Value As Boolean)

    If Value = mButtonDown Then Exit Property

    mButtonDown = Value

    Call pDrawButton
    
    Call pShowHideDropDown

End Property

Public Property Get ColumnCount() As Long
On Error Resume Next

    ColumnCount = mColumnCount
    
End Property
Public Property Let ColumnCount(ByVal Value As Long)
    Dim l As Long
    
    If Value < 1 Then
        Err.Raise 380
        Exit Property
    End If

    mColumnCount = Value
    ReDim Preserve mColumnWidth(1 To Value) As Long

    For l = 1 To mList.Count
        mList(l).ColCount = Value
    Next l

    Call PropertyChanged("ColumnCount")

End Property

Public Property Get ColumnWidth(ByVal Index As Long) As Long
On Error Resume Next

    ColumnWidth = mColumnWidth(Index)
    
End Property
Public Property Let ColumnWidth(ByVal Index As Long, ByVal Value As Long)

    mColumnWidth(Index) = Value

End Property

Public Property Get DropDownHeight() As Long
On Error Resume Next

    DropDownHeight = mDropDownHeight

End Property
Public Property Let DropDownHeight(ByVal Value As Long)

    mDropDownHeight = Value
    
    Call PropertyChanged("DropDownHeight")

End Property

Public Property Get DropDownWidth() As Long
On Error Resume Next

    DropDownWidth = mDropDownWidth
    
End Property
Public Property Let DropDownWidth(ByVal Value As Long)

    If (Value < -1) Then
        Err.Raise 380
        Exit Property
    End If

    mDropDownWidthAutoSize = (Value = UserControl.Width)

    mDropDownWidth = Value

    Call PropertyChanged("DrowDownWidth")
    
End Property

Public Property Get Enabled() As Boolean
On Error Resume Next
    
    Enabled = UserControl.Enabled

End Property
Public Property Let Enabled(ByVal Value As Boolean)

    UserControl.Enabled = Value
    txtText.Enabled = Value
    
    If (Not Value) Then
        Me.ButtonDown = False
    Else
        Call pDrawButton
    End If

    Call PropertyChanged("Enabled")

End Property

Public Property Get Font() As StdFont
On Error Resume Next

    Set Font = txtText.Font
    
End Property
Public Property Set Font(ByVal Value As StdFont)

    Set txtText.Font = Value
    Set lstList.Font = Value

    Call PropertyChanged("Font")

End Property

Public Property Get FontBold() As Boolean
On Error Resume Next

    FontBold = txtText.FontBold
    
End Property
Public Property Let FontBold(ByVal Value As Boolean)

    txtText.FontBold = Value
    lstList.FontBold = Value
    
    Call PropertyChanged("FontBold")

End Property

Public Property Get FontItalic() As Boolean
On Error Resume Next

    FontItalic = txtText.FontItalic
    
End Property
Public Property Let FontItalic(ByVal Value As Boolean)

    txtText.FontItalic = Value
    lstList.FontItalic = Value
    
    Call PropertyChanged("FontItalic")

End Property

Public Property Get FontName() As String
On Error Resume Next

    FontName = txtText.FontName
    
End Property
Public Property Let FontName(ByVal Value As String)

    txtText.FontName = Value
    lstList.FontName = Value
    
    Call PropertyChanged("FontName")

End Property

Public Property Get FontSize() As Single
On Error Resume Next

    FontSize = txtText.FontSize
    
End Property
Public Property Let FontSize(ByVal Value As Single)

    txtText.FontSize = Value
    lstList.FontSize = Value
    
    Call PropertyChanged("FontSize")

End Property

Public Property Get FontStrikethru() As Boolean
On Error Resume Next

    FontStrikethru = txtText.FontStrikethru
    
End Property
Public Property Let FontStrikethru(ByVal Value As Boolean)

    txtText.FontStrikethru = Value
    
    Call PropertyChanged("FontSFontStrikethruize")

End Property

Public Property Get FontUnderline() As Boolean
On Error Resume Next

    FontUnderline = txtText.FontUnderline
    
End Property
Public Property Let FontUnderline(ByVal Value As Boolean)

    txtText.FontUnderline = Value
    
    Call PropertyChanged("FontUnderline")

End Property

Public Property Get ForeColor() As OLE_COLOR
On Error Resume Next

    ForeColor = txtText.ForeColor
    
End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)

    txtText.ForeColor = Value
    
    Call PropertyChanged("ForeColor")
    
End Property

Public Property Get Height() As Long
On Error Resume Next
    
    Height = UserControl.Height

End Property

Public Property Get hWnd() As Long
On Error Resume Next

    hWnd = UserControl.hWnd
    
End Property

Public Property Get hWndEdit() As Long
On Error Resume Next

    hWndEdit = txtText.hWnd
    
End Property

Public Property Get hWndList() As Long
On Error Resume Next

    hWndList = lstList.hWnd

End Property

Public Property Get ItemText(ByVal ListIndex As Long, ByVal ColIndex As Long) As String
On Error Resume Next

    ItemText = mList(ListIndex).Text(ColIndex)

End Property
Public Property Let ItemText(ByVal ListIndex As Long, ByVal ColIndex As Long, ByVal Value As String)
    Dim l As Long
    Dim sText As String
    
    mList(ListIndex).Text(ColIndex) = Value
    
    sText = mList(ListIndex).Text(1)
    For l = 2 To mColumnCount
        sText = sText & vbTab & mList(ListIndex).Text(l)
    Next l
    lstList.List(ListIndex - 1) = sText

    If (ListIndex = 1) And (ColIndex = 1) Then Me.Text = Value

End Property

Public Property Get ListCount() As Long
On Error Resume Next

    ListCount = mList.Count
    
End Property

Public Property Get ListIndex() As Long
On Error Resume Next

    ListIndex = mListIndex
    
End Property
Public Property Let ListIndex(ByVal Value As Long)

    mListIndex = Value
    If (mListIndex = 0) Then
        Me.Text = vbNullString
    Else
        Me.Text = mList(Value).Text(1)
    End If

End Property

Public Property Get Locked() As Boolean
On Error Resume Next

    Locked = mLocked

End Property
Public Property Let Locked(ByVal Value As Boolean)

    mLocked = Value
    If (mStyle = vbComboDropdown) Then txtText.Locked = Value

    Call PropertyChanged("Locked")

End Property

Public Property Get MaxLength() As Long
On Error Resume Next

    MaxLength = txtText.MaxLength
    
End Property
Public Property Let MaxLength(ByVal Value As Long)

    txtText.MaxLength = Value
    
    Call PropertyChanged("MaxLength")
    
End Property

Public Property Get SelLength() As Long
On Error Resume Next

    SelLength = txtText.SelLength
    
End Property
Public Property Let SelLength(ByVal Value As Long)

    txtText.SelLength = Value
    
End Property

Public Property Get SelStart() As Long
On Error Resume Next

    SelStart = txtText.SelStart
    
End Property
Public Property Let SelStart(ByVal Value As Long)

    txtText.SelStart = Value
    
End Property

Public Property Get SelText() As String
On Error Resume Next

    SelText = txtText.SelText
    
End Property
Public Property Let SelText(ByVal Value As String)

    txtText.SelText = Value

End Property

Public Property Get Style() As dwdComboBoxStyle
On Error Resume Next
    
    Style = mStyle
    
End Property
Public Property Let Style(ByVal Value As dwdComboBoxStyle)

    mStyle = Value
    If Value = vbComboDropdownList Then
        txtText.Locked = True
    Else
        txtText.Locked = mLocked
    End If

    Call PropertyChanged("Style")

End Property

Public Property Get Text() As String
On Error Resume Next

    Text = txtText.Text

End Property
Public Property Let Text(ByVal Value As String)

    txtText.Text = Value
    
    Call PropertyChanged("Text")

End Property

Public Property Get ToolTip() As String
On Error Resume Next

    ToolTip = txtText.ToolTipText
    
End Property
Public Property Let ToolTip(ByVal Value As String)
On Error Resume Next

    txtText.ToolTipText = Value
    UserControl.Extender.ToolTipText = Value

End Property
'***********************************************************************************************
'Private Methods
'***********************************************************************************************
'Draws the drop-down button.
Private Sub pDrawButton()

    If UserControl.Enabled Then
        If mButtonDown Then
            Call DrawFrameControlAPI(ByVal UserControl.hdc, pRect((txtText.Width / Screen.TwipsPerPixelX), 0, GetSystemMetricsAPI(ByVal SM_CXHTHUMB), ((txtText.Height + 30) / Screen.TwipsPerPixelY) - 4), DFC_SCROLL, DFCS_SCROLLDOWN Or DFCS_PUSHED Or DFCS_FLAT)
        Else
            Call DrawFrameControlAPI(ByVal UserControl.hdc, pRect((txtText.Width / Screen.TwipsPerPixelX), 0, GetSystemMetricsAPI(ByVal SM_CXHTHUMB), ((txtText.Height + 30) / Screen.TwipsPerPixelY) - 4), DFC_SCROLL, DFCS_SCROLLDOWN)
        End If
    Else
        Call DrawFrameControlAPI(ByVal UserControl.hdc, pRect((txtText.Width / Screen.TwipsPerPixelX), 0, GetSystemMetricsAPI(ByVal SM_CXHTHUMB), ((txtText.Height + 30) / Screen.TwipsPerPixelY) - 4), DFC_SCROLL, DFCS_SCROLLDOWN Or DFCS_INACTIVE)
    End If

End Sub

'Returns the Hi & Lo values of the lParam.
Private Sub pGetHiLoWord(ByVal lParam As Long, ByRef LoWord As Long, ByRef HiWord As Long)
On Error Resume Next

    LoWord = lParam And &HFFFF&
    HiWord = lParam \ &H10000 And &HFFFF&

End Sub

Private Function pGetListRECT() As RECT
On Error Resume Next
    Dim rct As RECT
    Dim rctUC As RECT
    Dim rctSCR As RECT
    Dim lSCRH As Long
    Dim lHeight As Long
    Dim lBottom As Long
    Dim lRight As Long
    Dim lWidth As Long
    Dim lCnt As Long

    Call GetWindowRectAPI(ByVal UserControl.hWnd, rctUC)
    Call SystemParametersInfoAPI(ByVal SPI_GETWORKAREA, ByVal 0, rctSCR, ByVal 0)

    lSCRH = rctSCR.Bottom - rctSCR.Top
    
    lCnt = mList.Count
    If (mDropDownHeight = 0) Then
        If (lCnt > 8) Then
            lHeight = (8 * 195)
        ElseIf (lCnt = 0) Then
            lHeight = 230
        Else
            lHeight = 30 + (lCnt * 195)
        End If
    ElseIf (lCnt = 0) Then
        lHeight = 230
    Else
        lHeight = mDropDownHeight
    End If

    rct.Bottom = (lHeight / Screen.TwipsPerPixelY)
    rct.Top = rctUC.Bottom + 1

    If (mDropDownWidth = -1) Then
        For lCnt = 1 To mColumnCount
            If mColumnWidth(lCnt) = 0 Then
                lWidth = lWidth + 1440
            Else
                lWidth = lWidth + mColumnWidth(lCnt)
            End If
        Next lCnt
    Else
        lWidth = mDropDownWidth
    End If
    If lWidth < UserControl.Width Then lWidth = UserControl.Width

    rct.Right = (lWidth / Screen.TwipsPerPixelX)
    rct.Left = rctUC.Left
    lRight = rct.Left + rct.Right

    pGetListRECT = rct

End Function

'Returns a rectangle type from specified parameters.
Private Function pRect(Left As Long, Top As Long, Width As Long, Height As Long) As RECT
On Error Resume Next
    
    With pRect
      .Left = Left
      .Top = Top
      .Right = Left + Width
      .Bottom = Top + Height
   End With
   
End Function

Private Sub pShowHideDropDown()
On Error Resume Next
    Const CONST_DLUPERTWIP As Single = 21.9
    Dim rct As RECT
    Dim lPhWnd As Long
    Dim l As Long
    Dim mCols() As Long

    If mButtonDown Then
        ReDim mCols(mColumnCount - 1) As Long
        For l = 0 To mColumnCount - 1
            If (l <> 0) Then
                mCols(l) = mCols(l - 1) + (mColumnWidth(l + 1) / CONST_DLUPERTWIP)
            Else
                mCols(l) = (mColumnWidth(l + 1) / CONST_DLUPERTWIP)
            End If
        Next l

        Call SendMessageAPI(lstList.hWnd, LB_SETTABSTOPS, 0&, ByVal 0&)
        Call SendMessageAPI(lstList.hWnd, LB_SETTABSTOPS, 3, mCols(0))
        lstList.Refresh

        mNoClick = True
        lstList.ListIndex = lstList.ListCount - 1
        lstList.ListIndex = 0
        lstList.ListIndex = mListIndex - 1
        mNoClick = False
        mOldValue = Me.Text
        
        rct = pGetListRECT
        Call SetParentAPI(lstList.hWnd, GetDesktopWindowAPI)
        Call MoveWindowAPI(lstList.hWnd, rct.Left, rct.Top, rct.Right, rct.Bottom, -1)
        Call SetWindowThemeAPI(lstList.hWnd, 0, 0) 'Rev. (04/07/02)
        Call SetWindowLongAPI(lstList.hWnd, GWL_STYLE, WS_BORDER)
        Call SetWindowLongAPI(lstList.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)
        Call SetWindowPosAPI(lstList.hWnd, HWND_TOPMOST, rct.Left, rct.Top, rct.Right, rct.Bottom, (SWP_HIDEWINDOW Or SWP_NOSIZE)) 'Rev. (02/07/02)
        Call SetWindowPosAPI(lstList.hWnd, HWND_TOPMOST, rct.Left, rct.Top, rct.Right, rct.Bottom, (SWP_SHOWWINDOW Or SWP_NOSIZE))
        
        'Rev. (02/07/02)
        If UserControl.Parent.MDIChild Then
            lPhWnd = GetAncestorAPI(UserControl.Parent.hWnd, GA_ROOT)
        Else
            lPhWnd = UserControl.Parent.hWnd
        End If
        Call SendMessageAPI(lPhWnd, WM_NCACTIVATE, 1, 0)
        'End Rev. (02/07/02)

        Call AttachMessage(Me, lstList.hWnd, WM_ACTIVATE)
        DoEvents
        lstList.SetFocus
    Else
        Call SetParentAPI(lstList.hWnd, UserControl.hWnd)
        Call ShowWindowAPI(lstList.hWnd, SW_HIDE)
        Call DetachMessage(Me, lstList.hWnd, WM_ACTIVATE)
        Call SendMessageAPI(UserControl.Parent.hWnd, WM_ACTIVATE, 1, 0)
    End If
    RaiseEvent ButtonClick(mButtonDown)

End Sub
'***********************************************************************************************
'Public Methods
'***********************************************************************************************
Public Function AddItem(ByVal Value As String) As Long
On Error Resume Next
    Dim itm As dwdListItem

    Set itm = New dwdListItem
    itm.ColCount = mColumnCount
    itm.Text(1) = Value

    mList.Add itm
    Set itm = Nothing

    lstList.AddItem Value
    AddItem = mList.Count

End Function

Public Sub Clear()
On Error Resume Next

    Set mList = New Collection
    lstList.Clear

    mListIndex = 0

End Sub

Public Function FindItem(ByVal Value As String, Optional ByVal ColumnIndex As Long = 1, Optional ByVal StartIndex As Long, Optional ByVal Exact As Boolean) As Long
On Error Resume Next
    Dim l As Long
    Dim sFText As String
    Dim lLen As Long

    sFText = Value
    lLen = Len(sFText)

    If (sFText = vbNullString) Then
        FindItem = 0
        Exit Function
    End If

    If (StartIndex = 0) And (mList.Count > 0) Then StartIndex = 1
    If (ColumnIndex < 1) Or (ColumnIndex > mColumnCount) Then ColumnIndex = 1

    For l = StartIndex To mList.Count
        If Exact Then
            If StrComp(mList(l).Text(ColumnIndex), sFText, vbTextCompare) = 0 Then
                FindItem = l
                Exit For
            Else
                FindItem = 0
            End If
        Else
            If (StrComp(Left$(mList(l).Text(ColumnIndex), lLen), sFText, vbTextCompare) = 0) Then
                FindItem = l
                Exit For
            Else
                FindItem = 0
            End If
        End If
    Next l

End Function

Public Function RemoveItem(ByVal Index As Long)
On Error Resume Next

    mList.Remove Index
    mListIndex = mListIndex - 1
    lstList.RemoveItem Index - 1

End Function
'***********************************************************************************************
'Change/Validation Events
'***********************************************************************************************
Private Sub txtText_Change()
On Error Resume Next
    
    mListIndex = Me.FindItem(txtText.Text, , , True)

    RaiseEvent Change
    
End Sub
'***********************************************************************************************
'Mouse Events
'***********************************************************************************************
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If (X > txtText.Width) Then
        Me.ButtonDown = Not Me.ButtonDown
    Else
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub txtText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub lstList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim lIndex As Long
    Dim pt As POINTAPI
      
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    Call ClientToScreenAPI(lstList.hWnd, pt)

    lIndex = LBItemFromPtAPI(lstList.hWnd, pt.X, pt.Y, False)

    If (lIndex > -1) Then
        mHotTracking = True
        lstList.Selected(lIndex) = True
        mHotTracking = False
    End If

End Sub
'***********************************************************************************************
'Keyboard Events
'***********************************************************************************************
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error Resume Next

    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    RaiseEvent KeyDown(KeyCode, Shift)

    Select Case KeyCode
        Case vbKeyF4
            If (Shift = 0) Then
                KeyCode = 0
                Me.ButtonDown = Not Me.ButtonDown
            End If
        Case vbKeyUp
            If (Shift = 0) Then
                KeyCode = 0
                If (Me.ListIndex > 0) Then
                    Me.ListIndex = Me.ListIndex - 1
                End If
            End If
        Case vbKeyDown
            If (Shift = 0) Then
                KeyCode = 0
                If (Me.ListIndex = -1) Then
                    Me.ListIndex = 0
                ElseIf (Me.ListIndex < Me.ListCount) Then
                    Me.ListIndex = Me.ListIndex + 1
                End If
            End If
    End Select

End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
On Error Resume Next

    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub lstList_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
        Case vbKeyF4
            Me.ButtonDown = False
        Case vbKeyReturn
            If (Shift = 0) Then Call lstList_DblClick
        Case vbKeyEscape
            If (Shift = 0) Then
                Me.Text = mOldValue
                Me.ButtonDown = False
            End If
    End Select

End Sub
'***********************************************************************************************
'Click Events
'***********************************************************************************************
Private Sub txtText_Click()
On Error Resume Next

    RaiseEvent Click
    
End Sub

Private Sub txtText_DblClick()
On Error Resume Next

    RaiseEvent DblClick

End Sub

Private Sub lstList_Click()
On Error Resume Next

    If mNoClick Or mHotTracking Then Exit Sub
    
    Me.Text = mList(lstList.ListIndex + 1).Text(1)

    Me.ButtonDown = False 'Rev. (02/12/02)

End Sub

Private Sub lstList_DblClick()

    Me.Text = mList(lstList.ListIndex + 1).Text(1)
    
    Me.ButtonDown = False

End Sub
'***********************************************************************************************
'Focus Events
'***********************************************************************************************
Private Sub txtText_GotFocus()
On Error Resume Next

    With txtText
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub
'***********************************************************************************************
'Resize Events
'***********************************************************************************************
Private Sub UserControl_Resize()
On Error Resume Next
    Dim lWidth As Long
    Dim lHeight As Long

    lWidth = UserControl.Width - ((GetSystemMetricsAPI(ByVal SM_CXHTHUMB) + 4) * Screen.TwipsPerPixelX)
    lHeight = UserControl.Height - 30

    txtText.Move 0, 30, lWidth, lHeight
    
    If mDropDownWidthAutoSize Then mDropDownWidth = UserControl.Width

End Sub
'***********************************************************************************************
'Paint Events
'***********************************************************************************************
Private Sub UserControl_Paint()
On Error Resume Next

    Call pDrawButton

End Sub
'***********************************************************************************************
'Control Property Events
'***********************************************************************************************
Private Sub UserControl_InitProperties()
On Error Resume Next

    Me.ColumnCount = 1
    Me.DropDownWidth = UserControl.Width
    Me.Text = UserControl.Extender.Name

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

    Me.Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    Me.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Me.ColumnCount = PropBag.ReadProperty("ColumnCount", 1)
    Me.DropDownHeight = PropBag.ReadProperty("DropDownHeight", 0)
    Me.DropDownWidth = PropBag.ReadProperty("DropDownWidth", UserControl.Width)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Me.Font = PropBag.ReadProperty("Font", txtText.Font)
    Me.FontBold = PropBag.ReadProperty("FontBold", False)
    Me.FontItalic = PropBag.ReadProperty("FontItalic", False)
    Me.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    Me.FontSize = PropBag.ReadProperty("FontSize", 8)
    Me.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    Me.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Me.Locked = PropBag.ReadProperty("Locked", False)
    Me.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Me.Style = PropBag.ReadProperty("Style", vbComboDropdown)
    Me.Text = PropBag.ReadProperty("Text")
    Me.ToolTip = PropBag.ReadProperty("ToolTip")

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("Alignment", Me.Alignment, vbLeftJustify)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H80000005)
    Call PropBag.WriteProperty("ColumnCount", Me.ColumnCount, 1)
    Call PropBag.WriteProperty("DropDownHeight", Me.DropDownHeight, 0)
    Call PropBag.WriteProperty("DropDownWidth", Me.DropDownWidth, UserControl.Width)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("Font", Me.Font, txtText.Font)
    Call PropBag.WriteProperty("FontBold", Me.FontBold, False)
    Call PropBag.WriteProperty("FontItalic", Me.FontItalic, False)
    Call PropBag.WriteProperty("FontName", Me.FontName, "MS Sans Serif")
    Call PropBag.WriteProperty("FontSize", Me.FontSize, 8)
    Call PropBag.WriteProperty("FontStrikethru", Me.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", Me.FontUnderline, False)
    Call PropBag.WriteProperty("ForeColor", Me.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", Me.Locked, False)
    Call PropBag.WriteProperty("MaxLength", Me.MaxLength, 0)
    Call PropBag.WriteProperty("Style", Me.Style, vbComboDropdown)
    Call PropBag.WriteProperty("Text", Me.Text)
    Call PropBag.WriteProperty("ToolTip", Me.ToolTip)

End Sub
'***********************************************************************************************
'Control Initialize/Terminate Events
'***********************************************************************************************
Private Sub UserControl_Initialize()
On Error Resume Next
    
    Set mList = New Collection

End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
    
    Set mList = Nothing

End Sub

Private Property Let dwdSubClass_MsgResponse(ByVal RHS As EMsgResponse)
'Needed for Implements
End Property

Private Property Get dwdSubClass_MsgResponse() As EMsgResponse
    dwdSubClass_MsgResponse = emrConsume
End Property

Private Function dwdSubClass_WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WA_INACTIVE = 0
    Dim lLoW As Long
    Dim lHiW As Long

    Select Case wMsg
        Case WM_ACTIVATE
            Call pGetHiLoWord(wParam, lLoW, lHiW)
            If (lLoW = WA_INACTIVE) Then Me.ButtonDown = False
    End Select

End Function
