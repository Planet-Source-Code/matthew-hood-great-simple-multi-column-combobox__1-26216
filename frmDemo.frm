VERSION 5.00
Object = "*\AdwdMComboOCX.vbp"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dwdMCComboBox - Demo"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   9930
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFind 
      Caption         =   "Find"
      Height          =   1515
      Left            =   6960
      TabIndex        =   40
      Top             =   60
      Width           =   2895
      Begin VB.CommandButton cmdFindItem 
         Caption         =   "Find Item"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   45
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtFindText 
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Top             =   300
         Width           =   2655
      End
      Begin VB.CheckBox chkFindExact 
         Caption         =   "Find Exact Match"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   1635
      End
      Begin VB.TextBox txtFindStart 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2100
         TabIndex        =   42
         Text            =   "0"
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtFindColumn 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   780
         TabIndex        =   41
         Text            =   "1"
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblLabels 
         Caption         =   "Start:"
         Height          =   255
         Index           =   10
         Left            =   1620
         TabIndex        =   47
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblLabels 
         Caption         =   "Column:"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   46
         Top             =   780
         Width           =   675
      End
   End
   Begin VB.Frame fraDemo 
      Caption         =   "Control Demo"
      Height          =   1755
      Left            =   60
      TabIndex        =   28
      Top             =   3900
      Width           =   3795
      Begin dwdMCComboOCX.dwdMCCombo dwdMCCombo1 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   1380
         Width           =   3555
         _extentx        =   6271
         _extenty        =   556
         fontsize        =   8.25
         text            =   "dwdMCCombo1"
         tooltip         =   ""
      End
      Begin VB.ComboBox cboColumnText 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox txtColumnText 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   2835
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "dwdMCComboBox Control"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Column Text:"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Control Events"
      Height          =   3975
      Left            =   3960
      TabIndex        =   19
      Top             =   1680
      Width           =   5895
      Begin VB.ListBox lstDebug 
         Height          =   2790
         Left            =   120
         TabIndex        =   27
         Top             =   660
         Width           =   5655
      End
      Begin VB.CheckBox chkMouse 
         Caption         =   "Mouse"
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   300
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkClick 
         Caption         =   "Click"
         Height          =   195
         Left            =   4020
         TabIndex        =   24
         Top             =   300
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkKeyboard 
         Caption         =   "Keyboard"
         Height          =   195
         Left            =   1800
         TabIndex        =   23
         Top             =   300
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkFocus 
         Caption         =   "Focus"
         Height          =   195
         Left            =   3000
         TabIndex        =   22
         Top             =   300
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkChange 
         Caption         =   "Change"
         Height          =   195
         Left            =   4860
         TabIndex        =   21
         Top             =   300
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CommandButton cmdClearDebug 
         Caption         =   "Clear"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   3540
         Width           =   5655
      End
      Begin VB.Label lblLabels 
         Caption         =   "Watch:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "List"
      Height          =   1515
      Left            =   3960
      TabIndex        =   13
      Top             =   60
      Width           =   2895
      Begin VB.CommandButton cmdClearList 
         Caption         =   "Clear List"
         Height          =   315
         Left            =   1500
         TabIndex        =   34
         Top             =   1080
         Width           =   1260
      End
      Begin VB.CommandButton cmdAdd100Items 
         Caption         =   "Add 100 Items"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1260
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         TabIndex        =   16
         Top             =   720
         Width           =   1260
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add New Item"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1260
      End
      Begin VB.TextBox txtItemText 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraProps 
      Caption         =   "Properties"
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3795
      Begin VB.CommandButton cmdGetProperties 
         Caption         =   "Get Properties"
         Height          =   315
         Left            =   1980
         TabIndex        =   39
         Top             =   3300
         Width           =   1695
      End
      Begin VB.CommandButton cmdSetProperties 
         Caption         =   "Set Properties"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   3300
         Width           =   1695
      End
      Begin VB.TextBox txtColumnWidth 
         Height          =   315
         Left            =   2340
         TabIndex        =   36
         Top             =   2460
         Width           =   1335
      End
      Begin VB.ComboBox cboColumnWidth 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txtControlWidth 
         Height          =   315
         Left            =   1620
         TabIndex        =   18
         Top             =   2820
         Width           =   2055
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         ItemData        =   "frmDemo.frx":030A
         Left            =   1620
         List            =   "frmDemo.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2100
         Width           =   2055
      End
      Begin VB.TextBox txtMaxLength 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox txtDropDownWidth 
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   1380
         Width           =   2055
      End
      Begin VB.TextBox txtDropDownHeight 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   1020
         Width           =   2055
      End
      Begin VB.ComboBox cboAlignment 
         Height          =   315
         ItemData        =   "frmDemo.frx":033E
         Left            =   1620
         List            =   "frmDemo.frx":034B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtColumnCount 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Caption         =   "Column Widths"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   37
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Control Width:"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "Style:"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "MaxLength:"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "DropDownWidth:"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "DropDownHeight:"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "ColumnCount:"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "Alignment:"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Checks specified property values for valid data.
Private Sub pEnablePropChanges()
On Error GoTo On_Error
    Dim bDisabled As Boolean

    'Check for an appropriate value and at least 1 column.
    If (CLng(Nz(txtColumnCount.Text, -1)) < 1) Then bDisabled = True
        'Must have at least 1 column.

    'Check for an appropriate value.
    If (CLng(Nz(txtDropDownHeight.Text, -1)) < 0) Then bDisabled = True
        'Value must be in twips.
        'If the value is 0, then the DropDownHeight is automatic. (Lists up to 8 items)
    
    'Check for an appropriate value.
    If (CLng(Nz(txtDropDownWidth.Text, -2)) < -1) Then bDisabled = True
        'Value must be in twips.
        'If the value is 0, then the DropDownWidth will equal the control width.
    
    'Check for an appropriate value.
    If (CLng(Nz(txtMaxLength, -1)) < 0) Then bDisabled = True
        'This property behaves just like the TextBox.MaxLength property.
    
    If (CLng(Nz(txtControlWidth.Text, -1)) < 300) Or (CLng(Nz(txtControlWidth.Text, -1)) > (fraProps.Width - 240)) Then bDisabled = True

On_Exit:
    cmdSetProperties.Enabled = Not bDisabled
    Exit Sub
On_Error:
    Debug.Print Err.Number, Err.Description
    bDisabled = True
    Resume On_Exit
End Sub

'Applies the specified property values to the dwdMCComboBox control.
Private Sub cmdSetProperties_Click()
On Error GoTo On_Error
    Dim l As Long
    
    With dwdMCCombo1
        .Alignment = cboAlignment.ItemData(cboAlignment.ListIndex)
        .ColumnCount = Nz(txtColumnCount.Text, 1)
        For l = 1 To .ColumnCount
            If (l <= cboColumnWidth.ListCount) Then .ColumnWidth(l) = cboColumnWidth.ItemData(l - 1)
        Next l
        .DropDownHeight = Nz(txtDropDownHeight.Text, 0)
        .DropDownWidth = Nz(txtDropDownWidth.Text, 0)
        .MaxLength = Nz(txtMaxLength.Text, 0)
        .Style = cboStyle.ItemData(cboStyle.ListIndex)
        .Width = txtControlWidth.Text
        .Clear
        .ListIndex = 0
    End With

On_Exit:
    Call cmdGetProperties_Click
    Exit Sub
On_Error:
    MsgBox "Location: " & "cmdSetProperties_Click" & vbNewLine & Err.Description, vbCritical, "Program Error: " & Err.Number
    Resume On_Exit
End Sub

'Retrieves the property values from the dwdMCComboBox control.
Private Sub cmdGetProperties_Click()
On Error Resume Next
    Dim l As Long
    
    With dwdMCCombo1
        cboAlignment.ListIndex = .Alignment
        txtColumnCount.Text = .ColumnCount
        cboColumnWidth.Clear
        cboColumnText.Clear
        For l = 1 To .ColumnCount
            cboColumnWidth.AddItem l
            cboColumnText.AddItem l
            cboColumnWidth.ItemData(l - 1) = .ColumnWidth(l)
        Next l
        cboColumnWidth.ListIndex = 0
        cboColumnText.ListIndex = 0
        txtDropDownHeight.Text = .DropDownHeight
        txtDropDownWidth.Text = .DropDownWidth
        txtMaxLength.Text = .MaxLength
        cboStyle.ListIndex = .Style
        txtControlWidth.Text = .Width
    End With

End Sub

'Adds an item to the dwdMCComboBox list.
Private Sub cmdAddItem_Click()
    Dim lCol As Long
    Dim lIndex As Long
    
    lIndex = dwdMCCombo1.AddItem(txtItemText.Text)
    
    'Add Column text to the new list item.
    For lCol = 2 To dwdMCCombo1.ColumnCount
        dwdMCCombo1.ItemText(lIndex, lCol) = "Item #" & lIndex & " " & "Col. #" & lCol
    Next lCol

End Sub

'Adds 100 items to the dwdMCComboBox list.
Private Sub cmdAdd100Items_Click()
    Dim lCol As Long
    Dim lColumnCnt As Long
    Dim lIndex As Long
    Dim lStart As Long
    Dim l As Long
    
    lStart = dwdMCCombo1.ListCount
    lColumnCnt = dwdMCCombo1.ColumnCount

    For l = 1 To 100
        lIndex = dwdMCCombo1.AddItem("Item #" & (lStart + l))
        For lCol = 2 To lColumnCnt
            dwdMCCombo1.ItemText(lIndex, lCol) = "Item #" & lIndex & " " & "Col. #" & lCol
        Next lCol
    Next l

End Sub

'Removes the specified item (if found) from the dwdMCComboBox list.
Private Sub cmdRemoveItem_Click()
    Dim lIndex As Long
    
    lIndex = dwdMCCombo1.FindItem(txtItemText.Text, , , True)
    If (lIndex <> 0) Then
        dwdMCCombo1.RemoveItem lIndex
    Else
        MsgBox "No Item Found", vbExclamation
    End If

End Sub

'Clears the dwdMCComboBox list.
Private Sub cmdClearList_Click()

    dwdMCCombo1.Clear
    dwdMCCombo1.ListIndex = 0

End Sub

'Finds the specified item in the dwdMCComboBox list.
Private Sub cmdFindItem_Click()
    Dim lIndex As Long
    Dim bExact As Long
    Dim lFindInColumn As Long
    Dim lStart As Long
    Dim sTxt As String
    Dim lCol As Long

    'Search for exact string?
    bExact = (chkFindExact.Value = 1)

    'Search in what column?
    lFindInColumn = txtFindColumn.Text
    
    'Where do we start?
    lStart = txtFindStart.Text

    lIndex = dwdMCCombo1.FindItem(txtFindText.Text, lFindInColumn, lStart, bExact)

    If (lIndex = 0) Then
        'No Item found.
        MsgBox "No Item Found.", vbExclamation
    Else
        For lCol = 1 To dwdMCCombo1.ColumnCount
            sTxt = sTxt & dwdMCCombo1.ItemText(lIndex, lCol) & " "
        Next lCol
        sTxt = Trim$(sTxt)
        lstDebug.AddItem sTxt, 0
    End If

End Sub

'Returns an empty string or optional string if value is Null.
Private Function Nz(ByVal Value, Optional ByVal ValueIfNull As String) As String
On Error Resume Next

   Nz = IIf(IsNull(Value) Or Trim$(Value) = vbNullString, ValueIfNull, Trim$(Value))

End Function
'************************************
'dwdMCComboBox Events
'************************************
Private Sub dwdMCCombo1_ButtonClick(ByVal ButtonDown As Boolean)
    If (chkClick.Value = 1) Then lstDebug.AddItem "ButtonClick(" & ButtonDown & ")", 0
End Sub

Private Sub dwdMCCombo1_Change()
    If (chkChange.Value = 1) Then lstDebug.AddItem "Change() , ListIndex: " & dwdMCCombo1.ListIndex, 0
    Call cboColumnText_Click
End Sub

Private Sub dwdMCCombo1_Click()
    If (chkClick.Value = 1) Then lstDebug.AddItem "Click()", 0
End Sub

Private Sub dwdMCCombo1_DblClick()
    If (chkClick.Value = 1) Then lstDebug.AddItem "DblClick()", 0
End Sub

Private Sub dwdMCCombo1_DragDrop(Source As Control, X As Single, Y As Single)
    If (chkMouse.Value = 1) Then lstDebug.AddItem "DragDrop(" & Source.Name & " , " & X & " , " & Y & ")", 0
End Sub

Private Sub dwdMCCombo1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If (chkMouse.Value = 1) Then lstDebug.AddItem "DragOver(" & Source.Name & " , " & X & " , " & Y & ")", 0
End Sub

Private Sub dwdMCCombo1_GotFocus()
    If (chkFocus.Value = 1) Then lstDebug.AddItem "GotFocus()", 0
End Sub

Private Sub dwdMCCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (chkKeyboard.Value = 1) Then lstDebug.AddItem "KeyDown(" & KeyCode & " , " & Shift & ")", 0
End Sub

Private Sub dwdMCCombo1_KeyPress(KeyAscii As Integer)
    If (chkKeyboard.Value = 1) Then lstDebug.AddItem "KeyPress(" & KeyAscii & ")", 0
End Sub

Private Sub dwdMCCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (chkKeyboard.Value = 1) Then lstDebug.AddItem "KeyUp(" & KeyCode & " , " & Shift & ")", 0
End Sub

Private Sub dwdMCCombo1_LostFocus()
    If (chkFocus.Value = 1) Then lstDebug.AddItem "LostFocus()", 0
End Sub

Private Sub dwdMCCombo1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (chkMouse.Value = 1) Then lstDebug.AddItem "MouseDown(" & Button & " , " & Shift & " , " & X & " , " & Y & ")", 0
End Sub

Private Sub dwdMCCombo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (chkMouse.Value = 1) Then lstDebug.AddItem "MouseMove(" & Button & " , " & Shift & " , " & X & " , " & Y & ")", 0
End Sub

Private Sub dwdMCCombo1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (chkMouse.Value = 1) Then lstDebug.AddItem "MouseUp(" & Button & " , " & Shift & " , " & X & " , " & Y & ")", 0
End Sub

Private Sub dwdMCCombo1_Validate(Cancel As Boolean)
    If (chkFocus.Value = 1) Then lstDebug.AddItem "Validate(" & Cancel & ")", 0
End Sub
'************************************
'Demo Form/Control Events
'************************************
Private Sub Form_Load()
    Call cmdGetProperties_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmDemo = Nothing
End Sub

Private Sub cboColumnWidth_Click()
    txtColumnWidth.Text = cboColumnWidth.ItemData(cboColumnWidth.ListIndex)
End Sub

Private Sub cboColumnText_Click()
    txtColumnText.Text = dwdMCCombo1.ItemText(dwdMCCombo1.ListIndex, cboColumnText.Text)
End Sub

Private Sub txtColumnCount_Change()
    Call pEnablePropChanges
End Sub

Private Sub txtColumnWidth_Change()
On Error Resume Next
    cboColumnWidth.ItemData(cboColumnWidth.ListIndex) = Nz(txtColumnWidth.Text, 0)
    Call pEnablePropChanges
End Sub

Private Sub txtControlWidth_Change()
    Call pEnablePropChanges
End Sub

Private Sub txtDropDownHeight_Change()
    Call pEnablePropChanges
End Sub

Private Sub txtDropDownWidth_Change()
    Call pEnablePropChanges
End Sub

Private Sub txtFindStart_Change()
    Call txtFindText_Change
End Sub

Private Sub txtFindColumn_Change()
    Call txtFindText_Change
End Sub

Private Sub txtFindText_Change()
On Error GoTo On_Error
    Dim bDisabled As Boolean
    
    If (Nz(txtFindText.Text) = vbNullString) Then bDisabled = True
    If (CLng(Nz(txtFindStart.Text, -1)) < 0) Then bDisabled = True
    If (CLng(Nz(txtFindColumn.Text, -1)) < 1) Or (CLng(Nz(txtFindColumn.Text, -1)) > dwdMCCombo1.ColumnCount) Then bDisabled = True

On_Exit:
    cmdFindItem.Enabled = Not bDisabled
    Exit Sub
On_Error:
    Debug.Print Err.Number, Err.Description
    bDisabled = True
    Resume On_Exit
End Sub

Private Sub txtItemText_Change()
    cmdAddItem.Enabled = (Nz(txtItemText.Text) <> vbNullString)
    cmdRemoveItem.Enabled = (Nz(txtItemText.Text) <> vbNullString)
End Sub

Private Sub txtMaxLength_Change()
    Call pEnablePropChanges
End Sub

Private Sub cmdClearDebug_Click()
    lstDebug.Clear
End Sub

Private Sub mnuFile_Click(Index As Integer)
    If (Index = 0) Then Unload Me
End Sub
