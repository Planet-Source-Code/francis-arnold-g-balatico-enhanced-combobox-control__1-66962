VERSION 5.00
Begin VB.UserControl eCombobox 
   BackColor       =   &H00000000&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ClipBehavior    =   0  'None
   DataBindingBehavior=   1  'vbSimpleBound
   FillStyle       =   0  'Solid
   PaletteMode     =   4  'None
   ScaleHeight     =   765
   ScaleWidth      =   2715
   ToolboxBitmap   =   "eCombobox.ctx":0000
   Begin VB.Timer tmrPos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1050
      Top             =   150
   End
   Begin eCombo.FlatPanel fltContain 
      Height          =   690
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1217
      Begin VB.ComboBox cboUC2 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   345
         Width           =   2625
      End
      Begin VB.ComboBox cboUC 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   2625
      End
   End
   Begin VB.Shape shpShape 
      BorderColor     =   &H00000000&
      Height          =   240
      Left            =   30
      Top             =   45
      Width           =   2550
   End
End
Attribute VB_Name = "eCombobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'declarations for getting the mouse cursor coordinates
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Enum AlignConstants
   alignLeft
   alignRight
End Enum

'APIs to change the window style
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_RIGHT As Long = &H1000
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000

'SendMessage API
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'constants for combobox messages
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_ADDSTRING = &H143
Private Const CB_ERR = (-1)
Private Const WM_SETTEXT = &HC
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_RESETCONTENT = &H14B
Private Const CB_INITSTORAGE = &H161

Private Const WM_PAINT = &HF
Private Const WM_SETREDRAW = &HB

'usercontrol dimension
Dim uWidth As Long
Dim uHeight As Long

'backcolor variables
Dim oleNormalBackColor As OLE_COLOR
Dim oleFocusBackColor As OLE_COLOR
Dim oleDisabledBackColor As OLE_COLOR

'groove colors
Dim oleNormalGrooveBackColor As OLE_COLOR
Dim oleFocusGrooveBackColor As OLE_COLOR
Dim oleDisabledGrooveBackColor As OLE_COLOR

'fontcolor variables
Dim oleFocusFontColor As OLE_COLOR
Dim oleNormalFontColor As OLE_COLOR

'bordercolor variables
Dim oleNormalBorderColor As OLE_COLOR
Dim oleFocusBorderColor As OLE_COLOR
Dim oleDisabledBorderColor As OLE_COLOR

'alignment variables
Dim lAlign As Long

'border pattern enumeration
Enum BorderPattern
    Transparent
    Solid
    Dash
    Dot
    DashDot
    DashDotDot
    InsideSolid
End Enum

'type description for combobox
Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

'combo style enumeration
Enum ComboStyle
    DropDownCombo
    DropDownList
End Enum

'combo style variable
Dim lStyle As Long

'border pattern variable
Dim lNormalBorderPattern As Long
Dim lFocusBorderPattern As Long
Dim lDisabledBorderPattern As Long

'determines when a resize event of the user control will execute commands or not
'set to false when manually resizing through code
'reset to true after resize to accept inherent resize of the control
Dim sizeFlag As Boolean

'font variables
Dim fntNormal As StdFont

'button focus flag
Dim bFocus As Boolean

'flag for "exit when enter key pressed functionality"
Dim bEnterExit As Boolean

'clsDoEvents object
Dim oDoEvents As clsDoEvents

'recordset object
Dim rsData As Recordset

'flag for backspace key
Dim bkSpace As Boolean

'variable for the search string
Dim sString As String

'flag to control the change event firing
Dim cFlag As Boolean

'stores the previous text
Private prevString As String
Private nowString As String

Private bDown As Boolean

'list index
Dim lstIndex As Double
    
'Events declaration
Event DropDown()
Event DblClick()
Event Change()
Event Click()
Attribute Click.VB_MemberFlags = "200"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)

Private Sub cboUC_Click()
    RaiseEvent Click
        
    nowString = cboUC.Text
End Sub

Private Sub cboUC2_Click()
    RaiseEvent Click
End Sub

Private Sub cboUC_DropDown()
    RaiseEvent DropDown
    
    NormalColors True
End Sub

'set colors when on focus
Private Sub cboUC_GotFocus()
    
    tmrPos.Enabled = True
    
    NormalColors True
End Sub

'reset colors to normal when not in focus
Private Sub cboUC_LostFocus()
    Dim lstIndx As Long
    
    NormalColors False
    
    SendMessage cboUC.hwnd, CB_SHOWDROPDOWN, False, 1
    
    bFocus = False
    
    tmrPos.Enabled = False
                
    lstIndx = FindMatch(cboUC, nowString, True)
    
End Sub

Private Sub cboUC2_Change()
    If Ambient.UserMode = False Then Exit Sub
    RaiseEvent Change
End Sub

Private Sub cboUC2_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub cboUC2_DropDown()
    RaiseEvent DropDown
    
    NormalColors True
End Sub

Private Sub cboUC2_GotFocus()
    NormalColors True
End Sub

Private Sub cboUC2_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboUC2_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cboUC2_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    
    Dim lstIndx As Long
       
    If KeyCode = vbKeyReturn And bEnterExit = True Then
        On Error Resume Next
        
        SendKeys "{TAB}"
        
    End If
    
End Sub

Private Sub cboUC2_LostFocus()
    NormalColors False
End Sub

'checks the position of the mousecursor
Private Sub tmrPos_Timer()
    If GetMouseOver(cboUC.hwnd) = True And bFocus = False Then
        bFocus = True
        SendMessage cboUC.hwnd, CB_SHOWDROPDOWN, True, 1
    ElseIf GetMouseOver(cboUC.hwnd) = False And bFocus = False Then
        bFocus = True
    End If
End Sub

'*******************************************************************
'* Subs
'*******************************************************************

Private Sub UserControl_Resize()
    On Error GoTo errHandlerSize
    
    If sizeFlag = False Then Exit Sub
    
    uWidth = UserControl.Width
    uHeight = UserControl.Height
    
    If uWidth < 375 Then
        uWidth = 375
        UserControl.Width = uWidth
    End If
    
    fltContain.Move 25, 25, uWidth - 70, uHeight - 50
    
    cboUC.top = -25
    cboUC.left = -25
    cboUC2.top = -25
    cboUC2.left = -25
    
    cboUC.Width = fltContain.Width + 75
    cboUC.Height = fltContain.Height + 70
    
    cboUC2.Width = fltContain.Width + 75
    cboUC2.Height = fltContain.Height + 70
    
    shpShape.Move 0, 0, uWidth, uHeight
    
Exit Sub

errHandlerSize:
    sizeFlag = False
        
    fltContain.Height = cboUC.Height - 70
    
    UserControl.Height = fltContain.Height + 70
    
    uHeight = UserControl.Height
    
    sizeFlag = True
    Resume Next
End Sub

'*******************************************************************
'* Properties
'*******************************************************************

'==================================
'   Properties : Combobox Style
'==================================

Public Property Get Style() As ComboStyle

    Style = lStyle
    
End Property

Public Property Let Style(ByVal vNewValue As ComboStyle)
    If Ambient.UserMode = True Then
        MsgBox "Style property can only be set at design time.", vbOKOnly + vbInformation, "Denied"
        Exit Property
    End If
    
    lStyle = vNewValue
    
    If lStyle = 0 Then
        cboUC.Visible = True
        cboUC2.Visible = False
    Else
        cboUC.Visible = False
        cboUC2.Visible = True
    End If
        
    PropertyChanged "DisabledBackColor"

End Property

'==================================
'   Properties : TextAlign
'==================================

Public Property Get ItemsAlign() As AlignConstants

    ItemsAlign = lAlign

End Property

Public Property Let ItemsAlign(ByVal vNewValue As AlignConstants)

    lAlign = vNewValue
        
    ComboListAlign lAlign
        
    PropertyChanged "ItemsAlign"

End Property


'==================================
'   Properties : NormalBackColor
'==================================

Public Property Get NormalBackColor() As OLE_COLOR

    NormalBackColor = oleNormalBackColor

End Property

Public Property Let NormalBackColor(ByVal vNewValue As OLE_COLOR)

    oleNormalBackColor = vNewValue

    cboUC.BackColor = oleNormalBackColor   'this is the back Color
    cboUC2.BackColor = oleNormalBackColor   'this is the back Color
    
    PropertyChanged "NormalBackColor"

End Property

'==================================
'   Properties : FocusBackColor
'==================================

Public Property Get FocusBackColor() As OLE_COLOR

    FocusBackColor = oleFocusBackColor

End Property

Public Property Let FocusBackColor(ByVal vNewValue As OLE_COLOR)

    oleFocusBackColor = vNewValue
    
    PropertyChanged "FocusBackColor"

End Property

'==================================
'   Properties : DisabledBackColor
'==================================

Public Property Get DisabledBackColor() As OLE_COLOR

    DisabledBackColor = oleDisabledBackColor

End Property

Public Property Let DisabledBackColor(ByVal vNewValue As OLE_COLOR)

    oleDisabledBackColor = vNewValue
    
    If cboUC.Enabled = False Then
        cboUC.BackColor = oleDisabledBackColor
        cboUC2.BackColor = oleDisabledBackColor
    Else
        cboUC.BackColor = oleNormalBackColor
        cboUC2.BackColor = oleNormalBackColor
    End If
    
    PropertyChanged "DisabledBackColor"

End Property

'==================================
'   Properties : FocusBorderPattern
'==================================
Public Property Get FocusBorderPattern() As BorderPattern

    FocusBorderPattern = lFocusBorderPattern

End Property

Public Property Let FocusBorderPattern(ByVal vNewValue As BorderPattern)

    lFocusBorderPattern = vNewValue
    PropertyChanged "FocusBorderPattern"

End Property

'==================================
'   Properties : NormalBorderPattern
'==================================
Public Property Get NormalBorderPattern() As BorderPattern

    NormalBorderPattern = lNormalBorderPattern

End Property

Public Property Let NormalBorderPattern(ByVal vNewValue As BorderPattern)

    lNormalBorderPattern = vNewValue
    
    If cboUC.Enabled = True Then
        shpShape.BorderStyle = lNormalBorderPattern
    Else
        shpShape.BorderStyle = lDisabledBorderPattern
    End If
    
    PropertyChanged "NormalBorderPattern"

End Property

'==================================
'   Properties : DisabledBorderPattern
'==================================
Public Property Get DisabledBorderPattern() As BorderPattern

    DisabledBorderPattern = lDisabledBorderPattern

End Property

Public Property Let DisabledBorderPattern(ByVal vNewValue As BorderPattern)

    lDisabledBorderPattern = vNewValue
    
    If cboUC.Enabled = True Then
        shpShape.BorderStyle = lNormalBorderPattern
    Else
        shpShape.BorderStyle = lDisabledBorderPattern
    End If
    
    PropertyChanged "DisabledBorderPattern"

End Property

'==================================
'   Properties : FocusBorderColor
'==================================

Public Property Get FocusBorderColor() As OLE_COLOR

    FocusBorderColor = oleFocusBorderColor

End Property

Public Property Let FocusBorderColor(ByVal vNewValue As OLE_COLOR)

    oleFocusBorderColor = vNewValue
    PropertyChanged "FocusBorderColor"

End Property


'==================================
'   Properties : NormalBorderColor
'==================================

Public Property Get NormalBorderColor() As OLE_COLOR

    NormalBorderColor = oleNormalBorderColor

End Property

Public Property Let NormalBorderColor(ByVal vNewValue As OLE_COLOR)

    oleNormalBorderColor = vNewValue
    shpShape.BorderColor = vNewValue
    PropertyChanged ("NormalBorderColor")

End Property

'==================================
'   Properties : DisabledBorderColor
'==================================

Public Property Get DisabledBorderColor() As OLE_COLOR

    DisabledBorderColor = oleDisabledBorderColor

End Property

Public Property Let DisabledBorderColor(ByVal vNewValue As OLE_COLOR)

    oleDisabledBorderColor = vNewValue
    
    If cboUC.Enabled = False Then
        shpShape.BorderColor = oleDisabledBorderColor
    End If
    
    PropertyChanged "DisabledBorderColor"

End Property

'==================================
'   Properties : NormalFont
'==================================
Public Property Get NormalFont() As Font

    Set NormalFont = fntNormal

End Property

Public Property Set NormalFont(ByRef vNewValue As Font) 'Make sure this is Pass ByReference and method is set

    Set fntNormal = vNewValue
    
    'resize the uc depending on the font resize which results to cbo resize
    sizeFlag = False
    
    Set cboUC.Font = vNewValue
    Set cboUC2.Font = vNewValue
    
    If lStyle = 0 Then
        fltContain.Height = cboUC.Height - 70
    Else
        fltContain.Height = cboUC2.Height - 70
    End If
    
    UserControl.Height = fltContain.Height + 70
    
    uHeight = UserControl.Height
    
    shpShape.Move 0, 0, uWidth, uHeight
    
    sizeFlag = True
    
    PropertyChanged ("NormalFont")

End Property


'==================================
'   Properties : NormalFontColor
'==================================

Public Property Get NormalFontColor() As OLE_COLOR

    NormalFontColor = oleNormalFontColor

End Property

Public Property Let NormalFontColor(ByVal vNewValue As OLE_COLOR)

    oleNormalFontColor = vNewValue
    
    cboUC.ForeColor = oleNormalFontColor
    cboUC2.ForeColor = oleNormalFontColor
    
    PropertyChanged "NormalFontColor"

End Property

'==================================
'   Properties : FocusFontColor
'==================================
Public Property Get FocusFontColor() As OLE_COLOR

    FocusFontColor = oleFocusFontColor

End Property

Public Property Let FocusFontColor(ByVal vNewValue As OLE_COLOR)

    oleFocusFontColor = vNewValue
    
        
    PropertyChanged "FocusFontColor"

End Property


'==================================
'   Properties : DisabledGrooveBackColor
'==================================
Public Property Get DisabledGrooveBackColor() As OLE_COLOR

    DisabledGrooveBackColor = oleDisabledGrooveBackColor

End Property

Public Property Let DisabledGrooveBackColor(ByVal vNewValue As OLE_COLOR)

    oleDisabledGrooveBackColor = vNewValue
    
    UserControl.AutoRedraw = True
    
    If cboUC.Enabled = False Then
        UserControl.BackColor = oleDisabledGrooveBackColor
    End If
    
    UserControl.AutoRedraw = False
    
    PropertyChanged "DisabledGrooveBackColor"

End Property

'==================================
'   Properties : FocusGrooveBackColor
'==================================
Public Property Get FocusGrooveBackColor() As OLE_COLOR

    FocusGrooveBackColor = oleFocusGrooveBackColor

End Property

Public Property Let FocusGrooveBackColor(ByVal vNewValue As OLE_COLOR)

    oleFocusGrooveBackColor = vNewValue
    PropertyChanged "FocusGrooveBackColor"

End Property

'==================================
'   Properties : NormalGrooveBackColor
'==================================

Public Property Get NormalGrooveBackColor() As OLE_COLOR

    NormalGrooveBackColor = oleNormalGrooveBackColor

End Property

Public Property Let NormalGrooveBackColor(ByVal vNewValue As OLE_COLOR)

    oleNormalGrooveBackColor = vNewValue
    UserControl.AutoRedraw = True
    
    If cboUC.Enabled = True Then
        UserControl.BackColor = vNewValue   'this is the back Color
    End If
    
    UserControl.AutoRedraw = False
    PropertyChanged "NormalGrooveBackColor"

End Property

'==================================
'   Properties : ListCount
'==================================
Public Property Get ListCount() As Long
    If lStyle = 0 Then
        ListCount = cboUC.ListCount
    Else
        ListCount = cboUC2.ListCount
    End If
End Property

'==================================
'   Properties : Enabled
'==================================

Public Property Get Enabled() As Boolean

    Enabled = cboUC.Enabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)

    cboUC.Enabled = vNewValue
    cboUC2.Enabled = vNewValue
    
    If cboUC.Enabled = False And cboUC2.Enabled = False Then
        cboUC.BackColor = oleDisabledBackColor
        cboUC2.BackColor = oleDisabledBackColor
        shpShape.BorderColor = oleDisabledBorderColor
        UserControl.BackColor = oleDisabledGrooveBackColor
        shpShape.BorderStyle = lDisabledBorderPattern
      ElseIf cboUC.Enabled = True And cboUC2.Enabled = True Then
        cboUC.BackColor = oleNormalBackColor
        cboUC2.BackColor = oleNormalBackColor
        shpShape.BorderColor = oleNormalBorderColor
        UserControl.BackColor = oleNormalGrooveBackColor
        shpShape.BorderStyle = lNormalBorderPattern
    End If
   
    PropertyChanged "Enabled"

End Property

'==================================
'   Properties : Text
'==================================
Public Property Get Text() As String
    If lStyle = 0 Then
        Text = cboUC.Text
    Else
        Text = cboUC2.Text
    End If
End Property

Public Property Let Text(vNewValue As String)
    If lStyle = 1 Then
        MsgBox "Property is cannot be set while the style is DropDownList", vbOKOnly + vbInformation, "Denied"
        Exit Property
    End If
    
    cboUC.Text = vNewValue
    PropertyChanged "Text"
End Property

'==================================
'   Properties : Tag
'==================================

Public Property Get Tag() As String

    Tag = cboUC.Tag

End Property

Public Property Let Tag(ByVal vNewValue As String)

    cboUC.Tag = vNewValue
    cboUC2.Tag = cboUC.Tag
    
    PropertyChanged "Tag"

End Property


'==================================
'   Properties : ExitOnEnter
'==================================

Public Property Get ExitOnEnter() As Boolean
    ExitOnEnter = bEnterExit
End Property

Public Property Let ExitOnEnter(ByVal vNewValue As Boolean)
    bEnterExit = vNewValue
    PropertyChanged "ExitOnEnter"
End Property

'==================================
'   Properties : ListIndex
'==================================
Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    If lStyle = 0 Then
        ListIndex = cboUC.ListIndex
    Else
        ListIndex = cboUC2.ListIndex
    End If
End Property

Public Property Let ListIndex(vNewValue As Long)
    If Ambient.UserMode = False Then
        MsgBox "Property can not be set during design time.", vbOKOnly + vbCritical, "Property Locked"
        Exit Property
    End If
    
    If lStyle = 0 Then
        cboUC.ListIndex = vNewValue
    Else
        cboUC2.ListIndex = vNewValue
    End If
    
    PropertyChanged "ListIndex"
End Property

'*******************************************************************
'* READ / WRITE Properties
'*******************************************************************

'Read properties from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim lItemIndex As Long
    
    Set oDoEvents = New clsDoEvents
    
    oDoEvents.QueueUsed = Standard
    
    On Error Resume Next
    'Fonts
    Set fntNormal = PropBag.ReadProperty("NormalFont", Nothing)
    oleNormalFontColor = PropBag.ReadProperty("NormalFontColor", vbBlack)
    oleFocusFontColor = PropBag.ReadProperty("FocusFontColor", vbBlue)
    
    'BG Colors
    oleNormalBackColor = PropBag.ReadProperty("NormalBackColor", &HFFFFFF)
    oleFocusBackColor = PropBag.ReadProperty("FocusBackColor", &H80FFFF)
    oleDisabledBackColor = PropBag.ReadProperty("DisabledBackColor", &HE0E0E0)
    
    'Alignments
    lAlign = PropBag.ReadProperty("ItemsAlign", 0)
    
    'Combobox style
    lStyle = PropBag.ReadProperty("Style", 0)
    
    'Border Colors
    oleFocusBorderColor = PropBag.ReadProperty("FocusBorderColor", &H800000)
    oleNormalBorderColor = PropBag.ReadProperty("NormalBorderColor", &H800000)
    oleDisabledBorderColor = PropBag.ReadProperty("DisabledBorderColor", &H8000000F)
    
    'Border pattern
    lNormalBorderPattern = PropBag.ReadProperty("NormalBorderPattern", 1)
    lFocusBorderPattern = PropBag.ReadProperty("FocusBorderPattern", 1)
    lDisabledBorderPattern = PropBag.ReadProperty("DisabledBorderPattern", 1)
    
    'Groove colors
    oleNormalGrooveBackColor = PropBag.ReadProperty("NormalGrooveBackColor", "&H8000000F")
    oleFocusGrooveBackColor = PropBag.ReadProperty("FocusGrooveBackColor", "&H8000000F")
    oleDisabledGrooveBackColor = PropBag.ReadProperty("DisabledGrooveBackColor", &H8000000F)
    
    'Property assignment
    cboUC.BackColor = oleNormalBackColor
    cboUC2.BackColor = oleNormalBackColor
    
    shpShape.BorderColor = oleNormalBorderColor
    
    cboUC.Text = PropBag.ReadProperty("Text", "Text")
    
    'enabling and sorting
    cboUC.Enabled = PropBag.ReadProperty("Enabled", True)
    cboUC2.Enabled = PropBag.ReadProperty("Enabled", True)
 
    'ExitOnEnter
    bEnterExit = PropBag.ReadProperty("ExitOnEnter", False)
    
    'set combobox fonts
    Set cboUC.Font = fntNormal
    Set cboUC2.Font = fntNormal
    
    'set combobox text colors
    cboUC.ForeColor = oleNormalFontColor
    cboUC2.ForeColor = oleNormalFontColor
    
    'Misc properties
    cboUC.Tag = PropBag.ReadProperty("Tag", Nothing)
    cboUC2.Tag = PropBag.ReadProperty("Tag", Nothing)
    
    'apply combobox style
    If lStyle = 0 Then
        cboUC.Visible = True
        cboUC2.Visible = False
    Else
        cboUC.Visible = False
        cboUC2.Visible = True
    End If
    
    'apply text alignment
    ComboListAlign lAlign
    
    'apply necessary coloring
    If cboUC.Enabled = True And cboUC2.Enabled = True Then
        cboUC.BackColor = oleNormalBackColor
        cboUC2.BackColor = oleNormalBackColor
        UserControl.BackColor = oleNormalGrooveBackColor
        shpShape.BorderColor = oleNormalBorderColor
        shpShape.BorderStyle = lNormalBorderPattern
    Else
        cboUC.BackColor = oleDisabledBackColor
        cboUC2.BackColor = oleDisabledBackColor
        UserControl.BackColor = oleDisabledGrooveBackColor
        shpShape.BorderColor = oleDisabledBorderColor
        shpShape.BorderStyle = lDisabledBorderPattern
    End If
        
    UserControl_Resize
    
End Sub

'Write properties to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Fonts
    PropBag.WriteProperty "NormalFont", fntNormal
    PropBag.WriteProperty "NormalFontColor", oleNormalFontColor, vbBlack
    PropBag.WriteProperty "FocusFontColor", oleFocusFontColor, vbBlue
    
    'Combobox style
    PropBag.WriteProperty "Style", lStyle, 0
    
    'alignment
    PropBag.WriteProperty "ItemsAlign", lAlign, 0

    'border color
    PropBag.WriteProperty "NormalBorderColor", oleNormalBorderColor, &H800000
    PropBag.WriteProperty "DisabledBorderColor", oleDisabledBorderColor, &H8000000F
    PropBag.WriteProperty "FocusBorderColor", oleFocusBorderColor, &H800000
    
    'border pattern
    PropBag.WriteProperty "NormalBorderPattern", lNormalBorderPattern, 1
    PropBag.WriteProperty "FocusBorderPattern", lFocusBorderPattern, 1
    PropBag.WriteProperty "DisabledBorderPattern", lDisabledBorderPattern, 1
    
    'groove colors
    PropBag.WriteProperty "NormalGrooveBackColor", oleNormalGrooveBackColor, &H8000000F
    PropBag.WriteProperty "FocusGrooveBackColor", oleFocusGrooveBackColor, &H8000000F
    PropBag.WriteProperty "DisabledGrooveBackColor", oleDisabledGrooveBackColor, &H8000000F
    
    'bg colors
    PropBag.WriteProperty "NormalBackColor", oleNormalBackColor, &HFFFFFF
    PropBag.WriteProperty "FocusBackColor", oleFocusBackColor, &H80FFFF
    PropBag.WriteProperty "DisabledBackColor", oleDisabledBackColor, &HE0E0E0
    
    'enabling
    PropBag.WriteProperty "Enabled", cboUC.Enabled, True
 
    'ExitOnEnter
    PropBag.WriteProperty "ExitOnEnter", bEnterExit, False
    
    'text
    PropBag.WriteProperty "Text", cboUC.Text, "Text"
    
    'tag
    PropBag.WriteProperty "Tag", cboUC.Tag, Nothing
End Sub

Private Sub UserControl_Initialize()
    Set fntNormal = New StdFont
    
    sizeFlag = True
    
    lNormalBorderPattern = 1
    lFocusBorderPattern = 1
    lDisabledBorderPattern = 1
    
    oleNormalBackColor = &HFFFFFF
    oleDisabledBackColor = &HE0E0E0
    oleFocusBackColor = &H80FFFF
    
    oleNormalBorderColor = &H808080
    oleFocusBorderColor = &H80FF&
    oleDisabledBorderColor = &H8000000F
    
    oleNormalGrooveBackColor = &H8000000F
    oleFocusGrooveBackColor = &H8000000F
    oleDisabledGrooveBackColor = &H8000000F
    
    oleNormalFontColor = vbBlack
    oleFocusFontColor = vbBlue
    
    fntNormal.Name = "Arial"
    fntNormal.Size = 8
    
    bFocus = False
    
    bkSpace = False
    
    bDown = False
    
    cFlag = True
    
    prevString = ""
End Sub

Private Sub UserControl_InitProperties()
    cboUC.BackColor = vbWhite
    cboUC.ForeColor = vbBlack
    
    cboUC2.BackColor = vbWhite
    cboUC2.ForeColor = vbBlack
    
    UserControl.BackColor = oleNormalGrooveBackColor
    
    shpShape.BorderColor = oleNormalBorderColor
    
End Sub

Private Sub UserControl_Terminate()
    Set fntNormal = Nothing
    Set oDoEvents = Nothing
    Set rsData = Nothing
End Sub

'*******************************************************************
'* Procedures and functions
'*******************************************************************

'check position of mouse over the combobox
'returns true if mouse is over the dropdown arrow button
'returns false if otherwise
Private Function GetMouseOver(hwnd As Long) As Boolean

    Dim wRect As RECT
    Dim Mouse As PointAPI
    
    GetCursorPos Mouse
    GetWindowRect hwnd, wRect
    
    If (Mouse.X <= wRect.right - 4 And Mouse.X >= wRect.right - 20) And (Mouse.Y <= wRect.bottom - 4 And Mouse.Y >= wRect.top + 2) Then
        GetMouseOver = True
    Else
        GetMouseOver = False
    End If

End Function

Private Sub cboUC_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub cboUC_Change()
    If Ambient.UserMode = False Then Exit Sub
    
    RaiseEvent Change
    
    If cFlag = False Then Exit Sub
    If Not cboUC.Style = 0 Then Exit Sub
    
    Dim lcount As Long
    
    
    If bkSpace = False And cFlag = True Then
        prevString = cboUC.Text
        
        'force drop-down the combobox when typing text
        SendMessage cboUC.hwnd, CB_SHOWDROPDOWN, True, 1
        cboUC.Parent.MousePointer = 0
        
        lcount = Len(prevString)
        
        FindString cboUC.Text, False, False
        
        If StrComp(prevString, left$(sString, lcount), vbTextCompare) = 0 And Not Len(Trim$(cboUC.Text)) = 0 Then
            cFlag = False
            
            Call SendMessage(cboUC.hwnd, WM_SETTEXT, 0, ByVal sString)
                        
            cFlag = True
            
            cboUC.SelStart = lcount
            cboUC.SelLength = Len(sString) - lcount
        End If
    End If
    
    FindString cboUC.Text, True, False
    
End Sub

Private Sub cboUC_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
        
    If KeyCode = vbKeyReturn Then
                
        If bDown = False Then
            nowString = cboUC.Text
            'remove drop down state of combobox
            SendMessage cboUC.hwnd, CB_SHOWDROPDOWN, False, 1
            
            bDown = True
            Exit Sub
        Else
            
            bDown = False
        End If
        
    End If
End Sub

Private Sub cboUC_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    
    If KeyAscii = 8 Then
        
        bkSpace = True
        
    Else
        bkSpace = False
    End If
    
End Sub

Private Sub cboUC_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    Dim lstIndx As Long
   
    If KeyCode = vbKeyDelete Then
        bkSpace = True
        cFlag = False
        If Len(cboUC.SelText) > 0 Then
            cboUC.SelText = ""
        End If
        
        cboUC.ListIndex = -1
        
        cFlag = True
        bkSpace = False
        
        Exit Sub
    ElseIf KeyCode = vbKeyBack Then
        prevString = cboUC.Text
        
        cboUC.ListIndex = -1
    End If
    
    
    If KeyCode = vbKeyReturn And bEnterExit = True Then
        On Error Resume Next
        
        If cboUC.Style = 0 Then
            lstIndx = FindMatch(cboUC, nowString, True)
            
            cFlag = False
            
            If lstIndx = -1 Then
                
                SendKeys "{TAB}"
                cboUC.Text = nowString
                
            Else
                
                Call SendMessage(cboUC.hwnd, WM_SETTEXT, 0, ByVal cboUC.Text)
                SendKeys "{TAB}"
                
            End If
            
            cFlag = True
        Else
            cFlag = False
            SendKeys "{TAB}"
            cFlag = True
        End If
        
    End If
    
End Sub

'add an item to the combobox
Public Sub AddItem(itmName As String)
    Dim lItem As Long
    
    Select Case lStyle
        Case 0
            If cboUC.ListCount > 32766 Then
                MsgBox "Item limit reached.", vbOKOnly + vbExclamation, "Overflow check"
                Exit Sub
            End If
            
            lItem = SendMessage(cboUC.hwnd, CB_ADDSTRING, 0, ByVal itmName)
            Exit Sub
        Case 1
            If cboUC2.ListCount > 32766 Then
                MsgBox "Item limit reached. Item adding will be restricted", vbOKOnly + vbExclamation, "Overflow check"
                Exit Sub
            End If
            
            lItem = SendMessage(cboUC2.hwnd, CB_ADDSTRING, 0, ByVal itmName)
            Exit Sub
    End Select
    
End Sub

'clear combox content
Public Sub Clear()
    Call SendMessage(cboUC.hwnd, CB_RESETCONTENT, 0, 0)
    Call SendMessage(cboUC2.hwnd, CB_RESETCONTENT, 0, 0)
End Sub

'retrieve data from database to automatically fill the combobox
Public Sub GetDataFromDatabase(vRecordset As ADODB.Recordset, vField As String, Optional BlankInitial As Boolean = True, Optional Trimmed As Boolean = False)
    On Error GoTo ErrHandler
    
    Dim lItemIndex As Long
    
    Dim counter As Long
    Dim count As Long
    
    Set rsData = vRecordset
    
    With rsData
        'move to the first record
        If .RecordCount > 32767 Then
            MsgBox "Number of records exceed the limit for a combobox." & vbCrLf & vbCrLf & "Consider narrowing down items by supplying a criteria.", vbOKOnly + vbExclamation, "Overflow check"
            Exit Sub
        End If
        
        .MoveFirst
        
        count = .RecordCount
        
        If count < 1 Then
            Set rsData = Nothing
            Exit Sub
        End If
                
        'clear contents of combobox
        Clear
        
        'lock combobox from repainting
        lItemIndex = SendMessage(cboUC.hwnd, WM_SETREDRAW, False, 0)
        lItemIndex = SendMessage(cboUC2.hwnd, WM_SETREDRAW, False, 0)
          
        
        'adds a blank item to help in searching non existent items
        AddItem " "
        
        If Trimmed = False Then
            For counter = 1 To count
                AddItem .Collect(vField)
                .MoveNext
                oDoEvents.GetInputState
            Next counter
        Else
            For counter = 1 To count
                AddItem Trim$(.Collect(vField))
                .MoveNext
                oDoEvents.GetInputState
            Next counter
        End If
        
        'unlock combobox for repainting
        lItemIndex = SendMessage(cboUC.hwnd, WM_SETREDRAW, True, 0)
        lItemIndex = SendMessage(cboUC2.hwnd, WM_SETREDRAW, True, 0)
        
        .MoveFirst
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox UserControl.Name & " : An error occured while retrieving data from the database.", vbOKOnly + vbCritical, "Error in retrieving data"
    Set rsData = Nothing
End Sub

'uses the listindex found by findmatch function to move to the specified item in the listbox
Public Function FindString(TextToFind As String, Exact As Boolean, Trimmed As Boolean) As Boolean
    On Error Resume Next
    
    If lStyle = 0 Then
        If Trimmed = False Then
            lstIndex = FindMatch(cboUC, TextToFind, Exact)
        Else
            lstIndex = FindMatch(cboUC, Trim$(TextToFind), Exact)
        End If
        
        cboUC.ListIndex = lstIndex
    Else
        If Trimmed = False Then
            lstIndex = FindMatch(cboUC2, TextToFind, Exact)
        Else
            lstIndex = FindMatch(cboUC2, Trim$(TextToFind), Exact)
        End If
        
        cboUC2.ListIndex = lstIndex
    End If
    
    If lstIndex = -1 Then
        sString = vbNullString
        FindString = False
    Else
        sString = cboUC.Text
        FindString = True
    End If
    
End Function

'set miscellaneous colors
Private Sub NormalColors(bNormal As Boolean)
    If bNormal = True Then
        shpShape.BorderColor = oleFocusBorderColor
        UserControl.BackColor = oleFocusGrooveBackColor
        cboUC.BackColor = oleFocusBackColor
        cboUC.ForeColor = oleFocusFontColor
        
        cboUC2.BackColor = oleFocusBackColor
        cboUC2.ForeColor = oleFocusFontColor
        
        shpShape.BorderStyle = lFocusBorderPattern
    Else
        shpShape.BorderColor = oleNormalBorderColor
        cboUC.BackColor = oleNormalBackColor
        cboUC.ForeColor = oleNormalFontColor
        
        cboUC2.BackColor = oleNormalBackColor
        cboUC2.ForeColor = oleNormalFontColor
        
        UserControl.BackColor = oleNormalGrooveBackColor
        shpShape.BorderStyle = lNormalBorderPattern
    End If
End Sub

'returns the listindex when a string is found
Private Function FindMatch(objX As Object, sStr As String, _
        Optional sExact As Boolean = False, _
        Optional sStart As Long = -1) As Long

      If sExact = True Then
         FindMatch = SendMessage(objX.hwnd, CB_FINDSTRINGEXACT, _
                     sStart, ByVal sStr)
      Else
         FindMatch = SendMessage(objX.hwnd, CB_FINDSTRING, sStart, _
                     ByVal sStr)
      End If
End Function

'get handle of the combobox list
Private Function GetComboListHandle(ctl As ComboBox) As Long

   Dim CBI As COMBOBOXINFO

   CBI.cbSize = Len(CBI)
   Call GetComboBoxInfo(ctl.hwnd, CBI)
   GetComboListHandle = CBI.hwndList

End Function

'aligns items to the specified alignment
Private Sub ComboListAlign(Optional ByVal Align As AlignConstants = alignLeft)

   Dim hList As Long
   Dim hList2 As Long
   Dim nStyle As Long

  'obtain the handle to the list
  'portion of the combo
   hList = GetComboListHandle(cboUC)

  'if valid, change the style
   If hList <> 0 Then
   
      nStyle = GetWindowLong(hList, GWL_EXSTYLE)
      
      Select Case Align
          Case alignRight
              nStyle = nStyle Or WS_EX_RIGHT
          Case Else
              nStyle = nStyle And Not WS_EX_RIGHT
      End Select
      
      SetWindowLong hList, GWL_EXSTYLE, nStyle
   
   End If
   
   
   'obtain the handle to the list
  'portion of the combo
   hList2 = GetComboListHandle(cboUC2)

  'if valid, change the style
   If hList2 <> 0 Then
   
      nStyle = GetWindowLong(hList2, GWL_EXSTYLE)
      
      Select Case Align
          Case alignRight
              nStyle = nStyle Or WS_EX_RIGHT
          Case Else
              nStyle = nStyle And Not WS_EX_RIGHT
      End Select
      
      SetWindowLong hList2, GWL_EXSTYLE, nStyle
   
   End If
End Sub
