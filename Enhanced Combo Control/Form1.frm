VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enhanced Combobox Control"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "ENHANCED COMBO CONTROL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6570
      Left            =   135
      TabIndex        =   18
      Top             =   45
      Width           =   4500
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "* Control over item alignment"
         Height          =   255
         Left            =   105
         TabIndex        =   28
         Top             =   3555
         Width           =   4290
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "* Auto tab on enter functionality"
         Height          =   255
         Left            =   105
         TabIndex        =   27
         Top             =   3195
         Width           =   4290
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "- Francis Arnold Balatico"
         Height          =   420
         Left            =   60
         TabIndex        =   26
         Top             =   4695
         Width           =   4290
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "A vote from you is appreciated and comment / suggestions are welcome"
         Height          =   420
         Left            =   120
         TabIndex        =   25
         Top             =   3975
         Width           =   4290
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "* Two styles supported : DropDownList && DropDownCombo"
         Height          =   255
         Left            =   105
         TabIndex        =   24
         Top             =   2835
         Width           =   4290
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "* Basic autocomplete functionality"
         Height          =   255
         Left            =   105
         TabIndex        =   23
         Top             =   2460
         Width           =   4290
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "* Built-in add from database function"
         Height          =   255
         Left            =   105
         TabIndex        =   22
         Top             =   2100
         Width           =   4290
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "* Built-in search function"
         Height          =   255
         Left            =   105
         TabIndex        =   21
         Top             =   1770
         Width           =   4290
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "* Control over the colors of the different states namely Normal, Focused and Disabled."
         Height          =   435
         Left            =   105
         TabIndex        =   20
         Top             =   1275
         Width           =   4290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0000
         Height          =   825
         Left            =   105
         TabIndex        =   19
         Top             =   270
         Width           =   4275
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Control over disabled state colors"
      Height          =   1545
      Left            =   4740
      TabIndex        =   14
      Top             =   5070
      Width           =   5445
      Begin VB.CommandButton cmdEnDis 
         Caption         =   "Disable / Enable"
         Height          =   360
         Left            =   3315
         TabIndex        =   15
         Top             =   1080
         Width           =   1980
      End
      Begin eCombo.eCombobox ecbDis 
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   300
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusFontColor  =   13001271
         NormalBorderColor=   8421504
         DisabledBorderColor=   16777215
         FocusBorderColor=   33023
         NormalGrooveBackColor=   14737632
         FocusGrooveBackColor=   49344
         Text            =   ""
         Object.Tag             =   ""
      End
      Begin eCombo.eCombobox ecbDis2 
         Height          =   330
         Left            =   135
         TabIndex        =   17
         Top             =   705
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusFontColor  =   13001271
         NormalBorderColor=   8421504
         DisabledBorderColor=   12632319
         FocusBorderColor=   33023
         NormalGrooveBackColor=   14737632
         FocusGrooveBackColor=   49344
         DisabledGrooveBackColor=   8421631
         DisabledBackColor=   8421631
         Text            =   ""
         Object.Tag             =   ""
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fill combobox with user specified items [5000 items] and Search function"
      Height          =   1170
      Left            =   4740
      TabIndex        =   9
      Top             =   3840
      Width           =   5445
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   345
         Left            =   2070
         TabIndex        =   13
         Top             =   705
         Width           =   1170
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Top             =   705
         Width           =   1875
      End
      Begin VB.CommandButton cmdFillItems 
         Caption         =   "Start Fill"
         Height          =   360
         Left            =   4125
         TabIndex        =   10
         Top             =   690
         Width           =   1170
      End
      Begin eCombo.eCombobox ecbFill 
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   300
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusFontColor  =   13001271
         NormalBorderColor=   8421504
         FocusBorderColor=   33023
         NormalGrooveBackColor=   14737632
         FocusGrooveBackColor=   49344
         Text            =   ""
         Object.Tag             =   ""
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fill combobox with items from the database"
      Height          =   1170
      Left            =   4740
      TabIndex        =   6
      Top             =   2580
      Width           =   5445
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill Combobox from DB"
         Height          =   360
         Left            =   3315
         TabIndex        =   8
         Top             =   690
         Width           =   1980
      End
      Begin eCombo.eCombobox ecbDB 
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   300
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusFontColor  =   13001271
         NormalBorderColor=   8421504
         FocusBorderColor=   33023
         NormalGrooveBackColor=   14737632
         FocusGrooveBackColor=   49344
         Text            =   ""
         Object.Tag             =   ""
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Manipulate colors for a super flat look"
      Height          =   780
      Left            =   4740
      TabIndex        =   4
      Top             =   1710
      Width           =   5445
      Begin eCombo.eCombobox eCombobox4 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   300
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusFontColor  =   13001271
         NormalBorderColor=   12632256
         FocusBorderColor=   33023
         NormalGrooveBackColor=   16777215
         FocusGrooveBackColor=   8454143
         Text            =   ""
         Object.Tag             =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control over colors [Try focusing on each combobox for focus colors]"
      Height          =   1575
      Left            =   4740
      TabIndex        =   0
      Top             =   60
      Width           =   5445
      Begin eCombo.eCombobox eCombobox1 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NormalBorderColor=   8421504
         FocusBorderColor=   33023
         Text            =   ""
         Object.Tag             =   ""
      End
      Begin eCombo.eCombobox eCombobox2 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NormalBorderColor=   16752456
         FocusBorderColor=   33023
         NormalGrooveBackColor=   16777088
         FocusGrooveBackColor=   14674668
         NormalBackColor =   13001271
         FocusBackColor  =   15727868
         Text            =   ""
         Object.Tag             =   ""
      End
      Begin eCombo.eCombobox eCombobox3 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
         BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NormalBorderColor=   16777215
         FocusBorderColor=   12648384
         NormalGrooveBackColor=   49344
         FocusGrooveBackColor=   49152
         NormalBackColor =   65535
         FocusBackColor  =   65280
         Text            =   ""
         Object.Tag             =   ""
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bDisabled As Boolean

Dim objConn As New ADODB.Connection
Dim rsData As New ADODB.Recordset

Private Sub cmdEnDis_Click()
    bDisabled = Not bDisabled
    
    If bDisabled = False Then
        ecbDis.Enabled = True
        ecbDis2.Enabled = True
    Else
        ecbDis.Enabled = False
        ecbDis2.Enabled = False
    End If
End Sub

Private Sub cmdFill_Click()
    With ecbDB
        .Clear
        .GetDataFromDatabase rsData, "Name", True, True
    End With
End Sub

Private Sub cmdFillItems_Click()
    Dim counter As Long
    
    ecbFill.Clear
    
    For counter = 1 To 5000
        ecbFill.AddItem "Sample " & counter
    Next counter
    
End Sub

Private Sub cmdSearch_Click()
    ecbFill.FindString txtSearch.Text, True, True
End Sub


Private Sub Form_Load()
    With objConn
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Sample.mdb;Persist Security Info=False"
    End With
    
    With rsData
        If .State = adStateOpen Then
            .Close
        End If
        
        .Open "SELECT * FROM tblSample", objConn, adOpenForwardOnly, adLockReadOnly
    End With
    
    bDisabled = False
End Sub
