VERSION 5.00
Begin VB.Form frmAddBasketItem 
   Caption         =   "Add Item"
   ClientHeight    =   4980
   ClientLeft      =   2790
   ClientTop       =   3180
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   4980
   Begin VB.Frame Frame2 
      Caption         =   "Customer Input"
      Height          =   975
      Left            =   840
      TabIndex        =   18
      Top             =   3240
      Width           =   3255
      Begin VB.TextBox txtOrderQty 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTotalCost 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H0000FF00&
         Caption         =   "Enter Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmbAddItem 
      Caption         =   "Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmbClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Information"
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4695
      Begin VB.Label lblActive 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblUnitPrice 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblStocks 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label lblItem 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblItemType 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblSuplier 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Unit Price"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Active"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "Stocks"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.TextBox txtItemCodeSearch 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label aaa 
      BackColor       =   &H0000FF00&
      Caption         =   "Item Code Search"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddBasketItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private newSearch As Boolean
Private rs As ADODB.Recordset

Private Sub cmbClose_Click()
   Unload Me
End Sub
Private Sub Form_Load()
  newSearch = False
End Sub
Private Sub txtItemCodeSearch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    newSearch = True
    Set rs = DataCrudDao.getItemReg(txtItemCodeSearch)
    If rs.RecordCount > 0 Then
      lblSuplier = CommonHelper.extractStringValue(rs!SUPPLIER)
      lblItemType = CommonHelper.extractStringValue(rs!ITEM_TYPE)
      lblItem = CommonHelper.extractStringValue(rs!ITEM_NAME)
      lblStocks = Val(CommonHelper.extractStringValue(rs!Quantity))
      lblUnitPrice = Format(Val(CommonHelper.extractStringValue(rs!UNIT_PRICE)), Constants.CURRENCY_FORMAT)
      lblActive = CommonHelper.extractStringValue(rs!active)
      txtOrderQty.SetFocus
    Else
      MsgBox "Item Does not Exist", vbCritical
    End If
  ElseIf (newSearch) Then
    txtItemCodeSearch = ""
    newSearch = False
  End If
End Sub

Private Sub txtOrderQty_Change()
  lblTotalCost = Format(Val(lblUnitPrice) * Val(txtOrderQty), Constants.CURRENCY_FORMAT)
End Sub

Private Sub txtOrderQty_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtOrderQty) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub