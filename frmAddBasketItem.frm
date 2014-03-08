VERSION 5.00
Begin VB.Form frmAddBasketItem 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add Item"
   ClientHeight    =   5310
   ClientLeft      =   2790
   ClientTop       =   3180
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   Picture         =   "frmAddBasketItem.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   4980
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   4695
      Begin VB.TextBox txtOrderQty 
         Height          =   285
         Left            =   2160
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
         Left            =   2160
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   4215
         Left            =   0
         Picture         =   "frmAddBasketItem.frx":750D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15975
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
      Left            =   1200
      TabIndex        =   2
      Top             =   4560
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
      Left            =   2760
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4695
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblActive 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         BorderStyle     =   1  'Fixed Single
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
         BorderStyle     =   1  'Fixed Single
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
         BorderStyle     =   1  'Fixed Single
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
         BorderStyle     =   1  'Fixed Single
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
         BorderStyle     =   1  'Fixed Single
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
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Suppliers:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Stocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   4215
         Left            =   0
         Picture         =   "frmAddBasketItem.frx":7AAC1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15975
      End
   End
   Begin VB.TextBox txtItemCodeSearch 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label aaa 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code Search"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
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
Private tempRs As ADODB.Recordset
Private Sub cmbAddItem_Click()

  If rs.RecordCount = 0 Then
    MsgBox "Please select an Item First", vbCritical
    txtOrderQty.SetFocus
    Exit Sub
  ElseIf Val(txtOrderQty) = 0 Then
    MsgBox "Please enter a Quantity", vbCritical
    txtOrderQty.SetFocus
    Exit Sub
  ElseIf Val(lblStocks) - Val(txtOrderQty) < 0 Then
    MsgBox "Requested Quantity is greater that the current stock", vbCritical
    txtOrderQty.SetFocus
    txtOrderQty.SelStart = 0
    txtOrderQty.SelLength = Len(txtOrderQty)
    Exit Sub
  End If
  
  Set tempRs = DataCrudDao.getTmpBasketItem(UserSession.getLoginUser, Val(rs!supplier_id), Val(rs!id))
  If tempRs.RecordCount = 0 Then
    tempRs.AddNew
    tempRs!QUANTITY = Val(txtOrderQty)
  Else
    tempRs.MoveFirst
    tempRs!QUANTITY = Val(txtOrderQty) + Val(CommonHelper.extractStringValue(tempRs!QUANTITY))
    If Val(lblStocks) - Val(tempRs!QUANTITY) < 0 Then
      MsgBox "Requested Quantity is greater that the current stock", vbCritical
      txtOrderQty.SetFocus
      txtOrderQty.SelStart = 0
      txtOrderQty.SelLength = Len(txtOrderQty)
      Exit Sub
    End If
  End If
  tempRs!username = UserSession.getLoginUser
  tempRs!supplier_id = rs!supplier_id
  tempRs!item_id = rs!id
  tempRs!UNIT_PRICE = Val(lblUnitPrice)
  tempRs.Update
  Call DbInstance.closeRecordSet(tempRs)
  MsgBox "Item Added to Basket", vbInformation
  Call frmItemSell.reloadBasketItems
  Call clearForm
  txtItemCodeSearch = ""
  newSearch = False
  txtItemCodeSearch.SetFocus
End Sub
Private Sub clearForm()
   lblSuplier = ""
   lblItemType = ""
   lblItem = ""
   lblStocks = ""
   lblUnitPrice = ""
   lblActive = ""
   txtOrderQty = ""
   lblTotalCost = ""
End Sub
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
      lblStocks = Val(CommonHelper.extractStringValue(rs!QUANTITY))
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
  If KeyAscii = 13 Then
    Call cmbAddItem_Click
  ElseIf (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtOrderQty) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub
