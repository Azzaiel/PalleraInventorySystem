VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   2280
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtOrderDate 
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   3360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Format          =   106496003
         CurrentDate     =   41671
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox cmbItems 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblTotalPrice 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblUnitPrice 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FF00&
         Caption         =   "Total Price"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblOrderID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Status"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Order ID"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Received By"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Received Date"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  By"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Retail Price"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  Date"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private suplierIdList As Variant
Private itemTypeIdList As Variant
Private itemsList As Variant
Private tempRs As ADODB.Recordset
Private Sub lblCreatedDate_Click()

End Sub

Private Sub cmbItems_Click()
  lblUnitPrice = itemsList(cmbItems.ListIndex, 1)
  Call computeTotalPrice
End Sub

Private Sub cmbItemType_Click()
  cmbItems.Clear
  lblUnitPrice = ""
  Call computeTotalPrice
  Set tempRs = DataCrudDao.getItemByItemType(Val(itemTypeIdList(cmbItemType.ListIndex)))
  ReDim itemsList(0 To tempRs.RecordCount, 0 To 1) As Long
  Dim index As Integer
  index = 0
   While Not tempRs.EOF
    cmbItems.AddItem tempRs!ITEM_NAME
    itemsList(index, 0) = tempRs!id
    itemsList(index, 1) = tempRs!retail_price
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub

Private Sub cmbSupplier_Click()
  cmbItemType.Clear
  Set tempRs = DataCrudDao.getItemTypeRSBySupplierID(Val(suplierIdList(cmbSupplier.ListIndex)))
  ReDim itemTypeIdList(0 To tempRs.RecordCount) As Long
  Dim index As Integer
  index = 0
   While Not tempRs.EOF
    cmbItemType.AddItem tempRs!ITEM_TYPE_NAME
    itemTypeIdList(index) = tempRs!id
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub

Private Sub Form_Load()
  dtOrderDate.CustomFormat = Constants.DEFAULT_FORMAT
  dtOrderDate = Now
  Call populateLov
End Sub
Private Sub populateLov()
  Set tempRs = DataCrudDao.getSupplierRS("", "", "")
  cmbSupplier.Clear
  ReDim suplierIdList(0 To tempRs.RecordCount) As Long
  Dim index As Integer
  index = 0
  While Not tempRs.EOF
    cmbSupplier.AddItem tempRs!Name
    suplierIdList(index) = tempRs!id
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

End Sub

Private Sub computeTotalPrice()
  If (Val(txtQuantity) > 0 And Val(lblUnitPrice) > 0) Then
    lblTotalPrice = Val(txtQuantity) * Val(lblUnitPrice)
  Else
    lblTotalPrice = ""
  End If
End Sub

Private Sub txtQuantity_Change()
  Call computeTotalPrice
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtQuantity) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub
