VERSION 5.00
Begin VB.Form frmAddOrderItem 
   Appearance      =   0  'Flat
   Caption         =   "Add Order Item"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   Picture         =   "frmAddOrderItem.frx":0000
   ScaleHeight     =   4245
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      Begin VB.Label lblSuplier 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XXXXXXXXXXXXXXXXXXX"
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
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblOrderID 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Suplier:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order ID:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   3975
         Left            =   0
         Picture         =   "frmAddOrderItem.frx":750D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4815
      Begin VB.TextBox txtRetailPrice 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbItems 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Retail Price"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Price"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblTotalPrice 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Image Image5 
         Height          =   4335
         Left            =   0
         Picture         =   "frmAddOrderItem.frx":7AAC1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15975
      End
   End
End
Attribute VB_Name = "frmAddOrderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public suplierID As Integer
Public orderItemID As Integer
Private itemTypeIdList As Variant
Private itemsInfoList As Variant
Private tempRs As ADODB.Recordset
Const ID_INDEX As Integer = 0
Const PRICE_INDEX As Integer = 1
Public Sub initializeForm()
  If suplierID >= 0 Then
    cmbItemType.Clear
    Set tempRs = DataCrudDao.getItemTypeRSBySupplierID(Val(suplierID))
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
  End If
End Sub

Private Sub cmbClose_Click()
  Unload Me
End Sub

Private Sub cmbItems_Click()
  txtRetailPrice = 0
  If (cmbItems.Text <> vbNullString) Then
    txtRetailPrice = itemsInfoList(cmbItems.ListIndex, PRICE_INDEX)
  End If
   Call computeTotalPrice
End Sub
Private Sub cmbItemType_Click()
  cmbItems.Clear
  txtRetailPrice = ""
  If cmbItemType.ListIndex > -1 Then
    
    Set tempRs = DataCrudDao.getItemByItemsRS(Val(itemTypeIdList(cmbItemType.ListIndex)))
    ReDim itemsInfoList(0 To tempRs.RecordCount, 0 To 1) As Long
    Dim index As Integer
    index = 0
    While Not tempRs.EOF
      cmbItems.AddItem tempRs!ITEM_NAME
      itemsInfoList(index, ID_INDEX) = tempRs!id
      itemsInfoList(index, PRICE_INDEX) = tempRs!RETAIL_PRICE
      index = index + 1
      tempRs.MoveNext
    Wend
    Call DbInstance.closeRecordSet(tempRs)
    
  End If
  Call computeTotalPrice
End Sub
Private Function hasValidForm() As Boolean
  Dim isValid As Boolean
  isValid = True
  If Val(txtQuantity) <= 0 Then
    isValid = False
    MsgBox "Please enter a valid Quantity to Continue", vbCritical
  End If
  hasValidForm = isValid
End Function
Private Sub cmdAdd_Click()
  If (hasValidForm = False) Then
    Exit Sub
  End If
  If (cmdAdd.Caption = "Add") Then
    Set tempRs = DataCrudDao.getFakeOrderItems
    tempRs.AddNew
    tempRs!ORDER_ID = lblOrderID
    tempRs!supplier_id = suplierID
    tempRs!ITEM_TYPE_ID = Val(itemTypeIdList(cmbItemType.ListIndex))
    tempRs!item_id = itemsInfoList(cmbItems.ListIndex, ID_INDEX)
    tempRs!retil_price = Val(txtRetailPrice)
    tempRs!QUANTITY = Val(txtQuantity)
    tempRs!CREATED_BY = UserSession.getLoginUser
    tempRs!CREATED_DATE = Now
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    MsgBox "Record Added", vbInformation
    Unload Me
    Call DbInstance.closeRecordSet(tempRs)
  Else
    Set tempRs = DataCrudDao.getOrderItemsByID(orderItemID)
    If (tempRs.RecordCount > 0) Then
      tempRs!ORDER_ID = lblOrderID
      tempRs!supplier_id = suplierID
      tempRs!ITEM_TYPE_ID = Val(itemTypeIdList(cmbItemType.ListIndex))
      tempRs!item_id = itemsInfoList(cmbItems.ListIndex, ID_INDEX)
      tempRs!retil_price = Val(txtRetailPrice)
      tempRs!QUANTITY = Val(txtQuantity)
      tempRs!CREATED_BY = UserSession.getLoginUser
      tempRs!CREATED_DATE = Now
      tempRs!LAST_MOD_BY = UserSession.getLoginUser
      tempRs!LAST_MOD_DATE = Now
      tempRs.Update
    End If
    MsgBox "Record Updated!!", vbInformation
    Unload Me
    Call DbInstance.closeRecordSet(tempRs)
  End If
End Sub

Private Sub Form_Load()
  suplierID = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmOrder.showSelectedData
End Sub

Private Sub txtQuantity_Change()
   Call computeTotalPrice
End Sub
Private Sub computeTotalPrice()
  lblTotalPrice = Format(Val(txtRetailPrice) * Val(txtQuantity), Constants.CURRENCY_FORMAT)
End Sub
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtQuantity) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txtRetailPrice_Change()
  Call computeTotalPrice
End Sub
