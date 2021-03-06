VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   Caption         =   "Order Form"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   15945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   6360
      TabIndex        =   27
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbSearchStatus 
         Height          =   315
         ItemData        =   "�.frx":0000
         Left            =   1200
         List            =   "�.frx":000A
         TabIndex        =   28
         Text            =   "Pending"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Status Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3615
         Left            =   0
         Picture         =   "�.frx":0022
         Stretch         =   -1  'True
         Top             =   -1800
         Width           =   9495
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   2040
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   3120
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "New"
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
      Left            =   960
      TabIndex        =   19
      Top             =   120
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
      Left            =   4200
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frmOrderItem 
      Caption         =   "Order Items (Double cllick to view Detail)"
      Height          =   3975
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   6015
      Begin MSDataGridLib.DataGrid dgOrderItems 
         Height          =   3255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Items (Double Click to View Details)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label lblTotalCost 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost:"
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
         Left            =   1560
         TabIndex        =   25
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblAddItemLink 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "�.frx":750F6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order Info"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton cmbReceiveOrder 
         Caption         =   "Receive Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   23
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Pending"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbSupplier 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtOrderDate 
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Format          =   107937795
         CurrentDate     =   41671
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   3720
         TabIndex        =   24
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Order ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblOrderID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   1935
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Order  Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Order  By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Received Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Received By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblReceviedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblOrderBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblReceviedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "�.frx":EA1CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8055
      End
   End
   Begin MSDataGridLib.DataGrid dgOrders 
      Height          =   6615
      Left            =   6360
      TabIndex        =   15
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11668
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   8175
      Left            =   0
      Picture         =   "�.frx":15F29E
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   19335
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private rsOrderItems As ADODB.Recordset
Private rsTemp As ADODB.Recordset
Private suplierIdList As Variant
Private itemTypeIdList As Variant
Private itemsList As Variant
Private tempRs As ADODB.Recordset
Private Sub lblCreatedDate_Click()

End Sub

Private Sub cmbItems_Click()
  'lblUnitPrice = itemsList(cmbItems.ListIndex, 1)
  'Call computeTotalPrice
End Sub

Private Sub cmbItemType_Click()
  'cmbItems.Clear
  'lblUnitPrice = ""
  'Call computeTotalPrice
  'Set tempRs = DataCrudDao.getItemByItemType(Val(itemTypeIdList(cmbItemType.ListIndex)))
  'ReDim itemsList(0 To tempRs.RecordCount, 0 To 1) As Long
  'Dim index As Integer
  'index = 0
   'While Not tempRs.EOF
   ' cmbItems.AddItem tempRs!ITEM_NAME
   ' itemsList(index, 0) = tempRs!id
   ' itemsList(index, 1) = tempRs!RETAIL_PRICE
   ' index = index + 1
   ' tempRs.MoveNext
  'Wend
  'Call DbInstance.closeRecordSet(tempRs)
End Sub

Private Sub cmbClose_Click()
  Unload Me
End Sub

Private Sub cmbReceiveOrder_Click()
  frmOrderReceive.lblOrderID = lblOrderID
  frmOrderReceive.lblSuplier = cmbSupplier.Text
  frmOrderReceive.lblStatus = txtStatus
  frmOrderReceive.lblOrderBy = lblOrderBy
  frmOrderReceive.lblOrderDate = dtOrderDate.value
  Set frmOrderReceive.rs = rsOrderItems
  Set frmOrderReceive.dgOrderItems.DataSource = frmOrderReceive.rs
  With frmOrderReceive.dgOrderItems
   .Columns(0).Visible = False
   .Columns(1).Visible = False
   
   .Columns(4).NumberFormat = Constants.CURRENCY_FORMAT
   .Columns(5).NumberFormat = Constants.CURRENCY_FORMAT
   .Columns(6).NumberFormat = Constants.CURRENCY_FORMAT
  End With
  Dim totalCost As Long
  If (rsOrderItems.RecordCount > 0) Then
    While Not rsOrderItems.EOF
      totalCost = totalCost + Val(rsOrderItems!TOTAL_PRICE)
      rsOrderItems.MoveNext
    Wend
    rsOrderItems.MoveFirst
    frmOrderReceive.lblTotalCost = Format(totalCost, Constants.CURRENCY_FORMAT)
  Else
    frmOrderReceive.lblTotalCost = 0
  End If
  frmOrderReceive.isfromMain = False
  frmOrderReceive.Show vbModal
End Sub
Private Sub cmdAdd_Click()
  If cmdAdd.Caption = "New" Then
    Call toogelInsertMode(True)
  Else
    Set rsTemp = DataCrudDao.getFakeOrdersRs
    rsTemp.AddNew
    rsTemp!status = txtStatus
    rsTemp!suplier_id = suplierIdList(cmbSupplier.ListIndex)
    rsTemp!Order_Date = dtOrderDate.value
    rsTemp!Ordered_By = UserSession.getLoginUser
    rsTemp.Update
    Call DbInstance.closeRecordSet(rsTemp)
    MsgBox "Record Added", vbInformation
    Call clearForm
    Call populateDataGrid
  End If
End Sub
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmdAdd.Caption = "ADD"
    txtStatus = "Pending"
    cmdDelete.Enabled = False
    frmOrderItem.Enabled = False
    cmbSupplier.Enabled = True
    cmbReceiveOrder.Enabled = False
    lblAddItemLink.ForeColor = vbGrayText
    Set rsOrderItems = DataCrudDao.getOrderItemsByOrderID(Val(0))
    Set dgOrderItems.DataSource = rsOrderItems
  
  Else
    cmdAdd.Caption = "New"
    cmdDelete.Enabled = True
    frmOrderItem.Enabled = True
    lblAddItemLink.ForeColor = vbBlue
    cmbSupplier.Enabled = False
  End If
End Sub
Private Sub clearForm()
  Call toogelInsertMode(False)
  dtOrderDate.CustomFormat = Constants.DEFAULT_FORMAT
  cmbSupplier.ListIndex = -1
  lblOrderBy = UserSession.getLoginUser
End Sub

Private Sub cmdclear_Click()
  Call clearForm
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim ans
  ans = MsgBox("Are you sure you want to Delete the order?", vbYesNo)
  If ans = vbYes Then
    Set rsTemp = DataCrudDao.getOrderByIDRs(lblOrderID)
    rsTemp.Delete
    Call DbInstance.closeRecordSet(rsTemp)
    MsgBox "Record Added", vbInformation
    Call populateDataGrid
  End If
End Sub

Private Sub Command1_Click()
  Call populateDataGrid
End Sub

Private Sub Command2_Click()
  cmbSearchStatus.Text = ""
End Sub

Private Sub dgOrderItems_DblClick()
  If (rsOrderItems.RecordCount > 0) Then
    frmAddOrderItem.orderItemID = rsOrderItems!id
    frmAddOrderItem.lblOrderID = lblOrderID
    frmAddOrderItem.lblSuplier = cmbSupplier.Text
    frmAddOrderItem.suplierID = suplierIdList(cmbSupplier.ListIndex)
    Call frmAddOrderItem.initializeForm
    frmAddOrderItem.cmbItemType.Text = rsOrderItems!ITEM_TYPE
    frmAddOrderItem.cmbItems.Text = rsOrderItems!Name
    frmAddOrderItem.txtRetailPrice = rsOrderItems!retil_price
    frmAddOrderItem.txtQuantity = rsOrderItems!QUANTITY
    frmAddOrderItem.cmdAdd.Caption = "Save"
    frmAddOrderItem.Show vbModal
  End If
End Sub

Private Sub dgOrders_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub

Public Sub Form_Load()
  dtOrderDate.CustomFormat = Constants.DEFAULT_FORMAT
  dtOrderDate = Now
  Call populateLov
  Call populateDataGrid
End Sub
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getOrders(cmbSearchStatus.Text)
  Set dgOrders.DataSource = rs
  If (rs.RecordCount > 0) Then
   rs.MoveFirst
   Call showSelectedData
  End If
End Sub
Public Sub showSelectedData()
  lblOrderID = CommonHelper.extractStringValue(rs!ORDER_ID)
  cmbSupplier.Text = rs!Suplier_Name
  txtStatus = CommonHelper.extractStringValue(rs!status)
  dtOrderDate.value = CommonHelper.extractDateValue(rs!Order_Date)
  lblOrderBy = CommonHelper.extractStringValue(rs!Ordered_By)
  lblReceviedDate = CommonHelper.extractDateValue(rs!RECIVED_DATE)
  lblReceviedBy = CommonHelper.extractStringValue(rs!RECIVED_BY)
  
  Set rsOrderItems = DataCrudDao.getOrderItemsByOrderID(Val(lblOrderID))
  Set dgOrderItems.DataSource = rsOrderItems
  
  If (rsOrderItems.RecordCount > 0) Then
    With dgOrderItems
      .Columns(0).Visible = False
      .Columns(1).Visible = False
    End With
    Dim totalCost As Long
    While Not rsOrderItems.EOF
      totalCost = totalCost + Val(rsOrderItems!TOTAL_PRICE)
      rsOrderItems.MoveNext
    Wend
    lblTotalCost = Format(totalCost, Constants.CURRENCY_FORMAT)
    rsOrderItems.MoveFirst
  Else
    lblTotalCost = 0
  End If
  
  
  If (txtStatus = "Pending") Then
    cmbReceiveOrder.Enabled = True
    lblAddItemLink.Enabled = True
    cmdDelete.Enabled = True
  Else
    cmbReceiveOrder.Enabled = False
    lblAddItemLink.Enabled = False
    cmdDelete.Enabled = False
  End If
  
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
  'If (Val(txtQuantity) > 0 And Val(lblUnitPrice) > 0) Then
  '  lblTotalPrice = Val(txtQuantity) * Val(lblUnitPrice)
  'Else
  '  lblTotalPrice = ""
  'End If
End Sub
Private Sub txtQuantity_Change()
  Call computeTotalPrice
End Sub
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    'If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtQuantity) > 11)) Then
    'KeyAscii = 0
    'Beep
  'End If
End Sub

Private Sub gdOrderItems_Click()

End Sub

Private Sub lblAddItemLink_Click()
  If (Val(lblOrderID) > 0) Then
    frmAddOrderItem.cmdAdd.Caption = "Add"
    frmAddOrderItem.lblOrderID = lblOrderID
    frmAddOrderItem.lblSuplier = cmbSupplier.Text
    frmAddOrderItem.suplierID = suplierIdList(cmbSupplier.ListIndex)
    Call frmAddOrderItem.initializeForm
    frmAddOrderItem.Show vbModal
  Else
    MsgBox "Please select a valid Order Item"
  End If
End Sub
