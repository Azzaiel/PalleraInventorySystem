VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Pallera's  Inventory System"
   ClientHeight    =   7755
   ClientLeft      =   420
   ClientTop       =   1635
   ClientWidth     =   18765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   18765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Items in Critical Level"
      Height          =   7095
      Left            =   8520
      TabIndex        =   3
      Top             =   480
      Width           =   9855
      Begin MSDataGridLib.DataGrid dgCriticalItems 
         Height          =   6735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   11880
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Items in Critical Level"
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
         TabIndex        =   6
         Top             =   0
         Width           =   4695
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pending Orders (Double Click to View Details)"
      Height          =   7095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      Begin MSDataGridLib.DataGrid dgPendinOrders 
         Height          =   6735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11880
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Orders ( Double Click to View Details)"
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
         TabIndex        =   5
         Top             =   0
         Width           =   4695
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "frmMain.frx":750D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   8175
      Left            =   0
      Picture         =   "frmMain.frx":EA1A8
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   19335
   End
   Begin VB.Menu mnSupplier 
      Caption         =   "Supplier"
      Begin VB.Menu mnReSupplier 
         Caption         =   "Regster Supplier"
      End
      Begin VB.Menu mnRegItemType 
         Caption         =   "Register Item Type"
      End
   End
   Begin VB.Menu mnInventory 
      Caption         =   "Inventory"
      Begin VB.Menu mnRegisterItem 
         Caption         =   "Item Registry"
      End
      Begin VB.Menu mnOder 
         Caption         =   "Order Itmes"
      End
      Begin VB.Menu MnItemLoss 
         Caption         =   "Item Loss"
      End
   End
   Begin VB.Menu mnSellItem 
      Caption         =   "Sell Item"
   End
   Begin VB.Menu mnReports 
      Caption         =   "Reports"
      Begin VB.Menu mnFastMoving 
         Caption         =   "Fast  Moving"
      End
      Begin VB.Menu mnOrderReport 
         Caption         =   "Orders"
      End
      Begin VB.Menu mnSalesReport 
         Caption         =   "Sales"
      End
   End
   Begin VB.Menu mnUsers 
      Caption         =   "Account"
      Begin VB.Menu mnChangePass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnRegUsers 
         Caption         =   "Regiter Users"
      End
      Begin VB.Menu mnLogout 
         Caption         =   "Logout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pendingOrderRs As ADODB.Recordset
Private itemsCriticalRS As ADODB.Recordset

Private Sub dgPendinOrders_DblClick()
  If pendingOrderRs.RecordCount > 0 Then
    frmOrderReceive.lblOrderID = pendingOrderRs!ORDER_ID
    frmOrderReceive.lblSuplier = pendingOrderRs!Suplier_Name
    frmOrderReceive.lblStatus = "Pending"
    frmOrderReceive.lblOrderBy = pendingOrderRs!Ordered_By
    frmOrderReceive.lblOrderDate = pendingOrderRs!Order_Date
    Set frmOrderReceive.rs = DataCrudDao.getOrderItemsByOrderID(Val(pendingOrderRs!ORDER_ID))
    Set frmOrderReceive.dgOrderItems.DataSource = frmOrderReceive.rs
    With frmOrderReceive.dgOrderItems
     .Columns(0).Visible = False
     .Columns(1).Visible = False
     
     .Columns(4).NumberFormat = Constants.CURRENCY_FORMAT
     .Columns(5).NumberFormat = Constants.CURRENCY_FORMAT
     .Columns(6).NumberFormat = Constants.CURRENCY_FORMAT
     
    End With
    
    frmOrderReceive.lblTotalCost = Format(pendingOrderRs!Total_Cost, Constants.CURRENCY_FORMAT)
    frmOrderReceive.isfromMain = True
    frmOrderReceive.Show vbModal
  End If
End Sub

Private Sub Form_Load()
  lblWelcome.Caption = "Welcome " & UserSession.Name & " you are logged in as " & UserSession.Role
  Call populatePendingOrderDash
  Call populateItemsInCritical
End Sub
Public Sub populateItemsInCritical()
  Set itemsCriticalRS = DataCrudDao.getCriticalLevelItemRS
  Set dgCriticalItems.DataSource = itemsCriticalRS
  With dgCriticalItems
    .Columns(0).Width = 1200
    .Columns(0).Alignment = dbgCenter
    
    .Columns(1).Width = 2000
    
    .Columns(2).Width = 1500
    
    .Columns(3).Width = 2000
    
    .Columns(4).Width = 1000
    .Columns(4).Alignment = dbgCenter
    
    .Columns(5).Width = 1500
    .Columns(5).Alignment = dbgCenter
    
  End With
End Sub
Public Sub populatePendingOrderDash()
  Set pendingOrderRs = DataCrudDao.getPendingOrderDash
  Set dgPendinOrders.DataSource = pendingOrderRs
  With dgPendinOrders
    .Columns(0).Width = 800
    .Columns(0).Alignment = dbgCenter
    
    .Columns(2).Width = 1200
    
    .Columns(3).Width = 1500
    .Columns(3).NumberFormat = Constants.DEFAULT_FORMAT
    
    .Columns(4).Width = 800
    .Columns(4).Alignment = dbgCenter
    
    .Columns(5).Width = 900
    .Columns(5).Alignment = dbgCenter
    .Columns(5).NumberFormat = Constants.CURRENCY_FORMAT
    
  End With
End Sub

Private Sub mnChangePass_Click()
  frmChangePassword.Show vbModal
End Sub

Private Sub mnFastMoving_Click()
  frmFastMovingItems.Show vbModal
End Sub

Private Sub MnItemLoss_Click()
  frmItemLoss.Show vbModal
End Sub

Private Sub mnLogout_Click()
   Unload Me
   LogIn.Show vbModal
End Sub

Private Sub mnOder_Click()
 frmOrder.Show vbModal
 Call populatePendingOrderDash
 Call populateItemsInCritical
End Sub

Private Sub mnOrderReport_Click()
  frmOrderReport.Show vbModal
End Sub

Private Sub mnRegisterItem_Click()
  frmItemReg.Show vbModal
  Call populateItemsInCritical
End Sub

Private Sub mnRegItemType_Click()
  frmItemType.Show vbModal
End Sub

Private Sub mnRegUsers_Click()
  frmAccnt.Show vbModal
End Sub

Private Sub mnReSupplier_Click()
  frmSupplier.Show vbModal
End Sub

Private Sub mnSalesReport_Click()
frmSalesReport.Show vbModal
End Sub

Private Sub mnSellItem_Click()
  frmItemSell.Show vbModal
  Call populateItemsInCritical
End Sub
