VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Pallera's  Inventory System"
   ClientHeight    =   7755
   ClientLeft      =   420
   ClientTop       =   1635
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   14145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Pending Orders (Double Click to View Details)"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   240
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
   End
   Begin VB.Label lblWelcome 
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
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Width           =   5895
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
   End
   Begin VB.Menu mnSellItem 
      Caption         =   "Sell Item"
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
    
    frmOrderReceive.lblTotalCost = Format(pendingOrderRs!Total_cost, Constants.CURRENCY_FORMAT)
  
    frmOrderReceive.Show vbModal
  End If
End Sub

Private Sub Form_Load()
  lblWelcome.Caption = "Welcome " & UserSession.Name & " you are logged in as " & UserSession.Role
  Call populatePendingOrderDash
End Sub
Private Sub populatePendingOrderDash()
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

Private Sub mnLogout_Click()
   Unload Me
   LogIn.Show vbModal
End Sub

Private Sub mnOder_Click()
  frmOrder.Show vbModal
End Sub

Private Sub mnRegisterItem_Click()
  frmItemReg.Show vbModal
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

Private Sub mnSellItem_Click()
  frmItemSell.Show vbModal
End Sub
