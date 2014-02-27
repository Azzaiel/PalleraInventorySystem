VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItemSell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Counter"
   ClientHeight    =   5415
   ClientLeft      =   5385
   ClientTop       =   3105
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   12705
   Begin VB.Frame Frame2 
      Caption         =   "Commands"
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add Item"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
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
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmbClear 
         Caption         =   "Clear"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton cmbReceiveOrder 
         Caption         =   "Check Out"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Items"
      Height          =   5175
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin MSDataGridLib.DataGrid dgBasket 
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7435
         _Version        =   393216
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
         Left            =   4920
         TabIndex        =   3
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
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
         Left            =   3600
         TabIndex        =   2
         Top             =   4680
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmItemSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tmpRs As ADODB.Recordset
Private tmpBasketRs As ADODB.Recordset
Private salesRS As ADODB.Recordset
Public totalCost As Double
Public payment As Long
Private Sub cmbClear_Click()
  Dim ans
  ans = MsgBox("Are you sure you want to clear the basket?", vbYesNo)
  If ans = vbYes Then
    Call DataCrudDao.deleteTmpUserBasket(UserSession.getLoginUser)
    MsgBox "Basket Cleared", vbInformation
    Call reloadBasketItems
  End If
End Sub
Private Sub cmbReceiveOrder_Click()
  If Val(totalCost) = 0 Then
    MsgBox "Please purchase an Item First", vbCritical
    Exit Sub
  End If
  
  payment = -1
  frmEntePayment.lblTotalCost = lblTotalCost
  frmEntePayment.Show vbModal
  If payment <> -1 Then
    Set RepSalesInvoice.DataSource = rs
    RepSalesInvoice.Sections(2).Controls("lblDate").Caption = Format(Now, Constants.DEFAULT_FORMAT)
    RepSalesInvoice.Sections(5).Controls("lblTotalCost").Caption = Format(totalCost, Constants.CURRENCY_FORMAT)
    RepSalesInvoice.Sections(5).Controls("lblTendred").Caption = Format(payment, Constants.CURRENCY_FORMAT)
    RepSalesInvoice.Sections(5).Controls("lblChange").Caption = Format(payment - totalCost, Constants.CURRENCY_FORMAT)
    
    Set tmpBasketRs = DataCrudDao.getUserTmpBasket(UserSession.getLoginUser)
    Set salesRS = DataCrudDao.getFakeSalesRs
    
    While Not tmpBasketRs.EOF
      
      Set tmpRs = DataCrudDao.getItemRSByID(tmpBasketRs!item_id)
      tmpRs!quantity = Val(tmpRs!quantity) - Val(tmpBasketRs!quantity)
      tmpRs.Update
      Call DbInstance.closeRecordSet(tmpRs)
      
      salesRS.AddNew
      salesRS!username = UserSession.getLoginUser
      salesRS!supplier_id = tmpBasketRs!supplier_id
      salesRS!item_id = tmpBasketRs!item_id
      salesRS!sale_date = Now
      salesRS!quantity = tmpBasketRs!quantity
      salesRS!unit_price = tmpBasketRs!unit_price
      salesRS.Update
      
      tmpBasketRs.Delete
      tmpBasketRs.MoveNext
    Wend
    
    Call DbInstance.closeRecordSet(tmpBasketRs)
    Call DbInstance.closeRecordSet(salesRS)
    
    RepSalesInvoice.Show vbModal
    Call reloadBasketItems
  End If
End Sub
Private Sub cmdAddItem_Click()
  frmAddBasketItem.Show vbModal
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdRemoveItem_Click()
  If rs.RecordCount = 0 Or rs.BOF Then
    MsgBox "Please select a record to delete", vbCritical
    Exit Sub
  End If
  Dim ans
  ans = MsgBox("Are you sure you want to remove the item?", vbYesNo)
  If ans = vbYes Then
    Set tmpRs = DataCrudDao.getTmpBasketItem(UserSession.getLoginUser, Val(rs!supplier_id), Val(rs!item_id))
    If tmpRs.RecordCount > 0 Then
      tmpRs.Delete
      tmpRs.Update
      Call DbInstance.closeRecordSet(tmpRs)
      MsgBox "Item Removed from basket", vbInformation
      Call reloadBasketItems
    End If
  End If
End Sub

Private Sub Form_Load()
  Call reloadBasketItems
End Sub
Public Sub reloadBasketItems()
  Set rs = DataCrudDao.getBasketItemsByUser(UserSession.getLoginUser)
  Set dgBasket.DataSource = rs
  totalCost = 0
  If rs.RecordCount > 0 Then
    While Not rs.EOF
      totalCost = totalCost + Val(CommonHelper.extractStringValue(rs!Total_Cost))
      rs.MoveNext
    Wend
    rs.MoveFirst
    lblTotalCost = Format(totalCost, Constants.CURRENCY_FORMAT)
  Else
    lblTotalCost = 0
  End If
  With dgBasket
    .Columns(0).Width = 1000
    .Columns(1).Width = 2000
    .Columns(2).Width = 3500
    .Columns(3).Width = 800
    .Columns(3).NumberFormat = Constants.CURRENCY_FORMAT
    .Columns(4).Width = 800
    .Columns(4).Width = 1000
    .Columns(5).NumberFormat = Constants.CURRENCY_FORMAT
  End With
End Sub
