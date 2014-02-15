VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   15945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      TabIndex        =   22
      Top             =   120
      Width           =   975
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
      Left            =   1560
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
      Left            =   3720
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
      Left            =   480
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
      Left            =   4800
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
      Begin MSDataGridLib.DataGrid gdOrderItems 
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   6165
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
      Begin VB.Label lblAddItemLink 
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
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order Info"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Pending"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtOrderDate 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Format          =   107151363
         CurrentDate     =   41671
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Status"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Order ID"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblOrderID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  Date"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  By"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Received Date"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Received By"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblReceviedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblOrderBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblReceviedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid dgOrders 
      Height          =   6135
      Left            =   6360
      TabIndex        =   15
      Top             =   1680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
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
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private rsTemp As ADODB.Recordset
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

Private Sub cmbClose_Click()
  Unload Me
End Sub

Private Sub cmbSupplier_Click()
  'cmbItemType.Clear
  'Set tempRs = DataCrudDao.getItemTypeRSBySupplierID(Val(suplierIdList(cmbSupplier.ListIndex)))
  'ReDim itemTypeIdList(0 To tempRs.RecordCount) As Long
  'Dim index As Integer
  'index = 0
  ' While Not tempRs.EOF
  '  cmbItemType.AddItem tempRs!ITEM_TYPE_NAME
  '  itemTypeIdList(index) = tempRs!id
  '  index = index + 1
  '  tempRs.MoveNext
  'Wend
  'Call DbInstance.closeRecordSet(tempRs)
End Sub
Private Sub cmdAdd_Click()
  If cmdAdd.Caption = "New" Then
    Call toogelInsertMode(True)
  Else
    Set rsTemp = DataCrudDao.getFakeOrdersRs
    rsTemp.AddNew
    rsTemp!Status = txtStatus
    rsTemp!Suplier_id = suplierIdList(cmbSupplier.ListIndex)
    rsTemp!Order_Date = dtOrderDate.value
    rsTemp!Ordered_by = UserSession.getLoginUser
    rsTemp.Update
    Call DbInstance.closeRecordSet(rsTemp)
    MsgBox "Record Added", vbInformation
  End If
End Sub
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call ClearForm
    cmdAdd.Caption = "ADD"
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    frmOrderItem.Enabled = False
    lblAddItemLink.ForeColor = vbGrayText
  Else
    cmdAdd.Caption = "New"
    cmdDelete.Enabled = True
    cmdEdit.Enabled = True
    frmOrderItem.Enabled = True
    lblAddItemLink.ForeColor = vbBlue
  End If
End Sub
Private Sub ClearForm()
  Call toogelInsertMode(False)
End Sub

Private Sub cmdClear_Click()
  Call ClearForm
End Sub

Private Sub dgOrders_Click()

End Sub

Private Sub dgOrders_SelChange(Cancel As Integer)
  Call r
End Sub

Private Sub Form_Load()
  dtOrderDate.CustomFormat = Constants.DEFAULT_FORMAT
  dtOrderDate = Now
  Call populateLov
  Call populateDatagrid
End Sub
Private Sub populateDatagrid()
  Set rs = DataCrudDao.getPendingOrdersRs
  Set dgOrders.DataSource = rs
  If (rs.RecordCount > 0) Then
   rs.MoveFirst
   Call showSelectedData
  End If
  
End Sub
Private Sub showSelectedData()
  lblOrderID = CommonHelper.extractStringValue(rs!Order_ID)
  cmbSupplier.Text = rs!Suplier_Name
  txtStatus = CommonHelper.extractStringValue(rs!Status)
  dtOrderDate.value = CommonHelper.extractDateValue(rs!Order_Date)
  lblOrderBy = CommonHelper.extractStringValue(rs!Ordered_by)
  lblReceviedDate = CommonHelper.extractDateValue(rs!Recived_Date)
  lblReceviedBy = CommonHelper.extractStringValue(rs!RECIVED_BY)
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

Private Sub lblAddItemLink_Click()
  frmAddOrderItem.Show
End Sub
