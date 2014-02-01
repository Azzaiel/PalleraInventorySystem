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
      TabIndex        =   20
      Top             =   3240
      Width           =   1095
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
      TabIndex        =   18
      Top             =   3240
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
      TabIndex        =   17
      Top             =   3240
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
      TabIndex        =   16
      Top             =   3240
      Width           =   975
   End
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
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Order Items (Double cllick to view Detail)"
      Height          =   3975
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   6015
      Begin MSDataGridLib.DataGrid dgAccounts 
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
      Begin VB.Label Label2 
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
         Left            =   4440
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order Info"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Receive Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3840
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
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
         Format          =   106496003
         CurrentDate     =   41671
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Status"
         Height          =   255
         Left            =   480
         TabIndex        =   23
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
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7575
      Left            =   6360
      TabIndex        =   19
      Top             =   240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   13361
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
