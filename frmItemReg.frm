VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItemReg 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   19380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   19380
   Begin VB.CommandButton cmdActivation 
      Caption         =   "De-Activate"
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
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Type Form"
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtRetailPrice 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Active:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label txtActive 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Unit Price"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Retail Price"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Created by:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Created date:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Last mod by:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "Last mod date:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   975
      Left            =   6240
      TabIndex        =   12
      Top             =   120
      Width           =   12975
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   3120
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         Height          =   315
         Left            =   6000
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtSearchItemType 
         Height          =   285
         Left            =   6720
         TabIndex        =   14
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtSearchSuppliers 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FF00&
         Caption         =   "Supplier Name"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmbClear 
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
      Left            =   3840
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
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
      Left            =   5040
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmbEdit 
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
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmbNewRec 
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
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgItems 
      Height          =   3975
      Left            =   6240
      TabIndex        =   11
      Top             =   1080
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7011
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
End
Attribute VB_Name = "frmItemReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private suplierIdList As Variant
Private itemTypeIdList As Variant
Private tempRs As ADODB.Recordset

Private Sub cmbClear_Click()
Call clearForm
Call toogelInsertMode(False)
End Sub

Private Sub cmbClose_Click()
    Unload Me
End Sub

Private Sub cmbEdit_Click()

  Set tempRs = DataCrudDao.getItemRSByID(rs!id)
  tempRs!ITeM_CODE = txtItemCode
  tempRs!SUPPLIER_ID = suplierIdList(cmbSupplier.ListIndex)
  tempRs!ITEM_TYPE_ID = itemTypeIdList(cmbItemType.ListIndex)
  tempRs!Name = txtItemName
  tempRs!RETAIL_PRICE = txtRetailPrice
  tempRs!UNIT_PRICE = txtUnitPrice
  tempRs!CREATED_BY = UserSession.getLoginUser
  tempRs!LAST_MOD_DATE = Now
  
  tempRs.Update
  Call DbInstance.closeRecordSet(tempRs)
  MsgBox "Record Updated!! ", vbInformation
  Call populateDataGrid

End Sub

Private Sub cmbNewRec_Click()
  If (cmbNewRec.Caption = "New") Then
     toogelInsertMode (True)
  Else
    Call toogelInsertMode(False)
    Set tempRs = DataCrudDao.getFakeItemsRS
    tempRs.AddNew
    tempRs!ITeM_CODE = txtItemCode
    tempRs!SUPPLIER_ID = suplierIdList(cmbSupplier.ListIndex)
    tempRs!ITEM_TYPE_ID = itemTypeIdList(cmbItemType.ListIndex)
    tempRs!Name = txtItemName
    tempRs!RETAIL_PRICE = txtRetailPrice
    tempRs!UNIT_PRICE = txtUnitPrice
    tempRs!CREATED_BY = UserSession.getLoginUser
    tempRs!CREATED_DATE = Now
    tempRs!LAST_MOD_DATE = Now
    tempRs!active = txtActive
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Created", vbInformation
    Call populateDataGrid
    Call toogelInsertMode(False)
    cmbNewRec.Caption = "Add"
  
  
  End If
End Sub

Private Sub cmbSupplier_Click()
If cmbSupplier.ListIndex >= 0 Then
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
End If
End Sub

Private Sub cmdActivation_Click()

  Set tempRs = DataCrudDao.getItemRSByID(rs!id)
    
  If rs!active = "N" Then
    cmdActivation.Caption = "De-Activate"
    tempRs!active = "Y"
  Else
    tempRs!active = "N"
    cmdActivation.Caption = "Activate"
  End If
  tempRs.Update
  Call DbInstance.closeRecordSet(tempRs)
  MsgBox "Status Update "
    
  Call clearForm
  Call populateDataGrid


End Sub

Private Sub dgItems_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub

Private Sub Form_Load()
  Call populateLov
  Call populateDataGrid
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
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getItemReg()
  Set dgItems.DataSource = rs
  dgItems.Refresh
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
 End If
Call formatDataGrid
End Sub

Private Sub txtRetailPrice_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtRetailPrice) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub
Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtUnitPrice) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub showSelectedData()
 txtItemCode = CommonHelper.extractStringValue(rs!ITeM_CODE)
 cmbSupplier.Text = CommonHelper.extractStringValue(rs!SUPPLIER)
 cmbItemType.Text = CommonHelper.extractStringValue(rs!Item_Type)
 txtItemName = CommonHelper.extractStringValue(rs!ITEM_NAME)
 txtRetailPrice = CommonHelper.extractStringValue(rs!RETAIL_PRICE)
 txtUnitPrice = CommonHelper.extractStringValue(rs!UNIT_PRICE)
 txtActive = CommonHelper.extractStringValue(rs!active)
 lblCreatedBy = CommonHelper.extractStringValue(rs!CREATED_BY)
 lblCreatedDate = CommonHelper.extractDateValue(rs!CREATED_DATE)
 lblLatModBy = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
 lblLastModDate = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)

End Sub

Private Sub formatDataGrid()

End Sub

Private Sub clearForm()

txtItemCode = ""
cmbSupplier.ListIndex = -1
cmbItemType.ListIndex = -1
txtItemName = ""
txtRetailPrice = ""
txtUnitPrice = ""
lblCreatedBy = ""
lblCreatedDate = ""
lblLastModDate = ""
lblLatModBy = ""


End Sub

Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmbNewRec.Caption = "ADD"
    cmdActivation.Enabled = False
    cmbEdit.Enabled = False
  Else
    cmbNewRec.Caption = "New"
    cmdActivation.Enabled = True
    cmbEdit.Enabled = True
  End If
End Sub
