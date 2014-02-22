VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItemType 
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16245
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   16245
   StartUpPosition =   3  'Windows Default
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
      Left            =   240
      TabIndex        =   24
      Top             =   3360
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
      Left            =   1440
      TabIndex        =   23
      Top             =   3360
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
      Left            =   3960
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
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
      Left            =   2760
      TabIndex        =   21
      Top             =   3360
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgItemType 
      Height          =   2895
      Left            =   5280
      TabIndex        =   18
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   5106
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
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   975
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   10935
      Begin VB.TextBox txtSearchSuppliers 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtSearchItemType 
         Height          =   285
         Left            =   6720
         TabIndex        =   14
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         Height          =   315
         Left            =   6000
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   1695
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
      Begin VB.Label Label14 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type"
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Type Form"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.TextBox txtItemType 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox cmSuppliers 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "Last mod date:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Last mod by:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Created date:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Created by:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmItemType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private suplierIdList As Variant
Private tempRs As ADODB.Recordset
Private Sub dgCategories_Click()

End Sub
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getItemTypeRS(txtSearchItemType, txtSearchSuppliers)
  Set dgItemType.DataSource = rs
  dgItemType.Refresh
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  End If
  Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  
End Sub

Private Sub clearForm()
  txtItemType = ""
  cmSuppliers.ListIndex = -1
  lblCreatedBy = ""
  lblCreatedDate = ""
  lblLatModBy = ""
  lblLastModDate = ""
End Sub

Private Sub cmbClear_Click()
  Call clearForm
  Call toogelInsertMode(False)
End Sub

Private Sub cmbClose_Click()
  Unload Me
End Sub

Private Sub cmbDelete_Click()
End Sub

Private Sub cmbEdit_Click()

  Set tempRs = DataCrudDao.getItemTypeRSByID(rs!id)
  tempRs!SUPPLIER_ID = suplierIdList(cmSuppliers.ListIndex)
  tempRs!Name = txtItemType
  tempRs!LAST_MOD_BY = UserSession.getLoginUser
  tempRs!LAST_MOD_DATE = Now
  tempRs.Update
  Call DbInstance.closeRecordSet(tempRs)
  MsgBox "Record Updated", vbInformation
  Call populateDataGrid
End Sub

Private Sub cmbNewRec_Click()
  If (cmbNewRec.Caption = "New") Then
    Call toogelInsertMode(True)
  Else
    Set tempRs = DataCrudDao.getFakeItemTypeRS
    tempRs.AddNew
    tempRs!SUPPLIER_ID = suplierIdList(cmSuppliers.ListIndex)
    tempRs!Name = txtItemType
    tempRs!CREATED_BY = UserSession.getLoginUser
    tempRs!CREATED_DATE = Now
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Created", vbInformation
    Call populateDataGrid
    Call toogelInsertMode(False)
  End If
End Sub
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmbNewRec.Caption = "Add"
    cmbEdit.Enabled = False
  Else
    cmbNewRec.Caption = "New"
    cmbEdit.Enabled = True
  End If
End Sub
Private Sub cmdClearSearch_Click()
  txtSearchItemType = ""
  txtSearchSuppliers = ""
  cmSearchCategory.ListIndex = -1
End Sub

Private Sub cmdSearch_Click()
  Call populateDataGrid
End Sub

Private Sub cmSearchActive_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
    Call populateDataGrid
  End If
End Sub

Private Sub dgItemType_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub
Private Sub showSelectedData()
  cmSuppliers.Text = CommonHelper.extractStringValue(rs!Supplier_name)
  txtItemType = CommonHelper.extractStringValue(rs!ITEM_TYPE_NAME)
  lblCreatedBy = CommonHelper.extractStringValue(rs!CREATED_BY)
  lblCreatedDate = CommonHelper.extractDateValue(rs!CREATED_DATE)
  lblLatModBy = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
  lblLastModDate = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
    
End Sub

Private Sub Form_Load()
  Call populateLov
  Call populateDataGrid
End Sub
Private Sub populateLov()
  Set tempRs = DataCrudDao.getSupplierRS("", "", "")
  cmSuppliers.Clear
  ReDim suplierIdList(0 To tempRs.RecordCount) As Long
  Dim index As Integer
  index = 0
  While Not tempRs.EOF
    cmSuppliers.AddItem tempRs!Name
    suplierIdList(index) = tempRs!id
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub

Private Sub txtSearchSuppliers_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call populateDataGrid
  End If
End Sub

Private Sub txtSearchItemType_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
    Call populateDataGrid
  End If
End Sub

