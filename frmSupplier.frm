VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSupplier 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   270
   ClientTop       =   765
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   18645
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
      TabIndex        =   31
      Top             =   5040
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
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
   End
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
      Left            =   2520
      TabIndex        =   29
      Top             =   5040
      Width           =   1215
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
      Left            =   1320
      TabIndex        =   28
      Top             =   5040
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
      TabIndex        =   27
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   975
      Left            =   6240
      TabIndex        =   17
      Top             =   120
      Width           =   12375
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   7440
         TabIndex        =   26
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         Height          =   315
         Left            =   9480
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtSearchSales 
         Height          =   285
         Left            =   7440
         TabIndex        =   24
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox cmSearchActive 
         Height          =   315
         ItemData        =   "frmSupplier.frx":0000
         Left            =   1800
         List            =   "frmSupplier.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtSearchName 
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000FF00&
         Caption         =   "Sales Contact"
         Height          =   255
         Left            =   6120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Active"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FF00&
         Caption         =   "Supplier Name"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sales Form"
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtSalesEmail 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtSalesPhone 
         Height          =   285
         Left            =   1800
         TabIndex        =   35
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtSales 
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtComPhone 
         Height          =   285
         Left            =   1800
         TabIndex        =   32
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCompanyAddress 
         Height          =   645
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Sales Email"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "*Sales Phone no"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblActive 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "Last mde date"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "Last mod by"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Created date"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Created by"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Active"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "*Sales Contact"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Company Phone no"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Company Address"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "*Supplier Name"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dgSupplier 
      Height          =   4335
      Left            =   6240
      TabIndex        =   0
      Top             =   1200
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      RowDividerStyle =   3
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
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private Sub dgCategories_Click()

End Sub
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getSupplierRS(cmSearchActive.Text, txtSearchName, txtSearchSales)
  Set dgSupplier.DataSource = rs
  dgSupplier.Refresh
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  End If
  Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  
End Sub

Private Sub cmbClear_Click()
  Call toogelInsertMode(False)
End Sub

Private Sub cmbClose_Click()
  Unload Me
End Sub

Private Sub cmbEdit_Click()
  Call resetFormSkin
  If (isFormDetailValid) Then
    rs!Name = txtName
    rs!active = lblActive
    rs!COMPANY_PHONE_NUMBER = txtComPhone
    rs!COMPANY_ADDRESS = txtCompanyAddress
    rs!SALES_CONTACT = txtSales
    rs!SALES_EMAIL = txtSalesEmail
    rs!SALES_PHONE_NUMBER = txtSalesPhone
    rs!CREATED_BY = UserSession.getLoginUser
    rs!CREATED_DATE = Now
    rs!LAST_MOD_BY = UserSession.getLoginUser
    rs!LAST_MOD_DATE = Now
    rs.Update
    MsgBox "Record Updated", vbInformation
    Call populateDataGrid
  End If
End Sub
Private Function resetFormSkin()
  Call CommonHelper.toDefaultSkin(txtName)
  Call CommonHelper.toDefaultSkin(txtSales)
  Call CommonHelper.toDefaultSkin(txtSalesPhone)
End Function

Private Sub cmbNewRec_Click()
  Call resetFormSkin
  If (cmbNewRec.Caption = "New") Then
    Call toogelInsertMode(True)
  Else
    If (isFormDetailValid) Then
      rs.AddNew
      rs!Name = txtName
      rs!active = lblActive
      rs!COMPANY_PHONE_NUMBER = txtComPhone
      rs!COMPANY_ADDRESS = txtCompanyAddress
      rs!SALES_CONTACT = txtSales
      rs!SALES_EMAIL = txtSalesEmail
      rs!SALES_PHONE_NUMBER = txtSalesPhone
      rs!LAST_MOD_BY = UserSession.getLoginUser
      rs!LAST_MOD_DATE = Now
      rs.Update
      MsgBox "Record Created", vbInformation
      Call populateDataGrid
      Call toogelInsertMode(False)
    End If
  End If
End Sub

Private Function isFormDetailValid() As Boolean
  If (Not CommonHelper.hasValidValue(txtName)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtName, "Please enter the Supplier Name")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(txtSales)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtSales, "Please sales contact")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(txtSalesPhone)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtSalesPhone, "Please sales phone number")
    Exit Function
  End If
  
  isFormDetailValid = True
End Function

Private Sub cmdActivation_Click()
  If (cmdActivation.Caption = "De-Activate") Then
    rs!active = "N"
    rs!LAST_MOD_BY = UserSession.getLoginUser
    rs!LAST_MOD_DATE = Now
    rs.Update
    MsgBox "Supplier De-Activated"
    cmdActivation.Caption = "Activate"
  Else
    rs!active = "Y"
    rs!LAST_MOD_BY = UserSession.getLoginUser
    rs!LAST_MOD_DATE = Now
    rs.Update
    MsgBox "Supplier Activated"
    cmdActivation.Caption = "Activate"
  End If
  Call populateDataGrid
End Sub

Private Sub cmdClearSearch_Click()
  txtSearchName = ""
  txtSearchSales = ""
  cmSearchActive.ListIndex = -1
End Sub

Private Sub cmdSearch_Click()
  Call populateDataGrid
End Sub

Private Sub cmSearchActive_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
    Call populateDataGrid
  End If
End Sub

Private Sub dgSupplier_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub
Private Sub showSelectedData()
  txtName = CommonHelper.extractStringValue(rs!Name)
  lblActive = CommonHelper.extractStringValue(rs!active)
  txtComPhone = CommonHelper.extractStringValue(rs!COMPANY_PHONE_NUMBER)
  txtCompanyAddress = CommonHelper.extractStringValue(rs!COMPANY_ADDRESS)
  txtSales = CommonHelper.extractStringValue(rs!SALES_CONTACT)
  txtSalesEmail = CommonHelper.extractStringValue(rs!SALES_EMAIL)
  txtSalesPhone = CommonHelper.extractStringValue(rs!SALES_PHONE_NUMBER)
  lblCreatedBy = CommonHelper.extractStringValue(rs!CREATED_BY)
  lblCreatedDate = CommonHelper.extractDateValue(rs!CREATED_DATE)
  lblLatModBy = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
  lblLastModDate = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
  
  If (lblActive = "Y") Then
    cmdActivation.Caption = "De-Activate"
  Else
    cmdActivation.Caption = "Activate"
  End If
  
End Sub
Private Sub clearForm()
  txtName = ""
  lblActive = "Y"
  txtComPhone = ""
  txtCompanyAddress = ""
  txtSales = ""
  txtSalesEmail = ""
  txtSalesPhone = ""
  lblCreatedBy = ""
  lblCreatedDate = ""
  lblLatModBy = ""
  lblLastModDate = ""
End Sub
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmbNewRec.Caption = "Add"
    cmdActivation.Enabled = False
    cmbClear.Enabled = False
  Else
    cmbNewRec.Caption = "New"
    cmdActivation.Enabled = True
    cmbClear.Enabled = True
  End If
End Sub

Private Sub Form_Load()
    Call populateDataGrid
End Sub

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call populateDataGrid
  End If
End Sub

Private Sub txtSearchSales_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
    Call populateDataGrid
  End If
End Sub
