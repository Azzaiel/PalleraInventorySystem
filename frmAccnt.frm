VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccnt 
   Caption         =   "Accounts"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   16005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   15
      Left            =   9000
      TabIndex        =   17
      Top             =   2160
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration Form"
      Height          =   3375
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtFirstname 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtMiddlename 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtLastname 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   2640
         Width           =   2295
      End
      Begin VB.ComboBox cmbRole 
         Height          =   315
         ItemData        =   "frmAccnt.frx":0000
         Left            =   2160
         List            =   "frmAccnt.frx":000A
         TabIndex        =   7
         Text            =   "cmbRole"
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "USERNAME"
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "FIRST NAME"
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "MIDDLE NAME"
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "LAST NAME"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "ROLE"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   3480
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
      Left            =   2280
      TabIndex        =   4
      Top             =   3480
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
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
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
      Left            =   3360
      TabIndex        =   2
      Top             =   3480
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
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgAccounts 
      Height          =   3855
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6800
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
Attribute VB_Name = "frmAccnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private Sub dgCategories_Click()
End Sub
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getAccount()
  Set dgAccounts.DataSource = rs
  dgAccounts.Refresh
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  End If
  Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  With dgAccounts
    .Columns(6).Visible = False
  End With
End Sub
Private Sub cmdAdd_Click()
  If (cmdAdd.Caption = "New") Then
    Call toogelInsertMode(True)
  Else
    rs.AddNew
    rs!Username = txtUsername.Text
    rs!Password = txtUsername.Text
    rs!Role = cmbRole.Text
    rs!FIRST_NAME = txtFirstname.Text
    rs!LAST_NAME = txtLastname.Text
    rs!MIDDLE_NAME = txtMiddlename.Text
    rs.Update
    MsgBox "Record Created, Default password was set", vbInformation
    Call populateDataGrid
    Call toogelInsertMode(False)
  End If
End Sub
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmdAdd.Caption = "Add"
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
  Else
    cmdAdd.Caption = "New"
    cmdDelete.Enabled = True
    cmdEdit.Enabled = True
  End If
End Sub

Private Sub cmdclear_Click()
  Call clearForm
  Call toogelInsertMode(False)
End Sub
Private Sub clearForm()
  txtID = ""
  txtUsername = ""
  txtFirstname = ""
  txtLastname = ""
  txtMiddlename = ""
  cmbRole.Text = ""
  cmbRole.ListIndex = -1
End Sub

Private Sub dgAccount_Click()

End Sub

Private Sub cmdDelete_Click()
    Dim response As String
    response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
    If (response = vbOK) Then
      rs.Delete
      MsgBox "Record Deleted", vbInformation
      Call populateDataGrid
    End If
End Sub

Private Sub cmdEdit_Click()
    rs!id = txtID
    rs!Username = txtUsername
    rs!Role = cmbRole.Text
    rs!FIRST_NAME = txtFirstname
    rs!LAST_NAME = txtLastname
    rs!MIDDLE_NAME = txtMiddlename
    rs.Update
    MsgBox "Record Updated", vbInformation
    Call clearForm
    Call populateDataGrid
End Sub

Private Sub dgGetAccount_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedData
End Sub

Private Sub dgGetAccount_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Call populateDataGrid
End Sub
Private Sub showSelectedData()
  'txtID = CommonHelper.extractStringValue(rs!ID)
  txtUsername = CommonHelper.extractStringValue(rs!Username)
  'txtPassword = CommonHelper.extractStringValue(rs!PASSWORD)
  cmbRole.Text = CommonHelper.extractStringValue(rs!Role)
  txtFirstname = CommonHelper.extractStringValue(rs!FIRST_NAME)
  txtLastname = CommonHelper.extractStringValue(rs!LAST_NAME)
  txtMiddlename = CommonHelper.extractStringValue(rs!MIDDLE_NAME)
  End Sub
  
Private Sub Label1_Click()

End Sub
