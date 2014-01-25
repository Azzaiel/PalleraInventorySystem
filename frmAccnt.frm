VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccnt 
   Caption         =   "Accounts"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   16005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "exit"
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtRole 
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtLastname 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtMiddlename 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtFirstname 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dgGetAccount 
      Height          =   4455
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7858
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
   Begin VB.Label Label7 
      Caption         =   "ROLE"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "LAST NAME"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "MIDDLE NAME"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "FIRST NAME"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "PASSWORD"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "USERNAME"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmAccnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRS As ADODB.Recordset
Private Sub dgCategories_Click()

End Sub
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getAccount(txtID, txtUsername, txtPassword, txtRole, txtFirstname, txtLastname, txtMiddlename)
  Set dgGetAccount.DataSource = rs
  dgGetAccount.Refresh
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  End If
  
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdADD_Click()
  
  rs.AddNew
  rs!ID = txtID.Text
  rs!USERNAME = txtUsername.Text
  rs!PASSWORD = txtPassword.Text
  rs!ROLE = txtRole.Text
  rs!FIRST_NAME = txtFirstname.Text
  rs!LAST_NAME = txtLastname.Text
  rs!MIDDLE_NAME = txtMiddlename.Text
  rs.Update
  MsgBox "Record Created", vbInformation
  Call populateDataGrid
End Sub

Private Sub cmdclear_Click()
txtID.Text = ""
txtUsername.Text = ""
txtPassword.Text = ""
txtFirstname.Text = ""
txtLastname.Text = ""
txtMiddlename.Text = ""
txtRole.Text = ""

End Sub

Private Sub dgAccount_Click()

End Sub

Private Sub cmdEdit_Click()
    rs!ID = txtID
    rs!USERNAME = txtUsername
    rs!PASSWORD = txtPassword
    rs!ROLE = txtRole
    rs!FIRST_NAME = txtFirstname
    rs!LAST_NAME = txtLastname
    rs!MIDDLE_NAME = txtMiddlename
    MsgBox "Record Updated", vbInformation
End Sub

Private Sub Form_Load()
    Call populateDataGrid
End Sub
Private Sub showSelectedData()
  txtID = CommonHelper.extractStringValue(rs!ID)
  txtUsername = CommonHelper.extractStringValue(rs!USERNAME)
  'txtPassword = CommonHelper.extractStringValue(rs!PASSWORD)
  txtRole = CommonHelper.extractStringValue(rs!ROLE)
  txtFirstname = CommonHelper.extractStringValue(rs!FIRST_NAME)
  txtLastname = CommonHelper.extractStringValue(rs!LAST_NAME)
  txtMiddlename = CommonHelper.extractStringValue(rs!MIDDLE_NAME)
  
  End Sub
  
