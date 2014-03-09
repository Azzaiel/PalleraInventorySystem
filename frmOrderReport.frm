VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderReport 
   Caption         =   "Order Report"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   14910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   14415
      Begin MSDataGridLib.DataGrid dgOrders 
         Height          =   6495
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   11456
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
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "frmOrderReport.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   14535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin MSComCtl2.DTPicker dtStartDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   6
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM, dd yyyy"
         Format          =   60686339
         CurrentDate     =   41697
      End
      Begin VB.CommandButton cmdClearSearch 
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
         Height          =   315
         Left            =   6120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtEndDate 
         Height          =   375
         Left            =   10560
         TabIndex        =   8
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM, dd yyyy"
         Format          =   60686339
         CurrentDate     =   41697
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Field"
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
         TabIndex        =   11
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   9600
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "frmOrderReport.frx":750D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13335
      End
   End
   Begin VB.Image Image2 
      Height          =   9135
      Left            =   0
      Picture         =   "frmOrderReport.frx":EA1A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19815
   End
End
Attribute VB_Name = "frmOrderReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private suplierIdList As Variant
Private tempRs As ADODB.Recordset
Private Sub cmdClearSearch_Click()
  cmbSupplier.ListIndex = -1
  dtStartDate = DateAdd("m", -1, Now)
  dtEndDate = Now
End Sub
Private Sub cmdSearch_Click()
  Dim supplierID As Long
  If cmbSupplier.ListIndex > -1 Then
    supplierID = Val(suplierIdList(cmbSupplier.ListIndex))
  Else
    supplierID = -1
  End If

  Set rs = DataCrudDao.getOrdersReport(supplierID, dtStartDate.value, DateAdd("d", 1, dtEndDate.value))
  Set dgOrders.DataSource = rs
  With dgOrders
    .Columns(0).Width = 800
    .Columns(0).Alignment = dbgCenter

    .Columns(1).Width = 1200
    .Columns(1).Alignment = dbgCenter
    
    .Columns(2).Width = 2500
    
    .Columns(3).Width = 1300
    
    .Columns(4).Width = 2500
    
    .Columns(5).Width = 750
    .Columns(5).Alignment = dbgCenter
    
    .Columns(6).Width = 900
    .Columns(6).NumberFormat = Constants.CURRENCY_FORMAT
    .Columns(6).Alignment = dbgCenter
    
    .Columns(7).Width = 1000
    .Columns(7).NumberFormat = Constants.CURRENCY_FORMAT
    .Columns(7).Alignment = dbgCenter
    
    .Columns(8).Width = 1000
    .Columns(8).Alignment = dbgCenter
    
    .Columns(9).Width = 1500
    .Columns(9).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(9).Alignment = dbgCenter
    
    .Columns(10).Width = 1500
    .Columns(10).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(10).Alignment = dbgCenter
    
    .Columns(11).Width = 1000
    .Columns(11).Alignment = dbgCenter
    
  End With
End Sub

Private Sub Form_Load()
  Call populateLov
  Call cmdClearSearch_Click
  Call cmdSearch_Click
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

