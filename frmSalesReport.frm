VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesReport 
   Caption         =   "Sales Report"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   13215
      Begin MSDataGridLib.DataGrid dgSales 
         Height          =   6495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   12735
         _ExtentX        =   22463
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
         Picture         =   "frmSalesReport.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13335
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
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   8055
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
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   1695
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
         Left            =   4680
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
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
         Left            =   1440
         TabIndex        =   1
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
         Format          =   16580611
         CurrentDate     =   41697
      End
      Begin MSComCtl2.DTPicker dtEndDate 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
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
         Format          =   16580611
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
         TabIndex        =   9
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   7215
         Left            =   0
         Picture         =   "frmSalesReport.frx":750D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13335
      End
   End
   Begin VB.Image Image2 
      Height          =   8895
      Left            =   0
      Picture         =   "frmSalesReport.frx":EA1A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19695
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClearSearch_Click()
  dtStartDate = DateAdd("m", -1, Now)
  dtEndDate = Now
End Sub

Private Sub cmdSearch_Click()
  Set rs = DataCrudDao.getSalesReport(dtStartDate.value, DateAdd("d", 1, dtEndDate.value))
  Set dgSales.DataSource = rs
  With dgSales
    .Columns(0).Width = 1500
    .Columns(0).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(0).Alignment = dbgCenter
    
    .Columns(1).Width = 1250
    .Columns(1).Alignment = dbgCenter
    
    .Columns(2).Width = 1900
    
    .Columns(3).Width = 2000
    
    .Columns(4).Width = 2250
    
    .Columns(5).Width = 800
    .Columns(5).Alignment = dbgCenter
    
    .Columns(6).Width = 800
    .Columns(6).Alignment = dbgCenter
    .Columns(6).NumberFormat = Constants.CURRENCY_FORMAT
    
    .Columns(7).Width = 1500
    .Columns(7).Alignment = dbgCenter
    .Columns(7).NumberFormat = Constants.CURRENCY_FORMAT
  End With
End Sub

Private Sub Form_Load()
  Call cmdClearSearch_Click
  Call cmdSearch_Click
End Sub
