VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFastMovingItems 
   Caption         =   "Fast Moving Query"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgFastMovingItems 
      Height          =   6735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11880
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
Attribute VB_Name = "frmFastMovingItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private Sub Form_Load()
  Set rs = DataCrudDao.getFastMovingItems
  Set dgFastMovingItems.DataSource = rs
  With dgFastMovingItems
    
    .Columns(0).Width = 900
    .Columns(0).Alignment = dbgCenter
         
    .Columns(1).Width = 2500

    .Columns(2).Width = 1500
    
    .Columns(3).Width = 2000
    
    .Columns(4).Width = 900
    .Columns(4).Alignment = dbgCenter
    
    .Columns(5).Width = 900
    .Columns(5).Alignment = dbgCenter
    .Columns(5).NumberFormat = Constants.CURRENCY_FORMAT
    
  End With
End Sub
