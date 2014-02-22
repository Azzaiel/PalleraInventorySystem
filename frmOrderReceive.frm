VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOrderReceive 
   Caption         =   "Order Receive"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmbAccpectOrder 
      Caption         =   "Accept Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order Items"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   7815
      Begin MSDataGridLib.DataGrid dgOrderItems 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4895
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
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Total Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblTotalCost 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order Detail"
      Height          =   2175
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Label lblOrderBy 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblOrderDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblSuplier 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblOrderID 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  By"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  Date"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Order ID"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Status"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOrderReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset
Private tempRs As ADODB.Recordset

Private Sub cmbAccpectOrder_Click()

  If (rs.RecordCount = 0) Then
    MsgBox "No Order to accpet", vbCritical
    Exit Sub
  End If

  Dim ans
  ans = MsgBox("Are you sure you want to Continue?", vbYesNo)
  If ans = vbYes Then
    rs.MoveFirst
    While Not rs.EOF
      Set tempRs = DataCrudDao.getItemRSByID(rs!ITEM_ID)
      If (tempRs.RecordCount > 0) Then
        tempRs!quantity = Val(CommonHelper.extractStringValue(tempRs!quantity)) + Val(rs!quantity)
        tempRs.Update
      End If
      Call DbInstance.closeRecordSet(tempRs)
      rs.MoveNext
    Wend
    
    Set tempRs = DataCrudDao.getOrderByIDRs(Val(lblOrderID))
    If (tempRs.RecordCount > 0) Then
      tempRs!status = "Completed"
      tempRs!RECIVED_DATE = Now
      tempRs!RECIVED_BY = UserSession.getLoginUser
      tempRs.Update
    End If
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Order Accpected!!", vbInformation
    
    Unload Me
    
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call frmOrder.Form_Load
End Sub
