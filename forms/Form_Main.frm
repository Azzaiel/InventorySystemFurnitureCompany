VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales and Inventory System for "
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Main.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_customer 
      Height          =   1095
      Left            =   5040
      Picture         =   "Form_Main.frx":247E1
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Customer"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btn_order 
      Height          =   1095
      Left            =   5040
      Picture         =   "Form_Main.frx":25ED7
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Order"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btn_products 
      Height          =   1095
      Left            =   4080
      Picture         =   "Form_Main.frx":275C6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Products"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   6600
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Notification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   6240
      TabIndex        =   13
      Top             =   2520
      Width           =   7215
      Begin MSDataGridLib.DataGrid dg_products 
         Height          =   4215
         Left            =   0
         TabIndex        =   15
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7435
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
         BackStyle       =   0  'Transparent
         Caption         =   "Products under critical point"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   8400
      Width           =   13575
      Begin VB.Label lbl_datetime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   18
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbl_username 
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
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Date and Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Logged as: "
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   13455
      Begin VB.CommandButton btn_logout 
         Height          =   1095
         Left            =   10560
         Picture         =   "Form_Main.frx":28DBC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Logout"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_about 
         Height          =   1095
         Left            =   9480
         Picture         =   "Form_Main.frx":2A714
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Help"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_database 
         Height          =   1095
         Left            =   8400
         Picture         =   "Form_Main.frx":2BCE1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Database"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_report 
         Height          =   1095
         Left            =   7320
         Picture         =   "Form_Main.frx":2D5FD
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Report"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_users 
         Height          =   1095
         Left            =   6240
         Picture         =   "Form_Main.frx":2EDB7
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Users"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_purchase 
         Height          =   1095
         Left            =   5160
         Picture         =   "Form_Main.frx":30353
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Purchase"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_supplier 
         Height          =   1095
         Left            =   3000
         Picture         =   "Form_Main.frx":31A10
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Supplier"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btn_company 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1920
         Picture         =   "Form_Main.frx":3329C
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Company"
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label lbl_phone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "address"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   600
      TabIndex        =   23
      Top             =   5520
      Width           =   5415
   End
   Begin VB.Label lbl_address 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "address"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   600
      TabIndex        =   20
      Top             =   5040
      Width           =   5415
   End
   Begin VB.Label security_password 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Security Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TEST 123456"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   -120
      TabIndex        =   16
      Top             =   2520
      Width           =   6105
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_logs As New ADODB.Recordset
Dim sql_string As String
Public rs_product As New ADODB.Recordset

Private Sub btn_about_Click()
    Call load_form(Form_Help, True)
End Sub

Private Sub btn_company_Click()
    Call load_form(Form_CompanySettings, True)
End Sub

Private Sub btn_customer_Click()
    Call load_form(Form_Customer, True)
End Sub

Private Sub btn_database_Click()
   
    Call load_form(Form_Database, True)
End Sub

Private Sub btn_logout_Click()
    
       Unload Me
End Sub

Private Sub btn_order_Click()
    Call load_form(Form_Order, True)
End Sub

Private Sub btn_products_Click()
    Call load_form(Form_Choose, True)
End Sub

Private Sub btn_purchase_Click()
    Call load_form(Form_Purchase, True)
End Sub

Private Sub btn_report_Click()
     Call load_form(Form_Report, True)
End Sub

Private Sub btn_supplier_Click()
    Call load_form(Form_Supplier, True)
End Sub

Private Sub btn_users_Click()
    Call load_form(Form_Accounts, True)
End Sub

Private Sub dg_products_Click()
    MsgBox "Please restock this product."
    If rs_product.RecordCount = 0 Then
        MsgBox "No selected record."
    Else
        Form_Order.tab_order.Tab = 0
        Form_Order.txt_product_id.Text = rs_product.Fields("Product_ID").Value
        Form_Order.txt_product_name.Text = rs_product.Fields("Product_Name").Value
        Form_Order.txt_price.Text = rs_product.Fields("Cost").Value
        Call load_form(Form_Order, True)
    End If
End Sub

Public Sub Form_Load()
     Call mysql_select(public_rs, "SELECT * FROM tbl_company")
     
     lbl_name.Caption = public_rs.Fields!Name
     lbl_address.Caption = "Address: " & public_rs.Fields!address
     lbl_phone.Caption = "Mobile no: " & public_rs.Fields!mobile_number
     
     Form_Main.Caption = "Sales and Inventory System for " & lbl_name.Caption
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Quantity <= Critical_Point")
                                        
  With dg_products
     .Columns(7).NumberFormat = "##,##0.00"
     .Columns(9).NumberFormat = "##,##0.00"
  End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sql_string = "UPDATE " _
                                & "tbl_logs " _
                            & "SET " _
                                & "Logout = '" & Now & "'" _
                            & "WHERE " _
                                & " Username = '" & lbl_username.Caption & "' AND Logout='None'"
                Call mysql_select(rs_logs, sql_string)
                MsgBox "You have successfully logged out."
       Form_Login.txt_username.Text = ""
       Form_Login.txt_password.Text = ""
       Call load_form(Form_Login, True)
       Call Form_Login.Form_Load
End Sub



Private Sub mn_company_Click()

End Sub

Private Sub security_password_Click()
     Call load_form(Form_Change_Password, True)
End Sub

Private Sub Timer1_Timer()
    lbl_datetime.Caption = Now
End Sub
