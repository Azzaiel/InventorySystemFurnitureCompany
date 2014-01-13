VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Sales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Sales.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Day Sales"
      TabPicture(0)   =   "Form_Sales.frx":A636
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_sales"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "date_sales_to"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dg_sales"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "date_sales_from"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btn_report_sales"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton btn_report_sales 
         Height          =   495
         Left            =   8040
         Picture         =   "Form_Sales.frx":A652
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7680
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker date_sales_from 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   73072641
         CurrentDate     =   41584
      End
      Begin MSDataGridLib.DataGrid dg_sales 
         Height          =   5775
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10186
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
      Begin MSComCtl2.DTPicker date_sales_to 
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   73072641
         CurrentDate     =   41584
      End
      Begin VB.Label lbl_sales 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   7320
         Width           =   9375
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search for Product Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton btn_report_prod 
         Height          =   495
         Left            =   2400
         Picture         =   "Form_Sales.frx":B6DA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7800
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker date_from 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   73072641
         CurrentDate     =   41584
      End
      Begin VB.TextBox txt_search 
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3615
      End
      Begin MSDataGridLib.DataGrid dg_purchase 
         Height          =   5535
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9763
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
      Begin MSComCtl2.DTPicker date_to 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   73072641
         CurrentDate     =   41584
      End
      Begin VB.Label lbl_total 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   7440
         Width           =   3855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   2160
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form_Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_purchase As New ADODB.Recordset
Public rs_sales As New ADODB.Recordset
Public total, sales As Double

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub btn_report_prod_Click()
    If dg_purchase.DataSource Is Nothing Then
        MsgBox "No defined product."
    Exit Sub
Else
    
    If rs_purchase.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
        
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            
        dr_product_sales.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
        dr_product_sales.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_product_sales.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_product_sales.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
        dr_product_sales.Sections(2).Controls("lbl_prod_id").Caption = "Product ID: " & rs_purchase.Fields("Product_ID").Value
         dr_product_sales.Sections(2).Controls("lbl_prod_name").Caption = "Product Name: " & rs_purchase.Fields("Product_Name").Value
          dr_product_sales.Sections(2).Controls("lbl_from").Caption = date_from.Value
          dr_product_sales.Sections(2).Controls("lbl_to").Caption = date_to.Value
          dr_product_sales.Sections(2).Controls("lbl_total").Caption = total
         Set dr_product_sales.DataSource = rs_purchase
    dr_product_sales.Show vbModal, Me
End If
End If
End Sub

Private Sub btn_report_sales_Click()
      If dg_sales.DataSource Is Nothing Then
        MsgBox "No defined product."
    Exit Sub
Else
    
    If rs_sales.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
        
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            
        dr_sales.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
        dr_sales.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_sales.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_sales.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
          dr_sales.Sections(2).Controls("lbl_from").Caption = date_sales_from.Value
          dr_sales.Sections(2).Controls("lbl_to").Caption = date_sales_to.Value
          dr_sales.Sections(2).Controls("lbl_total").Caption = sales
         Set dr_sales.DataSource = rs_sales
    dr_sales.Show vbModal, Me
End If
End If
End Sub

Private Sub date_from_Change()
    Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (a.Product_ID = '" & txt_search.Text & "' OR b.Product_Name = '" & txt_search.Text & "') AND (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_to.Value & "','%m/%d/%Y'))")
   
    Call mysql_select(public_rs, "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (a.Product_ID = '" & txt_search.Text & "' OR b.Product_Name = '" & txt_search.Text & "') AND (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_to.Value & "','%m/%d/%Y'))")
    If public_rs.RecordCount = 0 Then
        lbl_total.Caption = "Total: P 0"
        total = 0
    Else
        
        total = 0
        While Not public_rs.EOF
            total = total + val(public_rs.Fields("Total").Value)
            public_rs.MoveNext
        Wend
        lbl_total.Caption = "Total: P " & total
    End If
End Sub

Private Sub date_sales_from_Change()
    Call set_datagrid(dg_sales, rs_sales, _
                                        "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_sales_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_sales_to.Value & "','%m/%d/%Y'))")
    Call mysql_select(public_rs, "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_sales_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_sales_to.Value & "','%m/%d/%Y'))")
    If public_rs.RecordCount = 0 Then
        lbl_sales.Caption = "Total Sales: P 0"
        sales = 0
    Else
        
        sales = 0
        While Not public_rs.EOF
            sales = sales + val(public_rs.Fields("Total").Value)
            public_rs.MoveNext
        Wend
        lbl_sales.Caption = "Total Sales: P " & sales
    End If
End Sub

Private Sub date_sales_to_Change()
    Call set_datagrid(dg_sales, rs_sales, _
                                        "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_sales_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_sales_to.Value & "','%m/%d/%Y'))")
    Call mysql_select(public_rs, "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_sales_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_sales_to.Value & "','%m/%d/%Y'))")
    If public_rs.RecordCount = 0 Then
        lbl_sales.Caption = "Total Sales: P 0"
        sales = 0
    Else
        
        sales = 0
        While Not public_rs.EOF
            sales = sales + val(public_rs.Fields("Total").Value)
            public_rs.MoveNext
        Wend
        lbl_sales.Caption = "Total Sales: P " & sales
    End If
End Sub

Private Sub date_to_Change()
    Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (a.Product_ID = '" & txt_search.Text & "' OR b.Product_Name = '" & txt_search.Text & "') AND (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_to.Value & "','%m/%d/%Y'))")
   
    Call mysql_select(public_rs, "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (a.Product_ID = '" & txt_search.Text & "' OR b.Product_Name = '" & txt_search.Text & "') AND (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_to.Value & "','%m/%d/%Y'))")
    If public_rs.RecordCount = 0 Then
        lbl_total.Caption = "Total: P 0"
        sales = 0
    Else
        
        total = 0
        While Not public_rs.EOF
            total = total + val(public_rs.Fields("Total").Value)
            public_rs.MoveNext
        Wend
        lbl_total.Caption = "Total: P " & total
    End If
End Sub

Private Sub Form_Load()
    date_from.Value = Now
    date_to.Value = Now
    date_sales_from.Value = Now
    date_sales_to.Value = Now
    Call set_datagrid(dg_sales, rs_sales, _
                                        "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & Now & "','%m/%d/%Y') AND STR_TO_DATE('" & Now & "','%m/%d/%Y'))")
    Call mysql_select(public_rs, "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & Now & "','%m/%d/%Y') AND STR_TO_DATE('" & Now & "','%m/%d/%Y'))")
    If public_rs.RecordCount = 0 Then
        lbl_sales.Caption = "Total Sales: P 0"
        sales = 0
    Else
        
        sales = 0
        While Not public_rs.EOF
            sales = sales + val(public_rs.Fields("Total").Value)
            public_rs.MoveNext
        Wend
        lbl_sales.Caption = "Total Sales: P " & sales
    End If

End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
    Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (a.Product_ID = '" & txt_search.Text & "' OR b.Product_Name = '" & txt_search.Text & "') AND (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_to.Value & "','%m/%d/%Y'))")
   
    Call mysql_select(public_rs, "SELECT a.*, b.Product_Name FROM tbl_purchase a LEFT JOIN tbl_product b ON a.Product_ID=b.Product_ID WHERE (a.Product_ID = '" & txt_search.Text & "' OR b.Product_Name = '" & txt_search.Text & "') AND (STR_TO_DATE(Purchase_Date,'%m/%d/%Y') BETWEEN STR_TO_DATE('" & date_from.Value & "','%m/%d/%Y') AND STR_TO_DATE('" & date_to.Value & "','%m/%d/%Y'))")
    If public_rs.RecordCount = 0 Then
        lbl_total.Caption = "Total: P 0"
        total = 0
    Else
        
        total = 0
        While Not public_rs.EOF
            total = total + val(public_rs.Fields("Total").Value)
            public_rs.MoveNext
        Wend
        lbl_total.Caption = "Total: P " & total
    End If
End Sub
