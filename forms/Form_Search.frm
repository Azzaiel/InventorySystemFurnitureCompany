VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Search.frx":0000
   ScaleHeight     =   3930
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search for Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton btn_search 
         Height          =   495
         Left            =   1920
         Picture         =   "Form_Search.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1575
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   5415
      End
      Begin MSDataGridLib.DataGrid dg_products 
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
End
Attribute VB_Name = "Form_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_product As New ADODB.Recordset
Dim sql_string As String

Private Sub btn_search_Click()
    Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Product_ID = '" & txt_search.Text & "' OR Product_Name = '" & txt_search.Text & "' OR Category = '" & txt_search.Text & "' OR Brand = '" & txt_search.Text & "' OR Initial_Supplier = '" & txt_search.Text & "' OR Remark = '" & txt_search.Text & "'")
    If rs_product.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
End Sub

Private Sub dg_products_DblClick()
If operation = "order" Then
    Form_Order.txt_product_id.Text = rs_product.Fields("Product_ID")
    Form_Order.txt_product_name.Text = rs_product.Fields("Product_Name")
     Form_Order.txt_price.Text = rs_product.Fields("Cost")
    Unload Me
Else
     Form_Purchase.txt_product_id.Text = rs_product.Fields("Product_ID")
    Form_Purchase.txt_product_name.Text = rs_product.Fields("Product_Name")
     Form_Purchase.txt_price.Text = rs_product.Fields("Unit_Price")
    Unload Me
End If
End Sub

Private Sub Form_Load()
     Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark='Active'")
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Product_ID LIKE '%" & txt_search.Text & "%' OR Product_Name LIKE '%" & txt_search.Text & "%' OR Category LIKE '%" & txt_search.Text & "%' OR Brand LIKE '%" & txt_search.Text & "%' OR Initial_Supplier LIKE '%" & txt_search.Text & "%' OR Remark LIKE '%" & txt_search.Text & "%'")
End Sub
