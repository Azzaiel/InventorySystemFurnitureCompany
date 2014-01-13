VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Category 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Category"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Category.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldcode 
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search for Category"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4575
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
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton btn_search 
         Height          =   495
         Left            =   1920
         Picture         =   "Form_Category.frx":BF53
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dg_categories 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
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
   Begin VB.CommandButton btn_clear 
      Height          =   495
      Left            =   8160
      Picture         =   "Form_Category.frx":CD60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton btn_save 
      Height          =   495
      Left            =   5640
      Picture         =   "Form_Category.frx":DA39
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txt_name 
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
      Left            =   4680
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name:"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_category As New ADODB.Recordset
Dim sql_string As String

Private Sub btn_clear_Click()
    txt_op.Text = "add"
    txt_code.Text = ""
    txt_name.Text = ""
End Sub

Private Sub btn_save_Click()
    If txt_op.Text = "add" Then
        If txt_name.Text = "" Then
            MsgBox "Please complete all fields."
        Else
            If is_duplicate("tbl_category", "Category_Code", txt_name.Text) Then
                MsgBox "Category Code already exists."
                Exit Sub
            End If
            sql_string = "INSERT INTO " _
                            & "tbl_category (Category_Code, Category_Name)" _
                        & " VALUES (" _
                            & "'" & txt_name.Text & "','" & txt_name.Text & "')"
            Call mysql_select(Form_Category.rs_category, sql_string)
            MsgBox "Category added."
            Call Form_Load
        End If
    Else
        If txt_code.Text = "" Or txt_name.Text = "" Then
            MsgBox "Please complete all fields."
        Else
            If txt_code.Text <> txt_oldcode.Text Then
                If is_duplicate("tbl_category", "Category_Code", txt_code.Text) Then
                    MsgBox "Category Code already exists."
                    Exit Sub
                End If
                sql_string = "UPDATE tbl_category SET Category_Code = '" & txt_code.Text & "', Category_Name = '" & txt_name.Text & "' WHERE Category_Code = '" & txt_oldcode.Text & "' "
                Call mysql_select(Form_Category.rs_category, sql_string)
                MsgBox "Category updated."
                Call Form_Load
            Else
                 sql_string = "UPDATE tbl_category SET  Category_Name = '" & txt_name.Text & "' WHERE Category_Code = '" & txt_oldcode.Text & "' "
                Call mysql_select(Form_Category.rs_category, sql_string)
                MsgBox "Category updated."
                Call Form_Load
            End If
        End If
    End If
End Sub

Private Sub btn_search_Click()
      Call set_datagrid(dg_categories, rs_category, _
                                        "SELECT  Category_Name FROM tbl_category WHERE Category_Code = '" & txt_search.Text & "' OR Category_Name = '" & txt_search.Text & "'")
                                        
                    
    If rs_category.RecordCount = 0 Then
        MsgBox "No record found."
    End If
End Sub

Private Sub dg_categories_DblClick()

    If rs_category.RecordCount = 0 Then
        MsgBox "No selected record."
        Exit Sub
    Else
        txt_code.Text = rs_category.Fields("Category_Code")
        txt_oldcode.Text = rs_category.Fields("Category_Code")
        txt_name.Text = rs_category.Fields("Category_Name")
        txt_op.Text = "edit"
    End If
End Sub

Private Sub Form_Load()
      Call set_datagrid(dg_categories, rs_category, _
                                        "SELECT Category_Name FROM tbl_category")
                                        
                    
        txt_op.Text = "add"
        txt_oldcode.Text = ""
        txt_name.Text = ""
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_categories, rs_category, _
                                        "SELECT Category_Name FROM tbl_category WHERE Category_Code LIKE '%" & txt_search.Text & "%' OR Category_Name LIKE '%" & txt_search.Text & "%'")
                                        
                    
        
End Sub
