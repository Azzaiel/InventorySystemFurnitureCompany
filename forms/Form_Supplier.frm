VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_Supplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Supplier.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton btn_report 
         Height          =   495
         Left            =   2880
         Picture         =   "Form_Supplier.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
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
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton btn_search 
         Height          =   495
         Left            =   4440
         Picture         =   "Form_Supplier.frx":4AD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dg_suppliers 
         Height          =   4575
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8070
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
   Begin TabDlg.SSTab tab_supplier 
      Height          =   6015
      Left            =   8160
      TabIndex        =   26
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Supplier Information"
      TabPicture(0)   =   "Form_Supplier.frx":58E3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_address"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_representative2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_representative1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_name"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_id"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_mobile"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_op"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_oldid"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Order History"
      TabPicture(1)   =   "Form_Supplier.frx":58FF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btn_order_history"
      Tab(1).Control(1)=   "dg_history"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Pending Transactions"
      TabPicture(2)   =   "Form_Supplier.frx":591B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btn_report_pending"
      Tab(2).Control(1)=   "dg_pending"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txt_oldid 
         Height          =   375
         Left            =   5520
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_op 
         Height          =   375
         Left            =   5880
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_mobile 
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   8
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txt_id 
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txt_name 
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txt_representative1 
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   6
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txt_representative2 
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   7
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txt_address 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3240
         Width           =   5415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2880
         TabIndex        =   16
         Top             =   4080
         Width           =   3735
         Begin VB.CommandButton btn_edit 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1920
            Picture         =   "Form_Supplier.frx":5937
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btn_add 
            Height          =   495
            Left            =   240
            Picture         =   "Form_Supplier.frx":65E1
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton btn_order_history 
         Height          =   495
         Left            =   -69000
         Picture         =   "Form_Supplier.frx":754C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton btn_report_pending 
         Height          =   495
         Left            =   -69000
         Picture         =   "Form_Supplier.frx":85D4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5400
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dg_history 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid dg_pending 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   17
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
      Begin VB.Label Label1 
         Caption         =   "Supplier ID:"
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
         Left            =   360
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier Name:"
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
         Left            =   360
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Representative:"
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
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Representative:"
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
         Left            =   360
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile Number:"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Address:"
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
         Left            =   360
         TabIndex        =   18
         Top             =   3360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_supplier As New ADODB.Recordset
Public rs_order As New ADODB.Recordset
Public rs_pending As New ADODB.Recordset
Dim sql_string As String

Private Sub btn_add_Click()
    txt_op.Text = "add"
    txt_oldid.Text = ""
    Call submitForm
End Sub

Private Sub btn_clear_Click(Index As Integer)
    btn_edit.Enabled = False
    txt_op.Text = "add"
    txt_oldid.Text = ""
End Sub

Private Sub btn_edit_Click()
    txt_op.Text = "edit"
    Call submitForm
End Sub

Private Sub btn_order_history_Click()
    If dg_history.DataSource Is Nothing Then
        MsgBox "No selected supplier."
        Exit Sub
    Else
        If rs_order.RecordCount = 0 Then
            MsgBox "No record to display."
            Exit Sub
        Else
          Call mysql_select(public_rs, "SELECT * FROM tbl_company")
           
                company_name = public_rs.Fields("Name").Value
                dr_supplier_order.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
            dr_supplier_order.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
            dr_supplier_order.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
            dr_supplier_order.Sections(2).Controls("lbl_prod_id").Caption = "Supplier Name: " & rs_supplier.Fields("Supplier_Name").Value
             Set dr_supplier_order.DataSource = rs_order
        dr_supplier_order.Show vbModal, Me
        End If
    End If
End Sub

Private Sub btn_report_Click()
    If rs_supplier.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
       Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
        dr_suppliers.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
        dr_suppliers.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_suppliers.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_suppliers.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_suppliers.DataSource = rs_supplier
    dr_suppliers.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_pending_Click()
If dg_pending.DataSource Is Nothing Then
    MsgBox "No selected supplier."
    Exit Sub
    Else
    If rs_pending.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_supplier_pending.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_supplier_pending.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_supplier_pending.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
        dr_supplier_pending.Sections(2).Controls("lbl_prod_id").Caption = "Supplier Name: " & rs_supplier.Fields("Supplier_Name").Value
         Set dr_supplier_pending.DataSource = rs_pending
    dr_supplier_pending.Show vbModal, Me
    End If
    End If
End Sub

Private Sub submitForm()
     If txt_id.Enabled = False Then
        MsgBox "Nothing to edit."
        Exit Sub
    End If
    If txt_op.Text = "add" Then
        If txt_id.Text = "" Or txt_name.Text = "" Or txt_representative1.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            If is_duplicate("tbl_supplier", "Supplier_ID", txt_id.Text) Then
                MsgBox "Supplier ID exists."
                Exit Sub
            End If
            sql_string = "INSERT INTO " _
                        & "tbl_supplier (Supplier_ID,Supplier_Name,Representative1," _
                        & "Representative2,Mobile_Number,Address)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','" & txt_name.Text & "','" _
                        & txt_representative1.Text & "','" & txt_representative2.Text & "','" _
                        & txt_mobile.Text & "','" & txt_address.Text & "')"
            Call mysql_select(Form_Supplier.rs_supplier, sql_string)
            MsgBox "Supplier added."
            Call Form_Load
        End If
    Else
        If txt_id.Text <> txt_oldid.Text Then
             If txt_id.Text = "" Or txt_name.Text = "" Or txt_representative1.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If is_duplicate("tbl_supplier", "Supplier_ID", txt_id.Text) Then
                    MsgBox "Supplier ID exists."
                    Exit Sub
                End If
                 sql_string = "UPDATE " _
                                & "tbl_supplier " _
                            & "SET " _
                                & "Supplier_ID = '" & txt_id.Text & "', Supplier_Name = '" & txt_name.Text & "'," _
                                & "Representative1 = '" & txt_representative1.Text & "',Representative2 = '" _
                                & txt_representative2.Text & "',Mobile_Number = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "'" _
                            & "WHERE " _
                                & " Supplier_ID = '" & txt_oldid.Text & "'"
                Call mysql_select(Form_Supplier.rs_supplier, sql_string)
                MsgBox "Supplier updated."
                Call Form_Load
            End If
        Else
             If txt_id.Text = "" Or txt_name.Text = "" Or txt_representative1.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                sql_string = "UPDATE " _
                                & "tbl_supplier " _
                            & "SET " _
                                & "Supplier_ID = '" & txt_id.Text & "', Supplier_Name = '" & txt_name.Text & "'," _
                                & "Representative1 = '" & txt_representative1.Text & "',Representative2 = '" _
                                & txt_representative2.Text & "',Mobile_Number = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "'" _
                            & "WHERE " _
                                & " Supplier_ID = '" & txt_oldid.Text & "'"
                Call mysql_select(Form_Supplier.rs_supplier, sql_string)
                MsgBox "Supplier updated."
                Call Form_Load
            End If
        End If
    End If
End Sub

Private Sub btn_search_Click()
    Call set_datagrid(dg_suppliers, rs_supplier, _
                                        "SELECT * FROM tbl_supplier WHERE Supplier_ID = '" & txt_search.Text & "' OR Supplier_Name = '" & txt_search.Text & "' OR Representative1 = '" & txt_search.Text & "' OR Representative2 = '" & txt_search.Text & "'")
    If rs_supplier.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub dg_suppliers_Click()
If rs_supplier.RecordCount = 0 Then
    MsgBox "No selected record."
    Exit Sub
Else
    btn_edit.Enabled = True
    txt_op.Text = "edit"
    txt_id.Text = rs_supplier.Fields("Supplier_ID").Value
    txt_oldid.Text = rs_supplier.Fields("Supplier_ID").Value
    txt_name.Text = rs_supplier.Fields("Supplier_Name").Value
    txt_representative1.Text = rs_supplier.Fields("Representative1").Value
    txt_representative2.Text = rs_supplier.Fields("Representative2").Value
    txt_mobile.Text = rs_supplier.Fields("Mobile_Number").Value
    txt_address.Text = rs_supplier.Fields("Address")
     Call set_datagrid(dg_history, rs_order, _
                                        "SELECT *  FROM tbl_order  WHERE Supplier_Name='" & rs_supplier.Fields("Supplier_Name").Value & "'")
    Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_order WHERE Supplier_Name='" & rs_supplier.Fields("Supplier_Name").Value & "' AND Remark='Pending'")
End If
End Sub

Private Sub Form_Load()
    Call set_datagrid(dg_suppliers, rs_supplier, _
                                        "SELECT * FROM tbl_supplier")
    txt_op.Text = "add"
    txt_oldid.Text = ""
    'btn_edit.Enabled = False
    Call clear_all
    'Call disable_all
    tab_supplier.Tab = 0
     Call enable_all
     btn_edit.Enabled = True
End Sub
Public Sub enable_all()
    txt_id.Enabled = True
    txt_name.Enabled = True
    txt_representative1.Enabled = True
    txt_representative2.Enabled = True
    txt_mobile.Enabled = True
    txt_address.Enabled = True
End Sub
Public Sub disable_all()
    'txt_id.Enabled = False
    'txt_name.Enabled = False
    'txt_representative1.Enabled = False
    'txt_representative2.Enabled = False
    'txt_mobile.Enabled = False
    'txt_address.Enabled = False
End Sub
Public Sub clear_all()
    txt_id.Text = ""
    txt_name.Text = ""
    txt_representative1.Text = ""
    txt_representative2.Text = ""
    txt_mobile.Text = ""
    txt_address.Text = ""
End Sub
Private Sub txt_mobile_KeyPress(KeyAscii As Integer)
  If (Not isFunctionAscii(KeyAscii) And (Not isNumberAscii(KeyAscii) Or Len(txt_mobile) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub
Private Function isFunctionAscii(ascii As Integer) As Boolean
  If (ascii = 13 Or ascii = 8 Or ascii = 32) Then
    isFunctionAscii = True
  Else
    isFunctionAscii = False
  End If
End Function


Private Function isNumberAscii(ascii As Integer) As Boolean
  If (ascii >= 48 And ascii <= 57) Then
    isNumberAscii = True
  Else
    isNumberAscii = False
  End If
End Function


Private Sub txt_representative1_KeyUp(KeyCode As Integer, Shift As Integer)
      If IsNumeric(txt_representative1.Text) Then
         txt_representative1.Text = ""
         MsgBox "Number is not allowed."
         Exit Sub
    End If
End Sub

Private Sub txt_representative2_KeyUp(KeyCode As Integer, Shift As Integer)
      If IsNumeric(txt_representative2.Text) Then
         txt_representative2.Text = ""
         MsgBox "Number is not allowed."
         Exit Sub
    End If
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
    Call set_datagrid(dg_suppliers, rs_supplier, _
                                        "SELECT * FROM tbl_supplier WHERE Supplier_ID LIKE '%" & txt_search.Text & "%' OR Supplier_Name LIKE '%" & txt_search.Text & "%' OR Representative1 LIKE '%" & txt_search.Text & "%' OR Representative2 LIKE '%" & txt_search.Text & "%'")
End Sub
