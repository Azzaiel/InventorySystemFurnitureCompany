VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_Accounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Accounts.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldusername 
      Height          =   375
      Left            =   10440
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_oldid 
      Height          =   375
      Left            =   9600
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
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
      Height          =   855
      Left            =   5880
      TabIndex        =   28
      Top             =   6000
      Width           =   3855
      Begin VB.CommandButton btn_edit 
         Height          =   495
         Left            =   2040
         Picture         =   "Form_Accounts.frx":BF53
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton btn_add 
         Height          =   495
         Left            =   240
         Picture         =   "Form_Accounts.frx":CBFD
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
   End
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
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton btn_report 
         Height          =   495
         Left            =   2160
         Picture         =   "Form_Accounts.frx":DB68
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6120
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
         TabIndex        =   4
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton btn_search 
         Height          =   495
         Left            =   1080
         Picture         =   "Form_Accounts.frx":EBF0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dg_users 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7858
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
   Begin TabDlg.SSTab tab_users 
      Height          =   5655
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User Account"
      TabPicture(0)   =   "Form_Accounts.frx":F9FD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_id"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_lastname"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_mobile"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_address"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_middlename"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_firstname"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_op"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "User Log History"
      TabPicture(1)   =   "Form_Accounts.frx":FA19
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dg_logs"
      Tab(1).Control(1)=   "btn_logs"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txt_op 
         Height          =   375
         Left            =   5040
         TabIndex        =   33
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton btn_logs 
         Height          =   495
         Left            =   -69120
         Picture         =   "Form_Accounts.frx":FA35
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4920
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dg_logs 
         Height          =   4215
         Left            =   -74520
         TabIndex        =   31
         Top             =   600
         Width           =   6975
         _ExtentX        =   12303
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
      Begin VB.TextBox txt_firstname 
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
         Left            =   1680
         TabIndex        =   14
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txt_middlename 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   2280
         Width           =   2775
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
         Height          =   975
         Left            =   4800
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   2775
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
         Left            =   4800
         TabIndex        =   16
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txt_lastname 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   1320
         Width           =   2775
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
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Account Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   7335
         Begin VB.ComboBox cmb_usertype 
            Appearance      =   0  'Flat
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
            Height          =   360
            ItemData        =   "Form_Accounts.frx":10ABD
            Left            =   2160
            List            =   "Form_Accounts.frx":10AC7
            TabIndex        =   19
            Text            =   "Select"
            Top             =   960
            Width           =   2775
         End
         Begin VB.CommandButton btn_clear 
            Height          =   495
            Left            =   5400
            Picture         =   "Form_Accounts.frx":10AE0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton btn_save 
            Height          =   495
            Left            =   5400
            Picture         =   "Form_Accounts.frx":117B9
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txt_retype 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            PasswordChar    =   "*"
            TabIndex        =   21
            Top             =   1920
            Width           =   2775
         End
         Begin VB.TextBox txt_password 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   1440
            Width           =   2775
         End
         Begin VB.TextBox txt_username 
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
            TabIndex        =   18
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lbl_retype 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Type Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   27
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lbl_password 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
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
            Left            =   240
            TabIndex        =   26
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Usertype:"
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
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
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
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   1215
         End
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
         Left            =   4800
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
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
         Left            =   4800
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Middle Name:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "First Name:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Last Name:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "User ID:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form_Accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_users As New ADODB.Recordset
Public rs_logs As New ADODB.Recordset
Dim sql_string As String

Private Sub btn_add_Click()
    txt_op.Text = "add"
    Call enable_all
    Call clear_all
End Sub

Private Sub btn_clear_Click()
    Call clear_all
    Call disable_all
    txt_op.Text = ""
    txt_oldid.Text = ""
    txt_oldusername.Text = ""
End Sub

Private Sub btn_edit_Click()
    Call enable_all
    txt_op.Text = "edit"
End Sub

Private Sub btn_logs_Click()
      If rs_logs.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
     Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_logs.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_logs.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_logs.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_logs.DataSource = rs_logs
    dr_logs.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_Click()

    If rs_users.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
     Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_accounts.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
            dr_accounts.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_accounts.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_accounts.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_accounts.DataSource = rs_users
    dr_accounts.Show vbModal, Me
    End If
End Sub

Private Sub btn_save_Click()
    If txt_id.Enabled = False Then
        MsgBox "Nothing to edit."
        Exit Sub
    End If
    If txt_op.Text = "add" Then
        If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Or txt_password.Text = "" Or txt_retype.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            If is_duplicate("tbl_users", "ID", txt_id.Text) Then
                MsgBox "User ID exists."
                Exit Sub
            End If
            If is_duplicate("tbl_users", "Username", txt_username.Text) Then
                MsgBox "Username exists."
                Exit Sub
            End If
            If txt_password.Text <> txt_retype.Text Then
                MsgBox "Password did not match."
                Exit Sub
            End If
            sql_string = "INSERT INTO " _
                        & "tbl_users (ID,Lastname,Firstname," _
                        & "Middlename,MobileNumber,Address,Username,Usertype,Password)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','" & txt_lastname.Text & "','" _
                        & txt_firstname.Text & "','" & txt_middlename.Text & "','" _
                        & txt_mobile.Text & "','" & txt_address.Text & "','" & txt_username.Text & "','" & cmb_usertype.Text & "','" & txt_password.Text & "')"
            Call mysql_select(rs_users, sql_string)
            MsgBox "User added."
            Call Form_Load
        End If
    Else
        If txt_id.Text <> txt_oldid.Text Then
            If txt_username.Text <> txt_oldusername.Text Then
             If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If is_duplicate("tbl_users", "ID", txt_id.Text) Then
                    MsgBox "User ID exists."
                    Exit Sub
                End If
                If is_duplicate("tbl_users", "Username", txt_username.Text) Then
                    MsgBox "Username exists."
                    Exit Sub
                End If
                If txt_password.Text = "" And txt_retype.Text = "" Then
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                If Form_Main.lbl_username.Caption = txt_oldusername.Text Then
                    Form_Main.lbl_username.Caption = txt_username.Text
                End If
                Call Form_Load
                Else
                    If txt_password.Text <> txt_retype.Text Then
                        MsgBox "Password did not match."
                        Exit Sub
                    End If
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "', Password='" & txt_password.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                If Form_Main.lbl_username.Caption = txt_oldusername.Text Then
                    Form_Main.lbl_username.Caption = txt_username.Text
                End If
                Call Form_Load
                End If
                 
            End If
            Else
                If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If is_duplicate("tbl_users", "ID", txt_id.Text) Then
                    MsgBox "User ID exists."
                    Exit Sub
                End If
                If txt_password.Text = "" And txt_retype.Text = "" Then
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                Call Form_Load
                Else
                    If txt_password.Text <> txt_retype.Text Then
                        MsgBox "Password did not match."
                        Exit Sub
                    End If
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "', Password='" & txt_password.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                Call Form_Load
                End If
                 
            End If
        End If
        Else
              If txt_username.Text <> txt_oldusername.Text Then
             If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If is_duplicate("tbl_users", "Username", txt_username.Text) Then
                    MsgBox "Username exists."
                    Exit Sub
                End If
                If txt_password.Text = "" And txt_retype.Text = "" Then
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                If Form_Main.lbl_username.Caption = txt_oldusername.Text Then
                    Form_Main.lbl_username.Caption = txt_username.Text
                End If
                MsgBox "User updated."
                Call Form_Load
                Else
                    If txt_password.Text <> txt_retype.Text Then
                        MsgBox "Password did not match."
                        Exit Sub
                    End If
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "', Password='" & txt_password.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                If Form_Main.lbl_username.Caption = txt_oldusername.Text Then
                    Form_Main.lbl_username.Caption = txt_username.Text
                End If
                Call Form_Load
                End If
                 
            End If
            Else
                If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Or txt_mobile.Text = "" Or txt_address.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If txt_password.Text = "" And txt_retype.Text = "" Then
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                Call Form_Load
                Else
                    If txt_password.Text <> txt_retype.Text Then
                        MsgBox "Password did not match."
                        Exit Sub
                    End If
                    sql_string = "UPDATE " _
                                & "tbl_users " _
                            & "SET " _
                                & "ID = '" & txt_id.Text & "', Lastname = '" & txt_lastname.Text & "'," _
                                & "Firstname = '" & txt_firstname.Text & "',Middlename = '" _
                                & txt_middlename.Text & "',MobileNumber = '" & txt_mobile.Text & "',Address" _
                                & " = '" & txt_address.Text & "', Username= '" & txt_username.Text & "', Usertype='" & cmb_usertype.Text & "', Password='" & txt_password.Text & "'" _
                            & "WHERE " _
                                & " ID = '" & txt_oldid.Text & "'"
                Call mysql_select(rs_users, sql_string)
                MsgBox "User updated."
                Call Form_Load
                End If
                 
            End If
        End If
    End If
End If
End Sub

Private Sub btn_search_Click()
     Call set_datagrid(dg_users, rs_users, _
                                        "SELECT Username,Usertype,ID,Lastname,Firstname,Middlename,MobileNumber, Address FROM tbl_users WHERE Username = '" & txt_search.Text & "'OR Usertype = '" & txt_search.Text & "' OR ID = '" & txt_search.Text & "' OR Lastname = '" & txt_search.Text & "' OR Firstname = '" & txt_search.Text & "'")
    If rs_users.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dg_users_DblClick()
    If rs_users.RecordCount = 0 Then
        MsgBox "No selected record."
        Exit Sub
    Else
        txt_id.Text = rs_users.Fields("ID").Value
        txt_oldid.Text = rs_users.Fields("ID").Value
        txt_lastname.Text = rs_users.Fields("Lastname").Value
        txt_firstname.Text = rs_users.Fields("Firstname").Value
        txt_middlename.Text = rs_users.Fields("Middlename").Value
        txt_username.Text = rs_users.Fields("Username").Value
        txt_oldusername.Text = rs_users.Fields("Username").Value
        txt_mobile.Text = rs_users.Fields("MobileNumber").Value
        txt_address.Text = rs_users.Fields("Address").Value
        cmb_usertype.Text = rs_users.Fields("Usertype").Value
        btn_edit.Enabled = True
        Call set_datagrid(dg_logs, rs_logs, _
                                        "SELECT *  FROM tbl_logs WHERE Username='" & rs_users.Fields("Username").Value & "'")
    End If
    
End Sub

Public Sub Form_Load()
    Call set_datagrid(dg_users, rs_users, _
                                        "SELECT Username,Usertype,ID,Lastname,Firstname,Middlename,MobileNumber, Address FROM tbl_users")
    txt_op.Text = "add"
    txt_oldid.Text = ""
    btn_edit.Enabled = False
    Call clear_all
    Call disable_all
     Call set_datagrid(dg_logs, rs_logs, _
                                        "SELECT *  FROM tbl_logs")
    tab_users.Tab = 0
End Sub

Private Sub txt_firstname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(txt_firstname.Text) Then
         txt_firstname.Text = ""
         MsgBox "Number is not allowed."
         Exit Sub
    End If
End Sub

Private Sub txt_lastname_KeyUp(KeyCode As Integer, Shift As Integer)
      If IsNumeric(txt_lastname.Text) Then
         txt_lastname.Text = ""
         MsgBox "Number is not allowed."
         Exit Sub
    End If
End Sub

Private Sub txt_middlename_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(txt_middlename.Text) Then
         txt_middlename.Text = ""
         MsgBox "Number is not allowed."
         Exit Sub
    End If
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
    Call set_datagrid(dg_users, rs_users, _
                                        "SELECT Username,Usertype,ID,Lastname,Firstname,Middlename,MobileNumber, Address FROM tbl_users WHERE Username LIKE '%" & txt_search.Text & "%'OR Usertype LIKE '%" & txt_search.Text & "%' OR ID LIKE '%" & txt_search.Text & "%' OR Lastname LIKE '%" & txt_search.Text & "%' OR Firstname LIKE '%" & txt_search.Text & "%'")
End Sub
Public Sub enable_all()
    txt_id.Enabled = True
    txt_lastname.Enabled = True
    txt_firstname.Enabled = True
    txt_middlename.Enabled = True
    txt_mobile.Enabled = True
    txt_address.Enabled = True
    txt_username.Enabled = True
    cmb_usertype.Enabled = True
    txt_password.Enabled = True
    txt_retype.Enabled = True
End Sub
Public Sub disable_all()
    txt_id.Enabled = False
    txt_lastname.Enabled = False
    txt_firstname.Enabled = False
    txt_middlename.Enabled = False
    txt_mobile.Enabled = False
    txt_address.Enabled = False
    txt_username.Enabled = False
    cmb_usertype.Enabled = False
    txt_password.Enabled = False
    txt_retype.Enabled = False
End Sub
Public Sub clear_all()
    txt_id.Text = ""
    txt_lastname.Text = ""
    txt_firstname.Text = ""
    txt_middlename.Text = ""
    txt_mobile.Text = ""
    txt_address.Text = ""
    txt_username.Text = ""
    cmb_usertype.Text = "Select"
    txt_password.Text = ""
    txt_retype.Text = ""
    txt_oldid.Text = ""
    txt_oldusername.Text = ""
    txt_op.Text = "add"
End Sub
