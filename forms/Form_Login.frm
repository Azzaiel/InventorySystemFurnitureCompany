VERSION 5.00
Begin VB.Form Form_Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Form"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Login.frx":0000
   ScaleHeight     =   2505
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_clear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      Picture         =   "Form_Login.frx":3A4E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton btn_login 
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      Picture         =   "Form_Login.frx":4727
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txt_password 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txt_username 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_logs As New ADODB.Recordset
Dim sql_string As String
Public no_log As Integer

Private Sub Command1_Click()
    
End Sub

Private Sub btn_clear_Click()
    txt_username.Text = ""
    txt_password.Text = ""
End Sub

Private Sub btn_login_Click()
If txt_username.Text = "" Or txt_password.Text = "" Then
    MsgBox "Please input your username and password."
    Exit Sub
Else
If no_log > 0 Then
    Call mysql_select(public_rs, "SELECT * FROM tbl_users WHERE  BINARY  Username ='" & txt_username.Text & "' AND BINARY  Password = '" & txt_password.Text & "' ")
    If public_rs.RecordCount = 0 Then
        no_log = no_log - 1
        If no_log <> 0 Then
            MsgBox "Wrong username or password. You have " & no_log & " chance(s) to log in. Please input your correct username and password."
            Exit Sub
        End If
        If no_log = 0 Then
            MsgBox "You have reached the maximum number of tries in accessing your account. Please contact your administrator to unlock the program."
            Call load_form(Form_Security, True)
            Exit Sub
        End If
    Else
        user_name = public_rs.Fields("Username").Value
        user_type = public_rs.Fields("Usertype").Value
        Form_Main.lbl_username.Caption = user_name
        sql_string = "INSERT INTO " _
                        & "tbl_logs(Username,Login,Logout)" _
                    & " VALUES (" _
                        & "'" & user_name & "','" & Now & "','None')"
            Call mysql_select(rs_logs, sql_string)
            If user_type = "Administrator" Then
                Form_Main.btn_users.Enabled = True
                Form_Main.security_password.Visible = True
            Else
                Form_Main.btn_users.Enabled = False
                Form_Main.btn_database.Enabled = False
                Form_Choose.btn_sales.Enabled = False
                Form_Main.security_password.Visible = False
            End If
        MsgBox "You have successfully logged in."
          Unload Me
         Call load_form(Form_Main, True)
   End If
End If
End If
End Sub

Public Sub Form_Load()
      Call connect_db
    'public_rs.Open "SELECT * FROM tbl_users", db, adOpenStatic, adLockOptimistic
    no_log = 3
End Sub
