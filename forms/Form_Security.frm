VERSION 5.00
Begin VB.Form Form_Security 
   BorderStyle     =   0  'None
   Caption         =   "Security Password"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Security.frx":0000
   ScaleHeight     =   2250
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_login 
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      Picture         =   "Form_Security.frx":3A4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input security password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txt_password 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   600
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form_Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_login_Click()
    If txt_password.Text = "" Then
        MsgBox "Please input the security password."
        Exit Sub
    Else
    Call mysql_select(public_rs, "SELECT * FROM tbl_password")
    If public_rs.Fields("Password").Value = txt_password.Text Then
        Form_Login.txt_username.Text = ""
        Form_Login.txt_password.Text = ""
        Unload Me
        Call Form_Login.Form_Load
    Else
        MsgBox "Wrong security password."
    End If
End If
End Sub
