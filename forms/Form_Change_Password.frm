VERSION 5.00
Begin VB.Form Form_Change_Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Security Password"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Change_Password.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Security Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txt_confirm 
         BackColor       =   &H00E0E0E0&
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
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox txt_new 
         BackColor       =   &H00E0E0E0&
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
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txt_old 
         BackColor       =   &H00E0E0E0&
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
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
      Begin VB.CommandButton btn_change 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Picture         =   "Form_Change_Password.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
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
         TabIndex        =   7
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
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
         TabIndex        =   6
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
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
         TabIndex        =   5
         Top             =   480
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form_Change_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_password As New ADODB.Recordset
Dim sql_string As String
Private Sub btn_change_Click()
    If txt_old.Text = "" Or txt_new.Text = "" Or txt_confirm.Text = "" Then
        MsgBox "Please complete all fields."
        Exit Sub
    Else
        Call mysql_select(public_rs, "SELECT * FROM tbl_password")
        If public_rs.Fields("Password").Value <> txt_old.Text Then
            MsgBox "Wrong security password."
            Exit Sub
        Else
            If txt_new.Text <> txt_confirm.Text Then
                MsgBox "Password did not match."
                Exit Sub
            Else
                 sql_string = "UPDATE " _
                            & "tbl_password " _
                        & "SET " _
                            & "Password = '" & txt_new.Text & "' " _
                        & "WHERE " _
                            & " ID = '1'"
                 Call mysql_select(rs_password, sql_string)
                MsgBox "System's security password has been updated."
                Call Form_Main.Form_Load
                Unload Me
            End If
        End If
    End If
End Sub

