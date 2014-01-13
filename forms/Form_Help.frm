VERSION 5.00
Begin VB.Form Form_Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About and Help"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Help.frx":0000
   ScaleHeight     =   2745
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton btn_about 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         Picture         =   "Form_Help.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "About"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btn_help 
         Height          =   1815
         Left            =   2760
         Picture         =   "Form_Help.frx":5F71
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Help"
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_about_Click()
    Call load_form(Form_About, True)
End Sub

Private Sub btn_help_Click()
    Call load_form(Form_Help2, True)
End Sub
