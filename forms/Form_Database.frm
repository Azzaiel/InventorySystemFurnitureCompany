VERSION 5.00
Begin VB.Form Form_Database 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Database.frx":0000
   ScaleHeight     =   2775
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
      Begin VB.CommandButton btn_restore 
         Height          =   1815
         Left            =   2760
         Picture         =   "Form_Database.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Restore Database"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btn_backup 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         Picture         =   "Form_Database.frx":652B
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Back up Database"
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_backup_Click()
    Dim response As String
    response = MsgBox("Are you sure you want to proceed?", vbYesNo)
    If (response = vbYes) Then
      Dim my_date As Date
      myDate = Format(Now, "mm-dd-yyyy")
      backup_db (GetShortName(App.Path & "\back-up database") & "\db_inventory.sql")
      MsgBox "Database successfully copied."
    End If
End Sub

Private Sub btn_restore_Click()
     Call load_form(Form_Restore, True)
End Sub
