VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_Restore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Restore.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Browse Database File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin MSComDlg.CommonDialog cdDB 
         Left            =   3240
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btn_restore 
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
         Picture         =   "Form_Restore.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txt_path 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   4575
      End
      Begin VB.CommandButton btn_browse 
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
         Picture         =   "Form_Restore.frx":47D3
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Restore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_browse_Click()
    cdDB.Filter = ("Sql(*.sql)|*.sql")
    cdDB.ShowOpen
    If Not cdDB.FileName = "" Then
        txt_path.Text = cdDB.FileName
        
    Else
        txt_path.Text = ""
    End If
End Sub

Private Sub btn_restore_Click()
    If txt_path.Text = "" Then
        MsgBox "Please browse for an SQL file."
    Else
        restore_db (GetShortName(txt_path.Text))
        MsgBox "Database successfully restored."
        Call Form_Main.Form_Load
        Unload Me
    End If
End Sub
