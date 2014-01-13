VERSION 5.00
Begin VB.Form Form_Choose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales and Inventory"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Choose.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Products for"
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
      Top             =   240
      Width           =   5055
      Begin VB.CommandButton btn_inventory 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         Picture         =   "Form_Choose.frx":3A4E
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Porduct Inventory"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btn_sales 
         Height          =   1815
         Left            =   2760
         Picture         =   "Form_Choose.frx":69A3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Product Sales"
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Choose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_inventory_Click()
     Call load_form(Form_Product, True)
End Sub

Private Sub btn_sales_Click()
     Call load_form(Form_Sales, True)
End Sub
