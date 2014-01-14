VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Purchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Form"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Purchase.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tab_purchase 
      Height          =   7215
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Purchase Form"
      TabPicture(0)   =   "Form_Purchase.frx":BF53
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label25"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label26"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label27"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label28"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label29"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label30"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label31"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label32"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "date_delivery"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dg_purchase"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "date_purchase"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_total"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmb_remark"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "btn_search"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "btn_clear(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "btn_save"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt_total_purchase"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt_price"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt_quantity"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txt_product_name"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txt_product_id"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txt_purchase_no"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt_tax"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txt_amount"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txt_change"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "btn_payment"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "btn_purchase"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txt_op"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "btn_remove"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt_customer"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).ControlCount=   42
      TabCaption(1)   =   "Pending Purchase"
      TabPicture(1)   =   "Form_Purchase.frx":BF6F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_customer_2"
      Tab(1).Control(1)=   "txt_purchase_no2"
      Tab(1).Control(2)=   "txt_product_id2"
      Tab(1).Control(3)=   "txt_product_name2"
      Tab(1).Control(4)=   "txt_quantity2"
      Tab(1).Control(5)=   "txt_price2"
      Tab(1).Control(6)=   "txt_search"
      Tab(1).Control(7)=   "btn_save2"
      Tab(1).Control(8)=   "cmb_remark2"
      Tab(1).Control(9)=   "txt_total2"
      Tab(1).Control(10)=   "btn_search2"
      Tab(1).Control(11)=   "btn_report"
      Tab(1).Control(12)=   "dg_pending"
      Tab(1).Control(13)=   "date_purchase2"
      Tab(1).Control(14)=   "date_delivery2"
      Tab(1).Control(15)=   "Label34"
      Tab(1).Control(16)=   "Label33"
      Tab(1).Control(17)=   "Label14"
      Tab(1).Control(18)=   "Label15"
      Tab(1).Control(19)=   "Label17"
      Tab(1).Control(20)=   "Label18"
      Tab(1).Control(21)=   "Label19"
      Tab(1).Control(22)=   "Label20"
      Tab(1).Control(23)=   "Label21"
      Tab(1).Control(24)=   "Label22"
      Tab(1).Control(25)=   "Label24"
      Tab(1).Control(26)=   "Label23"
      Tab(1).ControlCount=   27
      Begin VB.TextBox txt_customer_2 
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
         Height          =   375
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txt_customer 
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton btn_remove 
         Height          =   495
         Left            =   5400
         Picture         =   "Form_Purchase.frx":BF8B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txt_op 
         Height          =   285
         Left            =   1080
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btn_purchase 
         Enabled         =   0   'False
         Height          =   495
         Left            =   8160
         Picture         =   "Form_Purchase.frx":CDC4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton btn_payment 
         Height          =   495
         Left            =   8160
         Picture         =   "Form_Purchase.frx":DE11
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txt_change 
         BackColor       =   &H00C0C0C0&
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
         Left            =   7320
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   6240
         Width           =   2415
      End
      Begin VB.TextBox txt_amount 
         BackColor       =   &H00C0C0C0&
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
         Left            =   7320
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   5760
         Width           =   2415
      End
      Begin VB.TextBox txt_tax 
         BackColor       =   &H00C0C0C0&
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
         Left            =   7320
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   5280
         Width           =   2415
      End
      Begin VB.TextBox txt_purchase_no 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_product_id 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txt_product_name 
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
         Height          =   375
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txt_quantity 
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txt_price 
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
         Height          =   375
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txt_total_purchase 
         BackColor       =   &H00C0C0C0&
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
         Left            =   7320
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton btn_save 
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
         Left            =   1080
         Picture         =   "Form_Purchase.frx":EDF0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CommandButton btn_clear 
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
         Index           =   0
         Left            =   2880
         Picture         =   "Form_Purchase.frx":FABD
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CommandButton btn_search 
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
         Left            =   2880
         Picture         =   "Form_Purchase.frx":10796
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox cmb_remark 
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
         ItemData        =   "Form_Purchase.frx":115A3
         Left            =   2280
         List            =   "Form_Purchase.frx":115B0
         TabIndex        =   3
         Text            =   "Select"
         Top             =   5400
         Width           =   2775
      End
      Begin VB.TextBox txt_total 
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
         Height          =   375
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox txt_purchase_no2 
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
         Height          =   375
         Left            =   -72720
         TabIndex        =   22
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txt_product_id2 
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
         Height          =   375
         Left            =   -72720
         TabIndex        =   21
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txt_product_name2 
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
         Height          =   375
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txt_quantity2 
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
         Height          =   375
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txt_price2 
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
         Height          =   375
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3960
         Width           =   2535
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
         Left            =   -69600
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton btn_save2 
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
         Left            =   -72360
         Picture         =   "Form_Purchase.frx":115D1
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6000
         Width           =   1575
      End
      Begin VB.ComboBox cmb_remark2 
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
         ItemData        =   "Form_Purchase.frx":1229E
         Left            =   -72720
         List            =   "Form_Purchase.frx":122AB
         TabIndex        =   10
         Text            =   "Select"
         Top             =   4920
         Width           =   2775
      End
      Begin VB.TextBox txt_total2 
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
         Height          =   375
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton btn_search2 
         Height          =   495
         Left            =   -68280
         Picture         =   "Form_Purchase.frx":122CC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton btn_report 
         Height          =   495
         Left            =   -66840
         Picture         =   "Form_Purchase.frx":130D9
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker date_purchase 
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108068865
         CurrentDate     =   41518
      End
      Begin MSDataGridLib.DataGrid dg_purchase 
         Height          =   3255
         Left            =   5400
         TabIndex        =   25
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
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
      Begin MSDataGridLib.DataGrid dg_pending 
         Height          =   4335
         Left            =   -69600
         TabIndex        =   29
         Top             =   1800
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7646
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
      Begin MSComCtl2.DTPicker date_purchase2 
         Height          =   375
         Left            =   -72720
         TabIndex        =   30
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108068865
         CurrentDate     =   41518
      End
      Begin MSComCtl2.DTPicker date_delivery 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   5880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108068865
         CurrentDate     =   41518
      End
      Begin MSComCtl2.DTPicker date_delivery2 
         Height          =   375
         Left            =   -72720
         TabIndex        =   11
         Top             =   5400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108068865
         CurrentDate     =   41518
      End
      Begin VB.Label Label34 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72720
         TabIndex        =   67
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label33 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72720
         TabIndex        =   66
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   65
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   64
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label Label30 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   63
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label29 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   62
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   61
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   60
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "Change:"
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
         Left            =   5400
         TabIndex        =   57
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "Tendered:"
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
         Left            =   5400
         TabIndex        =   55
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "VAT (12%):"
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
         Left            =   5400
         TabIndex        =   53
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Purchase Number:"
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
         TabIndex        =   52
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Purchase Date:"
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
         TabIndex        =   51
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Customer Name:"
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
         TabIndex        =   50
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Product ID:"
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
         TabIndex        =   49
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Product Name:"
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
         TabIndex        =   48
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Quantity:"
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
         TabIndex        =   47
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Price:"
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
         TabIndex        =   46
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Total:"
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
         TabIndex        =   45
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Remark:"
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
         TabIndex        =   44
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Total Purchase:"
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
         Left            =   5400
         TabIndex        =   43
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Expected Delivery:"
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
         TabIndex        =   42
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Purchase Number:"
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
         Left            =   -74760
         TabIndex        =   41
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Customer Name:"
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
         Left            =   -74760
         TabIndex        =   40
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Product ID:"
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Product Name:"
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
         Left            =   -74760
         TabIndex        =   38
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Quantity:"
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
         Left            =   -74760
         TabIndex        =   37
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Price:"
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Total:"
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
         Left            =   -74760
         TabIndex        =   35
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "Remark:"
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
         Left            =   -74760
         TabIndex        =   34
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Expected Delivery:"
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
         Left            =   -74760
         TabIndex        =   33
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Label23 
         Caption         =   "Purchase Date:"
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
         Left            =   -74760
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_purchase As New ADODB.Recordset
Public rs_temp As New ADODB.Recordset
Public rs_pending As New ADODB.Recordset
Public rs_cashier As New ADODB.Recordset
Dim sql_string As String
Dim no As Integer
Dim cost, total, quan2, critical, all, tax As Double
Public quan As Integer

Private Sub btn_clear_Click(Index As Integer)
    Call Form_Load
End Sub

Private Sub btn_payment_Click()
    If txt_total_purchase.Text = " 0" Then
        MsgBox "Please purchase an item first before paying an amount."
        Exit Sub
    End If
    
    Call mysql_select(public_rs, "SELECT PRODUCT_ID, SUM(Quantity) as REQ_Quantity FROM tbl_temp_purchase Where Remark <> 'Pending' GROUP BY PRODUCT_ID")
    
    Dim productRS As New ADODB.Recordset
    
    While Not public_rs.EOF
      Call mysql_select(productRS, "SELECT * FROM tbl_product WHERE Product_ID='" & public_rs!PRODUCT_ID & "'")
      If (productRS.RecordCount > 0) Then
        If (val(public_rs!REQ_Quantity) > val(productRS!Quantity)) Then
          MsgBox "Product  " & productRS!PRODUCT_NAME & " has in sufficient stock. Requested quantity is " & val(public_rs!REQ_Quantity) & " in stock is " & val(productRS!Quantity)
          Exit Sub
        End If
      End If
      public_rs.MoveNext
    Wend
    
    Dim amount As String
    Dim payment As Double
    amount = InputBox("Enter an amount:", "Payment")
    txt_amount.Text = Format(amount, "###,###.00")
    
    payment = val(amount)
    If payment < all Then
        MsgBox "Please pay higher amount."
        txt_amount.Text = ""
        txt_change.Text = ""
        btn_purchase.Enabled = False
        Exit Sub
    Else
        txt_change.Text = Format(Str(payment - all), "###,###.00")
        btn_purchase.Enabled = True
    End If
End Sub

Private Sub btn_purchase_Click()

    If txt_amount.Text = "" Then
        MsgBox "Please input your payment first."
        Exit Sub
    Else
     Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
      company_name = public_rs.Fields("Name").Value
            dr_receipt.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_receipt.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_receipt.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
       Call mysql_select(rs_cashier, "SELECT * FROM tbl_users WHERE Username='" & Form_Main.lbl_username.Caption & "'")
        dr_receipt.Sections(2).Controls("lbl_date").Caption = Now
        dr_receipt.Sections(2).Controls("lbl_cashier").Caption = rs_cashier.Fields("Lastname").Value & ", " & rs_cashier.Fields("Firstname").Value
         dr_receipt.Sections(5).Controls("lbl_total").Caption = txt_total_purchase.Text
        dr_receipt.Sections(5).Controls("lbl_tax").Caption = txt_tax.Text
         dr_receipt.Sections(5).Controls("lbl_amount").Caption = txt_amount.Text
        dr_receipt.Sections(5).Controls("lbl_change").Caption = txt_change.Text
    Set dr_receipt.DataSource = rs_temp
    dr_receipt.Show vbModal, Me
End If
     
     Dim rsTempPurchase As New ADODB.Recordset
     Dim rsProduct As New ADODB.Recordset
     Call mysql_select(rsTempPurchase, "SELECT * FROM tbl_temp_purchase Where Remark <> 'Pending' ")
     
     rsTempPurchase.MoveFirst
     While Not (rsTempPurchase.EOF)
       Call mysql_select(rsProduct, "SELECT * FROM tbl_product WHERE Product_ID='" & rsTempPurchase!PRODUCT_ID & "' ")
       If (rsProduct.RecordCount > 0) Then
         rsProduct!Quantity = val(rsProduct!Quantity) - val(rsTempPurchase!Quantity)
         rsProduct.Update
       End If
       rsTempPurchase.MoveNext
     Wend
     

      Call mysql_select(public_rs, "DELETE FROM tbl_temp_purchase where Remark <> 'Pending' ")
      Call Form_Load
      txt_tax.Text = "0"
    txt_total_purchase.Text = "0"
    txt_amount.Text = "0"
    txt_change.Text = "0"
    Call Form_Main.Form_Load
End Sub

Private Sub btn_remove_Click()
    If rs_temp.RecordCount = 0 Then
        MsgBox "No item to remove."
        Exit Sub
    Else
    Call mysql_select(public_rs, "DELETE FROM tbl_temp_purchase WHERE Purchase_ID='" & rs_temp.Fields("Purchase_ID").Value & "' ")
    Call mysql_select(public_rs, "DELETE FROM tbl_purchase WHERE Purchase_ID='" & rs_temp.Fields("Purchase_ID").Value & "' ")
     Call Form_Load
    End If
End Sub

Private Sub btn_report_Click()
    If rs_pending.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_pending_purchase.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_pending_purchase.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_pending_purchase.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_pending_purchase.DataSource = rs_pending
    dr_pending_purchase.Show vbModal, Me
    End If
End Sub

Private Sub btn_save_Click()
     If txt_op.Text = "add" Then
        If txt_purchase_no.Text = "" Or txt_customer.Text = "" Or txt_customer.Text = "" Or txt_product_id.Text = "" Or txt_product_name.Text = "" Or txt_price.Text = "" Or txt_quantity.Text = "" Or txt_total.Text = "" Or cmb_remark.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            If txt_quantity.Text = "0" Then
                MsgBox "Please input number of purchase item."
                Exit Sub
            Else
            
                
             sql_string = "INSERT INTO " _
                        & "tbl_purchase (Purchase_ID, Purchase_Date,Customer_Name," _
                        & "Person_In_Charge,Product_ID, Quantity,Total,Remark,Expected_Delivery)" _
                    & " VALUES (" _
                        & "'" & txt_purchase_no.Text & "', " & "'" & date_purchase.Value & "','" & txt_customer.Text & "','" _
                        & txt_customer.Text & "','" & txt_product_id.Text & "','" _
                        & txt_quantity.Text & "','" & txt_total.Text & "','" & cmb_remark.Text & "','" & date_delivery.Value & "')"
                        
            
           Call mysql_select(rs_purchase, sql_string)
            
            sql_string = "INSERT INTO " _
                        & "tbl_temp_purchase (Purchase_ID,Purchase_Date,Customer_Name," _
                        & "Person_In_Charge,Product_ID,Quantity,Total,Remark,Expected_Delivery)" _
                    & " VALUES (" _
                        & "'" & txt_purchase_no.Text & "','" & date_purchase.Value & "','" & txt_customer.Text & "','" _
                        & txt_customer.Text & "','" & txt_product_id.Text & "','" _
                        & txt_quantity.Text & "','" & txt_total.Text & "','" & cmb_remark.Text & "','" & date_delivery.Value & "')"
            
            Call mysql_select(rs_purchase, sql_string)
                  
            MsgBox "Purchase added."
            txt_op.Text = "add"
            Call Form_Load
            End If
        End If
    Else
        If txt_purchase_no.Text = "" Or txt_customer.Text = "" Or txt_customer.Text = "" Or txt_product_id.Text = "" Or txt_product_name.Text = "" Or txt_price.Text = "" Or txt_quantity.Text = "" Or txt_total.Text = "" Or cmb_remark.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            If txt_quantity.Text = "0" Then
                MsgBox "Please input number of purchase item."
                Exit Sub
            Else
            sql_string = "UPDATE " _
                                & "tbl_purchase " _
                            & "SET " _
                                & "Purchase_Date = '" & date_purchase.Value & "', Customer_Name = '" & txt_customer.Text & "'," _
                                & "Person_In_Charge = '" & txt_customer.Text & "',Product_ID = '" _
                                & txt_product_id.Text & "',Quantity = '" & txt_quantity.Text & "',Total" _
                                & " = '" & txt_total.Text & "', Remark= '" & cmb_remark.Text & "', Expected_Delivery='" & date_delivery.Value & "'" _
                            & "WHERE " _
                                & " Purchase_ID = " & txt_purchase_no.Text & ""
            Call mysql_select(rs_purchase, sql_string)
          sql_string = "UPDATE " _
                                & "tbl_temp_purchase " _
                            & "SET " _
                                & "Purchase_Date = '" & date_purchase.Value & "', Customer_Name = '" & txt_customer.Text & "'," _
                                & "Person_In_Charge = '" & txt_customer.Text & "',Product_ID = '" _
                                & txt_product_id.Text & "',Quantity = '" & txt_quantity.Text & "',Total" _
                                & " = '" & txt_total.Text & "', Remark= '" & cmb_remark.Text & "', Expected_Delivery='" & date_delivery.Value & "'" _
                            & "WHERE " _
                                & " Purchase_ID = '" & txt_purchase_no.Text & "'"
            Call mysql_select(rs_purchase, sql_string)
                    MsgBox "Purchase updated."
                    txt_op.Text = "add"
            Call Form_Load
        End If
        End If
    End If
   Call Form_Main.Form_Load
End Sub

Private Sub btn_save2_Click()
      If txt_purchase_no2.Text = "" Then
        MsgBox "Nothing to edit."
        Exit Sub
    Else
    
    If txt_purchase_no2.Text = "" Or txt_customer_2.Text = "" Or txt_customer_2.Text = "" Or txt_product_id2.Text = "" Or txt_product_name2.Text = "" Or txt_price2.Text = "" Or txt_quantity2.Text = "" Or txt_total2.Text = "" Or cmb_remark2.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
             Call mysql_select(public_rs, "SELECT * FROM tbl_purchase WHERE Purchase_ID = " & txt_purchase_no2.Text)
             public_rs!REMARK = cmb_remark2.Text
             public_rs!EXPECTED_DELIVERY = date_delivery2
             public_rs.Update
             
             Call mysql_select(public_rs, "SELECT * FROM tbl_temp_purchase WHERE Purchase_ID = " & txt_purchase_no2.Text)
             public_rs!REMARK = cmb_remark2.Text
             public_rs!EXPECTED_DELIVERY = date_delivery2
             public_rs.Update
           
             MsgBox "Pending purchase updated."
             Call clear_all2
            Call Form_Load
        End If
    End If
    Call Form_Main.Form_Load
End Sub

Private Sub btn_search_Click()
    operation = "purchase"
     Call load_form(Form_Search, True)
End Sub

Private Sub btn_search2_Click()
     Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_purchase WHERE Remark = 'Pending' AND (Purchase_ID = '" & txt_search.Text & "' OR Purchase_Date = '" & txt_search.Text & "' OR Customer_Name = '" & txt_search.Text & "' OR Person_In_Charge = '" & txt_search.Text & "' OR Product_ID = '" & txt_search.Text & "' OR Remark = '" & txt_search.Text & "' OR Expected_Delivery = '" & txt_search.Text & "') ")
     Call formatPendingDataGrid
     dg_pending.Columns(3).Visible = False
            If rs_pending.RecordCount = 0 Then
                MsgBox "Record not found."
                Exit Sub
            End If
End Sub

Private Sub cmb_customer_Click()
       Call mysql_select(public_rs, "SELECT * FROM tbl_customer WHERE Customer_Name='" & cmb_customer.Text & "'")
    cmb_person.Clear
    While Not public_rs.EOF
        cmb_person.AddItem (public_rs.Fields("Representative1"))
        If public_rs.Fields("Representative2").Value <> "" Then
            
         cmb_person.AddItem (public_rs.Fields("Representative2"))
        End If
        public_rs.MoveNext
    Wend
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
    
End Sub

Private Sub cmb_customer_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose customer from the list."
    cmb_customer.Text = ""
End Sub

Private Sub cmb_person_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose person-in-charge from the list."
    cmb_person.Text = ""
End Sub

Public Sub checkRemainingStock()
    quan = val(txt_quantity.Text)
     Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID= '" & txt_product_id.Text & "'")
    If public_rs.RecordCount <> 0 Then
        Dim remaing As Integer
        remaing = val(public_rs!Quantity) - quan
        If remaing < 0 Then
            MsgBox "Your order is greater than the number of stocks."
            txt_quantity.Text = "0"
            Exit Sub
        ElseIf remaing <= val(public_rs!Critical_Point) Then
            MsgBox "You have reached the critical point of this product."
        End If
      End If
End Sub

Private Sub cmb_remark_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose remark from the list."
    cmb_remark.Text = ""
End Sub

Private Sub cmb_remark2_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose remark from the list."
    cmb_remark2.Text = ""
End Sub

Private Sub dg_pending_DblClick()
If rs_pending.RecordCount = 0 Then
    MsgBox "No record found."
    Exit Sub
Else

    txt_customer_2 = ""
    date_purchase2.Value = Now
    date_purchase2.Value = Now
     txt_op.Text = "edit"
    txt_purchase_no2.Text = rs_pending.Fields("Purchase_ID").Value
    date_purchase2.Value = rs_pending.Fields("Purchase_Date").Value
     txt_customer_2.Text = rs_pending.Fields("Customer_Name").Value
     txt_product_id2.Text = rs_pending.Fields("Product_ID").Value
       txt_quantity2.Text = rs_pending.Fields("Quantity").Value
     txt_total2.Text = rs_pending.Fields("Total").Value
      cmb_remark2.Text = rs_pending.Fields("Remark").Value
       date_delivery2.Value = rs_pending.Fields("Expected_Delivery").Value
       Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID = '" & rs_pending.Fields("Product_ID").Value & "' ")
       txt_product_name2.Text = public_rs.Fields("Product_Name").Value
       txt_price2.Text = public_rs.Fields("Cost").Value
End If
End Sub

Private Sub dg_purchase_Click()
If rs_temp.RecordCount = 0 Then
    MsgBox "No record found."
    Exit Sub
Else
     txt_op.Text = "edit"
    txt_purchase_no.Text = rs_temp.Fields("Purchase_ID").Value
    date_purchase.Value = rs_temp.Fields("Purchase_Date").Value
     txt_customer.Text = rs_temp.Fields("Customer_Name").Value
      txt_product_id.Text = rs_temp.Fields("Product_ID").Value
       txt_quantity.Text = rs_temp.Fields("Quantity").Value
     txt_total.Text = rs_temp.Fields("Total").Value
      cmb_remark.Text = rs_temp.Fields("Remark").Value
       date_delivery.Value = rs_temp.Fields("Expected_Delivery").Value
       Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID = '" & rs_temp.Fields("Product_ID").Value & "' ")
       txt_product_name.Text = public_rs.Fields("Product_Name").Value
       txt_price.Text = public_rs.Fields("Cost").Value
End If
End Sub
Private Sub formatPendingDataGrid()
  With dg_pending
    .Columns(6).NumberFormat = "###,###.00"
  End With
End Sub

Public Sub Form_Load()
      Call clear_all
    Call mysql_select(public_rs, "SELECT * FROM tbl_purchase ORDER BY Purchase_ID DESC LIMIT 1")
    If public_rs.RecordCount = 0 Then
        no = 1
    Else
        no = public_rs.Fields("Purchase_ID").Value + 1
    End If
    txt_purchase_no.Text = Str(no)
    date_purchase.Value = Now
    date_delivery.Value = Now
    txt_total_purchase.Text = "0"
    txt_op.Text = "add"
    Call set_datagrid(dg_purchase, rs_temp, _
                                        "SELECT * FROM tbl_temp_purchase where Remark <> 'Pending' ")
      dg_purchase.Columns(3).Visible = False
     
     Call mysql_select(public_rs, "SELECT * FROM tbl_temp_purchase where Remark <> 'Pending' ")
     all = 0
     While Not public_rs.EOF
        all = all + val(public_rs.Fields("Total").Value)
        public_rs.MoveNext
    Wend
    txt_total_purchase.Text = Format(Str(all), "###,###.00")
    tax = all * 0.12
    txt_tax.Text = Format(Str(tax), "###,###.00")
     Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_purchase WHERE Remark='Pending'")
     Call formatPendingDataGrid
    dg_pending.Columns(3).Visible = False
    tab_purchase.Tab = 0
    
    Call formatDataGrid
    date_purchase = Now
    
End Sub
Private Sub formatDataGrid()
  With dg_purchase
    .Columns(6).NumberFormat = "###,###.00"
  End With
End Sub

Public Sub clear_all()
    txt_purchase_no.Text = ""
    txt_customer.Text = ""
    date_purchase.Value = Now
    txt_product_id.Text = ""
    txt_product_name.Text = ""
    txt_quantity.Text = ""
    txt_price.Text = ""
    txt_total.Text = ""
    cmb_remark.Text = ""
    date_delivery.Value = Now
End Sub
Public Sub clear_all2()
    txt_purchase_no2.Text = ""
    txt_customer_2.Text = ""
    date_purchase2.Value = Now
    txt_product_id2.Text = ""
    txt_product_name2.Text = ""
    txt_quantity2.Text = ""
    txt_price2.Text = ""
    txt_total2.Text = ""
    cmb_remark2.Text = ""
    date_delivery2.Value = Now
End Sub

Private Sub Text7_Change()

End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    
End Sub

Private Sub Text7_LinkClose()

End Sub

Private Sub txt_quantity_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_quantity.Text) Then
    txt_quantity.Text = ""
     MsgBox "Invalid input."
Else
        Call checkRemainingStock
         quan = val(txt_quantity.Text)
        cost = val(txt_price.Text)
        
        total = quan * cost
        txt_total.Text = Str(total)
        
  End If
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_purchase WHERE Remark='Pending' AND (Purchase_ID LIKE '%" & txt_search.Text & "%' OR Purchase_Date LIKE '%" & txt_search.Text & "%' OR Customer_Name LIKE '%" & txt_search.Text & "%' OR Person_In_Charge LIKE '%" & txt_search.Text & "%' OR Product_ID LIKE '%" & txt_search.Text & "%' OR Remark LIKE '%" & txt_search.Text & "%' OR Expected_Delivery LIKE '%" & txt_search.Text & "%' )")
    Call formatPendingDataGrid
End Sub

