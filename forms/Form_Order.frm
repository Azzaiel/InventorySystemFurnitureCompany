VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Order 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Order.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tab_order 
      Height          =   7215
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Order Form"
      TabPicture(0)   =   "Form_Order.frx":BF53
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(11)=   "lbl_deliver"
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(13)=   "Label25"
      Tab(0).Control(14)=   "Label26"
      Tab(0).Control(15)=   "date_deliver"
      Tab(0).Control(16)=   "txt_order_no"
      Tab(0).Control(17)=   "txt_product_id"
      Tab(0).Control(18)=   "txt_product_name"
      Tab(0).Control(19)=   "txt_quantity"
      Tab(0).Control(20)=   "txt_price"
      Tab(0).Control(21)=   "dg_orders"
      Tab(0).Control(22)=   "txt_total_order"
      Tab(0).Control(23)=   "btn_save"
      Tab(0).Control(24)=   "btn_clear_order"
      Tab(0).Control(25)=   "btn_search"
      Tab(0).Control(26)=   "date_order"
      Tab(0).Control(27)=   "cmb_supplier"
      Tab(0).Control(28)=   "cmb_remark"
      Tab(0).Control(29)=   "txt_total"
      Tab(0).Control(30)=   "cmb_person"
      Tab(0).Control(31)=   "txt_op"
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Pending Orders"
      TabPicture(1)   =   "Form_Order.frx":BF6F
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label19"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label20"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label21"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label22"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label24"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label23"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label27"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label28"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "date_deliver2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "date_order2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "dg_pending"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txt_order_no2"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txt_product_id2"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txt_product_name2"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txt_quantity2"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txt_price2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txt_search"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "btn_save_pending"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmb_supplier2"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmb_remark2"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txt_total2"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "btn_search2"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "btn_report"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cmb_person2"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).ControlCount=   29
      Begin VB.ComboBox cmb_person2 
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
         Height          =   360
         Left            =   2280
         TabIndex        =   56
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txt_op 
         Height          =   285
         Left            =   -69720
         TabIndex        =   55
         Top             =   6360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmb_person 
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
         Left            =   -72720
         TabIndex        =   2
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CommandButton btn_report 
         Height          =   495
         Left            =   8160
         Picture         =   "Form_Order.frx":BF8B
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton btn_search2 
         Height          =   495
         Left            =   6720
         Picture         =   "Form_Order.frx":D013
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1080
         Width           =   1575
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
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   4800
         Width           =   2535
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
         ItemData        =   "Form_Order.frx":DE20
         Left            =   2280
         List            =   "Form_Order.frx":DE30
         TabIndex        =   48
         Text            =   "Select"
         Top             =   5280
         Width           =   2775
      End
      Begin VB.ComboBox cmb_supplier2 
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
         Height          =   360
         Left            =   2280
         TabIndex        =   47
         Text            =   "Select"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CommandButton btn_save_pending 
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
         Left            =   2040
         Picture         =   "Form_Order.frx":DE59
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   6360
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
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   600
         Width           =   4335
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
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   4320
         Width           =   2535
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   3840
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3360
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
         Left            =   2280
         TabIndex        =   33
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txt_order_no2 
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
         TabIndex        =   29
         Top             =   600
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
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   4920
         Width           =   2535
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
         ItemData        =   "Form_Order.frx":EB26
         Left            =   -72720
         List            =   "Form_Order.frx":EB36
         TabIndex        =   5
         Top             =   5400
         Width           =   2775
      End
      Begin VB.ComboBox cmb_supplier 
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
         Left            =   -72720
         TabIndex        =   1
         Text            =   "Select"
         Top             =   1680
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker date_order 
         Height          =   375
         Left            =   -72720
         TabIndex        =   0
         Top             =   1200
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
         Format          =   73072641
         CurrentDate     =   41518
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
         Left            =   -72120
         Picture         =   "Form_Order.frx":EB5F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton btn_clear_order 
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
         Left            =   -72120
         Picture         =   "Form_Order.frx":F96C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6480
         Width           =   1575
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
         Left            =   -73920
         Picture         =   "Form_Order.frx":10645
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6480
         Width           =   1575
      End
      Begin VB.TextBox txt_total_order 
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
         Left            =   -67920
         TabIndex        =   25
         Text            =   "0"
         Top             =   720
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid dg_orders 
         Height          =   4695
         Left            =   -69600
         TabIndex        =   23
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8281
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
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4440
         Width           =   2535
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
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3960
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
         Left            =   -72720
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   3480
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
         Left            =   -72720
         TabIndex        =   14
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txt_order_no 
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
         TabIndex        =   57
         Top             =   720
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid dg_pending 
         Height          =   4335
         Left            =   5400
         TabIndex        =   42
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
      Begin MSComCtl2.DTPicker date_order2 
         Height          =   375
         Left            =   2280
         TabIndex        =   46
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
         Format          =   73072641
         CurrentDate     =   41518
      End
      Begin MSComCtl2.DTPicker date_deliver 
         Height          =   375
         Left            =   -72720
         TabIndex        =   6
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
         Format          =   73072641
         CurrentDate     =   41518
      End
      Begin MSComCtl2.DTPicker date_deliver2 
         Height          =   375
         Left            =   2280
         TabIndex        =   49
         Top             =   5760
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
         Format          =   73072641
         CurrentDate     =   41518
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
         TabIndex        =   62
         Top             =   4800
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
         TabIndex        =   61
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label26 
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
         TabIndex        =   60
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label25 
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
         TabIndex        =   59
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label13 
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
         Left            =   -68160
         TabIndex        =   58
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Order Date:"
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
         Top             =   1200
         Width           =   1815
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
         Left            =   240
         TabIndex        =   45
         Top             =   5880
         Width           =   2175
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
         Left            =   240
         TabIndex        =   43
         Top             =   5400
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
         Left            =   240
         TabIndex        =   41
         Top             =   4920
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
         Left            =   240
         TabIndex        =   40
         Top             =   4440
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
         Left            =   240
         TabIndex        =   38
         Top             =   3960
         Width           =   1575
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
         Left            =   240
         TabIndex        =   36
         Top             =   3480
         Width           =   1575
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
         Left            =   240
         TabIndex        =   34
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Person-in-charge:"
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
         TabIndex        =   32
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Order Number:"
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
         TabIndex        =   30
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lbl_deliver 
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
         TabIndex        =   27
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Total Order:"
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
         Left            =   -69600
         TabIndex        =   26
         Top             =   840
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
         Left            =   -74760
         TabIndex        =   24
         Top             =   5520
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   5040
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
         Left            =   -74760
         TabIndex        =   20
         Top             =   4560
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
         Left            =   -74760
         TabIndex        =   18
         Top             =   4080
         Width           =   1575
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
         Left            =   -74760
         TabIndex        =   17
         Top             =   3600
         Width           =   1575
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
         Left            =   -74760
         TabIndex        =   15
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Person-in-charge:"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Order Date:"
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Order Number:"
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
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Label Label9 
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
      Left            =   480
      TabIndex        =   21
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "Form_Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_orders As New ADODB.Recordset
Public rs_temp As New ADODB.Recordset
Public rs_pending As New ADODB.Recordset
Dim sql_string As String
Dim no As Integer
Dim cost, total, quan, all As Double

Private Sub btn_clear_order_Click()
    Call Form_Load
End Sub

Private Sub btn_report_Click()
    If rs_pending.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
       Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
        dr_pending_order.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
        dr_pending_order.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_pending_order.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_pending_order.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_pending_order.DataSource = rs_pending
    dr_pending_order.Show vbModal, Me
    End If
End Sub

Private Sub btn_save_Click()
    If txt_op.Text = "add" Then
        If txt_order_no.Text = "" Or cmb_supplier.Text = "" Or cmb_person.Text = "" Or txt_product_id.Text = "" Or txt_product_name.Text = "" Or txt_price.Text = "" Or txt_quantity.Text = "" Or txt_total.Text = "" Or cmb_remark.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            
             sql_string = "INSERT INTO " _
                        & "tbl_order (Order_Date,Supplier_Name," _
                        & "Person_In_Charge,Product_ID,Quantity,Total,Remark,Expected_Delivery)" _
                    & " VALUES (" _
                        & "'" & date_order.Value & "','" & cmb_supplier.Text & "','" _
                        & cmb_person.Text & "','" & txt_product_id.Text & "','" _
                        & txt_quantity.Text & "','" & txt_total.Text & "','" & cmb_remark.Text & "','" & date_deliver.Value & "')"
            Call mysql_select(rs_orders, sql_string)
            sql_string = "INSERT INTO " _
                        & "tbl_temp_order (Order_ID,Order_Date,Supplier_Name," _
                        & "Person_In_Charge,Product_ID,Quantity,Total,Remark,Expected_Delivery)" _
                    & " VALUES (" _
                        & "'" & txt_order_no.Text & "','" & date_order.Value & "','" & cmb_supplier.Text & "','" _
                        & cmb_person.Text & "','" & txt_product_id.Text & "','" _
                        & txt_quantity.Text & "','" & txt_total.Text & "','" & cmb_remark.Text & "','" & date_deliver.Value & "')"
            Call mysql_select(rs_orders, sql_string)
            MsgBox "Order added."
            txt_op.Text = "add"
            Call Form_Load
        End If
    Else
        If txt_order_no.Text = "" Or cmb_supplier.Text = "" Or cmb_person.Text = "" Or txt_product_id.Text = "" Or txt_product_name.Text = "" Or txt_price.Text = "" Or txt_quantity.Text = "" Or txt_total.Text = "" Or cmb_remark.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            sql_string = "UPDATE " _
                                & "tbl_order " _
                            & "SET " _
                                & "Order_Date = '" & date_order.Value & "', Supplier_Name = '" & cmb_supplier.Text & "'," _
                                & "Person_In_Charge = '" & cmb_person.Text & "',Product_ID = '" _
                                & txt_product_id.Text & "',Quantity = '" & txt_quantity.Text & "',Total" _
                                & " = '" & txt_total.Text & "', Remark= '" & cmb_remark.Text & "', Expected_Delivery='" & date_deliver.Value & "'" _
                            & "WHERE " _
                                & " Order_ID = " & txt_order_no.Text & ""
            Call mysql_select(rs_orders, sql_string)
          sql_string = "UPDATE " _
                                & "tbl_temp_order " _
                            & "SET " _
                                & "Order_Date = '" & date_order.Value & "', Supplier_Name = '" & cmb_supplier.Text & "'," _
                                & "Person_In_Charge = '" & cmb_person.Text & "',Product_ID = '" _
                                & txt_product_id.Text & "',Quantity = '" & txt_quantity.Text & "',Total" _
                                & " = '" & txt_total.Text & "', Remark= '" & cmb_remark.Text & "', Expected_Delivery='" & date_deliver.Value & "'" _
                            & "WHERE " _
                                & " Order_ID = '" & txt_order_no.Text & "'"
            Call mysql_select(rs_orders, sql_string)
            MsgBox "Order updated."
            txt_op.Text = "add"
            Call Form_Load
        End If
    End If
    If cmb_remark.Text = "Accepted" Then
        Dim no As Integer
        Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID='" & txt_product_id.Text & "'")
        no = val(public_rs.Fields("Quantity").Value)
        no = no + val(txt_quantity.Text)
         sql_string = "UPDATE " _
                                & "tbl_product " _
                            & "SET " _
                                & "Quantity = '" & no & "'" _
                            & "WHERE " _
                                & " Product_ID = '" & txt_product_id.Text & "'"
            Call mysql_select(rs_orders, sql_string)
    End If
    Call Form_Main.Form_Load
End Sub

Private Sub btn_save_pending_Click()
    If txt_order_no2.Text = "" Then
        MsgBox "Nothing to edit."
        Exit Sub
    Else
    
    If txt_order_no2.Text = "" Or cmb_supplier2.Text = "" Or cmb_person2.Text = "" Or txt_product_id2.Text = "" Or txt_product_name2.Text = "" Or txt_price2.Text = "" Or txt_quantity2.Text = "" Or txt_total2.Text = "" Or cmb_remark2.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            sql_string = "UPDATE " _
                                & "tbl_order " _
                            & "SET " _
                                & "Order_Date = '" & date_order2.Value & "', Supplier_Name = '" & cmb_supplier2.Text & "'," _
                                & "Person_In_Charge = '" & cmb_person2.Text & "',Product_ID = '" _
                                & txt_product_id2.Text & "',Quantity = '" & txt_quantity2.Text & "',Total" _
                                & " = '" & txt_total2.Text & "', Remark= '" & cmb_remark2.Text & "', Expected_Delivery='" & date_deliver2.Value & "'" _
                            & "WHERE " _
                                & " Order_ID = '" & txt_order_no2.Text & "'"
            Call mysql_select(rs_orders, sql_string)
            If cmb_remark2.Text = "Accepted" Then
                Dim no As Integer
                Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID ='" & txt_product_id2.Text & "'")
                no = val(public_rs.Fields("Quantity").Value)
                no = no + val(txt_quantity2.Text)
                 sql_string = "UPDATE " _
                                        & "tbl_product " _
                                    & "SET " _
                                        & "Quantity = '" & Str(no) & "'" _
                                    & "WHERE " _
                                        & "Product_ID = '" & txt_product_id2.Text & "'"
                    Call mysql_select(rs_orders, sql_string)
            End If
             MsgBox "Pending order updated."
             Call clear_all2
            Call Form_Load
        End If
    End If
    Call Form_Main.Form_Load
End Sub

Private Sub btn_search_Click()
    operation = "order"
    Call load_form(Form_Search, True)
    
End Sub

Private Sub btn_search2_Click()
     Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_order WHERE Remark = 'Pending' AND (Order_ID = '" & txt_search.Text & "' OR Order_Date = '" & txt_search.Text & "' OR Supplier_Name = '" & txt_search.Text & "' OR Person_In_Charge = '" & txt_search.Text & "' OR Product_ID = '" & txt_search.Text & "' OR Remark = '" & txt_search.Text & "' OR Expected_Delivery = '" & txt_search.Text & "') ")
            If rs_pending.RecordCount = 0 Then
                MsgBox "Record not found."
                Exit Sub
            End If
End Sub

Private Sub cmb_person_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose person-in-charge from the list."
    cmb_person.Text = ""
End Sub

Private Sub cmb_remark_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose remark from the list."
    cmb_remark.Text = ""
End Sub

Private Sub cmb_remark2_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose remark from the list."
    cmb_remark2.Text = ""
End Sub

Private Sub cmb_supplier_Click()
    Call mysql_select(public_rs, "SELECT * FROM tbl_supplier WHERE Supplier_Name='" & cmb_supplier.Text & "'")
    cmb_person.Clear
    While Not public_rs.EOF
        cmb_person.AddItem (public_rs.Fields("Representative1"))
        If public_rs.Fields("Representative2").Value <> "" Then
            
         cmb_person.AddItem (public_rs.Fields("Representative2"))
        End If
        public_rs.MoveNext
    Wend
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmb_supplier_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose supplier from the list."
    cmb_supplier.Text = ""
End Sub

Private Sub cmb_supplier2_Click()
      Call mysql_select(public_rs, "SELECT * FROM tbl_supplier WHERE Supplier_Name='" & cmb_supplier.Text & "'")
    cmb_person2.Clear
    While Not public_rs.EOF
        cmb_person2.AddItem (public_rs.Fields("Representative1"))
        If public_rs.Fields("Representative2").Value <> "" Then
            
         cmb_person.AddItem (public_rs.Fields("Representative2"))
        End If
        public_rs.MoveNext
    Wend
End Sub

Private Sub dg_orders_DblClick()
    If rs_temp.RecordCount = 0 Then
        MsgBox "No order."
        Exit Sub
    End If
    If dg_orders.DataSource Is Nothing Then
        MsgBox "No order."
        Exit Sub
    End If
    txt_op.Text = "edit"
    txt_order_no.Text = rs_temp.Fields("Order_ID").Value
    date_order.Value = rs_temp.Fields("Order_Date").Value
     cmb_supplier.Text = rs_temp.Fields("Supplier_Name").Value
      cmb_person.Text = rs_temp.Fields("Person_In_Charge").Value
      txt_product_id.Text = rs_temp.Fields("Product_ID").Value
       txt_quantity.Text = rs_temp.Fields("Quantity").Value
     txt_total.Text = rs_temp.Fields("Total").Value
      cmb_remark.Text = rs_temp.Fields("Remark").Value
       date_deliver.Value = rs_temp.Fields("Expected_Delivery").Value
       Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID = '" & rs_temp.Fields("Product_ID").Value & "' ")
       txt_product_name.Text = public_rs.Fields("Product_Name").Value
       txt_price.Text = public_rs.Fields("Cost").Value
End Sub

Private Sub dg_pending_DblClick()
    Call mysql_select(public_rs, "SELECT * FROM tbl_supplier")
    cmb_supplier2.Clear
    While Not public_rs.EOF
        cmb_supplier2.AddItem (public_rs.Fields("Supplier_Name"))
        public_rs.MoveNext
    Wend
    date_order2.Value = Now
    date_deliver2.Value = Now
     txt_op.Text = "edit"
    txt_order_no2.Text = rs_pending.Fields("Order_ID").Value
    date_order2.Value = rs_pending.Fields("Order_Date").Value
     cmb_supplier2.Text = rs_pending.Fields("Supplier_Name").Value
      cmb_person2.Text = rs_pending.Fields("Person_In_Charge").Value
      txt_product_id2.Text = rs_pending.Fields("Product_ID").Value
       txt_quantity2.Text = rs_pending.Fields("Quantity").Value
     txt_total2.Text = rs_pending.Fields("Total").Value
      cmb_remark2.Text = rs_pending.Fields("Remark").Value
       date_deliver2.Value = rs_pending.Fields("Expected_Delivery").Value
       Call mysql_select(public_rs, "SELECT * FROM tbl_product WHERE Product_ID = '" & rs_pending.Fields("Product_ID").Value & "' ")
       txt_product_name2.Text = public_rs.Fields("Product_Name").Value
       txt_price2.Text = public_rs.Fields("Cost").Value
       
End Sub

Public Sub Form_Load()
    Call clear_all
    Call mysql_select(public_rs, "SELECT * FROM tbl_order")
    If public_rs.RecordCount = 0 Then
        no = 1
    Else
        no = public_rs.RecordCount + 1
    End If
    txt_order_no.Text = Str(no)
    Call mysql_select(public_rs, "SELECT * FROM tbl_supplier")
    cmb_supplier.Clear
    While Not public_rs.EOF
        cmb_supplier.AddItem (public_rs.Fields("Supplier_Name"))
        public_rs.MoveNext
    Wend
    
    date_order.Value = Now
    date_deliver.Value = Now
    txt_total_order.Text = "0"
    txt_op.Text = "add"
    Call set_datagrid(dg_orders, rs_temp, _
                                        "SELECT * FROM tbl_temp_order")
     Call mysql_select(public_rs, "SELECT * FROM tbl_temp_order")
     all = 0
     While Not public_rs.EOF
        all = all + val(public_rs.Fields("Total").Value)
        public_rs.MoveNext
    Wend
    txt_total_order.Text = Str(all)
     Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_order WHERE Remark='Pending'")
    tab_order.Tab = 0
End Sub
Public Sub clear_all()
    txt_order_no.Text = ""
    cmb_supplier.Text = ""
    date_order.Value = Now
    cmb_person.Text = ""
    txt_product_id.Text = ""
    txt_product_name.Text = ""
    txt_quantity.Text = ""
    txt_price.Text = ""
    txt_total.Text = ""
    cmb_remark.Text = ""
    date_deliver.Value = Now
End Sub
Public Sub clear_all2()
    txt_order_no2.Text = ""
    cmb_supplier2.Text = ""
    date_order2.Value = Now
    cmb_person2.Text = ""
    txt_product_id2.Text = ""
    txt_product_name2.Text = ""
    txt_quantity2.Text = ""
    txt_price2.Text = ""
    txt_total2.Text = ""
    cmb_remark2.Text = ""
    date_deliver2.Value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Call mysql_select(public_rs, "DELETE FROM tbl_temp_order")
End Sub

Private Sub Text7_Change()

End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
   
End Sub

Private Sub txt_quantity_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_quantity.Text) Then
    txt_quantity.Text = ""
     MsgBox "Invalid input."
Else
    quan = val(txt_quantity.Text)
    cost = val(txt_price.Text)
    total = quan * cost
    txt_total.Text = Str(total)
 End If
End Sub

Private Sub txt_search_Change()
      Call set_datagrid(dg_pending, rs_pending, _
                                        "SELECT * FROM tbl_order WHERE Remark='Pending' AND (Order_ID LIKE '%" & txt_search.Text & "%' OR Order_Date LIKE '%" & txt_search.Text & "%' OR Supplier_Name LIKE '%" & txt_search.Text & "%' OR Person_In_Charge LIKE '%" & txt_search.Text & "%' OR Product_ID LIKE '%" & txt_search.Text & "%' OR Remark LIKE '%" & txt_search.Text & "%' OR Expected_Delivery LIKE '%" & txt_search.Text & "%' )")
End Sub
