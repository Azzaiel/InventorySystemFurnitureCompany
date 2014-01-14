VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_Product 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Information"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Product.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   14460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton btn_report 
         Height          =   495
         Left            =   2400
         Picture         =   "Form_Product.frx":A636
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7800
         Width           =   1575
      End
      Begin VB.OptionButton opt_pull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pull-Out"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton opt_reserved 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reserved"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton opt_damaged 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Damaged"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton btn_search 
         Height          =   495
         Left            =   1200
         Picture         =   "Form_Product.frx":B6BE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opt_discontinue 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phase-Out"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton opt_active 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Active"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton opt_all 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid dg_products 
         Height          =   5415
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9551
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
   Begin TabDlg.SSTab tab_product 
      Height          =   8415
      Left            =   4320
      TabIndex        =   24
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   14843
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Product Information"
      TabPicture(0)   =   "Form_Product.frx":C4CB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label27"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label13"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "btn_clear"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "btn_save"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_description"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_brand"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_name"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_id"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_cost"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_quantity"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_price"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txt_critical"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmb_category"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmb_supplier"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmb_remark"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt_op"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt_oldid"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt_supplierID"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Order Information"
      TabPicture(1)   =   "Form_Product.frx":C4E7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_search_order"
      Tab(1).Control(1)=   "btn_search_order"
      Tab(1).Control(2)=   "btn_report_order"
      Tab(1).Control(3)=   "dg_order"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Purchase Information"
      TabPicture(2)   =   "Form_Product.frx":C503
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "dg_purchase"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "btn_report_purchase"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btn_search_purchase"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txt_search_purchase"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.TextBox txt_supplierID 
         Height          =   375
         Left            =   -67560
         TabIndex        =   50
         Top             =   2760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_oldid 
         Height          =   375
         Left            =   -67320
         TabIndex        =   49
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_op 
         Height          =   375
         Left            =   -67320
         TabIndex        =   47
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_search_purchase 
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
         Left            =   1200
         TabIndex        =   44
         Top             =   600
         Width           =   5655
      End
      Begin VB.CommandButton btn_search_purchase 
         Height          =   495
         Left            =   7080
         Picture         =   "Form_Product.frx":C51F
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton btn_report_purchase 
         Height          =   495
         Left            =   8160
         Picture         =   "Form_Product.frx":D32C
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   7680
         Width           =   1575
      End
      Begin VB.TextBox txt_search_order 
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
         Left            =   -73800
         TabIndex        =   29
         Top             =   600
         Width           =   5655
      End
      Begin VB.CommandButton btn_search_order 
         Height          =   495
         Left            =   -67920
         Picture         =   "Form_Product.frx":E3B4
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton btn_report_order 
         Height          =   495
         Left            =   -66840
         Picture         =   "Form_Product.frx":F1C1
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7680
         Width           =   1575
      End
      Begin VB.ComboBox cmb_remark 
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
         ItemData        =   "Form_Product.frx":10249
         Left            =   -72960
         List            =   "Form_Product.frx":1025C
         TabIndex        =   19
         Top             =   6000
         Width           =   4815
      End
      Begin VB.ComboBox cmb_supplier 
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
         Left            =   -72960
         TabIndex        =   14
         Top             =   3600
         Width           =   4815
      End
      Begin VB.ComboBox cmb_category 
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
         Left            =   -72960
         TabIndex        =   11
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txt_critical 
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
         Left            =   -72960
         TabIndex        =   18
         Top             =   5520
         Width           =   4815
      End
      Begin VB.TextBox txt_price 
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
         TabIndex        =   17
         Top             =   5040
         Width           =   4575
      End
      Begin VB.TextBox txt_quantity 
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
         Left            =   -72960
         TabIndex        =   16
         Top             =   4560
         Width           =   4815
      End
      Begin VB.TextBox txt_cost 
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
         TabIndex        =   15
         Top             =   4080
         Width           =   4575
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
         Left            =   -72960
         TabIndex        =   9
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txt_name 
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
         Left            =   -72960
         TabIndex        =   10
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox txt_brand 
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
         Left            =   -72960
         TabIndex        =   12
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txt_description 
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
         Height          =   735
         Left            =   -72960
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   2760
         Width           =   4815
      End
      Begin VB.CommandButton btn_save 
         Height          =   495
         Left            =   -72960
         Picture         =   "Form_Product.frx":10290
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CommandButton btn_clear 
         Height          =   495
         Left            =   -71160
         Picture         =   "Form_Product.frx":10F5D
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6600
         Width           =   1575
      End
      Begin VB.Frame Frame1 
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
         Height          =   975
         Left            =   -68880
         TabIndex        =   27
         Top             =   7320
         Width           =   3735
         Begin VB.CommandButton btn_edit 
            Height          =   495
            Left            =   2040
            Picture         =   "Form_Product.frx":11C36
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton btn_add 
            Height          =   495
            Left            =   240
            Picture         =   "Form_Product.frx":128E0
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid dg_order 
         Height          =   6255
         Left            =   -74760
         TabIndex        =   42
         Top             =   1200
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11033
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid dg_purchase 
         Height          =   6255
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11033
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         Left            =   -72960
         TabIndex        =   52
         Top             =   5040
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
         Left            =   -72960
         TabIndex        =   51
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Click to Add/Update Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -67920
         TabIndex        =   45
         Top             =   1920
         Width           =   2775
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
         TabIndex        =   41
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Crirtical Point:"
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
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Unit Price:"
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
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label8 
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
         TabIndex        =   38
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Cost:"
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
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         TabIndex        =   36
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         TabIndex        =   35
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Category:"
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
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Brand:"
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
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
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
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Initial Supplier:"
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
         TabIndex        =   28
         Top             =   3600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_category As New ADODB.Recordset
Public rs_product As New ADODB.Recordset
Public rs_order As New ADODB.Recordset
Public rs_purchase As New ADODB.Recordset
Dim sql_string As String
Private Sub txt_representative1_Change()

End Sub

Private Sub txt_representative2_Change()

End Sub

Private Sub btn_add_Click()
    txt_op.Text = "add"
    txt_oldid.Text = ""
    Call clear_all
    Call enable_all
    btn_edit.Enabled = False
    cmb_remark.Text = "Active"
End Sub

Private Sub btn_clear_Click()
    txt_op.Text = "add"
    txt_oldid.Text = ""
    Call clear_all
    btn_edit.Enabled = False
End Sub

Private Sub btn_edit_Click()
    Call enable_all
    txt_op.Text = "edit"
End Sub

Private Sub btn_report_Click()
    If rs_product.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
        Call mysql_select(public_rs, "SELECT * FROM tbl_company")
        company_name = public_rs.Fields("Name").Value
        dr_products.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
        dr_products.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_products.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_products.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
        Set dr_products.DataSource = rs_product
        
        Dim totalUnitPrice As Long
        Dim totalCost As Long
        
        totalUnitPrice = 0
        totalCost = 0
        rs_product.MoveFirst
        While Not rs_product.EOF
          totalUnitPrice = totalUnitPrice + val(rs_product!UNIT_PRICE)
          totalCost = totalCost + val(rs_product!cost)
          rs_product.MoveNext
        Wend
        
        dr_products.Sections(2).Controls("lblTotalUnitPrice").Caption = totalUnitPrice
        dr_products.Sections(2).Controls("lblTotalCost").Caption = totalCost
        
        rs_product.MoveFirst
        
        dr_products.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_order_Click()
If dg_order.DataSource Is Nothing Then
    MsgBox "No selected product."
    Exit Sub
Else
    If rs_order.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
       Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_product_order.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_product_order.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_product_order.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
        dr_product_order.Sections(2).Controls("lbl_prod_id").Caption = "Product_ID: " & rs_product.Fields("Product_ID").Value
         Set dr_product_order.DataSource = rs_order
    dr_product_order.Show vbModal, Me
    End If
End If
End Sub

Private Sub btn_report_purchase_Click()
If dg_purchase.DataSource Is Nothing Then
    MsgBox "No selected product."
    Exit Sub
Else
    
    If rs_purchase.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
        
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_product_purchase.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_product_purchase.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_product_purchase.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
        dr_product_purchase.Sections(2).Controls("lbl_prod_id").Caption = "Product_ID: " & rs_product.Fields("Product_ID").Value
         Set dr_product_purchase.DataSource = rs_purchase
    dr_product_purchase.Show vbModal, Me
End If
End If
End Sub


Private Sub btn_save_Click()
    If txt_id.Enabled = False Then
        MsgBox "Nothing to edit."
        Exit Sub
    End If
    
    If txt_op.Text = "add" Then
        If txt_id.Text = "" Or txt_name.Text = "" Or cmb_category.Text = "" Or txt_brand.Text = "" Or cmb_supplier.Text = "" Or txt_description.Text = "" Or txt_cost.Text = "" Or txt_quantity.Text = "" Or txt_price.Text = "" Or txt_critical.Text = "" Or cmb_remark.Text = "" Then
            MsgBox "Please complete all fields."
            Exit Sub
        Else
            If is_duplicate("tbl_product", "Product_ID", txt_id.Text) Then
                MsgBox "Product ID exists."
                Exit Sub
            End If
            sql_string = "INSERT INTO " _
                        & "tbl_product (Product_ID,Product_Name,Category,Brand," _
                        & "Description,Initial_Supplier,Cost,Quantity," _
                        & "Unit_Price,Critical_Point,Remark)" _
                    & " VALUES (" _
                        & "'" & txt_id.Text & "','" & txt_name.Text & "','" _
                        & cmb_category.Text & "','" & txt_brand.Text & "','" _
                        & txt_description.Text & "','" & cmb_supplier.Text & "','" _
                        & txt_cost.Text & "','" _
                        & txt_quantity.Text & "', '" & txt_price.Text & "','" & txt_critical.Text & "','" & cmb_remark.Text & "')"
            Call mysql_select(Form_Product.rs_product, sql_string)
            MsgBox "Product added."
            Call Form_Load
        End If
    Else
        If txt_id.Text <> txt_oldid.Text Then
            If txt_id.Text = "" Or txt_name.Text = "" Or cmb_category.Text = "" Or txt_brand.Text = "" Or cmb_supplier.Text = "" Or txt_description.Text = "" Or txt_cost.Text = "" Or txt_quantity.Text = "" Or txt_price.Text = "" Or txt_critical.Text = "" Or cmb_remark.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                If is_duplicate("tbl_product", "Product_ID", txt_id.Text) Then
                    MsgBox "Product ID exists."
                    Exit Sub
                End If
                sql_string = "UPDATE " _
                                & "tbl_product " _
                            & "SET " _
                                & "Product_ID = '" & txt_id.Text & "', Product_Name = '" & txt_name.Text & "'," _
                                & "Category = '" & cmb_category.Text & "',Brand = '" _
                                & txt_brand.Text & "',Description = '" & txt_description.Text & "',Initial_Supplier" _
                                & " = '" & cmb_supplier.Text & "'" _
                                & ",Cost = '" & txt_cost.Text _
                                & "',Quantity = '" & txt_quantity.Text & "',Unit_Price ='" & txt_price.Text & "', Critical_Point ='" & txt_critical.Text & "',Remark ='" & cmb_remark.Text & "'" _
                            & "WHERE " _
                                & " Product_ID = '" & txt_oldid.Text & "'"
                Call mysql_select(Form_Product.rs_product, sql_string)
                MsgBox "Product updated."
                Call Form_Load
            End If
        Else
            If txt_id.Text = "" Or txt_name.Text = "" Or cmb_category.Text = "" Or txt_brand.Text = "" Or cmb_supplier.Text = "" Or txt_description.Text = "" Or txt_cost.Text = "" Or txt_quantity.Text = "" Or txt_price.Text = "" Or txt_critical.Text = "" Or cmb_remark.Text = "" Then
                MsgBox "Please complete all fields."
                Exit Sub
            Else
                sql_string = "UPDATE " _
                                & "tbl_product " _
                            & "SET " _
                                & "Product_ID = '" & txt_id.Text & "', Product_Name = '" & txt_name.Text & "'," _
                                & "Category = '" & cmb_category.Text & "',Brand = '" _
                                & txt_brand.Text & "',Description = '" & txt_description.Text & "',Initial_Supplier" _
                                & " = '" & cmb_supplier.Text & "'" _
                                & ",Cost = '" & txt_cost.Text _
                                & "',Quantity = '" & txt_quantity.Text & "',Unit_Price ='" & txt_price.Text & "', Critical_Point ='" & txt_critical.Text & "',Remark ='" & cmb_remark.Text & "'" _
                            & "WHERE " _
                                & " Product_ID = '" & txt_oldid.Text & "'"
                Call mysql_select(Form_Product.rs_product, sql_string)
                MsgBox "Product updated."
                Call Form_Load
            End If
        End If
    End If
End Sub

Private Sub btn_search_Click()
     opt_all.Value = False
    opt_active.Value = False
    opt_discontinue.Value = False
    Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Product_ID = '" & txt_search.Text & "' OR Product_Name = '" & txt_search.Text & "' OR Category = '" & txt_search.Text & "' OR Brand = '" & txt_search.Text & "' OR Initial_Supplier = '" & txt_search.Text & "' OR Remark = '" & txt_search.Text & "'")
    Call formatProductDataGrid
    If rs_product.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
    
End Sub

Private Sub btn_search_order_Click()
    Call set_datagrid(dg_order, rs_order, _
                                        "SELECT *  FROM tbl_order WHERE (Product_ID='" & rs_product.Fields("Product_ID").Value & "') AND (Order_ID = '" & txt_search_order.Text & "' OR Order_Date = '" & txt_search_order.Text & "' OR Supplier_Name = '" & txt_search_order.Text & "' OR Person_In_Charge = '" & txt_search_order.Text & "' OR Expected_Delivery = '" & txt_search_order.Text & "' OR Remark = '" & txt_search_order.Text & "')")
        If rs_order.RecordCount = 0 Then
            MsgBox "Record not found."
        End If
End Sub

Private Sub btn_search_purchase_Click()
     Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT *  FROM tbl_purchase WHERE (Product_ID='" & rs_product.Fields("Product_ID").Value & "') AND (Purchase_ID = '" & txt_search_purchase.Text & "' OR Purchase_Date = '" & txt_search_purchase.Text & "' OR Customer_Name = '" & txt_search_purchase.Text & "' OR Person_In_Charge = '" & txt_search_purchase.Text & "' OR Expected_Delivery = '" & txt_search_purchase.Text & "' OR Remark = '" & txt_search_purchase.Text & "')")
                
     Call formatPurchaseDataGrid
                If rs_purchase.RecordCount = 0 Then
                    MsgBox "Record not found."
                    Exit Sub
                End If
End Sub

Private Sub cmb_category_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose product category from the list."
    cmb_category.Text = ""
End Sub

Private Sub cmb_remark_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please choose remark from the list."
    cmb_remark.Text = ""
End Sub

Private Sub formatPurchaseDataGrid()
  With dg_purchase
    .Columns(6).NumberFormat = "###,###.00"
  End With
End Sub

Private Sub dg_products_Click()
      txt_op.Text = "edit"
    btn_edit.Enabled = True
    Call disable_all
    txt_id.Text = rs_product.Fields("Product_ID").Value
    txt_oldid.Text = rs_product.Fields("Product_ID").Value
    txt_name.Text = rs_product.Fields("Product_Name").Value
    cmb_category.Text = rs_product.Fields("Category").Value
    txt_brand.Text = rs_product.Fields("Brand").Value
    txt_description.Text = rs_product.Fields("Description").Value
    cmb_supplier.Text = rs_product.Fields("Initial_Supplier").Value
    txt_cost.Text = rs_product.Fields("Cost").Value
    txt_quantity.Text = rs_product.Fields("Quantity").Value
    txt_price.Text = rs_product.Fields("Unit_Price").Value
    txt_critical.Text = rs_product.Fields("Critical_Point").Value
    cmb_remark.Text = rs_product.Fields("Remark").Value
     Call set_datagrid(dg_order, rs_order, _
                                        "SELECT *  FROM tbl_order WHERE Product_ID='" & rs_product.Fields("Product_ID").Value & "'")
     Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT *  FROM tbl_purchase WHERE Product_ID='" & rs_product.Fields("Product_ID").Value & "'")
    Call formatPurchaseDataGrid
End Sub

Private Sub Form_Load()
     Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product")
    Call formatProductDataGrid
                                        
    Call mysql_select(public_rs, "SELECT * FROM tbl_category")
    cmb_category.Clear
    While Not public_rs.EOF
        cmb_category.AddItem (public_rs.Fields("Category_Name"))
        public_rs.MoveNext
    Wend
    Call mysql_select(public_rs, "SELECT * FROM tbl_supplier")
    cmb_supplier.Clear
    While Not public_rs.EOF
        cmb_supplier.AddItem (public_rs.Fields("Supplier_Name"))
        public_rs.MoveNext
    Wend
    txt_op.Text = "add"
    txt_oldid.Text = ""
    btn_edit.Enabled = False
    Call clear_all
    Call disable_all
    tab_product.Tab = 0
End Sub
Public Sub enable_all()
    txt_id.Enabled = True
    txt_name.Enabled = True
    cmb_category.Enabled = True
    txt_brand.Enabled = True
    txt_description.Enabled = True
    cmb_supplier.Enabled = True
    txt_cost.Enabled = True
    txt_quantity.Enabled = True
    txt_price.Enabled = True
    txt_critical.Enabled = True
    cmb_remark.Enabled = True
End Sub
Public Sub disable_all()
    txt_id.Enabled = False
    txt_name.Enabled = False
    cmb_category.Enabled = False
    txt_brand.Enabled = False
    txt_description.Enabled = False
    cmb_supplier.Enabled = False
    txt_cost.Enabled = False
    txt_quantity.Enabled = False
    txt_price.Enabled = False
    txt_critical.Enabled = False
    cmb_remark.Enabled = False
End Sub
Public Sub clear_all()
    txt_id.Text = ""
    txt_name.Text = ""
    cmb_category.Text = ""
    txt_brand.Text = ""
    txt_description.Text = ""
    cmb_supplier.Text = ""
    txt_cost.Text = ""
    txt_quantity.Text = ""
    txt_price.Text = ""
    txt_critical.Text = ""
    cmb_remark.Text = ""
End Sub

Private Sub formatProductDataGrid()
  With dg_products
    .Columns(7).NumberFormat = "###,###.00"
    .Columns(9).NumberFormat = "###,###.00"
  End With
End Sub

Private Sub Label12_Click()
     Call load_form(Form_Category, True)
    Call mysql_select(public_rs, "SELECT * FROM tbl_category")
    cmb_category.Clear
    While Not public_rs.EOF
        cmb_category.AddItem (public_rs.Fields("Category_Name"))
        public_rs.MoveNext
    Wend
End Sub

Private Sub opt_active_Click()
     Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Active'")
    Call formatProductDataGrid
    txt_search.Text = ""
End Sub

Private Sub opt_all_Click()
     Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product ")
     Call formatProductDataGrid
    txt_search.Text = ""
End Sub

Private Sub opt_damaged_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Damaged'")
    Call formatProductDataGrid
    txt_search.Text = ""
End Sub

Private Sub opt_discontinue_Click()
    Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Phase-Out'")
    Call formatProductDataGrid
    txt_search.Text = ""
End Sub

Private Sub opt_pull_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Pull-Out'")
    Call formatProductDataGrid
    txt_search.Text = ""
End Sub

Private Sub opt_reserved_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Reserved'")
    Call formatProductDataGrid
    txt_search.Text = ""
End Sub

Private Sub txt_cost_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_cost.Text) Then
        txt_cost.Text = ""
        MsgBox "Invalid input."
        Exit Sub
    End If
End Sub

Private Sub txt_critical_KeyUp(KeyCode As Integer, Shift As Integer)
     If Not IsNumeric(txt_critical.Text) Then
        txt_critical.Text = ""
        MsgBox "Invalid input."
        Exit Sub
    End If
End Sub

Private Sub txt_price_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_price.Text) Then
        txt_price.Text = ""
        MsgBox "Invalid input."
        Exit Sub
    End If
End Sub

Private Sub txt_quantity_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_quantity.Text) Then
        MsgBox "Invalid input."
        txt_quantity.Text = ""
        Exit Sub
    End If
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
    opt_all.Value = False
    opt_active.Value = False
    opt_discontinue.Value = False
    Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Product_ID LIKE '%" & txt_search.Text & "%' OR Product_Name LIKE '%" & txt_search.Text & "%' OR Category LIKE '%" & txt_search.Text & "%' OR Brand LIKE '%" & txt_search.Text & "%' OR Initial_Supplier LIKE '%" & txt_search.Text & "%' OR Remark LIKE '%" & txt_search.Text & "%'")
    Call formatProductDataGrid
End Sub

Private Sub txt_search_order_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_order, rs_order, _
                                        "SELECT *  FROM tbl_order WHERE (Product_ID='" & rs_product.Fields("Product_ID").Value & "') AND (Order_ID LIKE '%" & txt_search_order.Text & "%' OR Order_Date LIKE '%" & txt_search_order.Text & "%' OR Supplier_Name LIKE '%" & txt_search_order.Text & "%' OR Person_In_Charge LIKE '%" & txt_search_order.Text & "%' OR Expected_Delivery LIKE '%" & txt_search_order.Text & "%' OR Remark LIKE '%" & txt_search_order.Text & "%')")
End Sub

Private Sub txt_search_purchase_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT *  FROM tbl_purchase WHERE (Product_ID='" & rs_product.Fields("Product_ID").Value & "') AND (Purchase_ID LIKE '%" & txt_search_purchase.Text & "%' OR Purchase_Date LIKE '%" & txt_search_purchase.Text & "%' OR Customer_Name LIKE '%" & txt_search_purchase.Text & "%' OR Person_In_Charge LIKE '%" & txt_search_purchase.Text & "%' OR Expected_Delivery LIKE '%" & txt_search_purchase.Text & "%' OR Remark LIKE '%" & txt_search_purchase.Text & "%')")
     Call formatPurchaseDataGrid
End Sub
