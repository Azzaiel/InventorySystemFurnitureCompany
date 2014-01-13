VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Report.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   24
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
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
      TabCaption(0)   =   "Products"
      TabPicture(0)   =   "Form_Report.frx":BF53
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Suppliers"
      TabPicture(1)   =   "Form_Report.frx":BF6F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Customers"
      TabPicture(2)   =   "Form_Report.frx":BF8B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Orders"
      TabPicture(3)   =   "Form_Report.frx":BFA7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Purchase"
      TabPicture(4)   =   "Form_Report.frx":BFC3
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "User Accounts"
      TabPicture(5)   =   "Form_Report.frx":BFDF
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User Accounts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74640
         TabIndex        =   35
         Top             =   840
         Width           =   10695
         Begin VB.CommandButton btn_search_logs 
            Height          =   495
            Left            =   7800
            Picture         =   "Form_Report.frx":BFFB
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txt_search_logs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Width           =   6135
         End
         Begin VB.CommandButton btn_report_logs 
            Height          =   495
            Left            =   8760
            Picture         =   "Form_Report.frx":CE08
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   5520
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dg_users 
            Height          =   4455
            Left            =   480
            TabIndex        =   36
            Top             =   960
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   7858
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purchase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   360
         TabIndex        =   33
         Top             =   840
         Width           =   10695
         Begin VB.CommandButton btn_search_purchase 
            Height          =   495
            Left            =   7800
            Picture         =   "Form_Report.frx":DE90
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txt_search_purchase 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   6135
         End
         Begin VB.CommandButton btn_report_purchase 
            Height          =   495
            Left            =   8760
            Picture         =   "Form_Report.frx":EC9D
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5520
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dg_purchase 
            Height          =   4455
            Left            =   480
            TabIndex        =   34
            Top             =   960
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   7858
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74640
         TabIndex        =   31
         Top             =   840
         Width           =   10695
         Begin VB.CommandButton btn_search_orders 
            Height          =   495
            Left            =   7800
            Picture         =   "Form_Report.frx":FD25
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txt_search_orders 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   0
            Top             =   240
            Width           =   6135
         End
         Begin VB.CommandButton btn_report_orders 
            Height          =   495
            Left            =   8760
            Picture         =   "Form_Report.frx":10B32
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   5520
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dg_orders 
            Height          =   4455
            Left            =   480
            TabIndex        =   32
            Top             =   960
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   7858
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74640
         TabIndex        =   29
         Top             =   840
         Width           =   10695
         Begin VB.CommandButton btn_search_customers 
            Height          =   495
            Left            =   7800
            Picture         =   "Form_Report.frx":11BBA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txt_search_customers 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   9
            Top             =   240
            Width           =   6135
         End
         Begin VB.CommandButton btn_report_customers 
            Height          =   495
            Left            =   8760
            Picture         =   "Form_Report.frx":129C7
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   5520
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dg_customers 
            Height          =   4455
            Left            =   480
            TabIndex        =   30
            Top             =   960
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   7858
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Suppliers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74640
         TabIndex        =   27
         Top             =   840
         Width           =   10695
         Begin VB.CommandButton btn_search_supplier 
            Height          =   495
            Left            =   7800
            Picture         =   "Form_Report.frx":13A4F
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txt_search_supplier 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   6135
         End
         Begin VB.CommandButton btn_report_supplier 
            Height          =   495
            Left            =   8760
            Picture         =   "Form_Report.frx":1485C
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   5520
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dg_suppliers 
            Height          =   4455
            Left            =   480
            TabIndex        =   28
            Top             =   960
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   7858
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Products"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74640
         TabIndex        =   25
         Top             =   960
         Width           =   10695
         Begin VB.OptionButton opt_all 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All"
            Height          =   255
            Left            =   3240
            TabIndex        =   17
            Top             =   840
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt_active 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Active"
            Height          =   255
            Left            =   4680
            TabIndex        =   18
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton opt_discontinue 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Phase-Out"
            Height          =   255
            Left            =   5880
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton opt_damaged 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Damaged"
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton opt_reserved 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reserved"
            Height          =   255
            Left            =   4680
            TabIndex        =   21
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton opt_pull 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pull-Out"
            Height          =   255
            Left            =   5880
            TabIndex        =   22
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton btn_report_product 
            Height          =   495
            Left            =   8760
            Picture         =   "Form_Report.frx":158E4
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   5520
            Width           =   1575
         End
         Begin VB.TextBox txt_search_product 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   15
            Top             =   240
            Width           =   6135
         End
         Begin VB.CommandButton btn_search_product 
            Height          =   495
            Left            =   7800
            Picture         =   "Form_Report.frx":1696C
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dg_products 
            Height          =   3855
            Left            =   480
            TabIndex        =   26
            Top             =   1560
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6800
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
   End
End
Attribute VB_Name = "Form_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_users As New ADODB.Recordset
Public rs_product As New ADODB.Recordset
Public rs_customer As New ADODB.Recordset
Public rs_supplier As New ADODB.Recordset
Public rs_orders As New ADODB.Recordset
Public rs_purchase As New ADODB.Recordset
Private Sub txt_search_Change()

End Sub

Private Sub btn_report_customers_Click()
    If rs_customer.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_customers.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
            dr_customers.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_customers.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_customers.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_customers.DataSource = rs_customer
    dr_customers.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_logs_Click()
    If rs_users.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_accounts.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
            dr_accounts.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_accounts.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_accounts.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_accounts.DataSource = rs_users
    dr_accounts.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_orders_Click()
    If rs_orders.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
      Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
             dr_order.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
            dr_order.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_order.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_order.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
       
         
        Dim totalUnitPrice As Long
        Dim totalCost As Long
    
        totalCost = 0
        rs_product.MoveFirst
        While Not rs_orders.EOF
          totalCost = totalCost + val(rs_orders!total)
          rs_orders.MoveNext
        Wend
        dr_order.Sections(2).Controls("lblTotalCost").Caption = totalCost
        
        rs_product.MoveFirst
        
        Set dr_order.DataSource = rs_orders
        dr_order.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_product_Click()
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

Private Sub btn_report_purchase_Click()
    If rs_purchase.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
       Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_purchase.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
            dr_purchase.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_purchase.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_purchase.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
    
        Dim totalUnitPrice As Long

        totalUnitPrice = 0
        totalCost = 0
        rs_purchase.MoveFirst
        While Not rs_purchase.EOF
          totalUnitPrice = totalUnitPrice + val(rs_purchase!total)
          rs_purchase.MoveNext
        Wend
        
        dr_purchase.Sections(2).Controls("lblTotalUnitPrice").Caption = totalUnitPrice
        
        rs_product.MoveFirst
        
        Set dr_purchase.DataSource = rs_purchase
        dr_purchase.Show vbModal, Me
    End If
End Sub

Private Sub btn_report_supplier_Click()
    If rs_supplier.RecordCount = 0 Then
        MsgBox "No record to display."
        Exit Sub
    Else
     Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
            dr_suppliers.Sections(2).Controls("lbl_date").Caption = Format(Now, "MMMM, dd yyyy h:n AM/PM")
            dr_suppliers.Sections(2).Controls("lbl_name").Caption = public_rs.Fields("Name").Value
        dr_suppliers.Sections(2).Controls("lbl_address").Caption = public_rs.Fields("Address").Value
        dr_suppliers.Sections(2).Controls("lbl_mobile").Caption = public_rs.Fields("Mobile_Number").Value
         Set dr_suppliers.DataSource = rs_supplier
    dr_suppliers.Show vbModal, Me
    End If
End Sub

Private Sub btn_search_customers_Click()
     Call set_datagrid(dg_customers, rs_customer, _
                                        "SELECT * FROM tbl_customer WHERE Customer_ID = '" & txt_search_customers.Text & "' OR Customer_Name = '" & txt_search_customers.Text & "' OR Representative1 = '" & txt_search_customers.Text & "' OR Representative2 = '" & txt_search_customers.Text & "'")
    If rs_customer.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
End Sub

Private Sub btn_search_logs_Click()
       Call set_datagrid(dg_users, rs_users, _
                                        "SELECT Username,Usertype,ID,Lastname,Firstname,Middlename,MobileNumber, Address FROM tbl_users WHERE Username = '" & txt_search_logs.Text & "'OR Usertype = '" & txt_search_logs.Text & "' OR ID = '" & txt_search_logs.Text & "' OR Lastname = '" & txt_search_logs.Text & "' OR Firstname = '" & txt_search_logs.Text & "'")
    If rs_users.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
End Sub

Private Sub btn_search_orders_Click()
     Call set_datagrid(dg_orders, rs_orders, _
                                        "SELECT * FROM tbl_order WHERE Order_ID = '" & txt_search_orders.Text & "' OR Order_Date = '" & txt_search_orders.Text & "' OR Supplier_Name = '" & txt_search_orders.Text & "' OR Person_In_Charge = '" & txt_search_orders.Text & "' OR Product_ID = '" & txt_search_orders.Text & "' OR Remark = '" & txt_search_orders.Text & "' OR Expected_Delivery = '" & txt_search_orders.Text & "' ")
            If rs_orders.RecordCount = 0 Then
                MsgBox "Record not found."
                Exit Sub
            End If
        
     Call formatDataGrid
End Sub

Private Sub btn_search_product_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Product_ID = '" & txt_search_product.Text & "' OR Product_Name = '" & txt_search_product.Text & "' OR Category = '" & txt_search_product.Text & "' OR Brand = '" & txt_search_product.Text & "' OR Initial_Supplier = '" & txt_search_product.Text & "' OR Remark = '" & txt_search_product.Text & "'")
    Call formatDataGrid
    If rs_product.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
End Sub

Private Sub btn_search_purchase_Click()
     Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT * FROM tbl_purchase WHERE Purchase_ID = '" & txt_search_purchase.Text & "' OR Purchase_Date = '" & txt_search_purchase.Text & "' OR Customer_Name = '" & txt_search_purchase.Text & "' OR Person_In_Charge = '" & txt_search_purchase.Text & "' OR Product_ID = '" & txt_search_purchase.Text & "' OR Remark = '" & txt_search_purchase.Text & "' OR Expected_Delivery = '" & txt_search_purchase.Text & "'")
                
                If rs_purchase.RecordCount = 0 Then
                    MsgBox "Record not found."
                End If
End Sub

Private Sub btn_search_supplier_Click()
     Call set_datagrid(dg_suppliers, rs_supplier, _
                                        "SELECT * FROM tbl_supplier WHERE Supplier_ID = '" & txt_search_supplier.Text & "' OR Supplier_Name = '" & txt_search_supplier.Text & "' OR Representative1 = '" & txt_search_supplier.Text & "' OR Representative2 = '" & txt_search_supplier.Text & "'")
    If rs_supplier.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
    
    Call formatDataGrid
End Sub

Private Sub Form_Load()
         Call set_datagrid(dg_users, rs_users, _
                                        "SELECT Username,Usertype,ID,Lastname,Firstname,Middlename,MobileNumber, Address FROM tbl_users")
         Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product")
        
         Call set_datagrid(dg_customers, rs_customer, _
                                        "SELECT * FROM tbl_customer")
        
         Call set_datagrid(dg_suppliers, rs_supplier, _
                                        "SELECT * FROM tbl_supplier")
                        
        Call set_datagrid(dg_orders, rs_orders, _
                                        "SELECT * FROM tbl_order")
                                        
          Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT * FROM tbl_purchase")

        
  
     Call formatDataGrid
  


End Sub

Public Sub formatDataGrid()
    dg_orders.Columns(6).NumberFormat = "##,##0.00"
    dg_purchase.Columns(6).NumberFormat = "##,##0.00"
    dg_products.Columns(7).NumberFormat = "##,##0.00"
    dg_products.Columns(9).NumberFormat = "##,##0.00"
End Sub

Private Sub Option3_Click()

End Sub

Private Sub opt_active_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Active'")
    txt_search_product.Text = ""
End Sub

Private Sub opt_all_Click()
     Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product ")
    txt_search_product.Text = ""
End Sub

Private Sub opt_damaged_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Damaged'")
    txt_search_product.Text = ""
End Sub

Private Sub opt_discontinue_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Phase-Out'")
    txt_search_product.Text = ""
End Sub

Private Sub opt_pull_Click()
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Pull-Out'")
    txt_search_product.Text = ""
End Sub

Private Sub opt_reserved_Click()
       Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Remark = 'Reserved'")
    txt_search_product.Text = ""
End Sub



Private Sub txt_search_customers_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_customers, rs_customer, _
                                        "SELECT * FROM tbl_customer WHERE Customer_ID LIKE '%" & txt_search_customers.Text & "%' OR Customer_Name LIKE '%" & txt_search_customers.Text & "%' OR Representative1 LIKE '%" & txt_search_customers.Text & "%' OR Representative2 LIKE '%" & txt_search_customers.Text & "%'")
End Sub

Private Sub txt_search_logs_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_users, rs_users, _
                                        "SELECT Username,Usertype,ID,Lastname,Firstname,Middlename,MobileNumber, Address FROM tbl_users WHERE Username LIKE '%" & txt_search_logs.Text & "%'OR Usertype LIKE '%" & txt_search_logs.Text & "%' OR ID LIKE '%" & txt_search_logs.Text & "%' OR Lastname LIKE '%" & txt_search_logs.Text & "%' OR Firstname LIKE '%" & txt_search_logs.Text & "%'")
End Sub

Private Sub txt_search_orders_KeyUp(KeyCode As Integer, Shift As Integer)
    Call set_datagrid(dg_orders, rs_orders, _
                                        "SELECT * FROM tbl_order WHERE Order_ID LIKE '%" & txt_search_orders.Text & "%' OR Order_Date LIKE '%" & txt_search_orders.Text & "%' OR Supplier_Name LIKE '%" & txt_search_orders.Text & "%' OR Person_In_Charge LIKE '%" & txt_search_orders.Text & "%' OR Product_ID LIKE '%" & txt_search_orders.Text & "%' OR Remark LIKE '%" & txt_search_orders.Text & "%' OR Expected_Delivery LIKE '%" & txt_search_orders.Text & "%'")
    Call formatDataGrid
End Sub

Private Sub txt_search_product_KeyUp(KeyCode As Integer, Shift As Integer)
      Call set_datagrid(dg_products, rs_product, _
                                        "SELECT * FROM tbl_product WHERE Product_ID LIKE '%" & txt_search_product.Text & "%' OR Product_Name LIKE '%" & txt_search_product.Text & "%' OR Category LIKE '%" & txt_search_product.Text & "%' OR Brand LIKE '%" & txt_search_product.Text & "%' OR Initial_Supplier LIKE '%" & txt_search_product.Text & "%' OR Remark LIKE '%" & txt_search_product.Text & "%'")
     Call formatDataGrid
End Sub

Private Sub txt_search_purchase_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_purchase, rs_purchase, _
                                        "SELECT * FROM tbl_purchase WHERE Purchase_ID LIKE '%" & txt_search_purchase.Text & "%' OR Purchase_Date LIKE '%" & txt_search_purchase.Text & "%' OR Customer_Name LIKE '%" & txt_search_purchase.Text & "%' OR Person_In_Charge LIKE '%" & txt_search_purchase.Text & "%' OR Product_ID LIKE '%" & txt_search_purchase.Text & "%' OR Remark LIKE '%" & txt_search_purchase.Text & "%' OR Expected_Delivery LIKE '%" & txt_search_purchase.Text & "%'")
End Sub

Private Sub txt_search_supplier_KeyUp(KeyCode As Integer, Shift As Integer)
      Call set_datagrid(dg_suppliers, rs_supplier, _
                                        "SELECT * FROM tbl_supplier WHERE Supplier_ID LIKE '%" & txt_search_supplier.Text & "%' OR Supplier_Name LIKE '%" & txt_search_supplier.Text & "%' OR Representative1 LIKE '%" & txt_search_supplier.Text & "%' OR Representative2 LIKE '%" & txt_search_supplier.Text & "%'")
    Call formatDataGrid
End Sub
