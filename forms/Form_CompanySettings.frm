VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_CompanySettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Settings"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_CompanySettings.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Company Profile"
      TabPicture(0)   =   "Form_CompanySettings.frx":BF53
      Tab(0).ControlEnabled=   -1  'True
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
      Tab(0).Control(7)=   "txt_mission"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_description"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btn_save"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "btn_clear(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_vision"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_address"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_owner"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_name"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_mobile"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox txt_mobile 
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
         TabIndex        =   3
         Top             =   2880
         Width           =   7095
      End
      Begin VB.TextBox txt_name 
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
         TabIndex        =   0
         Top             =   720
         Width           =   7095
      End
      Begin VB.TextBox txt_owner 
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
         TabIndex        =   1
         Top             =   1320
         Width           =   7095
      End
      Begin VB.TextBox txt_address 
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1920
         Width           =   7095
      End
      Begin VB.TextBox txt_vision 
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   5400
         Width           =   7095
      End
      Begin VB.CommandButton btn_clear 
         Height          =   495
         Index           =   0
         Left            =   7800
         Picture         =   "Form_CompanySettings.frx":BF6F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton btn_save 
         Height          =   495
         Left            =   6000
         Picture         =   "Form_CompanySettings.frx":CC48
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6360
         Width           =   1575
      End
      Begin VB.TextBox txt_description 
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3480
         Width           =   7095
      End
      Begin VB.TextBox txt_mission 
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
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   4440
         Width           =   7095
      End
      Begin VB.Label Label7 
         Caption         =   "Mobile Number:"
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
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
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
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Owner(s):"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
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
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Mission:"
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
         TabIndex        =   11
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Vision:"
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
         TabIndex        =   10
         Top             =   5520
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form_CompanySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_company As New ADODB.Recordset
Dim sql_string As String

Private Sub btn_clear_Click(Index As Integer)
     
             txt_name.Text = ""
             txt_owner.Text = ""
               txt_address.Text = ""
                txt_description.Text = ""
                 txt_mission.Text = ""
                 txt_vision.Text = ""
                 txt_mobile.Text = ""
End Sub

Private Sub btn_save_Click()
If txt_name.Text = "" Or txt_address.Text = "" Or txt_mobile.Text = "" Then
    MsgBox "Please supply important information."
Else
    Call mysql_select(public_rs, "SELECT * FROM tbl_company")
    If public_rs.RecordCount = 0 Then
        sql_string = "INSERT INTO " _
                        & "tbl_company (Name,Owner,Mobile_Number,Address,Description," _
                        & "Mission,Vision)" _
                    & " VALUES (" _
                        & "'" & txt_name.Text & "','" & txt_owner.Text & "','" & txt_mobile.Text & "','" _
                        & txt_address.Text & "','" & txt_description.Text & "','" _
                        & txt_mission.Text & "', '" & txt_vision.Text & "')"
        Call mysql_select(Form_CompanySettings.rs_company, sql_string)
       
        MsgBox "Company Information Added."
        Form_Main.lbl_name.Caption = txt_name.Text
        Form_Main.Caption = "Sales and Inventory System for " & txt_name.Text
        Unload Me
        
    Else
        sql_string = "UPDATE " _
                            & "tbl_company " _
                        & "SET " _
                            & "Name = '" & txt_name.Text & "', Owner = '" & txt_owner.Text & "',Mobile_Number = '" & txt_mobile.Text & "', Address = '" & txt_address.Text & "',Description = '" & txt_description.Text & "',Mission = '" & txt_mission.Text & "',Vision = '" & txt_vision.Text & "' " _
                        & "WHERE " _
                            & " ID = '1'"
           Call mysql_select(Form_CompanySettings.rs_company, sql_string)
          MsgBox "Company Information Updated."
          Form_Main.lbl_name.Caption = txt_name.Text
          Form_Main.Caption = "Sales and Inventory System for " & txt_name.Text
          Unload Me
    End If
     
End If
End Sub



Private Sub Form_Load()
     Call mysql_select(public_rs, "SELECT * FROM tbl_company")
       
            company_name = public_rs.Fields("Name").Value
             txt_name.Text = company_name
             txt_owner.Text = public_rs.Fields("Owner").Value
               txt_address.Text = public_rs.Fields("Address").Value
                txt_description.Text = public_rs.Fields("Description").Value
                 txt_mission.Text = public_rs.Fields("Mission").Value
                 txt_vision.Text = public_rs.Fields("Vision").Value
                 txt_mobile.Text = public_rs.Fields("Mobile_Number").Value
                 
End Sub
Function get_File_Ext(file_name As String) As String
    file = Split(file_name, ".")
    get_File_Ext = file(UBound(file))
End Function

Private Sub txt_mission_KeyPress(KeyAscii As Integer)
  If (isNumberAscii(KeyAscii)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txt_mobile_KeyPress(KeyAscii As Integer)
  If (Not isFunctionAscii(KeyAscii) And (Not isNumberAscii(KeyAscii) Or Len(txt_mobile) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub
Private Function isFunctionAscii(ascii As Integer) As Boolean
  If (ascii = 13 Or ascii = 8 Or ascii = 32) Then
    isFunctionAscii = True
  Else
    isFunctionAscii = False
  End If
End Function

Private Sub txt_name_KeyPress(KeyAscii As Integer)
  If (isNumberAscii(KeyAscii)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Function isNumberAscii(ascii As Integer) As Boolean
  If (ascii >= 48 And ascii <= 57) Then
    isNumberAscii = True
  Else
    isNumberAscii = False
  End If
End Function

Private Sub txt_owner_KeyPress(KeyAscii As Integer)
  If (isNumberAscii(KeyAscii)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txt_vision_KeyPress(KeyAscii As Integer)
  If (isNumberAscii(KeyAscii)) Then
    KeyAscii = 0
    Beep
  End If
End Sub
