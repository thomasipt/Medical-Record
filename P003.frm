VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT DATA PASIEN"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   6405
   ClientWidth     =   8925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1710
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "P003.frx":0000
      Top             =   915
      Width           =   4920
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1710
      TabIndex        =   1
      Text            =   "2"
      Top             =   510
      Width           =   4920
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1710
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "1"
      Top             =   90
      Width           =   2040
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "P003.frx":0005
      Left            =   1710
      List            =   "P003.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2130
      Width           =   1275
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6705
      TabIndex        =   3
      Text            =   "4"
      Top             =   1410
      Width           =   2040
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1710
      TabIndex        =   6
      Text            =   "5"
      Top             =   2940
      Width           =   870
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3060
      TabIndex        =   7
      Text            =   "6"
      Top             =   2940
      Width           =   870
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2130
      Width           =   1185
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2535
      Width           =   2040
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2970
      Width           =   2670
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3390
      Width           =   825
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6720
      TabIndex        =   17
      Top             =   6810
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   255
      TabIndex        =   16
      Top             =   6795
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7830
      OleObjectBlob   =   "P003.frx":0009
      Top             =   285
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   1710
      TabIndex        =   5
      Top             =   2475
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   20381697
      CurrentDate     =   39286
      MinDate         =   2
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P003.frx":023D
      TabIndex        =   23
      Top             =   165
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P003.frx":0298
      TabIndex        =   24
      Top             =   540
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P003.frx":02F7
      TabIndex        =   25
      Top             =   945
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   195
      Left            =   150
      OleObjectBlob   =   "P003.frx":035A
      TabIndex        =   26
      Top             =   2595
      Width           =   1185
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P003.frx":03C5
      TabIndex        =   27
      Top             =   2175
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   225
      Left            =   6705
      OleObjectBlob   =   "P003.frx":0436
      TabIndex        =   28
      Top             =   1140
      Width           =   2040
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P003.frx":049B
      TabIndex        =   29
      Top             =   3015
      Width           =   4110
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P003.frx":057C
      TabIndex        =   30
      Top             =   2175
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P003.frx":05DD
      TabIndex        =   31
      Top             =   2580
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P003.frx":0646
      TabIndex        =   32
      Top             =   3015
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P003.frx":06B1
      TabIndex        =   33
      Top             =   3435
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   -45
      TabIndex        =   35
      Top             =   1860
      Width           =   9015
   End
   Begin VB.PictureBox Picture2 
      Height          =   1365
      Left            =   -90
      ScaleHeight     =   1305
      ScaleWidth      =   9045
      TabIndex        =   34
      Top             =   6675
      Width           =   9105
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATA KELUARGA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   30
      TabIndex        =   18
      Top             =   4065
      Width           =   8850
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         TabIndex        =   12
         Text            =   "7"
         Top             =   450
         Width           =   4920
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1695
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "P003.frx":071A
         Top             =   855
         Width           =   4920
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6705
         TabIndex        =   14
         Text            =   "9"
         Top             =   1350
         Width           =   2040
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         TabIndex        =   15
         Text            =   "10"
         Top             =   1935
         Width           =   4920
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "P003.frx":071F
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "P003.frx":077E
         TabIndex        =   20
         Top             =   885
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   225
         Left            =   6705
         OleObjectBlob   =   "P003.frx":07F3
         TabIndex        =   21
         Top             =   1080
         Width           =   2040
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   225
         Left            =   135
         OleObjectBlob   =   "P003.frx":0858
         TabIndex        =   22
         Top             =   1965
         Width           =   1095
      End
   End
End
Attribute VB_Name = "P003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RKTG2, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SKTG2, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub cmdCLOSE_Click()
Unload Me
P001.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
P001.Show 1
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text2.SetFocus
    Exit Sub
End If
Dim Tanya
Tanya = MsgBox("YAKIN AKAN MERUBAH DATA", vbOKCancel, "KONFIRMASI")
If Tanya = vbOK Then
    SCari = "Select * From ms_pasienwhere NO_PASIEN = '" + Trim(BR) + "'"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    RCari.Edit
        RCari("NAMA") = Text2
        RCari("ALAMAT") = Text3
        RCari("TELEPON") = Text4
        RCari("CM") = Text5
        RCari("KG") = Text6
        RCari("NAMA_K") = Text7
        RCari("ALAMAT_K") = Text8
        RCari("TELP_K") = Text9
        RCari("HUB_K") = Text10
        RCari("KELAMIN") = Combo1
        RCari("AGAMA") = Combo2
        RCari("PEKERJAAN") = Combo3
        RCari("PENDIDIKAN") = Combo4
        RCari("DARAH") = Combo5
        RCari("TGL_LAHIR") = DTPicker1
        RCari("STS") = 1
    RCari.Update
    RCari.Close
    MsgBox "DATABASE TELAH DI UPDATE", vbCritical, "KONFIRMASI"
End If
Unload Me
P001.Show 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
ClearTextBoxes Me
Call IsiCombo
Call Cari
End Sub

Private Sub Cari()
SCari = "Select * From ms_pasienwhere NO_PASIEN = '" + Trim(BR) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
    Text1 = RCari("NO_PASIEN")
    Text2 = RCari("NAMA")
    Text3 = RCari("ALAMAT")
    Text4 = RCari("TELEPON")
    Text5 = RCari("CM")
    Text6 = RCari("KG")
    Text7 = RCari("NAMA_K")
    Text8 = RCari("ALAMAT_K")
    Text9 = RCari("TELP_K")
    Text10 = RCari("HUB_K")
    Combo1 = RCari("KELAMIN")
    Combo2 = RCari("AGAMA")
    Combo3 = RCari("PEKERJAAN")
    Combo4 = RCari("PENDIDIKAN")
    Combo5 = RCari("DARAH")
    DTPicker1 = RCari("TGL_LAHIR")
End If
RCari.Close
Set RCari = Nothing
End Sub
Private Sub IsiCombo()
Combo1.AddItem "LAKI - LAKI", 0
Combo1.AddItem "PEREMPUAN", 1
Combo1.ListIndex = 0

Combo2.AddItem "ISLAM", 0
Combo2.AddItem "KRISTEN", 1
Combo2.AddItem "KATHOLIK", 2
Combo2.AddItem "HINDU", 3
Combo2.AddItem "BUDHA", 4
Combo2.AddItem "LAINNYA", 5
Combo2.ListIndex = 0

Combo3.AddItem "PEGAWAI NEGERI", 0
Combo3.AddItem "TNI / ABRI", 1
Combo3.AddItem "PENSIUNAN", 2
Combo3.AddItem "PEGAWAI SWASTA", 3
Combo3.AddItem "PEDAGANG", 4
Combo3.AddItem "NELAYAN", 5
Combo3.AddItem "PETANI", 6
Combo3.AddItem "PEKERJA LEPAS", 7
Combo3.AddItem "IBU RUMAH TANGGA", 8
Combo3.AddItem "PELAJAR", 9
Combo3.AddItem "DI BAWAH UMUR", 10
Combo3.AddItem "TIDAK BEKERJA", 11
Combo3.AddItem "TIDAK TAHU", 12
Combo3.AddItem "LAINNYA", 13
Combo3.ListIndex = 0

Combo4.AddItem "TIDAK SEKOLAH", 0
Combo4.AddItem "BELUM / TIDAK TAMAT SD", 1
Combo4.AddItem "SD", 2
Combo4.AddItem "SLTP", 3
Combo4.AddItem "SLTA", 4
Combo4.AddItem "PERGURUAN TINGGI", 5
Combo4.AddItem "LAINNYA", 6
Combo4.ListIndex = 0

Combo5.AddItem "A", 0
Combo5.AddItem "AB", 1
Combo5.AddItem "B", 2
Combo5.AddItem "O", 3
Combo5.ListIndex = 0
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
'Call CekData
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text8_LostFocus()
Text8 = Format(Text8, ">")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, ">")
End Sub
