VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT DATA PASIEN"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
      MaxLength       =   15
      TabIndex        =   6
      Text            =   "5"
      Top             =   3060
      Width           =   2040
   End
   Begin VB.CommandButton cmdEDIT 
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
      Left            =   3517
      TabIndex        =   22
      Top             =   3765
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "SIMPAN"
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
      TabIndex        =   10
      Top             =   3765
      Width           =   1890
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
      TabIndex        =   11
      Top             =   3780
      Width           =   1890
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      TabIndex        =   9
      Text            =   "Combo5"
      Top             =   3015
      Width           =   825
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      TabIndex        =   8
      Text            =   "Combo4"
      Top             =   2595
      Width           =   2670
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6165
      TabIndex        =   7
      Text            =   "Combo3"
      Top             =   2160
      Width           =   2670
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
      Top             =   1620
      Width           =   2040
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   2040
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
      Top             =   630
      Width           =   4920
   End
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
      Text            =   "P002.frx":0000
      Top             =   1125
      Width           =   4920
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7830
      OleObjectBlob   =   "P002.frx":0005
      Top             =   270
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   1710
      TabIndex        =   5
      Top             =   2535
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
      Format          =   21561345
      CurrentDate     =   39286
      MinDate         =   2
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P002.frx":0239
      TabIndex        =   12
      Top             =   195
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P002.frx":0294
      TabIndex        =   13
      Top             =   660
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P002.frx":02F3
      TabIndex        =   14
      Top             =   1155
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   195
      Left            =   150
      OleObjectBlob   =   "P002.frx":0356
      TabIndex        =   15
      Top             =   2655
      Width           =   1185
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P002.frx":03C1
      TabIndex        =   16
      Top             =   2205
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   225
      Left            =   6705
      OleObjectBlob   =   "P002.frx":0432
      TabIndex        =   17
      Top             =   1350
      Width           =   2040
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P002.frx":0497
      TabIndex        =   18
      Top             =   2205
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P002.frx":04FA
      TabIndex        =   19
      Top             =   2640
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   225
      Left            =   4605
      OleObjectBlob   =   "P002.frx":055F
      TabIndex        =   20
      Top             =   3060
      Width           =   1365
   End
   Begin VB.PictureBox Picture2 
      Height          =   1365
      Left            =   -90
      ScaleHeight     =   1305
      ScaleWidth      =   9045
      TabIndex        =   21
      Top             =   3645
      Width           =   9105
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   225
      Left            =   150
      OleObjectBlob   =   "P002.frx":05C8
      TabIndex        =   23
      Top             =   3135
      Width           =   1365
   End
End
Attribute VB_Name = "P002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RCombo, RSLUser, RSave, RSave2, REdit, RKTG, RKTG2, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SCombo, SQLUser, SSave, SSave2, SEdit, SKTG, SKTG2, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

    Call Simpan

Text1.SetFocus

Unload Me
P002.Show 1
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Call Simpan

Unload Me
P002.Show 1

End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Call Edit

Unload Me
P002.Show 1

End Sub

Private Sub Edit()
Dim kode_pasien

kode_pasien = Trim(Text1)

SCari = "Select * From ms_pasien where kode = " + kode_pasien
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Edit
    RCari("nama") = Trim(Text2)
    RCari("alamat") = Trim(Text3)
    RCari("TELEPON") = Trim(Text4)
    RCari("JENIS_KELAMIN") = Combo1
    RCari("TANGGAL_LAHIR") = DTPicker1
    RCari("BAGIAN") = Combo3
    RCari("JABATAN") = Combo4
    RCari("GOL_DARAH") = Combo5
RCari.Update
RCari.Close
Set RCari = Nothing
MsgBox "DATABASE TELAH DI UPDATE", vbCritical, "KONFIRMASI"
End Sub

Private Sub Simpan()
SSave = "Select * From ms_pasien"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("kode") = Trim(Text1)
    RSave("nama") = Trim(Text2)
    RSave("alamat") = Trim(Text3)
    RSave("TELEPON") = Trim(Text4)
    RSave("JENIS_KELAMIN") = Combo1
    RSave("TANGGAL_LAHIR") = DTPicker1
    RSave("BAGIAN") = Combo3
    RSave("JABATAN") = Combo4
    RSave("GOL_DARAH") = Combo5
    RSave("NO_JAMSOSTEK") = Trim(Text5)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
DTPicker1 = Date
ClearTextBoxes Me

Call IsiCombo
'Call NoPasien

cmdOK.Visible = True
cmdEDIT.Visible = False

End Sub

Private Sub NoPasien()
SCari = "SELECT Count(ms_pasien.kode) AS CountOfkode FROM ms_pasien"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text1 = RCari("CountOfkode") + 1
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiCombo()
Combo1.AddItem "LAKI - LAKI", 0
Combo1.AddItem "PEREMPUAN", 1
Combo1.ListIndex = 0

SCombo = "Select * from Bagian order by Bagian asc"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdOpenKeyset)
RCombo.MoveFirst
Do While Not RCombo.EOF
    Combo3.AddItem Format(RCombo("Bagian"), ">")
RCombo.MoveNext
Loop
RCombo.Close
Set RCombo = Nothing
Combo3.ListIndex = 0

SCombo = "Select * from Jabatan order by Jabatan asc"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdOpenKeyset)
RCombo.MoveFirst
Do While Not RCombo.EOF
    Combo4.AddItem Format(RCombo("Jabatan"), ">")
RCombo.MoveNext
Loop
RCombo.Close
Set RCombo = Nothing
Combo4.ListIndex = 0

SCombo = "Select * from Gol_Darah order by Gol_Darah asc"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdOpenKeyset)
RCombo.MoveFirst
Do While Not RCombo.EOF
    Combo5.AddItem Format(RCombo("Gol_Darah"), ">")
RCombo.MoveNext
Loop
RCombo.Close
Set RCombo = Nothing
Combo5.ListIndex = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
Call CekData
End Sub

Private Sub CekData()
Dim Tanya

If Text1.Text = "" Then Exit Sub

SCari = "Select * From ms_pasien where kode like '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        Tanya = MsgBox("KARYAWAN SUDAH TERDAFTAR, EDIT DATA...?", vbOKCancel, "KONFIRMASI")
        If Tanya = vbOK Then
            Text1 = RCari("kode")
            Text2 = RCari("nama")
            Text3 = RCari("alamat")
            Text4 = RCari("telepon")
            Text5 = RCari("NO_JAMSOSTEK")
            Combo1.Text = Format(RCari("JENIS_KELAMIN"), ">")
            DTPicker1.Value = RCari("tanggal_lahir")
            Combo3.Text = Format(RCari("bagian"), ">")
            Combo4.Text = Format(RCari("jabatan"), ">")
            Combo5.Text = Format(RCari("gol_darah"), ">")
            
            cmdOK.Visible = False
            cmdEDIT.Visible = True
            
            Text1.Enabled = False
        Else
            Text1 = ""
            Text1.SetFocus
        End If
    Else
       Text2.SetFocus
    Exit Sub
    End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, ">")
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

