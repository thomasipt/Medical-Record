VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form OBAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT DATA OBAT"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
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
      Left            =   5805
      TabIndex        =   11
      Top             =   1575
      Width           =   900
   End
   Begin VB.TextBox Text4 
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
      Left            =   4905
      TabIndex        =   10
      Text            =   "4"
      Top             =   765
      Visible         =   0   'False
      Width           =   2490
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
      Height          =   360
      Left            =   1650
      TabIndex        =   2
      Text            =   "3"
      Top             =   945
      Width           =   2490
   End
   Begin VB.TextBox Text1 
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
      Left            =   1650
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "1"
      Top             =   90
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
      Left            =   1650
      TabIndex        =   1
      Text            =   "2"
      Top             =   510
      Width           =   2490
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
      Left            =   2250
      TabIndex        =   4
      Top             =   1575
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
      Left            =   300
      TabIndex        =   3
      Top             =   1575
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
      Left            =   6765
      TabIndex        =   5
      Top             =   1575
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4995
      OleObjectBlob   =   "OBAT.frx":0000
      Top             =   7515
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   3690
      Top             =   7605
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   225
      OleObjectBlob   =   "OBAT.frx":0234
      TabIndex        =   6
      Top             =   165
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   225
      OleObjectBlob   =   "OBAT.frx":0293
      TabIndex        =   7
      Top             =   540
      Width           =   1365
   End
   Begin VB.PictureBox Picture2 
      Height          =   1365
      Left            =   -45
      ScaleHeight     =   1305
      ScaleWidth      =   9045
      TabIndex        =   8
      Top             =   1440
      Width           =   9105
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   225
      OleObjectBlob   =   "OBAT.frx":02F4
      TabIndex        =   9
      Top             =   975
      Width           =   1365
   End
End
Attribute VB_Name = "OBAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lokasi As String
Private Filter As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RKTG2, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SKTG2, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Call Edit

Unload Me
OBAT.Show 1
End Sub

Private Sub Edit()
SSave = "Select * From ms_obat_his"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("tanggal") = Date
    RSave("nama_obat") = Trim(Text1)
    RSave("harga") = CCur(Text2)
    RSave("saldo_awal") = CCur(Text4)
    RSave("masuk") = CCur(Text3) - CCur(Text4)
    RSave("keluar") = 0
    RSave("saldo_akhir") = CCur(Text3)
    RSave("keterangan") = "EDIT DATA OBAT"
RSave.Update
RSave.Close
Set RSave = Nothing

SCari = "Select * From ms_obat where nama_obat = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Edit
    RCari("nama_obat") = Trim(Text1)
    RCari("harga") = CCur(Text2)
    RCari("qty") = CCur(Text3)
RCari.Update
RCari.Close
Set RCari = Nothing
MsgBox "DATABASE TELAH DI UPDATE", vbCritical, "KONFIRMASI"
End Sub

Private Sub Simpan()
SSave = "Select * From ms_obat"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("nama_obat") = Trim(Text1)
    RSave("harga") = CCur(Text2)
    RSave("qty") = CCur(Text3)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Call Simpan

Unload Me
OBAT.Show 1
End Sub

Private Sub Command1_Click()
crpt.ReportFileName = App.Path & "\ReportMR\Obat.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
crpt.Reset
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

ClearTextBoxes Me

cmdOK.Visible = True
cmdEDIT.Visible = False
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

SCari = "Select * From ms_obat where nama_obat = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        Tanya = MsgBox("EDIT DATA OBAT...?", vbOKCancel, "KONFIRMASI")
        If Tanya = vbOK Then
            Text1 = RCari("nama_obat")
            Text2 = RCari("harga")
            Text3 = RCari("qty")
            Text4 = RCari("qty")
          
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub



