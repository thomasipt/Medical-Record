VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PASS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PASSWORD"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
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
      Left            =   53
      TabIndex        =   3
      Top             =   1507
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
      Left            =   2738
      TabIndex        =   4
      Top             =   1492
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2273
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   89
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2273
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   517
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2273
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   952
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3555
      OleObjectBlob   =   "PASS.frx":0000
      Top             =   3345
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   353
      OleObjectBlob   =   "PASS.frx":0234
      TabIndex        =   5
      Top             =   142
      Width           =   1830
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   353
      OleObjectBlob   =   "PASS.frx":02A6
      TabIndex        =   6
      Top             =   577
      Width           =   1830
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   353
      OleObjectBlob   =   "PASS.frx":0318
      TabIndex        =   7
      Top             =   1012
      Width           =   1830
   End
End
Attribute VB_Name = "PASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSkin, RTOKO, RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private SSkin, STOKO, STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "DATA MASIH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    Exit Sub
Else
    Call CEK
End If
Unload Me
LOGIN.Show 1
End Sub

Private Sub CEK()
SCari = "Select * From C013 where Password = '" + Text1 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Call CEK2
Else
    MsgBox "PASSWORD LAMA TIDAK TERDAFTAR", vbCritical, "KONFIRMASI"
    ClearTextBoxes Me
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End
End Sub

Private Sub CEK2()
If Text2.Text = Text3.Text Then
    Call Simpan
Else
    MsgBox "PASSWORD BARU TIDAK SESUAI DENGAN KONFIRMASI", vbCritical, "KONFIRMASI"
    Text2 = ""
    Text3 = ""
    Text2.SetFocus
End If
End Sub

Private Sub Simpan()
SSave = "Select * From C013"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.Edit
    RSave("Password") = Trim(Text3)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
ClearTextBoxes Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = RGB(255, 255, 255)
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = RGB(255, 255, 255)
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text3_GotFocus()
Text3.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = RGB(255, 255, 255)
Text3 = Format(Text3, ">")
End Sub


