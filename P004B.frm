VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P004B 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT DATA PERIKSA LUAR"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   5175
      TabIndex        =   4
      Text            =   "10"
      Top             =   6300
      Width           =   4380
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   5175
      TabIndex        =   3
      Text            =   "9"
      Top             =   5535
      Width           =   4380
   End
   Begin VB.PictureBox Picture2 
      Height          =   1065
      Left            =   -112
      ScaleHeight     =   1005
      ScaleWidth      =   10710
      TabIndex        =   17
      Top             =   7095
      Width           =   10770
      Begin VB.CommandButton Command2 
         Caption         =   "&CLOSE"
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
         Left            =   8096
         TabIndex        =   6
         Top             =   135
         Width           =   1890
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&SIMPAN"
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
         Left            =   274
         TabIndex        =   5
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   5175
      TabIndex        =   2
      Text            =   "8"
      Top             =   4770
      Width           =   4380
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   173
      TabIndex        =   1
      Text            =   "6"
      Top             =   5535
      Width           =   4380
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   173
      TabIndex        =   0
      Text            =   "5"
      Top             =   4770
      Width           =   4380
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10530
      OleObjectBlob   =   "P004B.frx":0000
      Top             =   4005
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATA PASIEN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3510
      Left            =   105
      TabIndex        =   10
      Top             =   45
      Width           =   10050
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   810
         TabIndex        =   9
         Text            =   "2"
         Top             =   600
         Width           =   5040
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   810
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "1"
         Top             =   240
         Width           =   1800
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   195
         Left            =   90
         OleObjectBlob   =   "P004B.frx":0234
         TabIndex        =   11
         Top             =   735
         Width           =   825
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Left            =   90
         OleObjectBlob   =   "P004B.frx":0293
         TabIndex        =   12
         Top             =   360
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2355
         Left            =   45
         TabIndex        =   18
         Top             =   1080
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   4154
         _Version        =   393216
         Rows            =   4
         FixedRows       =   2
         FixedCols       =   0
         BackColorFixed  =   14737632
         BackColorBkg    =   16777152
         GridColor       =   0
         MergeCells      =   2
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   195
         Left            =   3225
         OleObjectBlob   =   "P004B.frx":02EE
         TabIndex        =   20
         Top             =   360
         Width           =   2625
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004B.frx":0355
      TabIndex        =   13
      Top             =   3938
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   1710
      TabIndex        =   7
      Top             =   3825
      Width           =   2850
      _ExtentX        =   5027
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004B.frx":03C6
      TabIndex        =   14
      Top             =   4500
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004B.frx":0441
      TabIndex        =   15
      Top             =   5265
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   195
      Left            =   5175
      OleObjectBlob   =   "P004B.frx":04BA
      TabIndex        =   16
      Top             =   4500
      Width           =   1905
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   195
      Left            =   5175
      OleObjectBlob   =   "P004B.frx":0533
      TabIndex        =   19
      Top             =   5265
      Width           =   1905
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   195
      Left            =   5175
      OleObjectBlob   =   "P004B.frx":0592
      TabIndex        =   21
      Top             =   6030
      Width           =   1905
   End
End
Attribute VB_Name = "P004B"
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
End Sub

Private Sub cmdOK_Click()
If Text5 = "" Or Text6 = "" Or Text8 = "" Or Text9 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text5.SetFocus
    Exit Sub
End If

Call Simpan

Unload Me
'P001.Show 1
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Command1_Click()
If Text4 = "" Then Exit Sub
    
SSave = "Select * From temp_obat"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("nama_obat") = Format(Combo2, ">")
        RSave("qty") = CCur(Text4)
RSave.Update
RSave.Close
Set RSave = Nothing

'Call IsiGridObat

Combo2.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
'P001.Show 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin2 + ".skn"
Skin1.ApplySkin hWnd
'SkinLabel4.Caption = Date
DTPicker1 = Now
ClearTextBoxes Me

SDel = "Delete From temp_obat"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)

Call NoTrans
Call IsiData
Call SiapkanGrid
Call IsiGrid

DTPicker1 = Date
End Sub

Private Sub IsiGrid()
SKTG = "Select * From tabel_periksa where NO_PASIEN like '" + Trim(Text1) + "'"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurRowVer)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 2
   Do Until RKTG.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RKTG("No_Urut"): .CellAlignment = 4
            .Col = 1: .Text = RKTG("Tgl_Kunj"): .CellAlignment = 4
            .Col = 2: .Text = RKTG("anamnese")
            .Col = 3: .Text = RKTG("fisik")
            .Col = 4: .Text = RKTG("diagnosa"): .CellAlignment = 4
            .Col = 5: .Text = RKTG("terapi"): .CellAlignment = 4
            .Col = 6: .Text = RKTG("obat"): .CellAlignment = 4
            .Col = 7: .Text = RKTG("qty"): .CellAlignment = 4
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub SiapkanGrid()
With Grid
    .Rows = 3
    .Cols = 8
    .Row = 0
    .Col = 0: .ColWidth(0) = 0: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "TGL": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2000: .Text = "ANAMNESE": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 2000: .Text = "FISIK": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 2000: .Text = "DIAGNOSA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2000: .Text = "TERAPI": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 2000: .Text = "OBAT": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 2000: .Text = "QTY": .CellAlignment = 4
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeRow(0) = True
    .MergeRow(1) = True
    
    .Row = 1
    .Col = 0: .ColWidth(0) = 0: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "TGL": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2000: .Text = "ANAMNESE": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 2000: .Text = "FISIK": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 2000: .Text = "DIAGNOSA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2000: .Text = "TERAPI": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 2000: .Text = "OBAT": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 2000: .Text = "QTY": .CellAlignment = 4
End With
End Sub

Private Sub NoTrans()
SCari = "Select * From tabel_periksa order by No_Urut desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    SkinLabel4 = RCari("No_Urut") + 1
Else
    SkinLabel4 = 1
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiData()
SCari = "Select * From ms_pasien where kode like '" + Trim(BR) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text1 = RCari("kode")
    Text2 = RCari("Nama")
    Me.Caption = "INPUT DATA KUNJUNGAN PASIEN NO " + Trim(Text1) + " NAMA " + Trim(Text2)
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Simpan()
SSave = "Select * From Tabel_Periksa"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.AddNew
    RSave("No_Urut") = SkinLabel4
    RSave("No_Pasien") = Text1
    RSave("Tgl_Kunj") = DateValue(DTPicker1)
    RSave("anamnese") = Text5
    RSave("fisik") = Text6
    RSave("diagnosa") = "-"
    RSave("terapi") = Text8
    RSave("obat") = Text9
    RSave("biaya") = CCur(Text10)
    RSave("Tgl") = Date
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
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
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_LostFocus()
Text9 = Format(Text9, ">")
End Sub

