VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P004A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT DATA KUNJUNGAN POLIKLINIK"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   195
      Left            =   4920
      OleObjectBlob   =   "P004A.frx":0000
      TabIndex        =   27
      Top             =   4433
      Width           =   2535
   End
   Begin VB.TextBox Text200 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   8565
      TabIndex        =   25
      Text            =   "200"
      Top             =   7635
      Width           =   1020
   End
   Begin VB.TextBox Text100 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   7320
      TabIndex        =   24
      Text            =   "100"
      Top             =   7635
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+ OBAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9105
      TabIndex        =   6
      Top             =   4365
      Width           =   960
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   8745
      TabIndex        =   5
      Text            =   "4"
      Top             =   3960
      Width           =   1320
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3960
      Width           =   3525
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6360
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      Height          =   1065
      Left            =   -112
      ScaleHeight     =   1005
      ScaleWidth      =   10710
      TabIndex        =   20
      Top             =   8040
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   180
      TabIndex        =   3
      Text            =   "8"
      Top             =   7125
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
      OleObjectBlob   =   "P004A.frx":006F
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
      TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Text            =   "1"
         Top             =   240
         Width           =   1800
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   195
         Left            =   90
         OleObjectBlob   =   "P004A.frx":02A3
         TabIndex        =   13
         Top             =   735
         Width           =   825
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Left            =   90
         OleObjectBlob   =   "P004A.frx":0302
         TabIndex        =   14
         Top             =   360
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2355
         Left            =   45
         TabIndex        =   21
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   195
         Left            =   3195
         OleObjectBlob   =   "P004A.frx":035D
         TabIndex        =   26
         Top             =   360
         Width           =   2625
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004A.frx":03C4
      TabIndex        =   15
      Top             =   3938
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   1710
      TabIndex        =   9
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
      Format          =   20185089
      CurrentDate     =   39286
      MinDate         =   2
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004A.frx":0435
      TabIndex        =   16
      Top             =   4500
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004A.frx":04B0
      TabIndex        =   17
      Top             =   5265
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004A.frx":0529
      TabIndex        =   18
      Top             =   6075
      Width           =   1905
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "P004A.frx":0590
      TabIndex        =   19
      Top             =   6840
      Width           =   1905
   End
   Begin MSFlexGridLib.MSFlexGrid GridObat 
      Height          =   2850
      Left            =   4815
      TabIndex        =   22
      ToolTipText     =   "Double Klik Untuk Menghapus"
      Top             =   4770
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   5027
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      MergeCells      =   4
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "DAFTAR OBAT"
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
      Height          =   4335
      Left            =   4815
      TabIndex        =   23
      Top             =   3645
      Width           =   5325
   End
End
Attribute VB_Name = "P004A"
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
If Text5 = "" Or Text6 = "" Or Text8 = "" Then
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
    Call CariObat
End If
End Sub

Private Sub CariObat()
SCari = "SELECT * From ms_obat Where Nama_Obat = '" + Trim(Combo2) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    SkinLabel4 = "Stock Obat = " + Trim(RCari("QTY"))
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Command1_Click()
If Text4 = "" Then Exit Sub
    
SSave = "Select * From temp_obat"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("nama_obat") = Format(Combo2, ">")
        
        SCari = "SELECT * From ms_obat Where Nama_Obat = '" + Trim(Combo2) + "'"
        Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
        If RCari.RowCount <> 0 Then
            RSave("harga") = CCur(RCari("harga"))
        End If
        RCari.Close
        Set RCari = Nothing

        RSave("qty") = CCur(Text4)
RSave.Update
RSave.Close
Set RSave = Nothing

Call IsiGridObat

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
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
'SkinLabel4.Caption = Date
DTPicker1 = Now
ClearTextBoxes Me

SDel = "Delete From temp_obat"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)

Call NoTrans
Call IsiData
Call SiapkanGrid
Call IsiGridObat
Call IsiGrid

DTPicker1 = Date

End Sub

Private Sub NoTrans()
SCari = "Select * From tabel_periksa order by No_Urut desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    SkinLabel3 = RCari("No_Urut") + 1
Else
    SkinLabel3 = 1
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub SiapkanGridObat()
GridObat.Refresh

With GridObat
    .Cols = 3
    .Row = 0
    .RowHeight(0) = 400
    .Col = 0: .ColWidth(0) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1250: .Text = "HARGA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "QTY": .CellAlignment = 4
End With
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

Private Sub IsiGridObat()
GridObat.Refresh
GridObat.Clear

Call SiapkanGridObat

SKTG = "Select * From temp_obat order by NO Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      GridObat.Rows = B + 1
      GridObat.Row = B
         With GridObat
            .Col = 0: .Text = RKTG("nama_obat")
            .Col = 1: .Text = CCur(RKTG("harga"))
            .Col = 2: .Text = CCur(RKTG("qty")): .CellAlignment = 4
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing

SCari = "SELECT Sum(temp_obat.harga) AS SumOfharga, Sum(temp_obat.qty) AS SumOfqty From temp_obat"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    On Error Resume Next
    Text100 = CCur(RCari("SumOfharga"))
    Text200 = CCur(RCari("SumOfqty"))
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

SCombo = "Select * from ms_diagnosa order by diagnosa asc"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdOpenKeyset)
RCombo.MoveFirst
Do While Not RCombo.EOF
    Combo1.AddItem Format(RCombo("diagnosa"), ">")
RCombo.MoveNext
Loop
RCombo.Close
Set RCombo = Nothing
Combo1.ListIndex = 0

SCombo = "Select * from ms_obat order by nama_obat asc"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdOpenKeyset)
RCombo.MoveFirst
Do While Not RCombo.EOF
    Combo2.AddItem Format(RCombo("nama_obat"), ">")
RCombo.MoveNext
Loop
RCombo.Close
Set RCombo = Nothing
Combo2.ListIndex = 0
End Sub

Private Sub Simpan()
SCari = "Select * From temp_obat"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.MoveFirst
Do While Not RCari.EOF
    NAMAOBAT = RCari("nama_obat")
    HARGAOBAT = RCari("harga")
    JUMLAHBARANG = RCari("qty")
    
    SSave = "Select * From Tabel_Periksa"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.AddNew
        RSave("No_Urut") = SkinLabel3
        RSave("No_Pasien") = Text1
        RSave("Tgl_Kunj") = DateValue(DTPicker1)
        RSave("anamnese") = Text5
        RSave("fisik") = Text6
        RSave("diagnosa") = Combo1
        RSave("terapi") = Text8
        RSave("obat") = NAMAOBAT
        RSave("harga") = HARGAOBAT
        RSave("qty") = JUMLAHBARANG
        RSave("Tgl") = Date
    RSave.Update
    RSave.Close
    Set RSave = Nothing
    
    SCari2 = "Select * From ms_obat where nama_obat = '" + Trim(NAMAOBAT) + "'"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.Edit
        JMLOBAT = RCari2("qty")
        RCari2("qty") = CCur(JMLOBAT) - CCur(JUMLAHBARANG)
        
            SSave = "Select * From ms_obat_his"
            Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
            RSave.AddNew
                RSave("tanggal") = Date
                RSave("nama_obat") = Trim(NAMAOBAT)
                RSave("harga") = CCur(HARGAOBAT)
                RSave("saldo_awal") = CCur(JMLOBAT)
                RSave("masuk") = 0
                RSave("keluar") = CCur(JUMLAHBARANG)
                RSave("saldo_akhir") = CCur(JMLOBAT) - CCur(JUMLAHBARANG)
                RSave("keterangan") = Trim(SkinLabel3)
            RSave.Update
            RSave.Close
            Set RSave = Nothing
        
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    


        
RCari.MoveNext
Loop
RCari.Close
Set RCari = Nothing

'SCari = "Select * From ms_diagnosa where diagnosa = '" + Trim(Combo1) + "'"
'Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
'RCari.Edit
'    RCari("sts") = RCari("sts") + 1
'RCari.Update
'RCari.Close
'Set RCari = Nothing

End Sub

Private Sub GridObat_DblClick()
SDel = "Delete From temp_obat Where Nama_Obat = '" + Trim(GridObat.TextMatrix(GridObat.Row, 0)) + "'"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)

Call IsiGridObat
Combo2.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, ">")
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

