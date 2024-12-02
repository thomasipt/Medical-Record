VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   990
      OleObjectBlob   =   "Grafik.frx":0000
      Top             =   7605
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6180
      Left            =   45
      OleObjectBlob   =   "Grafik.frx":0234
      TabIndex        =   0
      Top             =   45
      Width           =   8295
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Grafik"
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

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

ClearTextBoxes Me

Call SiapkanGrafik
End Sub

Private Sub IsiGrid()
SKTG = "Select * From ms_pasien order by kode Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RKTG("kode"): .CellAlignment = 4
            .Col = 1: .Text = RKTG("nama")
            .Col = 2: .Text = RKTG("alamat")
            .Col = 3: .Text = RKTG("BAGIAN"): .CellAlignment = 4
            .Col = 4: .Text = RKTG("JABATAN"): .CellAlignment = 4
            .Col = 5: .Text = RKTG("GOL_DARAH"): .CellAlignment = 4
            .Col = 6: .Text = RKTG("TELEPON"): .CellAlignment = 3
            .Col = 7: .Text = RKTG("JENIS_KELAMIN"): .CellAlignment = 3
            .Col = 8: .Text = RKTG("TANGGAL_LAHIR"): .CellAlignment = 4
            .Col = 9: .Text = RKTG("NO_JAMSOSTEK"): .CellAlignment = 4
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub SiapkanGrafik()
B = 1

SCari = "SELECT * From V_DIAGNOSA order by STS Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
MSChart1.ColumnCount = 10
MSChart1.RowCount = 1
    Do Until RCari.EOF
        MSChart1.RowLabel = RCari("DIAGNOSA")
        MSChart1.Data = RCari("STS")
    B = B + 1
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub
