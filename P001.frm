VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form P001 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA PASIEN"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "PERIKSA NON &KLINIK"
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
      Left            =   3825
      TabIndex        =   11
      Top             =   5220
      Width           =   3465
   End
   Begin VB.CommandButton Command3 
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
      Left            =   12210
      TabIndex        =   9
      Top             =   5220
      Width           =   1980
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   7020
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&HISTORY"
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
      Left            =   9045
      TabIndex        =   3
      Top             =   5220
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&PERIKSA KLINIK"
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
      Left            =   270
      TabIndex        =   2
      Top             =   5220
      Width           =   3465
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1530
      TabIndex        =   1
      Text            =   "2"
      Top             =   540
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   1530
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "1"
      Top             =   135
      Width           =   1725
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "P001.frx":0000
      TabIndex        =   4
      Top             =   225
      Width           =   1185
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1440
      OleObjectBlob   =   "P001.frx":005B
      Top             =   8235
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "P001.frx":028F
      TabIndex        =   5
      Top             =   630
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3885
      Left            =   45
      TabIndex        =   6
      ToolTipText     =   "Double Klik Untuk Melihat History Per Kode"
      Top             =   1080
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   6853
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
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   45
      ScaleHeight     =   855
      ScaleWidth      =   10665
      TabIndex        =   7
      Top             =   5040
      Width           =   10725
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Tampilkan Semua"
      Height          =   360
      Left            =   12555
      TabIndex        =   8
      Top             =   540
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   45
      TabIndex        =   10
      Top             =   -45
      Width           =   5685
   End
End
Attribute VB_Name = "P001"
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

Private Sub Command1_Click()
BR = ""
BR = Trim(Text1)

If Text1 = "" Then
    MsgBox "PILIH DATA PASIEN", vbCritical, "KONFIRMASI"
    Exit Sub
End If
    
'Unload Me
P004B.Show 1

End Sub

Private Sub Command10_Click()
SKTG = "Select * From ms_pasien order by kode asc"
Call IsiGrid2
End Sub

Private Sub Command2_Click()
BR = ""
BR = Trim(Text1)

If Text1 = "" Then
    MsgBox "PILIH DATA PASIEN", vbCritical, "KONFIRMASI"
    Exit Sub
End If
    
'Unload Me
P004A.Show 1

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command6_Click()
crpt.ReportFileName = App.Path & "\ReportMR\RekamMedis.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

ClearTextBoxes Me

Call SiapkanGrid
'Call IsiGrid

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 10
    .Row = 0
    .RowHeight(0) = 400
    .Col = 0: .ColWidth(0) = 1000:: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 3000: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 5000: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1250: .Text = "BAGIAN": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "JABATAN": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1000: .Text = "GOL.DARAH": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1500: .Text = "TELEPON": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1250: .Text = "KELAMIN": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 1250: .Text = "TGL LAHIR": .CellAlignment = 4
    .Col = 9: .ColWidth(9) = 2000: .Text = "JAMSOSTEK": .CellAlignment = 4
    
End With
End Sub

Private Sub IsiGrid()
SKTG = "Select * From ms_pasien order by kode Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
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

Private Sub CariNo()
SKTG = "Select * From ms_pasien where kode like '" + Trim(Text1) + "'"
Call IsiGrid2
End Sub

Private Sub CariNama()
SKTG = "Select * From ms_pasien where NAMA like '" + Trim(Text2) + "%'"
Call IsiGrid2
End Sub

Private Sub IsiGrid2()
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
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

Private Sub grid_dblClick()
'If Text1 = "" Then Exit Sub

Text1 = Format(grid.TextMatrix(grid.Row, 0), ">")
Text2 = Format(grid.TextMatrix(grid.Row, 1), ">")

'Text1.Enabled = False
'Text2.Enabled = False

crpt.ReportFileName = App.Path & "\ReportMR\RekamMedis.rpt"
crpt.SelectionFormula = "{ms_pasien.kode} = " + Text1
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
crpt.Reset
End Sub

Private Sub Text1_GotFocus()
'Call IsiGrid
Text2 = ""
Text3 = ""
Text1.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text2_GotFocus()
'Call IsiGrid
Text1 = ""
Text3 = ""
Text2.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text3_GotFocus()
'Call IsiGrid
Text1 = ""
Text2 = ""
Text3.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1 = "" Then Exit Sub
    If KeyAscii = 13 Then
        Call CariNo
    End If
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = RGB(255, 255, 255)
    Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2 = "" Then Exit Sub
    If KeyAscii = 13 Then
        Call CariNama
    End If
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = RGB(255, 255, 255)
    Text2 = Format(Text2, ">")
End Sub

