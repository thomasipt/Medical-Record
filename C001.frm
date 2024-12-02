VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form C001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENCARIAN"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text14 
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
      Left            =   1800
      TabIndex        =   30
      Text            =   "14"
      Top             =   1200
      Width           =   3990
   End
   Begin VB.TextBox Text13 
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
      Left            =   1800
      TabIndex        =   29
      Text            =   "13"
      Top             =   720
      Width           =   3990
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   3990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEMUA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   7920
      TabIndex        =   2
      Top             =   240
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARI KUNJUNGAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   7920
      TabIndex        =   1
      Top             =   240
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      Caption         =   "DETAIL DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   10335
      Begin VB.TextBox Text5 
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
         Left            =   1440
         TabIndex        =   27
         Text            =   "5"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.TextBox Text6 
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
         Left            =   960
         TabIndex        =   26
         Text            =   "6"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.TextBox Text2 
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
         Left            =   960
         TabIndex        =   25
         Text            =   "2"
         Top             =   360
         Width           =   2280
      End
      Begin VB.TextBox Text3 
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
         Left            =   960
         TabIndex        =   24
         Text            =   "3"
         Top             =   840
         Width           =   2280
      End
      Begin VB.TextBox Text4 
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
         Height          =   720
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "C001.frx":0000
         Top             =   1320
         Width           =   2280
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
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
         Left            =   8040
         TabIndex        =   16
         Text            =   "12"
         Top             =   2040
         Width           =   2160
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "C001.frx":0002
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "C001.frx":0006
         Top             =   2040
         Width           =   2130
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "C001.frx":000A
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "C001.frx":000E
         Top             =   2040
         Width           =   2130
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   8040
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "C001.frx":0013
         Top             =   600
         Width           =   2130
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C001.frx":0018
         TabIndex        =   6
         Top             =   495
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C001.frx":0073
         TabIndex        =   7
         Top             =   975
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C001.frx":00D2
         TabIndex        =   8
         Top             =   1440
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C001.frx":0135
         TabIndex        =   9
         Top             =   2295
         Visible         =   0   'False
         Width           =   1125
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C001.frx":01A0
         TabIndex        =   10
         Top             =   2775
         Visible         =   0   'False
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   195
         Left            =   8040
         OleObjectBlob   =   "C001.frx":01FF
         TabIndex        =   17
         Top             =   1800
         Width           =   1125
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   195
         Left            =   3480
         OleObjectBlob   =   "C001.frx":0270
         TabIndex        =   18
         Top             =   360
         Width           =   870
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   195
         Left            =   3480
         OleObjectBlob   =   "C001.frx":02D5
         TabIndex        =   19
         Top             =   1800
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   195
         Left            =   5760
         OleObjectBlob   =   "C001.frx":034E
         TabIndex        =   20
         Top             =   360
         Width           =   870
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   195
         Left            =   5760
         OleObjectBlob   =   "C001.frx":03B5
         TabIndex        =   21
         Top             =   1800
         Width           =   960
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   195
         Left            =   8040
         OleObjectBlob   =   "C001.frx":041C
         TabIndex        =   22
         Top             =   360
         Width           =   555
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9240
      OleObjectBlob   =   "C001.frx":047B
      Top             =   2040
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
      Height          =   780
      Left            =   900
      TabIndex        =   3
      Top             =   7575
      Width           =   8595
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "C001.frx":06AF
      TabIndex        =   4
      Top             =   803
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Height          =   1500
      Left            =   -728
      ScaleHeight     =   1440
      ScaleWidth      =   11790
      TabIndex        =   28
      Top             =   7440
      Width           =   11850
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "C001.frx":070E
      TabIndex        =   31
      Top             =   1283
      Width           =   1185
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "C001.frx":0771
      TabIndex        =   32
      Top             =   323
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2355
      Left            =   15
      TabIndex        =   33
      Top             =   1665
      Width           =   10365
      _ExtentX        =   18283
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
End
Attribute VB_Name = "C001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim Filter As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RKTG2, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SKTG2, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Text1.Enabled = True
Text13.Enabled = True
Text14.Enabled = True

Text1.SetFocus
Command2.ZOrder

End Sub

Private Sub Command2_Click()
Unload Me
C001.Show 1
End Sub


Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

ClearTextBoxes Me

Call SiapkanGrid2
Call IsiGrid

Command1.ZOrder

Text1.Enabled = False
Text13.Enabled = False
Text14.Enabled = False

End Sub

Private Sub CariNo()
SKTG = "Select * From Tabel_Periksa where NO_PASIEN like '%" + Trim(Text1) + "'"
Call IsiGrid2
End Sub

Private Sub CariNama()
SKTG = "Select * From Tabel_Periksa where NAMA like '%" + Trim(Text13) + "%'"
Call IsiGrid2
End Sub

Private Sub CariAlamat()
SKTG = "Select * From Tabel_Periksa where ALAMAT like '%" + Trim(Text14) + "%'"
Call IsiGrid2
End Sub


Private Sub Text1_GotFocus()
Call IsiGrid
Text13 = ""
Text14 = ""
Text1.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text13_GotFocus()
Call IsiGrid
Text1 = ""
Text14 = ""
Text13.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text14_GotFocus()
Call IsiGrid
Text1 = ""
Text13 = ""
Text14.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CariNo
    End If
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = RGB(255, 255, 255)
    Text1 = Format(Text1, ">")
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CariNama
    End If
End Sub

Private Sub Text13_LostFocus()
Text13.BackColor = RGB(255, 255, 255)
    Text13 = Format(Text13, ">")
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CariAlamat
    End If
End Sub

Private Sub Text14_LostFocus()
Text14.BackColor = RGB(255, 255, 255)
    Text14 = Format(Text14, ">")
End Sub

Private Sub SiapkanGrid2()
With Grid
    .Rows = 3
    .Cols = 10
    .Row = 0
    .Col = 0: .ColWidth(0) = 1000: .Text = "NO. PASIEN": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2750: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "TGL. KUNJUNGAN": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "BIAYA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2500: .Text = "KELUHAN": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 2700: .Text = "PEMERIKSAAN FISIK": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 2500: .Text = "DIAGNOSA": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 2250: .Text = "TINDAKAN": .CellAlignment = 4
    .Col = 9: .ColWidth(9) = 2500: .Text = "OBAT": .CellAlignment = 4
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeRow(0) = True
    .MergeRow(1) = True
    
    .Row = 1
    .Col = 0: .ColWidth(0) = 1000: .Text = "NO. PASIEN": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2750: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "TGL. KUNJUNGAN": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "BIAYA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2500: .Text = "KELUHAN": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 2700: .Text = "PEMERIKSAAN FISIK": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 2500: .Text = "DIAGNOSA": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 2250: .Text = "TINDAKAN": .CellAlignment = 4
    .Col = 9: .ColWidth(9) = 2500: .Text = "OBAT": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid2()
'SKTG = "Select * From Tabel_Periksa where NAMA like '%" + Trim(Text1) + "%' or ALAMAT like '%" + Trim(Text1) + "%' or KELUH like '%" + Trim(Text1) + "%' or PERIKSA like '%" + Trim(Text1) + "%'"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid2
   RKTG.MoveFirst
   B = 2
   Do Until RKTG.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RKTG("NO_PASIEN"): .CellAlignment = 4
            .Col = 1: .Text = RKTG("NAMA")
            .Col = 2: .Text = RKTG("ALAMAT")
            .Col = 3: .Text = RKTG("TGL_KUNJ"): .CellAlignment = 4
            .Col = 4: .Text = Format(RKTG("BIAYA"), "##,###.00")
            .Col = 5: .Text = RKTG("KELUH"): .CellAlignment = 4
            .Col = 6: .Text = RKTG("PERIKSA"): .CellAlignment = 4
            .Col = 7: .Text = RKTG("DIAG"): .CellAlignment = 4
            .Col = 8: .Text = RKTG("TINDAKAN"): .CellAlignment = 4
            .Col = 9: .Text = RKTG("OBAT"): .CellAlignment = 4
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub IsiGrid()
SKTG = "Select * From Tabel_Periksa order by NO_PASIEN Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid2
   RKTG.MoveFirst
   B = 2
   Do Until RKTG.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RKTG("NO_PASIEN"): .CellAlignment = 4
            .Col = 1: .Text = RKTG("NAMA")
            .Col = 2: .Text = RKTG("ALAMAT")
            .Col = 3: .Text = RKTG("TGL_KUNJ"): .CellAlignment = 4
            .Col = 4: .Text = Format(RKTG("BIAYA"), "##,###.00")
            .Col = 5: .Text = RKTG("KELUH"): .CellAlignment = 4
            .Col = 6: .Text = RKTG("PERIKSA"): .CellAlignment = 4
            .Col = 7: .Text = RKTG("DIAG"): .CellAlignment = 4
            .Col = 8: .Text = RKTG("TINDAKAN"): .CellAlignment = 4
            .Col = 9: .Text = RKTG("OBAT"): .CellAlignment = 4
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub grid_dblClick()
Text2 = Format(Grid.TextMatrix(Grid.Row, 0), ">")
Text3 = Format(Grid.TextMatrix(Grid.Row, 1), ">")
Text4 = Format(Grid.TextMatrix(Grid.Row, 2), ">")
'Text5 = Format(grid.TextMatrix(grid.Row, 3), ">")
'Text6 = Format(grid.TextMatrix(grid.Row, 4), ">")
Text7 = Format(Grid.TextMatrix(Grid.Row, 5), ">")
Text8 = Format(Grid.TextMatrix(Grid.Row, 6), ">")
Text9 = Format(Grid.TextMatrix(Grid.Row, 7), ">")
Text10 = Format(Grid.TextMatrix(Grid.Row, 8), ">")
Text11 = Format(Grid.TextMatrix(Grid.Row, 9), ">")
Text12 = Format(Grid.TextMatrix(Grid.Row, 4), ">")
End Sub
