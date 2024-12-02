VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form C002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENCARIAN CANGGIH"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      TabIndex        =   11
      Top             =   4560
      Width           =   10335
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   8040
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "C002.frx":0000
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "C002.frx":0005
         Top             =   2040
         Width           =   2130
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "C002.frx":000A
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "C002.frx":000F
         Top             =   2040
         Width           =   2130
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1080
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "C002.frx":0014
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox Text17 
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
         TabIndex        =   17
         Text            =   "17"
         Top             =   2040
         Width           =   2160
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   16
         Text            =   "C002.frx":0019
         Top             =   1320
         Width           =   2280
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   15
         Text            =   "8"
         Top             =   840
         Width           =   2280
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   14
         Text            =   "7"
         Top             =   360
         Width           =   2280
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   13
         Text            =   "11"
         Top             =   2640
         Width           =   2280
      End
      Begin VB.TextBox Text10 
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
         TabIndex        =   12
         Text            =   "10"
         Top             =   2160
         Width           =   1800
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C002.frx":001D
         TabIndex        =   23
         Top             =   495
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C002.frx":0078
         TabIndex        =   24
         Top             =   975
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C002.frx":00D7
         TabIndex        =   25
         Top             =   1440
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C002.frx":013A
         TabIndex        =   26
         Top             =   2295
         Width           =   1125
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   225
         Left            =   120
         OleObjectBlob   =   "C002.frx":01A5
         TabIndex        =   27
         Top             =   2775
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   195
         Left            =   8040
         OleObjectBlob   =   "C002.frx":0204
         TabIndex        =   28
         Top             =   1800
         Width           =   1125
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   195
         Left            =   3480
         OleObjectBlob   =   "C002.frx":0275
         TabIndex        =   29
         Top             =   360
         Width           =   870
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   195
         Left            =   3480
         OleObjectBlob   =   "C002.frx":02DA
         TabIndex        =   30
         Top             =   1800
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   195
         Left            =   5760
         OleObjectBlob   =   "C002.frx":0353
         TabIndex        =   31
         Top             =   360
         Width           =   870
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   195
         Left            =   5760
         OleObjectBlob   =   "C002.frx":03BA
         TabIndex        =   32
         Top             =   1800
         Width           =   960
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   195
         Left            =   8040
         OleObjectBlob   =   "C002.frx":0421
         TabIndex        =   33
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PARAMETER PENCARIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   10335
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
         Height          =   435
         Left            =   8280
         TabIndex        =   34
         Top             =   1890
         Width           =   1890
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
         Height          =   435
         Left            =   5760
         TabIndex        =   10
         Top             =   1890
         Width           =   2250
      End
      Begin VB.TextBox Text6 
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
         Left            =   7920
         TabIndex        =   9
         Text            =   "6"
         Top             =   1320
         Width           =   2325
      End
      Begin VB.TextBox Text5 
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
         Left            =   7920
         TabIndex        =   8
         Text            =   "5"
         Top             =   840
         Width           =   2325
      End
      Begin VB.TextBox Text4 
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
         Left            =   7920
         TabIndex        =   7
         Text            =   "4"
         Top             =   360
         Width           =   2325
      End
      Begin VB.TextBox Text3 
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
         Left            =   2760
         TabIndex        =   6
         Text            =   "3"
         Top             =   1800
         Width           =   2325
      End
      Begin VB.TextBox Text1 
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
         Left            =   2760
         TabIndex        =   5
         Text            =   "1"
         Top             =   840
         Width           =   2325
      End
      Begin VB.TextBox Text2 
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
         Left            =   2760
         TabIndex        =   4
         Text            =   "2"
         Top             =   1320
         Width           =   2325
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   420
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   2310
         _ExtentX        =   4075
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
         Format          =   57737217
         CurrentDate     =   39286
         MinDate         =   39083
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Left            =   240
         OleObjectBlob   =   "C002.frx":0480
         TabIndex        =   35
         Top             =   473
         Width           =   2010
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   195
         Left            =   240
         OleObjectBlob   =   "C002.frx":04F9
         TabIndex        =   36
         Top             =   923
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   195
         Left            =   240
         OleObjectBlob   =   "C002.frx":0558
         TabIndex        =   37
         Top             =   1403
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   195
         Left            =   240
         OleObjectBlob   =   "C002.frx":05BD
         TabIndex        =   38
         Top             =   1883
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   195
         Left            =   5880
         OleObjectBlob   =   "C002.frx":0636
         TabIndex        =   39
         Top             =   443
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   195
         Left            =   5880
         OleObjectBlob   =   "C002.frx":069D
         TabIndex        =   40
         Top             =   923
         Width           =   1770
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   195
         Left            =   5880
         OleObjectBlob   =   "C002.frx":0704
         TabIndex        =   41
         Top             =   1403
         Width           =   1770
      End
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
      Left            =   -135
      TabIndex        =   0
      Top             =   7920
      Width           =   10650
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "C002.frx":0763
      Top             =   3240
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1845
      Left            =   0
      TabIndex        =   1
      Top             =   2610
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   3254
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "C002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim a, Isi, Pusing As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String




Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Grid.Clear
Call SiapkanGrid2
Call IsiGrid2
End Sub

Private Sub Command2_Click()
Call SiapkanGrid
Call IsiGrid
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=MR", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
DTPicker1 = Now

Call SiapkanGrid
'Call IsiGrid
Grid.Refresh
End Sub

Private Sub SiapkanGrid2()
Grid.Rows = 2
With Grid
    .Cols = 10
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO. PASIEN": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "TGL. KUNJUNGAN": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "BIAYA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "KELUHAN": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1700: .Text = "PEMERIKSAAN FISIK": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1500: .Text = "DIAGNOSA": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 1250: .Text = "TINDAKAN": .CellAlignment = 4
    .Col = 9: .ColWidth(9) = 1500: .Text = "OBAT": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid2()
SKTG = "Select * From P004 where NAMA = '" + Trim(Text1) + "'"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RKTG("NO_PASIEN"): .CellAlignment = 4
            .Col = 1: .Text = RKTG("NAMA"): .CellAlignment = 4
            .Col = 2: .Text = RKTG("ALAMAT"): .CellAlignment = 4
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

Private Sub SiapkanGrid()
With Grid
    .Cols = 10
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO. PASIEN": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "ALAMAT": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "TGL. KUNJUNGAN": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "BIAYA": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "KELUHAN": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1700: .Text = "PEMERIKSAAN FISIK": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1500: .Text = "DIAGNOSA": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 1250: .Text = "TINDAKAN": .CellAlignment = 4
    .Col = 9: .ColWidth(9) = 1500: .Text = "OBAT": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SKTG = "Select * From P004 order by NAMA Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      Grid.Rows = B + 1
      Grid.Row = B
         With Grid
            .Col = 0: .Text = RKTG("NO_PASIEN"): .CellAlignment = 4
            .Col = 1: .Text = RKTG("NAMA"): .CellAlignment = 4
            .Col = 2: .Text = RKTG("ALAMAT"): .CellAlignment = 4
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
Text7 = Format(Grid.TextMatrix(Grid.Row, 0), ">")
Text8 = Format(Grid.TextMatrix(Grid.Row, 1), ">")
Text9 = Format(Grid.TextMatrix(Grid.Row, 2), ">")
'Text10 = Format(grid.TextMatrix(grid.Row, 3), ">")
'Text11 = Format(grid.TextMatrix(grid.Row, 4), ">")
Text12 = Format(Grid.TextMatrix(Grid.Row, 5), ">")
Text13 = Format(Grid.TextMatrix(Grid.Row, 6), ">")
Text14 = Format(Grid.TextMatrix(Grid.Row, 7), ">")
Text15 = Format(Grid.TextMatrix(Grid.Row, 8), ">")
Text16 = Format(Grid.TextMatrix(Grid.Row, 9), ">")
Text17 = Format(Grid.TextMatrix(Grid.Row, 4), ">")
End Sub

