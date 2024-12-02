VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LP01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "MUTASI OBAT"
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
      Left            =   3675
      TabIndex        =   10
      Top             =   2340
      Width           =   2190
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DAFTAR OBAT"
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
      Left            =   510
      TabIndex        =   9
      Top             =   2340
      Width           =   2190
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GRAFIK DIAGNOSA"
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
      Left            =   3675
      TabIndex        =   7
      Top             =   1260
      Width           =   2190
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
      Left            =   2122
      TabIndex        =   1
      Top             =   3375
      Width           =   2190
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LAPORAN TRANSAKSI"
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
      Left            =   510
      TabIndex        =   0
      Top             =   1260
      Width           =   2190
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   765
      OleObjectBlob   =   "LP01.frx":0000
      Top             =   7440
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   510
      TabIndex        =   2
      Top             =   240
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
      Format          =   20185089
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   420
      Left            =   3555
      TabIndex        =   3
      Top             =   240
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
      Format          =   20185089
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   3930
      OleObjectBlob   =   "LP01.frx":0234
      TabIndex        =   4
      Top             =   765
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   240
      Left            =   885
      OleObjectBlob   =   "LP01.frx":02AC
      TabIndex        =   5
      Top             =   765
      Width           =   1560
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   225
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      Height          =   1230
      Left            =   -30
      ScaleHeight     =   1170
      ScaleWidth      =   6480
      TabIndex        =   6
      Top             =   3195
      Width           =   6540
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   -180
      TabIndex        =   8
      Top             =   -180
      Width           =   7125
   End
End
Attribute VB_Name = "LP01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RR, RCari As rdoResultset
Private SR, SCari As String

Private T, M, D, T2, M2, D2

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call TGL
crpt.ReportFileName = App.Path & "\ReportMR\LapHar.rpt"
crpt.SelectionFormula = "{P004.TGL_KUNJ} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ")"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
crpt.Reset
End Sub

Private Sub Command2_Click()

SCari = "Select * From ms_diagnosa"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
    Do Until RCari.EOF
    RCari.Edit
        RCari("sts") = 0
    RCari.Update
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing

Call TGL

SCari = "Select * From tabel_periksa Where Tgl_Kunj >= #" + Trim(DTPicker1) + "# and Tgl_Kunj <= #" + Trim(DTPicker2) + "#"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
    Do Until RCari.EOF
        SCari2 = "Select * From ms_diagnosa where diagnosa = '" + Trim(RCari("DIAGNOSA")) + "'"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.Edit
            RCari2("sts") = RCari2("sts") + 1
            RCari2("Tgl_Awal") = DTPicker1
            RCari2("Tgl_Akhir") = DTPicker2
        RCari2.Update
        RCari2.Close
        Set RCari2 = Nothing
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing

B = 1
SCari = "Select * From ms_diagnosa Order by STS Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
    Do Until RCari.EOF
    RCari.Edit
        RCari("NO") = B
    RCari.Update
    RCari.MoveNext
    B = B + 1
    Loop
End If
RCari.Close
Set RCari = Nothing

crpt.ReportFileName = App.Path & "\ReportMR\Grafik.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
crpt.Reset

End Sub

Private Sub Command3_Click()
crpt.ReportFileName = App.Path & "\ReportMR\Obat.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
crpt.Reset
End Sub

Private Sub Command4_Click()
crpt.ReportFileName = App.Path & "\ReportMR\Mutasi_Obat.rpt"
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
DTPicker1 = Format(DateSerial(Year(Date), Month(Date), 1), "DD/MM/YYYY")
DTPicker2 = Date
End Sub

Private Sub TGL()
T = Year(DTPicker1)
M = Month(DTPicker1)
D = Day(DTPicker1)

T2 = Year(DTPicker2)
M2 = Month(DTPicker2)
D2 = Day(DTPicker2)
End Sub
