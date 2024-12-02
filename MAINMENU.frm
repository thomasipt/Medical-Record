VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MAINMENU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MEDICAL RECORD     |     PT GUNUNG SUBUR KARANGANYAR"
   ClientHeight    =   4125
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9375
   Begin VB.CommandButton Command1 
      Height          =   2670
      Left            =   37
      TabIndex        =   4
      Top             =   75
      Width           =   9300
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
      Left            =   -75
      TabIndex        =   3
      Top             =   3450
      Width           =   9525
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2835
      Width           =   3060
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3150
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2820
      Width           =   3060
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6255
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2835
      Width           =   3060
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6240
      OleObjectBlob   =   "MAINMENU.frx":0000
      Top             =   7635
   End
   Begin VB.Menu A 
      Caption         =   "ADMINISTRATOR"
      Index           =   0
      Begin VB.Menu ADM 
         Caption         =   "PASIEN"
         Index           =   1
      End
      Begin VB.Menu ADM 
         Caption         =   "DIAGNOSA"
         Index           =   2
      End
      Begin VB.Menu ADM 
         Caption         =   "OBAT"
         Index           =   3
      End
   End
   Begin VB.Menu P 
      Caption         =   "PASIEN"
      Index           =   1
      Begin VB.Menu PD 
         Caption         =   "DATA PASIEN"
         Index           =   11
         Shortcut        =   {F1}
      End
      Begin VB.Menu PD 
         Caption         =   "INPUT DATA PASIEN"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu PD 
         Caption         =   "EDIT DATA PASIEN"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu PD 
         Caption         =   "INPUT DATA KUNJUNGAN"
         Index           =   14
         Visible         =   0   'False
      End
   End
   Begin VB.Menu L 
      Caption         =   "LAPORAN"
      Index           =   4
      Begin VB.Menu LP 
         Caption         =   "CETAK LAPORAN"
         Index           =   41
         Shortcut        =   {F3}
      End
      Begin VB.Menu LP 
         Caption         =   "LAPORAN TAHUNAN"
         Index           =   42
         Visible         =   0   'False
      End
      Begin VB.Menu LP 
         Caption         =   "REKAM MEDIS PASIEN"
         Index           =   43
         Visible         =   0   'False
      End
   End
   Begin VB.Menu I 
      Caption         =   "INFO"
      Index           =   5
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MAINMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tAtur
    sIPT As String
End Type

Dim tSet As tAtur

Private Sub ADM_Click(Index As Integer)
Select Case Index
    Case 1
        P002.Show 1
    Case 2
        Diagnosa.Show 1
    Case 3
        OBAT.Show 1
End Select
End Sub

Private Sub cmdCLOSE_Click()
Unload Me
LOGIN.Show
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

Text1 = "USER : " + Operator
Text2 = Date
Text3 = "Copyrighted 2008 - IPT"

SkinLabel1 = NTOKO
SkinLabel4 = NAlamat
SkinLabel5 = NMOtto
SkinLabel6 = NTelepon
Me.Left = 0
Me.Top = 0
End Sub

Private Sub I_Click(Index As Integer)
INFO.Show 1
End Sub

Private Sub LoadSaveAtur(bLoad As Boolean)
    Dim Counter As String
    Dim nFile As String
    Dim ff As Integer

    Counter = 50
    nFile = App.Path & "\NOVI.dat"
    ff = FreeFile

    If bLoad = True Then
    
        Open nFile For Binary Access Read As #ff
        Get #ff, , tSet
        Close #ff
        
        With tSet
            Text1.Text = .sIPT
        End With

    Else

        With tSet
            .sIPT = Counter
        End With
        
        If Dir(nFile, 1 Or 2 Or 4 Or 32) <> "" Then Kill nFile
        
        Open nFile For Binary Access Read Write As #ff
        Put #ff, , tSet
        Close #ff

    End If
End Sub

Private Sub LP_Click(Index As Integer)
Select Case Index
    Case 41
        LP01.Show 1
    Case 42
        LP02.Show 1
    Case 43
        LP03.Show 1
End Select
End Sub

Private Sub PD_Click(Index As Integer)
Select Case Index
    Case 11
        P001.Show 1
    Case 12
        P002.Show 1
    Case 13
        P003.Show 1
    Case 14
        P004A.Show 1
End Select
End Sub

Private Sub PP_Click(Index As Integer)
Select Case Index
    Case 21
        C001.Show 1
    Case 22
        C002.Show 1
End Select
End Sub

Private Sub TT_Click(Index As Integer)
Select Case Index
    Case 31
        PASS.Show 1
    Case 32
        SETTING.Show 1
End Select
End Sub
