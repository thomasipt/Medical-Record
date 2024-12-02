VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INFO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFO"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4695
      OleObjectBlob   =   "INFO.frx":0000
      Top             =   7770
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
      Left            =   3795
      TabIndex        =   0
      Top             =   4920
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":0234
      TabIndex        =   1
      Top             =   5610
      Width           =   2970
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   285
      Left            =   870
      OleObjectBlob   =   "INFO.frx":02B7
      TabIndex        =   2
      Top             =   120
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":03C2
      TabIndex        =   3
      Top             =   120
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":041D
      TabIndex        =   4
      Top             =   480
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":0478
      TabIndex        =   5
      Top             =   840
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":04D3
      TabIndex        =   6
      Top             =   1425
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":052E
      TabIndex        =   7
      Top             =   2280
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":0589
      TabIndex        =   8
      Top             =   2865
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   285
      Left            =   870
      OleObjectBlob   =   "INFO.frx":05E4
      TabIndex        =   9
      Top             =   480
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   510
      Left            =   870
      OleObjectBlob   =   "INFO.frx":06E1
      TabIndex        =   10
      Top             =   840
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   780
      Left            =   870
      OleObjectBlob   =   "INFO.frx":0854
      TabIndex        =   11
      Top             =   1425
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   510
      Left            =   870
      OleObjectBlob   =   "INFO.frx":0A5B
      TabIndex        =   12
      Top             =   2280
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   285
      Left            =   870
      OleObjectBlob   =   "INFO.frx":0BCA
      TabIndex        =   13
      Top             =   2865
      Width           =   8550
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   285
      Left            =   6450
      OleObjectBlob   =   "INFO.frx":0CD7
      TabIndex        =   14
      Top             =   4305
      Width           =   2880
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   285
      Left            =   240
      OleObjectBlob   =   "INFO.frx":0D52
      TabIndex        =   15
      Top             =   3225
      Width           =   450
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   510
      Left            =   870
      OleObjectBlob   =   "INFO.frx":0DAD
      TabIndex        =   16
      Top             =   3225
      Width           =   8550
   End
End
Attribute VB_Name = "INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
End Sub
