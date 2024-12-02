VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LP03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REKAM MEDIS PASIEN"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1530
      TabIndex        =   4
      Text            =   "1"
      Top             =   120
      Width           =   2040
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
      Left            =   1530
      TabIndex        =   3
      Text            =   "2"
      Top             =   525
      Width           =   2805
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
      Left            =   1530
      TabIndex        =   2
      Text            =   "3"
      Top             =   930
      Width           =   6090
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   7830
      TabIndex        =   1
      Top             =   135
      Width           =   1125
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
      Top             =   5085
      Width           =   9330
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "LP03.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "LP03.frx":0234
      TabIndex        =   5
      Top             =   210
      Width           =   1185
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "LP03.frx":028F
      TabIndex        =   6
      Top             =   615
      Width           =   1185
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   195
      Left            =   240
      OleObjectBlob   =   "LP03.frx":02EE
      TabIndex        =   7
      Top             =   1020
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3480
      Left            =   120
      TabIndex        =   8
      Top             =   1425
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6138
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "LP03"
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
