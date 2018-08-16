VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIMRS- Update Checker"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   6255
      TabIndex        =   3
      Top             =   0
      Width           =   6255
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtlog 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KLIK ------->"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cfg As konfigurasi
Dim Update As Updater
Private Sub Form_Activate()
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
End Sub
Private Sub cmdUpdate_Click()
    Call Update.Update(cfg.PathFileLokal, cfg.PathFileServer)
    Shell cfg.PathFileLokal, vbNormalFocus
    Unload Me
End Sub
Private Sub Form_Load()
    Set cfg = New konfigurasi
    Set Update = New Updater
    cfg.GetCFG

    If Update.isAdaUpdate(cfg.PathFileLokal, cfg.PathFileServer) = True Then
        Label1.Visible = True
        cmdUpdate.Visible = True
        cmdUpdate.Enabled = True
        txtlog.text = "Ada Update Aplikasi terbaru, Tutup Semua Aplikasi SIMRS yang masih terbuka, Kemudian Klik Update"
    Else
        Shell cfg.PathFileLokal, vbNormalFocus
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
