VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "konfigurasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public NamaAplikasi As String
Public PathFileServer As String
Public PathFileLokal As String
Public Judul As String
Public Function GetCFG()
    PathFileLokal = "D:\App\test.exe"  'Path Lengkap FIle Lokal
    PathFileServer = "\\xxx.xxx.x.xxx\exeterbaru\test.exe" 'Path Lengkap FIle Server
End Function

