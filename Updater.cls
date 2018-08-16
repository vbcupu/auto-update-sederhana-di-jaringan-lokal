VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Updater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function isAdaUpdate(strFileLokal, strFileServer As String) As Boolean
    Dim fso As FileSystemObject
    Dim verServer As String
    Dim verLokal As String
    
    Set fso = New FileSystemObject
    'Cek Versi File di Server
    verServer = fso.GetFileVersion(strFileServer)
    'Cek Versi File di Lokal
    verLokal = fso.GetFileVersion(strFileLokal)
    
    'Compare
    If verServer > verLokal Then
        isAdaUpdate = True
        Exit Function
    Else
        isAdaUpdate = False
    End If
End Function
Public Function Update(strFileLokal, strFileServer As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    fso.CopyFile strFileServer, strFileLokal, True
End Function
