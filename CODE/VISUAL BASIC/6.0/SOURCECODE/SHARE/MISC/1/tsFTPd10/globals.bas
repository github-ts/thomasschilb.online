Attribute VB_Name = "Module3"
Public initialFormWidth As Long
Public initialFormHeight As Long

Public myState As STATE


'the following are for ftp between client and server
Public ftpFileName As String
Public ftpFilePath As String
Public ftpFileSize As String
Public myChunkSize As Long
Public theirChunkSize As Long
Public currFilePos As Long
Public counter As Integer 'counter for speed


Public allFiles() As String
'Public ftpFilesPath As String
Public filePointer As Long 'pointer pointing to the current path in allFiles

Public ftpFinished As Boolean

Public ftpFolderPath As String
Public myFilePath As String

Public tg As Long


Public fs As New FileSystemObject


Public logoDegree As Integer 'the degree in the logo circle (0 to 359)

Public clientName As String
Public clientPassword As String
'client may not be allowed to perform certain actions
Public clientPermission As Long

Public AccountFile As String
