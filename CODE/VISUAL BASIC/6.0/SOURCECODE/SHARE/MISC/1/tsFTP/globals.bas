Attribute VB_Name = "Module3"
Public initialFormWidth As Long
Public initialFormHeight As Long
Public initialtblWidth As Long
Public initialtblHeight As Long


Public myState As STATE
Public drivesLen As Long
Public drivesBuf As String

Public foldersLen As Long
Public foldersBuf As String

Public filesLen As Long
Public filesBuf As String



Public filesRequestPath As String

'the following are for ftp between client and server
'Public fileRequested As String
'Public fileSize As String

'ftpFolderPath stores the root path of the files you want to download
'from server
Public ftpFolderPath As String



Public ftpFileName As String
Public ftpFilePath As String
Public ftpFileSize As String
Public myFilePath As String
Public myFolderPath As String
'Public serverFilePath As String
Public myChunkSize As Long
Public theirChunkSize As Long
Public currFilePos As Long
Public counter As Integer 'counter for speed

Public allFiles() As String
Public filePointer As Long 'pointer pointing to the current path in allFiles


'Public first_request As Boolean


'keep track of # of spaces in front of the current dir
'the next level dirs need more spaces in front for distinction
Public numSpaces As Integer

Public expansionList() As String

Public myBasePath As String 'used in bulk ftp

'when use clicks download, he needs to choose option and pattern
Public tranOption As DOWNLOAD_TYPE
Public pattern As String
Public defaultFile As String
Public defaultDir As String


Public fs As New FileSystemObject

Public loginCancel As Boolean
Public loginName As String
Public loginPassword As String
Public loginRemember As Boolean


'isn't it annoying to have to pass parameter to a new form?
'i think it's possible to pass arg to Form.show but I am not sure
'i am fed up with searching in vain
Public fileForProperty As String

Public propertyType As Integer

Public fileFoundOption As FILE_FOUND_OPTION
Public newFilePath As String 'new file path filled in Form5
