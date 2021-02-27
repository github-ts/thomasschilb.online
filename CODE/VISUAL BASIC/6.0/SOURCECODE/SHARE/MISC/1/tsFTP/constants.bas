Attribute VB_Name = "Module2"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Const UNIT As Long = 120  'unit in twips for offset between each adjacent control

Public Const DELIMITER As String = "!"
Public Const CHUNK_SIZE As Long = 30 * 1024

'# of spaces between one level of dir and the next level of dir
Public Const DIR_OFF As Integer = 4

Public Const PA As String = "A"
Public Const PB As String = "B" 'from server, if the requested file exists and ready to send
Public Const PC As String = "C"
Public Const PD As String = "D" 'want drive list
Public Const PE As String = "E" 'from server, if the requested file doesn't exist
Public Const PF As String = "F" 'want folder list
Public Const PG As String = "G" 'user wants to login, send server user's login name and password
Public Const PI As String = "I" 'want file list
Public Const PJ As String = "J"
Public Const PL As String = "L" 'request file
Public Const PN As String = "N"
Public Const PP As String = "P"
Public Const PR As String = "R" 'request all files in a folder
Public Const PS As String = "S" 'send a file
Public Const PT As String = "T" 'send a folder
Public Const PQ As String = "Q" 'user cancels login, so server should re-listen
Public Const PX As String = "X" 'error
Public Const PY As String = "Y" 'user's identity is verified, client got this msg so that it can continue
Public Const PZ As String = "Z"


Public Enum RES

    NONE = -1
    SUCCESS = 0
    UNDER = 1
    OVER = 2
    
End Enum

Public Enum DOWNLOAD_TYPE
    TRAN_CANCEL = -2
    TRAN_FILE = -1
    TRAN_DIR_R = 0
    TRAN_DIR = 1
 
End Enum

Public Enum FILE_FOUND_OPTION
    OVERWRITE = 1
    OVERWRITE_ALL = 2
    DO_NOT_OVERWRITE = 3
    UNDECIDED = 4
    DISCARD = 5
    DISCARD_ALL = 6
End Enum
    
Public Enum STATE
    DORMANT = -1 'not connected
  '  INIT = 0    'attempt to connect or be connected
    ATTEMPT_CONNECT = 0
    CONN = 1    'connected
    ATTEMPT_SEND = 2
    SEND_FILE = 3  'file transfer is in progree
    RECV_FILE = 4
    WANT_DRIVES = 5
    WANT_DIRS = 6
    WANT_FILES = 7
    WANT_DIRS_A = 8    'want folders and append them to the list
    REQ_FILE = 9
    REQ_DIR = 10
    
    ATTEMPT_RECV_BULK = 11
    RECV_DIR = 12
    SEND_DIR = 13
    
End Enum



Public Declare Function SendMessageByNum Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
        wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const WANT_FILE_PROPERTY = 1
Public Const WANT_FOLDER_PROPERTY = 2
