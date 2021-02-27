Attribute VB_Name = "Module2"


Public Const UNIT As Integer = 120

Public Const DELIMITER As String = "!"
Public Const CHUNK_SIZE As Long = 30 * 1024

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
Public Const PS As String = "S"
Public Const PT As String = "T"
Public Const PQ As String = "Q" 'user cancels login or he logouts, so server should re-listen
Public Const PX As String = "X" 'error
Public Const PY As String = "Y" 'send it when server verifies client's identity
Public Const PZ As String = "Z" 'bulk ftp (entire folder)


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



Public Const PIE As Double = 22# / 7    'simulate the value of PI
Public Const A As Integer = 550 'how far from center along x-axis
Public Const b As Integer = 400 'how far from center along y-axis



Public Const PERMIT_VIEW_DRIVES As Long = 1
Public Const PERMIT_VIEW_DIRS As Long = 2
Public Const PERMIT_VIEW_FILES As Long = 4
Public Const PERMIT_DOWNLOAD As Long = 8
Public Const PERMIT_UPLOAD As Long = 16
Public Const PERMIT_UPLOAD_O As Long = 32
