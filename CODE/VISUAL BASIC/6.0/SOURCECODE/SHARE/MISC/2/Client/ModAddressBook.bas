Attribute VB_Name = "ModAddressBook"
Option Explicit

Public Const ADB_FILE As String = "\Book.dat"

'Address Book File Structure
'---------------------------
'Remote host|_|Remote port|_|Description*_*

Public Function AddressBookQuery(ByVal IPAddress As String, ByVal RemotePort As String) As Boolean
Dim FF As Integer: FF = FreeFile
Dim sBuff() As String, sDat() As String, lLoop As Long, sTmp As String, sNowLoop As String
If Not FileExists(App.Path & ADB_FILE) Then Exit Function
If FileLen(App.Path & ADB_FILE) = 0 Then Exit Function

Open App.Path & ADB_FILE For Binary Access Read As #FF
    sTmp = Input(LOF(FF), FF)
Close #FF

If InStr(1, sTmp, vbNewLine) = 0 Then
    If InStr(1, sTmp, "|_|") = 0 Then Exit Function
    sDat() = Split(sTmp, "|_|")
    If LCase$(sDat(0)) = LCase$(IPAddress) And LCase$(sDat(1)) = LCase$(RemotePort) Then
        AddressBookQuery = True
        Exit Function
    Else
        AddressBookQuery = False
        Exit Function
    End If
Else
    sBuff() = Split(sTmp, vbNewLine)
    For lLoop = 0 To UBound(sBuff)
        sNowLoop = sBuff(lLoop)
        If Len(sNowLoop) > 0 Then
            If InStr(1, sNowLoop, "|_|") > 0 Then
                sDat() = Split(sNowLoop, "|_|")
                If LCase$(sDat(0)) = LCase$(IPAddress) And LCase$(sDat(1)) = LCase$(RemotePort) Then
                    AddressBookQuery = True
                    Exit For
                End If
            End If
        End If
    Next lLoop
End If
End Function

Public Sub AddressBookAdd(ByVal IPAddress As String, ByVal RemotePort As String, Optional ByVal Description As String = "[ None ]")
Dim FF As Integer: FF = FreeFile
If AddressBookQuery(IPAddress, RemotePort) Then Exit Sub
Open App.Path & ADB_FILE For Append As #FF
    Print #FF, IPAddress & "|_|" & RemotePort & "|_|" & Description
Close #FF
End Sub

Public Sub AddressBookRemove(ByVal IPAddress As String, ByVal RemotePort As String, Optional ByVal Description As String = Empty)
On Error Resume Next
Dim FF As Integer: FF = FreeFile
Dim sBuff() As String, sDat() As String, lLoop As Long, sTmp As String, sNowLoop As String, sNewData As String
If Not FileExists(App.Path & ADB_FILE) Then Exit Sub
If FileLen(App.Path & ADB_FILE) = 0 Then Exit Sub

Open App.Path & ADB_FILE For Binary Access Read As #FF
    sTmp = Input(LOF(FF), FF)
Close #FF

If InStr(1, sTmp, vbNewLine) = 0 Then
    If InStr(1, sTmp, "|_|") = 0 Then Exit Sub
    sDat() = Split(sTmp, "|_|")
    If LCase$(sDat(0)) = LCase$(IPAddress) And LCase$(sDat(1)) = LCase$(RemotePort) Then
        Call KillFile(App.Path & ADB_FILE)
        Exit Sub
    End If
Else
    sBuff() = Split(sTmp, vbNewLine)
    For lLoop = 0 To UBound(sBuff)
        sNowLoop = sBuff(lLoop)
        If Len(sNowLoop) > 0 Then
            sDat() = Split(sNowLoop, "|_|")
            If LCase$(sDat(0)) <> LCase$(IPAddress) And LCase$(sDat(1)) <> LCase$(RemotePort) And LCase$(sDat(2)) <> LCase$(Description) Then
                sNewData = sNewData & sNowLoop & vbNewLine
            End If
        End If
    Next lLoop
    Call KillFile(App.Path & ADB_FILE)
    Open App.Path & ADB_FILE For Append As #FF
        Print #FF, sNewData
    Close #FF
End If
End Sub

Public Sub AddressBookLoadListView(objListView As Object)
' IP, Port, Description
Dim FF As Integer: FF = FreeFile
Dim sBuff() As String, sDat() As String, lLoop As Long, sTmp As String, sNowLoop As String
If Not FileExists(App.Path & ADB_FILE) Then Exit Sub
If FileLen(App.Path & ADB_FILE) = 0 Then Exit Sub

With objListView
    .ListItems.Clear
    Open App.Path & ADB_FILE For Binary Access Read As #FF
        sTmp = Input(LOF(FF), FF)
    Close #FF

    If InStr(1, sTmp, vbNewLine) = 0 Then
        If InStr(1, sTmp, "|_|") = 0 Then Exit Sub
        sDat() = Split(sTmp, "|_|")
        .ListItems.Add , , sDat(0)
        .ListItems(.ListItems.Count).ListSubItems.Add , , sDat(1)
        .ListItems(.ListItems.Count).ListSubItems.Add , , sDat(2)
        Exit Sub
    Else
        sBuff() = Split(sTmp, vbNewLine)
        For lLoop = 0 To UBound(sBuff)
            sNowLoop = sBuff(lLoop)
            If Len(sNowLoop) > 0 Then
                If InStr(1, sNowLoop, "|_|") > 0 Then
                    sDat() = Split(sNowLoop, "|_|")
                    .ListItems.Add , , sDat(0)
                    .ListItems(.ListItems.Count).ListSubItems.Add , , sDat(1)
                    .ListItems(.ListItems.Count).ListSubItems.Add , , sDat(2)
                End If
            End If
        Next lLoop
    End If
End With
End Sub

Public Function FileExists(ByVal Filename As String) As Boolean
On Error Resume Next
FileExists = (Dir(Filename, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Sub KillFile(ByVal FilePath As String)
On Error Resume Next
Kill FilePath
End Sub

Public Sub RemoveListViewItem(objLV As Object, ByVal ItemText As String)
On Error Resume Next
Dim lLoop As Long
With objLV
    If .ListItems.Count = 0 Then Exit Sub
    For lLoop = 1 To .ListItems.Count
        If LCase$(.ListItems(lLoop)) = LCase$(ItemText) Then
            .ListItems.Remove lLoop
            Exit For
        End If
    Next lLoop
End With
End Sub
