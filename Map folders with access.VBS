On Error Resume Next

Set objNetwork = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")

' Remove todos os mapeamentos de pastas existentes
Set colDrives = objNetwork.EnumNetworkDrives
For i = 0 to colDrives.Count-1 Step 2
    objNetwork.RemoveNetworkDrive colDrives.Item(i), True, True
Next

' Lista de pastas a serem mapeadas
Dim arrFolders(3)
arrFolders(0) = "\\servidor\pasta1"
arrFolders(1) = "\\servidor\pasta2"
arrFolders(2) = "\\servidor\pasta3"

' Testa acesso a cada pasta da lista e mapeia apenas as pastas que o usuário tem acesso
For Each folder In arrFolders
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(folder) Then
        Set objFolder = objFSO.GetFolder(folder)
        If objFolder.Attributes And 1 Then ' Verifica se a pasta é oculta
            objFolder.Attributes = objFolder.Attributes Xor 1 ' Remove atributo de oculto
        End If
        If objFolder.SubFolders.Count > 0 Or objFolder.Files.Count > 0 Then
            ' Pasta não vazia, não é possível mapear
        Else
            ' Tenta mapear a pasta
            strDriveLetter = GetNextAvailableDrive()
            If strDriveLetter <> "" Then
                objNetwork.MapNetworkDrive strDriveLetter, folder, False, objNetwork.UserName, objNetwork.UserDomain
            End If
        End If
    End If
Next

' Função para obter a próxima letra de unidade disponível para mapeamento
Function GetNextAvailableDrive()
    For i = 67 To 90 ' Letras de unidade A a Z
        strDriveLetter = Chr(i) & ":"
        If Not objFSO.DriveExists(strDriveLetter) Then
            GetNextAvailableDrive = strDriveLetter
            Exit Function
        End If
    Next
    GetNextAvailableDrive = "" ' Nenhum drive disponível
End Function
