' Constants (taken from WinReg.h)
Const HKEY_CLASSES_ROOT     = &H80000000
Const HKEY_CURRENT_USER     = &H80000001
Const HKEY_LOCAL_MACHINE    = &H80000002
Const HKEY_USERS            = &H80000003
Const REG_SZ                = 1
Const REG_EXPAND_SZ         = 2
Const REG_BINARY            = 3
Const REG_DWORD             = 4
Const REG_MULTI_SZ          = 7
Dim intServiceType, intStartupType, strDisplayName , result
Set WshShell    = CreateObject("Wscript.Shell")
Set objFSO      = Wscript.CreateObject("Scripting.FilesystemObject")
strNow          = Month(Date) & Day(Date) & Year(Date) 'Deffenir data e hora atual para salvarmos no nome do script e nao só
strBackupFile   = WshShell.CurrentDirectory & "\Services_Backup_" & strNow & ".REG" 'Criar Backup file na pasta atual
Set b           = objFSO.CreateTextFile ( strBackupFile , True ) 'defenir b como objecto que vai guardar e escrever tudo que selecionarmos
b.WriteLine "Windows Registry Editor Version 5.00"
b.WriteBlankLines 1
b.WriteLine ";Services Startup Configuration Backup " & Now
b.WriteBlankLines 1
' Chose computer name, registry tree and key path
strComputer = "." ' Use . for current machine
hDefKey     = HKEY_LOCAL_MACHINE
strKeyPath  = "SYSTEM\CurrentControlSet\Services"
FullKeyPath = "HKEY_LOCAL_MACHINE\" & strKeyPath
' Connect to registry provider on target machine with current user
Set oReg    = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
oReg.EnumKey hDefKey , strKeyPath, arrSubKeys
result = concatenate( b , FullKeyPath , arrSubKeys, "2" ) ' Chamar a Function para Auto
result = concatenate( b , FullKeyPath , arrSubKeys, "3" ) ' Chamar a Function para Manual
result = concatenate( b , FullKeyPath , arrSubKeys, "4" ) ' Chamar a Function para Disable
' funcao para achar todas as chaves do serviços e savar as que queremos
Function concatenate( bb , FullKeyPath , arrSubKeys, StartupType )
    For Each ServiceNames In arrSubKeys
        IsWin32Service = False
        intStartupType = ""
        On Error Resume Next
        intServiceType = WshShell.RegRead( FullKeyPath & "\" & ServiceNames & "\type" )
        On Error Goto 0
        If intServiceType   = 16 Or intServiceType = 32 Or intServiceType = 256 Then
            IsWin32Service  = True
            On Error Resume Next
            intStartupType  = Trim( WshShell.RegRead( FullKeyPath & "\" & ServiceNames & "\Start" ) )
            If intStartupType = StartupType Then
                intStartupType  = "0000000" & intStartupType
                bb.WriteLine "[" & FullKeyPath & "\" & ServiceNames & "]"
                bb.WriteLine chr(34) & "Start" & Chr(34) & "=dword:" & StartupType
                bb.WriteBlankLines 1
            End If
            On Error Goto 0
        End If
    Next
End Function
b.Close                                     'Fechar ficheiro
WshShell.Run "Notepad " & strBackupFile     'Abrir Bloco de Notas
Set objFSO      = Nothing
Set WshShell    = Nothing