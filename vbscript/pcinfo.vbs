Set wshNetwork = CreateObject("WScript.Network")
computerName = wshNetwork.ComputerName
nombreArchivo = "\INFO_" + computerName + ".txt"
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objFile=objFSO.CreateTextFile("."+nombreArchivo,2,true)
Function line(text)
    objFile.WriteLine(text)
End Function



Dim ObjExec
Dim strFromProc
Set oShell = WScript.CreateObject("WScript.Shell")

Function directorioActual()
    Dim FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    directorioActual = FSO.GetAbsolutePathName(".")
End Function

fechaHoy = CStr(Date + Time)
separador=vbNewLine + "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" + vbNewLine
lineaTitulo="-----------------------------"

Function ConvertToKey(Key)
    Const KeyOffset = 52
    i = 28
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    Do
    Cur = 0
    x = 14
    Do
    Cur = Cur * 256
    Cur = Key(x + KeyOffset) + Cur
    Key(x + KeyOffset) = (Cur \ 24) And 255
    Cur = Cur Mod 24
    x = x -1
    Loop While x >= 0
    i = i -1
    KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
    If (((29 - i) Mod 6) = 0) And (i <> -1) Then
    i = i -1
    KeyOutput = "-" & KeyOutput
    End If
    Loop While i >= 0
    ConvertToKey = KeyOutput
End Function


Function obtenerComando(ejecucion)
    Set ObjExec = oShell.Exec("cmd.exe /c " & ejecucion)
    obtenerComando = ObjExec.stdOut.ReadAll
End Function

dirActual = directorioActual()



' :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: '
'                                   PRINCIPAL                                  '
' :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: '

'-------------
' ENCABEZADO
'-------------

line(separador)
line("Usuario: ")
line("Departamento: ")
line("Computadora de consulta: " + computerName)
line("Fecha de consulta: " + fechaHoy)
line(separador)

'----------------------
' LICENCIA DE WINDOWS
'----------------------

line(lineaTitulo)
line("LICENCIA DE WINDOWS")
line(lineaTitulo + vbNewline)

texto = obtenerComando("wmic path softwareLicensingService get OA3xOriginalProductKey")
texto = Split(texto)
licenciaWmic = texto(2)
line("Servicio de licenciamiento (wmic): " + licenciaWmic)
llaveRegistro = ConvertToKey(oShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
line("Registro de windows: " & llaveRegistro)




'--------
' FINAL
'--------

line("----------------------------- final del archivo ----------------------------")
tituloFinal = "Ejecucion terminada"
mensajeFinal = "Datos de " + computerName + " obtenidos" + vbNewLine + "Ver: " + dirActual + nombreArchivo
MsgBox mensajeFinal,, tituloFinal


' ------------------------------ principal (fin) ----------------------------- '

