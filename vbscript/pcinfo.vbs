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
lineaTitulo   ="--------------------------------"
lineaSubTitulo="'''''''''''''''''''''''"

Function ObtenerLlaveLicencia(Key)
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
    ObtenerLlaveLicencia = KeyOutput
End Function


Function obtenerComando(ejecucion)
    Set ObjExec = oShell.Exec("cmd.exe /c " & ejecucion)
    obtenerComando = ObjExec.stdOut.ReadAll
End Function

Function convertirBytes(bytes)
	Dim sufijos(9)
	sufijos(0) = " B" 
	sufijos(1) = " KB"
	sufijos(2) = " MB"
	sufijos(3) = " GB"
	sufijos(4) = " TB"
	sufijos(5) = " PT"
	sufijos(6) = " EB"
	sufijos(7) = " ZB"
	sufijos(8) = " YB"
	factor     = 1024
	paso       = bytes
	indice     = 0
	Do
		If paso >= factor Then:
			paso = paso / factor
			indice = indice + 1
	Loop While paso >= factor
	convertirBytes = Round(paso, 2) & sufijos(indice)
End Function

dirActual = directorioActual()

Function lenguajeEspIngIta(numero)
	Select Case numero
		Case 9
			lenguajeEspIngIta = "Ingles"
		Case 1033
			lenguajeEspIngIta = "Ingles Estados Unidos"
		Case 1034
			lenguajeEspIngIta = "Espanol Tradicional"
		Case 1040
			lenguajeEspIngIta = "Italiano Italia"
		Case 2057
			lenguajeEspIngIta = "Ingles Reino Unido"
		Case 2058
			lenguajeEspIngIta = "Espanol Mexico"
		Case 2064
			lenguajeEspIngIta = "Italian Suiza"
		Case 3081
			lenguajeEspIngIta = "Ingles Australia"
		Case 3082
			lenguajeEspIngIta = "Espanol Internacional"
		Case 4105
			lenguajeEspIngIta = "Ingles Canada"
		Case 4106
			lenguajeEspIngIta = "Espanol Guatemala"
		Case 5129
			lenguajeEspIngIta = "Ingles Nueva Zelanda"
		Case 5130
			lenguajeEspIngIta = "Espanol Costa Rica"
		Case 6153
			lenguajeEspIngIta = "Ingles Irlanda"
		Case 6154
			lenguajeEspIngIta = "Espanol Panama"
		Case 7177
			lenguajeEspIngIta = "Ingles Sudafrica"
		Case 7178
			lenguajeEspIngIta = "Espanol Republica Dominicana"
		Case 8201
			lenguajeEspIngIta = "Ingles Jamaica"
		Case 8202
			lenguajeEspIngIta = "Espanol Venezuela"
		Case 9226
			lenguajeEspIngIta = "Espanol Colombia"
		Case 10249
			lenguajeEspIngIta  "Ingles Belize"
		Case 10250
			lenguajeEspIngIta  "Espanol Peru"
		Case 11273
			lenguajeEspIngIta  "Ingles Trinidad"
		Case 11274
			lenguajeEspIngIta  "Espanol Argentina"
		Case 12298
			lenguajeEspIngIta  "Espanol Ecuador"
		Case 13322
			lenguajeEspIngIta  "Espanol Chile"
		Case 14346
			lenguajeEspIngIta  "Espanol Uruguay"
		Case 15370
			lenguajeEspIngIta  "Espanol Paraguay"
		Case 16394
			lenguajeEspIngIta  "Espanol Bolivia"
		Case 17418
			lenguajeEspIngIta  "Espanol El Salvador"
		Case 18442
			lenguajeEspIngIta  "Espanol Honduras"
		Case 19466
			lenguajeEspIngIta  "Espanol Nicaragua"
		Case 20490
			lenguajeEspIngIta  "Espanol Puerto Rico"
		Case Else
			lenguajeEspIngIta = ""
	End Select
End Function

Function suite(numero)
	Select Case numero
		Case 1 
			suite = "Microsoft Small Business Server (upgraded)"
		Case 2 
			suite = "Windows Server 2008 Enterprise"
		Case 4 
			suite = "Windows BackOffice components"
		Case 8 
			suite = "Communication Server"
		Case 16 
			suite = "Terminal Services"
		Case 32 
			suite = "Microsoft Small Business Server with restrictive client license"
		Case 64 
			suite = "Windows Embedded"
		Case 128 
			suite = "Datacenter"
		Case 256 
			suite = "Terminal Services (only one interactive session)"
		Case 512 
			suite = "Windows Home Edition"
		Case 1024 
			suite = "Web Server Edition"
		Case 8192 
			suite = "Storage Server Edition"
		Case 16384 
			suite = "Compute Cluster Edition"
		Case Else
			suite = ""
	End Select
End Function

Function tipoProducto(numero)
	Select Case numero
		Case 1 
			tipoProducto = "Work Station"
		Case 2 
			tipoProducto = "Domain Controller"
		Case 3
			tipoProducto = "Server"
		Case Else
			tipoProducto = ""
	End Select
End Function

Function tipoOs(numero)
	Select Case numero
		Case 0
			tipoOs = "Unknown"
		Case 1
			tipoOs = "Other"
		Case 2
			tipoOs = "MACOS"
		Case 3
			tipoOs = "ATTUNIX"
		Case 4
			tipoOs = "DGUX"
		Case 5
			tipoOs = "DECNT"
		Case 6
			tipoOs = "Digital Unix"
		Case 7
			tipoOs = "OpenVMS"
		Case 8
			tipoOs = "HPUX"
		Case 9
			tipoOs = "AIX"
		Case 10
			tipoOs = "MVS"
		Case 11
			tipoOs = "OS400"
		Case 12
			tipoOs = "OS/2"
		Case 13
			tipoOs = "JavaVM"
		Case 14
			tipoOs = "MSDOS"
		Case 15
			tipoOs = "WIN3x"
		Case 16
			tipoOs = "WIN95"
		Case 17
			tipoOs = "WIN98"
		Case 18
			tipoOs = "WINNT"
		Case 19
			tipoOs = "WINCE"
		Case 20
			tipoOs = "NCR3000"
		Case 21
			tipoOs = "NetWare"
		Case 22
			tipoOs = "OSF"
		Case 23
			tipoOs = "DC/OS"
		Case 24
			tipoOs = "Reliant UNIX"
		Case 25
			tipoOs = "SCO UnixWare"
		Case 26
			tipoOs = "SCO OpenServer"
		Case 27
			tipoOs = "Sequent"
		Case 28
			tipoOs = "IRIX"
		Case 29
			tipoOs = "Solaris"
		Case 30
			tipoOs = "SunOS"
		Case 31
			tipoOs = "U6000"
		Case 32
			tipoOs = "ASERIES"
		Case 33
			tipoOs = "TandemNSK"
		Case 34
			tipoOs = "TandemNT"
		Case 35
			tipoOs = "BS2000"
		Case 36
			tipoOs = "LINUX"
		Case 37
			tipoOs = "Lynx"
		Case 38
			tipoOs = "XENIX"
		Case 39
			tipoOs = "VM/ESA"
		Case 40
			tipoOs = "Interactive UNIX"
		Case 41
			tipoOs = "BSDUNIX"
		Case 42
			tipoOs = "FreeBSD"
		Case 43
			tipoOs = "NetBSD"
		Case 44
			tipoOs = "GNU Hurd"
		Case 45
			tipoOs = "OS9"
		Case 46
			tipoOs = "MACH Kernel"
		Case 47
			tipoOs = "Inferno"
		Case 48
			tipoOs = "QNX"
		Case 49
			tipoOs = "EPOC"
		Case 50
			tipoOs = "IxWorks"
		Case 51
			tipoOs = "VxWorks"
		Case 52
			tipoOs = "MiNT"
		Case 53
			tipoOs = "BeOS"
		Case 54
			tipoOs = "HP MPE"
		Case 55
			tipoOs = "NextStep"
		Case 56
			tipoOs = "PalmPilot"
		Case 57
			tipoOs = "Rhapsody"
		Case 58
			tipoOs = "Windows 2000"
		Case 59
			tipoOs = "Dedicated"
		Case 60
			tipoOs = "OS/390"
		Case 61
			tipoOs = "VSE"
		Case 62
			tipoOs = "TPF"
		Case Else
			tipoOs = ""
	End Select
End Function			

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

'--------------------
' SISTEMA OPERATIVO
'--------------------

line(lineaTitulo)
line("SISTEMAS OPERATIVOS")
line(lineaTitulo + vbNewline)



Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

For Each os in oss
    dtmConvertedDate.Value = os.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate

    line(vbNewLine + "Version (nombre):    " + os.Caption)
    line("Manufacturador:      " & os.Manufacturer)
    line("Arquitectura:        " & os.OSArchitecture)
    line("Fecha instalacion:   " & dtmInstallDate)
    line("SKU:                 " & os.OperatingSystemSKU)
    line("Tipo producto:       " & tipoProducto(os.ProductType))
    line("Codigo del pais:     " & "+" & os.CountryCode)
    line("Lenguaje:            " & lenguajeEspIngIta(os.OsLanguage))
    line("Suite:               " & suite(os.OSProductSuite))
    line("Tipo:                " & tipoOs(os.OSType))
    line("Otra descripcion:    " & os.OtherTypeDescription)
    line("Es portable:         " & os.PortableOperatingSystem)
    line("Es el principal:     " & os.Primary)
    line("Numero serial:       " & os.SerialNumber)
    line("Version:             " & os.Version)
    line("Build:               " & os.BuildNumber)
    line("Service pack:        " & os.ServicePackMajorVersion & "." & os.ServicePackMinorVersion)
    line("Tipo build:          " & os.BuildType)
    line("Estado del sistema:  " & os.Status)
    line("Memoria virtual:     " & convertirBytes(os.TotalVirtualMemorySize))
    line("Memoria visible:     " & convertirBytes(os.TotalVisibleMemorySize))
    line("Disp arranque:       " & os.BootDevice)
    line("Usuario registrado:  " & os.RegisteredUser)
Next

texto = obtenerComando("wmic path softwareLicensingService get OA3xOriginalProductKey")
texto = Split(texto)
licenciaWmic = texto(2)
line("Licencia (wmic):     " + licenciaWmic)
llaveRegistro = ObtenerLlaveLicencia(oShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
line("Licencia (registro): " & llaveRegistro)




'--------
' FINAL
'--------

line("----------------------------- final del archivo ----------------------------")
tituloFinal = "Ejecucion terminada"
mensajeFinal = "Datos de " + computerName + " obtenidos" + vbNewLine + "Ver: " + dirActual + nombreArchivo
MsgBox mensajeFinal,, tituloFinal


' ------------------------------ principal (fin) ----------------------------- '