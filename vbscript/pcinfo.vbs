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
			suite = numero
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
			tipoProducto = numero
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
			tipoOs = numero
	End Select
End Function

Function estadoDeBooteoChasis(numero)
	Select Case numero
		Case 1 
			estadoDeBooteoChasis = "Otro"
		Case 2 
			estadoDeBooteoChasis = "Desconocido"
		Case 3
			estadoDeBooteoChasis = "Seguro"
		Case 4
			estadoDeBooteoChasis = "Alerta"
		Case 5
			estadoDeBooteoChasis = "Critico"
		Case 6
			estadoDeBooteoChasis = "Irrecuperable"
		Case Else
			estadoDeBooteoChasis = numero
	End Select
End Function
	
Function rolDominio(numero)
	Select Case numero
		Case 0 
			rolDominio = "Estacion de trabajo independiente"
		Case 1 
			rolDominio = "Estacion de trabajo miembro"
		Case 2
			rolDominio = "Servidor independiente"
		Case 3
			rolDominio = "Servidor miembro"
		Case 4
			rolDominio = "Controlador de dominio de respaldo"
		Case 5
			rolDominio = "Controlador primario de dominio"
		Case Else
			rolDominio = numero
	End Select
End Function

Function tipoPc(numero)
	Select Case numero
		Case 0 
			tipoPc = "Sin especificar"
		Case 1 
			tipoPc = "Escritorio"
		Case 2
			tipoPc = "Movil"
		Case 3
			tipoPc = "Estacion de trabajo"
		Case 4
			tipoPc = "Servidor empresarial"
		Case 5
			tipoPc = "Servidor SOHO"
		Case 6
			tipoPc = "Appliance PC"
		Case 7
			tipoPc = "Servidor de rendimiento"
		Case 8
			tipoPc = "Maximo"
		Case Else
			tipoPc = numero
	End Select
End Function

Function tipoPcX(numero)
	Select Case numero
		Case 0 
			tipoPcX = "Sin especificar"
		Case 1 
			tipoPcX = "Escritorio"
		Case 2
			tipoPcX = "Movil"
		Case 3
			tipoPcX = "Estacion de trabajo"
		Case 4
			tipoPcX = "Servidor empresarial"
		Case 5
			tipoPcX = "Servidor SOHO"
		Case 6
			tipoPcX = "Appliance PC"
		Case 7
			tipoPcX = "Servidor de rendimiento"
		Case 8
			tipoPcX = "Pizarra"
		Case 9
			tipoPcX = "Maximo"
		Case Else
			tipoPcX = numero
	End Select
End Function

Function arquitecturaCPU(numero)
	Select Case numero
		Case 0 
			arquitecturaCPU = "x86"
		Case 1 
			arquitecturaCPU = "MIPS"
		Case 2
			arquitecturaCPU = "Alpha"
		Case 3
			arquitecturaCPU = "PowerPC"
		Case 5
			arquitecturaCPU = "ARM"
		Case 6
			arquitecturaCPU = "ia64"
		Case 9
			arquitecturaCPU = "x64"
		Case 12
			arquitecturaCPU = "ARM64"
		Case Else
			arquitecturaCPU = numero
	End Select
End Function

Function disponibilidadCPU(numero)
	Select Case numero
		Case 1
			disponibilidadCPU = "Other" 
		Case 2
			disponibilidadCPU = "Unknown" 
		Case 3
			disponibilidadCPU = "Running/Full Power" 
		Case 4
			disponibilidadCPU = "Warning" 
		Case 5
			disponibilidadCPU = "In Test" 
		Case 6
			disponibilidadCPU = "Not Applicable" 
		Case 7
			disponibilidadCPU = "Power Off" 
		Case 8
			disponibilidadCPU = "Off Line" 
		Case 9
			disponibilidadCPU = "Off Duty" 
		Case 10
			disponibilidadCPU = "Degraded" 
		Case 11
			disponibilidadCPU = "Not Installed" 
		Case 12
			disponibilidadCPU = "Install Error" 
		Case 13
			disponibilidadCPU = "Power Save - Unknown" 
		Case 14
			disponibilidadCPU = "Power Save - Low Power Mode" 
		Case 15
			disponibilidadCPU = "Power Save - Standby" 
		Case 16
			disponibilidadCPU = "Power Cycle" 
		Case 17
			disponibilidadCPU = "Power Save - Warning" 
		Case 18
			disponibilidadCPU = "Paused" 
		Case 19
			disponibilidadCPU = "Not Ready" 
		Case 20
			disponibilidadCPU = "Not Configured" 
		Case 21
			disponibilidadCPU = "Quiesced" 
		Case Else
			disponibilidadCPU = numero	
	End Select
End Function

Function familaCPU(numero)
	Select Case numero
		Case 1
			familaCPU = "Other" 
		Case 2
			familaCPU = "Unknown" 
		Case 3
			familaCPU = "8086" 
		Case 4
			familaCPU = "80286" 
		Case 5
			familaCPU = "80386" 
		Case 6
			familaCPU = "80486" 
		Case 7
			familaCPU = "8087" 
		Case 8
			familaCPU = "80287" 
		Case 9
			familaCPU = "80387" 
		Case 10
			familaCPU = "80487" 
		Case 11
			familaCPU = "Pentium(R) brand" 
		Case 12
			familaCPU = "Pentium(R) Pro" 
		Case 13
			familaCPU = "Pentium(R) II" 
		Case 14
			familaCPU = "Pentium(R) processor with MMX(TM) technology" 
		Case 15
			familaCPU = "Celeron(TM)" 
		Case 16
			familaCPU = "Pentium(R) II Xeon(TM)" 
		Case 17
			familaCPU = "Pentium(R) III" 
		Case 18
			familaCPU = "M1 Family" 
		Case 19
			familaCPU = "M2 Family" 
		Case 20
			familaCPU = "Intel(R) Celeron(R) M processor" 
		Case 21
			familaCPU = "Intel(R) Pentium(R) 4 HT processor" 
		Case 24
			familaCPU = "K5 Family" 
		Case 25
			familaCPU = "K6 Family" 
		Case 26
			familaCPU = "K6-2" 
		Case 27
			familaCPU = "K6-3" 
		Case 28
			familaCPU = "AMD Athlon(TM) Processor Family" 
		Case 29
			familaCPU = "AMD(R) Duron(TM) Processor" 
		Case 30
			familaCPU = "AMD29000 Family" 
		Case 31
			familaCPU = "K6-2+" 
		Case 32
			familaCPU = "Power PC Family" 
		Case 33
			familaCPU = "Power PC 601" 
		Case 34
			familaCPU = "Power PC 603" 
		Case 35
			familaCPU = "Power PC 603+" 
		Case 36
			familaCPU = "Power PC 604" 
		Case 37
			familaCPU = "Power PC 620" 
		Case 38
			familaCPU = "Power PC X704" 
		Case 39
			familaCPU = "Power PC 750" 
		Case 40
			familaCPU = "Intel(R) Core(TM) Duo processor" 
		Case 41
			familaCPU = "Intel(R) Core(TM) Duo mobile processor" 
		Case 42
			familaCPU = "Intel(R) Core(TM) Solo mobile processor" 
		Case 43
			familaCPU = "Intel(R) Atom(TM) processor" 
		Case 48
			familaCPU = "Alpha Family" 
		Case 49
			familaCPU = "Alpha 21064" 
		Case 50
			familaCPU = "Alpha 21066" 
		Case 51
			familaCPU = "Alpha 21164" 
		Case 52
			familaCPU = "Alpha 21164PC" 
		Case 53
			familaCPU = "Alpha 21164a" 
		Case 54
			familaCPU = "Alpha 21264" 
		Case 55
			familaCPU = "Alpha 21364" 
		Case 56
			familaCPU = "AMD Turion(TM) II Ultra Dual-Core Mobile M Processor Family" 
		Case 57
			familaCPU = "AMD Turion(TM) II Dual-Core Mobile M Processor Family" 
		Case 58
			familaCPU = "AMD Athlon(TM) II Dual-Core Mobile M Processor Family" 
		Case 59
			familaCPU = "AMD Opteron(TM) 6100 Series Processor" 
		Case 60
			familaCPU = "AMD Opteron(TM) 4100 Series Processor" 
		Case 64
			familaCPU = "MIPS Family" 
		Case 65
			familaCPU = "MIPS R4000" 
		Case 66
			familaCPU = "MIPS R4200" 
		Case 67
			familaCPU = "MIPS R4400" 
		Case 68
			familaCPU = "MIPS R4600" 
		Case 69
			familaCPU = "MIPS R10000" 
		Case 80
			familaCPU = "SPARC Family" 
		Case 81
			familaCPU = "SuperSPARC" 
		Case 82
			familaCPU = "microSPARC II" 
		Case 83
			familaCPU = "microSPARC IIep" 
		Case 84
			familaCPU = "UltraSPARC" 
		Case 85
			familaCPU = "UltraSPARC II" 
		Case 86
			familaCPU = "UltraSPARC IIi" 
		Case 87
			familaCPU = "UltraSPARC III" 
		Case 88
			familaCPU = "UltraSPARC IIIi" 
		Case 96
			familaCPU = "68040" 
		Case 97
			familaCPU = "68xxx Family" 
		Case 98
			familaCPU = "68000" 
		Case 99
			familaCPU = "68010" 
		Case 100
			familaCPU = "68020" 
		Case 101
			familaCPU = "68030" 
		Case 112
			familaCPU = "Hobbit Family" 
		Case 120
			familaCPU = "Crusoe(TM) TM5000 Family" 
		Case 121
			familaCPU = "Crusoe(TM) TM3000 Family" 
		Case 122
			familaCPU = "Efficeon(TM) TM8000 Family" 
		Case 128
			familaCPU = "Weitek" 
		Case 130
			familaCPU = "Itanium(TM) Processor" 
		Case 131
			familaCPU = "AMD Athlon(TM) 64 Processor Family" 
		Case 132
			familaCPU = "AMD Opteron(TM) Processor Family" 
		Case 133
			familaCPU = "AMD Sempron(TM) Processor Family" 
		Case 134
			familaCPU = "AMD Turion(TM) 64 Mobile Technology" 
		Case 135
			familaCPU = "Dual-Core AMD Opteron(TM) Processor Family" 
		Case 136
			familaCPU = "AMD Athlon(TM) 64 X2 Dual-Core Processor Family" 
		Case 137
			familaCPU = "AMD Turion(TM) 64 X2 Mobile Technology" 
		Case 138
			familaCPU = "Quad-Core AMD Opteron(TM) Processor Family" 
		Case 139
			familaCPU = "Third-Generation AMD Opteron(TM) Processor Family" 
		Case 140
			familaCPU = "AMD Phenom(TM) FX Quad-Core Processor Family" 
		Case 141
			familaCPU = "AMD Phenom(TM) X4 Quad-Core Processor Family" 
		Case 142
			familaCPU = "AMD Phenom(TM) X2 Dual-Core Processor Family" 
		Case 143
			familaCPU = "AMD Athlon(TM) X2 Dual-Core Processor Family" 
		Case 144
			familaCPU = "PA-RISC Family" 
		Case 145
			familaCPU = "PA-RISC 8500" 
		Case 146
			familaCPU = "PA-RISC 8000" 
		Case 147
			familaCPU = "PA-RISC 7300LC" 
		Case 148
			familaCPU = "PA-RISC 7200" 
		Case 149
			familaCPU = "PA-RISC 7100LC" 
		Case 150
			familaCPU = "PA-RISC 7100" 
		Case 160
			familaCPU = "V30 Family" 
		Case 161
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 3200 Series" 
		Case 162
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 3000 Series" 
		Case 163
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 5300 Series" 
		Case 164
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 5100 Series" 
		Case 165
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 5000 Series" 
		Case 166
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor LV" 
		Case 167
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor ULV" 
		Case 168
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 7100 Series" 
		Case 169
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 5400 Series" 
		Case 170
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor" 
		Case 171
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 5200 Series" 
		Case 172
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 7200 Series" 
		Case 173
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 7300 Series" 
		Case 174
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 7400 Series" 
		Case 175
			familaCPU = "Multi-Core Intel(R) Xeon(R) processor 7400 Series" 
		Case 176
			familaCPU = "Pentium(R) III Xeon(TM)" 
		Case 177
			familaCPU = "Pentium(R) III Processor with Intel(R) SpeedStep(TM) Technology" 
		Case 178
			familaCPU = "Pentium(R) 4" 
		Case 179
			familaCPU = "Intel(R) Xeon(TM)" 
		Case 180
			familaCPU = "AS400 Family" 
		Case 181
			familaCPU = "Intel(R) Xeon(TM) processor MP" 
		Case 182
			familaCPU = "AMD Athlon(TM) XP Family" 
		Case 183
			familaCPU = "AMD Athlon(TM) MP Family" 
		Case 184
			familaCPU = "Intel(R) Itanium(R) 2" 
		Case 185
			familaCPU = "Intel(R) Pentium(R) M processor" 
		Case 186
			familaCPU = "Intel(R) Celeron(R) D processor" 
		Case 187
			familaCPU = "Intel(R) Pentium(R) D processor" 
		Case 188
			familaCPU = "Intel(R) Pentium(R) Processor Extreme Edition" 
		Case 189
			familaCPU = "Intel(R) Core(TM) Solo Processor" 
		Case 190
			familaCPU = "K7" 
		Case 191
			familaCPU = "Intel(R) Core(TM)2 Duo Processor" 
		Case 192
			familaCPU = "Intel(R) Core(TM)2 Solo processor" 
		Case 193
			familaCPU = "Intel(R) Core(TM)2 Extreme processor" 
		Case 194
			familaCPU = "Intel(R) Core(TM)2 Quad processor" 
		Case 195
			familaCPU = "Intel(R) Core(TM)2 Extreme mobile processor" 
		Case 196
			familaCPU = "Intel(R) Core(TM)2 Duo mobile processor" 
		Case 197
			familaCPU = "Intel(R) Core(TM)2 Solo mobile processor" 
		Case 198
			familaCPU = "Intel(R) Core(TM) i7 processor" 
		Case 199
			familaCPU = "Dual-Core Intel(R) Celeron(R) Processor" 
		Case 200
			familaCPU = "S/390 and zSeries Family" 
		Case 201
			familaCPU = "ESA/390 G4" 
		Case 202
			familaCPU = "ESA/390 G5" 
		Case 203
			familaCPU = "ESA/390 G6" 
		Case 204
			familaCPU = "z/Architectur base" 
		Case 205
			familaCPU = "Intel(R) Core(TM) i5 processor" 
		Case 206
			familaCPU = "Intel(R) Core(TM) i3 processor" 
		Case 207
			familaCPU = "Intel(R) Core(TM) i9 processor" 
		Case 210
			familaCPU = "VIA C7(TM)-M Processor Family" 
		Case 211
			familaCPU = "VIA C7(TM)-D Processor Family" 
		Case 212
			familaCPU = "VIA C7(TM) Processor Family" 
		Case 213
			familaCPU = "VIA Eden(TM) Processor Family" 
		Case 214
			familaCPU = "Multi-Core Intel(R) Xeon(R) processor" 
		Case 215
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 3xxx Series" 
		Case 216
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 3xxx Series" 
		Case 217
			familaCPU = "VIA Nano(TM) Processor Family" 
		Case 218
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 5xxx Series" 
		Case 219
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 5xxx Series" 
		Case 221
			familaCPU = "Dual-Core Intel(R) Xeon(R) processor 7xxx Series" 
		Case 222
			familaCPU = "Quad-Core Intel(R) Xeon(R) processor 7xxx Series" 
		Case 223
			familaCPU = "Multi-Core Intel(R) Xeon(R) processor 7xxx Series" 
		Case 224
			familaCPU = "Multi-Core Intel(R) Xeon(R) processor 3400 Series" 
		Case 230
			familaCPU = "Embedded AMD Opteron(TM) Quad-Core Processor Family" 
		Case 231
			familaCPU = "AMD Phenom(TM) Triple-Core Processor Family" 
		Case 232
			familaCPU = "AMD Turion(TM) Ultra Dual-Core Mobile Processor Family" 
		Case 233
			familaCPU = "AMD Turion(TM) Dual-Core Mobile Processor Family" 
		Case 234
			familaCPU = "AMD Athlon(TM) Dual-Core Processor Family" 
		Case 235
			familaCPU = "AMD Sempron(TM) SI Processor Family" 
		Case 236
			familaCPU = "AMD Phenom(TM) II Processor Family" 
		Case 237
			familaCPU = "AMD Athlon(TM) II Processor Family" 
		Case 238
			familaCPU = "Six-Core AMD Opteron(TM) Processor Family" 
		Case 239
			familaCPU = "AMD Sempron(TM) M Processor Family" 
		Case 250
			familaCPU = "i860" 
		Case 251
			familaCPU = "i960" 
		Case 254
			familaCPU = "Reserved (SMBIOS Extension)" 
		Case 255
			familaCPU = "Reserved (Un-initialized Flash Content - Lo)" 
		Case 260
			familaCPU = "SH-3" 
		Case 261
			familaCPU = "SH-4" 
		Case 280
			familaCPU = "ARM" 
		Case 281
			familaCPU = "StrongARM" 
		Case 300
			familaCPU = "6x86" 
		Case 301
			familaCPU = "MediaGX" 
		Case 302
			familaCPU = "MII" 
		Case 320
			familaCPU = "WinChip" 
		Case 350
			familaCPU = "DSP" 
		Case 500
			familaCPU = "Video Processor" 
		Case Else
			familaCPU = numero
	End Select
End Function
	
Function arquitecturaCPU(numero)
	Select Case numero
		Case 0 
			arquitecturaCPU = "x86"
		Case 1 
			arquitecturaCPU = "MIPS"
		Case 2
			arquitecturaCPU = "Alpha"
		Case 3
			arquitecturaCPU = "PowerPC"
		Case 5
			arquitecturaCPU = "ARM"
		Case 6
			arquitecturaCPU = "ia64"
		Case 9
			arquitecturaCPU = "x64"
		Case 12
			arquitecturaCPU = "ARM64"
		Case Else
			arquitecturaCPU = numero
	End Select
End Function

Function tipoCPU(numero)
	Select Case numero
		Case 1
			tipoCPU = "Other"
		Case 2 
			tipoCPU = "Unknown"
		Case 3
			tipoCPU = "Central Processor"
		Case 4
			tipoCPU = "Math Processor"
		Case 5
			tipoCPU = "DSP Processor"
		Case 6
			tipoCPU = "Video Processor"
		Case Else
			tipoCPU = numero
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

    line(vbNewLine + "Version (nombre):      " & os.Caption)
    line("Arquitectura:          " & os.OSArchitecture)
    line("Fecha instalacion:     " & dtmInstallDate)
    line("SKU:                   " & os.OperatingSystemSKU)
    line("Tipo producto:         " & tipoProducto(os.ProductType))
    line("Codigo del pais:       " & os.CountryCode)
    line("Lenguaje:              " & os.OsLanguage)
    line("Suite:                 " & suite(os.OSProductSuite))
    line("Tipo:                  " & tipoOs(os.OSType))
    line("Es el principal:       " & os.Primary)
    line("Numero serial:         " & os.SerialNumber)
    line("Version:               " & os.Version)
    line("Build:                 " & os.BuildNumber)
    line("Service pack (major):  " & os.ServicePackMajorVersion)
    line("Service pack (minor):  " & os.ServicePackMinorVersion)
    line("Tipo build:            " & os.BuildType)
    line("Estado del sistema:    " & os.Status)
    line("Memoria virtual:       " & convertirBytes(os.TotalVirtualMemorySize))
    line("Memoria visible:       " & convertirBytes(os.TotalVisibleMemorySize))
    line("Disp arranque:         " & os.BootDevice)
    line("Usuario registrado:    " & os.RegisteredUser)
Next

texto = obtenerComando("wmic path softwareLicensingService get OA3xOriginalProductKey")
texto = Split(texto)
licenciaWmic = texto(2)
line(vbNewLine + vbNewLine + "Licencia (wmic):     " + licenciaWmic)
llaveRegistro = ObtenerLlaveLicencia(oShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
line("Licencia (registro):   " & llaveRegistro)


line(vbNewLine + vbNewLine + lineaTitulo)
line("BIOS")
line(lineaTitulo + vbNewline)
Set bioss = objWMIService.ExecQuery ("Select * from Win32_BIOS  ")

For Each bios in bioss
	dtmConvertedDate.Value = bios.ReleaseDate
    dtmReleaseDate = dtmConvertedDate.GetVarDate

	line(vbNewLine + "Version BIOS:          " & Join(bios.BIOSVersion))
	line("Build number:          " & bios.BuildNumber)
	line("Leyenda:               " & bios.Caption)
	line("Descripcion:           " & bios.Description)
	line("Version Cont. (major): " & bios.EmbeddedControllerMajorVersion)
	line("Version Cont. (minor): " & bios.EmbeddedControllerMinorVersion)
	line("Codigo identificador:  " & bios.IdentificationCode)
	line("Manufacturador:        " & bios.Manufacturer)
	line("nombre:                " & bios.Name)
	line("Fecha liberacion:      " & dtmReleaseDate)
	line("Numero serial:         " & bios.SerialNumber)
	line("Version SMBIOS:        " & bios.SMBIOSBIOSVersion)
	line("Version SMBIOS(major): " & bios.SMBIOSMajorVersion)
	line("Version SMBIOS(minor): " & bios.SMBIOSMinorVersion)
	line("SMBIOS Presente:       " & bios.SMBIOSPresent)
	line("Ver. Sys.Bios (major): " & bios.SystemBiosMajorVersion)
	line("Ver. Sys.Bios (minor): " & bios.SystemBiosMinorVersion)
	line("Version:               " & bios.Version)
Next

'--------------------------
' INFORMACION DE HARDWARE  
'--------------------------


line(vbNewLine + vbNewLine + lineaTitulo)
line("COMPUTER SYSTEM")
line(lineaTitulo + vbNewline)
Set compus = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem ")

For Each compu in compus
	line(vbNewLine + "Estado de booteo:      " & compu.BootupState)
	line("Leyenda:               " & compu.Caption)
	line("SKU chasis:            " & compu.ChassisSKUNumber)
	line("Descripcion PC:        " & compu.Description)
	line("Nombre Servidor:       " & compu.DNSHostName)
	line("Dominio:               " & compu.Domain)
	line("Rol en dominio:        " & rolDominio(compu.DomainRole))
	line("Manufacturadora:       " & compu.Manufacturer)
	line("Modelo:                " & compu.Model)
	line("Nombre:                " & compu.Name)
	line("Servidor de red:       " & compu.NetworkServerModeEnabled)
	line("Procesadores logicos:  " & compu.NumberOfLogicalProcessors)
	line("Procesadores fisicos:  " & compu.NumberOfProcessors)
	line("Tipo sistema:          " & tipoPc(compu.PCSystemType))
	line("Tipo sistema ex:       " & tipoPcX(compu.PCSystemTypeEx))
	line("Nombre propietario:    " & compu.PrimaryOwnerName)
	line("Familia sistema:       " & compu.SystemFamily)
	line("SKU sistema:           " & compu.SystemSKUNumber)
	line("Tipo de sistema:       " & compu.SystemType)
	line("Memoria fisica (RAM):  " & convertirBytes(compu.TotalPhysicalMemory))
	line("Grupo de trabajo:      " & compu.Workgroup)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("PROCESSOR")
line(lineaTitulo + vbNewline)
Set procesadores = objWMIService.ExecQuery ("Select * from Win32_Processor ")

For Each procesador in procesadores
	line(vbNewLine + "Arquitectura:          " & arquitecturaCPU(procesador.Architecture))
	line("Disponibilidad:        " & disponibilidadCPU(procesador.Availability))
	line("Leyenda:               " & procesador.Caption)
	line("CurrentClockSpeed:     " & procesador.CurrentClockSpeed & " MHz")
	line("Descripcion:           " & procesador.Description)
	line("DeviceID:              " & procesador.DeviceID)
	line("Family:                " & procesador.Family)
	line("Cache L2:              " & convertirBytes(procesador.L2CacheSize))
	line("Cache L3:              " & convertirBytes(procesador.L3CacheSize))
	line("Nivel:                 " & procesador.Level)
	line("Manufacturador:        " & procesador.Manufacturer)
	line("MaxClockSpeed:         " & procesador.MaxClockSpeed & " MHz")
	line("Nombre (modelo):       " & procesador.Name)
	line("Nucleos:               " & procesador.NumberOfCores)
	line("Nucles habilitados:    " & procesador.NumberOfEnabledCore)
	line("Procesadores logicos:  " & procesador.NumberOfLogicalProcessors)
	line("Numero de parte:       " & procesador.PartNumber)
	line("Id procesador:         " & procesador.ProcessorId)
	line("Tipo procesador:       " & tipoCPU(procesador.ProcessorType))
	line("Revision:              " & procesador.Revision)
	line("Rol:                   " & procesador.Role)
	line("Numero de serie:       " & procesador.SerialNumber)
	line("Socket:                " & procesador.SocketDesignation)
	line("Stepping:              " & procesador.Stepping)
	line("Hilos:                 " & procesador.ThreadCount)
	line("UniqueId:              " & procesador.UniqueId)
	line("UpgradeMethod:         " & procesador.UpgradeMethod)
	line("Version:               " & procesador.Version)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("BASEBOARD")
line(lineaTitulo + vbNewline)
Set baseboards = objWMIService.ExecQuery ("Select * from Win32_BaseBoard")

For Each placaMadre in baseboards
	line(vbNewLine + "Leyenda:               " & placaMadre.Caption)
	line("Opciones de config.:   " & Join(placaMadre.ConfigOptions))
	line("Descripcion:           " & placaMadre.Description)
	line("Hot swappable:         " & placaMadre.HotSwappable)
	line("Manufacturador:        " & placaMadre.Manufacturer)
	line("Modelo:                " & placaMadre.Model)
	line("Nombre:                " & placaMadre.Name)
	line("Numero de parte:       " & placaMadre.PartNumber)
	line("Producto (modelo):     " & placaMadre.Product)
	line("Removible:             " & placaMadre.Removable)
	line("Reemplazable:          " & placaMadre.Replaceable)
	line("Require placa hija:    " & placaMadre.RequiresDaughterBoard)
	line("Numero de serie:       " & placaMadre.SerialNumber)
	line("SKU:                   " & placaMadre.SKU)
	line("Etiqueta:              " & placaMadre.Tag)
	line("Version:               " & placaMadre.Version)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("DISKDRIVES")
line(lineaTitulo + vbNewline)
Set discos = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")

For Each disco in discos
	' line("Leyenda:               " & placaMadre.Caption)
	line(vbNewLine + "Availability: " & disco.Availability)
	line("BytesPerSector: " & disco.BytesPerSector)
	line("CapabilityDescriptions: " & Join(disco.CapabilityDescriptions, ", "))
	line("Caption: " & disco.Caption)
	line("CompressionMethod: " & disco.CompressionMethod)
	line("ConfigManagerErrorCode: " & disco.ConfigManagerErrorCode)
	line("ConfigManagerUserConfig: " & disco.ConfigManagerUserConfig)
	line("CreationClassName: " & disco.CreationClassName)
	line("DefaultBlockSize: " & disco.DefaultBlockSize)
	line("Description: " & disco.Description)
	line("DeviceID: " & disco.DeviceID)
	line("ErrorCleared: " & disco.ErrorCleared)
	line("ErrorDescription: " & disco.ErrorDescription)
	line("ErrorMethodology: " & disco.ErrorMethodology)
	line("FirmwareRevision: " & disco.FirmwareRevision)
	line("Index: " & disco.Index)
	line("InstallDate: " & disco.InstallDate)
	line("InterfaceType: " & disco.InterfaceType)
	line("LastErrorCode: " & disco.LastErrorCode)
	line("Manufacturer: " & disco.Manufacturer)
	line("MaxBlockSize: " & disco.MaxBlockSize)
	line("MaxMediaSize: " & disco.MaxMediaSize)
	line("MediaLoaded: " & disco.MediaLoaded)
	line("MediaType: " & disco.MediaType)
	line("MinBlockSize: " & disco.MinBlockSize)
	line("Model: " & disco.Model)
	line("Name: " & disco.Name)
	line("NeedsCleaning: " & disco.NeedsCleaning)
	line("NumberOfMediaSupported: " & disco.NumberOfMediaSupported)
	line("Partitions: " & disco.Partitions)
	line("PNPDeviceID: " & disco.PNPDeviceID)
	line("PowerManagementSupported: " & disco.PowerManagementSupported)
	line("SCSIBus: " & disco.SCSIBus)
	line("SCSILogicalUnit: " & disco.SCSILogicalUnit)
	line("SCSIPort: " & disco.SCSIPort)
	line("SCSITargetId: " & disco.SCSITargetId)
	line("SectorsPerTrack: " & disco.SectorsPerTrack)
	line("SerialNumber: " & disco.SerialNumber)
	line("Signature: " & disco.Signature)
	line("Size: " & disco.Size)
	line("Status: " & disco.Status)
	line("StatusInfo: " & disco.StatusInfo)
	line("SystemCreationClassName: " & disco.SystemCreationClassName)
	line("SystemName: " & disco.SystemName)
	line("TotalCylinders: " & disco.TotalCylinders)
	line("TotalHeads: " & disco.TotalHeads)
	line("TotalSectors: " & disco.TotalSectors)
	line("TotalTracks: " & disco.TotalTracks)
	line("TracksPerCylinder: " & disco.TracksPerCylinder)
Next


'--------
' FINAL
'--------

line(vbNewLine + vbNewLine +"----------------------------- final del archivo ----------------------------")
tituloFinal = "Ejecucion terminada"
mensajeFinal = "Datos de " + computerName + " obtenidos" + vbNewLine + "Ver: " + dirActual + nombreArchivo
MsgBox mensajeFinal,, tituloFinal


' ------------------------------ principal (fin) ----------------------------- '