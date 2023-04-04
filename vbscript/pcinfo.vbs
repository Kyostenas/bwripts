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
	On Error Resume Next
    Dim FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    directorioActual = FSO.GetAbsolutePathName(".")
End Function

fechaHoy = CStr(Date + Time)
separador=vbNewLine + "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" + vbNewLine
lineaTitulo   ="--------------------------------"
lineaSubTitulo="'''''''''''''''''''''''"

Function ObtenerLlaveLicencia(Key)
	On Error Resume Next
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
	On Error Resume Next
    Set ObjExec = oShell.Exec("cmd.exe /c " & ejecucion)
    obtenerComando = ObjExec.stdOut.ReadAll
End Function

Function convertirCapacidaDeBytes(bytes)
	On Error Resume Next
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
	convertirCapacidaDeBytes = Round(paso, 2) & sufijos(indice)
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


'---------------------------------------------------------------------------------------+
'     ------------------------- tabla de caracteres ascii ------------------------      |
'                                                                                       |
' ------------------------------------------------------------------------------------- |
' Char  Dec  Oct  Hex | Char  Dec  Oct  Hex | Char  Dec  Oct  Hex | Char  Dec  Oct  Hex |
' ------------------------------------------------------------------------------------- |
' (nul)   0 0000 0x00 |        32 0040 0x20 | @      64 0100 0x40 | `      96 0140 0x60 |
' (soh)   1 0001 0x01 | !      33 0041 0x21 | A      65 0101 0x41 | a      97 0141 0x61 |
' (stx)   2 0002 0x02 | "      34 0042 0x22 | B      66 0102 0x42 | b      98 0142 0x62 |
' (etx)   3 0003 0x03 | #      35 0043 0x23 | C      67 0103 0x43 | c      99 0143 0x63 |
' (eot)   4 0004 0x04 | $      36 0044 0x24 | D      68 0104 0x44 | d     100 0144 0x64 |
' (enq)   5 0005 0x05 | %      37 0045 0x25 | E      69 0105 0x45 | e     101 0145 0x65 |
' (ack)   6 0006 0x06 | &      38 0046 0x26 | F      70 0106 0x46 | f     102 0146 0x66 |
' (bel)   7 0007 0x07 | '      39 0047 0x27 | G      71 0107 0x47 | g     103 0147 0x67 |
' (bs)    8 0010 0x08 | (      40 0050 0x28 | H      72 0110 0x48 | h     104 0150 0x68 |
' (ht)    9 0011 0x09 | )      41 0051 0x29 | I      73 0111 0x49 | i     105 0151 0x69 |
' (nl)   10 0012 0x0a | *      42 0052 0x2a | J      74 0112 0x4a | j     106 0152 0x6a |
' (vt)   11 0013 0x0b | +      43 0053 0x2b | K      75 0113 0x4b | k     107 0153 0x6b |
' (np)   12 0014 0x0c | ,      44 0054 0x2c | L      76 0114 0x4c | l     108 0154 0x6c |
' (cr)   13 0015 0x0d | -      45 0055 0x2d | M      77 0115 0x4d | m     109 0155 0x6d |
' (so)   14 0016 0x0e | .      46 0056 0x2e | N      78 0116 0x4e | n     110 0156 0x6e |
' (si)   15 0017 0x0f | /      47 0057 0x2f | O      79 0117 0x4f | o     111 0157 0x6f |
' (dle)  16 0020 0x10 | 0      48 0060 0x30 | P      80 0120 0x50 | p     112 0160 0x70 |
' (dc1)  17 0021 0x11 | 1      49 0061 0x31 | Q      81 0121 0x51 | q     113 0161 0x71 |
' (dc2)  18 0022 0x12 | 2      50 0062 0x32 | R      82 0122 0x52 | r     114 0162 0x72 |
' (dc3)  19 0023 0x13 | 3      51 0063 0x33 | S      83 0123 0x53 | s     115 0163 0x73 |
' (dc4)  20 0024 0x14 | 4      52 0064 0x34 | T      84 0124 0x54 | t     116 0164 0x74 |
' (nak)  21 0025 0x15 | 5      53 0065 0x35 | U      85 0125 0x55 | u     117 0165 0x75 |
' (syn)  22 0026 0x16 | 6      54 0066 0x36 | V      86 0126 0x56 | v     118 0166 0x76 |
' (etb)  23 0027 0x17 | 7      55 0067 0x37 | W      87 0127 0x57 | w     119 0167 0x77 |
' (can)  24 0030 0x18 | 8      56 0070 0x38 | X      88 0130 0x58 | x     120 0170 0x78 |
' (em)   25 0031 0x19 | 9      57 0071 0x39 | Y      89 0131 0x59 | y     121 0171 0x79 |
' (sub)  26 0032 0x1a | :      58 0072 0x3a | Z      90 0132 0x5a | z     122 0172 0x7a |
' (esc)  27 0033 0x1b | ;      59 0073 0x3b | [      91 0133 0x5b | {     123 0173 0x7b |
' (fs)   28 0034 0x1c | <      60 0074 0x3c | \      92 0134 0x5c | |     124 0174 0x7c |
' (gs)   29 0035 0x1d | =      61 0075 0x3d | ]      93 0135 0x5d | }     125 0175 0x7d |
' (rs)   30 0036 0x1e | >      62 0076 0x3e | ^      94 0136 0x5e | ~     126 0176 0x7e |
' (us)   31 0037 0x1f | ?      63 0077 0x3f | _      95 0137 0x5f | (del) 127 0177 0x7f |
'---------------------------------------------------------------------------------------+


Function decimalA_ASCII(numero)
	Select Case numero
		Case 32 decimalA_ASCII  = " "
		Case 33 decimalA_ASCII  = "!"
		Case 34 decimalA_ASCII  = "''"
		Case 35 decimalA_ASCII  = "#"
		Case 36 decimalA_ASCII  = "$"
		Case 37 decimalA_ASCII  = "%"
		Case 38 decimalA_ASCII  = "&"
		Case 39 decimalA_ASCII  = "'"
		Case 40 decimalA_ASCII  = "("
		Case 41 decimalA_ASCII  = ")"
		Case 42 decimalA_ASCII  = "*"
		Case 43 decimalA_ASCII  = "+"
		Case 44 decimalA_ASCII  = ","
		Case 45 decimalA_ASCII  = "-"
		Case 46 decimalA_ASCII  = "."
		Case 47 decimalA_ASCII  = "/"
		Case 48 decimalA_ASCII  = "0"
		Case 49 decimalA_ASCII  = "1"
		Case 50 decimalA_ASCII  = "2"
		Case 51 decimalA_ASCII  = "3"
		Case 52 decimalA_ASCII  = "4"
		Case 53 decimalA_ASCII  = "5"
		Case 54 decimalA_ASCII  = "6"
		Case 55 decimalA_ASCII  = "7"
		Case 56 decimalA_ASCII  = "8"
		Case 57 decimalA_ASCII  = "9"
		Case 58 decimalA_ASCII  = ":"
		Case 59 decimalA_ASCII  = ";"
		Case 60 decimalA_ASCII  = "<"
		Case 61 decimalA_ASCII  = "="
		Case 62 decimalA_ASCII  = ">"
		Case 63 decimalA_ASCII  = "?"
		Case 64 decimalA_ASCII  = "@"
		Case 65 decimalA_ASCII  = "A"
		Case 66 decimalA_ASCII  = "B"
		Case 67 decimalA_ASCII  = "C"
		Case 68 decimalA_ASCII  = "D"
		Case 69 decimalA_ASCII  = "E"
		Case 70 decimalA_ASCII  = "F"
		Case 71 decimalA_ASCII  = "G"
		Case 72 decimalA_ASCII  = "H"
		Case 73 decimalA_ASCII  = "I"
		Case 74 decimalA_ASCII  = "J"
		Case 75 decimalA_ASCII  = "K"
		Case 76 decimalA_ASCII  = "L"
		Case 77 decimalA_ASCII  = "M"
		Case 78 decimalA_ASCII  = "N"
		Case 79 decimalA_ASCII  = "O"
		Case 80 decimalA_ASCII  = "P"
		Case 81 decimalA_ASCII  = "Q"
		Case 82 decimalA_ASCII  = "R"
		Case 83 decimalA_ASCII  = "S"
		Case 84 decimalA_ASCII  = "T"
		Case 85 decimalA_ASCII  = "U"
		Case 86 decimalA_ASCII  = "V"
		Case 87 decimalA_ASCII  = "W"
		Case 88 decimalA_ASCII  = "X"
		Case 89 decimalA_ASCII  = "Y"
		Case 90 decimalA_ASCII  = "Z"
		Case 91 decimalA_ASCII  = "["
		Case 92 decimalA_ASCII  = "\"
		Case 93 decimalA_ASCII  = "]"
		Case 94 decimalA_ASCII  = "^"
		Case 95 decimalA_ASCII  = "_"
		Case 96 decimalA_ASCII  = "`"
		Case 97 decimalA_ASCII  = "a"
		Case 98 decimalA_ASCII  = "b"
		Case 99 decimalA_ASCII  = "c"
		Case 100 decimalA_ASCII = "d"
		Case 101 decimalA_ASCII = "e"
		Case 102 decimalA_ASCII = "f"
		Case 103 decimalA_ASCII = "g"
		Case 104 decimalA_ASCII = "h"
		Case 105 decimalA_ASCII = "i"
		Case 106 decimalA_ASCII = "j"
		Case 107 decimalA_ASCII = "k"
		Case 108 decimalA_ASCII = "l"
		Case 109 decimalA_ASCII = "m"
		Case 110 decimalA_ASCII = "n"
		Case 111 decimalA_ASCII = "o"
		Case 112 decimalA_ASCII = "p"
		Case 113 decimalA_ASCII = "q"
		Case 114 decimalA_ASCII = "r"
		Case 115 decimalA_ASCII = "s"
		Case 116 decimalA_ASCII = "t"
		Case 117 decimalA_ASCII = "u"
		Case 118 decimalA_ASCII = "v"
		Case 119 decimalA_ASCII = "w"
		Case 120 decimalA_ASCII = "x"
		Case 121 decimalA_ASCII = "y"
		Case 122 decimalA_ASCII = "z"
		Case 123 decimalA_ASCII = "{"
		Case 124 decimalA_ASCII = "|"
		Case 125 decimalA_ASCII = "}"
		Case 126 decimalA_ASCII = "~"
		Case Else decimalA_ASCII = ""
	End Select
End Function
	
Function covnertirArregloDecimalA_ASCII(arreglo)
	On Error Resume Next
	caracteresListos = ""
	For Each charDecimal in arreglo
		caracteresListos = caracteresListos & decimalA_ASCII(charDecimal)
	Next
	covnertirArregloDecimalA_ASCII = caracteresListos
End Function

Function paramsConeccionPantalla(numero)
	Select Case numero
		Case -2 paramsConeccionPantalla = "Nada"
		Case -1 paramsConeccionPantalla = "Otro"
		Case 0 paramsConeccionPantalla = "HD15 (VGA)"
		Case 1 paramsConeccionPantalla = "S-video"
		Case 2 paramsConeccionPantalla = "Composite video"
		Case 3 paramsConeccionPantalla = "Component video"
		Case 4 paramsConeccionPantalla = "DVI"
		Case 5 paramsConeccionPantalla = "HDMI"
		Case 6 paramsConeccionPantalla = "Low Voltage Differential Swing (LVDS)/ Mobile Industry Processor Interface (MIPI) Digital Serial Interface (DSI)"
		Case 8 paramsConeccionPantalla = "D-Jpn"
		Case 9 paramsConeccionPantalla = "SDI"
		Case 10 paramsConeccionPantalla = "External Display Port"
		Case 11 paramsConeccionPantalla = "Embedded Display Port"
		Case 12 paramsConeccionPantalla = "External Unified Display Interface (UDI)"
		Case 13 paramsConeccionPantalla = "Embedded Unified Display Interface (UDI)"
		Case 14 paramsConeccionPantalla = "Dongle cable that supports SDTV"
		Case 15 paramsConeccionPantalla = "Wireless Miracast session"
		Case 16 paramsConeccionPantalla = "Wired indirect display device"
		Case 0x80000000 paramsConeccionPantalla = "Internal connection (Possibly a laptop)"
		Case Else paramsConeccionPantalla = "?"
	End Select
End Function

Function intentarJoin(arreglo)
	On Error Resume Next
	intentarJoin = Join(arreglo)
End Function


Function inputUsuario(prompt, defecto)
'-------------------------------------------------------------------------+
' This function prompts the user for some input.                          |
' When the script runs in CSCRIPT.EXE, StdIn is used,                     |
' otherwise the VBScript InputBox( ) function is used.                    |
' myPrompt is the the text used to prompt the user for input.             |
' The function returns the input typed either on StdIn or in InputBox( ). |
' Written by Rob van der Woude                                            |
' http://www.robvanderwoude.com                                           |
'-------------------------------------------------------------------------+
	' Check if the script runs in CSCRIPT.EXE
	If UCase( Right( WScript.FullName, 12) ) = "\CSCRIPT.EXE" Then
		' If so, use StdIn and StdOut
		WScript.StdOut.Write prompt & " "
		inputUsuario = WScript.StdIn.ReadLine
	Else
		' If not, use InputBox()
		inputUsuario = InputBox( prompt, "Inventario de computadoras", defecto)
	End If
End Function


Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer  = "."
Set objCIMV2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
		 				& strComputer & "\root\cimv2")
Set objWMI   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
						& strComputer & "\root\wmi")


' :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: '
'                                   PRINCIPAL                                  '
' :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: '

'-------------
' ENCABEZADO
'-------------

line(separador)
line("Usuario:                 " & inputUsuario("Nombre del usuario", "Usuario"))
line("Departamento:            " & inputUsuario("Departamento donde trabaja", "Departamento"))
line("Computadora de consulta: " + computerName)
line("Fecha de consulta:       " + fechaHoy)
line(separador)

'--------------------
' SISTEMA OPERATIVO
'--------------------

line(lineaTitulo)
line("SISTEMAS OPERATIVOS")
line(lineaTitulo + vbNewline)
Set oss = objCIMV2.ExecQuery ("Select * from Win32_OperatingSystem")

For Each os in oss
    dtmConvertedDate.Value = os.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate

    line(vbNewLine + "Version (nombre):      " & os.Caption)
    line("Arquitectura:          " & os.OSArchitecture)
    line("Fecha instalacion:     " & dtmInstallDate)
    ' line("SKU:                   " & os.OperatingSystemSKU)
    line("Tipo producto:         " & tipoProducto(os.ProductType))
    ' line("Codigo del pais:       " & os.CountryCode)
    ' line("Lenguaje:              " & os.OsLanguage)
    ' line("Suite:                 " & suite(os.OSProductSuite))
    ' line("Tipo:                  " & tipoOs(os.OSType))
    line("Es el principal:       " & os.Primary)
    ' line("Numero serial:         " & os.SerialNumber)
    line("Version:               " & os.Version)
    line("Build:                 " & os.BuildNumber)
    line("Service pack (major):  " & os.ServicePackMajorVersion)
    line("Service pack (minor):  " & os.ServicePackMinorVersion)
    ' line("Tipo build:            " & os.BuildType)
    ' line("Estado del sistema:    " & os.Status)
    ' line("Memoria virtual:       " & convertirCapacidaDeBytes(os.TotalVirtualMemorySize))
    ' line("Memoria visible:       " & convertirCapacidaDeBytes(os.TotalVisibleMemorySize))
    line("Disp arranque:         " & os.BootDevice)
    ' line("Usuario registrado:    " & os.RegisteredUser)
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
Set bioss = objCIMV2.ExecQuery ("Select * from Win32_BIOS  ")

For Each bios in bioss
	dtmConvertedDate.Value = bios.ReleaseDate
    dtmReleaseDate = dtmConvertedDate.GetVarDate

	line(vbNewLine + "Version BIOS:          " & Join(bios.BIOSVersion))
	line("Build number:          " & bios.BuildNumber)
	line("Leyenda:               " & bios.Caption)
	line("Descripcion:           " & bios.Description)
	line("Version Cont. (major): " & bios.EmbeddedControllerMajorVersion)
	line("Version Cont. (minor): " & bios.EmbeddedControllerMinorVersion)
	' line("Codigo identificador:  " & bios.IdentificationCode)
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
Set compus = objCIMV2.ExecQuery ("Select * from Win32_ComputerSystem ")

For Each compu in compus
	line(vbNewLine + "Estado de booteo:      " & compu.BootupState)
	line("Leyenda:               " & compu.Caption)
	' line("SKU chasis:            " & compu.ChassisSKUNumber)
	line("Descripcion PC:        " & compu.Description)
	line("Nombre Servidor:       " & compu.DNSHostName)
	line("Dominio:               " & compu.Domain)
	line("Rol en dominio:        " & rolDominio(compu.DomainRole))
	line("Manufacturadora:       " & compu.Manufacturer)
	line("Modelo:                " & compu.Model)
	line("Nombre:                " & compu.Name)
	' line("Servidor de red:       " & compu.NetworkServerModeEnabled)
	' line("Procesadores logicos:  " & compu.NumberOfLogicalProcessors)
	' line("Procesadores fisicos:  " & compu.NumberOfProcessors)
	line("Tipo sistema:          " & tipoPc(compu.PCSystemType))
	line("Tipo sistema ex:       " & tipoPcX(compu.PCSystemTypeEx))
	line("Nombre propietario:    " & compu.PrimaryOwnerName)
	' line("Familia sistema:       " & compu.SystemFamily)
	' line("SKU sistema:           " & compu.SystemSKUNumber)
	' line("Tipo de sistema:       " & compu.SystemType)
	line("Memoria fisica (RAM):  " & convertirCapacidaDeBytes(compu.TotalPhysicalMemory))
	' line("Grupo de trabajo:      " & compu.Workgroup)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("COMPUTER SYSTEM PRODUCT")
line(lineaTitulo + vbNewline)
Set csproducts = objCIMV2.ExecQuery ("Select * from Win32_ComputerSystemProduct")

For Each csproduct in csproducts
	line(vbNewLine + "Leyenda:               " & csproduct.Caption)
	line("Descripcion:           " & csproduct.Description)
	line("Numero identificador:  " & csproduct.IdentifyingNumber)
	line("Nombre:                " & csproduct.Name)
	line("Numero SKU:            " & csproduct.SKUNumber)
	line("Vendedor:              " & csproduct.Vendor)
	line("Version:               " & csproduct.Version)
	line("UUID:                  " & csproduct.UUID)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("BASEBOARD")
line(lineaTitulo + vbNewline)
Set baseboards = objCIMV2.ExecQuery ("Select * from Win32_BaseBoard")

For Each placaMadre in baseboards
	line(vbNewLine + "Leyenda:               " & placaMadre.Caption)
	' line("Opciones de config.:   " & Join(placaMadre.ConfigOptions))
	line("Descripcion:           " & placaMadre.Description)
	line("Hot swappable:         " & placaMadre.HotSwappable)
	line("Manufacturador:        " & placaMadre.Manufacturer)
	line("Modelo:                " & placaMadre.Model)
	line("Nombre:                " & placaMadre.Name)
	line("Numero de parte:       " & placaMadre.PartNumber)
	line("Producto (modelo):     " & placaMadre.Product)
	' line("Removible:             " & placaMadre.Removable)
	line("Reemplazable:          " & placaMadre.Replaceable)
	line("Require placa hija:    " & placaMadre.RequiresDaughterBoard)
	line("Numero de serie:       " & placaMadre.SerialNumber)
	' line("SKU:                   " & placaMadre.SKU)
	' line("Etiqueta:              " & placaMadre.Tag)
	line("Version:               " & placaMadre.Version)
Next


line(vbNewLine + vbNewLine + lineaTitulo)
line("PROCESSOR")
line(lineaTitulo + vbNewline)
Set procesadores = objCIMV2.ExecQuery ("Select * from Win32_Processor ")

For Each procesador in procesadores
	line(vbNewLine + "Arquitectura:          " & arquitecturaCPU(procesador.Architecture))
	' line("Disponibilidad:        " & disponibilidadCPU(procesador.Availability))
	' line("Leyenda:               " & procesador.Caption)
	line("CurrentClockSpeed:     " & procesador.CurrentClockSpeed & " MHz")
	line("MaxClockSpeed:         " & procesador.MaxClockSpeed & " MHz")
	' line("Descripcion:           " & procesador.Description)
	' line("DeviceID:              " & procesador.DeviceID)
	' line("Family:                " & procesador.Family)
	line("Cache L2:              " & convertirCapacidaDeBytes(procesador.L2CacheSize))
	line("Cache L3:              " & convertirCapacidaDeBytes(procesador.L3CacheSize))
	' line("Nivel:                 " & procesador.Level)
	line("Manufacturador:        " & procesador.Manufacturer)
	line("Nombre (modelo):       " & procesador.Name)
	line("Nucleos:               " & procesador.NumberOfCores)
	line("Nucles habilitados:    " & procesador.NumberOfEnabledCore)
	line("Procesadores logicos:  " & procesador.NumberOfLogicalProcessors)
	' line("Numero de parte:       " & procesador.PartNumber)
	' line("Id procesador:         " & procesador.ProcessorId)
	line("Tipo procesador:       " & tipoCPU(procesador.ProcessorType))
	' line("Revision:              " & procesador.Revision)
	' line("Rol:                   " & procesador.Role)
	line("Numero de serie:       " & procesador.SerialNumber)
	line("Socket:                " & procesador.SocketDesignation)
	' line("Stepping:              " & procesador.Stepping)
	line("Hilos:                 " & procesador.ThreadCount)
	' line("UniqueId:              " & procesador.UniqueId)
	' line("UpgradeMethod:         " & procesador.UpgradeMethod)
	' line("Version:               " & procesador.Version)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("DISKDRIVES")
line(lineaTitulo + vbNewline)
Set discos = objCIMV2.ExecQuery ("Select * from Win32_DiskDrive")

For Each disco in discos
	line(vbNewLine & "Capacidades:           " & Join(disco.CapabilityDescriptions, ", "))
	line("Leyenda:               " & disco.Caption)
	' line("Descripcion:           " & disco.Description)
	line("ID dispositivo:        " & disco.DeviceID)
	line("Revision Firmware:     " & disco.FirmwareRevision)
	line("Index:                 " & disco.Index)
	line("Tipo interfaz:         " & disco.InterfaceType)
	' line("Manufacturador:        " & disco.Manufacturer)
	' line("Tipo media:            " & disco.MediaType)
	line("Modelo:                " & disco.Model)
	' line("Nombre:                " & disco.Name)
	' line("Particiones:           " & disco.Partitions)
	' line("ID dispositivo PNP:    " & disco.PNPDeviceID)
	line("Numero serial:         " & disco.SerialNumber)
	line("Capacidad:             " & convertirCapacidaDeBytes(disco.Size))
	line("Estatus:               " & disco.Status)
	line("Info estatus:          " & disco.StatusInfo)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("VIDEO CONTROLLERS (GPU's - GRAFICAS)")
line(lineaTitulo + vbNewline)
Set controladores = objCIMV2.ExecQuery ("Select * from Win32_VideoController")

For Each controlador in controladores
	line("Leyenda:               " & controlador.Caption)
	line("Descripcion:           " & controlador.Description)
	line("Nombre:                " & controlador.Name)
	line("Procesador video:      " & controlador.VideoProcessor)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("PHYSICAL MEMORY (RAM)")
line(lineaTitulo + vbNewline)
Set memorias = objCIMV2.ExecQuery ("Select * from Win32_PhysicalMemory ")

For Each memoria in memorias
	line(vbNewLine + "Etiqueta Slot:         " & memoria.BankLabel)
	line("Capacidad:             " & convertirCapacidaDeBytes(memoria.Capacity))
	line("Leyenda:               " & memoria.Caption)
	line("Velocidad configurada: " & memoria.ConfiguredClockSpeed & " MHz")
	line("Velocidad:             " & memoria.Speed & " MHz")
	line("Descripcion:           " & memoria.Description)
	line("Manufacturador:        " & memoria.Manufacturer)
	line("Modelo:                " & memoria.Model)
	line("Nombre:                " & memoria.Name)
	line("Numero de parte:       " & memoria.PartNumber)
	line("Numero serial:         " & memoria.SerialNumber)
	line("Tag:                   " & memoria.Tag)
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("DESKTOP MONITORS")
line(lineaTitulo + vbNewline)
Set monitores = objWMI.ExecQuery("Select * from WMIMonitorID")

For Each monitor in monitores
	line(vbNewLine + "Activo:                " & monitor.Active)
	line("Nombre instancia:      " & monitor.InstanceName)
	line("Manufacturador:        " & covnertirArregloDecimalA_ASCII(monitor.ManufacturerName))
	line("Codigo de producto:    " & covnertirArregloDecimalA_ASCII(monitor.ProductCodeID))
	line("Numero de serie:       " & covnertirArregloDecimalA_ASCII(monitor.SerialNumberID))
	line("Semana de manufactura: " & monitor.WeekOfManufacture)
	line("Year de manufactora:   " & monitor.YearOfManufacture)
	line("Nombre amigable:       " & covnertirArregloDecimalA_ASCII(monitor.UserFriendlyName))
Next

line(vbNewLine + "----")
Set conjuntosParamsMonitores = objWMI.ExecQuery("Select * from WmiMonitorConnectionParams")

For Each parametrosMonitor in conjuntosParamsMonitores
	line(vbNewLine + "Nombre instancia:      " & parametrosMonitor.InstanceName)
	line("Tecnologia salida:     " & paramsConeccionPantalla(parametrosMonitor.VideoOutputTechnology))
Next

line(vbNewLine + vbNewLine + lineaTitulo)
line("NETWORK")
line(lineaTitulo + vbNewline)
Set adaptadoresRed = objCIMV2.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=True")

For Each adaptadorRed in adaptadoresRed
	line(vbNewLine + "Leyenda:               " & adaptadorRed.Caption)
	line("Descripcion:           " & adaptadorRed.Description)
	line("Default IP Gateway:    " & intentarJoin(adaptadorRed.DefaultIPGateway))
	line("DHCP activado:         " & adaptadorRed.DHCPEnabled)
	line("Servidor DHCP:         " & adaptadorRed.DHCPServer)
	line("Direccion IP:          " & intentarJoin(adaptadorRed.IPAddress))
	line("Direccion MAC:         " & adaptadorRed.MACAddress)
Next


'--------
' FINAL
'--------

line(vbNewLine + vbNewLine +"----------------------------- final del archivo ----------------------------")
tituloFinal = "Ejecucion terminada"

mensajeFinal = "Datos de " + computerName + " obtenidos" + vbNewLine + "Ver: " + dirActual + nombreArchivo
MsgBox mensajeFinal,, tituloFinal


' ------------------------------ principal (fin) ----------------------------- '








