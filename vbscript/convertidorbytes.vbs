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

WScript.Echo convertirBytes(3221225472)         ' 3 GB
WScript.Echo convertirBytes(15247133900.8)      ' 14.2 GB
WScript.Echo convertirBytes(268435456)          ' 256 MB
WScript.Echo convertirBytes(281474976710656)    ' 256 TB
WScript.Echo convertirBytes(5629499534213120)   ' 5 PT