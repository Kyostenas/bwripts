:: VER https://www.tutorialspoint.com/batch_script/batch_script_commands.htm
:: VER https://ss64.com/vb/shell.html

:: Esto es para evitar que se vea la ejecucion 
:: de la script en la consola
@echo off

SETLOCAL

:: :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::                                  VARIABLES                                   
:: :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

set nl=^&echo.
set separador=::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
set separador_mini=+--------------------------------------------------------+
set linea_titulo=+-----+^&echo.


:: ----------------------------- variables (fin) -------------------------------

Set WshShell = CreateObject("Wscript.Shell")
WshShell.run "wmic path softwareLicensingService get OA3xOriginalProductKey"



(
    echo %separador%%nl%
    echo Fecha de consulta: %DATE% %TIME%
    echo Computadora consultada: %ComputerName%%nl%
    echo %separador%%nl%
    
    echo %linea_titulo%LICENCIA WINDOWS%nl%%linea_titulo%
    echo %licencia_wmic%

) > .\INFO_%ComputerName%


ENDLOCAL
EXIT /B %ERRORLEVEL%