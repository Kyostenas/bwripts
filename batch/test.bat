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


(
    echo %separador%%nl%
    echo Fecha de consulta: %DATE% %TIME%
    echo Computadora consultada: %ComputerName%%nl%
    echo %separador%%nl%
    
    echo %linea_titulo%PRUEBA%nl%%linea_titulo%

) > .\INFO_%ComputerName%



ENDLOCAL
EXIT /B %ERRORLEVEL%