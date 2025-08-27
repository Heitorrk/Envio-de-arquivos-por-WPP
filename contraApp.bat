@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

REM tenta achar o jar (ajuste o nome se necessário)
set "JAR=contraApp.jar"
if not exist "%JAR%" (
  if exist "target\contraApp-jar-with-dependencies.jar" set "JAR=target\contraApp-jar-with-dependencies.jar"
  if exist "out\artifacts\contraApp_jar\contraApp.jar" set "JAR=out\artifacts\contraApp_jar\contraApp.jar"
)

if not exist "%JAR%" (
  echo [ERRO] Nao encontrei o JAR. Deixe-o ao lado deste .bat ou em target.
  pause
  exit /b 1
)

REM silenciar aviso do log4j2 (opcional)
set "JAVA_OPTS=-Dfile.encoding=UTF-8 -Dlog4j2.status=OFF"

echo [INFO] Iniciando...
java %JAVA_OPTS% -jar "%JAR%"
echo.
pause