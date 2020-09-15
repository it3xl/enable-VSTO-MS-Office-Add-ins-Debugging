@SETLOCAL

@REM - Hi a developer or an admin. I am it3xl.
@REM - Your PC may be locked to debug VSTO MS Office Add-ins by the following Win-registry key.
@REM - \.NETFramework\Security\TrustManager\PromptingLevel
@REM -
@REM - By running this script with admin privileges you'll remove this limitation.
@REM - For details see
@REM - https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-configure-inclusion-list-security?view=vs-2019#enable-the-inclusion-list
@REM - https://msdn.microsoft.com/en-us/library/ms996418.aspx#clickoncetrustpub_topic2
@REM -
@REM - See the last version of this file in https://github.com/it3xl/enable-VSTO-MS-Office-Add-ins-Debugging

@ECHO:

@SET "reg_key=HKLM\SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"

@CALL REM Suppresses previous errors.
@REG QUERY %reg_key% > NUL 2>&1

@IF %ERRORLEVEL% NEQ 0 (

    @ECHO:&ECHO There is no the following Win-registry key
    @ECHO:&ECHO %reg_key%
    @ECHO:&ECHO This script is useless for you.
    @ECHO Search your troubles in another place.

    @ECHO:&ECHO ...delaying for some seconds before the completion.
    @TypePerf "\System\Processor Queue Length" -si 10 -sc 1 >nul

    @EXIT /B
)



@SET pause_for_sec=5

@REM The below code block detects if the script is being running with admin PRIVILEGES.
@NET SESSION >NUL 2>&1
@IF %ERRORLEVEL% NEQ 0 (

    @ECHO:
    @ECHO ######## ########  ########   #######  ########  
    @ECHO ##       ##     ## ##     ## ##     ## ##     ## 
    @ECHO ##       ##     ## ##     ## ##     ## ##     ## 
    @ECHO ######   ########  ########  ##     ## ########  
    @ECHO ##       ##   ##   ##   ##   ##     ## ##   ##   
    @ECHO ##       ##    ##  ##    ##  ##     ## ##    ##  
    @ECHO ######## ##     ## ##     ##  #######  ##     ## 
    @ECHO:
    @ECHO:
    @ECHO ####### ERROR: ADMINISTRATOR PRIVILEGES REQUIRED #########
    @ECHO This script must be run as administrator to work properly!  
    @ECHO:


 
    @ECHO ...delaying for %pause_for_sec% seconds before EXIT.
    @TypePerf "\System\Processor Queue Length" -si %pause_for_sec% -sc 1 >nul


    @EXIT 1
)

@REG ADD %reg_key% /v MyComputer /d "Enabled" /t REG_SZ /f  /reg:64
@IF %ERRORLEVEL% NEQ 0 (
    @ECHO There was some error for the MyComputer value. Analyse and fix.
)
@REG ADD %reg_key% /v LocalIntranet /d "Enabled" /t REG_SZ /f  /reg:64
@IF %ERRORLEVEL% NEQ 0 (
    @ECHO There was some error for the LocalIntranet value. Analyse and fix.
)

@ECHO:&ECHO ...delaying for some seconds before the completion.
@TypePerf "\System\Processor Queue Length" -si 10 -sc 1 >nul
