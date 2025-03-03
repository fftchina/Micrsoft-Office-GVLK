零售版的Office 轉換為VOL版腳本激活方法：

當您安装完成 Office Professional Plus  Retail 官方原版後，將下列的代碼用記事本貼上，再更改一下"kms.03k.org"個一欄為KMS Server目標IP後將內容後保存為 .bat 文件，然後以管理員的權限運行，即可一步完成所有的激活操作

<p>Microsoft Office LTSC Professional Plus 2024 ：<br/></p>
<pre><code>@echo off
title Activate Microsoft Office LTSC Professional Plus 2024 !
clskms.03k.org
echo ============================================================================
echo #Project: Activating Microsoft software products
echo ============================================================================
echo.
echo #Supported products:
echo - Microsoft Office LTSC Professional Plus 2024
echo.
echo.
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2024VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2024VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
echo.
echo ============================================================================
echo Activating your Office...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /inpkey:XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB >nul
set i=1
:server
if %i%==1 set KMS_Sev=kms.03k.org
if %i%==2 set KMS_Sev=kms.03k.org
if %i%==3 set KMS_Sev=kms.03k.org
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
echo ============================================================================
echo.
echo.
cscript //nologo ospp.vbs /act | find /i "successful" 

 (echo.
echo ============================================================================
echo #Please feel free to contact me at admin@gmail.com if you have any questions or concerns.
echo ============================================================================
choice /n /c YN /m "Would you like to visit my Page [Y,N]?" 
 if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... 
 echo Please wait... 
 echo. 
 echo. 
 set /a i+=1 
 goto server)
explorer "https://github.com/kwokaa2019"
goto halt
:notsupported
echo.
echo ============================================================================
echo Sorry! Your version is not supported.
echo.
:halt
pause >nul</code></pre></div>
            </div>



<p>Microsoft Office LTSC Professional Plus 2024 ：<br/></p>
<pre><code>if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2024VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2024VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /inpkey:XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
cscript ospp.vbs /sethst:192.168.xxx.xxxx 
cscript ospp.vbs /unpkey:8MBCX
cscript ospp.vbs /act</code></pre></div>
            </div>

<p>Microsoft Office LTSC Professional Plus 2021 ：<br/></p>
<pre><code>if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:192.168.xxx.xxxx 
cscript ospp.vbs /unpkey:8MBCX
cscript ospp.vbs /act</code></pre></div>
            </div>
  

<p> Office Professional Plus 2019 ：<br/></p>
<pre><code>if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:192.168.xxx.xxxx 
cscript ospp.vbs /unpkey:8MBCX
cscript ospp.vbs /act</code></pre></div>
            </div>


pause</code></pre></div>
            </div>
