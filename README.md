# VARIOS
Instaladores y archivos varios

//////////////////////////////////////////////////////////////////////////////
                                 PRO 2019
//////////////////////////////////////////////////////////////////////////////

cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:kms8.msguides.com
cscript ospp.vbs /act


//////////////////////////////////////////////////////////////////////////////
                            "quitar letrero 1"
//////////////////////////////////////////////////////////////////////////////

cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus
cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /unpkey:6MWKP

cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:s8.now.im
cscript ospp.vbs /act

//////////////////////////////////////////////////////////////////////////////
                            "quitar letrero 2"
//////////////////////////////////////////////////////////////////////////////

cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /dstatus
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /unpkey:VMFTK

cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:s8.now.im
cscript ospp.vbs /act

//////////////////////////////////////////////////////////////////////////////
                             2021 LTSC
//////////////////////////////////////////////////////////////////////////////
--------------DESCARGAR-----------------
"implementar office ltsc 2021" O
https://www.microsoft.com/en-us/download/details.aspx?id=49117

------------configuration.xml-----------

<Configuration>

	<Add OfficeClientEdition="64" Channel="PerpetualVL2021">
		<Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
			<Language ID="es-es" />
		</Product>

		<Product ID="VisioPro2021Volume">
			<Language ID="es-es" />
		</Product>

		<Product ID="ProjectPro2021Volume">
			<Language ID="es-es" />
		</Product>

	</Add>
	<RemoveMSI />
	<Updates Enabled="TRUE" />
	<Property Name="AUTOACTIVATE" Value="1" />

</Configuration>

--------------INSTALADOR---------------
cd *ruta*
setup /configure configuration.xml

--------------ACTIVADOR--------------- 

cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cd /d %ProgramFiles%\Microsoft Office\Office16

for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:e8.us.to
cscript ospp.vbs /act
