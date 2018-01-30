'###############################################################
'### Establece o crea clave de registro BasicAuthLevel a '2' ###
'### para permitir que Office abra documentos directamente   ###
'### desde Alfresco sin certificado SSL                      ###
'###############################################################
'### 06/07/2016 - Daniel Martinez                            ###
'###############################################################

wscript.echo "Éste proceso puede tardar de 1 a 5 minutos, puedes usar el equipo mientras tanto y saldrá otra ventana cuando esté listo"

strComputer = "." 
Set objShell = WScript.CreateObject("WScript.Shell")
Set comprobacion = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
        ("Select * from Win32_OperatingSystem")
		
'Buscando versiones de office
Set colSoft = objWMIService.ExecQuery("SELECT * FROM Win32_Product WHERE Name Like 'Microsoft Office%'")
	

If colSoft.Count = 0 Then
	wscript.echo "ERROR: No se ha detectado Office instalado" 
else
	'On Error Resume Next
	
    For Each objItem In colSoft
		' Estableciendo BasicAuthLevel para cada Office encontrado HKCU
		key = "HKEY_CURRENT_USER\Software\Microsoft\Office\"& Left(objItem.Version, InStr(1,objItem.Version,".")-1) & ".0\Common\Internet\BasicAuthLevel"
		objShell.RegWrite key, 2 , "REG_DWORD"
		wscript.echo "OK, configurada clave para version de Office: " & Left(objItem.Version, InStr(1,objItem.Version,".")-1) & ".0"
	exit for
	Next 
	
	
	
End If
