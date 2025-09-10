'This script checks to see if any version of Adobe Reader and optional components are
'installed. Any non-approved installations are removed and the new installation is completed.
'Last Update: 7-19-2010

Dim objShell, objReg
Dim strKeyPath, arrSubKeys, subKey, WinArch
Dim InstalledAppName, usrResponse
Dim strComputer, uninstString, GUID, DisplayName, SHGInstall, softVersion
const HKEY_LOCAL_MACHINE = &H80000002
Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next

'Checks to see if the PC is 32 or 64-bit
WinArch = objShell.RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")

Select Case WinArch

	Case "x86"

		Set SHGInstall = objShell.RegRead(HKEY_LOCAL_MACHINE & "\SOFTWARE\Adobe\Acrobat Reader\9.0\SHGInstall")
		
		If SHGInstall = "Yes" Then
		
			strComputer = "."
			Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
			
			objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
			
			For Each subKey In arrSubKeys
			    DisplayName = "Adobe Reader"
				InstalledAppName = ""
				InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayName")
			
				If InStr(InstalledAppName, DisplayName) > 0 Then
						softVersion = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayVersion")
						
						'This section is to be updated with each software patch that is released
						Select Case "softVersion"
						
							'This is the most current version at the time this script was published.
							Case "9.3.3"
								WScript.Quit
								
						End Select
			    
				End If
			
			Next
		
		Else
			
			'This section checks for and removes Adobe Reader 3.x
			rdrVersion = objShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Acrobat Reader 3.01\DisplayName") 

			If rdrVersion="Adobe Acrobat Reader 3.01" then
				objShell.run "C:\WINDOWS\ISUNINST.EXE -a -fC:\Acrobat3\Reader\DeIsL1.isu", 1, True  
		
			End If
			
			'This section checks for and removes Adobe Reader 4.x
			softVersion = objShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Acrobat 4.0\DisplayVersion") 

			If softVersion="4.0" then
				objShell.run "C:\WINDOWS\ISUNINST.EXE -a -f""C:\Program Files\Common Files\Adobe\Acrobat 4.0\NT\Uninst.isu"" -c""C:\Program Files\Common Files\Adobe\Acrobat 4.0\NT\Uninst.dll""", 1, True  
				objShell.run "rmdir /s /q ""C:\Program Files\Common Files\Adobe\Web\""", 1, True
			
			End If
			
			'This section checks for and removes Adobe Reader 5.x
			rdrVersion = objShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Acrobat 5.0\DisplayVersion") 
	
			If rdrVersion="5.0" then
				objShell.run "C:\WINDOWS\ISUNINST.EXE -a -f""C:\Program Files\Common Files\Adobe\Acrobat 5.0\NT\Uninst.isu"" -c""C:\Program Files\Common Files\Adobe\Acrobat 5.0\NT\Uninst.dll""", 1, True  
				objShell.run "rmdir /s /q ""C:\Program Files\Common Files\Adobe\Web\""", 1, True
			
			End If
		
			'This section will find and remove all software related to Acrobat Reader 6-9
			strComputer = "."
			Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
			
			objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
			
			For Each subKey In arrSubKeys
			    DisplayName = "Adobe Reader"
				InstalledAppName = ""
				InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayName")
			
				If InStr(InstalledAppName, DisplayName) > 0 then
				uninstString = ""
				GUID = ""
			    
				uninstString = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\UninstallString")
				GUID = Mid(uninstString, instr(uninstString, "{"), 38)
				
					If GUID<>"" then
						objShell.Run "msiexec /x " & GUID & " /qn", 1, True
						WScript.Echo "Done!"
					End If
			    
				End If
			
				'This section looks for a variation of Acrobat Reader 6
				DisplayName = "Adobe Acrobat - Reader 6"
				InstalledAppName = ""
			    InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\Comments")
			
			    If InStr(InstalledAppName, DisplayName) > 0 then
				uninstString = ""
				GUID = ""
			    
				uninstString = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\UninstallString")
				GUID = Mid(uninstString, instr(uninstString, "{"), 38)
				
					If GUID<>"" then
						objShell.Run "msiexec /x " & GUID & " /qn", 1, True
					End If
			    
				End If
			
				'This ection searches for the Adobe Atmosphere Player from Acrobat 6
				DisplayName = "Adobe Atmosphere Player"
				InstalledAppName = ""
			    InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayName")
			
			    If InStr(InstalledAppName, DisplayName) > 0 then
				uninstString = ""
				GUID = ""
			    
				uninstString = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\UninstallString")
				GUID = Mid(uninstString, instr(uninstString, "C"), 21)
				
					If GUID<>"" then
						objShell.Run "C:\WINDOWS\atmoUn.exe -y -a", 1, True
					End If
			    
				End If
			
			Next

			'This command installs Acrobat Reader 
			objShell.Run "msiexec /i \\files\deploy$\Adobe_Acrobat_Reader\AcroRead.msi TRANSFORMS=\\files\deploy$\Adobe_Acrobat_Reader\AcroRead.mst /qn", 1, True

			'This section asks the user to reboot the computer
			strComputer = "."
			Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Shutdown)}!\\" & strComputer & "\root\cimv2")
			Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

			InstalledAppName = "Adobe Acrobat Reader"

			usrResponse = MsgBox("The system must be restarted to apply critical security updates to " & softName & ". Please save and close all work, then click the OK button. ", vbOK)

			If usrResponse = vbOK Then
	
				For Each objOperatingSystem in colOperatingSystems
					ObjOperatingSystem.Reboot()
				Next

			Else
	
				WScript.Echo softName & " will not function until a reboot is completed. Please restart your computer as soon as posible."
				WScript.Quit
				
			End If
		
		End If

	Case "AMD64"
		
		Set Install = objShell.RegRead(HKEY_LOCAL_MACHINE & "\SOFTWARE\Wow6432Node\Adobe\Acrobat Reader\9.0\Install")
		
		If Install = "Yes" Then
		
			strComputer = "."
			Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
			
			objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
			
			For Each subKey In arrSubKeys
			    DisplayName = "Adobe Reader"
				InstalledAppName = ""
				InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayName")
			
				If InStr(InstalledAppName, DisplayName) > 0 Then
						softVersion = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayVersion")
						
						'This section is to be updated with each software patch that is released
						Select Case "softVersion"
						
							'this is the most current version at the time this script was published.
							Case "9.3.3"
								WScript.Quit
								
						End Select
			    
				End If
			
			Next
		
		Else
			
			'This section checks for and removes Adobe Reader 4.x
			softVersion = objShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Acrobat 4.0\DisplayVersion") 

			If softVersion="4.0" then
				objShell.run "C:\WINDOWS\ISUNINST.EXE -a -f""C:\Program Files (x86)\Common Files\Adobe\Acrobat 4.0\NT\Uninst.isu"" -c""C:\Program Files (x86)\Common Files\Adobe\Acrobat 4.0\NT\Uninst.dll""", 1, True  
				objShell.run "rmdir /s /q ""C:\Program Files (x86)\Common Files\Adobe\Web\""", 1, True
			
			End If
			
			'This section checks for and removes Adobe Reader 5.x
			rdrVersion = objShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe Acrobat 5.0\DisplayVersion") 
	
			If rdrVersion="5.0" then
				objShell.run "C:\WINDOWS\ISUNINST.EXE -a -f""C:\Program Files (x86)\Common Files\Adobe\Acrobat 5.0\NT\Uninst.isu"" -c""C:\Program Files (x86)\Common Files\Adobe\Acrobat 5.0\NT\Uninst.dll""", 1, True  
				objShell.run "rmdir /s /q ""C:\Program Files (x86)\Common Files\Adobe\Web\""", 1, True
			
			End If
		
			'This section will find and remove all software related to Acrobat Reader 6-9
			strComputer = "."
			Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
			
			objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
			
			For Each subKey In arrSubKeys
			    DisplayName = "Adobe Reader"
				InstalledAppName = ""
				InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayName")
			
				If InStr(InstalledAppName, DisplayName) > 0 then
				uninstString = ""
				GUID = ""
			    
				uninstString = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\UninstallString")
				GUID = Mid(uninstString, instr(uninstString, "{"), 38)
				
					If GUID<>"" then
						objShell.Run "msiexec /x " & GUID & " /qn", 1, True
						WScript.Echo "Done!"
					End If
			    
				End If
			
				'This section looks for a variation of Acrobat Reader 6
				DisplayName = "Adobe Acrobat - Reader 6"
				InstalledAppName = ""
			    InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\Comments")
			
			    If InStr(InstalledAppName, DisplayName) > 0 then
				uninstString = ""
				GUID = ""
			    
				uninstString = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\UninstallString")
				GUID = Mid(uninstString, instr(uninstString, "{"), 38)
				
					If GUID<>"" then
						objShell.Run "msiexec /x " & GUID & " /qn", 1, True
					End If
			    
				End If
			
				'This ection searches for the Adobe Atmosphere Player from Acrobat 6
				DisplayName = "Adobe Atmosphere Player"
				InstalledAppName = ""
			    InstalledAppName = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\DisplayName")
			
			    If InStr(InstalledAppName, DisplayName) > 0 then
				uninstString = ""
				GUID = ""
			    
				uninstString = objShell.RegRead(HKEY_LOCAL_MACHINE & strKeyPath & "\" & subKey & "\UninstallString")
				GUID = Mid(uninstString, instr(uninstString, "C"), 21)
				
					If GUID<>"" then
						objShell.Run "C:\WINDOWS\atmoUn.exe -y -a", 1, True
					End If
			    
				End If
			
			Next

			'This command installs Acrobat Reader 
			objShell.Run "msiexec /i \\files\deploy$\Adobe_Acrobat_Reader\AcroRead.msi TRANSFORMS=\\files\deploy$\Adobe_Acrobat_Reader\AcroRead.mst /qn", 1, True

			'This section asks the user to reboot the computer
			strComputer = "."
			Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Shutdown)}!\\" & strComputer & "\root\cimv2")
			Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

			InstalledAppName = "Adobe Acrobat Reader"

			usrResponse = MsgBox("The system must be restarted to apply critical security updates to " & softName & ". Please save and close all work, then click the OK button. ", vbOK)

			If usrResponse = vbOK Then
	
				For Each objOperatingSystem in colOperatingSystems
					ObjOperatingSystem.Reboot()
				Next

			Else
	
				WScript.Echo softName & " will not function until a reboot is completed. Please restart your computer as soon as posible."
				WScript.Quit
				
			End If
		
		End If

End Select
