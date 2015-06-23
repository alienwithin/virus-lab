Attribute VB_Name = "modSetup"
' This module will handle setup Functions Like Lab Creation, healing distortions and getting some machine information
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, _
nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function createLab()
    'This function creates a folder in the application's current directory to run tests most especially file infection by overwriting
    Dim LabPath As String
    LabPath = App.Path + "\vx_lab"
    If Dir(LabPath) <> "" Then
    MsgBox "Seems like the lab already exists at:" + LabPath
    ElseIf Dir(LabPath) = "" Then
    MkDir (LabPath)
    MsgBox "The Lab has been created in: " & vbNewLine & LabPath & vbNewLine & "Use this directory to do the file infection test", vbInformation, "Lab Created"
    End If
End Function

Public Function GetID() As String
    '************************************************************************************
    '*
    '* This function returns the ID of the currently logged in user.
    '*
    '************************************************************************************
    On Error Resume Next
    
    Const lLength As Long = 30
    
    Dim nErrorAction As Integer
    Dim sID As String
    Dim lResult As Long
    
    
    sID = Space(lLength)
    lResult = GetUserName(sID, lLength)

    sID = Trim$(sID)
    sID = Left$(sID, Len(sID) - 1)

    GetID = sID
    
    End Function
    
    Public Function GetComp() As String
    '************************************************************************************
    '*
    '* This function returns the ID of the currently logged in user domain.
    '*
    '************************************************************************************
    On Error Resume Next
    
    Const strlen As Long = 30
    
    
    Dim cName As String
    Dim lResult As Long

    cName = Space(strlen)
    lResult = GetComputerName(cName, strlen)
    

    cName = Trim$(cName)
    cName = Left$(cName, Len(cName) - 1)

    GetComp = cName

End Function


Public Sub initialInterface()
    Lab.imgInnocent.Visible = False
    Lab.imgCruel.Visible = False
    Lab.imgvxcopy1.Visible = False
    Lab.imgvxcopy2.Visible = False
    Lab.imgRegistry.Visible = False
    Lab.imgregwriteBad.Visible = False
    Lab.imgFWDisabled.Visible = False
    Lab.imgFWEnabled.Visible = False
    Lab.imgvxtofw.Visible = False
    Lab.lbltitle.Caption = "Welcome to Munir's Virus Lab"
    Lab.lblExplanation.Caption = "The features available in this setup are: " & vbNewLine & "- File Infector" & vbNewLine & "- Replicator" & vbNewLine & "- Service Disabling (Taskmanager, Run, Folder Options, Registry etc.)" & vbNewLine & "- Distortion Reversers" & vbNewLine & vbNewLine & "For more information or suggestions. http://munir.skilledsoft.com"
End Sub

Public Sub distortionFix()
    ' Enable Firewall
    Dim fwMgr
    Set fwMgr = CreateObject("HNetCfg.FwMgr")
    Dim profile
    Set profile = fwMgr.LocalPolicy.CurrentProfile
    profile.FIREWALLENABLED = True
    ' Registry Cleanup
    Set Get_shell = CreateObject("Wscript.Shell")
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoRun", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLogOff", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoClose", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoFolderOptions", 0, "REG_DWORD"
    Get_shell.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoWindowsUpdate", 0, "REG_DWORD"
    MsgBox "All Distortions have been reversed please note that File infection and replication are not included", vbInformation, "Distortion Healer"
End Sub
