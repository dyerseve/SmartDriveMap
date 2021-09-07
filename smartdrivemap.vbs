'Modified from https://ss64.com/vb/syntax-mapdrivepersistent.html
Option Explicit
Function MapDrivePersistent(strDrive,strPath)
   ' strDrive = Drive letter - e.g. "x:"
   ' strPath = Path to server/share - e.g. "\\server\share"
   ' Returns a boolean (True or False)

   Dim objNetwork, objDrives, blnFound, objReg, i
   Dim strLocalDrive, strRemoteShare, strRemembered, strMessage
   Const HKCU = &H80000001

   ' Check syntax of parameters passed
   If Right(strDrive, 1) <> ":" OR Left(strPath, 2) <> "\\" Then
      'WScript.echo "Usage: MapDrivePersistent.vbs ""X:"" ""\\server\share"" //NoLogo"
     WScript.Quit(1)
   End If

   Err.clear
   MapDrivePersistent = False

   Set objNetwork = WScript.CreateObject("WScript.Network")

   'Step 1: Get the current drives
   Set objDrives = objNetwork.EnumNetworkDrives
   If Err.Number <> 0 Then
        'Code here for error logging
        Err.Clear
        MapDrivePersistent = False
        Exit Function 
   End If

   'WScript.echo "   Connecting drive letter: " + strDrive + " to " + strPath
    
   'Step 2: Compare drive letters to the one requested
   blnFound = False
   For i = 0 To objDrives.Count - 1 Step 2
        If UCase(strDrive) = UCase(objDrives.Item(i)) Then
            blnFound = True
            'Drive letter was found.  Now see if the network share on it is the same as requested
            If UCase(strPath) = UCase(objDrives.Item(i+1)) Then
                'Correct mapping on the drive
                MapDrivePersistent = True
            Else
                'Wrong mapping on drive.  Disconnect and remap
                'WScript.Echo "--"
                objNetwork.RemoveNetworkDrive strDrive, True, True 'Disconnect drive
                If Err.Number <> 0 Then
                    'Code here for error logging
                    Err.clear
                    MapDrivePersistent = False
                    Exit Function
                End If

                ' To completely remove the previous remembered persistent mapping
                ' we also delete the associated registry key HKCU\Network\Drive\RemotePath
                ' In theory this should be covered by bUpdateProfile=True in
                ' the RemoveNetworkDrive section above but that doesn't always work.
                 Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
                 objReg.GetStringValue HKCU, "Network\" & Left(strDrive, 1), "RemotePath", strRemembered
                 If strRemembered <> "" Then
                   objReg.DeleteKey HKCU, "Network\" & Left(strDrive, 1)
                 End If

               ' Connect drive
               On Error Resume Next
               'WScript.Echo "++"
                objNetwork.MapNetworkDrive strDrive, strPath, True 
                If Err.Number <> 0 Then
                    'Code here for error logging
                    Err.clear
                    MapDrivePersistent = False
                    Exit Function 
                End If

                MapDrivePersistent = True
                
            End If
        End If
        
    Next'Drive in the list
    
   'If blnFound is still false, the drive letter isn't being used.  So let's map it.
   If Not blnFound Then
        On Error Resume Next
        objNetwork.MapNetworkDrive strDrive, strPath, True
        If Err.Number <> 0 Then
            'Code here for error logging
            Err.clear
            MapDrivePersistent = False
            Exit Function 
        End If

        MapDrivePersistent = True
   End If

   'WScript.Echo "   ____"
End Function


if not MapDrivePersistent("S:","\\AMIDC\accounting") Then
    'Wscript.Echo "   ERROR: Drive S: failed to connect!"
End If
if not MapDrivePersistent("Q:","\\AMIDC\quotes") Then
    'Wscript.Echo "   ERROR: Drive Q: failed to connect!"
End If
if not MapDrivePersistent("J:","\\AMIDC\jobcost") Then
    'Wscript.Echo "   ERROR: Drive J: failed to connect!"
End If
if not MapDrivePersistent("O:","\\AMIDC\office") Then
    'Wscript.Echo "   ERROR: Drive O: failed to connect!"
End If
if not MapDrivePersistent("X:","\\AMIDC\common") Then
    'Wscript.Echo "   ERROR: Drive X: failed to connect!"
End If
