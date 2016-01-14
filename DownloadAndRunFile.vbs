Sub Run(ByVal sFile)
  Dim shell
  Set shell = CreateObject("WScript.Shell")
  shell.Run Chr(34) & sFile & Chr(34), 1, false
  Set shell = Nothing
End Sub

'Settings
strFileURL    = "http://example.com"
strHDLocation = "Download.html"
successMsg    = "File downloaded successfully."
errorMsg      = "File could not be downloaded."

Set args = WScript.Arguments
intCount = args.Count

If intCount > 3 Then
  successMsg = args.Item(3)
End If
If intCount > 2 Then
  errorMsg = args.Item(2)
End If
If intCount > 1 Then
  strHDLocation = args.Item(1)
End If
If intCount > 0 Then
  strFileURL  = args.Item(0)
End If

'Fetch the file
On Error Resume Next

runFileFlag = true

Set objFSO = Createobject("Scripting.FileSystemObject")
'Check if file already exists (in case of updates rather than fresh downloads)
fileExistsFlag = objFSO.Fileexists(strHDLocation)

Set objXMLHTTP = CreateObject("Msxml2.XMLHttp.6.0")
objXMLHTTP.open "GET", strFileURL, false
objXMLHTTP.send()

If Err.Number <> 0 Then
  
  If fileExistsFlag Then  
    intAnswer = Msgbox(errorMsg & vbNewLine & _
      "Details: " & Err.Description & vbNewLine & _ 
      "Do you want to run the local version?", _
      vbYesNo + vbExclamation, "Download failed")                         
    If intAnswer = vbNo Then
      runFileFlag = false
    End If
  Else 
    Msgbox errorMsg & vbNewLine & "Details: " & Err.Description, _
      vbExclamation, "Download failed"
	runFileFlag = false
  End If
  
  Set objFSO = Nothing
  Err.Clear
  
ElseIf objXMLHTTP.Status = 200 Then
  
  Set objADOStream = CreateObject("ADODB.Stream")
  objADOStream.Open
  objADOStream.Type = 1 'adTypeBinary

  objADOStream.Write objXMLHTTP.ResponseBody
  objADOStream.Position = 0 'Set the stream position to the start

  If fileExistsFlag Then objFSO.DeleteFile strHDLocation
  Set objFSO = Nothing

  objADOStream.SaveToFile strHDLocation
  objADOStream.Close
  
  Msgbox successMsg, 0, "File downloaded"
  
  Set objADOStream = Nothing
  
End if

Set objXMLHTTP = Nothing

'Run the file
If runFileFlag Then
  Run strHDLocation
End If