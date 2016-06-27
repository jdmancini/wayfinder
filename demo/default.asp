<%
Dim displayCode  'holds application variable Application("displayStatus")
Dim displayMessage 'display message
Dim displayImage 'display image rotation
Dim displayLevel 'display level
Dim fadeTime 'set fade time between images

Dim objFSO 'file system object to find image folder
Dim objFolder 'folder where images reside
Dim i  'array counter variable
Dim strImages()  'array to hold images

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Server.MapPath("images/"))

displayCode = Application("displayStatus")

If displayCode = 2 Then 
	displayEmergency
ElseIf displaycode = 1 Then
	displayEvacuation
ElseIf displaycode = 0 Then
	displayNormal
Else
	Response.Write("Invalid display code.")
End If

Sub displayEmergency
	displayLevel = "emergency"
	i = -1
	For Each objFile in objFolder.Files
		If Left(objFile.Name, 4) = "emrg" Then
			i = i + 1
			Redim Preserve strImages(i)
			strImages(i) = objFile.Name
		End If
	Next
		If i >= 1 Then
		    fadeTime = "1000"
		Else
		    fadeTime = "0"
		End If
End Sub

Sub displayEvacuation
	displayLevel = "evacuation"
	i = -1
	For Each objFile in objFolder.Files
		If Left(objFile.Name, 4) = "evac" Then
			i = i + 1
			Redim Preserve strImages(i)
			strImages(i) = objFile.Name
		End If
	Next
		If i >= 1 Then
		    fadeTime = "1000"
		Else
		    fadeTime = "0"
		End If
End Sub

Sub displayNormal
	displayLevel = "normal"
	i = -1
	For Each objFile in objFolder.Files
		If Left(objFile.Name, 4) = "nrml" Then
			i = i + 1
			Redim Preserve strImages(i)
			strImages(i) = objFile.Name
		End If
	Next
		If i >= 1 Then
			fadeTime = "1000"
		Else
			fadeTime = "0"
		End If
End Sub
%>
<!DOCTYPE html>
<html>
<head>
    <!-- START: meta tags -->
    <meta charset="utf-8" />
    <!-- END: meta tags -->
    <title>PVD Airport</title>
    <!-- START: script links -->
	<script type="text/javascript">var fadeTime = <%= fadeTime %>;</script>
    <script type="text/javascript" src="../js/jquery.min.js"></script>
    <script type="text/javascript" src="../js/rotator.js"></script>
    <script type="text/javascript" src="../js/live.js"></script>
    <!-- END: script links -->
    <!-- START: stylesheet links -->
    <link rel="stylesheet" href="../css/style.css" type="text/css" media="screen" />
    <!-- END: stylesheet links -->
</head>
<body>
    <!-- START: image rotation link -->
	<div id="rotating-item-wrapper">
		<%
		Dim strImage
		For Each strImage in strImages
			Response.Write("<img src=""images/" & strImage & """ class=""rotating-item"" width=""1350"" height=""2410"" alt=""" & displayLevel & """>")
		Next
		%>
	</div>
    <!-- END: image rotation link -->
</body>
</html>
