Set objShell = CreateObject("WScript.Shell")

baseDir = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"  										' set your working directory here

Dim Arg, runType, Elements, minscore, mindec, minvmag, minobs, minseen, imageScale, skySeeing, imageOverhead, binning, minHorizon, uncertainty, object, score, ra, dec, vmag, obs, seen, RightAscension, Declination, dJD
Dim minESAPriority
set Arg = WScript.Arguments
runType = Arg(0)

' set your minimums here
minscore 			= 0				' what is the minumum score from the NEOCP, higher score, more desirable for MPC, used for Scheduler priority as well.
minESAPriority  	= 1
mindec 				= -10				' what is the minimum dec you can image at
minvmag 			= 20				' what is the dimmest object you can see
minobs 				= 3					' how many observations, fewer observations mean chance of being lost
minseen 			= 5					' what is the oldest object from the NEOCP, older objects have a good chance of being lost.
focalLength			= 1643
pixelSize			= 6.8
imageScale		 	= 1.71				' your imageScale for determining exposure duration for moving objects
skySeeing 			= 2.8					' your skyseeing in arcsec, used for figuring out max exposure duration for moving objects.
imageOverhead 		= 22 				' how much time to download (and calibrate) added to exposure duration to calculate total number of exposures and repoint
binning 			= 2 				' binning
minHorizon 			= 30				' minimum altitude that ACP/Scheduler will start imaging
maxuncertainty 		= 20				' maximum uncertainty in arcmin from scout for attempt 
getMPCORB 			= False				' do you want the full MPCORB.dat for reference, new NEOCP objects will be appended.
getCOMETS 			= False
getNEOCP 			= True
getESAPri	 		= True

strScriptFile = Wscript.ScriptFullName 													
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
runDir = objFSO.GetParentFolderName(objFile) 
objShell.CurrentDirectory = baseDir        												' E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO

esaPlistLink = "http://neo.ssa.esa.int/PSDB-portlet/plist.txt"							' link to the ESA Priority List.
mpcLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
neocpLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 							' minorplanetcenter URL, shouldnt need to change this
scoutLink = "https://ssd-api.jpl.nasa.gov/scout.api?tdes="								' link to jpl scout for uncertainty determination
newtonLinkBase = "https://newton.spacedys.com/~neodys2/mpcobs/"

neocpTmpFile = baseDir+"\neocp.txt"														' where to put the downloaded neocp.txt, adjust as required.
objectsSaveFile = baseDir+"\object_run.txt"												' where to put output of selected NEOCP objects reference
scoutTmpFile = "\scout.json"															' temporary scout json file
plistTmpFile = basedir+"\plist.txt"																' temporary plist file

mpcTmpFile = "C:\find_o64\mpc_fmt.txt"													' this is the output from find_orb in mpc 1-line element format
obsTmpFile = "C:\find_o64\observations.txt"												' this is from observations from the NEOCP after parsing and filtering NEOCP.txt

mpcorbSaveFile = baseDir+"\MPCORB.dat"													' the raw MPCORB.dat from MPC that we'll append our elements to.
												
fullMpcorbSave = "C:\Program Files (x86)\Common Files\ASCOM\MPCORB\MPCORB.dat"			' this is a copy for ACP should we decide to manually do an object run. 
fullMpcorbLink = "https://minorplanetcenter.net/iau/MPCORB/MPCORB.DAT"	
fullMpcorbDat = "\MPCORB.dat"

fullCometSave = "C:\Program Files (x86)\Common Files\ASCOM\MPCCOMET\CometEls.txt"		' this is a copy for ACP should we decide to manually do an object run. 
cometsLink = "https://minorplanetcenter.net/iau/MPCORB/CometEls.txt"
fullCometDat = "\CometEls.txt"															' I cant remember why I have this one and I'm too tired to figure it out.	
		
Include "VbsJson"	
Dim json, neocpStr, jsonDecoded
Set json = New VbsJson

if runType = "nightly" Then
	if objFSO.FileExists(mpcorbSaveFile) then												' remove the old MPCORB.dat '
		objFSO.DeleteFile mpcorbSaveFile
	end if
	
	if objFSO.FileExists(objectsSaveFile) then												' remove the old object_run.txt 
		objFSO.DeleteFile objectsSaveFile
	end if
	
	call clearACPSched()
	call downloadObjects()
	call updateACPObjects()	
	If getESAPri = True Then
		call getESAObjects()
	End If
End If

If getNEOCP = True Then
	call getNEOCPObjects()
End If 

'if runType = "nightly" Then	
'End If

if objFSO.FileExists(mpcTmpFile) then												' clean up temporary files
	objFSO.DeleteFile mpcTmpFile
end if

if objFSO.FileExists(baseDir+"\NEOCP.rtml") then											' clean up temporary files
	objFSO.DeleteFile baseDir+"\NEOCP.rtml"
end if

if objFSO.FileExists(baseDir+"\plist.txt") then											' clean up temporary files
	objFSO.DeleteFile baseDir+"\plist.txt"
end if

if objFSO.FileExists(neocpTmpFile) then
	objFSO.DeleteFile neocpTmpFile
end if

if objFSO.FileExists(baseDir+scoutTmpFile) then
	objFSO.DeleteFile baseDir+scoutTmpFile
end if

Set objFSO = Nothing
Set objectsFileToWrite = Nothing

Function clearACPSched()
	Set Sched = CreateObject("Acp.Util") 
	If Sched.Scheduler.Available = True Then
		DispStatus = Sched.Scheduler.DispatcherStatus 
		If DispStatus = 5 OR DispStatus = 99 Then
			DispWasEnabled = False 
			call deleteNEOProject()
		Else 
			DispWasEnabled = True
		End If
		If DispStatus = 0 Then
			Sched.Scheduler.DispatcherEnabled = False
			Wscript.Sleep 10000
			call deleteNEOProject()
		ElseIf DispStatus < 5 OR (DispStatus > 5 AND DispStatus < 99) Then
			Wscript.Quit
		End If
		If DispWasEnabled Then
			Sched.Scheduler.DispatcherEnabled = True
		End If
	End If
End Function

Function deleteNEOProject
	Dim RTML, REQ, TGT, ImageSet, FSO, FIL            
	Set RTML = CreateObject("DC3.RTML23.RTML")
	Set RTML.Contact = CreateObject("DC3.RTML23.Contact")
	RTML.Contact.User = "neocp"
		
	Set REQ =  CreateObject("DC3.RTML23.Request")
	REQ.UserName = "neocp"                                  
	REQ.Project = "NEOCP"                                    					' Proj for above user will be created if needed

	Set REQ.Schedule = CreateObject("DC3.RTML23.Schedule")
		
	RTML.RequestsC.Add REQ
	REQ.ID = "dummy"                                         					' This becomes the Plan name for the Request
	REQ.Description = "dummy"
	Set TGT = CreateObject("DC3.RTML23.Target")
	TGT.Name = "dummy"	
	TGT.count = 1	
	REQ.TargetsC.Add TGT
		
	Set ImageSet = CreateObject("DC3.RTML23.Picture")
	ImageSet.Name = "dummy"		
	ImageSet.ExposureSpec.ExposureTime = 1
	ImageSet.Count = 1
	TGT.PicturesC.Add ImageSet
		
	XML = RTML.XML(True)
		
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set FIL = FSO.CreateTextFile(baseDir+"\NEOCP.rtml", True)    		' **CHANGE FOR YOUR SYSTEM**
	FIL.Write XML                                           			' Has embedded line endings
	FIL.Close
	
	Dim I, DB, R
	Set DB = CreateObject("DC3.Scheduler.Database")
	Call DB.Connect()
	Set I = CreateObject("DC3.RTML23.Importer")
	Set I.DB = DB
	I.Import baseDir+"\NEOCP.rtml"
	Set R = I.Projects.Item(0)
	DB.DeleteProject(R)
	Call DB.Disconnect()
	
	Set RTML = Nothing
	Set REQ = Nothing
	Set DB = Nothing
	Set FIL = Nothing
	Set FSO = Nothing
End Function
																													
Function downloadObjects()
	if getCOMETS = true Then
		Wscript.Echo "Downloading CometEls.txt...."
		objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(cometsLink) & " -O" & " " & Quotes(baseDir) & fullCometDat,0,True  	' get the full comets file for reference
		Wscript.Echo "Done"
	End If

	if getMPCORB = true Then
		Wscript.Echo "Downloading MPCORB.dat...."
		objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(fullMpcorbLink) & " -O" & " " & Quotes(baseDir) & fullMpcorbDat,0,True  	' get the full MPCORB.dat file for reference
		Wscript.Echo "Done"
	End If
End Function

Function getESAObjects()
	objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(esaPlistLink) & " -N",0,True								'download ESA Priority List	
	score = "20"
	set plistTmpFileRead = objFSO.OpenTextFile(plistTmpFile, 1) 
	Set objectsFileToWrite = CreateObject("Scripting.FileSystemObject").openTextFile(objectsSaveFile,8,true)  
	Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat output to append NEOCP elements
	
	Do Until plistTmpFileRead.AtEndOfStream																			' read the downloaded neocp.txt and parse for object parameters
		strLine = plistTmpFileRead.ReadLine															
		esaPriority = Mid(strLine,1,1)
		object = replace(Mid(strLine, 5,10),chr(34), chr(32))															' temporary object designation
		dec = Mid(strLine, 27,5)																						' declination 
		vmag = Mid(strLine, 37,4)																						' vMag 
		uncertainty = Mid(strLine, 42,5)
		
		If IsNumeric(esaPriority) Then																					' this line cheats to bypass the date on the first line, doesnt cost much... I dont think
			If (esaPriority <= 0) AND (Csng(dec) >= mindec) AND (Csng(vmag) <= minvmag) AND (Csng(uncertainty) <= maxuncertainty) Then
				wscript.echo object & " " & esaPriority & " " & dec & " " & vmag & " " & uncertainty
				Set objectsFileToRead = objFSO.OpenTextFile(objectsSaveFile,1)
		
				Do Until objectsFileToRead.AtEndOfStream
					currentObject = objectsFileToRead.ReadLine
					if currentObject = object Then
						objectsFileHasMatches = 1
						Exit Do
					Else 
						objectsFileHasMatches = 0
					End If
				Loop
		
			objectsFileToRead.Close
			
				if objectsFileHasMatches = 0 Then	
					objectsFileToWrite.WriteLine(object)
					objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(newtonLinkBase) & Replace(object," ","")  & ".rwo" & " -O" & " " & obsTmpFile,0,True 
					Set objFile = objFSO.OpenTextFile(obsTmpFile, 1)
			
					If (objFile.AtEndOfStream <> True) Then
						objShell.CurrentDirectory = "C:\find_o64"										' lets change our cwd to run find_orb, it likes it's home
						objShell.Run "fo.exe observations.txt",0,True									' open the mpc_fmt.txt that find_orb created and append it to the MPCORB.dat
						objShell.CurrentDirectory = baseDir
						Set objFile = objFSO.OpenTextFile(mpcTmpFile, 1)
						Do Until objFile.AtEndOfStream
							Elements = objFile.ReadLine
							MPCorbFileToWrite.WriteLine(Elements)										'write elemets to MPCORB.dat
							Wscript.Sleep 10000
							call buildObjectDB(object, vmag,  seen, obs, uncertainty, Minutes)
						Loop
					objFile.Close
					End If
				End If
			End If
		End If
	Loop
End Function

Function getNEOCPObjects()	
	objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(neocpLink) & " -N",0,True 									'download current neocp.txt from MPC 
	Set neocpTmpFileRead = objFSO.OpenTextFile(neocpTmpFile, 1) 														' change path for input file from wget 
	Set objectsFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(objectsSaveFile,8,true)  		' create object_run.txt

	if objFSO.FileExists(mpcorbSaveFile) then												
		Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat output to append NEOCP elements
	else
		Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").CreateTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat didnt exist for some reason, lets create and empty one
	End If
	
	Do Until neocpTmpFileRead.AtEndOfStream													' read the downloaded neocp.txt and parse for object parameters
		strLine = neocpTmpFileRead.ReadLine													' its probably a good idea NOT to touch the positions as they are fixed position.
		object = Mid(strLine, 1,7)															' temporary object designation
		score = Mid(strLine, 9,3)															' neocp desirablility score from 0 to 100, 100 being most desirable.
		ra	  = Mid(strLine, 27,7)															' right ascension 
		dec = Mid(strLine, 35,6)															' declination 
		vmag = Mid(strLine, 44,4)															' vMag 
		obs = Mid(strLine, 79,4)															' how many observations has it had
		seen = Mid(strLine, 96,7)															' when was the object last seen
	
		objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(scoutLink) & object & " -O" & " " & Quotes(baseDir) & scoutTmpFile,0,True ' Get NEOCP from Scout
		scoutStr = objFSO.OpenTextFile(baseDir+scoutTmpFile).ReadAll
		Set jsonDecoded = json.Decode(scoutStr)
		uncertainty = jsonDecoded("unc")													' position uncertainty in arcmin
	
		if (CSng(score) >= minscore) AND (CSng(dec) >= mindec) AND (CSng(vmag) <= minvmag) AND (CSng(obs) >= minobs) AND (CSng(seen) <= minseen) AND ((CSng(uncertainty) <= maxuncertainty) AND (uncertainty <> "")) Then
		
			Set objectsFileToRead = objFSO.OpenTextFile(objectsSaveFile,1)
		
			Do Until objectsFileToRead.AtEndOfStream
				currentObject = objectsFileToRead.ReadLine
				if currentObject = object Then
					objectsFileHasMatches = 1
					Exit Do
				Else 
					objectsFileHasMatches = 0
				End If
			Loop
		
			objectsFileToRead.Close
			
			if objectsFileHasMatches = 0 Then
				objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(mpcLinkBase) & trim(object) & "&obs=y -O" & " " & obsTmpFile,0,True ' run wget to get observations from NEOCP
				
				Set objFile = objFSO.OpenTextFile(obsTmpFile, 1)
				objFile.ReadAll
				
				if objFSO.FileExists(mpcTmpFile) then
					objFSO.DeleteFile mpcTmpFile
				end if
				
				if objFile.Line >= 3 Then
					objShell.CurrentDirectory = "C:\find_o64"										' lets change our cwd to run find_orb, it likes it's home
					
					objShell.Run "fo.exe observations.txt",0,True									' open the mpc_fmt.txt that find_orb created and append it to the MPCORB.dat
					
					if objFSO.FileExists(mpcTmpFile) then	
						objShell.CurrentDirectory = baseDir												' change directory back to where we are running script
						Set objFile = objFSO.OpenTextFile(mpcTmpFile, 1)
						Do Until objFile.AtEndOfStream
							Elements = objFile.ReadLine
							MPCorbFileToWrite.WriteLine(Elements)										'write elemets to MPCORB.dat
						Loop
						objFile.Close
		
					Wscript.Echo object & "     " & score & "  " + ra & "  " & dec & "    " & vmag & "      " & obs & "     " & seen & "   " & uncertainty
					objectsFileToWrite.WriteLine(object)'+ "     " + score + "  " + ra + "  " + dec + "    " + vmag + "      " + obs + "     " + seen + " " + uncertainty)	' write out the objects_run for reference
					Name = object
					call buildObjectDB(object, vmag,  seen, obs, uncertainty, Minutes)
					End If
				End If
			End If
		End If
		
	Loop
	
	neocpTmpFileRead.Close																			' close any open files
	'objectsFileToWrite.Close
	MPCorbFileToWrite.Close
End Function

Sub GetExposureData(expTime,imageCount, objectRate, Minutes)
	
	call getRateFromFO(object, objectRate)
	expTime = round((60*(imageScale/objectRate)*skySeeing),0)
	Minutes = round(18/(objectRate/60))
	
	'If Minutes > 18 Then
		Minutes = 40
	'End If
	If expTime > 60 Then 
		expTime = 60 
	End If
												
	imageCount = round((Minutes*(60/(expTime+imageOverhead))),0)

End Sub

Function buildObjectDB(object, vmag,  seen, obs, uncertainty, Minutes)

	Name = object
	call GetExposureData(expTime,imageCount, objectRate, Minutes)
	
	If imageCount > 0 Then															' to overcome issues when object has been moved to PCCP
	
		set Util = CreateObject("ACP.Util")
		set AUtil = CreateObject("ASCOM.Utilities.Util")
		Dim RTML, REQ, TGT, ImageSet, COR, FSO, FIL, TR                 
		Set RTML = CreateObject("DC3.RTML23.RTML")
		Set RTML.Contact = CreateObject("DC3.RTML23.Contact")
		RTML.Contact.User = "neocp"
		RTML.Contact.Email = "brians@fl240.com"
		
		Set REQ =  CreateObject("DC3.RTML23.Request")
		REQ.UserName = "neocp"                                  
		REQ.Project = "NEOCP"                                    					' Proj for above user will be created if needed

		Set REQ.Schedule = CreateObject("DC3.RTML23.Schedule")
		REQ.Schedule.Horizon = minHorizon
		REQ.Schedule.Priority = score
		Set REQ.Schedule.Moon = CreateObject("DC3.RTML23.Moon")
		Set REQ.Schedule.Moon.Lorentzian = CreateObject("DC3.RTML23.Lorentzian")
		REQ.Schedule.Moon.Lorentzian.Distance  = 15
		REQ.Schedule.Moon.Lorentzian.Width = 6
			
		Set REQ.Correction = CreateObject("DC3.RTML23.Correction")
		REQ.Correction.zero = False
		REQ.Correction.flat = False
		REQ.Correction.dark = False
		
		RTML.RequestsC.Add REQ
		REQ.ID = object        		' This becomes the Plan name for the Request
		REQ.Description = object + " Score: " + score + " RA: " + Cstr(round(RightAscension,2)) + " DEC: " + Cstr(round(Declination,2)) + " vMag: " + vmag + " #Obs: " + obs + " Last Seen: " + seen  + " Rate: " + CStr(objectRate) + " arcsec/min" + " Unc: " + trim(uncertainty) + " arcsec"
		
		Set TGT = CreateObject("DC3.RTML23.Target")
		TGT.TargetType.OrbitalElements = Elements
		TGT.Description = Elements
		TGT.Name = object
			
		imageTotalTime = round(((expTime + imageOverhead) * imageCount) / 60)			' in minutes including overhead for download, etc
		
		If objectRate > 30 Then
			TGT.count = 10002
			TGT.Interval = 0
			imageCount = imageCount / 2
			REQ.TargetsC.Add TGT
		Else
			TGT.count = 1
			TGT.Interval = 0
			REQ.TargetsC.Add TGT
		End If
		
		'TGT.count = 1
		'TGT.Timefromprev = round((Minutes + 60) / 60)
		'TGT.Timefromprev = round((imageTotalTime + 16) / 60,2)
		'TGT.Tolfromprev = round((imageTotalTime) / 60,2)
		'REQ.TargetsC.Add TGT
		
		Set ImageSet = CreateObject("DC3.RTML23.Picture")
		ImageSet.Name = object+" Luminance"                               	
		ImageSet.ExposureSpec.ExposureTime = expTime
		ImageSet.Binning = binning
		ImageSet.Filter = "Luminance"
		ImageSet.Description = "#nopreview"	
		ImageSet.Count = imageCount
			
		TGT.PicturesC.Add ImageSet
		
		XML = RTML.XML(True)
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set FIL = FSO.CreateTextFile(baseDir+"\NEOCP.rtml", True)    		
		FIL.Write XML                                           			
		FIL.Close
		
		Dim I, DB, R
		Set DB = CreateObject("DC3.Scheduler.Database")
		Call DB.Connect()

		Set I = CreateObject("DC3.RTML23.Importer")
		Set I.DB = DB
			
		I.Import baseDir+"\NEOCP.rtml"
		Set R = I.Projects.Item(0)
		
		R.Disabled = false
		R.Update()
		Set NewPlan = I.Plans.Item(0)
		NewPlan.Resubmit()
            		
		Call DB.Disconnect()
		Set REQ =  nothing
		Set RTML = nothing
		Set TGT = nothing
		set ImageSet = Nothing
		
	End If
End Function
	
Function getRateFromFO(object, objectRate)
	
	ephemerisFile = "C:\find_o64\ephemeris.txt"
	
	objShell.CurrentDirectory = "C:\find_o64"			
	objShell.Run "fo.exe observations.txt -e ephemeris.txt -C V01",0,True	
	objShell.CurrentDirectory = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"
	
	Set ephemerisObjFile = objFSO.OpenTextFile(ephemerisFile)
	
	Do Until ephemerisObjFile.AtEndOfStream	
		current = ephemerisObjFile.ReadLine
	    if IsNumeric(Mid(current,87,3)) Then
			If (Mid(current,87,3)) > 0 Then
					
				alt2 = alt1
				alt1 = Mid(current,87,3)
			
				If alt1 < alt2 Then
					output = current
					exit Do
				End If
			End If
		End If 
	Loop	
	objectRate = Mid(output, 74,6)	
	Wscript.Echo object & " " & objectRate & " " & Mid(output, 1,17) & " " & alt1 & " " & alt2
	Wscript.Echo " "
End Function

Function updateACPObjects()
	if getMPCORB = True Then
		set mpccopy=CreateObject("Scripting.FileSystemObject")
		mpccopy.CopyFile baseDir+fullMpcorbDat, fullMpcorbSave, True

		objShell.CurrentDirectory = "C:\Program Files (x86)\Common Files\ASCOM\MPCORB"
		Wscript.Echo "compiling mpcorb.dat for ACP"
		objShell.Run "MakeDB.wsf",0,True

		set mpccopy = nothing
		objShell.CurrentDirectory = baseDir
	End If

	if getCOMETS = True Then
		set cmtccopy=CreateObject("Scripting.FileSystemObject")
		cmtccopy.CopyFile baseDir+fullCometDat, fullCometSave, True

		objShell.CurrentDirectory = "C:\Program Files (x86)\Common Files\ASCOM\MPCCOMET"
		objShell.Run "MakeCometDB.wsf",0,True

		set mpccopy = nothing
		objShell.CurrentDirectory = baseDir
	End If
End Function

Function Include(file)
	On Error Resume Next
	
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file & ".vbs", 1).ReadAll()
	Set FSO = Nothing

	If Err.Number <> 0 Then
		If Err.Number = 1041 Then
			Err.Clear
		Else
			WScript.Quit 1
		End If
	End If
End Function

Function LPad(s, l, c)
  Dim n : n = 0
  If l > Len(s) Then n = l - Len(s)
  LPad = String(n, c) & s
End Function

Function Quotes(strQuotes)																' Add Quotes to string
	Quotes = chr(34) & strQuotes & chr(34)												' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
End Function
