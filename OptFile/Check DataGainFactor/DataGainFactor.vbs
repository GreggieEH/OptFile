' DataGainFactor.vbs
' check the data gain factor for a data file

dim optFile
dim dataFile
dim tstWave
dim fDone		' completed flag
Set objArgs = WScript.Arguments
' read data file from arguments
dataFile = objArgs (0)
set optFile = CreateObject("Sciencetech.OPTFile.1")
if optFile.LoadFromFile(dataFile) then
	fDone = false
	while not fDone
		tstWave = inputBox("Enter test wavelength or negative value to stop")
	
		if tstWave > 0.0 then
			msgBox "Data gain factor for wave = " & tstWave & " = " & optFile.GetADGainFactor(tstWave)
		else
			fDone = true
		end if
	wend
else
	msgBox "Failed to load " & dataFile
end if