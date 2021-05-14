
'==================================================== 

'Script  for post patching validation . 
Created by : Sreejish K Nair ( sreejish.nair@gmail.com)
'=====================================================

Const CONVERT_TO_LOCAL_TIME = True
dim strLogType
dim usercount

intNumberID = 528 ' Event ID Number
intEventType = 4
strLogType = "'Security'"
dim rdpstring
 
On Error Resume Next

Const ForReading = 1
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile ("servers.txt", ForReading)

strFileName = Replace(Date,"/","-") & "-" & Left(Replace(Time,":",""),4) & "-checklist.csv" 


Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(strFileName, _ 
    ForAppending, True)
Set objFSO = CreateObject("Scripting.FileSystemObject")


objLogFile.Writeline

'==================================================== 

'Main loop Starts Here
'=====================================================
objLogFile.Writeline("Host Name,Ping Status,CDriveFreeSpace(GB),Memory,OS,Last Boot Time, Stopped Servies with Auto startup,Count of patch,hotfixes") 
Do Until objTextFile.AtEndOfStream

	strcomputer = Trim(objTextFile.Readline)
	hostno = hostno + 1
        wscript.echo "Processing "&strcomputer
       
	objLogFile.Write(strcomputer) & ","


'==================================================== 

'	Check server Ping Status
'=====================================================

      Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
            & strComputer & "'")
    	For Each objStatus in objPing
                
        	If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
            		objLogFile.Writeline("NotPingable") 
	    		
	    		PingStatus = 0	    
       		else    
                        
                        
	    		objLogFile.Write("Pingable,") 
	    		
	   	     	PingStatus = 1
 		End If
		
    	Next


	If pingstatus = 1 then
  
 'check C Drive Space  
'=====================================================
Const HARD_DISK = 3
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colDisks = objWMIService.ExecQuery ("SELECT * FROM Win32_LogicalDisk WHERE DriveType = " & HARD_DISK & "")
For Each objDisk in colDisks
 	'Wscript.Echo "Device ID: " & objDisk.DeviceID 
        If objDisk.DeviceID ="C:" Then
		'Wscript.Echo "Free Disk Space: " & objDisk.FreeSpace
		'Wscript.Echo "Free Disk Space: " & objDisk.Size
		perc=round((objDisk.FreeSpace/objDisk.Size)*100,2)
		'Wscript.Echo " % Free Disk Space: " & perc&"%"
                objLogFile.Write(round((objDisk.FreeSpace/1000000000),2))&","
       End If
Next





'check Operating System , Service Pack
'=====================================================
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        
	Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each objItem in colItems
        objLogFile.Write(objItem.TotalVisibleMemorySize/1000000)&","
    	objLogFile.Write(objItem.Caption)&","  
        
               	

	Next        
	

'




'check Last boot time
'=====================================================

Set dtmTargetDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
dtmTargetDate.SetVarDate "3/1/2004", LOCAL_TIME


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
 
For Each objOS in colOperatingSystems
    
dtmConvertedDate.Value = objOS.LastBootUpTime
dtmConvertedDate.GetVarDate(LOCAL_TIME)
objLogFile.Write(dtmConvertedDate.GetVarDate(LOCAL_TIME))&","
Wscript.Echo dtmConvertedDate.GetVarDate(LOCAL_TIME)    
Next
 
   
'check Auto services Status
'=====================================================


 str=""


 Set Services = GetObject("winmgmts:\\" & strComputer & "\Root\CIMv2").ExecQuery("SELECT * FROM Win32_Service where State='Stopped' AND StartMode='Auto'")
 
 For Each Service in Services
 str=str&"*"&Service.Caption&"*" 
 
 Next

 


objLogFile.Write str&","

'=====================================================


'check patches installed last month 
'=====================================================

 count=0
 patchstr=""

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colQuickFixes = objWMIService.ExecQuery("SELECT * FROM Win32_QuickFixEngineering")
For Each objQuickFix in colQuickFixes

 If DateDiff("d", objQuickFix.InstalledOn, Now) <= 31 Then
    patchstr=patchstr&"*"&objQuickFix.HotFixID&"*" 
    count=count+1
 End If
Next
objLogFile.Writeline count&","&patchstr

'=====================================================


	
	
	
end if          '--------------------end if of if ping  status =1 

Loop ' ---------- Main Loop ends here 
	
objLogFile.Close ' ----------- Closing the csv file .



Msgbox "Done"








