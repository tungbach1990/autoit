#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include "ObjectDynamicVari.au3"
Opt("SendCapslockMode", 0)
_InitLog("example")


$config = _InitObjectVari("vari.properties")
$dangnhapaddress = $config.dangnhapaddress
$SplitPath = StringSplit($config.userpath, ",")
$client_current = ProcessList("Game.exe")


;MsgBox("", "",Hex($SplitPath[1]))

for $i = 1 to $config.soluong
	do 		$client_current = ProcessList("Game.exe")

	   sleep(1000)
	   TrayTip($client_current[0][0], "", 4)
	Until $client_current[0][0] < $config.clientmax
	$Thread1 = ThreadCreate("Pro1")
	$Thread1.Start
	$Thread1.Send(Execute ("$config.nick" & $i))
	sleep( $config.thoigiancho * 1000)
Next


Sleep(120000)

Func userpasscallback($userpass)
	#NoTrayIcon
	_InitLog("thread")
	$userpass = SplitUserPass($userpass)
	$config = _InitObjectVari("vari.properties")
	$pid = ShellExecute($config.duongdan); get ProcessID trước
	sleep(1000)
	$handle = _WinGetByPID($pid)
	WinSetTitle($handle, "",$userpass.user )
	sleep(5000)
	
	$memory_object = MemoryObject($pid);
	$memory_object.Open
	$memory_object.GetBase("Game.exe")
	while $memory_object.Read($config.dangnhapaddress, "byte").RtRead = 0 
		ControlClick($handle, "", "", "left", 1, 717, 654)
		Sleep(1000)
	Wend
	Sleep(1000)
	ControlSend($handle, "", "", $userpass.pass)
	$memory_object.MultiWrite($config.usercount,"20", "dword")
	$memory_object.MultiWrite($config.userpath,$userpass.user, "wchar[20]")
	$memory_object.MultiWrite($config.passpath,$userpass.pass, "wchar[" & String(StringLen($userpass.pass)) & "]")
	sleep(500)
	ControlClick($handle, "", "", "left", 1, 563, 561)
	do 
		Sleep(100)
	until $memory_object.Read($config.dangnhapaddress, "byte").RtRead = 7 
	sleep(1000)
	ControlSend($handle, "", "", "{ENTER}")
	ControlClick($handle, "", "", "left", 1, 1286, 829)
	do 
		Sleep(100)
		ControlClick($handle, "", "", "left", 1, 1286, 829)
	until $memory_object.Read($config.dangnhapaddress, "byte").RtRead = 13 

	sleep(1000)
	ControlClick($handle, "", "", "left", 1, 372, 523)
	sleep(1000)
	ControlClick($handle, "", "", "left", 1, 1244, 543)
	ControlClick($handle, "", "", "left", 1, 1244, 543)
	ControlClick($handle, "", "", "left", 1, 1240, 525)
	;WinSetState ($handle,"", @SW_MINIMIZE )
EndFunc

Func Pro1()
	Thread_RegisterFunc("userpasscallback")
	While 1
		sleep(100)
	Wend
	
EndFunc   ;==>CoProc1


Func SplitUserPass($string)
	$object =  __createObject()
	$split = StringSplit($string, ",")
	$object.user = $split[1]
	$object.pass = $split[2]
	return $object
EndFunc


Func _WinGetByPID($iPID, $iArray = 1) ; 0 Will Return 1 Base Array & 1 Will Return The First Window.
    Local $aError[1] = [0], $aWinList, $sReturn
    If IsString($iPID) Then
        $iPID = ProcessExists($iPID)
    EndIf
    $aWinList = WinList()
    For $A = 1 To $aWinList[0][0]
        If WinGetProcess($aWinList[$A][1]) = $iPID And BitAND(WinGetState($aWinList[$A][1]), 2) Then
            If $iArray Then
                Return $aWinList[$A][1]
            EndIf
            $sReturn &= $aWinList[$A][1] & Chr(1)
        EndIf
    Next
    If $sReturn Then
        Return StringSplit(StringTrimRight($sReturn, 1), Chr(1))
    EndIf
    Return SetError(1, 0, $aError)
EndFunc   ;==>_WinGetByPID