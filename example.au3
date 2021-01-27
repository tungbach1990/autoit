#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include "ObjectDynamicVari.au3"

_InitLog("example")

$config = _InitObjectVari("vari.properties")

$Form1 = GUICreate("Main Process", 300, 100, 350, 125)
;~ $Label1 = GUICtrlCreateLabel("5", 10, 40, 148, 52)
;~ $CoProc1Label = GUICtrlCreateLabel("CoProc1:", 110, 40, 148, 30)
$bt = GUICtrlCreateButton("TEST",140,70,80,30)
;~ $CoProc1Process = _CoProc("CoProc1")
;~ $CoProc2Label = GUICtrlCreateLabel("CoProc2:", 210, 40, 148, 52)
$i1 = 5
GUISetState(@SW_SHOW)

$Thread1 = ThreadCreate("Pro1")
$Thread1.Start
$Thread1.MainCallBack("MsgCB")

While 1
	$msg = GUIGetMsg()
	Switch $msg
		Case -3
			Exit
		Case $bt
			$Thread1.Send("Từ main gọi thread")
	EndSwitch
WEnd

Func MsgCB($reic)
	MsgBox("", "","message: " & $reic )
	GUICtrlSetData($Label3, 3)
EndFunc

Func Pro1()
	$Form2 = GUICreate("CoProc1", 100, 100, 350, 260)
	Global $Label3 = GUICtrlCreateLabel("1", 49, 49, 148, 52)
	;_CoProc_Reciver("CoProcMessageReceiver")                  ;Nhân tin nhắn và hiện vào $label 3
;~ 	$i2 = 1
	Thread_RegisterFunc("MsgCB")
	GUISetState(@SW_SHOW)
	While 1
		$msg = GUIGetMsg()
		Switch $msg
			Case -3
				Thread_SendToMain("từ thread gọi tới main")
		EndSwitch
	WEnd
EndFunc   ;==>CoProc1
