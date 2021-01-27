#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <GUIConstants.au3>
#include "ObjectDynamicVari.au3"
#Region Main Process
$Form1 = GUICreate("Main Process", 300, 100, 350, 125)
;~ $Label1 = GUICtrlCreateLabel("5", 10, 40, 148, 52)
;~ $CoProc1Label = GUICtrlCreateLabel("CoProc1:", 110, 40, 148, 30)
$bt = GUICtrlCreateButton("TEST",140,70,80,30)
;~ $CoProc1Process = _CoProc("CoProc1")
;~ $CoProc2Label = GUICtrlCreateLabel("CoProc2:", 210, 40, 148, 52)
$i1 = 5
GUISetState(@SW_SHOW)
While 1
	$msg = GUIGetMsg()
	Switch $msg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $bt
			_CoProc_Send($gi_CoProcParent, "OK")  ;Tin nhắn cần gửi
			 $Process1 = _CoProc_Create("CoProc1")
	EndSwitch
WEnd
ProcessClose($Process1)
Exit
Func CoProcMessageReceiver($vParameter)
;~ 	$ParsedMessage = StringLeft($vParameter, 3)
		GUICtrlSetData($Label3, $vParameter)
EndFunc   ;==>CoProcMessageReceiver
 
Func CoProc1()
	$Form2 = GUICreate("CoProc1", 100, 100, 350, 260)
	Global $Label3 = GUICtrlCreateLabel("1", 49, 49, 148, 52)
	_CoProc_Reciver("CoProcMessageReceiver")                  ;Nhân tin nhắn và hiện vào $label 3
;~ 	$i2 = 1
	GUISetState(@SW_SHOW)
	While 1
		$msg = GUIGetMsg()
		Switch $msg
			Case $GUI_EVENT_CLOSE
				Exit
		EndSwitch
	WEnd
EndFunc   ;==>CoProc1