#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include "ObjectDynamicVari.au3"

_InitLog("example")

$config = __createObject()
$config.title = "Auto Cóc Ghẻ - tungbach1"
;$config.title = $CmdLine[2]
$config.choose = "quytu"
;$config.choose = $CmdLine[1]
Dim $a = [3, 2]
;$config.__defineGetter("Msg", "abc")

$a = 1
MsgBox("", "", Execute("a"))