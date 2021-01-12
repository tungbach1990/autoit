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

;$config.choose = $CmdLine[1]
Dim $a = [3, 2]
$config.__defineGetter("Msg", "abc")
$config.Msg

Func abc($this)
	MsgBox("", $this.parent.title, $this.parent.choose)
EndFunc
