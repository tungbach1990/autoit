#include-once
#include "Array.au3"
#include <Memory.au3>
#include <WinAPISys.au3>	
#include "WinApi.au3"
#include "Date.au3"
#include <WinHTTP.au3>
#include <GDIPlus.au3>
	;ocrkey = 09c854369688957

	
	#Region Const Vari Text
		Global $gs_SuperGlobalRegistryBase = "HKEY_CURRENT_USER\Software\AutoIt v3\CoProc"
		Global $gi_CoProcParent = 0
		Global $gs_CoProcReciverFunction = ""
		Global $gv_CoProcReviverParameter = 0
		Global $IsLog = False
		$log_prefix_info = "[Info]"
		$log_prefix_err = "[Error]"
		$log_prefix_Memory = "[Memory]"
		$log_prefix_InfoMemory = $log_prefix_info & $log_prefix_Memory
		$log_prefix_ErrMemory = $log_prefix_err & $log_prefix_Memory
		$log_prefix_Thread = "[Thread]"
		$log_prefix_InfoThread = $log_prefix_info & $log_prefix_Thread
		$log_prefix_ErrThread = $log_prefix_err & $log_prefix_Thread
		
	#EndRegion Const Vari Text
	$log_prefix = ''
	$File_Log_Extension = '.log'
	$File_Log_Path = ''
	;$Lang_log = _InitObjectVari("lang.properties")	
	;$Lang_log.Log.Text.Space = " "
	tut_all(); tắt func này để bỏ đoạn hướng dẫn ban đầu
	;WhoAmI()
	Func WhoAmI($info = '')
		;MsgBox("", "", StringInStr(StringLower($info), 'obj'))
		Select 			
			Case $info = '' or StringInStr(StringLower($info), 'tut')
				tut_vari()
				tut_obj()
				tut_thread()
				tut_mem()
				tut_bmp()
				tut_log()
			Case StringInStr(StringLower($info), 'vari')
				tut_vari()
			Case StringInStr(StringLower($info), 'object') or StringInStr(StringLower($info), 'obj')
				tut_obj()
			Case StringInStr(StringLower($info), 'thread') or StringInStr(StringLower($info), 'thr') or StringInStr(StringLower($info), 'coc')
				tut_thread()
			Case StringInStr(StringLower($info), 'mem') or StringInStr(StringLower($info), 'memory')
				tut_mem()
			Case StringInStr(StringLower($info), 'img') or StringInStr(StringLower($info), 'bmp') or StringInStr(StringLower($info), 'search') or StringInStr(StringLower($info), 'scan')
				tut_bmp()
			Case StringInStr(StringLower($info), 'log') or StringInStr(StringLower($info), 'cmt')
				tut_log()
		EndSelect 
	EndFunc
	
	Func tut_all()
		$temp = $IsLog
		$IsLog = False
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI()/WhoAmI('tut') để xem tất cả hướng dẫn")
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI('vari') để xem hướng dẫn về biến dạng Object")
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI('obj')/WhoAmI('object') để xem hướng dẫn về tạo Object")
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI('thr')/WhoAmI('thread')/WhoAmI('coc') để xem hướng dẫn về Threading")
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI('memory')/WhoAmI('mem') để xem hướng dẫn về Memory")
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI('search')/WhoAmI('bmp')/WhoAmI('img') để xem hướng dẫn về tìm ảnh và click ảnh")
		Logging("Hướng dẫn sử dụng thư viện: Hãy chạy hàm WhoAmI('log')/WhoAmI('cmt') để xem hướng dẫn về tìm ghi Log")
		$IsLog = $temp
	EndFunc
	
	Func tut_log()
		$temp = $IsLog
		$IsLog = False
		Logging("######## Log Everything _InitLog() #########")
		Logging("### Log vào file để dễ dang theo dõi bug")
		Logging("Example: _InitLog($id) ; Khai báo bắt đầu chạy log với $id là id của Log, có thể để dạng text")
		Logging("### Log 1 dòng vào file và console")
		Logging("Example: Logging('everything ') ; Log cụm từ everything vào console và log")
		Logging("Example: Logging('everything ',@scriptlinenumber) ; Log cụm từ everything vào console và log ngoài ra có thêm tiền tố số dòng ở trước, có thể thay bằng cụm từ yêu thích")
		Logging("######## Log Everything _InitLog() #########")
		$IsLog = $temp
	EndFunc
	
	Func tut_obj()
		$temp = $IsLog
		$IsLog = False
		Logging(" ###################### Start Tutorial Object ###################")
		Logging("### Thường được sử dụng khi muốn dùng dạng Object Trong AutoIT")
		Logging("###Function create Object __createObject()")
		Logging("Example: $newobject = __createObject()")
		Logging("###Set property cho object: ")
		Logging("Example: $newobject.newproperty = 'xxx';Set property tên là newproperty có giá trị là xxx")
		Logging("###Set property dạng method cho object: ")
		Logging("Example: $newobject.__defineGetter('Open', __MemoryOpen);Set method tên là Open sẽ gọi hàm __MemoryOpen")
		Logging("###Set property dạng method cho object; Hàm được gọi sẽ gọi toàn bộ bằng Object ")
		Logging("Example: Function __MemoryOpen($self)")
		Logging("Example: ControlWrite($self.parent.newproperty) ; ở đây sẽ in ra xxx trong đó parent là sẽ trở về object cha")
		Logging("Example: Return $self.parent;trở về object cha để có thể tiếp tục gọi dạng property.property2.proterty3")
		Logging(" ####################### End Tutorial Object ####################")
		$IsLog = $temp
	EndFunc
	
	
	Func tut_thread()
		$temp = $IsLog
		$IsLog = False
		Logging(" ###################### Start Thread Object ###################")
		Logging("### Thường được sử dụng khi muốn chạy MultiProcess(CoProc) theo dạng Object")
		Logging("### Function create Thread ThreadCreate()")
		Logging("Example: $newthread = ThreadCreate('functionabc'); Trong đó functionabc là 1 func được định nghĩa để chạy ở proc mới")
		Logging("### Start Thread: ")
		Logging("Example: $newthread.Start; có thể thêm agruments để gửi tham số tới functionabc, nhưng khi đó functionabc cần đặt là functionabc($agruments) và trông functionabc sẽ dùng biến là $agruments")
		Logging("### Abort Thread")
		Logging("Example: $newthread.Abort")
		Logging("### Suspend/Pause Thread")
		Logging("Example: $newthread.Suspend; $newthread.Pause")
		Logging("### Resume Thread")
		Logging("Example: $newthread.Resume; $newthread.Resume")
		Logging("### Kết hợp Object Thread")
		Logging("Example: $newthread.Start.Pause.Resume.Abort; chạy, tạm dừng rồi tiếp tục rồi thoát")
		Logging(" ####################### End Thread Object ####################")
		$IsLog = $temp
	EndFunc
	
	
	Func tut_vari()
		$temp = $IsLog
		$IsLog = False
		Logging(" ###################### Start Vari Object ###################")
		Logging("### Thường được sử dụng khi muốn tạo multi language hoặc thay thế các file biến bên ngoài rồi gọi dạng Object")
		Logging("### Function _InitObjectVari( $filename = 'variables.properties' )")
		Logging("Example: $lang = _InitObjectVari('en.properties'); Khai báo object tên là $lang, trong đó en.properties là tên 1 file cùng thư mục")
		Logging("### Nội dung file en.properties")
		Logging("Example: Log.Text.Prefix.Info.Thread=[Info][Thread]")
		Logging("### Chạy code trong au3:")
		Logging("Example: MsgBox('','',$lang.Log.Text.Prefix.Info.Thread') ; sẽ hiển thị ra nội dung trong file en.properties với mỗi dòng là 1 property ")
		Logging(" ####################### End Vari Object ####################")
		$IsLog = $temp
	EndFunc

	
	Func tut_mem()
		$temp = $IsLog
		$IsLog = False
		Logging(" ###################### Start Memory Object ###################")
		Logging("### Thường được sử dụng khi muốn sử dụng tác động Memory gọi dạng Object")
		Logging("### Function MemoryObject($pid) sử dụng để tạo ObjectMemory")
		Logging("Example: $mem = MemoryObject(WinGetProcess('Chrome')); Khai báo object tên là $mem với processid là WinGetProcess Title là Chrome")
		Logging("### OpenMemory ")
		Logging("Example: $mem.Open; Return là Object $mem, sẽ gán $mem.MemID là giá trị trả về của hàm MemoryOpen")
		Logging("### ReadMemory theo Address")
		Logging("Example: $mem.Read; Return là Object $mem, sẽ gán $mem.RtRead là giá trị trả về của hàm MemoryRead khi đọc tại địa chỉ $mem.RtRead trước đos với loại đọc là dword ")
		Logging("Example: $mem.Read($address); Return là Object $mem, sẽ gán $mem.RtRead là giá trị trả về của hàm MemoryRead khi đọc tại địa chỉ $address với loại đọc là dword ")
		Logging("Example: $mem.Read($address,$type); Return là Object $mem, sẽ gán $mem.RtRead là giá trị trả về của hàm MemoryRead khi đọc tại địa chỉ $address với loại đọc là $type ")
		Logging("Example: $mem.Read('Type = char[2]','Address = 0x122421'); Return là Object $mem, sẽ gán $mem.RtRead là giá trị trả về của hàm MemoryRead khi đọc tại địa chỉ 0x122421 với loại đọc là char[2] ")
		Logging("### WriteMemory theo Address")
		Logging("Example: $mem.Write($value); Return là Object $mem, sẽ gán $mem.RtWrite là giá trị trả về của hàm MemoryWrite khi viết tại địa chỉ $mem.RtRead với loại viết là dword, giá trị là $value ")
		Logging("Example: $mem.Write($address,$value); Return là Object $mem, sẽ gán $mem.RtWrite là giá trị trả về của hàm MemoryWrite khi viết tại địa chỉ $address với loại viết là dword, giá trị là $value ")
		Logging("Example: $mem.Write($address,$value,$type); Return là Object $mem, sẽ gán $mem.RtWrite là giá trị trả về của hàm MemoryWrite khi viết tại địa chỉ $address với loại viết là $type, giá trị là $value ")
		Logging("Example: $mem.Write('Type = char[2]','Address = 0x122421','valW = 123'); Return là Object $mem, sẽ gán $mem.RtWrite là giá trị trả về của hàm MemoryRead khi đọc tại địa chỉ 0x122421 với loại đọc là char[2] , giá trị là 123")
		Logging("### CloseMemory")
		Logging("Example: $mem.Close; sẽ đóng Memory với giá trị $mem.MemID  ")
		Logging(" ####################### End Memory Object ####################")
		$IsLog = $temp
	EndFunc
	
	
	Func tut_bmp()
		$temp = $IsLog
		$IsLog = False
		Logging(" ###################### Start BmpSearch Object ###################")
		Logging("### Thường được sử dụng khi muốn tìm ảnh như BMPSearch rồi gọi dạng Object")
		Logging("### BMPSearch($title)")
		Logging("Example: $bmpsearch = BMPSearch($title); Sẽ khởi tạo 1 Object tên là $bmpsearch để scan ảnh 1 ứng dụng có Title là $title. Có thể thay đổi Title bằng $bmpsearch.Title = $newtitle")
		Logging("### Search ảnh trong Title ban đầu")
		Logging("Example: $bmpsearch.Search($anh); Sẽ tìm ảnh với đường dẫn là $anh trong Cửa sổ có Title là $bmpsearch.Title , giá trị trả về chứa trong $bmpsearch.Status là true or false, tọa độ trong $bmpsearch.x và .y")
		Logging("### Click vào ảnh tìm được với lệnh Search ở trên")
		Logging("Example: $bmpsearch.Click ; click với tọa độ $bmpsearch.x và $bmpsearch.y")
		Logging("### Object kết hợp")
		Logging("Example: $bmpsearch.Search('zz.bmp').Click; Click vào tọa độ của ảnh zz.bmp trong hình ảnh của Title $bmpsearch.Title capture được'")
		Logging(" ####################### End BmpSearch Object ####################")
		$IsLog = $temp
	EndFunc
	
	If Not IsDeclared("ThreadArrayAll") Then Global $ThreadArrayAll = [0]
	If Not IsDeclared("IID_IUnknown") Then Global Const $IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
	If Not IsDeclared("IID_IDispatch") Then Global Const $IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
	If Not IsDeclared("IID_IConnectionPointContainer") Then Global Const $IID_IConnectionPointContainer = "{B196B284-BAB4-101A-B69C-00AA00341D07}"
	If Not IsDeclared("DISPATCH_METHOD") Then Global Const $DISPATCH_METHOD =               1
	If Not IsDeclared("DISPATCH_PROPERTYGET") Then Global Const $DISPATCH_PROPERTYGET =          2
	If Not IsDeclared("DISPATCH_PROPERTYPUT") Then Global Const $DISPATCH_PROPERTYPUT =          4
	If Not IsDeclared("DISPATCH_PROPERTYPUTREF") Then Global Const $DISPATCH_PROPERTYPUTREF =       8
	If Not IsDeclared("S_OK") Then Global Const $S_OK = 0x00000000
	If Not IsDeclared("E_NOTIMPL") Then Global Const $E_NOTIMPL = 0x80004001
	If Not IsDeclared("E_NOINTERFACE") Then Global Const $E_NOINTERFACE = 0x80004002
	If Not IsDeclared("E_POINTER") Then Global Const $E_POINTER = 0x80004003
	If Not IsDeclared("E_ABORT") Then Global Const $E_ABORT = 0x80004004
	If Not IsDeclared("E_FAIL") Then Global Const $E_FAIL = 0x80004005
	If Not IsDeclared("E_ACCESSDENIED") Then Global Const $E_ACCESSDENIED = 0x80070005
	If Not IsDeclared("E_HANDLE") Then Global Const $E_HANDLE = 0x80070006
	If Not IsDeclared("E_OUTOFMEMORY") Then Global Const $E_OUTOFMEMORY = 0x8007000E
	If Not IsDeclared("E_INVALIDARG") Then Global Const $E_INVALIDARG = 0x80070057
	If Not IsDeclared("E_UNEXPECTED") Then Global Const $E_UNEXPECTED = 0x8000FFFF
	If Not IsDeclared("DISP_E_UNKNOWNINTERFACE") Then Global Const $DISP_E_UNKNOWNINTERFACE = 0x80020001
	If Not IsDeclared("DISP_E_MEMBERNOTFOUND") Then Global Const $DISP_E_MEMBERNOTFOUND = 0x80020003
	If Not IsDeclared("DISP_E_PARAMNOTFOUND") Then Global Const $DISP_E_PARAMNOTFOUND = 0x80020004
	If Not IsDeclared("DISP_E_TYPEMISMATCH") Then Global Const $DISP_E_TYPEMISMATCH = 0x80020005
	If Not IsDeclared("DISP_E_UNKNOWNNAME") Then Global Const $DISP_E_UNKNOWNNAME = 0x80020006
	If Not IsDeclared("DISP_E_NONAMEDARGS") Then Global Const $DISP_E_NONAMEDARGS = 0x80020007
	If Not IsDeclared("DISP_E_BADVARTYPE") Then Global Const $DISP_E_BADVARTYPE = 0x80020008
	If Not IsDeclared("DISP_E_EXCEPTION") Then Global Const $DISP_E_EXCEPTION = 0x80020009
	If Not IsDeclared("DISP_E_OVERFLOW") Then Global Const $DISP_E_OVERFLOW = 0x8002000A
	If Not IsDeclared("DISP_E_BADINDEX") Then Global Const $DISP_E_BADINDEX = 0x8002000B
	If Not IsDeclared("DISP_E_UNKNOWNLCID") Then Global Const $DISP_E_UNKNOWNLCID = 0x8002000C
	If Not IsDeclared("DISP_E_ARRAYISLOCKED") Then Global Const $DISP_E_ARRAYISLOCKED = 0x8002000D
	If Not IsDeclared("DISP_E_BADPARAMCOUNT") Then Global Const $DISP_E_BADPARAMCOUNT = 0x8002000E
	If Not IsDeclared("DISP_E_PARAMNOTOPTIONAL") Then Global Const $DISP_E_PARAMNOTOPTIONAL = 0x8002000F
	If Not IsDeclared("DISP_E_BADCALLEE") Then Global Const $DISP_E_BADCALLEE = 0x80020010
	If Not IsDeclared("DISP_E_NOTACOLLECTION") Then Global Const $DISP_E_NOTACOLLECTION = 0x80020011
	Global Const $tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2"
	Global Const $tagDISPPARAMS = "ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;"
	Global Const $__AOI_LOCK_CREATE = 1
	Global Const $__AOI_LOCK_UPDATE = 2
	Global Const $__AOI_LOCK_DELETE = 4
	Global Const $__AOI_LOCK_CASE = 8
	Global Enum $VT_EMPTY,$VT_NULL,$VT_I2,$VT_I4,$VT_R4,$VT_R8,$VT_CY,$VT_DATE,$VT_BSTR,$VT_DISPATCH, _
		$VT_ERROR,$VT_BOOL,$VT_VARIANT,$VT_UNKNOWN,$VT_DECIMAL,$VT_I1=16,$VT_UI1,$VT_UI2,$VT_UI4,$VT_I8, _
		$VT_UI8,$VT_INT,$VT_UINT,$VT_VOID,$VT_HRESULT,$VT_PTR,$VT_SAFEARRAY,$VT_CARRAY,$VT_USERDEFINED, _
		$VT_LPSTR,$VT_LPWSTR,$VT_RECORD=36,$VT_FILETIME=64,$VT_BLOB,$VT_STREAM,$VT_STORAGE,$VT_STREAMED_OBJECT, _
		$VT_STORED_OBJECT,$VT_BLOB_OBJECT,$VT_CF,$VT_CLSID,$VT_BSTR_BLOB=0xfff,$VT_VECTOR=0x1000, _
		$VT_ARRAY=0x2000,$VT_BYREF=0x4000,$VT_RESERVED=0x8000,$VT_ILLEGAL=0xffff,$VT_ILLEGALMASKED=0xfff, _
		$VT_TYPEMASK=0xfff
	Global Const $tagProperty = "ptr Name;ptr Variant;ptr __getter;ptr __setter;ptr Next"



#Region Logging anythings
	


	Func Logging($log, $line = '')
		;Logging2($log, $line)
		$fn = DllCallbackRegister('Logging2', 'none', 'wstr;str');
		$lpfn = DllCallbackGetPtr($fn);
		DllCallAddress ('none', $lpfn, 'wstr', $log, 'str', $line);
		DllCallbackFree($fn)
	EndFunc
	
	Func _InitLog($name = '')
		Global $Global_Timer = TimerInit()
		OnAutoItExitRegister("OnExitAutoITGlobal")
		$IsLog = True	
		$log_prefix = '[Nolog on file]'		
		$File_Log_Path = @ScriptDir & '\log\' & $name & "_" & @YEAR & "_" & @MON & "_" & @MDAY & $File_Log_Extension
		If $IsLog Then
			If Not FileExists($File_Log_Path) Then
				If Not FileCreateNew($File_Log_Path) Then
					ConsoleWrite(" Error Creating/Resetting log.      error:" & @error & @CRLF)
					$IsLog = False
				EndIf
			EndIf			
		EndIf
		
		If $IsLog Then Logging("############### Program is starting ################")
	EndFunc
	
	Func OnExitAutoITGlobal()
		Global $g_hTimer, $g_iSecs, $g_iMins, $g_iHour, $g_sTime
		__TicksToTime(Round(TimerDiff($Global_Timer)), $g_iHour, $g_iMins, $g_iSecs)
		 Local $sTime = $g_sTime ; save current time to be able to test and avoid flicker..
		$g_sTime = StringFormat("%02i:%02i:%02i", $g_iHour, $g_iMins, $g_iSecs)
		If $sTime <> $g_sTime Then Logging("@@@@@@@@@@@@@@@@@@@@@@@@@@ Program is Exited After " & $g_sTime & " @@@@@@@@@@@@@@@@@@@@@@@@@@")
	EndFunc
	
	
	Func __TicksToTime($iTicks, ByRef $iHours, ByRef $iMins, ByRef $iSecs)
		If Number($iTicks) > 0 Then
			$iTicks = Int($iTicks / 1000)
			$iHours = Int($iTicks / 3600)
			$iTicks = Mod($iTicks, 3600)
			$iMins = Int($iTicks / 60)
			$iSecs = Mod($iTicks, 60)
			; If $iHours = 0 then $iHours = 24
			Return 1
		ElseIf Number($iTicks) = 0 Then
			$iHours = 0
			$iTicks = 0
			$iMins = 0
			$iSecs = 0
			Return 1
		Else
			Return SetError(1, 0, 0)
		EndIf
	EndFunc   ;==>_TicksToTime
	
	Func Logging2($log, $LineNumber)
		If $LineNumber <> '' and StringInStr(@ScriptFullPath, ".au3")Then
			$LineText = '[' & $LineNumber & ']'
		Else 
			$LineText = ''
		EndIf
		Local $current_time = '[' & @YEAR & "/" & @MON & "/" & @MDAY &  " " & _DateTimeFormat(_NowCalc(), 5) & ']' & $LineText
		Local $log_content = $current_time & $log & @CRLF
		If $IsLog Then 
			ConsoleWrite($log_content)
			$fopen = FileOpen($File_Log_Path,1 + 128)
			;MsgBox("", "", $fopen)
			If Not FileWrite($fopen, $log_content)	Then
				ConsoleWrite(" Error Write log.      error:" & @error & @CRLF)
				$IsLog = False
			EndIf
			FileClose($fopen)
		Else
			ConsoleWrite($log_prefix & $log_content)
		EndIf
		Return 1
		;Return _WinAPI_CallWindowProc($log, $LineNumber)
	EndFunc
	
	
	Func FileCreateNew($sFilePath,  $noidung = "")
		Local $hFileOpen = FileOpen($sFilePath, 2 + 8 + 128)
		If $hFileOpen = -1 Then Return SetError(1, 0, 0)
		Local $iFileWrite = FileWrite($hFileOpen, $noidung)
		FileClose($hFileOpen)
		If Not $iFileWrite Then 
			Return SetError(2, 0, 0)
		Else
			Logging("[Info] file is created " & $sFilePath)
		EndIf
		Return 1
	EndFunc
	
	
	
#EndRegion Logging anythings


#Region Memory
	Func MemoryObject($pid)
		If ProcessExists($pid) Then
			$MemoryObject =  __createObject()
			$MemoryObject.pid = $pid
			$MemoryObject.MemID = 0
			$MemoryObject.Address = 0
			$MemoryObject.InheritHandle = 1
			$MemoryObject.DesiredAcces = '0x1F0FFF'
			$MemoryObject.RtRead = ''
			$MemoryObject.Type = 'dword' ;byte, float, dword, char[], wchar[] https://docs.microsoft.com/en-us/windows/win32/winprog/windows-data-types
			$MemoryObject.__defineGetter("Open", __MemoryOpen); _MemoryOpen($oThis.parent.pid, $oThis.parent.iv_DesiredAcces, $oThis.parent.iv_InheritHandle)
			$MemoryObject.__defineGetter("Read", __MemoryRead)
			$MemoryObject.__defineGetter("Write", __MemoryWrite)
			$MemoryObject.__defineGetter("Close", __MemoryClose)
			;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryCreate & $Lang_log.Log.Text.Space & $pid & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
			Logging($log_prefix_InfoMemory & "MemoryObject đã tạo thành công", @ScriptLineNumber)
			Return $MemoryObject
		Else
			Logging($log_prefix_ErrMemory & "MemoryObject tạo không thành công [ProcessID không tồn tại]", @ScriptLineNumber)
			;Logging($Lang_log.Log.Text.Prefix.Error.Memory & $Lang_log.Log.Text.MemoryCreate & $Lang_log.Log.Text.Space & $pid & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
			Return SetError(1, 1, 0)
		EndIf
	EndFunc
	
	
	Func __MemoryOpen($oThis)
;~		If ProcessExists($oThis.parent.pid) Then
			$temp = _MemoryOpen($oThis.parent.pid);, $oThis.parent.DesiredAcces, $oThis.parent.InheritHandle)
			;, $oThis.parent.DesiredAcces, $oThis.parent.InheritHandle))
		If Not @error Then
			$oThis.parent.MemID = $temp
			;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryOpen & $Lang_log.Log.Text.Space & $oThis.parent.pid & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
			Logging($log_prefix_InfoMemory & "MemoryOpen thành công ", @ScriptLineNumber)
		Else 
			Logging($log_prefix_ErrMemory & "MemoryOpen không thành công ", @ScriptLineNumber)
			Return SetError(1, 1, 0)
			;Logging($Lang_log.Log.Text.Prefix.Error.Memory & $Lang_log.Log.Text.MemoryOpen & $Lang_log.Log.Text.Space & $oThis.parent.pid & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Exists,   @ScriptLineNumber)
		EndIf	

		Return $oThis.parent
	EndFunc
	
		
	Func __MemoryRead($oThis)
		
		Switch $oThis.arguments.length		
			Case 0
				;$oThis.parent.Address = $oThis.parent.RtRead	
				;$oThis.parent.Type = 'dword'
				$temp = _MemoryRead($oThis.parent.Address, $oThis.parent.MemID, $oThis.parent.Type)
				If Not @error Then
					$oThis.parent.RtRead = $temp
					Logging($log_prefix_InfoMemory & "MemoryRead thành công " & $oThis.arguments.length	, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
				Else 
					Logging($log_prefix_ErrMemory & "MemoryRead không thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
					Return SetError(4, 1, 0)
				EndIf		
			Case 1
				$oThis.parent.Address = $oThis.arguments.values[0]
				__DynamicArguments($oThis)			
				;$oThis.parent.Type = 'dword'
				$temp = _MemoryRead($oThis.parent.Address, $oThis.parent.MemID, $oThis.parent.Type)
				;MsgBox("", "", $temp)
				If Not @error Then
					$oThis.parent.RtRead = $temp
					Logging($log_prefix_InfoMemory & "MemoryRead thành công " & $oThis.arguments.length	, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
				Else 
					Logging($log_prefix_ErrMemory & "MemoryRead không thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
					Return SetError(3, 1, 0)
				EndIf				
			Case 2
				$oThis.parent.Address = $oThis.arguments.values[0]		
				$oThis.parent.Type = $oThis.arguments.values[1]	
				__DynamicArguments($oThis)	
				$temp = _MemoryRead($oThis.parent.Address, $oThis.parent.MemID, $oThis.parent.Type)
				If Not @error Then
					$oThis.parent.RtRead = $temp
					Logging($log_prefix_InfoMemory & "MemoryRead thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
				Else 
					Logging($log_prefix_ErrMemory & "MemoryRead không thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
					Return SetError(2, 1, 0)
				EndIf
			Case Else 
				Logging($log_prefix_ErrMemory & "MemoryRead không thành công [Agruments không đúng số lượng] " & $oThis.arguments.length, @ScriptLineNumber)
				Return SetError(1, 1, 0)

		EndSwitch
		Return $oThis.parent
	EndFunc
	
	
	Func __MemoryWrite($oThis)
		Switch $oThis.arguments.length		
			Case 1
				;$oThis.parent.Address = $oThis.parent.RtRead
				$oThis.parent.valW	 = $oThis.arguments.values[0]
				__DynamicArguments($oThis)			
				;$oThis.parent.Type = 'dword'
				$temp = _MemoryWrite($oThis.parent.Address, $oThis.parent.MemID,$oThis.parent.valW, $oThis.parent.Type)
				If Not @error Then
					$oThis.parent.RtWrite = $temp
					Logging($log_prefix_InfoMemory & "MemoryWrite thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
				Else 					
					Logging($log_prefix_ErrMemory & "MemoryRead không thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
					Return SetError(1, 1, 0)
				EndIf	
			Case 2
				$oThis.parent.Address = $oThis.arguments.values[0]
				$oThis.parent.valW	 = $oThis.arguments.values[1]
				;$oThis.parent.Type = 'dword'
				__DynamicArguments($oThis)			
				
				$temp = _MemoryWrite($oThis.parent.Address, $oThis.parent.MemID,$oThis.parent.valW, $oThis.parent.Type)
				If Not @error Then
					$oThis.parent.RtWrite = $temp
					Logging($log_prefix_InfoMemory & "MemoryWrite thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
				Else 
					Logging($log_prefix_ErrMemory & "MemoryRead không thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
					Return SetError(2, 1, 0)
				EndIf
			Case 3
				$oThis.parent.Address = $oThis.arguments.values[0]
				$oThis.parent.valW	 = $oThis.arguments.values[1]
				$oThis.parent.Type = $oThis.arguments.values[2]
				__DynamicArguments($oThis)
				$temp = _MemoryWrite($oThis.parent.Address, $oThis.parent.MemID,$oThis.parent.valW, $oThis.parent.Type)
				If Not @error Then
					$oThis.parent.RtWrite = $temp
					Logging($log_prefix_InfoMemory & "MemoryWrite thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
				Else 
					Logging($log_prefix_ErrMemory & "MemoryRead không thành công " & $oThis.arguments.length, @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
					Return SetError(3, 1, 0)
				EndIf
			Case Else
				Logging($log_prefix_ErrMemory & "MemoryWrite không thành công [Agruments không đúng số lượng] " & $oThis.arguments.length, @ScriptLineNumber)
				;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & $Lang_log.Log.Text.Space & $oThis.parent.Address & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Space & @error,   @ScriptLineNumber)
				Return SetError(4, 1, 0)
		EndSwitch 
		
		Return $oThis.parent
	EndFunc
	
	
	Func __MemoryClose($oThis)
		_MemoryClose($oThis.parent.MemID)
		;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.Space & $oThis.parent.MemID & $Lang_log.Log.Text.Close & @error,   @ScriptLineNumber)
	EndFunc
	
	
	Func __DynamicArguments($oThis)
		For $i = 0 To $oThis.arguments.length - 1
			$newstr = StringReplace(StringStripCR(StringStripWS($oThis.arguments.values[$i], 8)), ' ', '')
			$vari 	= StringSplit($newstr, '=')
			If StringInStr($newstr, '=') Then	
				$oThis.parent.__set($vari[1], $vari[2])
				Logging($log_prefix_info & "__DynamicArguments thành công " & $vari[1] & ":" & $vari[2], @ScriptLineNumber)
				;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.Vari & $Lang_log.Log.Text.Space & $vari[1] & $Lang_log.Log.Text.Set & $Lang_log.Log.Text.Space & $vari[2],   @ScriptLineNumber)
			EndIf
		Next
	EndFunc
#EndRegion Memory	


#Region Threading
	

	Func ThreadCreate($function)
		If VarGetType(Execute($function)) = 'UserFunction' Then
			$ThreadObject =  __createObject()
			$ThreadObject.Thread = $function
			$ThreadObject.ppid = @AutoItPID
			$ThreadObject.__defineGetter("Start",ThreadStart)
			$ThreadObject.__defineGetter("Send",ThreadCBSend)
			$ThreadObject.__defineGetter("Suspend",ThreadSuspend)
			$ThreadObject.__defineGetter("Pause",ThreadSuspend)
			$ThreadObject.__defineGetter("Resume",ThreadResume)
			$ThreadObject.__defineGetter("Abort",ThreadAbort)
			$ThreadObject.__defineGetter("MainCallBack",ThreadCallBack)
			$ThreadObject.pid = 0		
			;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.ThreadCreate & $Lang_log.Log.Text.Space & $function & $Lang_log.Log.Text.Success ,   @ScriptLineNumber)
			Logging($log_prefix_InfoThread & "ThreadObject tạo thành công")
			Return $ThreadObject
		Else
			Logging($log_prefix_InfoThread & "ThreadObject tạo không thành công [Không tồn tại Function " & $function & "]")
			;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Space & $function & $Lang_log.Log.Text.False & $Lang_log.Log.Text.Func,   @ScriptLineNumber) 
			;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.ThreadCreate & $Lang_log.Log.Text.Space & $function & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success,   @ScriptLineNumber)
			Return SetError(1, 1, 0)
		EndIf
	EndFunc


	Func ThreadAbort($oThis)
		If ProcessExists($oThis.parent.pid) Then	
			For $i = 1 to UBound($ThreadArrayAll)
				;Logging("[Info][Thread] $ThreadArrayAll thứ " & $i & " | Check with pid " & $oThis.parent.pid,  @ScriptLineNumber)
				If $ThreadArrayAll[$i] = $oThis.parent.pid Then 
					If not @error Then
						_ArrayDelete($ThreadArrayAll, $i)
						If Not @error Then
							ProcessClose($oThis.parent.pid)
							;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Space & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space &  $oThis.parent.pid & $Lang_log.Log.Text.Info.Aborted,  @ScriptLineNumber)
							Logging($log_prefix_InfoThread & " ThreadAbort đã xóa pid " & $oThis.parent.pid, @ScriptLineNumber)
							ExitLoop
						Else 
							;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Space & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space &  $oThis.parent.pid & $Lang_log.Log.Text.Error.DelFArray,  @ScriptLineNumber)
							Logging($log_prefix_ErrThread & "ThreadAbort không tồn tại trong array pid " & $oThis.parent.pid, @ScriptLineNumber)		
						EndIf
					EndIf
				EndIf
			Next			
		Else
			Logging($log_prefix_ErrThread & "ThreadAbort Không tìm thấy pid " & $oThis.parent.pid, @ScriptLineNumber)
			;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Error.NotFound & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $oThis.parent.pid,  @ScriptLineNumber)
		EndIf
		Return $oThis.parent
	EndFunc


	Func ThreadCallBack($oThis)
		If ($oThis.arguments.length == 0) Then
			$oThis.parent.pid = _CoProc_Create($oThis.parent.Thread)	
			Logging($log_prefix_ErrThread & "ThreadCallBack Không có func callback", @ScriptLineNumber)
		ElseIF ($oThis.arguments.length == 1) Then
			SetCallBack($oThis.arguments.values[0])
		EndIf	
		Return $oThis.parent
	EndFunc
	
	
	Func SetCallBack($function)
		If VarGetType(Execute($function)) = 'UserFunction' Then
			_CoProc_Reciver($function)
			Logging($log_prefix_InfoThread & "SetCallBack setCallBack " & $function & " cho pid " & @AutoItPID, @ScriptLineNumber)
		Else 
			Logging($log_prefix_ErrThread & "SetCallBack Không tìm thấy Function " & $function, @ScriptLineNumber)
		EndIf
	EndFunc 
		
		
	Func CBMsgToMain($msg)
		Return _CoProc_Send($gi_CoProcParent,$msg)		
	EndFunc
	

	Func ThreadResume($oThis)
		If ProcessExists($oThis.parent.pid) Then
			_CoProc_Resume($oThis.parent.pid)
			;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $oThis.parent.pid & $Lang_log.Log.Text.Resumed,  @ScriptLineNumber)
			Logging($log_prefix_ErrThread & "ThreadResume thành công pid " & $oThis.parent.pid, @ScriptLineNumber)
		Else			
			;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Error.NotFound & $Lang_log.Log.Text.Pid & $oThis.parent.pid,  @ScriptLineNumber)
			Logging($log_prefix_ErrThread & "ThreadResume Không tìm thấy pid " & $oThis.parent.pid, @ScriptLineNumber)
		EndIf
		Return $oThis.parent
	EndFunc
	
	
	Func ThreadCBSend($oThis)
		If ($oThis.arguments.length == 0) Then	
			Logging($log_prefix_ErrThread & "ThreadSend Không tìm thấy dữ liệu gửi", @ScriptLineNumber)
		ElseIF ($oThis.arguments.length == 1) Then
			_CoProc_Send($oThis.parent.pid,$oThis.arguments.values[0])
		EndIf			
		Return $oThis.parent
	EndFunc
	

	Func ThreadSuspend($oThis)
		If ProcessExists($oThis.parent.pid) Then
			If ($oThis.arguments.length == 0) Then
				_CoProc_Suspend($oThis.parent.pid)
				Logging($log_prefix_ErrThread & "ThreadSuspend thành công với Arg = 0", @ScriptLineNumber)
				;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $oThis.parent.pid & $Lang_log.Log.Text.Suppended & "Arg = 0",  @ScriptLineNumber)
			ElseIf ($oThis.arguments.length == 1) Then
				_CoProc_Suspend($oThis.parent.pid, $oThis.arguments.values[0])
				Logging($log_prefix_ErrThread & "ThreadSuspend thành công với Arg = " & $oThis.arguments.values[0], @ScriptLineNumber)
				;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $oThis.parent.pid & $Lang_log.Log.Text.Suppended & $oThis.arguments.values[0],  @ScriptLineNumber)
			EndIf
		Else
			Logging($log_prefix_ErrThread & "ThreadSuspend Không tìm thấy pid " & $oThis.parent.pid, @ScriptLineNumber)
			;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Error.NotFound & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $oThis.parent.pid,  @ScriptLineNumber)
		EndIf
		Return $oThis.parent
	EndFunc


	Func ThreadStart($oThis) 
		
		Switch $oThis.arguments.length
			Case 0
				$oThis.parent.pid = _CoProc_Create($oThis.parent.Thread)	
				$sleeptime = 400
			Case 1
				$oThis.parent.pid = _CoProc_Create($oThis.parent.Thread,$oThis.arguments.values[0] )			
				$sleeptime = $oThis.arguments.values[1]
		EndSwitch
		
		Sleep($sleeptime)
		
		If ProcessExists($oThis.parent.pid) Then
			_ArrayAdd($ThreadArrayAll, $oThis.parent.pid)
			If Not @error Then
				If OnAutoItExitRegister("ThreadKillSelfIfNotHaveParent") Then
					Logging($log_prefix_InfoThread & "ThreadStart thành công Register Exit Function", @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Register & $Lang_log.Log.Text.Func & $Lang_log.Log.Text.OnExit & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & GetAllThreadrunning(),  @ScriptLineNumber)
				Else 
					Logging($log_prefix_ErrThread & "ThreadStart không thành công Register Exit Function", @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Register & $Lang_log.Log.Text.Func & $Lang_log.Log.Text.OnExit & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & GetAllThreadrunning(),  @ScriptLineNumber)
				EndIf
				Logging($log_prefix_InfoThread & "ThreadStart Add thành công vào Array với pid " & $oThis.parent.pid, @ScriptLineNumber)
				;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Run & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $oThis.parent.pid & $Lang_log.Log.Text.Success,  @ScriptLineNumber)
			Else 
				Logging($log_prefix_ErrThread & "ThreadStart Add không thành công vào Array với pid " & $oThis.parent.pid, @ScriptLineNumber)
				;Logging(Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Add & $Lang_log.Log.Text.Array & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success,  @ScriptLineNumber)
			EndIf
		Else
			Logging($log_prefix_ErrThread & "ThreadStart Không tìm thấy pid " & $oThis.parent.pid, @ScriptLineNumber)			
			;Logging($Lang_log.Log.Text.Prefix.Error.Thread & $Lang_log.Log.Text.Run & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Not & $Lang_log.Log.Text.Success,  @ScriptLineNumber)
		EndIf
		
		Return $oThis.parent		
	EndFunc
	
	
	Func GetAllThreadrunning()
		Local $AllPidThread
		If IsArray($ThreadArrayAll) Then 
			For $i = 1 to UBound($ThreadArrayAll) - 1
				$AllPidThread = $AllPidThread & $i & '-' & $ThreadArrayAll[$i] & ' '
			Next
			Return SetError(0, 0, $AllPidThread)
		EndIf
		Return SetError(1, 1, 0)
	EndFunc
	
	
	Func ThreadKillSelfIfNotHaveParent()
		For $thread in $ThreadArrayAll
			If $thread <> 0 Then				
				ProcessClose($thread)
				If Not @error Then 
					Logging($log_prefix_InfoThread & "ThreadKillSelfIfNotHaveParent Main bị closed nên đóng theo thread" & $thread ,  @ScriptLineNumber)
					;Logging($Lang_log.Log.Text.Prefix.Info.Thread & $Lang_log.Log.Text.Main & $Lang_log.Log.Text.Close & $Lang_log.Log.Text.Close & $Lang_log.Log.Text.Pid & $Lang_log.Log.Text.Space & $thread,  @ScriptLineNumber)
				EndIf
			EndIf
		Next
	EndFunc
	

#EndRegion Threading

#Region Objectvari
	Func _InitObjectVari( $filename = "variables.properties" )
		$arrayline = FileReadToArray($filename)
		If Not @error Then
			$ObjectDynamicVari = __createObject()
			For $i = 0 to UBound($arrayline)-1
				;ConsoleWrite($arrayline[$i])
				$perline = $arrayline[$i]
				$perline = StringStripWS($perline, 3)
				if StringRegExp(StringLeft($perline,1),"[A-Za-z0-9_]" ) = 1 Then 
					$ObjectDynamicVari = AssignValue($perline, $ObjectDynamicVari)	
				EndIf    
			Next
			Logging("[Info][ObjectV] Đã load thành công ObjectVariable từ file " & $filename)
			return $ObjectDynamicVari
		ElseIf @error = 1 Then 
			Logging("[Error][ObjectV] Không tìm thấy file")
		Else
			Logging("[Error][ObjectV] File rỗng chưa khai báo hoặc không đúng định dạng")
		EndIf
		
	EndFunc


	Func Gui_object($oSelf)
		If Not IsObj($oSelf.val) Then
			Local $IDispatch = __createObject()
			$oSelf.val = $IDispatch
			Return $IDispatch
		EndIf	
		Return $oSelf.val
	EndFunc
	


	Func AssignValue($text,$object = "", $delimeter = "=")
		$b 		= StringSplit($text, $delimeter)
		$vari 	= StringStripCR(StringStripWS($b[1], 8))
		If StringRegExp($vari,"^[A-Za-z0-9_.]*$" ) Then
			$checkobject = StringInStr($vari, ".")
			If $checkobject <= 0 Then
				$pos 	= StringInStr($text, $delimeter)
				$value 	= StringRight($text, StringLen($text) -$pos)
				$object.__set($vari, $value)
			Else 		
				$object.__defineGetter(StringLeft($vari, $checkobject -1), Gui_object)
				AssignValue(StringTrimLeft($text, $checkobject), Execute("$object." & StringLeft($vari, $checkobject -1)))
			EndIf
			Return SetError(0, 0,$object)
		Else
			Return SetError(1, 1, 0)
		EndIf		
	EndFunc
#EndRegion ObjectVari


#Region BMPSearch
	
	#Region BMPSearch Object
	Func BMPSearch($title)
			$BmpSearchObject =  __createObject()
		If WinExists($title) Then 
			$BmpSearchObject.Title = $title
			$BmpSearchObject.__defineGetter("Search",CallBmpSearch)
			$BmpSearchObject.__defineGetter("Click",CallBmpSearchClick)
			$BmpSearchObject.Status =  False 
			$BmpSearchObject.x = 0
			$BmpSearchObject.y = 0
			return $BmpSearchObject
		Else 
			;MsgBox("", "ERROR", "Check lại title")
			Return SetError(1, 1, 0)
		EndIf 
	EndFunc 
	
	
	Func CallBmpSearchClick($oThis)		
		If Not WinExists($oThis.parent.Title) Then SetError(1, 1, 0)
		If $oThis.parent.Status Then
			
			If Not ($oThis.arguments.length == 1) Then
				_MouseClickPlus(WinGetHandle($oThis.parent.Title),"left",$oThis.parent.x,$oThis.parent.y,1, 1, "SendMessage")
			Else 
				_MouseClickPlus(WinGetHandle($oThis.parent.Title),"left",$oThis.parent.x,$oThis.parent.y,1, $oThis.arguments.values[0], "SendMessage")
			EndIf
		Else
			
		EndIf
		Return $oThis.parent
	EndFunc
	
	Func CallBmpSearch($oThis)
		If Not WinExists($oThis.parent.Title) Then SetError(1, 1, 0)
		If Not ($oThis.arguments.length == 1) Then
			SetError(1, 1, 0)
			$oThis.parent.Status = False
		Else
			Capture_Window(WinGetHandle($oThis.parent.Title))
			Local $temp = 0
			$temp = _BmpSearch('capture_window_objectvaridynamic.bmp', $oThis.arguments.values[0])
			If Not @error Then
				$oThis.parent.Status = True
				$wincord = WinGetPos($oThis.parent.Title)
				$oThis.parent.x = $temp[1][2] + 8;+ $wincord[0] - @DesktopWidth + 8;$temp[1][2]
				$oThis.parent.y = $temp[1][3] + 8;- $wincord[1] + @DesktopHeight + 8;$temp[1][3]
				
			Else			
				$oThis.parent.Status = False
				;Return SetError(1, 1, 0)
			EndIf
		EndIf
		Return $oThis.parent
	EndFunc
	#EndRegion BMPSearch Object
	
	
	#Region BMPSearch Include
		#Region Run Func
		   Func Bmp_Adv($bmps,$xs_x=0,$xs_y = 0)
				 Dim $timer = TimerInit()
				 Capture_Window($handle)

			  If IsArray($bmps) Then
				 For $bmp In $bmps
					;MsgBox("","",$bmp)
					BMP_Click($bmp,$xs_x,$xs_y)
				 Next
			  Else
				 If Not BMP_Click($bmps,$xs_x,$xs_y) Then BMP_Click($bmps,$xs_x,$xs_y)

			  EndIf
		   EndFunc

		   Func BMP_Click($bmp,$xs_x = 0,$xs_y = 0)
		;MsgBox("","",$bmp)
			  sleep(10)
			  $a = 0
			  $a = _BmpSearch("capture_window_objectvaridynamic.bmp",$bmp)
			  If Not @error Then
				 _MouseClickPlus($handle,"left",$a[1][2] + $wincord[2] - @DesktopWidth + $xs_x,$a[1][3]  - $wincord[3] + @DesktopHeight+ $xs_y,1,1, "SendMessage")
				 sleep(50)
				 ;_MouseClickPlus($handle,"left",100,100,1,"SendMessage")
				 ;ConsoleWrite($a[1][2] + $wincord[2] - @DesktopWidth + $xs_x & "  " &$a[1][3]  - $wincord[3] + @DesktopHeight+ $xs_y & @CRLF)
				 Return True
			  Else
				 Return False
			  EndIf

		   EndFunc

		   Func Get_Bmp($dir_bmp)
			  Dim $array = []
			  _ArrayDelete($array,0)
			  $hSearch = FileFindFirstFile(@ScriptDir&"/"&$dir_bmp&"/*.bmp")

			  While 1
				 $sFileName = FileFindNextFile($hSearch)
				 if @error then ExitLoop
				 ;ConsoleWrite($sFileName)
				 _ArrayAdd($array,@ScriptDir&"\"&$dir_bmp&"\"&$sFileName)
			  WEnd

			  Return $array
		   EndFunc
		#EndRegion Run Func




		#Region Include
		Func Capture_Window($hWnd)
			_GDIPlus_Startup()
			$w = _WinAPI_GetWindowWidth($hWnd)
			$h = _WinAPI_GetWindowHeight($hWnd)
			If Not IsHWnd($hWnd) Then Return SetError(1, 0, 0)
			If Int($w) < 1 Then Return SetError(2, 0, 0)
			If Int($h) < 1 Then Return SetError(3, 0, 0)
			Local Const $hDC_Capture = _WinAPI_GetDC(HWnd($hWnd))
			Local Const $hMemDC = _WinAPI_CreateCompatibleDC($hDC_Capture)
			Local Const $hHBitmap = _WinAPI_CreateCompatibleBitmap($hDC_Capture, $w, $h)
			Local Const $hObjectOld = _WinAPI_SelectObject($hMemDC, $hHBitmap)
			;DllCall("gdi32.dll", "int", "SetStretchBltMode", "hwnd", $hDC_Capture, "uint", 0)
			$user32_dll = DllOpen("user32.dll")
			DllCall($user32_dll, "int", "PrintWindow", "hwnd", $hWnd, "handle", $hMemDC, "int", 3)
			_WinAPI_DeleteDC($hMemDC)
			_WinAPI_SelectObject($hMemDC, $hObjectOld)
			_WinAPI_ReleaseDC($hWnd, $hDC_Capture)
			Local Const $hFullScreen = WinGetHandle("[TITLE:Program Manager;CLASS:Progman]")
			Local Const $aFullScreen = WinGetPos($hFullScreen)
			Local Const $c1 = $aFullScreen[2] - @DesktopWidth, $c2 = $aFullScreen[3] - @DesktopHeight
			Local Const $wc1 = $w - $c1, $hc2 = $h - $c2
			;If (($wc1 > 1 And $wc1 < $w) Or ($w - @DesktopWidth > 1) Or ($hc2 > 7 And $hc2 < $h) Or ($h - @DesktopHeight > 1)) And (BitAND(WinGetState(HWnd($hWnd)), 32) = 32) Then
			Local $hBmp_t = _GDIPlus_BitmapCreateFromHBITMAP($hHBitmap)
			$hBmp = _GDIPlus_BitmapCloneArea($hBmp_t, 8, 8, $w - 16, $h - 16)
			_GDIPlus_BitmapDispose($hBmp_t)
			;Else
			; $hBmp = _GDIPlus_BitmapCreateFromHBITMAP($hHBitmap)
			;EndIf
			DllClose($user32_dll)
			_WinAPI_DeleteObject($hHBitmap)
			_GDIPlus_ImageSaveToFile($hBmp, @ScriptDir & "\capture_window_objectvaridynamic.bmp")
			sleep(50)
			_GDIPlus_BitmapDispose($hBmp)
			_GDIPlus_Shutdown()
			EndFunc ;==>Capture_Window
		
		
		Func ConvertToBMP($File)
				_GDIPlus_Startup()
			$in=FileOpen($File,16)
			$in = FileRead($in)
			;MsgBox("","",$in)
			;ConsoleWrite($in & @CR)
			$Bitmap = _GDIPlus_BitmapCreateFromMemory (Binary($in))
			;ConsoleWrite($Bitmap & @CR)
			;$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($Bitmap)
			;ConsoleWrite($hBmp & @CR)
			_GDIPlus_ImageSaveToFile($Bitmap, $File & ".bmp")
				_GDIPlus_Shutdown()
				ShellExecute($File & ".bmp")
			return  $File & ".bmp"

		EndFunc

		; #FUNCTION# ====================================================================================================================
		; Name ..........: _BmpSearch
		; Description ...: Searches for Bitmap in a Bitmap
		; Syntax ........: _BmpSearch($hSource, $hFind, $iMax=5000)
		; Parameters ....: $hSource             - Handle to bitmap to search
		;                  $hFind               - Handle to bitmap to find
		;                  $iMax               	- Max matches to find
		; Return values .: Success: Returns a 2d array with the following format:
		;							$aCords[0][0] = Total Matches found
		;							$aCords[$i][0] = Width of bitmap
		;							$aCords[$i][1] = Hight of bitmap
		;							$aCords[$i][2] = X cordinate
		;							$aCords[$i][3] = Y cordinate
		;
		;					Failure: Returns 0 and sets @error to 1
		;
		; Author ........: Brian J Christy (Beege)
		; ===============================================================================================================================
		Func _BmpSearch($hSource, $hFind, $iMax = 5000)
			;ConsoleWrite($hSource &" " & $hFind & @CR)
			$hSource = _WinAPI_LoadImage(0, $hSource, $IMAGE_BITMAP, 0, 0, $LR_LOADFROMFILE)
			;MsgBox("","",$hSource)
			$hFind = _WinAPI_LoadImage(0, $hFind, $IMAGE_BITMAP, 0, 0, $LR_LOADFROMFILE)
			Static Local $aMemBuff, $tMem = 0, $fStartup = True
			;MsgBox("","",$IMAGE_BITMAP)
			If $fStartup Then
				;####### (BinaryStrLen = 490) #### (Base64StrLen = 328 )####################################################################################################
				Local $Opcode = 'yBAAAFCNRfyJRfSNRfiJRfBYx0X8AAAAAItVDP8yj0X4i10Ii0UYKdiZuQQAAAD38YnBi0X4OQN0CoPDBOL36akAAACDfSgAdB1TA10oO10YD4OVAAAAi1UkORN1A1vrBluDwwTrvVOLVSyLRTADGjtdGHd3iwg5C3UhA1oEi0gEO10Yd2Y5C3USA1oIi0gIO10Yc1c5' & _
						'C3UDW+sGW4PDBOuCi1UUid6LfQyLTRCJ2AHIO0UYczfzp4P5AHcLSoP6AHQNA3Uc6+KDwwTpVP///4tFIIkYg0UgBIPDBP9F/ItVNDlV/HQG6Tj///9bi0X8ycIwAA=='
				 $struct_demo = DllStructCreate("byte[254]")
				Local $aDecode = DllCall("Crypt32.dll", "bool", "CryptStringToBinary", "str", $Opcode, "dword", 0, "dword", 1, "struct*",$struct_demo , "dword*", 254, "ptr", 0, "ptr", 0)
				If @error Or (Not $aDecode[0]) Then Return SetError(1, 0, 0)
				$Opcode = BinaryMid(DllStructGetData($aDecode[4], 1), 1, $aDecode[5])

				$aMemBuff = DllCall("kernel32.dll", "ptr", "VirtualAlloc", "ptr", 0, "ulong_ptr", BinaryLen($Opcode), "dword", 4096, "dword", 64)
				$tMem = DllStructCreate('byte[' & BinaryLen($Opcode) & ']', $aMemBuff[0])
				DllStructSetData($tMem, 1, $Opcode)
				;####################################################################################################################################################################################
				$fStartup = False
			EndIf

			Local $iTime = TimerInit()

			Local $tSizeSource = _WinAPI_GetBitmapDimension($hSource)
			Local $tSizeFind = _WinAPI_GetBitmapDimension($hFind)

			Local $iRowInc = ($tSizeSource.X - $tSizeFind.X) * 4

			Local $tSource = _GetBmpPixelStruct($hSource)
			Local $tFind = _GetBmpPixelStruct($hFind)

			Local $aFD = _FindFirstDiff($tFind)
			Local $iFirstDiffIdx = $aFD[0]
			Local $iFirstDiffPix = $aFD[1]

			Local $iFirst_Diff_Inc = _FirstDiffInc($iFirstDiffIdx, $tSizeFind.X, $tSizeSource.X)
			If $iFirst_Diff_Inc < 0 Then $iFirst_Diff_Inc = 0

			Local $tCornerPixs = _CornerPixs($tFind, $tSizeFind.X, $tSizeFind.Y)
			Local $tCornerInc = _CornerInc($tSizeFind.X, $tSizeFind.Y, $tSizeSource.X)

			Local $pStart = DllStructGetPtr($tSource)
			Local $iEndAddress = Int($pStart + DllStructGetSize($tSource))

			Local $tFound = DllStructCreate('dword[' & $iMax & ']')

			Local $ret = DllCallAddress('dword', DllStructGetPtr($tMem), 'struct*', $tSource, 'struct*', $tFind, _
					'dword', $tSizeFind.X, 'dword', $tSizeFind.Y, _
					'dword', $iEndAddress, 'dword', $iRowInc, 'struct*', $tFound, _
					'dword', $iFirstDiffPix, 'dword', $iFirst_Diff_Inc, _
					'struct*', $tCornerInc, 'struct*', $tCornerPixs, _
					'dword', $iMax)

			If Not $ret[0] Then
			   _WinAPI_DeleteObject($hSource)
			  _WinAPI_DeleteObject($hFind)
			   Return SetError(1, 0, 0)
			  EndIf


			Local $aCords = _GetCordsArray($ret[0], $tFound, $tSizeSource.X, $pStart, $tSizeFind.X, $tSizeFind.Y)
			_WinAPI_DeleteObject($hSource)
			  _WinAPI_DeleteObject($hFind)
			Return SetExtended(Int(TimerDiff($iTime) * 1000), $aCords)

		EndFunc   ;==>_BmpSearch_MC


		#Region Internal Functions
		;Returns a Dllstructure will all pixels
		Func _GetBmpPixelStruct($hBMP)

			Local $tSize = _WinAPI_GetBitmapDimension($hBMP)
			local $tBits = 0
			Local $tBits = DllStructCreate('dword[' & ($tSize.X * $tSize.Y) & ']')

			_WinAPI_GetBitmapBits($hBMP, DllStructGetSize($tBits), DllStructGetPtr($tBits))

			Return $tBits

		#Tidy_Off
		#cs

			This is how the dllstructure index numbers correspond to the pixel cordinates:

			An 5x5 dimension bmp:
				X0	X1	X2	X3	X4
			Y0 	1   2	3	4	5
			Y1	6	7	8	9	10
			Y2	11	12	13	14	15
			Y3	16	17	18	19	20
			Y4	21	22	23	24	25

			An 8x8 dimension bmp:
				X0	X1	X2	X3	X4	X5	X6	X7
			Y0	1	2	3	4	5	6	7	8
			Y1	9	10	11	12	13	14	15	16
			Y2	17	18	19	20	21	22	23	24
			Y3	25	26	27	28	29	30	31	32
			Y4	33	34	35	36	37	38	39	40
			Y5	41	42	43	44	45	46	47	48
			Y6	49	50	51	52	53	54	55	56
			Y7	57	58	59	60	61	62	63	64

		#ce
			 #Tidy_On

		EndFunc   ;==>_GetBmpPixelStruct

		;Find first pixel that is diffrent than ....the first pixel
		Func _FindFirstDiff($tPix)

			;####### (BinaryStrLen = 106) ########################################################################################################################
			Static Local $Opcode = '0xC80000008B5D0C8B1383C3048B4D103913750C83C304E2F7B800000000EB118B5508FF338F028B451029C883C002EB00C9C20C00'
			Static Local $aMemBuff = DllCall("kernel32.dll", "ptr", "VirtualAlloc", "ptr", 0, "ulong_ptr", BinaryLen($Opcode), "dword", 4096, "dword", 64)
			Static Local $tMem = 0
			$tMem = DllStructCreate('byte[' & BinaryLen($Opcode) & ']', $aMemBuff[0])
			Static Local $fSet = DllStructSetData($tMem, 1, $Opcode)
			;#####################################################################################################################################################

			Local $iMaxLoops = (DllStructGetSize($tPix) / 4) - 1
			Local $aRet = DllCallAddress('dword', DllStructGetPtr($tMem), 'dword*', 0, 'struct*', $tPix, 'dword', $iMaxLoops)

			Return $aRet

		EndFunc   ;==>_FindFirstDiff

		; Calculates the value to increase pointer by to check first different pixel
		Func _FirstDiffInc($iDx, $iFind_Xmax, $iSource_Xmax)

			Local $aFirstDiffCords = _IdxToCords($iDx, $iFind_Xmax)
			Local $iXDiff = ($iDx - ($aFirstDiffCords[1] * $iFind_Xmax)) - 1

			Return (($aFirstDiffCords[1] * $iSource_Xmax) + $iXDiff) * 4

		EndFunc   ;==>_FirstDiffInc

		;Converts the pointer addresses to cordinates
		Func _GetCordsArray($iTotalFound, $tFound, $iSource_Xmax, $pSource, $iFind_Xmax, $iFind_Ymax)

			Local $aRet[$iTotalFound + 1][4]
			$aRet[0][0] = $iTotalFound

			For $i = 1 To $iTotalFound
				$iFoundIndex = ((DllStructGetData($tFound, 1, $i) - $pSource) / 4) + 1
				$aRet[$i][0] = $iFind_Xmax
				$aRet[$i][1] = $iFind_Ymax
				$aRet[$i][3] = Int(($iFoundIndex - 1) / $iSource_Xmax) ; Y
				$aRet[$i][2] = ($iFoundIndex - 1) - ($aRet[$i][3] * $iSource_Xmax) ; X
			Next

			Return $aRet

		EndFunc   ;==>_GetCordsArray

		;converts cordinates to dllstructure index number
		Func _CordsToIdx($iX, $iY, $iMaxX)
			Return ($iY * $iMaxX) + $iX + 1
		EndFunc   ;==>_CordsToIdx

		;convert dllstructure index number to cordinates
		Func _IdxToCords($iDx, $iMaxX)

			Local $aCords[2]
			$aCords[1] = Int(($iDx - 1) / $iMaxX) ; Y
			$aCords[0] = ($iDx - 1) - ($aCords[1] * $iMaxX) ; X

			Return $aCords

		EndFunc   ;==>_IdxToCords

		;Retrieves the Pixel Values of Right Top, Left Bottom, Right Bottom. Returns dllstructure
		Func _CornerPixs(ByRef $tFind, $iFind_Xmax, $iFind_Ymax)

			Local $tCornerPixs = DllStructCreate('dword[3]')

			DllStructSetData($tCornerPixs, 1, DllStructGetData($tFind, 1, $iFind_Xmax), 1) ; top right corner
			DllStructSetData($tCornerPixs, 1, DllStructGetData($tFind, 1, ($iFind_Xmax + ($iFind_Xmax * ($iFind_Ymax - 2)) + 1)), 2) ;  bottom left corner
			DllStructSetData($tCornerPixs, 1, DllStructGetData($tFind, 1, ($iFind_Xmax * $iFind_Ymax)), 3);	bottom right corner

			Return $tCornerPixs

		EndFunc   ;==>_CornerPixs

		;Retrieves the pointer adjust values for Right Top, Left Bottom, Right Bottom. Returns dllstructure
		Func _CornerInc($iFind_Xmax, $iFind_Ymax, $iSource_Xmax)

			Local $tCornerInc = DllStructCreate('dword[3]')

			DllStructSetData($tCornerInc, 1, ($iFind_Xmax - 1) * 4, 1)
			DllStructSetData($tCornerInc, 1, (($iSource_Xmax - $iFind_Xmax) + $iSource_Xmax * ($iFind_Ymax - 2) + 1) * 4, 2)
			DllStructSetData($tCornerInc, 1, ($iFind_Xmax - 1) * 4, 3)

			Return $tCornerInc

		EndFunc   ;==>_CornerInc
		#EndRegion Internal Functions

		Func _MouseClickPlus($Window, $Button = "left", $X = "", $Y = "", $Clicks = 1,$sTime = 1, $type = "SendMessage")
			Local $MK_LBUTTON       =  0x0001
			Local $WM_LBUTTONDOWN   =  0x0201
			Local $WM_LBUTTONUP     =  0x0202

			Local $MK_RBUTTON       =  0x0002
			Local $WM_RBUTTONDOWN   =  0x0204
			Local $WM_RBUTTONUP     =  0x0205

			Local $WM_MOUSEMOVE     =  0x0200

			Local $i                = 0

			Select
			Case $Button = "left"
			   $Button     =  $MK_LBUTTON
			   $ButtonDown =  $WM_LBUTTONDOWN
			   $ButtonUp   =  $WM_LBUTTONUP
			Case $Button = "right"
			   $Button     =  $MK_RBUTTON
			   $ButtonDown =  $WM_RBUTTONDOWN
			   $ButtonUp   =  $WM_RBUTTONUP
			EndSelect

			If $X = "" OR $Y = "" Then
			   $MouseCoord = MouseGetPos()
			   $X = $MouseCoord[0]
			   $Y = $MouseCoord[1]
			EndIf
			
			For $i = 1 to $Clicks
		;~        DllCall("user32.dll", "int", $type, _
		;~           "hwnd",  WinGetHandle( $Window ), _
		;~           "int",   $WM_MOUSEMOVE, _
		;~           "int",   0, _
		;~           "long",  _MakeLong($X, $Y))

			   DllCall("user32.dll", "int", $type, _
				  "hwnd",   $Window , _
				  "int",   $ButtonDown, _
				  "int",   $Button, _
				  "long",  _MakeLong($X, $Y))
				Sleep($sTime)
			   DllCall("user32.dll", "int", $type, _
				  "hwnd",   $Window , _
				  "int",   $ButtonUp, _
				  "int",   $Button, _
				  "long",  _MakeLong($X, $Y))
			   DllCall("user32.dll", "int", $type, _
				  "hwnd",  $Window, _
				  "int",   $WM_MOUSEMOVE, _
				  "int",   0, _
				  "long",  _MakeLong(0, 0))
			Next
				
				
		
		 EndFunc




		 Func _MakeLong($LoWord,$HiWord)
			Return BitOR($HiWord * 0x10000, BitAND($LoWord, 0xFFFF))
		 EndFunc



		#cs Test/verify Functions
			Func _TestDiff()

			Local $hFind = _ScreenCapture_Capture('', 80, 250, 140, 280)
			Local $hSource = _ScreenCapture_Capture('', 80, 250, 240, 350)

			Local $tSizeSource = _WinAPI_GetBitmapDimension($hSource)
			Local $tSizeFind = _WinAPI_GetBitmapDimension($hFind)

			Local $iSource_Xmax = DllStructGetData($tSizeSource, 'X')
			Local $iFind_Xmax = DllStructGetData($tSizeFind, 'X')
			Local $iFind_Ymax = DllStructGetData($tSizeFind, 'Y')

			Local $tSource = _GetBmpPixelStruct($hSource)
			Local $tFind = _GetBmpPixelStruct($hFind)

			Local $aFD = _FindFirstDiff($tFind)
			Local $iFirstDiffIdx = $aFD[0]
			Local $iFirstDiffPix = $aFD[1]

			Local $iFirst_Diff_Inc = _FirstDiffInc($iFirstDiffIdx, $iFind_Xmax, $iSource_Xmax)

			$p = DllStructGetPtr($tSource)
			$temp = DllStructCreate('dword', $p + ($iFirst_Diff_Inc))
			$sourcepix = Hex(DllStructGetData($temp, 1), 8)

			_WinAPI_DeleteObject($hSource)
			_WinAPI_DeleteObject($hFind)

			EndFunc

			Func _TestCorners()

			Local $hFind = _ScreenCapture_Capture('', 1, 1, 100, 100)
			Local $hSource = _ScreenCapture_Capture('', 1, 1, 250, 350)

			Local $tSizeSource = _WinAPI_GetBitmapDimension($hSource)
			Local $tSizeFind = _WinAPI_GetBitmapDimension($hFind)

			Local $iSource_Xmax = DllStructGetData($tSizeSource, 'X')
			Local $iFind_Xmax = DllStructGetData($tSizeFind, 'X')
			Local $iFind_Ymax = DllStructGetData($tSizeFind, 'Y')

			Local $tSource = _GetBmpPixelStruct($hSource)
			Local $tFind = _GetBmpPixelStruct($hFind)

			Local $tCornerPixs = _CornerPixs($tFind, $iFind_Xmax, $iFind_Ymax)
			Local $tCornerInc = _CornerInc($iFind_Xmax, $iFind_Ymax, $iSource_Xmax)

			ConsoleWrite('TLC = ' & Hex(DllStructGetData($tFind, 1, 1), 8) & @LF)
			ConsoleWrite('TRC = ' & Hex(DllStructGetData($tCornerPixs, 1, 1), 8) & @LF)
			ConsoleWrite('BLC = ' & Hex(DllStructGetData($tCornerPixs, 1, 2), 8) & @LF)
			ConsoleWrite('BRC = ' & Hex(DllStructGetData($tCornerPixs, 1, 3), 8) & @LF)
			ConsoleWrite(@LF)

			$iFirstInc = DllStructGetData($tCornerInc, 1, 1)
			$iSecond = DllStructGetData($tCornerInc, 1, 2)
			$iThird = DllStructGetData($tCornerInc, 1, 3)

			$p = DllStructGetPtr($tSource)
			$tTLC = DllStructCreate('dword', $p)
			$tTRC = DllStructCreate('dword', $p + ($iFirstInc))
			$tBLC = DllStructCreate('dword', $p + ($iFirstInc+$iSecond))
			$tBRC = DllStructCreate('dword', $p + ($iFirstInc+$iSecond+$iThird))
			ConsoleWrite('TLC = ' & Hex(DllStructGetData($tTLC, 1), 8) & @LF)
			ConsoleWrite('TRC = ' & Hex(DllStructGetData($tTRC, 1), 8) & @LF)
			ConsoleWrite('BLC = ' & Hex(DllStructGetData($tBLC, 1), 8) & @LF)
			ConsoleWrite('BRC = ' & Hex(DllStructGetData($tBRC, 1), 8) & @LF)

			_WinAPI_DeleteObject($hSource)
			_WinAPI_DeleteObject($hFind)

			EndFunc
		#ce

		#cs Assembly Source

		Func _BmpSearch($hSource, $hFind, $iMax = 5000)

			Local $iTime = TimerInit()

			Local $tSizeSource = _WinAPI_GetBitmapDimension($hSource)
			Local $tSizeFind = _WinAPI_GetBitmapDimension($hFind)

			Local $iRowInc = ($tSizeSource.X - $tSizeFind.X) * 4

			Local $tSource = _GetBmpPixelStruct($hSource)
			Local $tFind = _GetBmpPixelStruct($hFind)

			Local $aFD = _FindFirstDiff($tFind)
			Local $iFirstDiffIdx = $aFD[0]
			Local $iFirstDiffPix = $aFD[1]

			Local $iFirst_Diff_Inc = _FirstDiffInc($iFirstDiffIdx, $tSizeFind.X, $tSizeSource.X)
			If $iFirst_Diff_Inc < 0 Then $iFirst_Diff_Inc = 0

			Local $tCornerPixs = _CornerPixs($tFind, $tSizeFind.X, $tSizeFind.Y)
			Local $tCornerInc = _CornerInc($tSizeFind.X, $tSizeFind.Y, $tSizeSource.X)

			Local $pStart = DllStructGetPtr($tSource)
			Local $iEndAddress = Int($pStart + DllStructGetSize($tSource))

			_FasmFunc("'struct*', $pSource, 'struct*', $pFind, " & _
					"'dword', $iFind_Xmax, 'dword', $iFind_Ymax, " & _
					"'dword', $iEndAddress, 'dword', $iRowInc, 'struct*', $pFound, " & _
					"'dword', $iFirstDiffPix, 'dword', $iFirst_Diff_Inc, " & _
					"'struct*', $pCornerInc, 'struct*', $pCornerPixs, " & _
					"'dword', $iMax")

			_FasmLocalVar('$iTotalFound = 0') ; create var to hold total found matches

			_FasmAdd("mov edx, $pFind") ; set edx to address of $tFind
			_FasmLocalVar('$iFirstPixel = dword[edx]'); first pixel to find

			_FasmAdd("mov ebx, $pSource") ; set ebx to address of $tSource. As we walk through the bitmap, ebx will hold are current postion throughout the whole function.

			; Start of main loop
			_FasmAdd("FindFistPixel:")
			_FasmAdd("		mov eax, $iEndAddress") ; set eax to $iEndaddress.
			_FasmAdd("		sub eax, ebx") ; subtract current address -
			_FasmAdd('		cdq') ; convert dword in eax to qword. the value is placed in edx:eax  - (needed for division below)
			_FasmAdd('		mov ecx, 4') ; ecx = 4 - ecx is the divisor for "div" instruction
			_FasmAdd('		div ecx') ; eax = (edx:eax)/ecx   (fyi - remainder is placed in edx if ever needed)
			_FasmAdd('		mov ecx, eax');  Set Max Loops --  ecx = ($iEndAddress-$iCurrentAddress) / 4
			_FasmAdd("		mov eax, $iFirstPixel")
			_FasmAdd("		CheckNextPixel:")
			_FasmJumpIf('		dword[ebx] = eax, CheckFirstDiff') ; Pixels match. Check if first diff pixels match.
			_FasmAdd("			add ebx, 4") ; increse current address to next pixel
			_FasmAdd("			loop CheckNextPixel") ; "loop" instruction will decrease ecx each loop and exit when ecx = 0
			_FasmAdd("jmp Finished")

			_FasmAdd("CheckFirstDiff:")
			_FasmJumpIf('	$iFirst_Diff_Inc = 0, CheckCorners') ; if inc is 0 then 'find' pic is all one color. must skip the diff check.
			_FasmAdd("		push ebx") ; save current are current address ( position)
			_FasmAdd("		add ebx, $iFirst_Diff_Inc") ; increase current address to first pixel that is different
			_FasmJumpIf("	ebx >= $iEndAddress, PopFinished")
			_FasmAdd("		mov edx, $iFirstDiffPix") ; set edx to pixel value to check
			_FasmJumpIf("	dword[ebx] <> edx, FirstDiffFail") ; compare pixels. jump to FirstDiffFail if not equal
			_FasmAdd("		pop ebx") ; restore address
			_FasmAdd("		jmp CheckCorners") ;

			_FasmAdd("FirstDiffFail:")
			_FasmAdd("		pop ebx") ; restore address
			_FasmAdd("		add ebx, 4") ; increase current address
			_FasmAdd("		jmp FindFistPixel") ; start over

			_FasmAdd("CheckCorners:")
			_FasmAdd("		push ebx")
			_FasmAdd("		mov edx, $pCornerInc") ; set edx to structure that holds are values to adjust pointer to corners
			_FasmAdd("		mov eax, $pCornerPixs") ; set ebx to struct that holds pixel values
			_FasmAdd("		add ebx, dword[edx]") ; increase address to top right corner
			_FasmJumpIf("	ebx > $iEndAddress, PopFinished")
			_FasmAdd("		mov ecx, dword[eax]") ; set ecx to pixel value
			_FasmJumpIf('	dword[ebx] <> ecx, CornerFail') ; compare
			_FasmAdd("		add ebx, dword[edx+4]") ; increase address to bottom left corner
			_FasmAdd("		mov ecx, dword[eax+4]") ; set ecx to bottom left pixel value
			_FasmJumpIf("	ebx > $iEndAddress, PopFinished")
			_FasmJumpIf('	dword[ebx] <> ecx, CornerFail') ; compare
			_FasmAdd("		add ebx, dword[edx+8]") ; increase address to bottom right corner
			_FasmAdd("		mov ecx, dword[eax+8]") ; set pixel value
			_FasmJumpIf("	ebx >= $iEndAddress, PopFinished")
			_FasmJumpIf('	dword[ebx] <> ecx, CornerFail') ; compare
			_FasmAdd("		pop ebx") ; restore orignal address
			_FasmAdd("		jmp CheckMatch") ; if we are here all corners matched. Now do a full check

			_FasmAdd("CornerFail:")
			_FasmAdd("		pop ebx") ; restore address
			_FasmAdd("		add ebx, 4") ; increase address to next pixel
			_FasmAdd("		jmp FindFistPixel") ; start over

			_FasmAdd("CheckMatch:")
			_FasmAdd("		mov edx, $iFind_Ymax") ; set edx to max number of rows to check
			_FasmAdd("		mov esi, ebx") ; set esi to current address. esi and edi must be used for the "repe cmpsd" command below
			_FasmAdd("		mov edi, $pFind") ;	set edi to address of find pic
			_FasmAdd("		NextRow:")
			_FasmAdd("			mov ecx, $iFind_Xmax") ; set max number of compares
			_FasmAdd("			mov eax, ebx") ; move current address to ebx
			_FasmAdd("			add eax, ecx") ; add max compares.
			_FasmJumpIf("	    eax >= $iEndAddress, Finished") ; verify we will not go past end address
			_FasmAdd("			repe cmpsd") ; 			repeats comparing bits in edi and esi until bits are not equal or ecx is 0
			_FasmJumpIf("		ecx > 0, NextPixel") ; if ecx is > 0 then we found a bits not equal, else the whole row matched
			_FasmAdd("			dec edx") ; dec row count
			_FasmJumpIf("		edx = 0, FoundMatch") ; if row count reaches 0 the we have a match. exit loop
			_FasmAdd("			add esi, $iRowInc") ; increase source address to next row
			_FasmAdd("			jmp NextRow") ; check next row

			_FasmAdd("NextPixel:")
			_FasmAdd("		add ebx, 4"); ; increase address
			_FasmAdd("		jmp FindFistPixel") ; start over

			_FasmAdd("FoundMatch:")
			_FasmAdd("		mov eax, $pFound") ; move pointer of $tFound to eax
			_FasmAdd("		mov dword[eax], ebx") ; set with address found
			_FasmAdd("		add $pFound, 4") ; increase address of $pfound.
			_FasmAdd("		add ebx, 4") ; increase current address
			_FasmAdd("		inc $iTotalFound") ; increase iTotalfound
			_FasmAdd("		mov edx, $iMax") ; set edx to max found
			_FasmJumpIf('	$iTotalFound = edx, Finished') ; verify we are not over are limit. if we are exit
			_FasmAdd("		jmp FindFistPixel") ; start over
			;End main loop

			_FasmAdd("PopFinished:")
			_FasmAdd("pop ebx")

			_FasmAdd("Finished:")
			_FasmAdd("mov eax, $iTotalFound")
			_FasmEndFunc()

			Return _FasmCompileMC('_BmpSearch')
			Local $pMem = _FasmGetFuncPtr()

			Local $tFound = DllStructCreate('dword[' & $iMax & ']')

			Local $ret = DllCallAddress('dword', $pMem, 'struct*', $tSource, 'struct*', $tFind, _
					'dword', $tSizeFind.X, 'dword', $tSizeFind.Y, _
					'dword', $iEndAddress, 'dword', $iRowInc, 'struct*', $tFound, _
					'dword', $iFirstDiffPix, 'dword', $iFirst_Diff_Inc, _
					'struct*', $tCornerInc, 'struct*', $tCornerPixs, _
					'dword', $iMax)

			If Not $ret[0] Then Return SetError(1, 0, 0)

			Local $aCords = _GetCordsArray($ret[0], $tFound, $tSizeSource.X, $pStart, $tSizeFind.X, $tSizeFind.Y)
			Return SetExtended(Int(TimerDiff($iTime) * 1000), $aCords)

		EndFunc   ;==>_BmpSearch


			Func __FindFirstDiff($tPix)
			_FasmFunc("'dword*', $iPixFound, 'struct*', $tPix, 'dword', $iMaxLoops")

			_FasmAdd("mov ebx, $tPix")
			_FasmAdd("mov edx, dword[ebx]")
			_FasmAdd("add ebx, 4")
			_FasmAdd("mov ecx, $iMaxLoops")

			_FasmAdd("CheckNext:")
			_FasmJumpIf('	dword[ebx] <> edx, Found')
			_FasmAdd("		add ebx, 4")
			_FasmAdd("		loop CheckNext")

			_FasmAdd("mov eax, 0")
			_FasmAdd("jmp Finished")

			_FasmAdd("Found:")
			_FasmAdd("		mov edx, $iPixFound")
			_FasmAdd("		push dword[ebx]")
			_FasmAdd("		pop dword[edx]")
			_FasmAdd("		mov eax, $iMaxLoops")
			_FasmAdd("		sub eax, ecx")
			_FasmAdd("		add eax, 2")
			_FasmAdd("		jmp Finished")

			_FasmAdd("Finished:")

			_FasmEndFunc()

			_FasmCompileMC('_FindFirstDiff')

			Local $pMem = _FasmGetFuncPtr()

			Local $iMaxLoops = (DllStructGetSize($tPix) / 4) - 1
			Local $iDx = DllCallAddress('dword', $pMem, 'dword*', 0, 'struct*', $tPix, 'dword', $iMaxLoops)

			Return $iDx
			EndFunc   ;==>__FindFirstDiff
		#ce
		#EndRegion Include
	#EndRegion BMPSearch Include

#EndRegion BMPSearch


#Region Coprocex UDF
	#cs
		Author: piccaso - autoitsript.com

		User Calltips:
		_CoProc_Create([$sFunction],[$vParameter]) Starts another Process and Calls $sFunc, Returns PID
		_CoProc_ParentPID() Parent Process PID
		_CoProc_SuperGlobalSet($sName,[$vValue],[$sRegistryBase]) Sets or Deletes a Superglobal Variable
		_CoProc_SuperGlobalGet($sName,[$fOption],[$sRegistryBase]) Returns the Value of a Superglobal Variable
		_CoProc_Suspend($vProcess) Suspends all Threads in $vProcess (PID or Name)
		_CoProc_Resume($vProcess) Resumes all Threads in $vProcess (PID or Name)
		_CoProc_ProcessGetWinList($vProcess, $sTitle = Default, $iOption = 0) Enumerates Windows of a Process
		_CoProc_Reciver([$sFunction = ""]) Register/Unregister Reciver Function
		_CoProc_Send($vProcess, $vParameter,[$iTimeout = 500],[$fAbortIfHung = True]) Send Message to Process
		_CoProc_ConsoleForward($iPid1, [$iPid2], [$iPid3], [$iPidn])
		_CoProc_ProcessEmptyWorkingSet($vPid = @AutoItPID,[$hDll_psapi],[$hDll_kernel32]) Removes as many pages as possible from the working set of the specified process.
		_CoProc_DuplicateHandle($dwSourcePid, $hSourceHandle, $dwTargetPid = @AutoItPID, $fCloseSource = False) Returns a Duplicate handle
		_CoProc_CloseHandle($hAny) Close a Handle
		$gs_SuperGlobalRegistryBase
		$gi_CoProcParent
	#ce

	

	;===============================================================================
	;
	; Description:      Starts another Process
	; Parameter(s):     $sFunction - Optional, Name of Function to start in the new Process
	;					$vParameter - Optional, Parameter to pass
	; Requirement(s):   3.2.4.9
	; Return Value(s):  On Success - Pid of the new Process
	;                   On Failure - Set @error to 1
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			Inside the new Process $gi_CoProcParent holds the PID of the Parent Process.
	;					$vParameter must not be Binary, an Array or DllStruct.
	;					If $sFunction is a just a function name Like _CoProc("MyFunc","MyParameter") Then
	;					Call() will be used to Call the function and optional pass the Parameter (only one).
	;					If $sFunction is a expression like _CoProc("MyFunc('MyParameter')") Then
	;					Execute() will be used to evaluate this expression makeing more than one Parameter Possible.
	;					In the 'Execute()' Case the $vParameter Parameter will be ignored.
	;					In both cases there are Limitations, Read doc's (for Execute() and Call()) for more detail.
	;
	;					If $sFunction is Empty ("" or Default) $vParameter can be the name of a Reciver function.
	;
	;
	;===============================================================================
	Func _CoProc_Create($sFunction = Default, $vParameter = Default)
		Local $iPid
		If IsKeyword($sFunction) Or $sFunction = "" Then $sFunction = "__CoProcDummy"
		EnvSet("CoProc", "0x" & Hex(StringToBinary($sFunction, 4)))
		EnvSet("CoProcParent", @AutoItPID)
		If Not IsKeyword($vParameter) Then
			EnvSet("CoProcParameterPresent", "True")
			EnvSet("CoProcParameter", StringToBinary($vParameter, 4))
		Else
			EnvSet("CoProcParameterPresent", "False")
		EndIf
		If @Compiled Then
			$iPid = Run(FileGetShortName(@AutoItExe), @WorkingDir, @SW_HIDE, 1 + 2 + 4)
		Else
			$iPid = Run(FileGetShortName(@AutoItExe) & ' "' & @ScriptFullPath & '"', @WorkingDir, @SW_HIDE, 1 + 2 + 4)
		EndIf
		If @error Then SetError(1)
		Sleep(20)
		Return $iPid
	EndFunc


	Func _CoProc_ParentPID()
		$gi_CoProcParent = ProcessExists($gi_CoProcParent)
		Return SetError(Not $gi_CoProcParent, '', $gi_CoProcParent)
	EndFunc


	;===============================================================================
	;
	; Description:      Set a Superglobal Variable
	; Parameter(s):     $sName - Identifyer for the Superglobal Variable
	;					$vValue - Value to be Stored (optional)
	;					$sRegistryBase - Registry Base Key (optional)
	; Requirement(s):   3.2.4.9
	; Return Value(s):  On Success - Returns True
	;                   On Failure - Returns False and Set
	;												@error to:	1 - Wrong Value Type
	;															2 - Registry Problem
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			$vValue must not be an Array or Struct.
	;					if $vValue is Omitted the Superglobal Variable will be deleted.
	;
	;					Superglobal Variables are Stored in the Registry.
	;					$gs_SuperGlobalRegistryBase is holding the Default Base Key
	;
	;===============================================================================
	Func _CoProc_SuperGlobalSet($sName, $vValue = Default, $sRegistryBase = Default)
		Local $vTmp
		If $sRegistryBase = Default Then $sRegistryBase = $gs_SuperGlobalRegistryBase
		If $vValue = "" Or $vValue = Default Then
			RegDelete($sRegistryBase, $sName)
			If @error Then Return SetError(2, 0, False) ; Registry Problem
		Else
			RegWrite($sRegistryBase, $sName, "REG_BINARY", StringToBinary($vValue, 4))
			If @error Then Return SetError(2, 0, False) ; Registry Problem
		EndIf
		Return True
	EndFunc

	;===============================================================================
	;
	; Description:      Get a Superglobal Variable
	; Parameter(s):     $sName - Identifyer for the Superglobal Variable
	;					$fOption - Optional, if True The Superglobal Variable will be deleted after sucessfully Reading it out
	;					$sRegistryBase - Registry Base Key (optional)
	; Requirement(s):   3.2.4.9
	; Return Value(s):  On Success - Returns the Value of the Superglobal Variable
	;                   On Failure - Set
	;								@error to:	1 - Not Found / Registry Problem
	;											2 - Error Deleting
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			$vValue must not be an Array or Struct.
	;					Superglobal Variables are Stored in the Registry.
	;					$gs_SuperGlobalRegistryBase is holding the Default Base Key
	;
	;===============================================================================
	Func _CoProc_SuperGlobalGet($sName, $fOption = Default, $sRegistryBase = Default)
		Local $vTmp
		If $fOption = "" Or $fOption = Default Then $fOption = False
		If $sRegistryBase = Default Then $sRegistryBase = $gs_SuperGlobalRegistryBase
		$vTmp = RegRead($sRegistryBase, $sName)
		If @error Then Return SetError(1, 0, "") ; Registry Problem
		If $fOption Then
			_CoProc_SuperGlobalSet($sName)
			If @error Then SetError(2)
		EndIf
		Return BinaryToString($vTmp, 4)
	EndFunc

	;===============================================================================
	;
	; Description:      Suspend all Threads in a Process
	; Parameter(s):     $vProcess - Name or PID of Process
	; Requirement(s):   3.1.1.130, Win ME/2k/XP
	; Return Value(s):  On Success - Returns Nr. of Threads Suspended and Set @extended to Nr. of Threads Processed
	;                   On Failure - Returns False and Set
	;												@error to:	1 - Process not Found
	;															2 - Error Calling 'CreateToolhelp32Snapshot'
	;															3 - Error Calling 'Thread32First'
	;															4 - Error Calling 'Thread32Next'
	;															5 - Not all Threads Processed
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			Ported from: http://www.codeproject.com/threads/pausep.asp
	;					Better read the article (and the warnings!) if you want to use it :)
	;
	;===============================================================================
	Func _CoProc_Suspend($vProcess, $iReserved = 0)
		Local $iPid, $vTmp, $hThreadSnap, $ThreadEntry32, $iThreadID, $hThread, $iThreadCnt, $iThreadCntSuccess, $sFunction
		Local $TH32CS_SNAPTHREAD = 0x00000004
		Local $INVALID_HANDLE_VALUE = 0xFFFFFFFF
		Local $THREAD_SUSPEND_RESUME = 0x0002
		Local $THREADENTRY32_StructDef = "int;" _ ; 1 -> dwSize
				 & "int;" _ ; 2 -> cntUsage
				 & "int;" _ ; 3 -> th32ThreadID
				 & "int;" _ ; 4 -> th32OwnerProcessID
				 & "int;" _ ; 5 -> tpBasePri
				 & "int;" _ ; 6 -> tpDeltaPri
				 & "int" ; 7 -> dwFlags
		$iPid = ProcessExists($vProcess)
		If Not $iPid Then Return SetError(1, 0, False) ; Process not found.
		$vTmp = DllCall("kernel32.dll", "ptr", "CreateToolhelp32Snapshot", "int", $TH32CS_SNAPTHREAD, "int", 0)
		If @error Then Return SetError(2, 0, False) ; CreateToolhelp32Snapshot Failed
		If $vTmp[0] = $INVALID_HANDLE_VALUE Then Return SetError(2, 0, False) ; CreateToolhelp32Snapshot Failed
		$hThreadSnap = $vTmp[0]
		$ThreadEntry32 = DllStructCreate($THREADENTRY32_StructDef)
		DllStructSetData($ThreadEntry32, 1, DllStructGetSize($ThreadEntry32))
		$vTmp = DllCall("kernel32.dll", "int", "Thread32First", "ptr", $hThreadSnap, "long", DllStructGetPtr($ThreadEntry32))
		If @error Then Return SetError(3, 0, False) ; Thread32First Failed
		If Not $vTmp[0] Then
			DllCall("kernel32.dll", "int", "CloseHandle", "ptr", $hThreadSnap)
			Return SetError(3, 0, False) ; Thread32First Failed
		EndIf
		While 1
			If DllStructGetData($ThreadEntry32, 4) = $iPid Then
				$iThreadID = DllStructGetData($ThreadEntry32, 3)
				$vTmp = DllCall("kernel32.dll", "ptr", "OpenThread", "int", $THREAD_SUSPEND_RESUME, "int", False, "int", $iThreadID)
				If Not @error Then
					$hThread = $vTmp[0]
					If $hThread Then
						If $iReserved Then
							$sFunction = "ResumeThread"
						Else
							$sFunction = "SuspendThread"
						EndIf
						$vTmp = DllCall("kernel32.dll", "int", $sFunction, "ptr", $hThread)
						If $vTmp[0] <> -1 Then $iThreadCntSuccess += 1
						DllCall("kernel32.dll", "int", "CloseHandle", "ptr", $hThread)
					EndIf
				EndIf
				$iThreadCnt += 1
			EndIf
			$vTmp = DllCall("kernel32", "int", "Thread32Next", "ptr", $hThreadSnap, "long", DllStructGetPtr($ThreadEntry32))
			If @error Then Return SetError(4, 0, False) ; Thread32Next Failed
			If Not $vTmp[0] Then ExitLoop
		WEnd
		DllCall("kernel32.dll", "int", "CloseToolhelp32Snapshot", "ptr", $hThreadSnap) ; CloseHandle
		If Not $iThreadCntSuccess Or $iThreadCnt > $iThreadCntSuccess Then Return SetError(5, $iThreadCnt, $iThreadCntSuccess)
		Return SetError(0, $iThreadCnt, $iThreadCntSuccess)
	EndFunc

	;===============================================================================
	;
	; Description:      Resume all Threads in a Process
	; Parameter(s):     $vProcess - Name or PID of Process
	; Requirement(s):   3.1.1.130, Win ME/2k/XP
	; Return Value(s):  On Success - Returns Nr. of Threads Resumed and Set @extended to Nr. of Threads Processed
	;                   On Failure - Returns False and Set
	;												@error to:	1 - Process not Found
	;															2 - Error Calling 'CreateToolhelp32Snapshot'
	;															3 - Error Calling 'Thread32First'
	;															4 - Error Calling 'Thread32Next'
	;															5 - Not all Threads Processed
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			Ported from: http://www.codeproject.com/threads/pausep.asp
	;					Better read the article (and the warnings!) if you want to use it :)
	;
	;===============================================================================
	Func _CoProc_Resume($vProcess)
		Local $fRval = _CoProc_Suspend($vProcess, True)
		Return SetError(@error, @extended, $fRval)
	EndFunc

	;===============================================================================
	;
	; Description:      Enumerates Windows of a Process
	; Parameter(s):     $vProcess - Name or PID of Process
	;					$sTitle - Optional Title of window to Find
	;					$iOption - Optional, Can be added together
	;								0 - Matches any Window (Default)
	;								2 - Matches any Window Created by GuiCreate() (ClassName: AutoIt v3 GUI)
	;								4 - Matches AutoIt Main Window (ClassName: AutoIt v3)
	;								6 - Matches Any AutoIt Window
	;								16 - Return the first Window Handle found (No Array)
	; Requirement(s):   3.1.1.130
	; Return Value(s):  On Success - Retuns an Array/Handle of Windows found
	;                   On Failure - Set @ERROR to:	1 - Process not Found
	;												2 - Window(s) not Found
	;												3 - GetClassName Failed
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):
	;
	;===============================================================================
	Func _CoProc_ProcessGetWinList($vProcess, $sTitle = Default, $iOption = 0)
		Local $aWinList, $iCnt, $aTmp, $aResult[1], $iPid, $fMatch, $sClassname
		$iPid = ProcessExists($vProcess)
		If Not $iPid Then Return SetError(1) ; Process not Found
		If $sTitle = "" Or IsKeyword($sTitle) Then
			$aWinList = WinList()
		Else
			$aWinList = WinList($sTitle)
		EndIf
		For $iCnt = 1 To $aWinList[0][0]
			$hWnd = $aWinList[$iCnt][1]
			$iProcessId = WinGetProcess($hWnd)
			If $iProcessId = $iPid Then
				If $iOption = 0 Or IsKeyword($iOption) Or $iOption = 16 Then
					$fMatch = True
				Else
					$fMatch = False
					$sClassname = DllCall("user32.dll", "int", "GetClassName", "hwnd", $hWnd, "str", "", "int", 1024)
					If @error Then Return SetError(3) ; GetClassName
					If $sClassname[0] = 0 Then Return SetError(3) ; GetClassName
					$sClassname = $sClassname[2]
					If BitAND($iOption, 2) Then
						If $sClassname = "AutoIt v3 GUI" Then $fMatch = True
					EndIf
					If BitAND($iOption, 4) Then
						If $sClassname = "AutoIt v3" Then $fMatch = True
					EndIf
				EndIf
				If $fMatch Then
					If BitAND($iOption, 16) Then Return $hWnd
					ReDim $aResult[UBound($aResult) + 1]
					$aResult[UBound($aResult) - 1] = $hWnd
				EndIf
			EndIf
		Next
		$aResult[0] = UBound($aResult) - 1
		If $aResult[0] < 1 Then Return SetError(2, 0, 0) ; No Window(s) Found
		Return $aResult
	EndFunc

	;===============================================================================
	;
	; Description:      Register Reciver Function
	; Parameter(s):     $sFunction - Optional, Function name to Register.
	;								Omit to Disable/Unregister
	; Requirement(s):   3.2.4.9
	; Return Value(s):  On Success - Returns True
	;                   On Failure - Returns False and Set
	;												@error to:	1 - Unable to create Reciver Window
	;															2 - Unable to (Un)Register WM_COPYDATA or WM_USER+0x64
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			If the process doesent have a Window it will be created
	;					The Reciver Function must accept 1 Parameter
	;
	;===============================================================================
	Func _CoProc_Reciver($sFunction = Default)
		Local $sHandlerFuction = "__CoProcReciverHandler", $hWnd, $aTmp
		If IsKeyword($sFunction) Then $sFunction = ""
		$hWnd = _CoProc_ProcessGetWinList(@AutoItPID, "", 16 + 2)
		If Not IsHWnd($hWnd) Then
			$hWnd = GUICreate("CoProcEventReciver")
			If @error Then Return SetError(1, 0, False)
		EndIf
		If $sFunction = "" Or IsKeyword($sFunction) Then $sHandlerFuction = ""
		If Not GUIRegisterMsg(0x4A, $sHandlerFuction) Then Return SetError(2, 0, False) ; WM_COPYDATA
		If Not GUIRegisterMsg(0x400 + 0x64, $sHandlerFuction) Then Return SetError(2, 0, False) ; WM_USER+0x64
		$gs_CoProcReciverFunction = $sFunction
		Return True
	EndFunc

	Func __CoProcReciverHandler($hWnd, $iMsg, $WParam, $LParam)
		If $iMsg = 0x4A Then ; WM_COPYDATA
			Local $COPYDATA, $MyData
			$COPYDATA = DllStructCreate("ptr;dword;ptr", $LParam)
			$MyData = DllStructCreate("wchar [" & DllStructGetData($COPYDATA, 2) & "]", DllStructGetData($COPYDATA, 3))
			$gv_CoProcReviverParameter = DllStructGetData($MyData, 1)
			Return 256
		ElseIf $iMsg = 0x400 + 0x64 Then ; WM_USER+0x64
			If $gv_CoProcReviverParameter Then
				Call($gs_CoProcReciverFunction, $gv_CoProcReviverParameter)
				If @error And @Compiled = 0 Then MsgBox(16, "CoProc Error", "Unable to Call: " & $gs_CoProcReciverFunction)
				$gv_CoProcReviverParameter = 0
				Return 0
			EndIf
		EndIf
	EndFunc

	;===============================================================================
	;
	; Description:      Send a Message to a CoProcess
	; Parameter(s):     $vProcess - Name or PID of Process
	;					$vParameter - Parameter to pass
	;					$iTimeout - Optional, Defaults to 500 (msec)
	;					$fAbortIfHung - Optional, Default is True
	; Requirement(s):   3.2.4.9
	; Return Value(s):  On Success - Returns True
	;                   On Failure - Returns False and Set
	;												@error to:	1 - Process not found
	;															2 - Window not found
	;															3 - Timeout/Busy/Hung
	;															4 - PostMessage Falied
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):
	;
	;==========================================================================
	Func _CoProc_Send($vProcess, $vParameter, $iTimeout = 500, $fAbortIfHung = True)
		Local $iPid, $hWndTarget, $MyData, $aTmp, $COPYDATA, $iFuFlags
		$iPid = ProcessExists($vProcess)
		If Not $iPid Then Return SetError(1, 0, False) ; Process not Found
		$hWndTarget = _CoProc_ProcessGetWinList($vProcess, "", 16 + 2)
		If @error Or (Not $hWndTarget) Then Return SetError(2, 0, False) ; Window not found
		$MyData = DllStructCreate("wchar [" & StringLen($vParameter) + 1 & "]")
		$COPYDATA = DllStructCreate("ptr;dword;ptr")
		DllStructSetData($MyData, 1, $vParameter)
		DllStructSetData($COPYDATA, 1, 1)
		DllStructSetData($COPYDATA, 2, DllStructGetSize($MyData))
		DllStructSetData($COPYDATA, 3, DllStructGetPtr($MyData))
		If $fAbortIfHung Then
			$iFuFlags = 0x2 ; SMTO_ABORTIFHUNG
		Else
			$iFuFlags = 0x0 ; SMTO_NORMAL
		EndIf
		$aTmp = DllCall("user32.dll", "int", "SendMessageTimeout", "hwnd", $hWndTarget, "int", 0x4A _ ; WM_COPYDATA
				, "int", 0, "ptr", DllStructGetPtr($COPYDATA), "int", $iFuFlags, "int", $iTimeout, "long*", 0)
		If @error Then Return SetError(3, 0, False) ; SendMessageTimeout Failed
		If Not $aTmp[0] Then Return SetError(3, 0, False) ; SendMessageTimeout Failed
		If $aTmp[7] <> 256 Then Return SetError(3, 0, False)
		$aTmp = DllCall("user32.dll", "int", "PostMessage", "hwnd", $hWndTarget, "int", 0x400 + 0x64, "int", 0, "int", 0)
		If @error Then Return SetError(4, 0, False)
		If Not $aTmp[0] Then Return SetError(4, 0, False)
		Return True
	EndFunc

	;===============================================================================
	;
	; Description:      Forwards StdOut and StdErr from specified Processes to Calling process
	; Parameter(s):     $iPid1 - Pid of Procces
	;					$iPidn - Optional, Up to 16 Processes
	; Requirement(s):   3.1.1.131
	; Return Value(s):  None
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			Processes must provide StdErr and StdOut Streams (See Run())
	;
	;==========================================================================
	Func _CoProc_ConsoleForward($iPid1, $iPid2 = Default, $iPid3 = Default, $iPid4 = Default, $iPid5 = Default, $iPid6 = Default, $iPid7 = Default, $iPid8 = Default, $iPid9 = Default, $iPid10 = Default, $iPid11 = Default, $iPid12 = Default, $iPid13 = Default, $iPid14 = Default, $iPid15 = Default, $iPid16 = Default)
		Local $iPid, $i, $iPeek
		For $i = 1 To 16
			$iPid = Eval("iPid" & $i)
			If $iPid = Default Or Not $iPid Then ContinueLoop
			If ProcessExists($iPid) Then
				$iPeek = StdoutRead($iPeek, 0, True)
				If Not @error And $iPeek > 0 Then
					ConsoleWrite(StdoutRead($iPid))
				EndIf
				$iPeek = StderrRead($iPeek, 0, True)
				If Not @error And $iPeek > 0 Then
					ConsoleWriteError(StderrRead($iPid))
				EndIf
			EndIf
		Next
	EndFunc

	;===============================================================================
	;
	; Description:      Removes as many pages as possible from the working set of the specified process.
	; Parameter(s):     $vPid - Optional, Pid or Process Name
	;					$hDll_psapi - Optional, Handle to psapi.dll
	;					$hDll_kernel32 - Optional, Handle to kernel32.dll
	; Requirement(s):   3.2.1.12
	; Return Value(s):  On Success - nonzero
	;                   On Failure - 0 and sets error to
	;												@error to:	1 - Process Doesent exist
	;															2 - OpenProcess Failed
	;															3 - EmptyWorkingSet Failed
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):			$vPid can be the -1 Pseudo Handle
	;
	;===============================================================================
	Func _CoProc_ProcessEmptyWorkingSet($vPid = @AutoItPID, $hDll_psapi = "psapi.dll", $hDll_kernel32 = "kernel32.dll")
		Local $av_EWS, $av_OP, $iRval
		If $vPid = -1 Then ; Pseudo Handle
			$av_EWS = DllCall($hDll_psapi, "int", "EmptyWorkingSet", "ptr", -1)
		Else
			$vPid = ProcessExists($vPid)
			If Not $vPid Then Return SetError(1, 0, 0) ; Process Doesent exist
			$av_OP = DllCall($hDll_kernel32, "int", "OpenProcess", "dword", 0x1F0FFF, "int", 0, "dword", $vPid)
			If $av_OP[0] = 0 Then Return SetError(2, 0, 0) ; OpenProcess Failed
			$av_EWS = DllCall($hDll_psapi, "int", "EmptyWorkingSet", "ptr", $av_OP[0])
			DllCall($hDll_kernel32, "int", "CloseHandle", "int", $av_OP[0])
		EndIf
		If $av_EWS[0] Then
			Return $av_EWS[0]
		Else
			Return SetError(3, 0, 0) ; EmptyWorkingSet Failed
		EndIf
	EndFunc


	;===============================================================================
	;
	; Description:      Duplicates a Handle from or for another process
	; Parameter(s):     $dwSourcePid - Pid from Source Process
	;					$hSourceHandle - The Handle to duplicate
	;					$dwTargetPid - Optional, Pid from Target Procces - Defaults to current process
	;					$fCloseSource - Optional, Close the source handle - Defaults to False
	; Requirement(s):   3.2.4.9
	; Return Value(s):  On Success - Duplicated Handle
	;                   On Failure - 0 and sets error to
	;												@error to:	1 - Api OpenProcess Failed
	;															2 - Api DuplicateHandle Falied
	; Author(s):        Florian 'Piccaso' Fida
	; Note(s):
	;
	;===============================================================================
	Func _CoProc_DuplicateHandle($dwSourcePid, $hSourceHandle, $dwTargetPid = @AutoItPID, $fCloseSource = False)
		Local $hTargetHandle, $hPrSource, $hPrTarget, $dwOptions
		$hPrSource = __dh_OpenProcess($dwSourcePid)
		$hPrTarget = __dh_OpenProcess($dwTargetPid)
		If $hPrSource = 0 Or $hPrTarget = 0 Then
			_CoProc_CloseHandle($hPrSource)
			_CoProc_CloseHandle($hPrTarget)
			Return SetError(1, 0, 0)
		EndIf
		; DUPLICATE_CLOSE_SOURCE = 0x00000001
		; DUPLICATE_SAME_ACCESS = 0x00000002
		If $fCloseSource <> False Then
			$dwOptions = 0x01 + 0x02
		Else
			$dwOptions = 0x02
		EndIf
		$hTargetHandle = DllCall("kernel32.dll", "int", "DuplicateHandle", "ptr", $hPrSource, "ptr", $hSourceHandle, "ptr", $hPrTarget, "long*", 0, "dword", 0, "int", 1, "dword", $dwOptions)
		If @error Then Return SetError(2, 0, 0)
		If $hTargetHandle[0] = 0 Or $hTargetHandle[4] = 0 Then
			_CoProc_CloseHandle($hPrSource)
			_CoProc_CloseHandle($hPrTarget)
			Return SetError(2, 0, 0)
		EndIf
		Return $hTargetHandle[4]
	EndFunc


	Func _CoProc_CloseHandle($hAny)
		If $hAny = 0 Then Return SetError(1, 0, 0)
		Local $fch = DllCall("kernel32.dll", "int", "CloseHandle", "ptr", $hAny)
		If @error Then Return SetError(1, 0, 0)
		Return $fch[0]
	EndFunc



	#Region Internal Functions
		Func __dh_OpenProcess($dwProcessId)
			; PROCESS_DUP_HANDLE = 0x40
			Local $hPr = DllCall("kernel32.dll", "ptr", "OpenProcess", "dword", 0x40, "int", 0, "dword", $dwProcessId)
			If @error Then Return SetError(1, 0, 0)
			Return $hPr[0]
		EndFunc

		Func __CoProcStartup()
			Local $sCmd = EnvGet("CoProc")
			If StringLeft($sCmd, 2) = "0x" Then
				$sCmd = BinaryToString($sCmd, 4)
				$gi_CoProcParent = Number(EnvGet("CoProcParent"))
				If StringInStr($sCmd, "(") And StringInStr($sCmd, ")") Then
					Execute($sCmd)
					If @error And Not @Compiled Then MsgBox(16, "CoProc Error", "Unable to Execute: " & $sCmd)
					Exit
				EndIf
				If EnvGet("CoProcParameterPresent") = "True" Then
					Call($sCmd, BinaryToString(EnvGet("CoProcParameter"), 4))
					If @error And Not @Compiled Then MsgBox(16, "CoProc Error", "Unable to Call: " & $sCmd & @LF & "Parameter: " & BinaryToString(EnvGet("CoProcParameter"), 4))
				Else
					Call($sCmd)
					If @error And Not @Compiled Then MsgBox(16, "CoProc Error", "Unable to Call: " & $sCmd)
				EndIf
				Exit
			EndIf
		EndFunc
		Func __CoProcDummy($vPar = Default)
			If Not IsKeyword($vPar) Then _CoProc_Reciver($vPar)
			While ProcessExists($gi_CoProcParent)
				Sleep(500)
			WEnd
		EndFunc
		__CoProcStartup()
	#EndRegion


#EndRegion Coproex


#Region AutoitObject UDF
	#cs
	# AutoItObject Internal
	# @author genius257
	# @version 2.0.0
	#ce

	

	Func __createObject($QueryInterface=QueryInterface, $AddRef=AddRef, $Release=Release, $GetTypeInfoCount=GetTypeInfoCount, $GetTypeInfo=GetTypeInfo, $GetIDsOfNames=GetIDsOfNames, $Invoke=Invoke)
		Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties;BYTE lock;PTR __destructor"
		Local $tObject = DllStructCreate($tagObject)

		$QueryInterface = DllCallbackRegister($QueryInterface, "LONG", "ptr;ptr;ptr")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
		DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

		$AddRef = DllCallbackRegister($AddRef, "dword", "PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
		DllStructSetData($tObject, "Callbacks", $AddRef, 2)

		$Release = DllCallbackRegister($Release, "dword", "PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
		DllStructSetData($tObject, "Callbacks", $Release, 3)

		$GetTypeInfoCount = DllCallbackRegister($GetTypeInfoCount, "long", "ptr;ptr")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoCount), 4)
		DllStructSetData($tObject, "Callbacks", $GetTypeInfoCount, 4)

		$GetTypeInfo = DllCallbackRegister($GetTypeInfo, "long", "ptr;uint;int;ptr")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfo), 5)
		DllStructSetData($tObject, "Callbacks", $GetTypeInfo, 5)

		$GetIDsOfNames = DllCallbackRegister($GetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetIDsOfNames), 6)
		DllStructSetData($tObject, "Callbacks", $GetIDsOfNames, 6)

		$Invoke = DllCallbackRegister($Invoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Invoke), 7)
		DllStructSetData($tObject, "Callbacks", $Invoke, 7)

		DllStructSetData($tObject, "RefCount", 2) ; initial ref count is 1
		DllStructSetData($tObject, "Size", 7) ; number of interface methods

		Local $pData = MemCloneGlob($tObject)

		Local $tObject = DllStructCreate($tagObject, $pData)

		DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
		Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IDispatch, Default, True) ; pointer that's wrapped into object
	EndFunc

	#cs
	# @internal
	#ce
	Func QueryInterface($pSelf, $pRIID, $pObj)
		If $pObj=0 Then Return $E_POINTER
		Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
		If (Not ($sGUID=$IID_IDispatch)) And (Not ($sGUID=$IID_IUnknown)) Then Return $E_NOINTERFACE
		Local $tStruct = DllStructCreate("ptr", $pObj)
		DllStructSetData($tStruct, 1, $pSelf)
		Return $S_OK
	EndFunc

	#cs
	# @internal
	#ce
	Func AddRef($pSelf)
	   Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	   $tStruct.Ref += 1
	   Return $tStruct.Ref
	EndFunc

	#cs
	# @internal
	#ce
	Func Release($pSelf)
		Local $i
		Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
		$tStruct.Ref -= 1
		If $tStruct.Ref == 0 Then; initiate garbage collection
			Local $pDescructor = DllStructGetData(DllStructCreate("PTR", $pSelf + __AOI_GetPtrOffset("__destructor")),1)
			If Not ($pDescructor=0) Then;TODO
				Local $tVARIANT = DllStructCreate($tagVARIANT, $pDescructor)
				DllStructSetData(DllStructCreate("PTR", $pSelf + __AOI_GetPtrOffset("__destructor")),1,0)
				Local $IDispatch = __createObject()
				$IDispatch.a=0
				Local $pProperty = DllStructGetData(DllStructCreate("PTR", Ptr($IDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
				Local $pVARIANT = DllStructGetData(DllStructCreate($tagProperty, $pProperty),"Variant")
				VariantClear($pVARIANT)
				VariantCopy($pVARIANT, $tVARIANT)
				Local $f__destructor = $IDispatch.a
				VariantClear($pVARIANT)
				DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
				Local $tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
				$tVARIANT.vt = $VT_DISPATCH
				$tVARIANT.data = $pSelf
				Call($f__destructor, $IDispatch.a)
				VariantClear($pVARIANT)
				$IDispatch=0
			EndIf
			DllStructSetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2 + (@AutoItX64?8:4)),1,1);lock
			Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1);get first property
			DllStructSetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1,0);detatch properties from object
			While 1;releases all properties
				If $pProperty=0 Then ExitLoop
				Local $tProperty = DllStructCreate($tagProperty, $pProperty)
				Local $_pProperty = $pProperty
				$pProperty = $tProperty.Next
				If Not ($tProperty.__getter=0) Then
					VariantClear($tProperty.__getter)
					_MemGlobalFree(GlobalHandle($tProperty.__getter))
				EndIf
				If Not ($tProperty.__setter=0) Then
					VariantClear($tProperty.__setter)
					_MemGlobalFree(GlobalHandle($tProperty.__setter))
				EndIf
				VariantClear($tProperty.Variant)
				_MemGlobalFree(GlobalHandle($tProperty.Variant))
				_WinAPI_FreeMemory($tProperty.Name)
				$tProperty=0
				_MemGlobalFree(GlobalHandle($_pProperty))
			WEnd
			Local $pMethods = $pSelf + (@AutoItX64?8:4)
			_MemGlobalFree(GlobalHandle($pSelf-8))
			Return 0
		EndIf
		Return $tStruct.Ref
	EndFunc

	#cs
	# @internal
	#ce
	Func GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
		Local $tIds = DllStructCreate("long", $rgDispId); 2,147,483,647 properties available to define, per object. And 2,147,483,647 private properties to set in the negative space, per object.
		Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
		Local $tProperty = 0
		Local $iSize2

		Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
		Local $iSize = _WinAPI_StrLen($pStr, True)
		Local $t_rgszNames = DllStructCreate("WCHAR["&$iSize&"]", $pStr)
		Local $s_rgszName = DllStructGetData($t_rgszNames, 1)

		Switch $s_rgszName
			Case "__assign"
				DllStructSetData($tIds, 1, -2)
			Case "__isExtensible"
				DllStructSetData($tIds, 1, -3)
			Case "__case"
				DllStructSetData($tIds, 1, -4)
			Case "__freeze"
				DllStructSetData($tIds, 1, -5)
			Case "__isFrozen"
				DllStructSetData($tIds, 1, -6)
			Case "__isSealed"
				DllStructSetData($tIds, 1, -7)
			Case "__keys"
				DllStructSetData($tIds, 1, -8)
			Case "__preventExtensions"
				DllStructSetData($tIds, 1, -9)
			Case "__defineGetter"
				DllStructSetData($tIds, 1, -10)
			Case "__defineSetter"
				DllStructSetData($tIds, 1, -11)
			Case "__lookupGetter"
				DllStructSetData($tIds, 1, -12)
			Case "__lookupSetter"
				DllStructSetData($tIds, 1, -13)
			Case "__seal"
				DllStructSetData($tIds, 1, -14)
			Case "__destructor"
				DllStructSetData($tIds, 1, -16)
			Case "__unset"
				DllStructSetData($tIds, 1, -17)
			Case "__get"
				DllStructSetData($tIds, 1, -18)
			Case "__set"
				DllStructSetData($tIds, 1, -19)
			Case Else
				DllStructSetData($tIds, 1, -1)
		EndSwitch
		If DllStructGetData($tIds, 1) <> -1 Then Return $S_OK

		Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
		Local $bCase = Not (BitAND($iLock, $__AOI_LOCK_CASE)>0)
		Local $pProperty = __AOI_PropertyGetFromName(__AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("Properties"), "ptr"), DllStructGetData($t_rgszNames, 1), $bCase)
		Local $iID = @error<>0?-1:@extended
		Local $iIndex = @extended

		If ($iID==-1) And BitAND($iLock, $__AOI_LOCK_CREATE)=0 Then
			Local $pData = __AOI_PropertyCreate(DllStructGetData($t_rgszNames, 1))
			If $iIndex=-1 Then;first item in list
				DllStructSetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")), 1, $pData)
			Else
				$tProperty = DllStructCreate($tagProperty, $pProperty)
				$tProperty.next = $pData
			EndIf
			$iID = $iIndex+1
		EndIf

		If $iID==-1 Then Return $DISP_E_UNKNOWNNAME
		DllStructSetData($tIds, 1, $iID)
		Return $S_OK
	EndFunc

	#cs
	# @internal
	#ce
	Func GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
		If $iTInfo<>0 Then Return $DISP_E_BADINDEX
		If $ppTInfo=0 Then Return $E_INVALIDARG
		Return $S_OK
	EndFunc

	#cs
	# @internal
	#ce
	Func GetTypeInfoCount($pSelf, $pctinfo)
		DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 0)
		Return $S_OK
	EndFunc

	#cs
	# @internal
	#ce
	Func Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
		If $dispIdMember=-1 Then Return $DISP_E_MEMBERNOTFOUND
		Local $tVARIANT, $_tVARIANT, $tDISPPARAMS
		Local $t
		Local $i

		Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
		Local $tProperty = DllStructCreate($tagProperty, $pProperty)

		If $dispIdMember<-1 Then
			If $dispIdMember=-3 Then;__isExtensible
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				Local $iExtensible = $__AOI_LOCK_CREATE
				$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
				$tVARIANT.vt = $VT_BOOL
				$tVARIANT.data = (BitAND($iLock, $iExtensible) = $iExtensible)?1:0
				Return $S_OK
			EndIf

			If $dispIdMember=-4 Then;__case
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) Then
					If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
					$tVARIANT=DllStructCreate($tagVARIANT, $pVarResult)
					$tVARIANT.vt = $VT_BOOL
					Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
					$tVARIANT.data = (BitAND($iLock, $__AOI_LOCK_CASE)>0)?0:1
				Else; $DISPATCH_PROPERTYPUT
					If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
					$tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
					If $tVARIANT.vt<>$VT_BOOL Then Return $DISP_E_BADVARTYPE
					Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
					If BitAND($iLock, $__AOI_LOCK_UPDATE)>0 Then Return $DISP_E_EXCEPTION
					Local $tLock = DllStructCreate("BYTE", $pSelf + __AOI_GetPtrOffset("lock"))
					$b = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
					DllStructSetData($tLock, 1, _
						(Not $b) ? BitOR(DllStructGetData($tLock, 1), $__AOI_LOCK_CASE) : BitAND(DllStructGetData($tLock, 1), BitNOT(BitShift(1 , 0-(Log($__AOI_LOCK_CASE)/log(2))))) _
					)

				EndIf
				Return $S_OK
			EndIf

			If $dispIdMember=-13 Then;__lookupSetter
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT

				$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
				DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
				DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
				$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
				$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
				If Not GetIDsOfNames($pSelf, $riid, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) == $S_OK Then Return $DISP_E_EXCEPTION

				$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
				$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)
				If Not $tProperty.__setter=0 Then
					VariantClear($pVarResult)
					VariantCopy($pVarResult, $tProperty.__setter)
				EndIf
				Return $S_OK
			EndIf

			If $dispIdMember=-12 Then;__lookupGetter
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT

				$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
				DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
				DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
				$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
				$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
				If Not GetIDsOfNames($pSelf, $riid, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) == $S_OK Then Return $DISP_E_EXCEPTION

				$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
				$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)
				If Not $tProperty.__getter=0 Then
					VariantClear($pVarResult)
					VariantCopy($pVarResult, $tProperty.__getter)
				EndIf
				Return $S_OK
			EndIf

			If $dispIdMember=-2 Then;__assign
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION


				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs=0 Then Return $DISP_E_BADPARAMCOUNT

				Local $tVARIANT = DllStructCreate($tagVARIANT)
				Local $iVARIANT = DllStructGetSize($tVARIANT)
				Local $i
				Local $pExternalProperty, $tExternalProperty
				Local $pProperty, $tProperty
				Local $iID, $iIndex, $pData
				For $i=$tDISPPARAMS.cArgs-1 To 0 Step -1
					$tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+$iVARIANT*$i)
					If Not (DllStructGetData($tVARIANT, "vt")==$VT_DISPATCH) Then Return $DISP_E_BADVARTYPE
					$pExternalProperty = __AOI_GetPtrValue(DllStructGetData($tVARIANT, "data") + __AOI_GetPtrOffset("Properties"), "ptr")
					While 1
						If $pExternalProperty = 0 Then ExitLoop
						$tExternalProperty = DllStructCreate($tagProperty, $pExternalProperty)

						$pProperty = __AOI_PropertyGetFromName(__AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("Properties"), "ptr"), _WinAPI_GetString($tExternalProperty.Name), False)
						$iID = @error<>0?-1:@extended
						$iIndex = @extended

						If ($iID==-1) Then
							$pData = __AOI_PropertyCreate(_WinAPI_GetString($tExternalProperty.Name))
							$tProperty = DllStructCreate($tagProperty, $pData)
							VariantClear($tProperty.Variant)
							VariantCopy($tProperty.Variant, $tExternalProperty.Variant)

							If $iIndex=-1 Then;first item in list
								DllStructSetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")), 1, $pData)
							Else
								$tProperty = DllStructCreate($tagProperty, $pProperty)
								$tProperty.Next = $pData
							EndIf
						Else
							$tProperty = DllStructCreate($tagProperty, $pProperty)
							VariantClear($tProperty.Variant)
							VariantCopy($tProperty.Variant, $tExternalProperty.Variant)
						EndIf

						$pExternalProperty = $tExternalProperty.Next
					WEnd
				Next
				Return $S_OK
			EndIf

			If $dispIdMember=-7 Then;__isSealed
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				Local $iSeal = $__AOI_LOCK_CREATE + $__AOI_LOCK_DELETE
				$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
				$tVARIANT.vt = $VT_BOOL
				$tVARIANT.data = (BitAND($iLock, $iSeal) = $iSeal)?1:0
				Return $S_OK
			EndIf

			If $dispIdMember=-6 Then;__isFrozen
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				Local $iFreeze = $__AOI_LOCK_CREATE + $__AOI_LOCK_UPDATE + $__AOI_LOCK_DELETE
				$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
				$tVARIANT.vt = $VT_BOOL
				$tVARIANT.data = (BitAND($iLock, $iFreeze) = $iFreeze)?1:0
				Return $S_OK
			EndIf

			If $dispIdMember=-18 Then;__get
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				If $tVARIANT.vt<>$VT_BSTR Then Return $DISP_E_BADVARTYPE
				$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
				DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
				DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
				$t.str_ptr = DllStructGetData($tVARIANT, "data")
				$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
				If Not GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION
				Return Invoke($pSelf, $t.id, $riid, $lcid, $DISPATCH_PROPERTYGET, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
			EndIf

			If $dispIdMember=-19 Then;__set
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
				If $tVARIANT.vt<>$VT_BSTR Then Return $DISP_E_BADVARTYPE
				$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
				DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
				$t.str_ptr = DllStructGetData($tVARIANT, "data")
				$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
				If Not GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION
				$tDISPPARAMS.cArgs=1
				Return Invoke($pSelf, $t.id, $riid, $lcid, $DISPATCH_PROPERTYPUT, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
			EndIf

			If $dispIdMember=-16 Then;__destructor
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
				If (Not (DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs),"vt")==$VT_RECORD)) And (Not (DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs),"vt")==$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
				Local $tVARIANT = DllStructCreate($tagVARIANT)
				Local $pVARIANT = MemCloneGlob($tVARIANT)
				$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
				VariantInit($pVARIANT)
				VariantCopy($pVARIANT, $tDISPPARAMS.rgvargs)
				DllStructSetData(DllStructCreate("PTR", $pSelf + __AOI_GetPtrOffset("__destructor")),1,$pVARIANT)
				Return $S_OK
			EndIf

			If $dispIdMember=-5 Then;__freeze
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
				Local $tLock = DllStructCreate("BYTE", $pSelf + __AOI_GetPtrOffset("lock"))
				DllStructSetData($tLock, 1, BitOR(DllStructGetData($tLock, 1), $__AOI_LOCK_CREATE + $__AOI_LOCK_DELETE + $__AOI_LOCK_UPDATE))
				Return $S_OK
			EndIf

			If $dispIdMember=-14 Then;__seal
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
				Local $tLock = DllStructCreate("BYTE", $pSelf + __AOI_GetPtrOffset("lock"))
				DllStructSetData($tLock, 1, BitOR(DllStructGetData($tLock, 1), $__AOI_LOCK_CREATE + $__AOI_LOCK_DELETE))
				Return $S_OK
			EndIf

			If $dispIdMember=-9 Then;__preventExtensions
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
				Local $tLock = DllStructCreate("BYTE", $pSelf + __AOI_GetPtrOffset("lock"))
				DllStructSetData($tLock, 1, BitOR(DllStructGetData($tLock, 1), $__AOI_LOCK_CREATE))
				Return $S_OK
			EndIf

			If $dispIdMember=-17 Then;__unset
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				If BitAND($iLock, $__AOI_LOCK_DELETE)>0 Then Return $DISP_E_EXCEPTION

				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				If Not($VT_BSTR==$tVARIANT.vt) Then Return $DISP_E_BADVARTYPE
				Local $sProperty = _WinAPI_GetString($tVARIANT.data);the string to search for
				Local $tProperty=0,$tProperty_Prev
				While 1
					If $pProperty=0 Then ExitLoop
					$tProperty_Prev = $tProperty
					$tProperty = DllStructCreate($tagProperty, $pProperty)
					If _WinAPI_GetString($tProperty.Name)==$sProperty Then
						If $tProperty_Prev==0 Then
							DllStructSetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")), 1, $tProperty.Next)
						Else
							$tProperty_Prev.Next = $tProperty.next
						EndIf
						VariantClear($tProperty.Variant)
						_MemGlobalFree(GlobalHandle($tProperty.Variant))
						$tProperty = 0
						_MemGlobalFree(GlobalHandle($pProperty))
						Return $S_OK
					EndIf
					$pProperty = $tProperty.Next
				WEnd
				Return $DISP_E_MEMBERNOTFOUND
			EndIf

			If ($dispIdMember=-8) Then;__keys
				Local $aKeys[0]
				Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
				While 1
					If $pProperty=0 Then ExitLoop
					Local $tProperty = DllStructCreate($tagProperty, $pProperty)
					ReDim $aKeys[UBound($aKeys,1)+1]
					$aKeys[UBound($aKeys,1)-1] = DllStructGetData(DllStructCreate("WCHAR["&_WinAPI_StrLen($tProperty.Name)&"]", $tProperty.Name), 1)
					If $tProperty.next=0 Then ExitLoop
					$pProperty = $tProperty.next
				WEnd
				Local $oIDispatch = __createObject()
				$oIDispatch.a=$aKeys
				VariantClear($pVarResult)
				VariantCopy($pVarResult, DllStructGetData(DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + __AOI_GetPtrOffset("Properties")),1)), "Variant"))
				$oIDispatch=0
				Return $S_OK
			EndIf

			If ($dispIdMember=-10) Then;__defineGetter
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION

				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				If Not (($tVARIANT.vt==$VT_RECORD) Or ($tVARIANT.vt==$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
				Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
				If Not ($tVARIANT2.vt==$VT_BSTR) Then Return $DISP_E_BADVARTYPE

				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
				DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
				DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
				$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
				$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
				GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"))

				$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
				$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)

				If ($tProperty.__getter=0) Then
					Local $tVARIANT_Getter = DllStructCreate($tagVARIANT)
					$pVARIANT_Getter = MemCloneGlob($tVARIANT_Getter)
					VariantInit($pVARIANT_Getter)
				Else
					Local $pVARIANT_Getter = $tProperty.__getter
					VariantClear($pVARIANT_Getter)
				EndIf
				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				VariantCopy($pVARIANT_Getter, $tVARIANT)
				$tProperty.__getter = $pVARIANT_Getter
				Return $S_OK
			ElseIf ($dispIdMember=-11) Then;defineSetter
				Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
				If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION

				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				If Not (($tVARIANT.vt==$VT_RECORD) Or ($tVARIANT.vt==$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
				Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
				If Not ($tVARIANT2.vt==$VT_BSTR) Then Return $DISP_E_BADVARTYPE

				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
				DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
				DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
				$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
				$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
				GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"))

				$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AOI_GetPtrOffset("Properties")),1)
				$tProperty = DllStructCreate($tagProperty, $pProperty)

				$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				For $i=1 To $t.id
					$pProperty = $tProperty.Next
					$tProperty = DllStructCreate($tagProperty, $pProperty)
				Next
				If ($tProperty.__setter=0) Then
					Local $tVARIANT_Setter = DllStructCreate($tagVARIANT)
					Local $pVARIANT_Setter = MemCloneGlob($tVARIANT_Setter)
					VariantInit($pVARIANT_Setter)
				Else
					Local $pVARIANT_Setter = $tProperty.__setter
					VariantClear($pVARIANT_Setter)
				EndIf
				VariantCopy($pVARIANT_Setter, $tVARIANT)
				$tProperty.__setter = $pVARIANT_Setter
				Return $S_OK
			EndIf
			Return $DISP_E_EXCEPTION
		EndIf

		For $i=1 To $dispIdMember
			$pProperty = $tProperty.Next
			$tProperty = DllStructCreate($tagProperty, $pProperty)
		Next

		$tVARIANT = DllStructCreate($tagVARIANT, $tProperty.Variant)

		If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) Then
			$_tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
			If Not($tProperty.__getter = 0) Then
				$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
				Local $oIDispatch = __createObject()
				$oIDispatch.val = 0
				$oIDispatch.ret = 0
				DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
				$oIDispatch.parent = 0
				Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + __AOI_GetPtrOffset("Properties")),1))
				$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
				$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
				$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
				$tVARIANT.vt = $VT_DISPATCH
				$tVARIANT.data = $pSelf
				$oIDispatch.arguments = __createObject();
				$oIDispatch.arguments.length=$tDISPPARAMS.cArgs
				Local $aArguments[$tDISPPARAMS.cArgs], $iArguments=$tDISPPARAMS.cArgs-1
				Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + __AOI_GetPtrOffset("Properties")),1)
				Local $_tProperty = DllStructCreate($tagProperty, $_pProperty)
				For $i=0 To $iArguments
					VariantClear($_tProperty.Variant)
					VariantCopy($_tProperty.Variant, $tDISPPARAMS.rgvargs+(($iArguments-$i)*DllStructGetSize($_tVARIANT)))
					$aArguments[$i]=$oIDispatch.val
				Next
				$oIDispatch.arguments.values=$aArguments
				$oIDispatch.arguments.__seal()
				$oIDispatch.__defineSetter("parent", PrivateProperty)
				VariantClear($_tProperty.Variant)
				VariantCopy($_tProperty.Variant, $tProperty.__getter)
				Local $fGetter = $oIDispatch.val
				VariantClear($_tProperty.Variant)
				VariantCopy($_tProperty.Variant, $tProperty.Variant)
				$oIDispatch.__seal()
				Local $mRet = Call($fGetter, $oIDispatch)
				Local $iError = @error, $iExtended = @extended
				VariantClear($tProperty.Variant)
				VariantCopy($tProperty.Variant, $_tProperty.Variant)
				$oIDispatch.ret = $mRet
				$_tProperty = DllStructCreate($tagProperty, $_tProperty.Next)
				VariantCopy($pVarResult, $_tProperty.Variant)
				$oIDispatch=0
				Return ($iError<>0)?$DISP_E_EXCEPTION:$S_OK
				Return $S_OK
			EndIf

			VariantCopy($pVarResult, $tVARIANT)
			Return $S_OK
		Else; ~ $DISPATCH_PROPERTYPUT
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If Not ($tProperty.__setter=0) Then
				Local $oIDispatch = __createObject()
				$oIDispatch.val = 0
				$oIDispatch.ret = 0
				DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
				$oIDispatch.parent = 0
				Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + __AOI_GetPtrOffset("Properties")),1))
				$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
				$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
				$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
				$tVARIANT.vt = $VT_DISPATCH
				$tVARIANT.data = $pSelf
				Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + __AOI_GetPtrOffset("Properties")),1)
				Local $_tProperty = DllStructCreate($tagProperty, $_pProperty)
				Local $_tProperty2 = DllStructCreate($tagProperty, $_tProperty.Next)
				VariantClear($_tProperty.Variant)
				VariantCopy($_tProperty.Variant, $tProperty.__setter)
				VariantClear($_tProperty2.Variant)
				VariantCopy($_tProperty2.Variant, $tDISPPARAMS.rgvargs)
				Local $fSetter = $oIDispatch.val
				VariantClear($_tProperty.Variant)
				VariantCopy($_tProperty.Variant, $tProperty.Variant)
				$oIDispatch.__seal()
				Local $mRet = Call($fSetter, $oIDispatch)
				Local $iError = @error, $iExtended = @extended
				VariantClear($tProperty.Variant)
				VariantCopy($tProperty.Variant, $_tProperty.Variant)
				$oIDispatch.ret = $mRet
				$_tProperty = DllStructCreate($tagProperty, $_tProperty.Next)
				VariantCopy($pVarResult, $_tProperty.Variant)
				$oIDispatch=0
				Return ($iError<>0)?$DISP_E_EXCEPTION:$S_OK
				Return $S_OK
			EndIf

			Local $iLock = __AOI_GetPtrValue($pSelf + __AOI_GetPtrOffset("lock"), "BYTE")
			If BitAND($iLock, $__AOI_LOCK_UPDATE)>0 Then Return $DISP_E_EXCEPTION

			$_tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			VariantClear($tVARIANT)
			VariantCopy($tVARIANT, $_tVARIANT)
		EndIf
		Return $S_OK
	EndFunc

	Func MemCloneGlob($tObject);clones DllStruct to Global memory and return pointer to new allocated memory
	   Local $iSize = DllStructGetSize($tObject)
	   Local $hData = _MemGlobalAlloc($iSize, $GMEM_MOVEABLE)
	   Local $pData = _MemGlobalLock($hData)
	   _MemMoveMemory(DllStructGetPtr($tObject), $pData, $iSize)
	   Return $pData
	EndFunc

	Func GlobalHandle($pMem)
	   Local $aRet = DllCall("Kernel32.dll", "ptr", "GlobalHandle", "ptr", $pMem)
	   If @error<>0 Then Return SetError(@error, @extended, 0)
	   If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	   Return $aRet[0]
	EndFunc

	Func LocalSize($hMem)
		Local $aRet = DllCall("Kernel32.dll", "UINT", "LocalSize", "handle", $hMem)
		If @error<>0 Then Return SetError(@error, @extended, 0)
		If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
		Return $aRet[0]
	EndFunc

	Func LocalHandle($pMem)
		Local $aRet = DllCall("Kernel32.dll", "handle", "LocalHandle", "ptr", $pMem)
		If @error<>0 Then Return SetError(@error, @extended, 0)
		If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
		Return $aRet[0]
	EndFunc

	Func HeapSize($hHeap, $dwFlags, $lpMem)
		Local $aRet = DllCall("Kernel32.dll", "ULONG_PTR", "HeapSize", "handle", $hHeap, "dword", $dwFlags, "LONG_PTR", $lpMem)
		If @error<>0 Then Return SetError(@error, @extended, 0)
		If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
		Return $aRet[0]
	EndFunc

	Func GetProcessHeap()
		Local $aRet = DllCall("Kernel32.dll", "handle", "GetProcessHeap")
		If @error<>0 Then Return SetError(@error, @extended, 0)
		If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
		Return $aRet[0]
	EndFunc

	Func VariantInit($tVARIANT)
		Local $aRet=DllCall("OleAut32.dll","LONG","VariantInit",IsDllStruct($tVARIANT)?"struct*":"PTR",$tVARIANT)
		If @error<>0 Then Return SetError(-1, @error, 0)
		If $aRet[0]<>$S_OK Then SetError($aRet, 0, $tVARIANT)
		Return 1
	EndFunc

	Func VariantCopy($tVARIANT_Dest,$tVARIANT_Src)
		Local $aRet=DllCall("OleAut32.dll","LONG","VariantCopy",IsDllStruct($tVARIANT_Dest)?"struct*":"PTR",$tVARIANT_Dest, IsDllStruct($tVARIANT_Src)?"struct*":"PTR", $tVARIANT_Src)
		If @error<>0 Then Return SetError(-1, @error, 0)
		If $aRet[0]<>$S_OK Then SetError($aRet, 0, 0)
		Return 1
	EndFunc

	Func VariantClear($tVARIANT)
		Local $aRet=DllCall("OleAut32.dll","LONG","VariantClear",IsDllStruct($tVARIANT)?"struct*":"PTR",$tVARIANT)
		If @error<>0 Then Return SetError(-1, @error, 0)
		If $aRet[0]<>$S_OK Then SetError($aRet, 0, $tVARIANT)
		Return 1
	EndFunc

	Func VariantChangeType()
		;TODO
	EndFunc

	Func VariantChangeTypeEx()
		;TODO
	EndFunc

	Func PrivateProperty()
		Return SetError(1, 1, 0)
	EndFunc

	#cs
	# @internal
	#ce
	Func __AOI_PropertyGetFromName($pProperty, $sName, $bCase = True)
		Local $iID = -1
		Local $iIndex=-1
		Local $iName = StringLen($sName)
		Local $iPropertyName
		While 1
			If $pProperty=0 Then ExitLoop
			$iIndex+=1
			$tProperty = DllStructCreate($tagProperty, $pProperty)
			$iPropertyName = _WinAPI_StrLen($tProperty.Name, True)
			$sPropertyName = DllStructGetData(DllStructCreate("WCHAR["&$iPropertyName&"]", $tProperty.Name), 1)
			If ($iPropertyName=$iName) And ( ($bCase And $sPropertyName==$sName) Or ((Not $bCase) And $sPropertyName=$sName) ) Then
				$iID = $iIndex
				ExitLoop
			EndIf
			If $tProperty.next=0 Then ExitLoop
			$pProperty = $tProperty.next
		WEnd
		If $iID==-1 Then Return SetError(1, $iIndex, $pProperty)
		Return SetError(0, $iID, $pProperty)
	EndFunc

	#cs
	# @internal
	#ce
	Func __AOI_PropertyGetFromId($pProperty, $iID)
		Local $tProperty = DllStructCreate($tagProperty, $pProperty)
		Local $i
		For $i=1 To $iID
			$pProperty = $tProperty.Next
			$tProperty = DllStructCreate($tagProperty, $pProperty)
		Next
		Return $tProperty
	EndFunc

	#cs
	# Create new empty named property
	# @internal
	# @param String $sName new property name
	# @return Pointer new empty property
	#ce
	Func __AOI_PropertyCreate($sName)
		$tProp = DllStructCreate($tagProperty)
		$pData = MemCloneGlob($tProp)
		$tProp = DllStructCreate($tagProperty, $pData)
			$tProp.Name = _WinAPI_CreateString($sName)
				$tVARIANT = DllStructCreate($tagVARIANT)
				$pVARIANT = MemCloneGlob($tVARIANT)
				$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
				$tVARIANT.vt = $VT_EMPTY
				VariantInit($tVARIANT)
			$tProp.Variant = $pVARIANT
		Return $pData
	EndFunc

	#cs
	# Get IDispatch property pointer offset from Object property in bytes
	# @internal
	# @param String $sElement struct element name
	# @return Integer offset
	#ce
	Func __AOI_GetPtrOffset($sElement)
		Local Static $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties;BYTE lock;PTR __destructor")
		Local Static $iObject = Int(DllStructGetPtr($tObject, "Object"), @AutoItX64?2:1)
		Return DllStructGetPtr($tObject, $sElement) - $iObject
	EndFunc

	#cs
	# Gets single value from pointer
	# @internal
	# @param Pointer $pPointer the struct element pointer
	# @param String $sElementType struct element specification
	# @return Mixed the value of the pointer of type specified by $sElementType
	#ce
	Func __AOI_GetPtrValue($pPointer, $sElementType)
		Return DllStructGetData(DllStructCreate($sElementType, $pPointer), 1)
	EndFunc


#EndRegion


#region _Memory
	;==================================================================================
	; AutoIt Version:	3.1.127 (beta)
	; Language:			English
	; Platform:			All Windows
	; Author:			Nomad
	; Requirements:		These functions will only work with beta.
	;==================================================================================
	; Credits:	wOuter - These functions are based on his original _Mem() functions.
	;			But they are easier to comprehend and more reliable.  These
	;			functions are in no way a direct copy of his functions.  His
	;			functions only provided a foundation from which these evolved.
	;==================================================================================
	;
	; Functions:
	;
	;==================================================================================
	; Function:			_MemoryOpen($iv_Pid[, $iv_DesiredAccess[, $iv_InheritHandle]])
	; Description:		Opens a process and enables all possible access rights to the
	;					process.  The Process ID of the process is used to specify which
	;					process to open.  You must call this function before calling
	;					_MemoryClose(), _MemoryRead(), or _MemoryWrite().
	; Parameter(s):		$iv_Pid - The Process ID of the program you want to open.
	;					$iv_DesiredAccess - (optional) Set to 0x1F0FFF by default, which
	;										enables all possible access rights to the
	;										process specified by the Process ID.
	;					$iv_InheritHandle - (optional) If this value is TRUE, all processes
	;										created by this process will inherit the access
	;										handle.  Set to 1 (TRUE) by default.  Set to 0
	;										if you want it FALSE.
	; Requirement(s):	None.
	; Return Value(s): 	On Success - Returns an array containing the Dll handle and an
	;								 open handle to the specified process.
	;					On Failure - Returns 0
	;					@Error - 0 = No error.
	;							 1 = Invalid $iv_Pid.
	;							 2 = Failed to open Kernel32.dll.
	;							 3 = Failed to open the specified process.
	; Author(s):		Nomad
	; Note(s):
	;==================================================================================
	Func _MemoryOpen($iv_Pid, $iv_DesiredAccess = 0x1F0FFF, $iv_InheritHandle = 1)
		Logging($log_prefix_InfoMemory & "MemoryOpen " & '$iv_Pid :' & $iv_Pid & '; $iv_DesiredAccess :' & $iv_DesiredAccess & '; $iv_InheritHandle :' & $iv_InheritHandle,   @ScriptLineNumber)
		;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryOpen & '$iv_Pid :' & $iv_Pid & '; $iv_DesiredAccess :' & $iv_DesiredAccess & '; $iv_InheritHandle :' & $iv_InheritHandle,   @ScriptLineNumber)
		If Not ProcessExists($iv_Pid) Then
		Logging($log_prefix_ErrMemory & "MemoryOpen Invalid $iv_Pid", @ScriptLineNumber)
			SetError(1)
			Return 0
		EndIf
		
		Local $ah_Handle[2] = [DllOpen('kernel32.dll')]
		
		If @Error Then
			Logging($log_prefix_ErrMemory & "MemoryOpen Failed to open Kernel32.dll", @ScriptLineNumber)
			SetError(2)
			Return 0
		EndIf
		
		Local $av_OpenProcess = DllCall($ah_Handle[0], 'int', 'OpenProcess', 'int', $iv_DesiredAccess, 'int', $iv_InheritHandle, 'int', $iv_Pid)
		
		If @Error Then
			Logging($log_prefix_ErrMemory & "MemoryOpen Failed to open the specified process", @ScriptLineNumber)
			DllClose($ah_Handle[0])
			SetError(3)
			Return 0
		EndIf
		
		$ah_Handle[1] = $av_OpenProcess[0]
		
		Return $ah_Handle
		
	EndFunc

	;==================================================================================
	; Function:			_MemoryRead($iv_Address, $ah_Handle[, $sv_Type])
	; Description:		Reads the value located in the memory address specified.
	; Parameter(s):		$iv_Address - The memory address you want to read from. It must
	;								  be in hex format (0x00000000).
	;					$ah_Handle - An array containing the Dll handle and the handle
	;								 of the open process as returned by _MemoryOpen().
	;					$sv_Type - (optional) The "Type" of value you intend to read.
	;								This is set to 'dword'(32bit(4byte) signed integer)
	;								by default.  See the help file for DllStructCreate
	;								for all types.  An example: If you want to read a
	;								word that is 15 characters in length, you would use
	;								'char[16]' since a 'char' is 8 bits (1 byte) in size.
	; Return Value(s):	On Success - Returns the value located at the specified address.
	;					On Failure - Returns 0
	;					@Error - 0 = No error.
	;							 1 = Invalid $ah_Handle.
	;							 2 = $sv_Type was not a string.
	;							 3 = $sv_Type is an unknown data type.
	;							 4 = Failed to allocate the memory needed for the DllStructure.
	;							 5 = Error allocating memory for $sv_Type.
	;							 6 = Failed to read from the specified process.
	; Author(s):		Nomad
	; Note(s):			Values returned are in Decimal format, unless specified as a
	;					'char' type, then they are returned in ASCII format.  Also note
	;					that size ('char[size]') for all 'char' types should be 1
	;					greater than the actual size.
	;==================================================================================
	Func _MemoryRead($iv_Address, $ah_Handle, $sv_Type = 'dword')
		If _IsHex($iv_Address) Then
			Logging($log_prefix_InfoMemory & "MemoryRead " & '$iv_Address :' & $iv_Address & '; $ah_Handle :' & $ah_Handle & '; $sv_Type :' & $sv_Type,   @ScriptLineNumber)
			;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryRead & '$iv_Address :' & $iv_Address '; $ah_Handle :' & $ah_Handle '; $sv_Type :' & $sv_Type,   @ScriptLineNumber)
			If Not IsArray($ah_Handle) Then
				Logging($log_prefix_ErrMemory & "MemoryRead Invalid $ah_Handle", @ScriptLineNumber)
				SetError(1)
				Return 0
			EndIf
			
			Local $v_Buffer = DllStructCreate($sv_Type)
			
			If @Error Then
				Switch @error + 1
					
					Case 2
						Logging($log_prefix_ErrMemory & "MemoryRead $sv_Type was not a string", @ScriptLineNumber)
					Case 3
						Logging($log_prefix_ErrMemory & "MemoryRead $sv_Type is an unknown data type", @ScriptLineNumber)
					Case 4	
						Logging($log_prefix_ErrMemory & "MemoryRead Failed to allocate the memory needed for the DllStructure", @ScriptLineNumber)
					Case 5
						Logging($log_prefix_ErrMemory & "MemoryRead Error allocating memory for $sv_Type", @ScriptLineNumber)
				EndSwitch
				SetError(@Error + 1)
				Return 0
			EndIf
			
			DllCall($ah_Handle[0], 'int', 'ReadProcessMemory', 'int', $ah_Handle[1], 'int', $iv_Address, 'ptr', DllStructGetPtr($v_Buffer), 'int', DllStructGetSize($v_Buffer), 'int', '')
			
			If Not @Error Then
				Local $v_Value = DllStructGetData($v_Buffer, 1)
				Return $v_Value
			Else
				Logging($log_prefix_ErrMemory & "MemoryRead Failed to read from the specified process", @ScriptLineNumber)
				SetError(6)
				Return 0
			EndIf
		Else			
			Logging($log_prefix_ErrMemory & "MemoryRead $iv_Address không phải là Hex (Address)", @ScriptLineNumber)
			Return SetError(7, 1, 0)
		EndIf
	EndFunc
	
	
	Func _IsHex($sString)
		Return StringRegExp($sString, '^0x[[:xdigit:]]+$') > 0 ; Or the class [0-9A-Za-z].
	EndFunc   ;==>_IsHex
	;==================================================================================
	; Function:			_MemoryWrite($iv_Address, $ah_Handle, $v_Data[, $sv_Type])
	; Description:		Writes data to the specified memory address.
	; Parameter(s):		$iv_Address - The memory address which you want to write to.
	;								  It must be in hex format (0x00000000).
	;					$ah_Handle - An array containing the Dll handle and the handle
	;								 of the open process as returned by _MemoryOpen().
	;					$v_Data - The data to be written.
	;					$sv_Type - (optional) The "Type" of value you intend to write.
	;								This is set to 'dword'(32bit(4byte) signed integer)
	;								by default.  See the help file for DllStructCreate
	;								for all types.  An example: If you want to write a
	;								word that is 15 characters in length, you would use
	;								'char[16]' since a 'char' is 8 bits (1 byte) in size.
	; Return Value(s):	On Success - Returns 1
	;					On Failure - Returns 0
	;					@Error - 0 = No error.
	;							 1 = Invalid $ah_Handle.
	;							 2 = $sv_Type was not a string.
	;							 3 = $sv_Type is an unknown data type.
	;							 4 = Failed to allocate the memory needed for the DllStructure.
	;							 5 = Error allocating memory for $sv_Type.
	;							 6 = $v_Data is not in the proper format to be used with the
	;								 "Type" selected for $sv_Type, or it is out of range.
	;							 7 = Failed to write to the specified process.
	; Author(s):		Nomad
	; Note(s):			Values sent must be in Decimal format, unless specified as a
	;					'char' type, then they must be in ASCII format.  Also note
	;					that size ('char[size]') for all 'char' types should be 1
	;					greater than the actual size.
	;==================================================================================
	Func _MemoryWrite($iv_Address, $ah_Handle, $v_Data, $sv_Type = 'dword')
		If _IsHex($iv_Address) Then
			Logging($log_prefix_InfoMemory & "MemoryWrite " & '$iv_Address :' & $iv_Address & '; $ah_Handle :' & $ah_Handle & '; $v_Data :' & $v_Data & '; $sv_Type :' & $sv_Type,   @ScriptLineNumber)
			;Logging($Lang_log.Log.Text.Prefix.Info.Memory & $Lang_log.Log.Text.MemoryWrite & '$iv_Address :' & $iv_Address '; $ah_Handle :' & $ah_Handle '; $v_Data :' & $v_Data & '; $sv_Type :' & $sv_Type,   @ScriptLineNumber)
			If Not IsArray($ah_Handle) Then
				Logging($log_prefix_ErrMemory & "MemoryWrite Invalid $ah_Handle", @ScriptLineNumber)
				SetError(1)
				Return 0
			EndIf
			
			Local $v_Buffer = DllStructCreate($sv_Type)
			
			If @Error Then
				Switch @error + 1
					
					Case 2
						Logging($log_prefix_ErrMemory & "MemoryWrite $sv_Type was not a string", @ScriptLineNumber)
					Case 3
						Logging($log_prefix_ErrMemory & "MemoryWrite $sv_Type is an unknown data type", @ScriptLineNumber)
					Case 4	
						Logging($log_prefix_ErrMemory & "MemoryWrite Failed to allocate the memory needed for the DllStructure", @ScriptLineNumber)
					Case 5
						Logging($log_prefix_ErrMemory & "MemoryWrite Error allocating memory for $sv_Type", @ScriptLineNumber)
				EndSwitch
				SetError(@Error + 1)
				Return 0
			Else
				DllStructSetData($v_Buffer, 1, $v_Data)
				If @Error Then
					Logging($log_prefix_ErrMemory & "MemoryWrite $v_Data is not in the proper format to be used with the 'Type' selected for $sv_Type, or it is out of range", @ScriptLineNumber)
					SetError(6)
					Return 0
				EndIf
			EndIf
			
			DllCall($ah_Handle[0], 'int', 'WriteProcessMemory', 'int', $ah_Handle[1], 'int', $iv_Address, 'ptr', DllStructGetPtr($v_Buffer), 'int', DllStructGetSize($v_Buffer), 'int', '')
			
			If Not @Error Then
				Return 1
			Else
				Logging($log_prefix_ErrMemory & "MemoryWrite Failed to write to the specified process", @ScriptLineNumber)
				SetError(7)
				Return 0
			EndIf
		Else			
			Logging($log_prefix_ErrMemory & "MemoryWrite $iv_Address không phải là Hex (Address)", @ScriptLineNumber)
			Return SetError(7, 0, 0)
		EndIf
	EndFunc

	;==================================================================================
	; Function:			_MemoryClose($ah_Handle)
	; Description:		Closes the process handle opened by using _MemoryOpen().
	; Parameter(s):		$ah_Handle - An array containing the Dll handle and the handle
	;								 of the open process as returned by _MemoryOpen().
	; Return Value(s):	On Success - Returns 1
	;					On Failure - Returns 0
	;					@Error - 0 = No error.
	;							 1 = Invalid $ah_Handle.
	;							 2 = Unable to close the process handle.
	; Author(s):		Nomad
	; Note(s):
	;==================================================================================
	Func _MemoryClose($ah_Handle)
		Logging($log_prefix_InfoMemory & "MemoryClose $ah_Handle :" & $ah_Handle ,   @ScriptLineNumber)
		If Not IsArray($ah_Handle) Then
			Logging($log_prefix_ErrMemory & "MemoryClose Invalid $ah_Handle" ,   @ScriptLineNumber)
			SetError(1)
			Return 0
		EndIf
		
		DllCall($ah_Handle[0], 'int', 'CloseHandle', 'int', $ah_Handle[1])
		If Not @Error Then
			DllClose($ah_Handle[0])
			Return 1
		Else
			Logging($log_prefix_ErrMemory & "MemoryClose Unable to close the process handle" ,   @ScriptLineNumber)
			DllClose($ah_Handle[0])
			SetError(2)
			Return 0
		EndIf
		
	EndFunc

	;==================================================================================
	; Function:			SetPrivilege( $privilege, $bEnable )
	; Description:		Enables (or disables) the $privilege on the current process
	;                   (Probably) requires administrator privileges to run
	;
	; Author(s):		Larry (from autoitscript.com's Forum)
	; Notes(s):
	; http://www.autoitscript.com/forum/index.php?s=&showtopic=31248&view=findpost&p=223999
	;==================================================================================

	Func SetPrivilege( $privilege, $bEnable )
		
		Const $TOKEN_ADJUST_PRIVILEGES = 0x0020
		Const $TOKEN_QUERY = 0x0008
		Const $SE_PRIVILEGE_ENABLED = 0x0002
		Local $hToken, $SP_auxret, $SP_ret, $hCurrProcess, $nTokens, $nTokenIndex, $priv
		$nTokens = 1
		$LUID = DLLStructCreate("dword;int")
		If IsArray($privilege) Then    $nTokens = UBound($privilege)
		$TOKEN_PRIVILEGES = DLLStructCreate("dword;dword[" & (3 * $nTokens) & "]")
		$NEWTOKEN_PRIVILEGES = DLLStructCreate("dword;dword[" & (3 * $nTokens) & "]")
		$hCurrProcess = DLLCall("kernel32.dll","hwnd","GetCurrentProcess")
		$SP_auxret = DLLCall("advapi32.dll","int","OpenProcessToken","hwnd",$hCurrProcess[0],   _
				"int",BitOR($TOKEN_ADJUST_PRIVILEGES,$TOKEN_QUERY),"int_ptr",0)
		If $SP_auxret[0] Then
			$hToken = $SP_auxret[3]
			DLLStructSetData($TOKEN_PRIVILEGES,1,1)
			$nTokenIndex = 1
			While $nTokenIndex <= $nTokens
				If IsArray($privilege) Then
					$priv = $privilege[$nTokenIndex-1]
				Else
					$priv = $privilege
				EndIf
				$ret = DLLCall("advapi32.dll","int","LookupPrivilegeValue","str","","str",$priv,   _
						"ptr",DLLStructGetPtr($LUID))
				If $ret[0] Then
					If $bEnable Then
						DLLStructSetData($TOKEN_PRIVILEGES,2,$SE_PRIVILEGE_ENABLED,(3 * $nTokenIndex))
					Else
						DLLStructSetData($TOKEN_PRIVILEGES,2,0,(3 * $nTokenIndex))
					EndIf
					DLLStructSetData($TOKEN_PRIVILEGES,2,DllStructGetData($LUID,1),(3 * ($nTokenIndex-1)) + 1)
					DLLStructSetData($TOKEN_PRIVILEGES,2,DllStructGetData($LUID,2),(3 * ($nTokenIndex-1)) + 2)
					DLLStructSetData($LUID,1,0)
					DLLStructSetData($LUID,2,0)
				EndIf
				$nTokenIndex += 1
			WEnd
			$ret = DLLCall("advapi32.dll","int","AdjustTokenPrivileges","hwnd",$hToken,"int",0,   _
					"ptr",DllStructGetPtr($TOKEN_PRIVILEGES),"int",DllStructGetSize($NEWTOKEN_PRIVILEGES),   _
					"ptr",DllStructGetPtr($NEWTOKEN_PRIVILEGES),"int_ptr",0)
			$f = DLLCall("kernel32.dll","int","GetLastError")
		EndIf
		$NEWTOKEN_PRIVILEGES=0
		$TOKEN_PRIVILEGES=0
		$LUID=0
		If $SP_auxret[0] = 0 Then Return 0
		$SP_auxret = DLLCall("kernel32.dll","int","CloseHandle","hwnd",$hToken)
		If Not $ret[0] And Not $SP_auxret[0] Then Return 0
		return $ret[0]
	EndFunc   ;==>SetPrivilege

#endregion

#Region OCR
	Func OCR($filedir,$link =  "https://api.ocr.space/parse/image", $apikey ='09c854369688957' )
		
		$sAddress = $link; the address of the target (https or http, makes no difference - handled automatically)
		$sFileToUpload = $filedir
		If Not $sFileToUpload Then Exit 5 ; check if the file is selected and exit if not
		$sForm = _
				'<form action="' & $sAddress & '" method="post" enctype="multipart/form-data">' & _
				' <input type="file" name="fileUploaded"/>' & _
				' <input type="text" name="apikey" value="'& $apikey &'"/>' & _
				'</form>'
		$hOpenU = _WinHttpOpen()
		$hConnectU = $sForm ; will pass form as string so this is for coding correctness because $hConnect goes in byref
		$sHTML = _WinHttpSimpleFormFill($hConnectU, $hOpenU, _
				Default, _
				"name:fileUploaded", $sFileToUpload)
		$iErr = @error
		_WinHttpCloseHandle($hConnectU)
		_WinHttpCloseHandle($hOpenU)
		If $iErr Then 
			;ConsoleWrite("Error number = " & $iErr & @CRLF)
			Logging("[OCR][Error][JSON] Error number = " & $iErr)
			SetError($iErr, 1, 0)
		Else
			$str = StringReplace($sHTML, '"', "")
			Logging("[OCR][Info][JSON] " & $sHTML)
			$res = StringRegExp($str, '(?:ParsedText):([^\{,}]+)', 3)
			$ErrorMessage= StringRegExp($str, '(?:ErrorMessage):([^\{,}]+)', 3)
			;_ArrayDisplay($ErrorMessage)
			If IsArray($res) and Not IsArray($ErrorMessage) Then 
				;_ArrayDisplay($res)
				Return $res[0]
			Else 
				Return SetError(3, 1, $ErrorMessage[0])
			EndIf
			

		EndIf
		SetError(1, 1, 0)
	EndFunc
#EndRegion OCR
