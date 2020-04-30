# autoit
Mình có tổng hợp các thư viện mình hay dùng lại thành 1 file thư viện chạy theo dạng OOP

Lưu ý: Các code trong bài viết là mình viết thẳng lên file readme chưa chạy thử, có lỗi thì báo mình nhé

Sử dụng AutoIT theo dạng OOP để thao tác dễ dàng hơn cho người mới, bao gồm các chức năng:

1.[Log vào file và console](#log)
 
2.[Tạo Object Custom dễ dàng với thư viện có sẵn](#obj)

3.[Sử dụng Nomadmemory với Object](#memory)

4.[Sử dụng BMPSearch với Object](#bmp)

5.[Load file thành biến trong Object (easy multilang hoặc thay thế ini)](#vari)

6.[CoProc bằng Object](#coproc)

7.[OCR qua API](#ocr)



<a name="log"></a>
1. Log vào file và console:
Sử dụng Logging để log, sẽ sử dụng callback nên không sợ làm chậm chương trình hay không hiển thị ra
vd:
```
#include "ObjectDynamicVari.au3"

_InitLog('main'); khai báo bắt đầu chạy Logging, nếu không có sẽ chỉ ghi log ra console, main ở đây là tên file để phân biệt khi sử dụng multi process
Logging("Đây là câu log"); Thay thế cho consoleWrite ngoài ra thì còn log ra file theo ngày giờ dễ theo dõi
```

<a name="obj"></a>
2. Tạo Object dễ dàng với thư viện có sẵn:
Sử dụng tạo object và thao tác với nó
vd:
```
#include "ObjectDynamicVari.au3"

$exam_obj =  __createObject(); Khai báo biến $exam_obj thành Object
$exam_obj.a = "Hello"; Gán property của Object $exam_obj là a có giá trị là Hello
$exam_obj.__defineGetter("Show", show_msg); Khai báo method Show sẽ chạy function là show_msg

Func show_msg($oThis); bắt buộc phải có argu 
  MsgBox("","",$oThis.parent.a); $oThis.parent.a ở đây sẽ trả về giá trị của $exam_obj.a
EndFunc

MsgBox("","",$exam_obj.a); Sẽ Tạo ra 1 MessageBox là Hello
$exam_obj.Show;Sẽ Tạo ra 1 MessageBox là Hello như trên ( Lưu ý: Show hay Show() đều được và có phân biệt hoa thường)
```

<a name="memory"></a>
3. Sử dụng Nomad memory với Object:
Sử dụng memory với Object để dễ nhìn và dễ sử dụng hơn
vd:
```
#include "ObjectDynamicVari.au3"

$pid = ProcessExists("AutoIT3.exe"); get ProcessID trước
$memory_object = MemoryObject($pid); Khai báo biến $memory_object thành object có PID lấy được ở trên
$memory_object.Open; Open memory để lấy về memory ID; giá trị của memoryID sẽ trả về là $memory_object.MemID
$memory_object.Read("0x40000"); Đọc memory từ địa chỉ 0x40000 theo dạng dword; địa chỉ sẽ lưu ở $memory_object.Address, giá trị trả về $memory_object.RtRead
$memory_object.Read("0x40000","char[2]"); Đọc memory từ địa chỉ 0x40000 dạng char với 2 byte ký tự, giá trị trả về $memory_object.RtRead
;'dword' ;byte, float, dword, char[], wchar[] https://docs.microsoft.com/en-us/windows/win32/winprog/windows-data-types
;trên là các kiểu dữ liệu
$memory_object.Type = "wchar[3]"; Set Type đọc mặc định là wchar[3]
$memory_object.Address = "0x40000" ; Set Address mặc định là 0x40000
$memory_object.Read; Đọc memory từ địa chỉ 0x40000 dạng wchar với 3 byte ký tự, giá trị trả về $memory_object.RtRead
$memory_object.Read("Type = char[2]","Address=0x40000"); Đọc memory từ địa chỉ 0x40000 dạng char với 2 byte ký tự
$memory_object.Write("a"); Write vào memory tại địa chỉ 0x40000 được lưu ở $memory_object.Address và Type được lưu ở $memory_object.Type, giá trị trả về là $memory_object.RtWrite
$memory_object.Close; Đóng MemID
```
;có thể rút gọn như sau:

```
$pid = ProcessExists("AutoIT3.exe"); get ProcessID trước
$memory_object = MemoryObject($pid); Khai báo biến $memory_object thành object có PID lấy được ở trên
$memory_object.Open.Read("0x40000","wchar[3]").Write("a");Mở MemID sau đó đọc giá trị tại 0x40000 với Type wchar[3] rồi write vào giá trị là a
```

<a name="bmp"></a>
4. Sử dụng BMPSearch với Object:
```
#include "ObjectDynamicVari.au3"

$title = "AutoIT"; Tên title cửa sổ cần tìm hình ảnh
$object_bmp = BMPSearch($title); Khai báo object BMPSearch
$fileanh = "anhbmp24bit.bmp"; Ảnh 24bit bmp, lấy bằng cách dùng lightshot chụp và cắt rồi dán vào mspaint, sau đó save lại theo dạng bmp24bit
$object_bmp.Search($fileanh); Tìm kiếm file ảnh theo handle của $title, kết quả trả về $object_bmp.Status theo True/False
; Nếu True sẽ trả thêm tọa độ của ảnh trong ảnh tìm được là $object_bmp.x và $object_bmp.y
$object_bmp.Click;Click vào ảnh tìm được ở trên
;rút gọn
$object_bmp.Search($fileanh).Click;Tìm và Click vào ảnh
```

<a name="vari"></a>
5. Load file thành biến trong Object (easy multilang hoặc thay thế ini):

Bạn muốn làm pmem multi lang nhưng code quá khó ? Sử dụng thử như sau nhé:
Tạo 1 file tên là vi.properties với nội dung như sau:
```
Msg.Text=Đây là Text
Msg.Title=Đây là Title
```

Tạo 1 file tên là en.properties với nội dung như sau:
```
Msg.Text=I'm Text
Msg.Title=Hey Title Here
```

CODE:
```
#include "ObjectDynamicVari.au3"

$en_file = "en.properties"
$vi_file = "vi.properties"
$display = _InitObjectVari($vi_file); Khai báo $display với vi_file
MsgBox("",$display.Msg.Title,$display.Msg.Text)
$display = _InitObjectVari($en_file); Khai báo lại $display với en_file
MsgBox("",$display.Msg.Title,$display.Msg.Text)
```

<a name="coproc"></a>
6:CoProc bằng Object
Làm biếng viết quá, đọc code rồi sử dụng nha mấy bạn, mình mới làm mấy cái căn bản thôi


<a name="ocr"></a>
7:OCR qua API
Sử dụng rất đơn giản, mình lấy api của trang https://api.ocr.space/parse/image
Mọi người vào đăng ký free để lấy apikey rồi thay thế apikey nhé
```
#include "ObjectDynamicVari.au3"
$file_dir =@DesktopDir & "/test.png"
$result = OCR($file_dir)
If Not @error Then
	MsgBox("","SUCCESS",$result)
Else 
	MsgBox("","ERROR",$result)
EndIf

```
