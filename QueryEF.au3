#include <Excel.au3>

$oExcel = _Excel_Open()
$FilePath = _Excel_BookOpen($oExcel, @ScriptDir & "\filter.xlsx")
$HitungRow = $oExcel.ActiveSheet.UsedRange.Rows.Count
Global $DelayKetik = 500 ;dalam mili detik


Func AktifWin()
   WinWait("Filter Engine")
   WinActivate("Filter Engine")
EndFunc

Func TulisQuery1()
   Send("{DOWN 17}{TAB 3}")
   For $i = 1 to $HitungRow
	  $BacaRow = _Excel_RangeRead($FilePath, Default, "A" & $i)
	  Sleep($DelayKetik)
	  Send("^a")
	  Send("{DEL}")
	  Sleep($DelayKetik)
	  Send($BacaRow)
	  Sleep($DelayKetik)
	  Send("!o")
	  Sleep($DelayKetik)
	  Send("{TAB 8}")
	  Sleep($DelayKetik)
   Next
EndFunc

Func TulisQuery2()
   Send("{DOWN 7}{TAB 3}")
   For $i = 1 to $HitungRow
	  $BacaRow = _Excel_RangeRead($FilePath, Default, "A" & $i)
	  Sleep($DelayKetik)
	  Send("^a")
	  Send("{DEL}")
	  Sleep($DelayKetik)
	  Send($BacaRow)
	  Sleep($DelayKetik)
	  Send("!o")
	  Sleep($DelayKetik)
	  Send("{TAB 8}")
	  Sleep($DelayKetik)
   Next
EndFunc

Func SimpanQuery()
   Send("!s")
EndFunc

AktifWIN()
TulisQuery2()
;SimpanQuery()