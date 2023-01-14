; Random List Spinner © 2017–2023 T.D. Stoneheart. All rights reserved.
; Licensed under MIT License.

Global $SpinGUI ; Forward declaration; a customizable function uses this variable!

#Region Customization
Global $gift[][2] = [ _ ; [Option String, Quantity]
	["Alice", 1], _
	["Bob", 1], _
	["Carol", 1], _
	["David", 1], _
	["Erin", 1], _
	["Frank", 1], _
	["Grace", 1], _
	["Heidi", 1], _
	["Ivan", 1], _
	["Judy", 1]]

Global Const $SpinButtonString = "Spin!"
Global Const $SpinWindowTitle = "Lucky Spin"
Global Const $BackstageViewWindowTitle = "Backstage View"
Global Const $GUIFont = "Tahoma"
Global Const $SpinGUIWidth = 300, $SpinGUIHeight = 580
Global Const $BackstageViewWidth = 300, $BackstageViewHeight = 500

Func GiftAnnouncement($choice)
	MsgBox(0, "", $gift[$choice][0], 0, $SpinGUI)
EndFunc
Func EndingMessage()
	MsgBox(0, "End", "No more options!")
EndFunc
#EndRegion Customization

Opt("TrayAutoPause", 0)
Global $list, $last = 0
For $i = 0 To UBound($gift) - 1
	$list &= "|" & $gift[$i][0]
Next
$SpinGUI = GUICreate($SpinWindowTitle, $SpinGUIWidth + 16, $SpinGUIHeight + 80)
GUISetFont(24, 0, 0, $GUIFont, $SpinGUI)
Global $GiftList = GUICtrlCreateList("", 8, 8, $SpinGUIWidth, $SpinGUIHeight, 0)
GUICtrlSetData($GiftList, $list)
Global $SpinButton = GUICtrlCreateButton($SpinButtonString, 8, $SpinGUIHeight + 16, $SpinGUIWidth, 50, 1)
Global $BackstageView = GUICreate($BackstageViewWindowTitle, $BackstageViewWidth + 16, $BackstageViewHeight + 8)
GUISetFont(16, 0, 0, $GUIFont, $BackstageView)
Global $GiftDetails = GUICtrlCreateList("", 8, 8, $BackstageViewWidth, $BackstageViewHeight, 0)
GUICtrlSetData($GiftDetails, BackstageGiftList())
GUISetState(@SW_SHOW, $BackstageView)
GUISetState(@SW_SHOW, $SpinGUI)
EndingCheck()

While 1
	$msg = GUIGetMsg(1)
	; Only closing the backstage view can quit the script!
	If $msg[1] = $BackstageView And $msg[0] = -3 Then Exit
	If $msg[1] = $SpinGUI And $msg[0] = $SpinButton Then Spin()
WEnd

; Renders the next item as selected on GUI, and increase the delay time until the next selection by a value between the multipliers given.
Func SelectItem(ByRef $choice, ByRef $sleep, $min, $max)
	$choice = Mod($choice + 1, UBound($gift))
	ControlCommand($SpinGUI, "", "ListBox1", "SetCurrentSelection", $choice)
	If $min <> $max Then $sleep *= Random($min, $max)
	Sleep($sleep)
EndFunc

Func Spin()
	Local $choice = $last - 1, $sleep = 20, $select = $choice

	; Pre-select the gift
	For $i = 1 To Random(UBound($gift), UBound($gift) * 3, 1)
		Do
			$select = Mod($select + 1, UBound($gift))
		Until $gift[$select][1] > 0
	Next

	; Initial fast spinning
	For $i = 1 To Random(UBound($gift) * 3, UBound($gift) * 6, 1)
		SelectItem($choice, $sleep, 1, 1)
	Next

	; Slow down
	Do
		SelectItem($choice, $sleep, 1.1, 1.3)
	Until $sleep >= 250

	; Further slow down until the selected gift
	Do
		SelectItem($choice, $sleep, 1.0, 1.1)
	Until $sleep >= 500 And $choice = $select

	; Flash the selection!
	For $i = 1 To 4
		ControlCommand($SpinGUI, "", "ListBox1", "SetCurrentSelection", -1)
		Sleep(250)
		ControlCommand($SpinGUI, "", "ListBox1", "SetCurrentSelection", $choice)
		Sleep(250)
	Next

	$last = $choice
	$gift[$choice][1] -= 1
	GUICtrlSetData($GiftDetails, BackstageGiftList())
	ControlCommand($BackstageView, "", "ListBox1", "SetCurrentSelection", $choice)
	GiftAnnouncement($choice)
	EndingCheck()
EndFunc

Func EndingCheck()
	For $i = 0 To UBound($gift) - 1
		If $gift[$i][1] > 0 Then Return
	Next
	EndingMessage()
	Exit
EndFunc

Func BackstageGiftList()
	Local $return = ""
	For $i = 0 To UBound($gift) - 1
		$return &= "|" & $gift[$i][0] & ": " & $gift[$i][1]
	Next
	Return $return
EndFunc
