#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=border.ico
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=This app fix old tajik fonts to new.
#AutoIt3Wrapper_Res_Description=Document fixer by Odilshoh
#AutoIt3Wrapper_Res_Fileversion=1.0.0.9
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=© by Odilshoh
#AutoIt3Wrapper_Add_Constants=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#NoTrayIcon
#include <Word.au3>
#include <WordConstants.au3>

$oldChr = "Љ,Ќ,Ѓ,Њ,Ў,Ї"
$newChr = "Ҷ,Қ,Ғ,Ҳ,Ӯ,Ӣ"

$OSU = StringSplit(StringUpper($oldChr), ",")
$NSU = StringSplit(StringUpper($newChr), ",")
$OSL = StringSplit(StringLower($oldChr), ",")
$NSL = StringSplit(StringLower($newChr), ",")

$UI = GUICreate("Document fixer (© Odilshoh 2021)", 400, 260, Default, Default, Default, 0x00000010); $WS_EX_DROPACCEPT
GUISetBkColor(0xFFFFFF, $ui)
$title = GUICtrlCreateLabel("Ҳарф-ислоҳкунак", 0, 0, 400, 40, 0x0200+0x01)
GUICtrlSetFont(-1, 18, Default, Default, "SegoeUI", 5)

;$dropZone = GUICtrlCreateLabel("Ҳуҷҷатро партоед...", 400/2 - 150/2, 260/2 - 150/2, 150, 150, 0x0200+0x01)
;GUICtrlSetBkColor(-1, 0x00AA00)

If @Compiled Then
	$iconPath = @ScriptFullPath
Else
	$iconPath = @ScriptDir & "\border.ico"
EndIf

$dropZone = GUICtrlCreateIcon($iconPath, 99, 400/2 - 128/2, 260/2 - 128/2, 128, 128)
GUICtrlSetCursor(-1, 0)
GUICtrlSetState(-1, 8); $GUI_DROPACCEPTED

$path = GUICtrlCreateLabel("• Барномаи мазкур ҳафрҳои номаълумро ислоҳ мекунад!", 0, 260-20, 400, 20, 0x0200)
GUICtrlSetFont(-1, 10, Default, Default, "Consolas", 5)

GUISetState(@SW_SHOW, $ui)

While True
	$msg = GUIGetMsg()
	Switch $msg
		Case -3
			Exit
		Case $dropZone
			$fod = FileOpenDialog("Интихоби файл", @WorkingDir, "Хуччатхо (*.docx;*.doc)", Default, $ui)
			If Not @error Then
				GUICtrlSetData($path, $fod)
				GUICtrlSetColor($path, 0x00BB00)
				_Exeption($fod)
			Else
				GUICtrlSetData($path, "Хуччатро нодуруст интихоб кардед!")
				GUICtrlSetColor($path, 0xBB0000)
			EndIf
		Case -13 ;$GUI_EVENT_DROPPED
			If @GUI_DropId = $dropZone Then
				GUICtrlSetData($path, @GUI_DragFile)
				$sp = StringSplit(StringRight(@GUI_DragFile, 5), ".")
				If $sp[0] > 1 Then
					If StringInStr($sp[2], "doc") Or StringInStr($sp[2], "docx") Then
						GUICtrlSetColor($path, 0x00BB00)
						_Exeption(@GUI_DragFile)
					Else
						GUICtrlSetColor($path, 0xBB0000)
					EndIf
				EndIf
			EndIf

	EndSwitch
WEnd

Func _Exeption($Doc_Path)

	Local $oWord = _Word_Create()
	Local $oDoc = _Word_DocOpen($oWord, $Doc_Path, Default, Default)

	For $i = 1 To $OSU[0]
		_Word_DocFindReplace($oDoc, $OSU[$i], $NSU[$i], $WdReplaceAll, Default, True)
		_Word_DocFindReplace($oDoc, $OSL[$i], $NSL[$i], $WdReplaceAll, Default, True)
	Next
EndFunc