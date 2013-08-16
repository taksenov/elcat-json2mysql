#cs ------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         taksenov@gmail.com

 Script Function:
	Скрипт для обработки файлов JSON и внесения данных из них в БД электронного каталога.

#ce ------------------------------------------

; Script Start - Add your code below here

#include <GUIConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Include <JSMN.au3>


;объявление переменных------------------------
Global $var = ""
Global $NumbersOfArray
Global $ArrayTableIndex
Dim $1DArray, $TmpArray, $UIDArray
Dim Const $AppDBPrefix = "elcat_"

;Форма GUI------------------------------------
Opt("GUIOnEventMode", 1) 													;включить режим онэвент

;создаем форму и ее управляющие элементы------
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Перенос JSON-данных в БД elcat", 957, 420, 186, 120, 1, $WS_EX_TOPMOST)
$lstSprTable = GUICtrlCreateList("", 16, 104, 185, 162)
GUICtrlSetData(-1, "author|bbk1|bbk2|content|genre|language|library|librarytown|publisher|series|shelfplace")     ;Названия таблиц
$lbFileName = GUICtrlCreateLabel("<файл JSON>", 16, 16, 186, 24)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
$btnJsonOpen = GUICtrlCreateButton("Выбрать", 216, 16, 171, 24)                            ;кнопка выбора файла
GUICtrlSetOnEvent($btnJsonOpen, "_FileOpen")                                               ;_FileOpen()
$btnSprTablEx = GUICtrlCreateButton("Обработать спр-таблицу", 216, 104, 171, 25)
GUICtrlSetOnEvent($btnSprTablEx, "_ProcessSprTable")                                       ;кнопка для создания SQL для справочной таблицы
$inpMainTable2 = GUICtrlCreateInput("book", 16, 288, 185, 21)                              ;_ProcessSprTable()
$btnMainTablEx = GUICtrlCreateButton("Обработать основную таблицу", 216, 288, 171, 25)     ;кнопка для создания SQL для основной таблицы
GUICtrlSetOnEvent($btnMainTablEx, "_ProcessMainTable")                                     ;_ProcessMainTable()
$edtSqlText = GUICtrlCreateEdit("", 408, 40, 537, 305, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $WS_VSCROLL, $WS_HSCROLL, $ES_MULTILINE))
GUICtrlSetData(-1, "")																	   ;текст кода добавляется в процессе работы функции
GUICtrlSetFont(-1, 10, 800, 0, "Courier New")
$Label2 = GUICtrlCreateLabel("Лог работы программы:", 408, 16, 136, 17)
;$inpMysqlConnect = GUICtrlCreateInput("", 16, 360, 297, 21)
;$lbMySqlConnect = GUICtrlCreateLabel("Связка с MySQL", 16, 336, 88, 17)
$btnSqlRun = GUICtrlCreateButton("Очистить лог", 408, 360, 107, 25)                        ;кнопка запуска sql
GUICtrlSetOnEvent($btnSqlRun, "_LogClear")                                                 ;_LogClear
$inpMainTable1 = GUICtrlCreateInput("book", 16, 64, 185, 21)
$btnMainTablArray = GUICtrlCreateButton("Выгрузить осн.табл в массив", 216, 64, 171, 25)   ;кнопка для создания двумерного массива из основной таблицы
GUICtrlSetOnEvent($btnMainTablArray, "_ProcessMainTablArray")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
;---------------------------------------------

;ставим выполнение окна на бесконечный цикл---
While 1
	Sleep(1000)   ; Just idle around
WEnd
;---------------------------------------------

;события формы--------------------------------
Func CLOSEClicked()			    ;закрыть форму
	Exit
EndFunc
;---------------------------------------------

;события элементов формы----------------------
;нажатие на кнопку выбора файла.
Func _FileOpen()

	$message = "Выберите JSON-файл для обработки"
	$var = FileOpenDialog($message, "::{450D8FBA-AD25-11D0-98A8-0800361B1103}", "json (*.json)", 1 + 4 )	;папка "Мои документы"

	If @error Then
		$var = ""
		MsgBox(4096,"","Файл не выбран")
	Else
		$var = StringReplace($var, "|", @CRLF)
	EndIf

	GUICtrlSetData($edtSqlText, "Выбран файл: " & $var & @CRLF, 1)

EndFunc			;_FileOpen()

;Функции обработки JSON файлов----------------
;названия таблиц из заголовков json-файлов:
;вначале названия таблицы добавляем "elcat_"
;"author" "bbk1" "bbk2" "genre" "series" "content" "publisher" "language" "shelfplace" "library" "librarytown"
Func _ProcessSprTable()                                                       ;Нажатие на кнопку обработки справочной таблицы
	If (GUICtrlRead($lstSprTable) = "bbk1") Or (GUICtrlRead($lstSprTable) = "bbk2") Or (GUICtrlRead($lstSprTable) = "publisher") Then
			_ProcessSprTableEx3(GUICtrlRead($lstSprTable))                    ;обработка 3-х аргументной таблицы
	Else
			_ProcessSprTableEx2(GUICtrlRead($lstSprTable))		              ;обработка 2-х аргументной таблицы
	EndIf
EndFunc				;_ProcessSprTable()

Func _ProcessMainTablArray()                                                  ;создать двумерный массив для таблицы elcat_book из JSON-файла

	Local $Json1 = FileRead($var)
	Local $Data1 = Jsmn_Decode($Json1)

	Local $Json2 = Jsmn_Encode($Data1, $JSMN_UNESCAPED_UNICODE)
	Local $Data2 = Jsmn_Decode($Json2)

	Local $objJson = Jsmn_Decode($Json1)
	$1DArray = Jsmn_ObjTo2DArray($objJson)

	;определим двумерный массив
	$i = UBound($1DArray)
	Global $GrandArray2D[$i][20]					                        ;двумерный массив
	$NumbersOfArray = $i

	;набиваем двумерный и одмерный массивы данными
	$i = 0
	While $i <> 3434245														;дикая переменная (гвоздь)
		Jsmn_Get($objJson, '[' & $i & ']')
		If @error Then
			$i = 3434245
		Else
			$GrandArray2D[$i][0] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["bookname"]')
			$GrandArray2D[$i][1] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["inventory"]')
			$GrandArray2D[$i][2] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["volume"]')
			$GrandArray2D[$i][3] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["author"]')
			$GrandArray2D[$i][4] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["ISBN"]')
			$GrandArray2D[$i][5] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["ISBN10"]')
			$GrandArray2D[$i][6] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["bbk1"]')
			$GrandArray2D[$i][7] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["bbk2"]')
			$GrandArray2D[$i][8] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["genre"]')
			$GrandArray2D[$i][9] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["series"]')
			$GrandArray2D[$i][10] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["content"]')
			$GrandArray2D[$i][11] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["publisher"]')
			$GrandArray2D[$i][12] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["year"]')
			$GrandArray2D[$i][13] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["pages"]')
			$GrandArray2D[$i][14] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["language"]')
			$GrandArray2D[$i][15] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["price"]')
			$GrandArray2D[$i][16] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["shelfplace"]')
			$GrandArray2D[$i][17] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["library"]')
			$GrandArray2D[$i][18] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["librarytown"]')
			$GrandArray2D[$i][19] = uuid()															;крайний столбец массива содержит в себе айдишники книг
			$i = $i + 1
		EndIf
	WEnd

	;если в столбце с ценой есть 'null', то нужно заменить его на '0.00'
		For $i = 0 To $NumbersOfArray - 1
			If $GrandArray2D[$i][15] = "null" Then
				$GrandArray2D[$i][15] = "0.00"
			EndIf
		Next

	;если в столбце с инвентарным номером номера нет, то заменяем это значекние на 'null'
		For $i = 0 To $NumbersOfArray - 1
			If $GrandArray2D[$i][1] = "" Then
				$GrandArray2D[$i][1] = "null"
			EndIf
		Next

;	MsgBox(4096,"","Данные помещены в массив. (MainTable-->GrandArray)")
	GUICtrlSetData($edtSqlText, "Данные помещены в массив. (MainTable-->GrandArray)" & @CRLF, 1)

EndFunc	            ;_ProcessMainTablArray()

Func _ProcessSprTableEx2($SprTableName)                                        ;обработка справочной таблицы с двумя аргументами

	Local $Json1 = FileRead($var)
	Local $Data1 = Jsmn_Decode($Json1)

	Local $Json2 = Jsmn_Encode($Data1, $JSMN_UNESCAPED_UNICODE)
	Local $Data2 = Jsmn_Decode($Json2)

	Local $objJson = Jsmn_Decode($Json1)
	$1DArray = Jsmn_ObjTo2DArray($objJson)

	;определим два массива с размерностью по кол-ву записей из JSON-файла
	$BigContArray = $1DArray												;Основной контент
	$TmpContArray = $1DArray												;Временный контент

	$i = 0
	While $i <> 3434245														;дикая переменная (гвоздь)
		Jsmn_Get($objJson, '[' & $i & ']')
		If @error Then
			$i = 3434245
		Else
			$BigContArray[$i] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["' & $SprTableName & '"]')
			$i = $i + 1
		EndIf
	WEnd

	$TmpContArray = $BigContArray
	_ArraySort($TmpContArray)												;Основной массив сортировать нельзя

    ;Меняем "null" на "без автора" в большом массиве с контентом
	If $SprTableName = "author" Then
		For $i = 0 To UBound($BigContArray) - 1
			If $BigContArray[$i] = "null" Then
				$BigContArray[$i] = "без автора"
			EndIf
		Next
	EndIf

	;Удаляем дубликаты из временного массива с контентом
	For $i = UBound($TmpContArray) - 1 To 1 Step - 1
		If $TmpContArray[$i] = $TmpContArray[$i - 1] Then _ArrayDelete($TmpContArray, $i)
	Next

	;Меняем "null" на "без автора" во временном, маленьком массиве
	If $SprTableName = "author" Then
		For $i = 0 To UBound($TmpContArray) - 1
			If $TmpContArray[$i] = "null" Then
				$TmpContArray[$i] = "без автора"
			EndIf
		Next
	EndIf

	;Создаем малый массив с UID
	$UIDTmpContArray = $TmpContArray
	For $i = 0 To UBound($TmpContArray) - 1
		$UIDTmpContArray[$i] = uuid()
	Next
;	_ArrayDisplay($TmpContArray, "Итоговый массив малый(временный)")
;	_ArrayDisplay($UIDTmpContArray, "Малый UID-массив")

	Select
		Case $SprTableName = "author"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 3
		Case $SprTableName = "genre"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 8
		Case $SprTableName = "series"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 9
		Case $SprTableName = "content"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 10
		Case $SprTableName = "language"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 14
		Case $SprTableName = "shelfplace"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 16
		Case $SprTableName = "library"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name) VALUES"
			$ArrayTableIndex = 17
		Case $SprTableName = "librarytown"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, town) VALUES"
			$ArrayTableIndex = 18
	EndSelect

	;Создать UID-массив для большого не отсортированного массива
	$i = 0
	$i1 = 0
	$UIDBigContArray = $BigContArray
	For $i = 0 To UBound($TmpContArray) - 1
		For $i1 = 0 To UBound($BigContArray) - 1
			If $TmpContArray[$i] = $BigContArray[$i1] Then
				$UIDBigContArray[$i1] = $UIDTmpContArray[$i]
				$GrandArray2D[$i1][$ArrayTableIndex] = $UIDTmpContArray[$i]
			EndIf
		Next
	Next

	;Вывод SQL-запроса на внесение данных в справочную таблицу в MySQL
	;создаем временный файл, в него будет помещаться код sql запроса
	$fileOut = FileOpen("temp.txt" , 10)
	If $fileOut = -1 Then
		MsgBox(0, "Error", "Временный файл не создан.")
		Exit
	EndIf

	FileWriteLine($fileOut, $1stString & @CRLF)
;	GUICtrlSetData($edtSqlText, $1stString & @CRLF, 1)
	$i = 0
	$i1 = UBound($TmpContArray) - 1
	For $i = 0 To UBound($TmpContArray) - 1
		If $i = $i1 Then
			FileWriteLine($fileOut, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray[$i] & "');" & @CRLF)     ;последняя строка SQL-edit
;			GUICtrlSetData($edtSqlText, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray[$i] & "');" & @CRLF, 1)     ;последняя строка SQL-edit
		Else
			FileWriteLine($fileOut, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray[$i] & "')," & @CRLF)     ;рядовая строка SQL-edit
;			GUICtrlSetData($edtSqlText, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray[$i] & "')," & @CRLF, 1)     ;рядовая строка SQL-edit
		EndIf
	Next

	;Если обрабатывается таблица author , то создать дополнительный файл для промежуточной связной таблицы, т.к. есть связь многие ко многим
	If $SprTableName = "author" Then
		;создаем временный файл, в него будет помещаться код sql запроса
		$fileOut2 = FileOpen("temp2.txt" , 10)
		If $fileOut2 = -1 Then
			MsgBox(0, "Error", "Временный файл не создан.")
			Exit
		EndIf
		$1stString = "INSERT INTO " & $AppDBPrefix & "book_author_id ( book_id, author_id ) VALUES"
		FileWriteLine($fileOut2, $1stString & @CRLF)
		$i = 0
		$i1 = UBound($BigContArray) - 1
		For $i = 0 To UBound($BigContArray) - 1
			If $i = $i1 Then
				FileWriteLine($fileOut2, "(" & $GrandArray2D[$i][19] & ", " & $UIDBigContArray[$i] & ");" & @CRLF)     ;последняя строка SQL-edit
			Else
				FileWriteLine($fileOut2, "(" & $GrandArray2D[$i][19] & ", " & $UIDBigContArray[$i] & ")," & @CRLF)     ;рядовая строка SQL-edit
			EndIf
		Next
		FileCopy("temp2.txt", $var & "_elcat_book_author_id_SQL.txt", 8)
		FileDelete("temp2.txt")
		FileClose($fileOut2)
	EndIf

	;вывод итогового результата SQL-кода в файл для итоговой таблицы book
	FileCopy("temp.txt", $var & "_" & $AppDBPrefix & $SprTableName & "_SQL.txt", 8)
	FileDelete("temp.txt")
	$FileVar = $var & "_" & $AppDBPrefix & $SprTableName & "_SQL.txt"
	ShellExecute( $FileVar )
	GUICtrlSetData($edtSqlText, "Обработана таблица: " & $SprTableName & @CRLF, 1)
	FileClose($fileOut)

EndFunc				;_ProcessSprTableEx2()

Func _ProcessSprTableEx3($SprTableName)                                        ;обработка справочной таблицы с тремя аргументами

	Local $Json1 = FileRead($var)
	Local $Data1 = Jsmn_Decode($Json1)

	Local $Json2 = Jsmn_Encode($Data1, $JSMN_UNESCAPED_UNICODE)
	Local $Data2 = Jsmn_Decode($Json2)

	Local $objJson = Jsmn_Decode($Json1)
	$1DArray = Jsmn_ObjTo2DArray($objJson)

	;определим двумерный массив
	$i = UBound($1DArray)
	Local $BigContArray2D[$i][2]									        ;двумерный массив, Основной контент
	Local $TmpContArray2D[$i][2]                                            ;двумерный временный массив
	;определим два	массива с размерностью по кол-ву записей из JSON-файла
	$BigContArray = $1DArray												;одномерный массив, Основной контент
	$TmpContArray = $1DArray												;Временный контент

	Select
		Case $SprTableName = "bbk1"
			$SprTableName2 = "bbk1desc"                                      ;в БД это поле называется description1
		Case $SprTableName = "bbk2"
			$SprTableName2 = "bbk2desc"                                      ;в БД это поле называется description2
		Case $SprTableName = "publisher"
			$SprTableName2 = "city"
	EndSelect

	;набиваем двумерный и одмерный массивы данными
	$i = 0
	While $i <> 3434245														;дикая переменная (гвоздь)
		Jsmn_Get($objJson, '[' & $i & ']')
		If @error Then
			$i = 3434245
		Else
			$BigContArray2D[$i][0] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["' & $SprTableName & '"]')
			$BigContArray2D[$i][1] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["' & $SprTableName2 & '"]')
			$BigContArray[$i] = Jsmn_Get_ShowResult($objJson, '[' & $i & ']["' & $SprTableName & '"]')
			$i = $i + 1
		EndIf
	WEnd

	$TmpContArray = $BigContArray
	$TmpContArray2D = $BigContArray2D

	_ArraySort($TmpContArray)													;Основной массив сортировать нельзя
	_ArraySort($TmpContArray2D)

	;Удаляем дубликаты из временного массива с контентом
	For $i = UBound($TmpContArray) - 1 To 1 Step - 1
		If $TmpContArray2D[$i][0] = $TmpContArray2D[$i - 1][0] Then _ArrayDelete($TmpContArray2D, $i)
		If $TmpContArray[$i] = $TmpContArray[$i - 1] Then _ArrayDelete($TmpContArray, $i)
	Next

	;Создаем малый массив с UID
	$UIDTmpContArray = $TmpContArray
	For $i = 0 To UBound($TmpContArray) - 1
		$UIDTmpContArray[$i] = uuid()
	Next

	Select
		Case $SprTableName = "bbk1"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, code1, description1) VALUES"
			$ArrayTableIndex = 6
		Case $SprTableName = "bbk2"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, code2, description2) VALUES"
			$ArrayTableIndex = 7
		Case $SprTableName = "publisher"
			$1stString = "INSERT INTO " & $AppDBPrefix & $SprTableName & " (" & $SprTableName & "_id, name, city) VALUES"
			$ArrayTableIndex = 11
	EndSelect

	;Создать UID-массив для большого не отсортированного массива
	$i = 0
	$i1 = 0
	$UIDBigContArray = $BigContArray
	For $i = 0 To UBound($TmpContArray) - 1
		For $i1 = 0 To UBound($BigContArray) - 1
			If $TmpContArray[$i] = $BigContArray[$i1] Then
				$UIDBigContArray[$i1] = $UIDTmpContArray[$i]
				$GrandArray2D[$i1][$ArrayTableIndex] = $UIDTmpContArray[$i]
			EndIf
		Next
	Next

	;Вывод SQL-запроса на внесение данных в справочную таблицу в MySQL
		;создаем временный файл, в него будет помещаться код sql запроса
	$fileOut = FileOpen("temp.txt" , 10)
	If $fileOut = -1 Then
		MsgBox(0, "Error", "Временный файл не создан.")
		Exit
	EndIf

	FileWriteLine($fileOut, $1stString & @CRLF)
;	GUICtrlSetData($edtSqlText, $1stString & @CRLF, 1)
	$i = 0
	$i1 = UBound($TmpContArray) - 1
	For $i = 0 To UBound($TmpContArray) - 1
		If $i = $i1 Then
			FileWriteLine($fileOut, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray2D[$i][0] & "' , " & "'" & $TmpContArray2D[$i][1] & "');" & @CRLF)     ;последняя строка SQL-edit
;			GUICtrlSetData($edtSqlText, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray2D[$i][0] & "' , " & "'" & $TmpContArray2D[$i][1] & "');" & @CRLF, 1)     ;последняя строка SQL-edit
		Else
			FileWriteLine($fileOut, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray2D[$i][0] & "' , " & "'" & $TmpContArray2D[$i][1] & "')," & @CRLF)     ;рядовая строка SQL-edit
;			GUICtrlSetData($edtSqlText, "(" & $UIDTmpContArray[$i] & ", " & "'" & $TmpContArray2D[$i][0] & "' , " & "'" & $TmpContArray2D[$i][1] & "')," & @CRLF, 1)     ;рядовая строка SQL-edit
		EndIf
	Next

	;вывод итогового результата SQL-кода в файл для итоговой таблицы book
	FileCopy("temp.txt", $var & "_" & $AppDBPrefix & $SprTableName & "_SQL.txt", 8)
	FileDelete("temp.txt")
	$FileVar = $var & "_" & $AppDBPrefix & $SprTableName & "_SQL.txt"
	ShellExecute( $FileVar )

	GUICtrlSetData($edtSqlText, "Обработана таблица: " & $SprTableName & @CRLF, 1)
	FileClose($fileOut)

EndFunc				;_ProcessSprTableEx3()

Func _ProcessMainTable()                                                       ;обработка основной таблицы

	;создаем временный файл, в него будет помещаться код sql запроса
	$fileOut = FileOpen("temp.txt" , 10)
	If $fileOut = -1 Then
		MsgBox(0, "Error", "Временный файл не создан.")
		Exit
	EndIf

	;Вывод SQL-запроса на внесение данных в основную таблицу (elcat_book) в MySQL
	$1stString = "INSERT INTO " & $AppDBPrefix & "book (book_id, name, inventory, volume, bookauthor_id, isbn, isbn10, bbk1_id, bbk2_id, genre_id, series_id, content_id, publisher_id, year, pages, language_id, price, shelfplace_id, library_id, librarytown_id) VALUES"
	FileWriteLine($fileOut, $1stString & @CRLF)
;	GUICtrlSetData($edtSqlText, $1stString & @CRLF, 1)
	$i = 0
	$i1 = $NumbersOfArray - 1
	For $i = 0 To $NumbersOfArray - 1
		If $i = $i1 Then
			;                             book_id                         name                           inventory                      volume                         author_id                      isbn                            isbn10                         bbk1_id                       bbk2_id                       genre_id                      series_id                     content_id                     publisher_id                    year                             pages                           language_id                     price                           shelfplace_id                  library_id                     librarytown_id
			FileWriteLine($fileOut, "(" & $GrandArray2D[$i][19] & ", '" & $GrandArray2D[$i][0] & "', " & $GrandArray2D[$i][1] & ", '" & $GrandArray2D[$i][2] & "', " & $GrandArray2D[$i][3] & ", '" & $GrandArray2D[$i][4] & "', '" & $GrandArray2D[$i][5] & "', " & $GrandArray2D[$i][6] & ", " & $GrandArray2D[$i][7] & ", " & $GrandArray2D[$i][8] & ", " & $GrandArray2D[$i][9] & ", " & $GrandArray2D[$i][10] & ", " & $GrandArray2D[$i][11] & ", '" & $GrandArray2D[$i][12] & "', '" & $GrandArray2D[$i][13] & "', " & $GrandArray2D[$i][14] & ", '" & $GrandArray2D[$i][15] & "', " & $GrandArray2D[$i][16] & ", " & $GrandArray2D[$i][17] & ", " & $GrandArray2D[$i][18] & ");" & @CRLF)     ;последняя строка SQL-edit
;			GUICtrlSetData($edtSqlText, "('" & $GrandArray2D[$i][0] & "', " & $GrandArray2D[$i][1] & ", '" & $GrandArray2D[$i][2] & "', " & $GrandArray2D[$i][3] & ", '" & $GrandArray2D[$i][4] & "', '" & $GrandArray2D[$i][5] & "', " & $GrandArray2D[$i][6] & ", " & $GrandArray2D[$i][7] & ", " & $GrandArray2D[$i][8] & ", " & $GrandArray2D[$i][9] & ", " & $GrandArray2D[$i][10] & ", " & $GrandArray2D[$i][11] & ", '" & $GrandArray2D[$i][12] & "', '" & $GrandArray2D[$i][13] & "', " & $GrandArray2D[$i][14] & ", '" & $GrandArray2D[$i][15] & "', " & $GrandArray2D[$i][16] & ", " & $GrandArray2D[$i][17] & ", " & $GrandArray2D[$i][18] & "');" & @CRLF, 1)     ;последняя строка SQL-edit
		Else
			;                             book_id                         name                           inventory                      volume                         author_id                      isbn                            isbn10                         bbk1_id                       bbk2_id                       genre_id                      series_id                     content_id                     publisher_id                    year                             pages                           language_id                     price                           shelfplace_id                  library_id                     librarytown_id
			FileWriteLine($fileOut, "(" & $GrandArray2D[$i][19] & ", '" & $GrandArray2D[$i][0] & "', " & $GrandArray2D[$i][1] & ", '" & $GrandArray2D[$i][2] & "', " & $GrandArray2D[$i][3] & ", '" & $GrandArray2D[$i][4] & "', '" & $GrandArray2D[$i][5] & "', " & $GrandArray2D[$i][6] & ", " & $GrandArray2D[$i][7] & ", " & $GrandArray2D[$i][8] & ", " & $GrandArray2D[$i][9] & ", " & $GrandArray2D[$i][10] & ", " & $GrandArray2D[$i][11] & ", '" & $GrandArray2D[$i][12] & "', '" & $GrandArray2D[$i][13] & "', " & $GrandArray2D[$i][14] & ", '" & $GrandArray2D[$i][15] & "', " & $GrandArray2D[$i][16] & ", " & $GrandArray2D[$i][17] & ", " & $GrandArray2D[$i][18] & ")," & @CRLF)     ;последняя строка SQL-edit
;			GUICtrlSetData($edtSqlText, "('" & $GrandArray2D[$i][0] & "', " & $GrandArray2D[$i][1] & ", '" & $GrandArray2D[$i][2] & "', " & $GrandArray2D[$i][3] & ", '" & $GrandArray2D[$i][4] & "', '" & $GrandArray2D[$i][5] & "', " & $GrandArray2D[$i][6] & ", " & $GrandArray2D[$i][7] & ", " & $GrandArray2D[$i][8] & ", " & $GrandArray2D[$i][9] & ", " & $GrandArray2D[$i][10] & ", " & $GrandArray2D[$i][11] & ", '" & $GrandArray2D[$i][12] & "', '" & $GrandArray2D[$i][13] & "', " & $GrandArray2D[$i][14] & ", '" & $GrandArray2D[$i][15] & "', " & $GrandArray2D[$i][16] & ", " & $GrandArray2D[$i][17] & ", " & $GrandArray2D[$i][18] & "')," & @CRLF, 1)     ;последняя строка SQL-edit
		EndIf
	Next

	;вывод итогового результата SQL-кода в файл для итоговой таблицы book
	FileCopy("temp.txt", $var & "_book_SQL.txt", 8)
	FileDelete("temp.txt")
	$FileVar = $var & "_book_SQL.txt"
	ShellExecute( $FileVar )

	FileClose($var)
	FileClose($FileVar)
	FileClose($fileOut)

	GUICtrlSetData($edtSqlText, "Обработана основная таблица: book" & @CRLF, 1)
	GUICtrlSetData($edtSqlText, "---------------------------------" & @CRLF, 1)

EndFunc             ;_ProcessMainTable

Func _LogClear()                                                               ;Очистка окна с логом работы программы
	GUICtrlDelete($edtSqlText)
	$edtSqlText = GUICtrlCreateEdit("", 408, 40, 537, 305, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $WS_VSCROLL, $WS_HSCROLL, $ES_MULTILINE))
EndFunc             ;_LogClear()
;---------------------------------------------

;Локальные функции----------------------------------------------------------------------------------------------------------
Func Jsmn_Get_ShowResult($Var, $Key)                                     ;Показывает в консоли данные по конкретному ключу
    Local $Ret = Jsmn_Get($Var, $Key)
    If @Error Then
        Switch @error
            Case 1
                ConsoleWrite("Error 1: key not exists" & @LF)
            Case 2
                ConsoleWrite("Error 2: syntax error" & @LF)
        EndSwitch

    Else
;        ConsoleWrite($Key & " => " & VarGetType($Ret) & ": " & $Ret & @LF)
        ;ConsoleWrite( $Ret & @LF)
		Return $Ret
    EndIf
EndFunc     ;Jsmn_Get_ShowResult()

Func Jsmn_Get($Var, $Key)											     ;Вытаскивает объект по конкретному ключу
    If Not $Key Then Return $Var

    Local $Match = StringRegExp($Key, "(^\[([^\]]+)\])", 3)
    If IsArray($Match) Then
        Local $Index = Jsmn_Decode($Match[1])
        $Key = StringTrimLeft($Key, StringLen($Match[0]))

        If IsString($Index) And Jsmn_IsObject($Var) And Jsmn_ObjExists($Var, $Index) Then
            Local $Ret = Jsmn_Get(Jsmn_ObjGet($Var, $Index), $Key)
            Return SetError(@Error, 0, $Ret)

        ElseIf IsNumber($Index) And IsArray($Var) And $Index >= 0 And $Index < UBound($Var) Then
            Local $Ret = Jsmn_Get($Var[$Index], $Key)
            Return SetError(@Error, 0, $Ret)

        Else
            Return SetError(1, 0, "")

        EndIf
    EndIf
    Return SetError(2, 0, "")
EndFunc		;Jsmn_Get()

Func uuid()											       	             ;Генерирует UID, для уникальных значений ключей
    Return StringFormat('%02d%02d%02d%04d%04d', _
            Random(10, 99), _
			Random(0, 99), _
            Random(0, 99), _
            BitOR(Random(100, 5555), 1000), _
            BitOR(Random(5555, 9999), 1000) _
        )
EndFunc		;uuid()
