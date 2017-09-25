'********************************************************
'  FileCopyList
'  Date: 25.09.2017
'  v1.00
'  Authors:
'  Vladimir Svishch (IndianBiker)  <mail:5693031@gmail.com>
'  https://github.com/BikerIndian/VBS
'********************************************************
 
Dim FSO, arr_Line, File, Txt, my_arr(), objF,path
Dim pathIn


Set FSO = CreateObject("Scripting.FileSystemObject")

' Здесь указать путь к папке с файлами
pathIn = "C:\Files\"
 
'PachROOT
Set F = FSO.GetFile(Wscript.ScriptFullName)
PachROOT = FSO.GetParentFolderName(F)
PachROOT = PachROOT & "\file\"
 
 
 
Dim FileLogIn, FileLogInLogNameIn

FileLogInLogNameIn = PachROOT & ".\logIn.txt"

Dim FileLogOut, FileLogInLogNameOut
FileLogInLogNameOut = PachROOT & ".\logOut.txt"

'Проверка наличия аргументов


intCount = WScript.Arguments.Count

if 0=WScript.Arguments.Count Then
wscript.echo "1 - изменить параметр pathIn в скрипте" &  vbNewLine & "2 - перетащить мышкой файл на скрипт "
WScript.Quit 1 'Выход из скрипта
End if



'Если нет  каталога то создаем.
If FSO.FolderExists(PachROOT) Then

   Else
   FSO.CreateFolder PachROOT
 
End if
' ------------ FileLogInLogNameIn ------------------------ 
' >>

' Если файл не создан, то создать и закрыть

if Not FSO.FileExists(FileLogInLogNameIn) Then
	Set FileLogIn = FSO.CreateTextFile(FileLogInLogNameIn, true)
	FileLogIn.Close
	Else
	Set FileLogIn = FSO.CreateTextFile(FileLogInLogNameIn, true)
	FileLogIn.Close
End if


' Открытие файла и запись в него
Set FileLogIn = FSO.OpenTextFile(FileLogInLogNameIn, 8, True)

' <<<<  ------------ FileLogInLogNameIn ------------------------


' >>>> ------------ FileLogInLogNameOut ------------------------ 
'
' Если файл не создан, то создать и закрыть
if Not FSO.FileExists(FileLogInLogNameOut) Then
	Set FileLogOut = FSO.CreateTextFile(FileLogInLogNameOut, true)
	FileLogOut.Close
	Else
	Set FileLogOut = FSO.CreateTextFile(FileLogInLogNameOut, true)
	FileLogOut.Close
End if

' Открытие файла и запись в него
Set FileLogOut = FSO.OpenTextFile(FileLogInLogNameOut, 8, True)

' <<<<  ------------ FileLogInLogNameOut ------------------------


Set File = FSO.OpenTextFile(WScript.Arguments(0), 1)
 
 
arr_Line = 0
 
Do Until File.AtEndOfStream ' пока наступит конец файла
	Redim Preserve my_arr(arr_Line)
	my_arr(arr_Line) = File.ReadLine ' Считываем строку из файла
	
	'myTrim = Trim(Cell)
	'pathInFile = pathIn & my_arr(arr_Line)	' Путь файла
	pathInFile = pathIn & Trim(my_arr(arr_Line))	' Путь файла
	
	
	If (FSO.FileExists(pathInFile)) Then
		'MsgBox "Файл есть!"
		'Set fso = CreateObject("Scripting.FileSystemObject")
		fso.CopyFile pathInFile, PachROOT, True

		' Вставка строки с переносом на новую.
		FileLogIn.WriteLine(my_arr(arr_Line))
		
		Else
		' Вставка строки с переносом на новую.
		FileLogOut.WriteLine(arr_Line & " - " & my_arr(arr_Line))
		'MsgBox pathInFile  & " >>> " & PachROOT

	
	End if
	
	
	'MsgBox pathInFile  & " >>> " & PachROOT
	arr_Line = arr_Line + 1
Loop
 
' Закрытие файла.
File.Close 

' Закрытие файла.
FileLogIn.Close

' Закрытие файла.
FileLogOut.Close
