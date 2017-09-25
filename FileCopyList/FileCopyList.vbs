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

' ����� ������� ���� � ����� � �������
pathIn = "C:\Files\"
 
'PachROOT
Set F = FSO.GetFile(Wscript.ScriptFullName)
PachROOT = FSO.GetParentFolderName(F)
PachROOT = PachROOT & "\file\"
 
 
 
Dim FileLogIn, FileLogInLogNameIn

FileLogInLogNameIn = PachROOT & ".\logIn.txt"

Dim FileLogOut, FileLogInLogNameOut
FileLogInLogNameOut = PachROOT & ".\logOut.txt"

'�������� ������� ����������


intCount = WScript.Arguments.Count

if 0=WScript.Arguments.Count Then
wscript.echo "1 - �������� �������� pathIn � �������" &  vbNewLine & "2 - ���������� ������ ���� �� ������ "
WScript.Quit 1 '����� �� �������
End if



'���� ���  �������� �� �������.
If FSO.FolderExists(PachROOT) Then

   Else
   FSO.CreateFolder PachROOT
 
End if
' ------------ FileLogInLogNameIn ------------------------ 
' >>

' ���� ���� �� ������, �� ������� � �������

if Not FSO.FileExists(FileLogInLogNameIn) Then
	Set FileLogIn = FSO.CreateTextFile(FileLogInLogNameIn, true)
	FileLogIn.Close
	Else
	Set FileLogIn = FSO.CreateTextFile(FileLogInLogNameIn, true)
	FileLogIn.Close
End if


' �������� ����� � ������ � ����
Set FileLogIn = FSO.OpenTextFile(FileLogInLogNameIn, 8, True)

' <<<<  ------------ FileLogInLogNameIn ------------------------


' >>>> ------------ FileLogInLogNameOut ------------------------ 
'
' ���� ���� �� ������, �� ������� � �������
if Not FSO.FileExists(FileLogInLogNameOut) Then
	Set FileLogOut = FSO.CreateTextFile(FileLogInLogNameOut, true)
	FileLogOut.Close
	Else
	Set FileLogOut = FSO.CreateTextFile(FileLogInLogNameOut, true)
	FileLogOut.Close
End if

' �������� ����� � ������ � ����
Set FileLogOut = FSO.OpenTextFile(FileLogInLogNameOut, 8, True)

' <<<<  ------------ FileLogInLogNameOut ------------------------


Set File = FSO.OpenTextFile(WScript.Arguments(0), 1)
 
 
arr_Line = 0
 
Do Until File.AtEndOfStream ' ���� �������� ����� �����
	Redim Preserve my_arr(arr_Line)
	my_arr(arr_Line) = File.ReadLine ' ��������� ������ �� �����
	
	'myTrim = Trim(Cell)
	'pathInFile = pathIn & my_arr(arr_Line)	' ���� �����
	pathInFile = pathIn & Trim(my_arr(arr_Line))	' ���� �����
	
	
	If (FSO.FileExists(pathInFile)) Then
		'MsgBox "���� ����!"
		'Set fso = CreateObject("Scripting.FileSystemObject")
		fso.CopyFile pathInFile, PachROOT, True

		' ������� ������ � ��������� �� �����.
		FileLogIn.WriteLine(my_arr(arr_Line))
		
		Else
		' ������� ������ � ��������� �� �����.
		FileLogOut.WriteLine(arr_Line & " - " & my_arr(arr_Line))
		'MsgBox pathInFile  & " >>> " & PachROOT

	
	End if
	
	
	'MsgBox pathInFile  & " >>> " & PachROOT
	arr_Line = arr_Line + 1
Loop
 
' �������� �����.
File.Close 

' �������� �����.
FileLogIn.Close

' �������� �����.
FileLogOut.Close
