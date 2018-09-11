'###################################################
' 	Author:	Ezra Rieben
' 	Source:	https://github.com/ezrarieben/mssql-backup-vbs/
'	Ver:	1.1.0
'###################################################
'
'	MIT License
'
'	Copyright (c) 2018 Ezra Rieben
'
'	Permission is hereby granted, free of charge, to any person obtaining a copy
'	of this software and associated documentation files (the "Software"), to deal
'	in the Software without restriction, including without limitation the rights
'	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'	copies of the Software, and to permit persons to whom the Software is
'	furnished to do so, subject to the following conditions:
'
'	The above copyright notice and this permission notice shall be included in all
'	copies or substantial portions of the Software.
'
'	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'	SOFTWARE.
'
'###################################################

'Settings
const ServerName = "localhost\SQLEXPRESS" 'Name of MS SQL server
const BackupDir = "D:\foo" 'Folder to back up to. Trailing slash not needed
const DBName = "foobar" 'Name of DB to back up

const CopyOnCompleteBool = false 'set to true if .BAK needs to be copied after backup has completed
const CopyOnCompleteDir = "D:\bar" 'Path to copy .BAK file to if CopyOnCompleteBool is set to true. Trailing slash not needed

const MoveOnCompleteBool = false 'set to true if .BAK needs to be moved after backup has completed
const MoveOnCompleteDir = "D:\bar" 'Path to move .BAK file to if MoveOnCompleteBool is set to true. Trailing slash not needed

'NOTE: If both CopyOnCompleteBool and MoveOnCompleteBool are set to true,
'	   it will copy the file first and then move it


'Optional constant defenitions for auth method: SQL Server Authentication
const DBUser = ""
const DBPass = ""

'Check required vars
IF ServerName = "" THEN wscript.quit
IF BackupDir = "" THEN wscript.quit
IF DBName = "" THEN wscript.quit


backupFileName = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & DBName & ".BAK"


SET conn = CREATEOBJECT("ADODB.Connection")
SET cmd = CREATEOBJECT("ADODB.Command")
SET rs = CREATEOBJECT("ADODB.RecordSet")


IF DBUser = "" AND DBPass = "" THEN
	'Use windows auth (user that is running script)
	connString = "Provider=SQLOLEDB.1;Data Source=" & ServerName & ";Integrated Security=SSPI;Initial Catalog=" & DBName
ELSE
	'Use set user and pass to authenticate
	connString = "Provider=SQLOLEDB.1;Data Source=" & ServerName & ";Initial Catalog=" & DBName & ";User ID=" & DBUser & ";Password=" & DBPass & ";"
END IF

'Open DB connection according to the specified connection string
conn.open connString

call backupDB()

conn.close

IF CopyOnCompleteBool = true THEN
	call copyOnComplete()
END IF

IF MoveOnCompleteBool = true THEN
	call moveOnComplete()
END IF

SUB backupDB()
	backupFilePath = BackupDir & "\" & backupFileName

	'Start new DB command
	SET cmdbackup = CREATEOBJECT("ADODB.Command")
	cmdbackup.activeconnection = conn
	'Set command to be executed to generate backup file
	cmdbackup.commandtext = "backup database " & DBName & " to disk='" & backupFilePath & "'"
	'Execute DB command to generate file
	cmdbackup.EXECUTE

END SUB

SUB moveOnComplete()
	backupFilePath = BackupDir & "\" & backupFileName
	moveToPath = MoveOnCompleteDir & "\" & backupFileName
	
	'Move the file
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	fileSystem.MoveFile backupFilePath, moveToPath
END SUB

SUB copyOnComplete()
	backupFilePath = BackupDir & "\" & backupFileName
	copyToPath = CopyOnCompleteDir & "\" & backupFileName

	'Copy the file
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	fileSystem.CopyFile backupFilePath, copyToPath
END SUB
