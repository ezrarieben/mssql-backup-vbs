'###################################################
' 	Author:	Ezra Rieben
' 	Source:	https://github.com/ezrarieben/mssql-backup-vbs/
'	Ver:	1.0.0
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

'Required constant defenitions
const ServerName = "localhost\SQLEXPRESS" 'Name of MS SQL server
const BackupDir = "D:\foo" 'Folder to back up to. Trailing slash not needed
const DBName = "foobar" 'Name of DB to back up

'Optional constant defenitions for auth method: SQL Server Authentication
const DBUser = ""
const DBPass = ""

'Check required vars
IF ServerName = "" THEN wscript.quit
IF BackupDir = "" THEN wscript.quit
IF DBName = "" THEN wscript.quit

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

backupDB DBName

conn.close


SUB backupDB(databaseName)
	fileName = BackupDir & "\" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & databaseName & ".BAK"

	'Start new DB command
	SET cmdbackup = CREATEOBJECT("ADODB.Command")
	cmdbackup.activeconnection = conn
	'Set command to be executed to generate backup file
	cmdbackup.commandtext = "backup database " & databaseName & " to disk='" & fileName & "'"
	'Execute DB command to generate file
	cmdbackup.EXECUTE

END SUB