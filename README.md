# mssql-backup-vbs
A VBScript to generate a backup file of a MS SQL database
## Configuring VBScript
 > **NOTE:** By default the script uses the user executing the script to authenticate to the DB server. If you want to authenticate to the DB server using SQL Server Authentication set the options `DBUser` and `DBPass` (Line 47 & 48)
1. Edit `mssql_backup.vbs`
2. Set the required options (Line 32-34)
  ```vbs
  const ServerName = "[name of ms sql server]\[name of sql service]"
  const BackupDir = "[folder path]"
  const DBName = "[db name]"
  ```
3. Execute `mssql_backup.vbs` to generate backup

## Automated backups with task scheduler
1. Place the `start_backup.bat` file into the same directory as the `mssql_backup.vbs` script
2. Edit `start_backup.bat` and make sure the filename defined is the same as the VBScript's filename (Line 2)
```batch
SET fileName= [filename]
```
3. Open Windows Task Scheduler
4. Create a new task that executes the batch file with the corresponding user for the DB
> **NOTE:** The user that executes the backup script doesn't matter if you plan to use SQL Server Authentication

## Automatically copy backup on completion
1. Set the `CopyOnCompleteBool` to `true` (Line 36)
2. Set the path of the folder to copy the backup into to the variable (Line 37)
```vbs
const CopyOnCompleteBool = [true|false]
const CopyOnCompleteDir = "D:\bar"
```
3. Execute `mssql_backup.vbs`

## Automatically move backup on completion
1. Set the `MoveOnCompleteBool` variable to `true` (Line 39)
2. Set the path of the folder to copy the backup into to the variable `MoveOnCompleteDir` (Line 40)
```vbs
const MoveOnCompleteBool = [true|false]
const MoveOnCompleteDir = "D:\bar"
```
3. Execute `mssql_backup.vbs`
> **NOTE:** If both `CopyOnCompleteBool` and `MoveOnCompleteBool` are set to true, the script will make a copy of the .BAK file first and then move it to the desired location.
