# mssql-backup-vbs
A VBScript to generate a backup file of a MS SQL database
## Configuring VBScript
 > **NOTE:** By default the script uses the user executing the script to authenticate to the DB server. If you want to authenticate to the DB server using SQL Server Authentication set the options `DBUser` and `DBPass` (Line 7-8)
1. Edit `mssql_backup.vbs`
2. Set the required options (Line 2-3)
  ```vbs
  const ServerName = "[name of ms sql server]\[name of sql service]"
  const BackupDir = "[folder path]"
  const DBName = "[db name]"
  ```
3. Execute `mssql_backup.vbs` to generate backup

## Automated Backups with Task Scheduler
1. Place the `start_backup.bat` file into the same directory as the `mssql_backup.vbs` script
2. Edit `start_backup.bat` and make sure the filename defined is the same as the VBScript's filename (Line 2)
```batch
SET fileName= [filename]
```
3. Open Windows Task Scheduler
4. Create a new task that executes the batch file with the corresponding user for the DB
> **NOTE:** The user that executes the backup script doesn't matter if you plan to use SQL Server Authentication
