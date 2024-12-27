# James Lazo 011529541

# Try block
Try {
    # Set pwd to location
    Set-Location -Path $PSScriptRoot

    # Remove deprecated module
    if (Get-Module sqlps) { Remove-Module sqlps }

    # Get correct module
    Import-Module SqlServer

    # Set initialization variables
    $sqlInstance = ".\SQLEXPRESS"
    $myDatabase = "ClientDB"
    $schema = "dbo"
    $myTable = "Client_A_Contacts"

    # Create SQL server and database objects
    $sqlServObj = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlInstance
    $databaseObj = Get-SqlDatabase -ServerInstance $sqlInstance -Name $myDatabase -ErrorAction SilentlyContinue

    # Detect database, delete if exists
    if ($databaseObj) {
        Write-Host "$($myDatabase) detected, deleting..."
        $sqlServObj.KillAllProcesses($myDatabase)
        $databaseObj.UserAccess = "Single"
        $databaseObj.Drop()
    }

    # OK to proceed
    else {
        Write-Host "Okay to create $myDatabase"
    }

    # Import data, create data structures and insert data
    Write-Host "Importing data and inserting into data structures"
    Import-Csv -Path .\NewClientData.csv | Write-SqlTableData -ServerInstance $sqlInstance -DatabaseName $myDatabase -TableName $myTable -SchemaName $schema -Force

    # Export results
    Invoke-Sqlcmd -Database ClientDB -ServerInstance $sqlInstance -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SQLResults.txt
}

# Catch block
Catch {
    Write-Host "Exception"
}