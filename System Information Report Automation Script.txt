# Define output HTML file path
$DesktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "Server_Configuration_Report.html")
$HtmlFilePath = $DesktopPath

# Gather System Information
$hostname = $env:COMPUTERNAME
$ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
$processor = (Get-CimInstance Win32_Processor).Name
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$osArchitecture = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
$domain = (Get-CimInstance Win32_ComputerSystem).Domain

# Get Network Information
$networkTable = "<h2>Network Adapters</h2><table border='1'>
    <tr><th>Adapter Name</th><th>IP Address</th><th>MAC Address</th><th>Link Speed</th></tr>"

$networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | ForEach-Object {
    $ip = (Get-NetIPAddress -InterfaceIndex $_.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
    $networkTable += "<tr><td>$($_.Name)</td><td>$ip</td><td>$($_.MacAddress)</td><td>$($_.LinkSpeed) Mbps</td></tr>"
}
$networkTable += "</table>"

# Gather Disk Information
$diskTable = "<h2>Disk Information</h2><table border='1'>
    <tr><th>Drive</th><th>Total Space</th><th>Free Space</th></tr>"

Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object {
    $diskTable += "<tr><td>$($_.DeviceID)</td><td>{0} GB</td><td>{1} GB</td></tr>" -f 
        [math]::Round($_.Size / 1GB, 2), [math]::Round($_.FreeSpace / 1GB, 2)
}
$diskTable += "</table>"

# Ask user if they want to retrieve SQL Server Information
$includeSQL = Read-Host "Do you want to retrieve SQL Server details? (Yes/No)"
$sqlVersion = "Skipped"
$dbSizeTable = "Skipped"

if ($includeSQL -match "Yes") {
    $sqlInstance = Read-Host "Enter SQL Server Instance Name (e.g., ServerName\InstanceName)"
    $useSQLAuth = Read-Host "Use SQL Authentication? (Yes/No)"
    
    if ($useSQLAuth -match "Yes") {
        $sqlUser = Read-Host "Enter SQL Username"
        $sqlPassword = Read-Host "Enter SQL Password" -AsSecureString
        $sqlPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sqlPassword)
        )
        $connectionString = "Server=$sqlInstance;Database=master;User Id=$sqlUser;Password=$sqlPassword"
    } else {
        $connectionString = "Server=$sqlInstance;Database=master;Integrated Security=True"
    }

    try {
        Write-Host "Attempting to connect to SQL Server: $sqlInstance" -ForegroundColor Yellow
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = $connectionString
        $conn.Open()
        Write-Host "✅ SQL Connection Successful!" -ForegroundColor Green

        # Get SQL Server Version
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT SERVERPROPERTY('ProductVersion') AS Version"
        $sqlVersion = $cmd.ExecuteScalar()

        # Get all Database Names and Sizes
        $cmd.CommandText = @"
        SELECT d.name AS DatabaseName, 
        CAST(SUM(mf.size * 8.0 / 1024 / 1024) AS DECIMAL(10,2)) AS DBSize_GB
        FROM sys.databases d
        JOIN sys.master_files mf ON d.database_id = mf.database_id
        WHERE d.state_desc = 'ONLINE'
        AND d.database_id > 4          
        GROUP BY d.name
        ORDER BY d.name;
"@
        $reader = $cmd.ExecuteReader()

        # Build the SQL Database Size Table
        $dbSizeTable = "<h2>SQL Database Information</h2>
                        <table border='1'>
                        <tr><th>Database Name</th><th>Size (GB)</th></tr>"
        while ($reader.Read()) {
            #$dbSizeTable += "<tr><td>$($reader["DatabaseName"])</td><td>{0} GB</td></tr>" -f [math]::Round($reader["Size_MB"] / 1024, 2)
            $dbSizeTable += "<tr><td>$($reader["DatabaseName"])</td><td>{0} GB</td></tr>" -f $reader["DBSize_GB"]

        }
        $reader.Close()
        $dbSizeTable += "</table>"

        $conn.Close()
    } catch {
        Write-Host "❌ Error: Could not retrieve SQL details. Check instance name, username, or password." -ForegroundColor Red
        Write-Host "🔍 Detailed Error: $_" -ForegroundColor Cyan
        $sqlVersion = "Error"
        $dbSizeTable = "<h2>SQL Database Information</h2><p>Error retrieving SQL details.</p>"
    }
}

# Retrieve SSL Certificates
$sslCerts = Get-ChildItem -Path Cert:\LocalMachine\My, Cert:\LocalMachine\WebHosting | 
            Select-Object Subject, NotAfter, Issuer, PSParentPath

$sslTable = "<h2>SSL Certificate Information</h2><table border='1'>
    <tr><th>Subject Name</th><th>Expiry Date</th><th>Issuer</th><th>Certificate Type</th></tr>"
foreach ($cert in $sslCerts) {
    $certType = if ($cert.PSParentPath -match "WebHosting") { "Web Hosting" } else { "Personal" }
    $sslTable += "<tr><td>$($cert.Subject)</td><td>$($cert.NotAfter.ToString('yyyy-MM-dd'))</td>
                   <td>$($cert.Issuer)</td><td>$certType</td></tr>"
}
$sslTable += "</table>"

# Generate HTML Report
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Server Configuration Report</title>
<!DOCTYPE html>
<html>
<head>
    <title>Server Configuration Report</title>
    <style>
        body {
            font-family: 'Arial', sans-serif;
            background-color: #f4f4f4;
            color: #333;
            margin: 0;
            padding: 20px;
        }
        h1 {
            text-align: center;
            font-size: 28px;
            color: #007acc;
        }
        h2 {
            background-color: #e0e0e0;
            padding: 10px;
            border-radius: 5px;
            color: #333;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background-color: #fff;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        th, td {
            border: 1px solid #ccc;
            padding: 10px;
            text-align: left;
            color: #333;
        }
        th {
            background-color: #007acc;
            color: white;
            font-size: 16px;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .container {
            max-width: 900px;
            margin: auto;
            background: #fff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
    </style>
</body>
</html>

</head>
<body>
    <h1>Server Configuration Report</h1>

    <h2>System Information</h2>
    <table>
        <tr><td>Hostname</td><td>$hostname</td></tr>
        <tr><td>Total Physical Memory (GB)</td><td>$ram</td></tr>
        <tr><td>OS Version</td><td>$os</td></tr>
        <tr><td>OS Architecture</td><td>$osArchitecture</td></tr>
        <tr><td>Processor</td><td>$processor</td></tr>
        <tr><td>Domain</td><td>$domain</td></tr>
    </table>

    $networkTable
    $diskTable
    $sslTable
    $dbSizeTable
</body>
</html>
"@

# Write to HTML file
$html | Out-File -Encoding utf8 -FilePath $HtmlFilePath

# Open the generated HTML file
Start-Process $HtmlFilePath
