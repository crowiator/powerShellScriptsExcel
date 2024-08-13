$api_base_url = 'http://localhost:8081'
$api_key = 'odt_G4lnWF8hrxYcbPVQjdCGPJ5vfPsahlHN'
$output_file = 'C:\Users\User\Desktop\components.xlsx'
$headers = @{
    'accept' = 'application/json'
    'X-Api-Key' = $api_key
}

try
{
    Write-Host "Starting Excel application..."
    $my_excel = New-Object -ComObject excel.application
    $my_excel.visible = $false
    $my_workbook = $my_excel.workbooks.add()
    $sheet_1 = $my_workbook.worksheets.item(1)
    $sheet_1.name = "EPSS-CVSS"

    Write-Host "Setting up Excel headers..."
    $sheet_1.cells.item(1, 1) = 'NAME'
    $sheet_1.cells.item(1, 2) = 'VERSION'
    $sheet_1.cells.item(1, 3) = 'CRITICAL'
    $sheet_1.cells.item(1, 4) = 'HIGH'
    $sheet_1.cells.item(1, 5) = 'MEDIUM'
    $sheet_1.cells.item(1, 6) = 'LOW'
    $sheet_1.cells.item(1, 7) = 'UNASSIGNED'
    $sheet_1.cells.item(1, 8) = 'FINDINGS'

    $line = 2

    Write-Host "Retrieving projects from API..."
    $response = Invoke-WebRequest -Uri ($api_base_url + '/api/v1/project') -Method Get -Headers $headers
    $projects = $response.Content | ConvertFrom-Json

    foreach ($project in $projects)
    {
        if ($project.name -eq "SDA ZA")
        {
            Write-Host "Found project: $($project.name). Retrieving components..."

            $page = 1
            $pageSize = 100
            $hasMorePages = $true

            while ($hasMorePages)
            {
                $url = $api_base_url + "/api/v1/component/project/" + $project.uuid + "?pageNumber=$page&pageSize=$pageSize"
                $response = Invoke-WebRequest -Uri $url -Method Get -Headers $headers
                $components = $response.Content | ConvertFrom-Json

                if ($components.Count -eq 0)
                {
                    $hasMorePages = $false
                }
                else
                {
                    foreach ($comp in $components)
                    {
                        #Write-Host "Processing component: $($comp.name)"
                         #Write-Host "Processing component: $($comp.name)"
                        $sheet_1.cells.item($line, 1).NumberFormat = "@"
                        $sheet_1.cells.item($line, 1) = $comp.name
                        $sheet_1.cells.item($line, 2).NumberFormat = "@"
                        $sheet_1.cells.item($line, 2) = $comp.version
                        $sheet_1.cells.item($line, 3).NumberFormat = "@"
                        $sheet_1.cells.item($line, 3) = $comp.metrics.critical
                        $sheet_1.cells.item($line, 4).NumberFormat = "@"
                        $sheet_1.cells.item($line, 4) = $comp.metrics.high
                        $sheet_1.cells.item($line, 5).NumberFormat = "@"
                        $sheet_1.cells.item($line, 5) = $comp.metrics.medium
                        $sheet_1.cells.item($line, 6).NumberFormat = "@"
                        $sheet_1.cells.item($line, 6) = $comp.metrics.low
                        $sheet_1.cells.item($line, 7).NumberFormat = "@"
                        $sheet_1.cells.item($line, 7) = $comp.metrics.unassigned
                        $findingsTotal = $comp.metrics.findingsTotal
                        $sheet_1.cells.item($line, 8).NumberFormat = "@"
                        $sheet_1.cells.item($line, 8) = $findingsTotal

                        $line++
                    }

                    $page++
                }
            }
        }    
    }

    Write-Host "Saving Excel file..."
    $my_workbook.SaveAs($output_file)
    $my_excel.Quit()
    Write-Host "Process completed successfully."
}
catch
{
    Write-Host 'Error: ' $_.Exception.Message
    Write-Host $_.ScriptStackTrace

    if ($my_excel -ne $null) {
        $my_excel.Quit()
    }
}
finally
{
    if ($my_excel -ne $null) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($my_excel) | Out-Null
    }
}
