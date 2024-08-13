$api_base_url = 'http://localhost:8081'
$api_key = 'odt_G4lnWF8hrxYcbPVQjdCGPJ5vfPsahlHN'
$output_file = 'C:\Users\User\Desktop\cvss-epss.xlsx'
$cvssMin = 5
$epssMin = 0.5
$headers = @{
    'accept' = 'application/json'
    'X-Api-Key' = $api_key
}
try
{
    
    $response = Invoke-WebRequest -Uri ($api_base_url + '/api/v1/project') -Method Get -Headers $headers
    $projects = $response | ConvertFrom-Json
    Write-Output "Choose project from projects:"
    $array = @()
    foreach($project in $projects){
        $array += $project.name
    }
    foreach ($item in $array) {
        Write-Output $item
    }
    # Prompt the user for input
    $userInput = Read-Host "Please enter name of the project"
    # Display the input back to the user
    Write-Output "You entered: $userInput"

    $severity = ""
    $my_excel = New-Object -ComObject excel.application
    $my_excel.visible = $false
    $my_workbook = $my_excel.workbooks.add()
    $sheet_1 = $my_workbook.worksheets.item(1)
    $sheet_1.name = "EPSS-CVSS"

    $sheet_1.cells.item(1, 1) = 'NAME'
    $sheet_1.cells.item(1, 2) = 'VERSION'
    $sheet_1.cells.item(1, 3) = 'UUID'
    $sheet_1.cells.item(1, 4) = 'VULN-ID'
    $sheet_1.cells.item(1, 5) = 'CVSS'
    $sheet_1.cells.item(1, 6) = 'SEVERITY'
    $sheet_1.cells.item(1, 7) = 'EPSS'
    $sheet_1.cells.item(1, 8) = 'COMPONENT-NAME'
    $sheet_1.cells.item(1, 9) = 'COMPONENT-VERSION'

    $line = 2
    foreach ($project in $projects)
    {
        if($project.name -eq $userInput)
        {

            $response = Invoke-WebRequest -Uri ($api_base_url + '/api/v1/vulnerability/project/' + $project.uuid) -Method Get -Headers $headers
            $vulns = $response | ConvertFrom-Json
            foreach ($vuln in $vulns)
            {
                $cvss = [Float]$vuln.cvssV3BaseScore
                $epss = [Float]$vuln.epssScore
               
                    foreach ($comp in $vuln.components)
                    {

                        if ($vuln.cvssV3BaseScore -gt 0.0 -and $vuln.cvssV3BaseScore -le 3.9 ) {
                               $severity = "Low"
                               $color = 5296274   
                            } elseif ($vuln.cvssV3BaseScore -gt 3.9 -and $vuln.cvssV3BaseScore -le 6.9) {
                                $severity = "Medium"
                                $color = 65535   
                            } elseif ($vuln.cvssV3BaseScore -gt 6.9 -and $vuln.cvssV3BaseScore -le 8.9) {
                                $severity = "High"
                                $color = 255 
                            } elseif ($vuln.cvssV3BaseScore -gt 8.9 -and $vuln.cvssV3BaseScore -le 10.0) {
                                $severity = "Critical"
                                $color = 9109504 
                            }  
                                else {
                                 $severity = "Unasigned"
                                 $color = 12632256
                            }
                        # print into console    
                        #$project.name + ";" + $project.version + ";" + $project.uuid + ";" + $vuln.vulnID + ";" + $vuln.cvssV3BaseScore + ";" + $severity + ";"+ $vuln.epssScore + ";" + $comp.name + ";" + $comp.version
                        Write-Output $color 
                        # Set text format
                        $sheet_1.cells.item($line, 1).NumberFormat = "@"
                        $sheet_1.cells.item($line, 1) = $project.name
                        $sheet_1.cells.item($line, 2).NumberFormat = "@"
                        $sheet_1.cells.item($line, 2) = $project.version
                        $sheet_1.cells.item($line, 3).NumberFormat = "@"
                        $sheet_1.cells.item($line, 3) = $project.uuid
                        $sheet_1.cells.item($line, 4).NumberFormat = "@"
                        $sheet_1.cells.item($line, 4) = $vuln.vulnID
                        $sheet_1.cells.item($line, 5).NumberFormat = "@"
                        $sheet_1.cells.item($line, 5) = $vuln.cvssV3BaseScore
                        $sheet_1.cells.item($line, 6).NumberFormat = "@"
                        $sheet_1.cells.item($line, 6) = $severity
                       # $sheet_1.cells.item($line, 6).Interior.Color = $color
                        $sheet_1.cells.item($line, 7).NumberFormat = "@"
                        $sheet_1.cells.item($line, 7) = $vuln.epssScore
                        $sheet_1.cells.item($line, 8).NumberFormat = "@"
                        $sheet_1.cells.item($line, 8) = $comp.name
                        $sheet_1.cells.item($line, 9).NumberFormat = "@"
                        $sheet_1.cells.item($line, 9) = $comp.version
                        $line++
                    }
            
            }
        }
    }
    $my_workbook.Saveas($output_file)
    $my_excel.Quit()
}
catch
{
    'error: ' + $response
    $_.Exception.Message
    $_.ScriptStackTrace
}
