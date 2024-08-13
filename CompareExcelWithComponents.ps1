# Import the necessary module for handling Excel files
Import-Module ImportExcel

# Paths to the Excel files
$checkMarxPath = "C:\Users\User\Desktop\CheckMarx.xlsx"
$dependencyTrackPath = "C:\Users\User\Desktop\DepedencyTrack.xlsx"
$outputPath = "C:\Users\User\Desktop\ComparisonOutput.xlsx"

# Load the data from both Excel files
$checkMarxData = Import-Excel -Path $checkMarxPath
$dependencyTrackData = Import-Excel -Path $dependencyTrackPath

# Create a new array to store the comparison results
$comparisonResults = @()

# Loop through each entry in CheckMarx data
foreach ($checkMarxRow in $checkMarxData) {
    $componentName = $checkMarxRow.'NAME'
    $checkMarxVersion = $checkMarxRow.'VERSION'
    $checkMarxFindings = $checkMarxRow.'FINDINGS'
    
    # Find the matching component and version in DependencyTrack data
    $dependencyTrackRow = $dependencyTrackData | Where-Object { $_.'NAME' -eq $componentName -and $_.'VERSION' -eq $checkMarxVersion }
    
    if ($dependencyTrackRow) {
        $dependencyTrackVersion = $dependencyTrackRow.'VERSION'
        $dependencyTrackFindings = $dependencyTrackRow.'FINDINGS'
    } else {
        $dependencyTrackVersion = ""
        $dependencyTrackFindings = ""
    }

    # Add the results to the comparisonResults array
    $comparisonResults += [PSCustomObject]@{
        'Component name' = $componentName
        'Version_CheckMarx' = $checkMarxVersion
        'Version_DependencyTrack' = $dependencyTrackVersion
        'Findings_CheckMarx' = $checkMarxFindings
        'Findings_DependencyTrack' = $dependencyTrackFindings
    }
}

# Now loop through each entry in DependencyTrack data to find any missing components in CheckMarx data
foreach ($dependencyTrackRow in $dependencyTrackData) {
    $componentName = $dependencyTrackRow.'NAME'
    $dependencyTrackVersion = $dependencyTrackRow.'VERSION'
    $dependencyTrackFindings = $dependencyTrackRow.'FINDINGS'
    
    # Check if this component and version is missing in CheckMarx data
    $checkMarxRow = $checkMarxData | Where-Object { $_.'NAME' -eq $componentName -and $_.'VERSION' -eq $dependencyTrackVersion }
    
    if (-not $checkMarxRow) {
        $comparisonResults += [PSCustomObject]@{
            'Component name' = $componentName
            'Version_CheckMarx' = ""
            'Version_DependencyTrack' = $dependencyTrackVersion
            'Findings_CheckMarx' = ""
            'Findings_DependencyTrack' = $dependencyTrackFindings
        }
    }
}

# Export the comparison results to a new Excel file following the template
$comparisonResults | Export-Excel -Path $outputPath -WorkSheetname "Comparison" -AutoSize

# Inform the user that the process is complete
Write-Output "Comparison completed and exported to $outputPath"
