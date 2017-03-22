# Load the required module to parse XLSX
Import-Module ImportExcel

# Initiate and capture web request
$uri = "http://bernoullihealth.com/bernoulli-validated-windows-updates/"
$WebRequest = (Invoke-WebRequest $uri)

# Parse page for links
$links = ($WebRequest.ParsedHtml.getElementsByTagName("a") | where {($_.innertext -like "*Approved*2008*") -or ($_.innertext -like "*Approved*2012*") }) 


foreach ($link in $links) {
    $file = ($link | select -ExpandProperty nameProp) 
    $data = New-Object -TypeName psobject 
    Invoke-WebRequest $link.href -OutFile $file
    $data = (Import-Excel $file | where {$_.description -eq "Security"} | ConvertTo-Csv -NoTypeInformation )
    $data = (ConvertFrom-Csv $data | select -ExpandProperty hotfixid)
    $Results = @() 

    foreach ($row in $data) {

    #Find and replace for KB number to Q number
    $QNum = ($row.Replace("KB","Q"))

    # Configure output, add results to array.  
    $Output = ($file.ToString()).Replace("xlsx","txt")
    $Results += $QNum
    }

# Write report and notify user
$Results | Out-File $Output
Write-Host "Q Numbers written to $Output" -BackgroundColor DarkGreen

# Cleanup the downloaded temp files
Remove-Item $file
}