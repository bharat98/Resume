# Convert Word to PDF
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $docPath = "C:\Users\gurba\OneDrive - University of Maryland\Resume\Prod\Bharat Gurbaxani Resume Nov'24.docx"
    $pdfPath = "C:\Users\gurba\OneDrive - University of Maryland\Resume\Prod\PDF version\Bharat Gurbaxani Resume Nov'24.pdf"

    if (Test-Path $docPath) {
        $doc = $word.Documents.Open($docPath)
        $doc.SaveAs([ref] $pdfPath, [ref] 17)
        $doc.Close()
        Write-Host "PDF created successfully at: $pdfPath"
    } else {
        Write-Host "Word document not found at: $docPath"
        exit
    }
}
catch {
    Write-Host "An error occurred during PDF conversion: $_"
    exit
}
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# Upload PDF to GitHub
try {
    # GitHub API settings
    $owner = "bharat98"
    $repo = "Resume"
    $branch = "main"
    $filePath = "Bharat Gurbaxani resume.pdf"
    $currentDate = Get-Date -Format "dd-MMM-yyyy"
    $olderVersionPath = "Older version/Bharat Gurbaxani resume $currentDate.pdf"
    $commitMessage = "Resume Update: $currentDate"

    # Your GitHub Personal Access Token
    $token = '' # Replace with your actual token or use a secure method to retrieve it

    # Read the PDF file content
    $fileContent = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($pdfPath))

    # Set up the API request headers
    $headers = @{
        Authorization = "token $token"
        Accept = "application/vnd.github.v3+json"
    }

    # Get the current file content and SHA
    $uri = "https://api.github.com/repos/$owner/$repo/contents/$filePath"
    $existingFile = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue

    if ($existingFile) {
        # Move the existing file to "Older version" folder
        $moveBody = @{
            message = "Move existing file to Older version folder"
            content = $existingFile.content
            branch = $branch
            path = $olderVersionPath
        } | ConvertTo-Json

        Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$olderVersionPath" -Method Put -Headers $headers -Body $moveBody

        # Delete the original file
        $deleteBody = @{
            message = "Delete original file"
            sha = $existingFile.sha
            branch = $branch
        } | ConvertTo-Json

        Invoke-RestMethod -Uri $uri -Method Delete -Headers $headers -Body $deleteBody
    }

    # Upload the new file
    $uploadBody = @{
        message = $commitMessage
        content = $fileContent
        branch = $branch
    } | ConvertTo-Json

    $response = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $uploadBody

    Write-Host "PDF updated and old version moved to 'Older version' folder."
}
catch {
    Write-Host "An error occurred during GitHub upload: $_"
}
