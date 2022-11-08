Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[reflection.assembly]::loadwithpartialname("system.windows.forms")|Out-Null
Clear-Host

#Change photopath to folder holding pictures
$PhotoPath = "C:\Users\sfrancis3\Desktop\OneDrive_1_10-28-2022\1L photos Fall 2022"
$testlist = New-Object -TypeName 'System.Collections.ArrayList';

Function ResizeImage() {
    param([String]$ImagePath, [Int]$Quality = 90, [Int]$targetSize, [String]$OutputLocation, [String]$Name)
 
    Add-Type -AssemblyName "System.Drawing"
 
    $img = [System.Drawing.Image]::FromFile($ImagePath)
 
    $CanvasWidth = $targetSize
    $CanvasHeight = $targetSize
 
    #Encoder parameter for image quality
    $ImageEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($ImageEncoder, $Quality)
 
    # get codec
    $Codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where {$_.MimeType -eq 'image/jpeg'}
 
    #compute the final ratio to use
    $ratioX = $CanvasWidth / $img.Width;
    $ratioY = $CanvasHeight / $img.Height;
 
    $ratio = $ratioY
    if ($ratioX -le $ratioY) {
        $ratio = $ratioX
    }
 
    $newWidth = [int] ($img.Width * $ratio)
    $newHeight = [int] ($img.Height * $ratio)
 
    $bmpResized = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graph = [System.Drawing.Graphics]::FromImage($bmpResized)
    $graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
 
    $graph.Clear([System.Drawing.Color]::White)
    $graph.DrawImage($img, 0, 0, $newWidth, $newHeight)
 
    #save to file
    $bmpResized.Save($OutputLocation+$Name, $Codec, $($encoderParams))
    $bmpResized.Dispose()
    $img.Dispose()
}

# Get Source File Name
    Write-Host "Please Select CSV file"
    $msgBox =  [System.Windows.MessageBox]::Show("Please select CSV file containing the IDs","Select CSV File",[System.Windows.Forms.MessageBoxButtons]::OKCancel)
    switch($msgBox)
    {
        "OK"{}
        "Cancel"
         {
            Write-Host "Script Terminated"
            exit
         }
    }

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    [void]$FileBrowser.ShowDialog()


# Read File Contents
    $CsvContent = Get-Content $FileBrowser.FileName.Split(",")
    
    
    #line holds each line of data in csv email,id(s00),fullname
    foreach($line in $CsvContent)
    {
        Write-Output "Processing Pictures"
        #gatewayuser holds the email of the user
        $gatewayuser = $line.Split("@")[0]

        #id holds the id number in the csv 
        $id = $line.Split(",")[1]
	    
        #append csv
        

        #if the photo id exists
        if(Test-Path -Path ($PhotoPath+"\"+$id+".jpg"))
        {
	        $Date = Get-Date -Format "yyyy-MM-dd-HHmmss"
		    $Image = $PhotoPath+"\"+$id+".jpg"

            #Change to Path of where updated photographs are to be stored
		    $OutputFolder = "C:\Users\sfrancis3\Documents\UpdatedIDPhotos\"
		    $Name = ($id+".jpg")
 		    
            #Parameters: Image to be altered - Quality of image on a scale of 1/10 - Size of image - Folder to be output to - Name of image
		    ResizeImage $Image 90 300 ($OutputFolder) $Name
        }
	    else
	    {
		Write-Host ("No photo found "+$id)
		$testlist.Add($id)
	    }
     }



# Get Destination Folder
     Write-Host "Please Select Destination Folder"
     $msgBox =  [System.Windows.MessageBox]::Show("Please select the destination folder","Select Folder",[System.Windows.Forms.MessageBoxButtons]::OKCancel)
     switch($msgBox)
     {
        "OK"{}
        "Cancel"
	{
            Write-Host "Script Terminated"
            #exit
        }
     }

    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    [void]$FolderBrowser.ShowDialog()

#create new csv file
    $getdate = Get-Date -Format "MM-dd-yyyy-HH-mm"
    $csvfilename = ("\mapping-"+$getdate+".csv")
    $outfile =  ($FolderBrowser.SelectedPath+$csvfilename)
    
    foreach($line in $CsvContent)
    {
        Write-Output "Processing csv"
        #gatewayuser holds the email of the user
        $gatewayuser = $line.Split("@")[0]

        #id holds the id number in the csv 
        $id = $line.Split(",")[1]
            $record = [pscustomobject]@{
             'File_Name'= $id+".jpg"
             'External_Id'= $gatewayuser
           } 

       if(Test-Path -Path ($PhotoPath+"\"+$id+".jpg"))
       {
          # append record to CSV
          $record |Export-Csv $outfile -Append
       }   

    }
        

#writes to a file with error log
    $getdate = Get-Date -Format "MM-dd-yyyy-HH-mm"
    $filename = ("\errorlog-"+$getdate+".txt")
    "The following IDs do not have photos:" | Out-File -FilePath ($FolderBrowser.SelectedPath+$filename)
    $testlist | Out-File -FilePath ($FolderBrowser.SelectedPath+$filename) -Append
    $oReturn=[System.Windows.Forms.Messagebox]::Show("Photo Copy Completed")
