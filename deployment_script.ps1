# By Eric Post
#
# This is a deployment script I developed for deploying the front end code to WebServer01, the main DFS repllication webserver.
# The script has a gui. The dev would input an instance (customer code), code version and the TargetProcess URL involving the deployment.
#
# It then stops the IIS site/app pools, deletes the destination folder contents (save for some exceptions) then deploys the new code into that directory.
# After that, it starts the IIS site/app pools up again and fires off an email to changeMangamentSystem@company.com displaying what was deployed, who did it and where.
#
# I've sanitized it for future refence on GitHub.

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        
        Title="SRV Deployment" Height="481.833" Width="510.61">
    <Grid Margin="0,0,3.6,0">
        <TextBox Name ="t1" HorizontalAlignment="Left" Height="21" Margin="120,25,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="355"/>
        <TextBox Name ="t2" HorizontalAlignment="Left" Height="21" Margin="120,66,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="355"/>
        <TextBox Name ="t3" HorizontalAlignment="Left" Height="36" Margin="120,108,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="355"/>
        <Label Content="Web Instance" HorizontalAlignment="Left" Height="24" Margin="15,22,0,0" VerticalAlignment="Top" Width="135"/>
        <Label Content="Web Version" HorizontalAlignment="Left" Height="24" Margin="15,62,0,0" VerticalAlignment="Top" Width="102"/>
        <Label Content="TP URL" HorizontalAlignment="Left" Height="24" Margin="15,104,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBox Name ="Display" HorizontalAlignment="Left" Height="212" Margin="29,174,0,0" VerticalAlignment="Top" Width="446" IsReadOnly="True" />
        <Button Name ="enter" Content="Enter" HorizontalAlignment="Left" Height="25" Margin="379,146,0,0" VerticalAlignment="Top" Width="96"/>
        <Button Name ="close" Content="Close" HorizontalAlignment="Left" Margin="379,391,0,0" VerticalAlignment="Top" Width="96" Height="25"/>


    </Grid>
</Window>
"@
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
    try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
    catch{Write-Host "Unable to load Windows.Markup.XamlReader"; exit}

# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

#Events
$enter.Add_Click({

    $Display.Clear()
    
    $instanceName = $t1.Text
    $webVersion = $t2.Text
    $tpURL = $t3.Text

    if(!$instanceName)
    {
        Write-Host "Instance cannot be blank"
        if(!$webVersion)
            {
            Write-Host "Web Version cannot be blank"
            $Form.Close()
            exit
            }
        $Form.Close()
        exit
    }
    
    if(!$webVersion)
            {
            Write-Host "Web Version cannot be blank"
            $Form.Close()
            exit
            } 

    $Display.AppendText(($instanceName)+"`n")
    $Display.AppendText(($webVersion)+"`n")
    $Display.AppendText(($tpURL)+"`n"+"`n")

    $t1.Clear()
    $t2.Clear()
    $t3.Clear()

    $destination = "\\WebServer01\inetpub"
    $deploymentDirectory = "\\DeployServer01\deployments\WEB\QA\"
    $messageFrom = "operations@company.com"
    $messageTo = "changeMangamentSystem@company.com"
    $messageBody = "Updated WEB for $instanceName to $webVersion <br>by $env:username <br><br>Per $tpURL"
    $smtpServer = "emailServer.company.local"

    if (-not (Test-Path -Path $deploymentDirectory\$webVersion)) {
        throw 'The directory does not exist'
    } else {
    $Display.AppendText('The directory does exist and deployment will proceed'+"`n")
    #Write-Host 'The directory does exist and deployment will proceed' -ErrorAction Stop
    }

    Write-Host "Stopping IIS Site on WebServer01"
    # Stop the IIS Site on WebServer01 by passing $instanceName into the remote session.
    Invoke-Command -ComputerName WebServer01 -ScriptBlock {
        param($instanceName)
        Stop-IISSite -Name $instanceName -Confirm:$false
    } -ArgumentList $instanceName
    
    Write-Host "Stopping App Pool on WebServer01"
    # Stop the IIS App pool on WebServer01 by passing $instanceName into the remote session.
    Invoke-Command -ComputerName WebServer01 -ScriptBlock {
        param($instanceName)
        Stop-WebAppPool -Name $instanceName
    } -ArgumentList $instanceName
    
    Write-Host "Waiting 15 seconds for site and pools to finish stopping."
    # Pause the script to give WebServer01 time to stop both site and app pools.
    Start-Sleep -Seconds 15

    Write-Host "Deleting all files in the destination folder except for the QB folder, Web.config and newrelic.config."
    # Delete all the files in the $destination\$instanceName directory,
    # EXCLUDING the qb folder, Web.config, newrelic.config.
    # Using Get-ChildItem for the exclusion then piping it to Remove-Item 
    # was the only way I could prevent it from emptying the qb folder, despite it being excluded.
    Get-ChildItem -Path $destination\$instanceName -Exclude qb, Web.config, newrelic.config | Remove-Item -Recurse -Force

    Write-Host "Pausing for another 15 seconds to wait for the files in the destination folder to finish deleting."
    #Pause the script again to give it time to completely delete the files, just to be safe.
    Start-Sleep -Seconds 15
    
    Write-Host "Copying files from source to destination."
    # Copy files from $deploymentDirectory\$webVersion\ into the $instanceName.
    Copy-Item -Path $deploymentDirectory\$webVersion\* -Recurse -Destination $destination\$instanceName\

    Write-Host "Pausing for another 15 seconds to wait for the files to finish copying over."
    # Pause the script again to give it time to finish copying the files over.
    Start-Sleep -Seconds 15

    Write-Host "Starting the IIS Site on WebServer01"
    # Start the IIS Site on WebServer01 by passing $instanceName into the remote session.
    Invoke-Command -ComputerName WebServer01 -ScriptBlock {
        param($instanceName)
        Start-IISSite -Name $instanceName
    } -ArgumentList $instanceName

    Write-Host "Starting the App Pool on WebServer01"
    # Start the IIS App pool on WebServer01 by passing $instanceName into the remote session.
    Invoke-Command -ComputerName WebServer01 -ScriptBlock {
        param($instanceName)
        Start-WebAppPool -Name $instanceName
    } -ArgumentList $instanceName

    Write-Host "Sending email to changeMangamentSystem@company.com"
    # CMS Email
    Send-MailMessage -From $messageFrom -To $messageTo -Subject "$instanceName" -Body $messageBody -BodyAsHtml -SmtpServer $smtpServer;

    Write-Host "Deployment Complete!"
    $Display.AppendText('Deployment Complete'+"`n")
    })

$close.Add_Click({
    $Form.Close()
})
$Null = $Form.ShowDialog()
