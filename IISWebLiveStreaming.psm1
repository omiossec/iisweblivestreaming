function set-IISWebLiveStreamingPublishingPoint
{
<#
.SYNOPSIS
    This function start a publish point 
.DESCRIPTION
    This function start a publish point 
.PARAMETER ApplicationName
    The application you want to restart
.PARAMETER WebSiteName
    The WebSite where reside the application 
    By default "Default Web Site"
.PARAMETER PublishingPointCommand
    the action to apply to the Publishing point
    Start|Stop|Shutdown
    Default action Start
.EXAMPLE 
    restart-IISWebLiveStreamingPublishingPoint -ApplicationName "MyApplication" -WebSiteName "MyWebSite" 
.NOTES
    Olivier Miossec
#>
    PARAM (
		[parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
        $ApplicationName,
        [parameter(Mandatory = $False)]
        $WebSiteName = "Default Web Site",
        [parameter(Mandatory = $true)]
        [ValidateSet("Start","Stop","Shutdown")]
        $PublishingPointCommand = "Start"
    )

    BEGIN
	{
        write-debug "load web.administration DLL"
        [System.Reflection.Assembly]::LoadFrom( "C:\windows\system32\inetsrv\Microsoft.Web.Administration.dll" ) |Out-Null
    }

    PROCESS 
    {
       
       try {
            $site = (New-Object Microsoft.Web.Administration.ServerManager).Sites[$WebSiteName]       

            $section = $site.GetWebConfiguration().GetSection("system.webServer/media/liveStreaming")
            

            $instance = $section.Methods["GetPublishingPoints"].CreateInstance()

            $instance.Input["siteName"] = $site.Name
            $instance.Input["virtualPath"] = $ApplicationName


            $collection = $instance.Output.GetCollection()

            $collection

            foreach ($item in $collection.GetCollection())
            {

                $methode = $item.Methods[$PublishingPointCommand]
                $ItemInstance = $methode.CreateInstance()
                $ItemInstance.Execute()


            }
       }
       catch 
       {
            Write-Error "Error : " $error[0].Exception.Message
       }
       
       
       


    }

    END 
    {

    }
}

function Convert-IISWebLiveStreamingStatus
{
    PARAM (
        [parameter(Mandatory = $true)]
        [String]$state
    )
    <#
    Helper function take streaming status code and return the String value
    #>
    #Idle|Starting|Started|Stopping|Stopped|ShuttingDown|Error
    $stateString = @{"0"="Idle";"1"="Starting";"2"="Started";"3"="Stopping";"4"="Stopped";"5"="ShuttingDown";"6"="Error"}
    return $stateString.item($state)
}
function Convert-IISWebLiveSourceType
{
    PARAM (
        [parameter(Mandatory = $true)]
        [String]$SourceType
    )
    <#
    Helper function take streaming status code and return the String value
    #>
    #Push|Pull
    $SourceString = @{"0"="Push";"1"="Pull"}
    return $SourceString.item($SourceType)
}
function get-IISWebLiveStreamingPublishingPoint
{

<#
.SYNOPSIS
    This function get the attribute of the Live Streaming publishing point
.DESCRIPTION
    This function get the attribute of the Live Streaming publishing point
.PARAMETER ApplicationName
    The application you want to test
.PARAMETER WebSiteName
    The WebSite you want to test 
    By default "Default Web Site"
.EXAMPLE 
    get-IISWebLiveStreamingPublishingPoint -ApplicationName "MyApplication" -WebSiteName "MyWebSite" 
    return a PsObject 
.NOTES
    Olivier Miossec
#>
    PARAM (
		[parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
        $ApplicationName,
        [parameter(Mandatory = $False)]
        $WebSiteName = "Default Web Site"
        
    )


    BEGIN
	{
        write-debug "load web.administration DLL"
        [System.Reflection.Assembly]::LoadFrom( "C:\windows\system32\inetsrv\Microsoft.Web.Administration.dll" ) |Out-Null
    }

    PROCESS 
    {
        
        try {
                    #load the website
                    $site = (New-Object Microsoft.Web.Administration.ServerManager).Sites[$WebSiteName] 
                    #load the Live Streaming Section 
                    $section = $site.GetWebConfiguration().GetSection("system.webServer/media/liveStreaming")
                    #create an instance to access publishing points 
                    $instance = $section.Methods["GetPublishingPoints"].CreateInstance()

                    $instance.Input["siteName"] = $site.Name
                    $instance.Input["virtualPath"] = $ApplicationName

                    $instance.execute()


                    $collection = $instance.Output.GetCollection()



                    [System.Collections.ArrayList]$arrPublishPointsInfos = @()


                    foreach ($point in $collection.GetCollection())
                    {
                                $PointInfos = New-Object PSObject
                                $PointInfos | Add-Member -MemberType NoteProperty  -Name "Name" -Value $point.Attributes["name"].Value.ToString()
                                $PointInfos | Add-Member -MemberType NoteProperty  -Name "VirtualPath" -Value $point.Attributes["virtualpath"].Value.ToString()
                                $sourcestring = Convert-IISWebLiveSourceType -SourceType $point.Attributes["sourcetype"].Value.ToString()
                                $PointInfos | Add-Member -MemberType NoteProperty -Name "SourceType" -Value $sourcestring
                                $PointInfos | Add-Member -MemberType NoteProperty  -Name "Archives" -Value $point.Attributes["archives"].Value.ToString()
                                $PointInfos | Add-Member -MemberType NoteProperty  -Name "Streams" -Value $point.Attributes["streams"].Value.ToString()
                                $PointInfos | Add-Member -MemberType NoteProperty  -Name "Fragments" -Value $point.Attributes["fragments"].Value.ToString()
                                $PointInfos | Add-Member -MemberType NoteProperty  -Name "LastError" -Value $point.Attributes["lasterror"].Value.ToString()
                                
                                $stateString = Convert-IISWebLiveStreamingStatus -state $point.Attributes["state"].Value.ToString()
                                $PointInfos | Add-Member -MemberType NoteProperty -Name "State" -Value $stateString 
                                $arrPublishPointsInfos.Add($PointInfos)
                    }

                    return $arrPublishPointsInfos
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            write-error $FailedItem
            write-error $ErrorMessage
        }
        
        
        

    }

    END 
    {
        write-debug "End Function"
    }


}


function get-IISWebLiveStreamingPublishingPointStatus
{

<#
.SYNOPSIS
    This function get the attribute of the Live Streaming publishing point
.DESCRIPTION
    This function get the attribute of the Live Streaming publishing point
.PARAMETER ApplicationName
    The application you want to test
.PARAMETER PublishingPoint
    The PublishingPoint you want to test
.PARAMETER WebSiteName
    The WebSite you want to test 
    By default "Default Web Site"
.EXAMPLE 
    get-IISWebLiveStreamingPublishingPoint -ApplicationName "MyApplication" -WebSiteName "MyWebSite" 
    return a PsObject 
.NOTES
    Olivier Miossec
#>
    PARAM (
		[parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
        $ApplicationName,
        [parameter(Mandatory = $False)]
        $WebSiteName = "Default Web Site",
        [parameter(Mandatory =$true)]
        $PublishingPoint
        
    )

    BEGIN
	{
        write-debug "load web.administration DLL"
        [System.Reflection.Assembly]::LoadFrom( "C:\windows\system32\inetsrv\Microsoft.Web.Administration.dll" ) |Out-Null
    }

    PROCESS 
    {
        
        try {
                    #load the website
                    $site = (New-Object Microsoft.Web.Administration.ServerManager).Sites[$WebSiteName] 
                    #load the Live Streaming Section 
                    $section = $site.GetWebConfiguration().GetSection("system.webServer/media/liveStreaming")
                    #create an instance to access publishing points 
                    $instance = $section.Methods["GetPublishingPoints"].CreateInstance()

                    $instance.Input["siteName"] = $site.Name
                    $instance.Input["virtualPath"] = $ApplicationName

                    $instance.execute()


                    $collection = $instance.Output.GetCollection()

                    $stateNum = 1000

                    foreach ($point in $collection.GetCollection())
                    {

                        if ($point.Attributes["name"].Value.ToString() -eq $PublishingPoint)
                        {
                            $stateNum = $point.Attributes["state"].Value             
                        }

                    }

                    $stateString = Convert-IISWebLiveStreamingStatus -state $stateNum
                    return $stateString
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            write-error $FailedItem
            write-error $ErrorMessage
        }
        
    }

    END 
    {
        write-debug "End Function"
    }

}


Export-ModuleMember -Function get-IISWebLiveStreamingPublishingPointStatus, get-IISWebLiveStreamingPublishingPoint,set-IISWebLiveStreamingPublishingPoint