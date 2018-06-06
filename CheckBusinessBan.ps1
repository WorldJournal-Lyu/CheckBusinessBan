param([string]$Area = "TEST")

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Import-Module WorldJournal.Email -Verbose -Force
Import-Module WorldJournal.Server -Verbose -Force

$beda    = (Get-WJPath -Name beda).Path
$mailMsg = ""

switch($Area){

    "AT" {

        $temp = ""
        Get-WJEmail -Name AT | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @("45631")
        break

    }

    "BO" {

        $temp = ""
        Get-WJEmail -Name BO | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @("45635")
        break

    }

    "CH" {

        $temp = ""
        Get-WJEmail -Name CH | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @("45634", "45637")
        break

    }

    "DC" {

        $temp = ""
        Get-WJEmail -Name DC | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @("45632", "45601")
        break

    }

    "NJ" {

        $temp = ""
        Get-WJEmail -Name NJ | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @(
                        "45633", 
                        "45641", "45642", "45643", "45644", "45645", "45646", "45647", "45648", "45649",
                        "45651", "45652", "45653", "45654", "45655", "45656", "45657", "45658", "45659"
                    )
        break

    }

    "NY" {

        $temp = ""
        Get-WJEmail -Name NY | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @(
                        "45611", "45612", "45613", "45614", "45615", "45616", "45617", 
                        "45621", "45622", "45623", "45624", "45625", "45626", "45424"
                    )
        break

    }

    "TEST" {

        $temp = ""
        Get-WJEmail -Name lyu | ForEach-Object { $temp += ("<"+$_.MailAddress + ">, ") }
        $mailList = $temp.Substring(0, $temp.Length-2)
        $banId    = @("45001")
        break

    }

}

$banId | ForEach-Object{

    if(Test-Path ($beda+"\"+$_)){

        $regex1 = ($_+"-\d{2}-\d{2}$")
        $regex2 = ($_+"-\d{2}$")
        $storyCount = 0
        $imageCount = 0
        $mailMsg += $_ + "REPLACE`n"
        $mailMsg += "`n"

        Get-ChildItem -Path ($beda+"\"+$_) -Filter "*.txt" | ForEach-Object{

            $mailMsg += "  " + $_.BaseName + "`n"

            if($_.BaseName -match $regex1){
                $storyCount++
            }

            if($_.BaseName -match $regex2){
                $imageCount++
            }

        }

        $mailMsg = $mailMsg.Replace(($_+"REPLACE") ,($_+" ("+$storyCount + "文 " + $imageCount + "圖說)"))
        $mailMsg += "`n"

    }

}

#$mailMsg

if($mailMsg -ne ""){

    $mailFrom = (Get-WJEmail -Name noreply).MailAddress
    $mailPass = (Get-WJEmail -Name noreply).Password

    Emailv3 -From $mailFrom -Pass $mailPass -To $mailList -Subject ($Area+"工商文圖 " + (Get-Date).ToString("yyyy-MM-dd hh:mm")) -Body $mailMsg

}#else{"No " + $Area + " Business Ban files found"}
#Pause