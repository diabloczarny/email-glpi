#Initialize Session
$R=Invoke-WebRequest http://192.168.1.228/glpi/index.php -SessionVariable thrunksession


#Set Login and Password
$R.Forms[0].Fields["login_name"]="asarat"
$R.Forms[0].Fields["login_password"]="AS!230as"
#$R.Forms[0].Fields
#$R.Forms[0].Fields["_glpi_csrf_token"]
#$token=$R.Forms[0].Fields["_glpi_csrf_token"]

$R=Invoke-WebRequest http://192.168.1.228/glpi/login.php -WebSession $thrunksession -Method Post -Body $R.Forms[0].Fields
#$R > C:\Users\asarat\Desktop\glpi.html
#Paweł
#$tokenPaw=$R.Forms[0].Fields["_glpi_csrf_token"]
$pawel=Invoke-WebRequest ("http://192.168.1.228/glpi/front/ticket.php?is_deleted=0&criteria%5B0%5D%5Bfield%5D=12&criteria%5B0%5D%5Bsearchtype%5D=equals&criteria%5B0%5D%5Bvalue%5D=notold&criteria%5B1%5D%5Blink%5D=AND&criteria%5B1%5D%5Bfield%5D=5&criteria%5B1%5D%5Bsearchtype%5D=equals&criteria%5B1%5D%5Bvalue%5D=7&search=Szukaj&itemtype=Ticket&start=0&_glpi_csrf_token="+$tokenPaw) -WebSession $thrunksession

$pawel.content > C:\Users\asarat\Desktop\zgloszenia\pawel.html 

$pawel_report=$pawel.content

#Piotrek
#$tokenPio=$R.Forms[0].Fields["_glpi_csrf_token"]
$piotrek=Invoke-WebRequest ("http://192.168.1.228/glpi/front/ticket.php?is_deleted=0&criteria%5B0%5D%5Bfield%5D=12&criteria%5B0%5D%5Bsearchtype%5D=equals&criteria%5B0%5D%5Bvalue%5D=notold&criteria%5B1%5D%5Blink%5D=AND&criteria%5B1%5D%5Bfield%5D=5&criteria%5B1%5D%5Bsearchtype%5D=equals&criteria%5B1%5D%5Bvalue%5D=248&search=Szukaj&itemtype=Ticket&start=0&_glpi_csrf_token="+$tokenPio) -WebSession $thrunksession

$piotrek.content > C:\Users\asarat\Desktop\zgloszenia\piotrek.html

$piotrek_report=$piotrek.content

#Rafał
#$tokenRaf=$R.Forms[0].Fields["_glpi_csrf_token"]
$rafal=Invoke-WebRequest ("http://192.168.1.228/glpi/front/ticket.php?is_deleted=0&criteria%5B0%5D%5Bfield%5D=12&criteria%5B0%5D%5Bsearchtype%5D=equals&criteria%5B0%5D%5Bvalue%5D=notold&criteria%5B1%5D%5Blink%5D=AND&criteria%5B1%5D%5Bfield%5D=5&criteria%5B1%5D%5Bsearchtype%5D=equals&criteria%5B1%5D%5Bvalue%5D=13&search=Szukaj&itemtype=Ticket&start=0&_glpi_csrf_token="+$tokenRaf) -WebSession $thrunksession

$rafal.content > C:\Users\asarat\Desktop\zgloszenia\rafal.html

$rafal_report=$rafal.content

#Artur
#$tokenArt=$R.Forms[0].Fields["_glpi_csrf_token"]
$artur=Invoke-WebRequest ("http://192.168.1.228/glpi/front/ticket.php?is_deleted=0&criteria%5B0%5D%5Bfield%5D=12&criteria%5B0%5D%5Bsearchtype%5D=equals&criteria%5B0%5D%5Bvalue%5D=notold&criteria%5B1%5D%5Blink%5D=AND&criteria%5B1%5D%5Bfield%5D=5&criteria%5B1%5D%5Bsearchtype%5D=equals&criteria%5B1%5D%5Bvalue%5D=361&search=Szukaj&itemtype=Ticket&start=0&_glpi_csrf_token="+$tokenArt) -WebSession $thrunksession

$artur.content > C:\Users\asarat\Desktop\zgloszenia\artur.html

$artur_report=$artur.content


$utf8 = New-Object System.Text.utf8encoding

$mail=get-storedcredential -Target mail

#$String = Get-Content "C:\Users\ARTEK\Desktop\pawel.html"

$pawel_report = Get-Content "C:\Users\asarat\Desktop\zgloszenia\pawel.html"

$piotrek_report = Get-Content "C:\Users\asarat\Desktop\zgloszenia\piotrek.html" 

$rafal_report = Get-Content "C:\Users\asarat\Desktop\zgloszenia\rafal.html"

$artur_report = Get-Content "C:\Users\asarat\Desktop\zgloszenia\artur.html"    
               
function GetStringBetweenTwoStrings($firstString, $secondString, $importPath){

    #Get content from file
    $file = $importPath

    #Regex pattern to compare two strings
    $pattern = "$firstString(.*?)$secondString"

    #Perform the opperation
    $result = [regex]::Match($file,$pattern).Groups[1].Value

    #Return result
    return $result

}

#---------------------petla testowa

if(!(Test-Path C:\Users\asarat\Desktop\zgloszenia\))
{
New-Item -Path "C:\Users\asarat\Desktop\" -Name "zgloszenia" -ItemType "directory"
}

$array_users=$pawel_report,$piotrek_report,$rafal_report,$artur_report

#$array_report="","",""

$array_filename="pawel.html","piotrek.html","rafal.html","artur.html"

$styles= Get-Content "C:\Users\asarat\Desktop\zgloszenia\styles.css"

$path_array="Get-Content C:\Users\asarat\Desktop\zgloszenia\pawel.html -raw","Get-Content C:\Users\asarat\Desktop\zgloszenia\piotrek.html -raw","Get-Content C:\Users\asarat\Desktop\zgloszenia\rafal.html -raw","Get-Content C:\Users\asarat\Desktop\zgloszenia\artur.html -raw"

$mail_array="test1@centermed.pl","test2@centermed.pl","test3@centermed.pl","test4@centermed.pl"
$test_mail_array="test_pp@centermed.pl","test_ps@centermed.pl","test_rl@centermed.pl"
$array_report="","","",""

for ($i=0; $i -le $array_users.Length -1 ;$i++)
{
$array_report[$i] = GetStringBetweenTwoStrings -firstString "<table border='0' class='tab_cadrehov'>" -secondString "</table>" -importPath $array_users[$i]

$array_report[$i] = "<head><style>$styles</style><table border='0' class='tab_cadrehov'>"+$array_report[$i]+"</table></head>"

$path = 'C:\Users\asarat\Desktop\zgloszenia\'+$array_filename[$i]

$array_report[$i] > $path

$array_report[$i] = (Get-Content  $path -Raw) -replace '/glpi','http://192.168.1.228/glpi'

$array_report[$i] > $path

if((Get-Content $path -raw).Length/1KB -gt 70)
{
$body= Get-Content  $path -Raw

Send-MailMessage -To $mail_array[$i] -From asarat@centermed.pl  -Subject "GLPI BETA" -Body $body"<br /><br />Wiadomość wygenerowana automatycznie." -BodyAsHtml -Credential $mail -SmtpServer serwer1447049.home.pl -Port 587 -encoding $utf8

}

}
#---------------------petla testowa

#$body= Get-Content  C:\Users\$env:username\Desktop\zgloszenia\rafal.html -Raw
#Send-MailMessage -To @("asarat@centermed.pl") -From asarat@centermed.pl  -Subject "GLPI" -Body $body"<br /><br />Wiadomość wygenerowana automatycznie." -BodyAsHtml -Credential $mail -SmtpServer serwer1447049.home.pl -Port 587 -encoding $utf8


