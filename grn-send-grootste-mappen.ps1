<#
Functie : Script stuurt bericht met grootste mappen en grootste berichten naar mailbox of CSV-met-mailboxen of *|all|alle
Vereist : PowerShell 3, Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Versie  : 1.0.4
Datum   : 2017-02-10
Auteur  : Tim Oosterhoff
Email   : taoosterhoff@gmail.com
Licentie: GNU General Public License
#>


Param (
   [string] $AccountsCSV                                                 = "petreh1x",
   [string] [ValidateSet("Ja", "Nee")] [AllowEmptyString()] $HeaderJaNee = "Nee",
   [string] $PadNaarMetingBestanden                                      = "\\grnwi357\D\scripts\output\",
   [string] $MetingBestandPrefix                                         = "MBsize ",
   [int]    $StartJaarHistorie                                           = 2015,
   [string] $MetingOpDatums                                              = "",
   [string] $NaarAnderAccountOfEmailAdres                                = "",
   [switch] $Help
   )

if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null ) {
    write-host "PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 wordt geladen"
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    }

if ($Help) {
    Write-Host ""
    Write-Host "Script stuurt bericht met totalen, grootste mappen en grootste berichten naar mailbox of CSV-met-mailboxen of *|all|alle" -ForegroundColor Green
    Write-Host "-AccountsCSV: <mailbox> of CVS-met-mailboxen of '*' of 'all' of 'alle'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "-HeaderJaNee: Ja of Nee" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "-PadNaarMetingBestanden" -ForegroundColor Yellow
    Write-Host "Voorbeeld: \\grnwi357\D\scripts\output\ (let op achterliggende backslash)" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "-MetingBestandPrefix" -ForegroundColor Yellow
    Write-Host "Voorbeeld: 'MBsize ' (let op achterliggende spatie)" -ForegroundColor DarkYellow
    Write-Host "Voorbeeld volledige bestandsnaam: MBsize 2018-01-01.csv" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "-StartJaarHistorie" -ForegroundColor Yellow
    Write-Host "Geeft Totalen op 1 januari vanaf opgegeven jaar" -ForegroundColor DarkYellow
    Write-Host "Voorbeeld: 2017" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "-MetingOpDatums" -ForegroundColor Yellow
    Write-Host "Geef datum(s) op. In de tabel Totalen worden deze dan gepresenteerd. Nuttig om 0-meting van een schoonmaakactie te presenteren" -ForegroundColor DarkYellow
    Write-Host "Voorbeeld: 2018-06-31" -ForegroundColor DarkYellow
    Write-Host "Voorbeeld: '2018-06-30, 2018-08-31'" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "-NaarAnderAccountOfEmailAdres" -ForegroundColor Yellow
    Write-Host "Mailtje wordt naar ander account of emailadres verzonden" -ForegroundColor DarkYellow
    Write-Host "Voorbeeld: petoud1x" -ForegroundColor DarkYellow
    Write-Host "Voorbeeld: peter.den.oudsten@groningen.nl" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host  "Er worden 1 logbestand gemaakt" -ForegroundColor DarkYellow
    exit
    }

if ("*", "all", "alle" -contains $AccountsCSV) {
   $AccountsCSV = "Alle email accounts.csv"
   $HeaderJaNee = "Ja"
   write-host "Selectie     : Alle email accounts" -ForeGround "Cyan"
   write-host 'Let op       : Bestand "Alle email accounts.csv" bevat alleen Mailbox accounts' -ForeGround "Cyan"
   write-host "             : (dus geen Distribution Group en Mailcontact accounts)" -ForeGround "Cyan"
   Write-Progress  -Activity "Alle email accounts worden verzameld ..."
   Get-Mailbox -ResultSize unlimited  | Select Name | sort-object Name | export-csv $AccountsCSV -Encoding UTF8 -NoTypeInformation
   }

#als $accountCSV geen .csv bevat dan aanname dat waarde als accountnaam is bedoeld
if (-not $accountsCSV.toLower().Contains(".csv")) {
   $AccountsCSV | Out-File .\"Account "$AccountsCSV".csv"
   $AccountsCSV = "Account " + $AccountsCSV + ".csv"
   $HeaderJaNee = "Nee"
   }

write-host "Bestand      :" $accountsCSV -ForeGround "Cyan"
write-host "Header       :" $HeaderJaNee -ForeGround "Cyan"

if (-not(Test-Path $AccountsCSV)) {
   write-host "Bestand      : Niet aanwezig" -ForeGround "Red"
   write-host ""
   exit
   }

$ImportOpties = @{header = "veld1"}
$veld1 = "Veld1"

$records = (Get-Content $AccountsCSV).count

if ($HeaderJaNee -eq "Ja") {
   $ImportOpties = ""
   $veld1 = (get-content $accountsCSV -totalcount 1 ).Replace("`"","")
   $records -= 1
   write-host "Header       :" $veld1 -ForeGround "Cyan"
   }


# HTML code en tekst van het bericht
$head = @'

<style>
body { background-color:#f2f2f2;
       font-family:Calibri,Candara,Segoe,Segoe UI,Optima,Arial,sans-serif;
       font-size:10pt; }
td, th { border:1px solid black; padding-left: 5px }


th { color:white;
     background-color:#80AD15; }
#
table { margin-left:5px; border-collapse: collapse; font-family:Calibri,Candara,Segoe,Segoe UI,Optima,Arial,sans-serif; font-size:9pt }
td:nth-last-child(1) {font-family:Courier New; text-align: right; padding-right: 5px}

</style>
'@


$BerichtBegin = @'
Geachte medewerk(st)er,<br>
<br>
<b>Wat zit er in deze mailbox</b><br>
Bij deze 3 tabellen; totalen, de grootste mappen en de grootste berichten.<br>
Tegelijk een verzoek van onze kant: wil je berichten die geen waarde meer hebben verwijderen?<br>
<br>
<b>Hyperlinks in onderstaande tabellen</b><br>
In OutLook, op Windows, zijn de mappen en berichten hyperlinks.<br>
<small style:font-size=2px>(De links werken alleen in Outlook, dus niet in OWA.)</font></small><br>
<br>
<b>Tip om je Agenda op te ruimen</b><br>
Kies:
<ol>
<li>Agenda</li>
<li>Beeld</li>
<li>Weergave wijzigen</li>
<li>Lijst</li>
</ol>
Daarna wijst het zich vanzelf.<br> 
'@

$BerichtEinde= @'
<br>
Wil je graag een nieuw overzicht? Stuur even een mailtje.<br>

<br>
Vragen, tips, reageren?<br>
Mail gerust.<br>
<br>
Mans Kok<br>
Tim Oosterhoff<br>
'@


$TotalenHeader = @"
<style>
TABLE{border-width: 1px;border-style: solid; border-color: #324408; border-collapse: collapse; text-align: center; padding: 2px}
TD{border-width: 1px;border-style: solid;border-color: #324408;border-collapse: collapse; text-align: right; padding: 2px}
</style>
"@

[array]$Accounts = (import-csv $AccountsCSV @ImportOpties)

# Datums verzamelen
[array]$datums = @()
if ($StartJaarHistorie -le (get-date).year) {
    foreach ($jaar in $StartJaarHistorie ..((get-date).year) ) {
        $datums += "$($jaar)-01-01"
        }
    }

foreach ($datum in $MetingOpDatums.split(' ,;')) {
    if ($datum.trim()) {
        $datums += $datum
        }
    }

$datums = $datums | Sort-Object -Unique

# Inlezen eerdere metingen; Mailboxen met AantalItems en GrootteInMB
$TotalenOpDatum = @()
foreach ( $datum in $datums ) {
    Write-Progress -Activity "Inlezen meetgegevens $datum).."
    try {
        $ingelezen = Import-Csv "$PadNaarMetingBestanden$MetingBestandPrefix$datum.csv" -ErrorAction Stop
        # Hieronder benadering via hashtable. Performance veeel slechter!
        #$ingelezen = Import-Csv "$($PadNaarMetingBestanden)$MetingBestandPrefix$([string]$jaar)-01-01.csv" -ErrorAction Stop | Group-Object -AsHashTable -Property Name
        $TotalenOpDatum+= [PSCustomObject] @{ Datum   = $datum
                                              Meting =  $ingelezen
                                              }
        }
    catch {
        Write-Host "Waarschuwing : " -NoNewline -ForegroundColor Cyan
        Write-Host "$PadNaarMetingBestanden$MetingBestandPrefix$datum.csv" -ForegroundColor Yellow -NoNewline
        Write-Host " niet gevonden"
        }
    }


$tijd            = get-date -f 'yyyy-MM-dd HH-mm-ss'
$LogFile         = "send-grootste-mappen " + $tijd +  " " + $AccountsCSV.split('\/')[-1] #strip event. pad-naam
$stopwatch       = [Diagnostics.Stopwatch]::StartNew()
$records         = $Accounts.count
$ronde           = 0

Add-Type -AssemblyName System.Web

foreach ($mailaccount in $accounts) {
    try {
        $account = Get-Mailbox -identity $mailaccount.$veld1 -ErrorAction Stop
        }
    catch {
        write-host "Fout         : Account " -NoNewline -ForegroundColor Yellow
        Write-Host $mailaccount.$veld1 -NoNewline -ForegroundColor Red
        Write-Host " niet gevonden" -ForeGround Yellow
        continue
        }
    $ronde  += 1
    Write-Progress  -Activity $account.name "($ronde van $records) ..." -SecondsRemaining ($stopwatch.Elapsed.totalseconds/$ronde*($records-$ronde))


    $Totalen = @()
    foreach ( $datum in $TotalenOpDatum ) {
        $IndexAccountInArray = [array]::IndexOf( $datum.Meting.name, $account.name)
        if ($IndexAccountInArray -ge 0 ) {
            $Totalen += [PSCustomObject]@{ Datum       = $datum.Datum
                                           GrootteInMB = ('{0:N0}' -f [int]($datum.Meting[$IndexAccountInArray].SizeInMB))
                                           AantalItems = ('{0:N0}' -f [int]($datum.Meting[$IndexAccountInArray].Items))
                                           }
            }
        }

    $Start = new-object PSObject
    $Start | add-member -MemberType NoteProperty -Name "Datum" -Value (get-date -f yyyy-MM-dd)


    try {
        $GrootteNu = (Microsoft.Exchange.Management.PowerShell.E2010\get-mailboxstatistics $account.Name -ErrorAction Stop -WarningAction SilentlyContinue ).totalitemsize.value.tomb()
        }
    catch {
        write-host "Attentie     : Account " -NoNewline -ForegroundColor Yellow
        Write-Host $mailaccount.$veld1 $error[0] -NoNewline -ForegroundColor Red
        Write-Host " nog niet gebruikt" -ForeGround Yellow
        continue
        }

    $Start | add-member -MemberType NoteProperty -Name "GrootteInMB" -Value ('{0:n0}' -f [int]$GrootteNu)
    $AantalNu = ('{0:n0}' -f [int](get-mailboxstatistics $account.Name).itemcount)
    $Start | add-member -MemberType NoteProperty -Name "AantalItems" -Value $AantalNu
    $Totalen += $Start

    $Totalen = $Totalen | ConvertTo-Html -PreContent "<b><font color=#527A00>Totalen:</b></font>"  -Head $TotalenHeader
    #$Totalen = $Totalen | ConvertTo-Html -Fragment -PreContent "<b><font color=#527A00>Totalen:</b></font>"
    
    $GrootsteMappen = Get-MailboxFolderStatistics $account.Name | Where-Object {$_.FolderType -notmatch "RecoverableItems"} |
    Sort-Object foldersize -Descending| Select-Object @{n='Map';e={( $_.folderpath).substring(1)} }, @{n='GrootteInMB';e={'{0:n0}' -f [int]($_.Foldersize.ToMB())} }, @{n='AantalItems'; e={'{0:n0}' -f [int]($_.itemsinfolder)} } -First 100 |
    Where {$_.GrootteInMB -GT 0} |
    Convertto-Html -Fragment -PreContent "<b><font color=#527A00>De (max 100) grootste mappen<small style:font-size=2px> ( >= 1 MB )</small>:</b></font>" |
    % { $_ -replace '<tr><td>(.*?)</td>', '<tr><td><a href="outlook:$1">$1</a></td>'}

    $GrootsteBerichtInMap = Get-MailboxFolderStatistics -Identity $account.name -includeanalysis -FolderScope All |
    Where-Object {$_.FolderType -notmatch "RecoverableItems"} | Sort-Object TopSubjectSize -Descending | 
    Select-Object @{n='Map'; e={($_.folderPath).substring(1)} }, @{n='Bericht';e={$_.topsubject} }, @{n='GrootteInMB'; e={$_.TopSubjectSize.toMB()} } -First 200 |
    Where {$_.GrootteInMB -GT 0} |
    Convertto-Html -Fragment -PreContent "<b><font color=#527A00>De (max 200) grootste berichten per map<small style:font-size=2px> ( >= 1 MB)</small>:</b></font>"  |
    % { $_ -replace '<tr><td>(.*?)</td><td>(.*?)</td>', '<tr><td><a href="outlook:$1">$1</a></td><td><a href="outlook:$1/~$2">$2</a></td>'}

    
    $Bericht = ConvertTo-Html -Head $head -Body "<br>$Totalen<br>$GrootsteMappen<br>$GrootsteBerichtInMap"| Out-String
    
    $MailNaar = (Get-Mailbox $account).PrimarySMTPaddress
    if ($NaarAnderAccountOfEmailAdres) {
        if ( $NaarAnderAccountOfEmailAdres -match '@') {
            $MailNaar = $NaarAnderAccountOfEmailAdres
            }
        else {
            $MailNaar = (Get-Mailbox $NaarAnderAccountOfEmailAdres).PrimarySMTPaddress
            }
        }
    $postbus = $account.name

    try {
        send-mailmessage -from "exchange.beheer@groningen.nl" -to $MailNaar -subject "Wat zit er in deze mailbox ($postbus)?" -BodyAsHtml -body "$BerichtBegin $Bericht $BerichtEinde" -smtpServer mail2.groningen.nl -ErrorAction Stop
#        send-mailmessage -from "Exchange.Beheer@groningen.nl" -to 'tim.oosterhoff@groningen.nl' -subject "Wat zit er in deze mailbox ($postbus)?" -BodyAsHtml -body "$BerichtBegin $Bericht $BerichtEinde" -smtpServer mail2.groningen.nl -ErrorAction Stop
        }
    catch {
        write-host "Attentie     : Account " -NoNewline -ForegroundColor Yellow
        Write-Host $mailaccount.$veld1 -NoNewline -ForegroundColor Red
        Write-Host " heeft geen email adres" -ForeGround Yellow
        continue
        }


    $LogRegel = new-object PSObject
    $LogRegel | add-member -membertype NoteProperty -name "Mailbox" -value $postbus
    $LogRegel | add-member -membertype NoteProperty -name "Tijdstip" -value (get-date -f 'yyyy-MM-dd HH:mm:ss')
    $LogRegel | add-member -membertype NoteProperty -name "GrootteInMB" -value $GrootteNu
    $LogRegel | add-member -membertype NoteProperty -name "AantalItems" -value $AantalNu
    $LogRegel | Export-Csv $LogFile -Append -NoTypeInformation -Encoding UTF8

    }


if (Test-Path ($LogFile)) {
   write-host "LogFile      :" $LogFile -ForeGround Cyan
   }

write-host "Voltooid     :                     " (get-date -f 'yyyy-MM-dd HH:mm:ss') -ForeGround Cyan
