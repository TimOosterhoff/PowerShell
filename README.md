Param (
   [string] $AccountsCSV = "Accounts.csv",
   [string] [ValidateSet("Ja", "Nee")] [AllowEmptyString()] $HeaderJaNee = "Nee",
   [string] [ValidateSet("Ja", "Nee")] [AllowEmptyString()] $MBXnMetFullAccessCSVopnieuwAanmaken = "Nee",
   [switch] $Help
   )

if ($Help) {
   Write-Host ""
   Write-Host "Script leest FullAccees en AutoMapping gegevens van account of CSV-met-accounts of *|all|alle" -ForegroundColor Green
   Write-Host "Zonodig eerst bestand MBXnMetFullAccess.csv (opnieuw) aanmaken" -ForegroundColor Green
   Write-Host "FullAccess wordt alleen uitgelezen bij aanmaken MBXnMetFullAccess.csv" -ForegroundColor DarkGreen
   Write-Host "Automapping wordt altijd uitgelezen" -ForegroundColor DarkGreen
   Write-Host "-AccountsCSV: <account> of CVS-met-accounts of '*' of 'all' of 'alle'" -ForegroundColor Yellow
   Write-Host ""
   Write-Host "-HeaderJaNee: Ja of Nee" -ForegroundColor Yellow
   Write-Host "Bevat -AccountsCSV, als deze naar een bestand verwijst, een header" -ForegroundColor DarkYellow
   Write-Host ""      
   Write-Host "-MBXnMetFullAccessCSVopnieuwAanmaken: Ja of Nee" -ForegroundColor Yellow
   Write-Host "Maakt in huidige map bestand met alle FullAccess rechten (opnieuw) aan" -ForegroundColor DarkYellow
   Write-Host ""      
   Write-Host  "Er worden 1 logbestand gemaakt" -ForegroundColor Green
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

Write-Host "Start        :" ((get-date -f yyyy-MM-dd) + " " + (get-date -f HH-mm-ss))  -ForeGround "Cyan"
# inlezen alle mailboxen met FullAccess rechten, duurt circa 6 minuten
Write-Progress "Alle mailaccounts met FullAccess worden ingelezen (duurt circa 1 minuut per 1000 accounts) ..."

if ($MBXnMetFullAccessCSVopnieuwAanmaken -eq 'Ja') {
    # rare constructie; iets met rawidentity!!
    
    $MBXnMetFullAccess = Get-Mailbox -ResultSize unlimited | Get-MailboxPermission | where {$_.user.tostring().split('\')[0] -ne "NT AUTHORITY" -and $_.IsInherited -eq $false} |
                         Select @{name = 'Identity'; Expression = {$_.Identity.tostring()} }, # .name; aanmame: unieke usernamen
                         @{Name='User'; Expression = {$_.User.rawidentity.split('\\')[-1] }},
                         @{Name='Access Rights';Expression = {[string]::join(', ', $_.AccessRights)}} |
                         Sort-Object User, Identity # split() ivm ZELF en SELF
    $MBXnMetFullAccess | Export-Csv 'MBXnMetFullAccess.csv' -NoTypeInformation -Encoding UTF8
    }
else {
    if (-not(Test-Path 'MBXnMetFullAccess.csv')) {
        write-host "Bestand      : MBXnMetFullAcces.csv niet aanwezig. Gebruik optie -MBXnMetFullAccessCSVopnieuwAanmaken = Ja" -ForeGround "Red"
        write-host ""
        exit
        }
    else {
        write-host "FullAccess   : MBXnMetFullAccess.csv is van" (Get-ChildItem 'MBXnMetFullAccess.csv').LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")  -ForeGround "Yellow"
        $MBXnMetFullAccess = Import-Csv 'MBXnMetFullAccess.csv'
        }
    }

$Accounts = @()
foreach ($regel in import-csv $AccountsCSV @ImportOpties) {
    $Accounts += $regel.$veld1
    }

$tijd            = (get-date -f yyyy-MM-dd) + " " + (get-date -f HH-mm-ss)
$LogFile         = "get-FullAccessAutomapping " + $tijd +  " " + $AccountsCSV.split('\/')[-1] #strip event. pad-naam

$FullAccessRechten = @()

#$mbx.Identity.split('/')[-1]

foreach ($mbx in $MBXnMetFullAccess ) {

    Write-Progress $mbx.user
    
    if ($Accounts.Contains($mbx.User)) {

        try {
            # $ADaccount is van de mailbox waarop de User FullAccess heeft gekregen
            $ADaccount = Get-ADUser ((Get-Mailbox $mbx.Identity).SAMaccountName) -ErrorAction Stop # Discovery mailbox heeft geen ADaccount
            }
        catch {
            continue
            }

        # op scherm als er 1 gemachtigde is
        if ($Accounts.Length -eq 1) {
            Write-Host $mbx.User -ForegroundColor Gray -NoNewline
            Write-Host '', (Get-Mailbox $mbx.user).displayname -NoNewline
            Write-Host ' FullAccess bij:' -ForegroundColor Gray -NoNewline
            Write-Host '', (Get-Mailbox $mbx.Identity).Name -ForegroundColor DarkYellow -NoNewline
            Write-Host '', (Get-Mailbox $mbx.Identity).displayname -ForegroundColor Yellow -NoNewline
            }

        $FullAccessRecht = new-object PSObject
        $FullAccessRecht | add-member -membertype NoteProperty -name "Gemachtigde" -value $mbx.User
        $FullAccessRecht | add-member -membertype NoteProperty -name "GemachtigdeDisplayName" -value (Get-Mailbox $mbx.user).displayname
        $FullAccessRecht | add-member -membertype NoteProperty -name "FullAccessBij" -value (Get-Mailbox $mbx.Identity).Name

        $msExchDelegateListLink = (Get-ADUser $ADaccount -properties msExchDelegateListLink).msExchDelegateListLink
        if ( $msExchDelegateListLink.contains((get-mailbox $mbx.user).DistinguishedName)) {

            # op scherm als er 1 gemachtigde is
            if ($Accounts.Length -eq 1) {
                Write-Host ' Automapped' -ForegroundColor DarkGreen
                }
            $FullAccessRecht | add-member -membertype NoteProperty -name "AutoMapping" -value 'Ja'
            
            }
        else {
            # Mogelijk geen FullAccess meer indien na aanmaken MBXnMetFullAccess.csv FullAccess rechten zijn verwijderd
            if ((Get-MailboxPermission -Identity $mbx.identity -User $mbx.user | select AccessRights) -match 'FullAccess') {
                # op scherm als er 1 gemachigde is
                if ($Accounts.Length -eq 1) {
                    Write-Host ' Niet automapped' -ForegroundColor Green
                    }
                $FullAccessRecht | add-member -membertype NoteProperty -name "AutoMapping" -value 'Nee'
                }
            else {
                Write-Host ' Geen FullAccess meer' -ForegroundColor Red
                }
            }
        $FullAccessRechten += $FullAccessRecht
        }
        
    }

$FullAccessRechten | Export-Csv $LogFile -NoTypeInformation -Encoding UTF8
Write-Host "Voltooid     :" ((get-date -f yyyy-MM-dd) + " " + (get-date -f HH-mm-ss))  -ForeGround "Cyan"
#$FullAccessRechten | Out-GridView
