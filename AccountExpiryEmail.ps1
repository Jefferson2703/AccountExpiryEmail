#Quantidade de dias setado que devemos pegar 
$15Days = New-TimeSpan -Days 15

#Procura por contas expiradas 
$AcsExp = Search-ADAccount -AccountExpiring -TimeSpan $15Days

#Pega data atual e hora
$start = [datetime]::Now
$today = $start

#formata o e-mail para o padrão UTF8
$textEncoding = [System.Text.Encoding]::UTF8

 ForEach($AcExp in $AcsExp){

    $mgrEmail = $null
    $mgrName = $null
    $ExpDate = $AcExp.AccountExpirationDate 
	$AcName = $AcExp.Name

    #pega a data atual menos a data de expiração para calcular a quantidade de dias.        
    $daysToExpire = New-TimeSpan -Start $today -End $ExpDate

    #converter na quantidade de dias que faltam para expirar 
    $daysToExpire = [math]::Round($daysToExpire.TotalDays)

    #Pega o gerente do usuário que aconta vai expirar 
    $mgr = (Get-Aduser $AcExp -Properties *).manager

    if($mgr -ne $null )
		{   
            # Pega o e-mail do gerente e nome 
			$mgrEmail = (Get-Aduser $mgr -Properties *).mail
            $mgrName = (Get-Aduser $mgr -Properties *).name
		}
    #se o gerente estiver nulo adiona ela em uma lista separada 
    if ($mgrEmail -eq $null){
        Write-host "A conta do colaborador $Acname , vai expirar em $daysToExpire dias. O não achei o gerente dele $mgrEmail "
        $userS = Get-ADUser $AcExp.samaccountname -properties * | select  DisplayName, samaccountname, EmailAddress
        $user | Add-Member -MemberType NoteProperty -Name "daysToExpire" -Value "$daysToExpire" -Force
        $user | Export-Csv c:\export\Sem_Gerente.csv -Append -NoTypeInformation -Encoding Utf8

    }
    else{
        #Cria lista de usuários por gerente 
        Write-host "A conta do colaborador $Acname , vai expirar em $daysToExpire dias. O achei o gerente dele $mgrEmail "
        $user = Get-ADUser $AcExp.samaccountname -properties * | select DisplayName, samaccountname, EmailAddress 
        $user | Add-Member -MemberType NoteProperty -Name "daysToExpire" -Value "$daysToExpire" -Force
        $user | Export-Csv c:\export\$mgrName".csv" -Append -NoTypeInformation -Encoding Utf8
    }

 }
 
 #foreach para enviar o e-mail 
  ForEach ($AcExp in $AcsExp)
 {
    #Configurações variaveis de e-mail 
	$mgrEmail = $null
	$AdminEmail = "email@dominio.com"
	$Fromusr = "email@dominio.com"
	$ExpDate = $AcExp.AccountExpirationDate #ajusta a data 
	$AcName = $AcExp.Name
	$Sub = "USUÁRIOs QUE VÃO EXPIRAR NO PRÓXIMOS 15 DIAS" #Titulo e-mail 
    $body = "Olá, Em anexo os colaboradores que vão expirar nos próximos 15 dias." #Body E-mail
    

    #Configruações do servidor de e-mail 
	$anonUser = "email@dominio.com"
	$anonPass = ConvertTo-SecureString "Password" -AsPlainText -Force
	$anonCred = New-Object System.Management.Automation.PSCredential($anonUser, $anonPass)
	$Smtpsvr = "smtp.office365.com" #Anonymous allowed for the user or machine running this script
    $SMTPPort = 587

	$mgr = (Get-Aduser $AcExp -Properties *).manager
    
    $i = "_"+(Get-Date -format "ddMMyyyy_HHmmss") #pega data atual 

	if($mgr -ne $null)
		{
			$mgrEmail = (Get-Aduser $mgr -Properties *).mail
            $mgrName = (Get-Aduser $mgr -Properties *).name
		}
    
    #Se o arquivo existir na pasta, envia e-mail
    if ((Test-Path -Path ("C:\export\$mgrName.csv") ) -eq 'True'){
        $Attachment = "C:\export\$mgrName.csv"
        Send-MailMessage -to $AdminEmail -from $mgrEmail -Port $SMTPPort -UseSsl -subject $sub -SmtpServer $Smtpsvr -Credential $anonCred -Attachments $Attachment -Body $Body -priority High -Encoding $textEncoding -ErrorAction Stop #Envio do e-mail
        dir C:\export\$mgrName.csv | % {ren -Path $_.fullname ($_.name.substring(0, $_.name.length-4) +$i +$_.Extension ) } #Renomei o arquivo com a data atual.
    }
    #se existir o arquivo "Sem_gerente.csv" envia e-mail para o adminstrador 
    if ((Test-Path -Path ("c:\export\Sem_Gerente.csv") ) -eq 'True'){
        $Attachment = "c:\export\Sem_Gerente.csv"
        Send-MailMessage -to $AdminEmail -from $AdminEmail -Port $SMTPPort -UseSsl -subject $sub -SmtpServer $Smtpsvr -Credential $anonCred -Attachments $Attachment -Body $Body -priority High -Encoding $textEncoding -ErrorAction Stop
        dir c:\export\Sem_Gerente.csv | % {ren -Path $_.fullname ($_.name.substring(0, $_.name.length-4) +$i +$_.Extension ) }
    }

 }#For#>
 
 #move para pasta bkp 
 Move-Item -Path C:\export\*.csv -Destination C:\export\backup\ -PassThru