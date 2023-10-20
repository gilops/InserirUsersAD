$headers = @{
    "Authorization"="Inserir o usuario aqui"
};
#Dentro da url colocar a query
$url = ""
$result  = Invoke-WebRequest -UseBasicParsing -Uri $url -Headers $headers
$open_requests = $result.Content | ConvertFrom-Json

echo "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-"
Get-Date

if (!$open_requests.result) {

    echo "N�o existem requests abertos"
    Exit
}

$open_requests.result.sys_id | ForEach-Object {
    
    echo $_":"
    
    #Pega detalhes de uma request especifica (no formato json)
    $url = "https://service-now.com/api/now/table/sc_req_item/$_"
    $result  = Invoke-WebRequest -UseBasicParsing -Uri $url -Headers $headers
    
    $request_details = $result.Content | ConvertFrom-Json
    $request_description = $request_details.result.description
    $splitted = $request_description -split '\n'
    #echo $splitted
    $nome         = $($splitted[4] -split ": ")[1]
    $sobrenome    = $($splitted[6] -split ": ")[1]
    $nomecompleto = $nome + " " + $sobrenome
    $mail         = $($splitted[5] -split ": ")[1]
    $grupo        = $($splitted[7] -split ": ")[1]
    #verifica se a parte antes do arroba tem mais de 14 caracteres
    if($mail.Split("@")[0].Length -gt 14){
        #Pega o que tem antes do arroba e corta até o caracter 14
        $username = $mail.Split("@")[0].Substring(0, 14)
    }
    else{
        #pega o que tem antes do arroba
        $username   = $($mail -split "@")[0]
    }
    #OU Service Now
    $OU         = "Caminho da OU"
    $UPN          = "$username"
    $vpn          = "vpn"
    echo $request_details.result.number
    
    #Biblioteca Assembly Para criar a senha
    Add-Type -AssemblyName System.Web
    $pass = [System.Web.Security.Membership]::GeneratePassword(16,4)
    $password = ConvertTo-SecureString $pass -AsPlainText -Force
    
    #Cria��o do Usu�rio no AD
    #Checa se o usu�rio j� existe no AD
    if (Get-ADUser -F {SamAccountName -eq $username}) {

        echo "Nome de usu�rio j� existente: $username."
        #fechar o chamado:
        $url = "https://service-now.com/api/now/table/sc_req_item/$_"
        $contentType = "application/json"
        $json = '{"state":"3", "close_notes":"Prezado solicitante, usuario ' + $username + ' ja existente."}'
        $result = Invoke-RestMethod -Method PUT -Uri $url -ContentType $contentType -Headers $headers -Body $json;
    }
    else {
        
        
        #Criacao do Usuario no AD
        Import-Module ActiveDirectory
        New-AdUser -DisplayName $nomecompleto -Name $nomecompleto -GivenName $nome -SurName $sobrenome -Accountpassword ($password) -EmailAddress $mail -Description $description -SamAccountName $username -ChangePasswordAtLogon:$false -UserPrincipalName $UPN -Enabled:$true -Path $OU
        
        #Adiciona o usu�rio nos grupos
        Add-AdGroupMember -Identity $grupo -Members $username
        Add-ADGroupMember -Identity $vpn -Members $username
        
        
        #Envia E-mail com os dados
        $From = "lopesgi007@gmail.com"
        $To = $mail
        $Subject = "Detalhes sobre o ticket " + $request_details.result.number
           $Body = "Corpo do email em HTML"
           #inserção de anexos
           $Anexo = @("C:\Anexo aqui.pdf")
           $SMTPServer = "definicao do servidor SMTP aqui"
           $SMTPPORT = 25
           Send-MailMessage -From $From -To $mail -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -Attachments $Anexo -Encoding UTF8
        
        #fechar o chamado:
        $url = "https://service-now.com/api/now/table/sc_req_item/$_"
        $contentType = "application/json"
        $json = '{"state":"3"}'
        $result = Invoke-RestMethod -Method PUT -Uri $url -ContentType $contentType -Headers $headers -Body $json;
        
        echo "Usuario criado com sucesso: $username"
    }
}