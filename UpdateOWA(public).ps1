# connect to exchange
Connect-ExchangeOnline -UserPrincipalName {username@domain} -ShowProgress $true

# read list of people we need to update
$userPath = "C:\Users\$env:USERNAME\Documents\CustomScripts\OWA\userlist.txt"
$userList = Get-Content -Path $userPath

# for each user in the file
Foreach($username in $userList){

    # generate signature file
    if (!($user = Get-User $username)){
        Write-Host "$username does not exist, skipping"
        Continue
    }
    $save_location = "C:\Users\$env:USERNAME\Documents\CustomScripts\OWA\Signatures"
    #Signature information
    $full_name = $($user.DisplayName)
    # $account_name = $($user.UserPrincipalName) we don’t need account name for this script
    $job_title = $($user.Title)
    # $location = $($user.office) we don’t need location for now
    $comp = “CompanyName” # The company name
    $email = $($user.WindowsEmailAddress)
    $phone = $($user.Phone)
    # $logo = “C:/Path-to-photo” if they ever wanted to add logos you would have to update the HTML

    # Generating the unique signature name
    $sign = $username.split("@")[0]
    $signatureFile = “signature-$sign.htm”
    $signatureName = "signature-$sign"

    $output_file = "$save_location\$signatureFile"

    Write-Host "Generating HTML for: $user"
    $HTML = "<span style=`"font-family: calibri,sans-serif;`"><p style=`"margin: 0cm;`"><strong>$full_name<span style=`"color: #ffda1a;`"> | </span></strong>$job_title</p><p style=`"margin: 0cm;`">$comp</p><p style=`"margin: 0cm;`"><a href=`"mailto:$email`">$email</a></p><p style=`"margin: 0cm;`">$phone</p></span><br>"
    $HTML | Out-File $output_file

    # update signature file
    Write-Host "Setting signature for " $user
    set-mailboxmessageconfiguration -identity $username -signaturehtml (get-content $output_file) -autoaddsignature $true -AutoAddSignatureOnReply $true -AutoAddSignatureOnMobile $true

    #if you want to delete the signature
    #Set-MailboxMessageConfiguration -identity $_.alias -SignatureHtml "" -SignatureText "" -SignatureTextOnMobile ""
}

Disconnect-ExchangeOnline