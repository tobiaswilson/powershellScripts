import-module activedirectory
$givenname = Read-Host -Prompt 'What is their first name?'
$surname = Read-Host -Prompt 'What is their second name?'
$confirmed = 1
while($confirmed)
{
$password = Read-Host  -Prompt 'Enter Password' -AsSecureString
$password2= Read-Host  -Prompt 'Enter Password Again' -AsSecureString
$pwd1_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
$pwd2_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password2))
if($pwd1_text -eq $pwd2_text){
$confirmed = 0
}
else{$a = new-object -comobject wscript.shell

$b = $a.popup(“Try again, you muppet “,0,“Passwords did not match”,1)}
}

$username = $givenname + $surname.substring(0,1)
$displayName = $givenname + " " + $surname



-path = "OU=x (e.g london watchguard")

