cls

<#--------------------------------------------
    Kreditkarte
--------------------------------------------#>
$cardtype = "Mastercard" #"Visa"#"Diners"#"Discover"
$cardpan = "" #Kartennummer
$cardexpiremonth = ""
$cardexpireyear = ""
$cardcvc = ""

<#--------------------------------------------
    NBB-Login
--------------------------------------------#>
$user_email = "" #read-host "user_email eingeben"
$user_password = "" #read-host " eingeben"

<#--------------------------------------------
    Ziel-Adresse
--------------------------------------------#>
$first_name = "" #read-host "first_name eingeben"
$last_name = "" #read-host "last_name eingeben"
$street = "" #read-host "street eingeben"
$street_number = "" #read-host "street_number eingeben"
$postal_code = "" #read-host "postal_code eingeben"
$city = "" #read-host "city eingeben"

<#--------------------------------------------
    NBB-Gutschein-Code
--------------------------------------------#>
$gutschein_code = "" #read-host "gutschein_code eingeben"

<#--------------------------------------------
    NBB-Produkt-URL
--------------------------------------------#>
$URL = "https://www.notebooksbilliger.de/hp+15s+eq2147ng+693961/incrpc/topprod" #read-host "NBB URL eingeben"

<#--------------------------------------------
    Browser
--------------------------------------------#>
$try = 0
while($try -lt 5){
    try{
        Write-Host "Browser wird gestartet" -ForegroundColor Yellow
        $i_1 = New-Object -ComObject 'internetExplorer.Application' -ErrorAction Ignore -ErrorVariable global:Fehler
        $i_1.Visible = $true
        $i_1.Navigate("https://www.notebooksbilliger.de/index.php/action/login")
        do{Sleep 1}until($i_1.Busy -eq $false)
    }catch{
        Write-Host "Browser konnte nicht gestartet werden" -ForegroundColor Red
        $try++
        sleep 10
    }
    if($i_1 -ne 0){
        $try = 6
        Write-Host "Browser wurde gestartet" -ForegroundColor Green
    }
    sleep 1
}

<#--------------------------------------------
    Cookiebanner
--------------------------------------------#>
try{
    $btn_cookie = $i_1.Document.IHTMLDocument3_getElementById("uc-btn-accept-banner")
    $btn_cookie.click()
}catch{}
do{Sleep 1}until($i_1.Busy -eq $false);sleep 1

<#--------------------------------------------
    Login
--------------------------------------------#>
$try = 0
while($try -lt 3){
    try{
        Write-Host "Warte auf Login" -ForegroundColor Yellow
        $email_address = $i_1.Document.IHTMLDocument3_getElementById("layerEmailAddress")
        $email_address.click()
        $email_address.value = $user_email
        $password = $i_1.Document.IHTMLDocument3_getElementById("layerPassword")
        $password.click()
        $password.value = $user_password
        $btn_login = $i_1.Document.IHTMLDocument3_getElementById("layerLoginButton")
        $btn_login.click()
        do{Sleep 1}until($i_1.LocationURL -eq "https://www.notebooksbilliger.de/index.php");sleep 1
    }catch{}
    try{
        $Kundenkonto = $i_1.Document.body.getElementsByClassName("nbx-text-btn-base")
        if($Kundenkonto[0].innerText.Length -gt 2){
            Write-Host "Login erfolgreich" -ForegroundColor Green
            $try = 4
        }
    }catch{
            Write-Host "Login gescheitert" -ForegroundColor Red
            $try ++
    }
    $login_layer_validation_error = ""
    try{
        $login_layer_validation_error = $i_1.Document.IHTMLDocument3_getElementById("loginLayerValidation")
        if($login_layer_validation_error.innerText.Length -gt 2){
            write-host $login_layer_validation_error.innerText -ForegroundColor Red
        }
    }catch{}
}

<#--------------------------------------------
    Zum Warenkorb
--------------------------------------------#>
try{
    $i_1.Navigate2("https://www.notebooksbilliger.de/warenkorb")
    Write-Host "Warenkorb wird geladen" -ForegroundColor Yellow
    do{Sleep 1}until($i_1.Busy -eq $false);sleep 4
}catch{}

<#--------------------------------------------
    Aktuellen Warenkorb löschen
--------------------------------------------#>
$try = 0
while($try -lt 3){
    try{
    $count = 0
    $cart_items = $i_1.Document.body.getElementsByClassName("shopping-cart__delete-cta")
    if($cart_items.length -gt 0){
    $count = $cart_items.length
    Write-Host "Warenkorb wird geleert" -ForegroundColor Yellow
        foreach($cart_item in $cart_items){
            $cart_item.click()
            do{Sleep 1}until($i_1.Busy -eq $false)
        }
    }
    do{Sleep 1}until($i_1.Busy -eq $false);sleep 5
    }catch{$try++}
    $cart_items_2 = $i_1.Document.body.getElementsByClassName("empty_cart")
    if($cart_items_2[0].innertext -match "Zur Zeit befinden sich keine Produkte im Warenkorb."){
        Write-Host $count, "Artikel aus Warenkorb entfernt" -ForegroundColor Green
        $try = 4
    }
}

<#--------------------------------------------
    Produktseite
--------------------------------------------#>
$i_1.Navigate2($URL)
Write-Host "Produktseite wird geladen" -ForegroundColor Yellow
do{Sleep 1}until($i_1.Busy -eq $false);sleep 2
Write-Host "Ladevorgang abgeschlossen" -ForegroundColor Green

<#--------------------------------------------
    Dem Warenkorb hinzufügen
--------------------------------------------#>
do{Sleep 1}until($i_1.LocationURL -eq $url -and $i_1.Busy -eq $false)
$try = 0
while($try -lt 3){
    try{
    $btn_check_out = $i_1.Document.body.getElementsByClassName("js-pdp-head-add-to-cart")
    $btn_check_out[0].click()
    Write-Host "Lade Produkt in Warenkorb" -ForegroundColor Yellow
    do{Sleep 1}until($i_1.Busy -eq $false);sleep 2
    $try = 4
    }catch{
        Write-Host "Nicht im Warenkorb" -ForegroundColor Red
        sleep 4
        $try++
    }
}
do{Sleep 1}until($i_1.Busy -eq $false)

<#--------------------------------------------
    Zum Warenkorb
--------------------------------------------#>
$try = 0
while($try -lt 5){
    try{
        $i_1.Navigate("https://www.notebooksbilliger.de/warenkorb")
        Write-Host "Warenkorb wird geladen" -ForegroundColor Yellow
        do{Sleep 1}until($i_1.Busy -eq $false);sleep 1
        if($i_1.LocationURL -eq "https://www.notebooksbilliger.de/warenkorb"){
            Write-Host "Ladevorgang abgeschlossen" -ForegroundColor Green
            $try = 6
        }
        else{
            Write-Host "Warenkorb wurde nicht geladen" -ForegroundColor Red
            sleep 4
        }
    }catch{
        Write-Host "Warenkorb wurde nicht geladen" -ForegroundColor Red
        sleep 4
        $try++
    }
}
do{Sleep 1}until($i_1.Busy -eq $false)

<#--------------------------------------------
    Zur Kasse
--------------------------------------------#>
$try = 0
while($try -lt 5){
    try{
        $i_1.Navigate("https://www.notebooksbilliger.de/kasse")
        Write-Host "Kasse wird geladen" -ForegroundColor Yellow
        do{Sleep 1}until($i_1.Busy -eq $false);sleep 1
        if($i_1.LocationURL -eq "https://www.notebooksbilliger.de/kasse"){
            Write-Host "Ladevorgang abgeschlossen" -ForegroundColor Green
            $try = 6
        }
        else{
            sleep 4
            Write-Host "Kasse wurde nicht geladen" -ForegroundColor Red
        }
    }catch{
        Write-Host "Kasse wurde nicht geladen" -ForegroundColor Red
        sleep 4
        $try++
    }
}

<#--------------------------------------------
    Adresse ändern
--------------------------------------------#>
#Write-Host "Adresse wird angepasst" -ForegroundColor Yellow
#$btn_changedelivery = $i_1.Document.body.getElementsByClassName("changedelivery-text")
#$btn_changedelivery[0].click()

<#--------------------------------------------
    Zahlungsmethode wählen
--------------------------------------------#>
$try = 0
while($try -lt 4){
    try{
    Write-Host "Zahlungsmethode wird angepasst" -ForegroundColor Yellow
    $radio_payment_paycreditcard = $i_1.Document.IHTMLDocument3_getElementById("paycreditcard")
    $radio_payment_paycreditcard.checked = "checked"
    $radio_payment_paycreditcard.click()
    }catch{
    Write-Host "Zahlungsmethode wurde nicht angepasst" -ForegroundColor Red
    $try++
    sleep 3
    }
    $radio_payment_paycreditcard = $i_1.Document.IHTMLDocument3_getElementById("paycreditcard")
    if($radio_payment_paycreditcard.checked -eq "checked"){
        $try = 5
        Write-Host "Zahlungsmethode wurde angepasst" -ForegroundColor Green
    }
}
do{Sleep 1}until($i_1.Busy -eq $false);sleep 1

<#--------------------------------------------
    AGB annehmen & Weiter
--------------------------------------------#>
try{
$checkbox = $i_1.Document.IHTMLDocument3_getElementById("conditions")
$checkbox.checked = "checked"
Write-Host "AGB wird akzeptiert" -ForegroundColor Yellow
$checkbox.click()
}catch{}
do{Sleep 1}until($i_1.Busy -eq $false);sleep 1
try{
$btn_submit = $i_1.Document.body.getElementsByClassName("nbx-btn-disabled")
Write-Host "Weiterleitung zur Zahlungsmitteleingabe" -ForegroundColor Yellow
$btn_submit[0].click()
}catch{}
try{
do{Sleep 1}until($i_1.Busy -eq $false);sleep 1
$summary_subtotal = $i_1.Document.body.getElementsByClassName("summary subtotal gray")
Write-Host $summary_subtotal[0].innerText[0..14]
$btn_checkout_submit = $i_1.Document.IHTMLDocument3_getElementById("checkout_submit")
$btn_checkout_submit.click()
do{Sleep 1}until($i_1.Busy -eq $false);sleep 1
Write-Host "Weiterleitung erfolgreich" -ForegroundColor Green
}catch{}


<#--------------------------------------------
    Zahlung - Script dose not work at this point - Cant find the right iFrame - send solutions to admin@miruth.de
--------------------------------------------#>
#Write-Host "Zahlungsmittel werden eingegeben" -ForegroundColor Yellow
#$box_cardtype = $i_1.Document.IHTMLDocument3_getElementById("cardtype")
#$box_cardtype.value = $cardtype
#$box_cardpan = $i_1.Document.IHTMLDocument3_getElementById("cardpan")
#$box_cardpan.value = $cardpan
#$box_cardexpiremonth = $i_1.Document.IHTMLDocument3_getElementById("cardexpiremonth")
#$box_cardexpiremonth.value = $cardexpiremonth
#$box_cardexpireyear = $i_1.Document.IHTMLDocument3_getElementById("cardexpireyear")
#$box_cardexpireyear.value = $cardexpireyear
#$box_cardcvc = $i_1.Document.IHTMLDocument3_getElementById("cardcvc2")
#$box_cardcvc.value = $cardcvc
#$btn_checkoutCreditCardSubmit = $i_1.Document.IHTMLDocument3_getElementById("checkoutCreditCardSubmit")
#$btn_checkoutCreditCardSubmit.click()
#$i_1.quit()
