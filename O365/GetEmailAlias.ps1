get-mailbox | foreach{ 

 $host.UI.Write("White", $host.UI.RawUI.BackGroundColor, "`nUser Name: " + $_.DisplayName+"`n")

 for ($i=0;$i -lt $_.EmailAddresses.Count; $i++)
 {
    $address = $_.EmailAddresses[$i]
    
    $host.UI.Write("White", $host.UI.RawUI.BackGroundColor, $address.AddressString.ToString()+"`t")
 
    if ($address.IsPrimaryAddress)
    { 
    	$host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "Primary Email Address`n")
    }
   else
   {
    	$host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "Alias`n")
    }
 }
}