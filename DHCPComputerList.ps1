$ScopeList = (netsh dhcp server 192.168.10.15 show scope) #where 192.168.10.15 is your DHCP server
$Scopes = $ScopeList[5..($ScopeList.Length -4)]

Foreach($Line in $Scopes){

                         $Scope = $Line.split(" ")
                         $Subnet =$Scope[1]
                         
                         $SubnetClients = (netsh dhcp server 192.168.10.15 scope $Subnet show clients 1)
                         
                         $ClientList = $SubnetClients[8..($SubnetClients.Length -5)]
                         
                         Foreach($Client in $ClientList){
                                                        $test = ($Client -match "-D-")
                                                        if( $test -like "True"){
                                                                               $Client = $Client.Split(" ")
                                                                               $IP = $Client[0]
                                                                               $IP | Out-File -FilePath ./DHCPComputer.txt -Append
                                                                               $ClientName = $Client[$Client.Length-1]
                                                                               $ClientName | Out-File -FilePath ./DHCPName.txt -Append
                                                                               }
                                                        }
}
