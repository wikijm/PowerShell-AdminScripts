
# test server-port reachability 
test-netconnection <IP Address> -port <Port number>

# test server-port reachability on continuous bases

while ($true) {test-netconnection <IP Address> -port <Port number> | Format-Table @{n='Timestamp';e={Get-DAte}},tcptestsucceeded}
