$kind = "video"  
$folder = 'D:\Music'

$objConnection = New-Object -com ADODB.Connection  
$objRecordSet = New-Object -com ADODB.Recordset  
$objConnection.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")  
$objRecordSet.Open("SELECT System.ItemPathDisplay FROM SYSTEMINDEX WHERE System.Kind = '$kind' AND System.ItemPathDisplay LIKE '$folder\%'", $objConnection)  
if ($objRecordSet.EOF -eq $false) {$objRecordSet.MoveFirst() }  

while ($objRecordset.EOF -ne $true) {  
  $objRecordset.Fields.Item("System.ItemPathDisplay").Value  
  $objRecordset.MoveNext()  
} 