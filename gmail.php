<!doctype>
<html>
<head>
</head>
<body>
1) Logged in via browser to gmail account.<br/>
2) Open this url "https://www.google.com/settings/security/lesssecureapps"<br/>
3) Click at "turn on"<br/>
4) try following code <br/>

<form action="" method="post">
<input type="hidden" name="excel" value="EXCEL"">
<input type="Submit" value="EXPORT TO EXCEL">
</form>
<br/>
<?php
set_time_limit(4000);
require_once "excel.php";
// Connect to gmail
$imapPath = '{imap.gmail.com:993/imap/ssl}INBOX';
//$imapPath = '{imap.gmail.com:993/imap/ssl}[Gmail]/Trash';
$username = 'email';
$password = 'password';
 
// try to connect
$inbox = imap_open($imapPath,$username,$password) or die('Cannot connect to Gmail: ' . imap_last_error());
//$emails = imap_search($inbox, 'SUBJECT "importir.net"');
$emails = imap_search($inbox, 'ALL');

if(!empty($_POST)){
	$excel = new Excel();
	$excel->generate($emails, $inbox);
}

$output = '';
$no = 1;
foreach($emails as $mail) {
    
    $headerInfo = imap_headerinfo($inbox,$mail);
	
	$output .= '=======================================<b>'.$no.'</b>===============================================================<br><br/>';
	$output .= '<b>FROM &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: </b>'. $headerInfo->fromaddress.'<br/>';
	$output .= '<b>TO &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: </b>'. $headerInfo->toaddress.'<br/>';
    $output .= '<b>SUBJECT &nbsp;&nbsp;&nbsp;&nbsp;: </b>'. $headerInfo->subject.'<br/>';
    $output .= '<b>TANGGAL &nbsp;&nbsp;: </b>'.$headerInfo->date.'<br/>';
	
	$emailStructure = imap_fetchstructure($inbox,$mail);
	$msg = imap_fetchbody($inbox,$mail,1.1);
	
    if ($msg == "") { 
        $output .= '<b>BODY &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: </b>'.imap_fetchbody($inbox, $mail, 1).'<br><br/>';
    }
	$output .= '=======================================================================================================<br><br/>';

   echo $output;
   $output = '';
   $no++;
}
 
// colse the connection
imap_expunge($inbox);
imap_close($inbox);
?>

</body>
</html>