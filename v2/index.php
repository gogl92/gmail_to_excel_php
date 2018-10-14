<!doctype>
<html>
<head>
</head>
<body>

<?php
set_time_limit(4000);
// Connect to gmail
$imapPath = '{imap.gmail.com:993/imap/ssl}INBOX';
//$imapPath = '{imap.gmail.com:993/imap/ssl}[Gmail]/Trash';
$username = '';
$password = '';

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
$result = [];
foreach($emails as $key => $mail) {

  if($key==2){
    break;
  }
    $headerInfo = imap_headerinfo($inbox,$mail);
	echo $headerInfo->fromaddress.' </br>';
	echo $headerInfo->toaddress.' </br>';
	echo $headerInfo->subject.' </br>';
	echo $headerInfo->date.' </br>';
	$emailStructure = imap_fetchstructure($inbox,$mail);
	$msg = imap_fetchbody($inbox,$mail,1.1).' </br>';
	echo $msg.' </br>';
	echo $key.' </br>';
	$result[] = ['from'=>' '.$headerInfo->fromaddress,'to'=>' '.$headerInfo->toaddress] ;
}
include_once('../common/ExcelHelper.php');
$excelHelper = new ExcelHelper();
$excelHelper->createExportTable($result, [
    ['coordinate' => 'A1', 'title' => 'From'],
    ['coordinate' => 'B1', 'title' => 'To'],
]);
$excelHelper->saveExcel('files', 'Test');

// colse the connection
imap_expunge($inbox);
imap_close($inbox);
?>

</body>
</html>
