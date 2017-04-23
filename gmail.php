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
// Connect to gmail
//$imapPath = '{imap.gmail.com:993/imap/ssl}INBOX';
$imapPath = '{imap.gmail.com:993/imap/ssl}[Gmail]/Trash';
$username = 'email';
$password = 'password';
 
// try to connect
$inbox = imap_open($imapPath,$username,$password) or die('Cannot connect to Gmail: ' . imap_last_error());
//$emails = imap_search($inbox, 'SUBJECT "importir.net"');
$emails = imap_search($inbox, 'ALL');

if(!empty($_POST)){
	require_once "Classes/PHPExcel.php";
	$objPHPExcel = new PHPExcel();
	$objPHPExcel->setActiveSheetIndex(0);
	
	$style = array(
			'alignment' => array(
				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
			)
		);

	$border = array(
				'borders' => array(
				'allborders' => array(
				'style' => PHPExcel_Style_Border::BORDER_THIN
				)
			),
				'fill' => array(
				'type' => PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array('rgb' => 'FF0000')
				)
			);
				 
	// Add some data
	$objWorkSheet = $objPHPExcel->setActiveSheetIndex(0);
	$objWorkSheet = $objPHPExcel->getActiveSheet()->setTitle('GMAIL');

	$objWorkSheet->getStyle("A1:E1")->applyFromArray($style);
	$objWorkSheet->getStyle("A1:E1")->applyFromArray($border);

	// Set Header
	$objWorkSheet
				->setCellValue('A1', 'FROM ')
				->setCellValue('B1', 'TO')
				->setCellValue('C1', 'SUBJECT')
				->setCellValue('D1', 'TANGGAL')
				->setCellValue('E1', 'BODY');
				$objWorkSheet->getColumnDimension('A')->setWidth(25);
                $objWorkSheet->getColumnDimension('B')->setWidth(25);
                $objWorkSheet->getColumnDimension('C')->setWidth(25);
                $objWorkSheet->getColumnDimension('D')->setWidth(25);
                $objWorkSheet->getColumnDimension('E')->setWidth(70);
	
	foreach($emails as $yek => $mail) {
		$headerInfo = imap_headerinfo($inbox,$mail);
		$key = $yek + 2;
		$objPHPExcel->getActiveSheet()->SetCellValue('A'.$key, $headerInfo->fromaddress);
		$objPHPExcel->getActiveSheet()->SetCellValue('B'.$key, $headerInfo->toaddress);
		$objPHPExcel->getActiveSheet()->SetCellValue('C'.$key, $headerInfo->subject);
		$objPHPExcel->getActiveSheet()->SetCellValue('D'.$key, $headerInfo->date);
		$emailStructure = imap_fetchstructure($inbox,$mail);
		$msg = imap_fetchbody($inbox,$mail,1.1);
		if ($msg == "") { 
			$objPHPExcel->getActiveSheet()->SetCellValue('E'.$key, imap_fetchbody($inbox, $mail, 1));
		}
	}
	
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	ob_end_clean();
	// We'll be outputting an excel file
	header('Content-type: application/vnd.ms-excel');
	header('Content-Disposition: attachment; filename="gmail.xlsx"');
	$objWriter->save('php://output');exit;
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