<?php

require_once "Classes/PHPExcel.php";
	
class Excel
{
	public function generate($emails = '', $inbox = ''){
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
}
?>