<?php
require_once '../vendor/autoload.php';
require_once('includes/config.sample.php');
require_once('includes/auth.php');
include_once('../common/ExcelHelper.php');

use GmailWrapper\Messages;
$excelHelper = new ExcelHelper();

$pageToken = isset($_GET['pageToken']) ? $_GET['pageToken'] : null;
$msgs = new Messages($authenticate);
$messageList = $msgs->getMessages([], $pageToken);
if(!$messageList['status']) {
    echo $messageList['message'];
    exit;
}
foreach ($messageList['data'] as $key => $value) {
  if($key==10){
    break;
  }
  $messageDetails = $msgs->getMessageDetails($value->id);
  if($messageDetails['status']==1){
    print_r($messageDetails['data']);
    $id = $value->id;
    $from = $messageDetails['data']['headers']['From'];
    $to = isset($messageDetails['data']['headers']['Delivered-To']) ? $messageDetails['data']['headers']['Delivered-To']:'' ;
    $cc = isset($messageDetails['data']['headers']['CC']) ? $messageDetails['data']['headers']['CC']:'' ;
    $bcc = isset($messageDetails['data']['headers']['BCC']) ? $messageDetails['data']['headers']['BCC']:'';
    $subject = isset($messageDetails['data']['headers']['Subject']) ? $messageDetails['data']['headers']['Subject']:'';
    $body = "";
    foreach ($messageDetails['data']['body']['text/plain'] as $value) {
        $body .= $value;
    }

    foreach ($messageDetails['data']['body']['text/html'] as $value) {
        $body .= $value;
    }

    $body .= isset($messageDetails['data']['body']['snippet']) ? $body .= $messageDetails['data']['body']['snippet'] : '';
    echo $body;
    $result[] = ['ID'=>$id, 'From'=>$from, 'To'=>$to, 'CC'=>$cc,'BCC'=>$bcc,'Subject'=>$subject,'Body'=>$body];
  }
}
$excelHelper->createExportTable($result, [
    ['coordinate' => 'A1', 'title' => 'ID'],
    ['coordinate' => 'B1', 'title' => 'From'],
    ['coordinate' => 'C1', 'title' => 'To'],
    ['coordinate' => 'D1', 'title' => 'CC'],
    ['coordinate' => 'E1', 'title' => 'BCC'],
    ['coordinate' => 'F1', 'title' => 'Subject'],
    ['coordinate' => 'G1', 'title' => 'Body'],
]);
$excelHelper->saveExcel('files', 'Test2');

$nextToken = $messageList['nextToken'];
echo '<p><a href="messages.php?pageToken='.$nextToken.'">Next</a></p>';
