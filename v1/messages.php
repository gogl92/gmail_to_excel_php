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
if (!$messageList['status']) {
    echo $messageList['message'];
    exit;
}
foreach ($messageList['data'] as $key => $value) {
    if ($key == 10) {
        break;
    }
    $messageDetails = $msgs->getMessageDetails($value->id);
    if ($messageDetails['status'] == 1) {
        print_r($messageDetails['data']);
        $id = $value->id;
        $from = $messageDetails['data']['headers']['From'];
        $to = isset($messageDetails['data']['headers']['Delivered-To']) ? $messageDetails['data']['headers']['Delivered-To'] : '';
        $cc = isset($messageDetails['data']['headers']['CC']) ? $messageDetails['data']['headers']['CC'] : '';
        $bcc = isset($messageDetails['data']['headers']['BCC']) ? $messageDetails['data']['headers']['BCC'] : '';
        $subject = isset($messageDetails['data']['headers']['Subject']) ? $messageDetails['data']['headers']['Subject'] : '';
        $body = "";
        foreach ($messageDetails['data']['body']['text/plain'] as $value) {
            $body .= $value;
        }

        foreach ($messageDetails['data']['body']['text/html'] as $value) {
            $body .= $value;
        }

        $body .= isset($messageDetails['data']['body']['snippet']) ? $body .= $messageDetails['data']['body']['snippet'] : '';

        $toknized = explode('</br>', $body);

        $marca = 'N/A';
        foreach ($toknized as $innerKey => $item) {
            if ($innerKey === 0 || startsWith($item, 'Nueva Cotizacion de Energia Verde RMS')) {
                continue;
            }
            if ($innerKey === 1 || startsWith($item, '=============================')) {
                continue;
            }
            if ($innerKey === 2 || startsWith($item, 'Marca:')) {
                $marca = $item;
            }
        }

        $result[] = ['ID' => $id, 'From' => $from, 'To' => $to, 'CC' => $cc, 'BCC' => $bcc, 'Subject' => $subject, 'Marca' => $marca, 'Body' => $body];
    }
}
$excelHelper->createExportTable($result, [
    ['coordinate' => 'A1', 'title' => 'ID'],
    ['coordinate' => 'B1', 'title' => 'Remitente'],
    ['coordinate' => 'C1', 'title' => 'Destinatario'],
    ['coordinate' => 'D1', 'title' => 'Copia'],
    ['coordinate' => 'E1', 'title' => 'Copia Oculta'],
    ['coordinate' => 'F1', 'title' => 'Modelo'],
    ['coordinate' => 'G1', 'title' => 'Marca'],
    ['coordinate' => 'H1', 'title' => 'Nombre'],
    ['coordinate' => 'I1', 'title' => 'Nombre de la Empresa'],
    ['coordinate' => 'J1', 'title' => 'Teléfono'],
    ['coordinate' => 'K1', 'title' => 'Tipo de Empresa'],
    ['coordinate' => 'L1', 'title' => 'Estado'],
    ['coordinate' => 'M1', 'title' => 'País'],
    ['coordinate' => 'N1', 'title' => 'Correo Electrónico'],
    ['coordinate' => 'O1', 'title' => 'Mensaje'],
]);

$excelHelper->saveExcel('files', 'Test2');

$nextToken = $messageList['nextToken'];
echo '<p><a href="messages.php?pageToken=' . $nextToken . '">Next</a></p>';


/**
 * @param $haystack
 * @param $needle
 * @return bool
 */
function startsWith($haystack, $needle)
{
    $length = strlen($needle);
    return (substr($haystack, 0, $length) === $needle);
}

/**
 * @param $haystack
 * @param $needle
 * @return bool
 */
function endsWith($haystack, $needle)
{
    $length = strlen($needle);
    return $length === 0 ||
        (substr($haystack, -$length) === $needle);
}