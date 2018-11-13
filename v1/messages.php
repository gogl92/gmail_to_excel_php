<?php
require_once '../vendor/autoload.php';
require_once('includes/config.php');
require_once('includes/auth.php');
include_once('../common/ExcelHelper.php');

use GmailWrapper\Messages;
const ASUNTO = 'Prueba';
$excelHelper = new ExcelHelper();

$pageToken = isset($_GET['pageToken']) ? $_GET['pageToken'] : null;
$msgs = new Messages($authenticate);
$messageList = $msgs->getMessages([], $pageToken);
foreach ($messageList['data'] as $key => $value) {
    $messageDetails = $msgs->getMessageDetails($value->id);
    if ($messageDetails['status'] == 1) {
        $id = $value->id;
        $from = $messageDetails['data']['headers']['From'];
        $to = isset($messageDetails['data']['headers']['Delivered-To']) ? $messageDetails['data']['headers']['Delivered-To'] : '';
        $cc = isset($messageDetails['data']['headers']['CC']) ? $messageDetails['data']['headers']['CC'] : '';
        $bcc = isset($messageDetails['data']['headers']['BCC']) ? $messageDetails['data']['headers']['BCC'] : '';
        $subject = isset($messageDetails['data']['headers']['Subject']) ? $messageDetails['data']['headers']['Subject'] : '';
        if ($subject !== ASUNTO) {
            continue;
        }
        $body = '';
        foreach ($messageDetails['data']['body']['text/plain'] as $value) {
            $body .= $value;
        }

        foreach ($messageDetails['data']['body']['text/html'] as $value) {
            $body .= $value;
        }

        $body .= isset($messageDetails['data']['body']['snippet']) ? $body .= $messageDetails['data']['body']['snippet'] : '';

        $toknized = explode('<br />', $body);

        $modelo = 'N/A';
        $marca = 'N/A';
        $nombre = 'N/A';
        $nombre_empresa = 'N/A';
        $telefono = 'N/A';
        $tipo_empresa = 'N/A';
        $otro_tipo_empresa = 'N/A';
        $estado = 'N/A';
        $pais = 'N/A';
        $correo_electronico = 'N/A';
        $mensaje = 'N/A';

        foreach ($toknized as $innerKey => $item) {
            if ($innerKey === 0 || startsWith($item, 'Nueva Cotizacion de Energia Verde RMS')) {
                continue;
            }
            if ($innerKey === 1 || startsWith($item, '=============================')) {
                continue;
            }
            if ($innerKey === 2 || startsWith($item, 'Modelo: ')) {
                $modelo = substr($item, strlen('Modelo: ') + 2);
            }
            if ($innerKey === 3 || startsWith($item, 'Marca: ')) {
                $marca = substr($item, strlen('Marca: ') + 2);
            }
            if ($innerKey === 4 || startsWith($item, 'Nombre: ')) {
                $nombre = substr($item, strlen('Nombre: ') + 2);
            }
            if ($innerKey === 5 || startsWith($item, 'Nombre de la Empresa: ')) {
                $nombre_empresa = substr($item, strlen('Nombre de la Empresa: ') + 2);
            }
            if ($innerKey === 6 || startsWith($item, 'Telefono: ')) {
                $telefono = ' ' . substr($item, strlen('Telefono: ') + 2);
            }
            if ($innerKey === 7 || startsWith($item, 'Tipo de Empresa: ')) {
                $tipo_empresa = substr($item, strlen('Tipo de Empresa: ') + 2);
            }
            if ($innerKey === 8 || startsWith($item, 'Otro Tipo de Empresa: ')) {
                $otro_tipo_empresa = substr($item, strlen('Otro Tipo de Empresa: ') + 2);
            }
            if ($innerKey === 9 || startsWith($item, 'Estado: ')) {
                $estado = substr($item, strlen('Estado: ') + 2);
            }
            if ($innerKey === 10 || startsWith($item, 'Otro País: ')) {
                $pais = substr($item, strlen('Otro País: ') + 2);
            }
            if ($innerKey === 11 || startsWith($item, 'Correo Electrónico: ')) {
                $correo_electronico = substr($item, strlen('Correo Electrónico: ') + 2);
            }
            if ($innerKey === 11 || startsWith($item, 'Mensaje: ')) {
                $mensaje = substr($item, strlen('Mensaje: ') + 2);
            }
        }

        $result[] = [
            'ID' => $id,
            'From' => $from,
            'To' => $to,
            'CC' => $cc,
            'BCC' => $bcc,
            'Subject' => $subject,
            'Modelo' => $modelo,
            'Marca' => $marca,
            'Nombre' => $nombre,
            'NombreEmpresa' => $nombre_empresa,
            'Telefono' => $telefono,
            'TipoEmpresa' => $tipo_empresa,
            'OtroTipoEmpresa' => $otro_tipo_empresa,
            'Estado' => $estado,
            'OtroPais' => $pais,
            'CorreoElectronico' => $correo_electronico,
            'Mensaje' => $mensaje
        ];
    }
}
$excelHelper->createExportTable($result, [
    ['coordinate' => 'A1', 'title' => 'ID'],
    ['coordinate' => 'B1', 'title' => 'Remitente'],
    ['coordinate' => 'C1', 'title' => 'Destinatario'],
    ['coordinate' => 'D1', 'title' => 'Copia'],
    ['coordinate' => 'E1', 'title' => 'Copia Oculta'],
    ['coordinate' => 'F1', 'title' => 'Asunto'],
    ['coordinate' => 'G1', 'title' => 'Modelo'],
    ['coordinate' => 'H1', 'title' => 'Marca'],
    ['coordinate' => 'I1', 'title' => 'Nombre'],
    ['coordinate' => 'J1', 'title' => 'Nombre de la Empresa'],
    ['coordinate' => 'K1', 'title' => 'Télefono'],
    ['coordinate' => 'L1', 'title' => 'Tipo de Empresa'],
    ['coordinate' => 'M1', 'title' => 'Otro Tipo de Empresa'],
    ['coordinate' => 'N1', 'title' => 'Estado'],
    ['coordinate' => 'O1', 'title' => 'País'],
    ['coordinate' => 'P1', 'title' => 'Correo Electrónico'],
    ['coordinate' => 'Q1', 'title' => 'Mensaje'],
]);

$excelHelper->hideColumns(['A', 'B', 'C', 'D', 'E', 'F']);
$excelHelper->autoSizeColumns(['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']);

$excelHelper->saveExcel('files', 'Test2');

$nextToken = $messageList['nextToken'];
echo '<p><a href="messages.php?pageToken=' . $nextToken . '">Siguiente</a></p>';


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