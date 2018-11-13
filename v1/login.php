<?php
session_start();
require_once '../vendor/autoload.php';
require_once('includes/config.php');
use GmailWrapper\Authenticate;
$authenticate = Authenticate::getInstance(CLIENT_ID,CLIENT_SECRET,APPLICATION_NAME,DEVELOPER_KEY);
if(!$authenticate->isAuthenticated()) {
    $response = $authenticate->getLogInURL('http://127.0.0.1:7979/login.php', ['openid','https://www.googleapis.com/auth/gmail.readonly','https://mail.google.com/','https://www.googleapis.com/auth/gmail.modify','https://www.googleapis.com/auth/gmail.compose','https://www.googleapis.com/auth/gmail.send'],'offline', 'force');
    if(!$response['status']) {
        echo $response['message'];
        exit;
    }
    $loginUrl = $response['data'];
    echo "<a href='{$loginUrl}'>Login</a>";
}
if(isset($_GET['code'])) {
    $auth = $authenticate->logIn($_GET['code']);
    if($auth['status']) {
        $_SESSION['tokens'] = $authenticate->getTokens();
        echo '<pre>';
        var_dump($authenticate->getUserDetails());
    } else {
        echo $auth['message'];
    }
}
