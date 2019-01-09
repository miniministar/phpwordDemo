<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

$servername = "localhost";
$username   = "root";
$password   = "2222";

try {
    $conn = new PDO("mysql:host=$servername;dbname=nxdb_ty", $username, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $conn->exec("SET CHARACTER SET utf8");
} catch(PDOException $e) {
    echo "Connection failed: " . $e->getMessage();
}

//database Data
$stmt = $conn->prepare('SELECT ID as userId , USERNAME as  userFirstName , NAME as userName ,  TELEPHONE as  userPhone FROM nc_user WHERE id = ? or id = ? or id = ?');
$stmt->execute([1, 2, 3]);
$users = $stmt->fetchAll(PDO::FETCH_ASSOC);

if (empty($users)) {
    exit('查无记录');
}

require_once __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

use PhpOffice\PhpWord\Settings;
Settings::loadConfig();
$templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('template-table.docx');

// Variables on different parts of document
$templateProcessor->setValue('weekday', date('l'));            // On section/content
$templateProcessor->setValue('time', date('H:i'));             // On footer
$templateProcessor->setValue('serverName', realpath(__DIR__)); // On header

$templateProcessor->cloneRow('rowValue', 10);
$templateProcessor->setValue('rowValue#2', 'Mercury');
$templateProcessor->setValue('rowValue#1', 'Sun');
$templateProcessor->setValue('rowValue#3', 'Venus');
$templateProcessor->setValue('rowValue#4', 'Earth');
$templateProcessor->setValue('rowValue#5', 'Mars');
$templateProcessor->setValue('rowValue#6', 'Jupiter');
$templateProcessor->setValue('rowValue#7', 'Saturn');
$templateProcessor->setValue('rowValue#8', 'Uranus');
$templateProcessor->setValue('rowValue#9', 'Neptun');
$templateProcessor->setValue('rowValue#10', 'Pluto');

$templateProcessor->setValue('rowNumber#2', '2');
$templateProcessor->setValue('rowNumber#1', '1');
$templateProcessor->setValue('rowNumber#3', '3');
$templateProcessor->setValue('rowNumber#4', '4');
$templateProcessor->setValue('rowNumber#5', '5');
$templateProcessor->setValue('rowNumber#6', '6');
$templateProcessor->setValue('rowNumber#7', '7');
$templateProcessor->setValue('rowNumber#8', '8');
$templateProcessor->setValue('rowNumber#9', '9');
$templateProcessor->setValue('rowNumber#10', '10');


// Table with a spanned cell

$templateProcessor->cloneRow('userId', count($users));

foreach ($users as $rowKey => $rowData) {
    $rowNumber = $rowKey + 1;
    foreach ($rowData as $macro => $replace) {
        $templateProcessor->setValue($macro . '#' . $rowNumber, $replace);
    }
}

$pathName = time() . '.docx';
$templateProcessor->saveAs($pathName);

$fileName = '用户情况一览表.docx';
$fileName = iconv('utf-8', 'gb2312', $fileName);

$contentType = 'Content-type: application/vnd.openxmlformats-officedocument.wordprocessingml.document';
header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
header ( "Cache-Control: no-cache, must-revalidate" );
header ( "Pragma: no-cache" );
header ( $contentType );
header ( "Content-Disposition: attachment; filename=".$fileName);
readfile($pathName);
unlink($pathName);