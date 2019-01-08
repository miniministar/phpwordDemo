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

$userCode = isset($_GET['id']) ? $_GET['id'] : 1;

//database Data
$stmt = $conn->prepare('SELECT * FROM nc_user WHERE id = ?');
$stmt->execute([$userCode]);
$user = $stmt->fetch(PDO::FETCH_ASSOC);

if (empty($user)) {
    exit('查无记录');
}

require_once __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

$phpWord = new \PhpOffice\PhpWord\PhpWord();

$section = $phpWord->addSection();

//Font & style
$phpWord->addFontStyle('hFont', array('size' => 18, 'color' => '1B2232', 'bold' => true));
$phpWord->addFontStyle('bFont', array('size' => 9));
$phpWord->addParagraphStyle('pStyle', array('align' => 'center'));

// Header
$section->addText('用户信息表', 'hFont', 'pStyle');
$section->addTextBreak(1);
$section->addText('系：'  .$user['MOBILE'] . '（盖章）             专业：' . $user['MOBILE'] . '           班级：' . $user['MAJOR'] . '            学号：' . $userCode , 'bFont',  array('align' => 'left'));
$section->addTextBreak(1);

//Table
$styleTable = array('borderSize' => 6);
$phpWord->addTableStyle('myOwnTableStyle', $styleTable);

$table = $section->addTable('myOwnTableStyle');

$lefontStyle = array('align' => 'left');

$fontStyle = array('align' => 'center');
$fontStyleBold = array('align' => 'center', 'bold' => true);
$cell400Style = array('align' => 'center', 'lineHeigth' => 350, 'spaceBefore' => 100, 'spaceAfter' => 150);
$cell600Style = array('align' => 'center', 'lineHeigth' => 600, 'spaceBefore' => 300, 'spaceAfter' => 300);
$cell300Style = array('align' => 'center', 'lineHeigth' => 300, 'spaceBefore' => 200, 'spaceAfter' => 100);
$cellRowSpan = array('vMerge' => 'restart', 'valign' => 'center');
$cellRowContinue = array('vMerge' => 'continue');
$cellRowSpan2 = array('gridSpan' => 2, 'vMerge' => 'restart', 'valign' => 'center');
$cellRowContinue2 = array('gridSpan' => 2, 'vMerge' => 'continue');
$cellColSpan = array('gridSpan' => 2, 'valign' => 'center');
$cellColSpan3 = array('gridSpan' => 3, 'valign' => 'center');
$cellColSpan4 = array('gridSpan' => 4, 'valign' => 'center');
$cellColSpan5 = array('gridSpan' => 5, 'valign' => 'center');
$cellColSpan6 = array('gridSpan' => 6, 'valign' => 'center');

//First row
$table->addRow();
$table->addCell(800, $cellRowSpan)->addText('姓名', $fontStyle, $cell400Style);
$table->addCell(1300)->addText($user['NAME'], $fontStyle, $cell400Style);
$table->addCell(1000)->addText('现名', $fontStyle, $cell400Style);
$table->addCell(1300)->addText($user['NAME'], $fontStyle, $cell400Style);
$table->addCell(1100)->addText('性别', $fontStyle, $cell400Style);
$table->addCell(1200)->addText($user['GENDER'] == 1 ? '男' : '女', $fontStyle, $cell400Style);
$table->addCell(1000)->addText('民族', $fontStyle, $cell400Style);
$table->addCell(1200)->addText($user['POLITICS'], $fontStyle, $cell400Style);
if ($user['PHOTO']) {
    $table->addCell(1500, $cellRowSpan)->addImage(__DIR__ . DIRECTORY_SEPARATOR . 'photo.jpg', array(
        'width' => 70,
        'height' => 80
    ));
} else {
    $table->addCell(1500, $cellRowSpan)->addText('照片', $fontStyle, array('align' => 'center', 'lineHeigth' => 1800));
}
//
//$table->addRow();
//$table->addCell(null, $cellRowContinue);
//$table->addCell(1000)->addText('曾用名', $fontStyle, $cell400Style);
//$table->addCell(1300)->addText($user['NAME'], $fontStyle, $cell400Style);
//$table->addCell(1100)->addText('健康状况', $fontStyle, $cell400Style);
//$table->addCell(1200)->addText($user['SCHOOL'], $fontStyle, $cell400Style);
//$table->addCell(1000)->addText('邮政编码', $fontStyle, $cell400Style);
//$table->addCell(1200)->addText($user['postal'], $fontStyle, $cell400Style);
//$table->addCell(null, $cellRowContinue);
//
////Third row
//$table->addRow();
//$table->addCell(1800, $cellColSpan)->addText('出生日期', $fontStyle, $cell400Style);
//$table->addCell(1300)->addText($user['birthday'], $fontStyle, $cell400Style);
//$table->addCell(1100)->addText('身份证号', $fontStyle, $cell400Style);
//$table->addCell(3400, $cellColSpan3)->addText($user['personal_id'], $fontStyle, $cell400Style);
//$table->addCell(null, $cellRowContinue);
//
//// Fourth row
//$table->addRow();
//$table->addCell(800, $cellRowSpan)->addText('籍贯', $fontStyle, $cell400Style);
//$table->addCell(1000)->addText('原籍', $fontStyle, $cell400Style);
//$table->addCell(1300)->addText($user['original_birth_place'], $fontStyle, $cell400Style);
//$table->addCell(1100)->addText('入团时间', $fontStyle, $cell400Style);
//$table->addCell(1200)->addText($user['time_join_league'], $fontStyle, $cell400Style);
//$table->addCell(1000)->addText('学历', $fontStyle, $cell400Style);
//$table->addCell(1200)->addText($user['education'], $fontStyle, $cell400Style);
//$table->addCell(null, $cellRowContinue);
//
////Fifth row
//$table->addRow();
//$table->addCell(null, $cellRowContinue);
//$table->addCell(1000)->addText('出生地', $fontStyle, $cell400Style);
//$table->addCell(1300)->addText($user['brith_place'], $fontStyle, $cell400Style);
//$table->addCell(1100)->addText('入党时间', $fontStyle, $cell400Style);
//$table->addCell(1200)->addText($user['time_join_party'], $fontStyle, $cell400Style);
//$table->addCell(1000)->addText('身高', $fontStyle, $cell400Style);
//$table->addCell(1200, $cellColSpan)->addText($user['height'], $fontStyle, $cell400Style);
//
////Sixth row
//$table->addRow();
//$table->addCell(1800, $cellColSpan)->addText('家庭通讯地址', $fontStyle, $cell400Style);
//$table->addCell(2400, $cellColSpan4)->addText($user['home_address'], $fontStyle, $cell400Style);
//$table->addCell(1200)->addText('体重', $fontStyle, $cell400Style);
//$table->addCell(1500)->addText($user['weight'], $fontStyle, $cell400Style);
//
//$money=0;
//if(count($family)>0){
//    foreach($family as $fa){
//        if($fa['relation']=='母亲'){
//            $user['moth_race']=$fa['POLITICS'];
//        }
//        if($fa['relation']=='父亲'){
//            $user['fath_race']=$fa['POLITICS'];
//        }
//        $money=(int)$fa['money']+$money;
//    }
//}else{
//    $user['moth_race']='';
//    $user['fath_race']='';
//}
//
////KNSSQB
////Seventh row
//$table->addRow();
//$table->addCell(1800, $cellRowSpan2)->addText('联系方式', $fontStyle, $cell400Style);
//$table->addCell(1000)->addText('父亲', $fontStyle, $cell400Style);
//$table->addCell(1400)->addText($user['fath_race'], $fontStyle, $cell400Style);
//$table->addCell(1200)->addText('母亲', $fontStyle, $cell400Style);
//$table->addCell(1000)->addText($user['moth_race'], $fontStyle, $cell400Style);
//$table->addCell(1200)->addText('邮箱', $fontStyle, $cell400Style);
//$table->addCell(1500)->addText($user['email'], $fontStyle, $cell400Style);
//
////Eighty row
//$table->addRow();
//$table->addCell(1800, $cellRowContinue2);
//$table->addCell(1000)->addText('本人', $fontStyle, $cell400Style);
//$table->addCell(1400)->addText($user['phone'], $fontStyle, $cell400Style);
//$table->addCell(1200)->addText('宅电', $fontStyle, $cell400Style);
//$table->addCell(1000)->addText($user['home_phone'], $fontStyle, $cell400Style);
//$table->addCell(1200)->addText('QQ/微信号', $fontStyle, $cell400Style);
//$table->addCell(1500)->addText($user['qq'], $fontStyle, $cell400Style);
//
////New Table
//$table6 = $section->addTable('myOwnTableStyle');
//$table6->addRow();
//$table6->addCell(9100, $cellColSpan5)->addText('入 学 前 经 历', $fontStyleBold, $cell400Style);
//
////First row
//$table6->addRow();
//$table6->addCell(1500)->addText('自何年何月', $fontStyle, $cell400Style);
//$table6->addCell(1500)->addText('至何年何月', $fontStyle, $cell400Style);
//$cell = $table6->addCell(3100);
//$cell->addText('在何地何校上学', $fontStyle, $cell400Style);
//$cell->addText('或在何地何单位从事何种工作', $fontStyle, $cell400Style);
//$table6->addCell(1500)->addText('任何职务', $fontStyle, $cell400Style);
//$table6->addCell(1500)->addText('证明人', $fontStyle, $cell400Style);
//
//foreach($pre_schools as $school) {
//    $table6->addRow();
//    $table6->addCell(1500)->addText($school['date_from'], $fontStyle, $cell400Style);
//    $table6->addCell(1500)->addText($school['date_to'], $fontStyle, $cell400Style);
//    $table6->addCell(3100)->addText($school['school'], $fontStyle, $cell400Style);
//    $table6->addCell(1500)->addText($school['duty'], $fontStyle, $cell400Style);
//    $table6->addCell(1500)->addText($school['prove'], $fontStyle, $cell400Style);
//}
//
////New Table
//$table3 = $section->addTable('myOwnTableStyle');
//$table3->addRow();
//$table3->addCell(9100, $cellColSpan6)->addText('家　庭　成　员　情　况', $fontStyleBold, $cell400Style);
//
////First row
//$table3->addRow();
//$cell = $table3->addCell(1200);
//$cell->addText('家庭人', $fontStyle, $cell400Style);
//$cell->addText('口数量', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText(count($family)>1 ? count($family) : '', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText('户口性质', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText($user['nature'], $fontStyle, $cell400Style);
//$cell = $table3->addCell(3100);
//$cell->addText('目前家庭经济状况', $fontStyle, $cell400Style);
//$cell->addText('（人均月收入）', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText(count($family)>1 ? $money/count($family)/12 :'',$fontStyle, $cell400Style);
//
////Second row
//$table3->addRow();
//$cell = $table3->addCell(1200);
//$cell->addText('与本人关系', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText('姓名', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText('年龄', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText('联系电话', $fontStyle, $cell400Style);
//$table3->addCell(3100)->addText('现在何地何单位从事何种工作', $fontStyle, $cell400Style);
//$table3->addCell(1200)->addText('任何职务', $fontStyle, $cell400Style);
//
//foreach($family as $member) {
//    $table3->addRow();
//    $table3->addCell(1200)->addText($member['relation'], $fontStyle, $cell400Style);
//    $table3->addCell(1200)->addText($member['name'], $fontStyle, $cell400Style);
//    $table3->addCell(1200)->addText($member['birthday'], $fontStyle, $cell400Style);
//    $table3->addCell(1200)->addText($member['POLITICS'], $fontStyle, $cell400Style);
//    $table3->addCell(3100)->addText($member['work'], $fontStyle, $cell400Style);
//    $table3->addCell(1200)->addText($member['duty'], $fontStyle, $cell400Style);
//}
//
////New Table
//$table2 = $section->addTable('myOwnTableStyle');
//$table2->addRow();
//$table2->addCell(9100, $cellColSpan6)->addText('主　要　社　会　关　系　情　况', $fontStyleBold, $cell400Style);
//
////First row
//$table2->addRow();
//$table2->addCell(1200)->addText('与本人关系', $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText('姓　名', $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText('单位电话', $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText('手机号码', $fontStyle, $cell400Style);
//$table2->addCell(3100)->addText('现在何地何单位从事何种工作', $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText('任何职务', $fontStyle, $cell400Style);
//
////Second row
//$table2->addRow();
//$table2->addCell(1200)->addText($user['social1_relation'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social1_name'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social1_birthday'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social1_role'], $fontStyle, $cell400Style);
//$table2->addCell(3100)->addText($user['social1_work'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social1_duty'], $fontStyle, $cell400Style);
//
////third row
//$table2->addRow();
//$table2->addCell(1200)->addText($user['social2_relation'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social2_name'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social2_birthday'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social2_role'], $fontStyle, $cell400Style);
//$table2->addCell(3100)->addText($user['social2_work'], $fontStyle, $cell400Style);
//$table2->addCell(1200)->addText($user['social2_duty'], $fontStyle, $cell400Style);
//
////New Table
//$table4 = $section->addTable('myOwnTableStyle');
//$table4->addRow();
//$table4->addCell(9100, $cellColSpan3)->addText('奖励情况', $fontStyleBold, $cell400Style);
//
//$table4->addRow();
//$table4->addCell(2500)->addText('学年', $fontStyle, $cell400Style);
//$table4->addCell(2500)->addText('学期', $fontStyle, $cell400Style);
//$table4->addCell(4100)->addText('内容', $fontStyle, $cell400Style);
//
//foreach($awards as $award) {
//    $table4->addRow();
//    $table4->addCell(2500)->addText($award['year'], $fontStyle, $cell400Style);
//    $table4->addCell(2500)->addText($award['semester'], $fontStyle, $cell400Style);
//    $table4->addCell(4100)->addText($award['award'], $fontStyle, $cell400Style);
//}
//
////New Table
//$table5 = $section->addTable('myOwnTableStyle');
//$table5->addRow();
//$table5->addCell(9100, $cellColSpan4)->addText('违纪情况', $fontStyleBold, $cell400Style);
//
//$table5->addRow();
//$table5->addCell(2000)->addText('学年', $fontStyle, $cell400Style);
//$table5->addCell(2000)->addText('学期', $fontStyle, $cell400Style);
//$table5->addCell(3100)->addText('内容', $fontStyle, $cell400Style);
//$table5->addCell(2000)->addText('类别', $fontStyle, $cell400Style);
//
//foreach($punishes as $punish) {
//    $table5->addRow();
//    $table5->addCell(2000)->addText($punish['year'], $fontStyle, $cell400Style);
//    $table5->addCell(2000)->addText($punish['semester'], $fontStyle, $cell400Style);
//    $table5->addCell(3100)->addText($punish['punish'], $fontStyle, $cell400Style);
//    $table5->addCell(2000)->addText($punish['category'], $fontStyle, $cell400Style);
//}
//
//$section->addTextBreak(1);

$section->addText('制表时间：' . date('Y') . '年' . date('m') . '月' . date('d') . '日', $lefontStyle,array('align' => 'right', 'lineHeigth' => 50, 'spaceBefore' => 50, 'spaceAfter' => 50));

//Output
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$fileName = $userCode . '.docx';
$objWriter->save($fileName);

$contentType = 'Content-type: application/vnd.openxmlformats-officedocument.wordprocessingml.document';
header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
header ( "Cache-Control: no-cache, must-revalidate" );
header ( "Pragma: no-cache" );
header ( $contentType );
header ( "Content-Disposition: attachment; filename=".$fileName);
readfile($fileName);
unlink($fileName);