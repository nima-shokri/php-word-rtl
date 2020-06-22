<?php

use \phpWordRtl\phpWordRtl;
require __DIR__."/../src/phpWordRtl/phpWordRtl.php";

$ds = DIRECTORY_SEPARATOR;

$templatePath =  __DIR__."/template/variable.docx";
$word = new phpWordRtl($templatePath);
$word->setVarValue('نام', 'فلانی');
$word->setVarValue('سازمان', 'سازمان فلان');
$word->setVarValue('تاریخ', '02/04/1399');
$word->setVarValue('شماره', '35/858585');
$word->setVarValue('پیوست', 'ندارد');
$word->setVarValue('توضیح', 'شرح کار');
$word->setVarValue('مورد1', 'کالای مورد نیاز اول');
$word->setVarValue('مورد2', 'کالای مورد نیاز دوم');
$word->setVarValue('مورد3', 'کالای مورد نیاز سوم');
$word->setVarValue('نویسنده', 'نام شخص');
$word->setVarValue('قسمت', 'نام رئیس');

$word->output('aoutput.docx', true);

