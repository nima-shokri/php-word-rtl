<?php

use \phpWordRtl\phpWordRtl;
require __DIR__."/../src/phpWordRtl/phpWordRtl.php";

$ds = DIRECTORY_SEPARATOR;

$templatePath =  __DIR__."/template/bookmark.docx";
$word = new phpWordRtl($templatePath);
$word->setBookmarkValue('name', 'آقای فلانی');
$word->setBookmarkValue('organization', 'سازمان فلان');
$word->setBookmarkValue('date', '02/04/1399');
$word->setBookmarkValue('number', '35/858585');
$word->setBookmarkValue('attach', 'ندارد');
$word->setBookmarkValue('description', 'شرح کار');
$word->setBookmarkValue('item1', 'کالای مورد نیاز اول');
$word->setBookmarkValue('item2', 'کالای مورد نیاز دوم');
$word->setBookmarkValue('item3', 'کالای مورد نیاز سوم');
$word->setBookmarkValue('writer', 'نام شخص');
$word->setBookmarkValue('department', 'نام رئیس');

$word->output('aoutput.docx', true);

