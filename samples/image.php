<?php

use \phpWordRtl\phpWordRtl;
require __DIR__."/../src/phpWordRtl/phpWordRtl.php";

$templatePath =  __DIR__."/template/image.docx";
$imagePath =  __DIR__."/template/img/image.png";
$word = new phpWordRtl($templatePath);

$word->addImageToVar('تصویر',$imagePath,[
    'width'=>5195570,
    'height'=>2762250,
]);

$word->output('aoutput.docx', true);


