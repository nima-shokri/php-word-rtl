<?php

use \phpWordRtl\phpWordRtl;
require __DIR__."/../src/phpWordRtl/phpWordRtl.php";

$ds = DIRECTORY_SEPARATOR;

$templatePath =  __DIR__."/template/block.docx";
$word = new phpWordRtl($templatePath);
$word->deleteBlock('بلاک1');
$word->deleteBlock('بلاک2');

$word->output('aoutput.docx', true);

