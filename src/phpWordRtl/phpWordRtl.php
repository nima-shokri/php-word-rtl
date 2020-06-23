<?php

namespace phpWordRtl;

require __DIR__ . "/clsTbsZip.php";


/**
 *   php.word.Rtl version 1.0
 *   Date    : 2015-08-15
 *   Author  : Shokri (email: paybox.ir@gmail.com)
 *   Licence : MIT
 *   A pure PHP library for reading and writing RTL in Microsoft Word 2007+ documents.
 */


use phpWordRtl\clsTbsZip;

define('VAR_START', '((');   // the start of variables 
define('VAR_END', '))');   // the start of variables 

define('BLOCK_START_BEGIN', '(((');   // 
define('BLOCK_START_END', ')');   // 
define('BLOCK_END_BEGIN', '(');   // 
define('BLOCK_END_END', ')))');   // 

class phpWordRtl
{

    private $variables = array();
    private $delBlocks = array();
    private $bookmarks = array();
    private $headerBookmarks = array();
    private $footerBookmarks = array();

    private $images = array();
    public $errors = array();
    public $template;

    function __construct($template)
    {
        $this->template = $template;
    }

    public function setVarValue($var, $value)
    {
        $var = VAR_START . $var . VAR_END;
        $this->variables[count($this->variables)] = array('var' => $var, 'value' => $value);
    }

    // update bookmark values
    public function setBookmarkValue($bmName, $value)
    {
        $this->bookmarks[count($this->bookmarks)] = array(
            'name' => $bmName,
            'value' => $value,
        );
    }

    // update header bookmark values
    public function setHeaderBookmarkValue($bmName, $value)
    {
        $this->headerBookmarks[count($this->headerBookmarks)] = array(
            'name' => $bmName,
            'value' => $value,
        );
    }

    // update footer bookmark values
    public function setFooterBookmarkValue($bmName, $value)
    {
        $this->footerBookmarks[count($this->footerBookmarks)] = array(
            'name' => $bmName,
            'value' => $value,
        );
    }

    // set image to image array
    public function addImageToVar($var, $path, $params = null)
    {
        $this->images[count($this->images)] = array(
            'var' => $var,
            'path' => $path,
            'params' => $params,
        );
    }

    // add block to delete list
    public function deleteBlock($blockName)
    {
        $this->delBlocks[count($this->delBlocks)] = array(
            'start' => BLOCK_START_BEGIN . $blockName . BLOCK_START_END,
            'end' => BLOCK_END_BEGIN . $blockName . BLOCK_END_END
        );
    }

    public function output($fileName, $phpOutput = true)
    {

        $zip = new clsTbsZip(); // create a new instance of the TbsZip class
        $zip->Open($this->template); // open an existing archive for reading and/or modifying
        // --------------------------------------------------
        // Reading information and data in the opened archive
        // --------------------------------------------------
        $xml = $zip->FileRead('word/document.xml');

        $relations = $zip->FileRead('word/_rels/document.xml.rels');
        $contentTypes = $zip->FileRead('[Content_Types].xml');

        $fileExist = true;
        $c = 1;
        $headerXml = array();
        $footerXml = array();
        while ($fileExist) {
            $tmpFFE = true;
            $tmpHFE = true;
            if ($zip->FileExists("word/header$c.xml"))
                $headerXml[count($headerXml)] = $zip->FileRead("word/header$c.xml");
            else
                $tmpHFE = false;
            if ($zip->FileExists("word/footer$c.xml"))
                $footerXml[count($footerXml)] = $zip->FileRead("word/footer$c.xml");
            else
                $tmpFFE = false;
            $c++;

            if (!($tmpFFE && $tmpHFE))
                $fileExist = false;
        }

        // SET VARIABLES TO FILE 
        for ($i = 0; $i < count($this->variables); $i++) {
            $xml = str_replace($this->variables[$i]['var'], $this->variables[$i]['value'], $xml);
        }
        // DELETE BLOCKS  
        for ($i = 0; $i < count($this->delBlocks); $i++) {
            $xml = $this->deleteBlockFromXml($xml, $this->delBlocks[$i]['start'], $this->delBlocks[$i]['end']);
        }
        // UPDATE BOOKMARK
        for ($i = 0; $i < count($this->bookmarks); $i++) {
            $xml = $this->changeBookmarkValue($xml, $this->bookmarks[$i]['name'], $this->bookmarks[$i]['value']);
        }

        // UPDATE HEADERS BOOKMARK && VAR VALUES
        for ($j = 0; $j < count($headerXml); $j++) {
            // set bookmarks value
            for ($i = 0; $i < count($this->headerBookmarks); $i++) {
                $headerXml[$j] = $this->changeBookmarkValue($headerXml[$j], $this->headerBookmarks[$i]['name'], $this->headerBookmarks[$i]['value']);
            }

            // set variables value
            for ($i = 0; $i < count($this->variables); $i++) {
                $headerXml[$j] = str_replace($this->variables[$i]['var'], $this->variables[$i]['value'], $headerXml[$j]);
            }
        }
        // UPDATE FOOTERS BOOKMARK&& VAR VALUES
        for ($j = 0; $j < count($footerXml); $j++) {
            // set bookmarks value
            for ($i = 0; $i < count($this->footerBookmarks); $i++) {
                $footerXl[$j] = $this->changeBookmarkValue($footerXml[$j], $this->footerBookmarks[$i]['name'], $this->footerBookmarks[$i]['value']);
            }
            // set variables value
            for ($i = 0; $i < count($this->variables); $i++) {
                $footerXml[$j] = str_replace($this->variables[$i]['var'], $this->variables[$i]['value'], $footerXml[$j]);
            }
        }

        // add images to file
        for ($i = 0; $i < count($this->images); $i++) {

            // add image to zip
            $fPath = $this->images[$i]['path'];

            if (file_exists($fPath)) {
                // generate file name
                $dotPos = strrpos($fPath, '.', 0);
                $fileType = substr($fPath, $dotPos, strlen($fPath) - $dotPos);

                $contentTypes = $this->addContentTypes($contentTypes, $fileType);

                $imageName = 'RtlImage' . ($i + 1) . $fileType;
                // add file to zip archive
                $zip->FileAdd('word/media/' . $imageName, $fPath, TBSZIP_FILE, false);


                $idCount = substr_count($relations, ' Id=');
                $rid = 'rId' . ($idCount + 1);
                $relations = $this->addImagesToRelationXml($relations, $fPath, $imageName,  $rid);

                $xml = $this->addImageTagsToXml($xml, $this->images[$i]['var'], $imageName, $rid, $this->images[$i]['params']);
            } else {
                $this->errors[count($this->errors)] = "File \"$fPath\" not found!";
            }
        }




        // CLEAR UNUSED BLOCKS 
        $xml = $this->clearTemplate($xml);



        $zip->FileReplace('word/_rels/document.xml.rels', $relations, TBSZIP_STRING); // replace the file by giving the content
        $zip->FileReplace('[Content_Types].xml', $contentTypes, TBSZIP_STRING); // replace the file by giving the content
        $zip->FileReplace('word/document.xml', $xml, TBSZIP_STRING); // replace the file by giving the content
        for ($i = 1; $i <= count($headerXml); $i++)
            $zip->FileReplace("word/header$i.xml", $headerXml[$i - 1], TBSZIP_STRING); // replace the file by giving the content
        for ($i = 1; $i <= count($footerXml); $i++)
            $zip->FileReplace("word/footer$i.xml", $footerXml[$i - 1], TBSZIP_STRING); // replace the file by giving the content

        if (!$phpOutput) {
            $zip->Flush(TBSZIP_FILE, $fileName); // apply modifications as a new local file
        } else {
            header("Content-type: application/force-download");
            header("Content-Disposition: attachment; filename=$fileName");
            $zip->Flush(TBSZIP_DOWNLOAD + TBSZIP_NOHEADER);
        }


        $zip->Close(); // stop to work with the opened archive. Modifications are not applied to the opened archive, use Flush() to commit  
    }

    private function deleteBlockFromXml($xml, $blockStart, $blockEnd)
    {
        $bStart = strpos($xml, $blockStart, 0);
        $bEnd = strpos($xml, $blockEnd, 0);

        $below = substr($xml, 0, $bStart);
        $abov = substr($xml, $bEnd, strlen($xml) - $bEnd);
        $bb = strrpos($below, '<w:p ');
        $be = strpos($abov, '</w:p>', 0) + 6;

        $below = substr($below, 0, $bb);
        $abov = substr($abov, $be, strlen($abov) - $be);

        $block = substr($xml, $bb, $be - $bb);


        $xml = $below . $abov;

        return $xml;
    }

    private function clearTemplate($xml)
    {
        $tmp = $this->removeBlockStartAndEnd($xml, BLOCK_START_BEGIN);
        $xml = $this->removeBlockStartAndEnd($tmp, BLOCK_END_END);
        return $xml;
    }

    private function removeBlockStartAndEnd($xml, $keyWord, $c = 0)
    {

        $pos = strpos($xml, $keyWord, 0);

        if (!($pos > -1))
            return $xml;

        // remove unnecessary lines
        $below = substr($xml, 0, $pos);
        $abov = substr($xml, $pos, strlen($xml) - $pos);

        $bb = strrpos($below, '<w:p ');
        $be = strpos($abov, '</w:p>', 0) + 6;

        $below = substr($below, 0, $bb);
        $abov = substr($abov, $be, strlen($abov) - $be);

        $xml = $below . $abov;

        $pos = strpos($xml, $keyWord, 0);
        if ($pos > -1)
            $xml = $this->removeBlockStartAndEnd($xml, $keyWord,  ++$c);

        return $xml;
    }

    // update bookmark value in template
    private function changeBookmarkValue($xml, $bmName, $bmValue)
    {
        $bmName = 'w:name="' . $bmName . '"';

        $pos = strpos($xml, $bmName, 0);

        if ($pos < 1) return $xml;

        $abov = substr($xml, 0, $pos);
        $below = substr($xml, $pos, strlen($xml) - $pos);
        $tEnd = strpos($below, '</w:t>', 0);
        $below = substr($below, $tEnd, strlen($below) - $tEnd);


        $tEnd = strlen($abov) + $tEnd;
        $abov = substr($xml, 0, $tEnd);

        $tEnd = strrpos($abov, '>', 0) + 1;

        $abov = substr($xml, 0, $tEnd);

        $res = $abov . $bmValue . $below;
        return $res;
    }

    

    // add images to relation resource 
    function addImagesToRelationXml($rels, $path, $name, $newId)
    {

        $pos = strpos($rels, '</Relationships>');

        $abow = substr($rels, 0, $pos);
        $below = substr($rels, $pos, strlen($rels) - $pos);

        $newRel = '<Relationship Id="' . $newId . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' . $name . '"/>';
        $rels = $abow . $newRel . $below;
        return $rels;
    }

    // add image tags to xml 
    private function addImageTagsToXml($xml, $var, $name, $rId, $params = null)
    {

        $id = substr_count($xml, '<wp:docPr') + 1;

        $image = '<w:r>
                        <w:rPr>
                                <w:rFonts w:cs="Arial"/>
                                <w:noProof/>
                                <w:rtl/>
                                <w:lang w:bidi="ar-SA"/>
                        </w:rPr>
                        <w:drawing>
                                <wp:inline distT="0" distB="0" distL="0" distR="0">
                                <wp:extent cx="' . $params['width'] . '" cy="' . $params['height'] . '"/>
                                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                                <wp:docPr id="' . $id . '" name="Picture ' . $id . '" descr="' . $name . '"/>
                                <wp:cNvGraphicFramePr>
                                        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                                </wp:cNvGraphicFramePr>
                                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                        <pic:nvPicPr>
                                                                <pic:cNvPr id="0" name="Picture ' . $id . '" descr="' . $name . '"/>
                                                                <pic:cNvPicPr>
                                                                        <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                                                                </pic:cNvPicPr>
                                                        </pic:nvPicPr>
                                                        <pic:blipFill>
                                                                <a:blip r:embed="' . $rId . '"/>
                                                                        <a:srcRect/>
                                                                        <a:stretch>
                                                                                <a:fillRect/>
                                                                        </a:stretch>
                                                        </pic:blipFill>
                                                        <pic:spPr bwMode="auto">
                                                                <a:xfrm>
                                                                        <a:off x="0" y="0"/>
                                                                        <a:ext cx="' . $params['width'] . '" cy="' . $params['height'] . '" />
                                                                </a:xfrm>
                                                                <a:prstGeom prst="rect">
                                                                        <a:avLst/>
                                                                </a:prstGeom>
                                                                <a:noFill/>
                                                                <a:ln w="9525">
                                                                        <a:noFill/>
                                                                        <a:miter lim="800000"/>
                                                                        <a:headEnd/>
                                                                        <a:tailEnd/>
                                                                </a:ln>
                                                        </pic:spPr>
                                                </pic:pic>
                                        </a:graphicData>
                                </a:graphic>
                                </wp:inline>
                        </w:drawing>
                </w:r>';

        $var = VAR_START . $var . VAR_END;
        $pos = strpos($xml, $var);


        $abov = substr($xml, 0, $pos);
        $below = substr($xml, $pos, strlen($xml) - $pos);


        $sr = strrpos($abov, '<w:r>');
        $abov = substr($abov, 0, $sr);

        $er = strpos($below, '</w:r>') + 6;
        $below = substr($below, $er, strlen($below) - $er);



        $xml = $abov . $image . $below;

        return $xml;
    }

    // add content types to word file
    private function addContentTypes($contentType, $type)
    {
        $typeTag = '';
        switch ($type) {
            case '.png':
                $typeTag = '<Default Extension="png" ContentType="image/png"/>';
                break;
            case '.jpg':
                $typeTag = '<Default Extension="jpeg" ContentType="image/jpeg"/>';
                break;
            case '.gif':
                $typeTag = '<Default Extension="gif" ContentType="image/gif"/>';
                break;

            default:

                break;
        }

        $pos = strpos($contentType, $typeTag);
        if (!($pos > 0)) {
            $p = strpos($contentType, '</Types>');
            $abov = substr($contentType, 0, $p);
            $below = substr($contentType, $p, strlen($contentType) - $p);
            $contentType = $abov . $typeTag . $below;
        }
        return $contentType;
    }
}
