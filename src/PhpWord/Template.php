<?php

    /**
     * This file is part of PHPWord - A pure PHP library for reading and writing
     * word processing documents.
     *
     * PHPWord is free software distributed under the terms of the GNU Lesser
     * General Public License version 3 as published by the Free Software Foundation.
     *
     * For the full copyright and license information, please read the LICENSE
     * file that was distributed with this source code. For the full list of
     * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
     *
     * @link        https://github.com/PHPOffice/PHPWord
     * @copyright   2010-2016 PHPWord contributors
     * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
     */

    namespace PhpOffice\PhpWord;


    class Template extends TemplateProcessor
    {

        /**
         * ZipArchive
         *
         * @var ZipArchive
         */
        private $_objZip;

        /**
         * Temporary Filename
         *
         * @var string
         */
        private $_tempFileName;

        /**
         * Document XML
         *
         * @var string
         */
        private $_documentXML;
        private $_header1XML;



        private $_footer1XML;
        private $_footer2XML;
        private $_footer3XML;
        private $_rels;
        private $_types;
        private $_countRels;

        /**
         * Create a new Template Object
         *
         * @param string $strFilename
         */
        public function __construct($strFilename)
        {
            $path = dirname($strFilename);

            $this->_tempFileName = $path . DIRECTORY_SEPARATOR . time() . '.docx';

            //var_dump($path);exit;
            copy($strFilename, $this->_tempFileName); // Copy the source File to the temp File

            $this->_objZip = new \ZipArchive();
            $this->_objZip->open($this->_tempFileName);

            $this->_documentXML = $this->_objZip->getFromName('word/document.xml');
            $this->_header1XML = $this->_objZip->getFromName('word/header1.xml');
            $this->_footer1XML = $this->_objZip->getFromName('word/footer1.xml');
            $this->_footer2XML = $this->_objZip->getFromName('word/footer2.xml');
            $this->_footer3XML = $this->_objZip->getFromName('word/footer3.xml');
            $this->_rels = $this->_objZip->getFromName('word/_rels/document.xml.rels'); #erap 07/07/2015
            $this->_types = $this->_objZip->getFromName('[Content_Types].xml'); #erap 07/07/2015
            $this->_countRels = substr_count($this->_rels, 'Relationship') - 1; #erap 07/07/2015
        }

        /**
         * Set a Template value
         *
         * @param mixed $search
         * @param mixed $replace
         */
        public function setValue($search, $replace, $limit = -1)
        {
            /* if($search === 'rowAvgL1')
                 {
                     var_dump('1');
                     exit;}*/
            $replace = preg_replace('~\R~u', '</w:t><w:br/><w:t>', $replace);

            if (substr($search, 0, 1) !== '{' && substr($search, -1) !== '}')
            {
                $search = '{' . $search . '}';
            }

            preg_match_all('/\{[^}]+\}/', $this->_documentXML, $matches);

            foreach ($matches[0] as $k => $match)
            {
                $no_tag = strip_tags($match);
                if ($no_tag == $search)
                {
                    $match = '{' . $match . '}';
                    $this->_documentXML = preg_replace($match, $replace, $this->_documentXML, $limit);
                    $this->_header1XML = preg_replace($match, $replace, $this->_header1XML);
                    if ($limit == 1)
                    {
                        break;
                    }
                }
            }

            preg_match_all('/\{[^}]+\}/', $this->_header1XML, $matches);

            foreach ($matches[0] as $k => $match)
            {
                $no_tag = strip_tags($match);
                if ($no_tag == $search)
                {
                    $match = '{' . $match . '}';
                    $this->_header1XML = preg_replace($match, $replace, $this->_header1XML);
                    if ($limit == 1)
                    {
                        break;
                    }
                }
            }




            preg_match_all('/\{[^}]+\}/', $this->_footer1XML, $matches);
            foreach ($matches[0] as $k => $match)
            {
                $no_tag = strip_tags($match);
                if ($no_tag == $search)
                {
                    $match = '{' . $match . '}';
                    $this->_footer1XML = preg_replace($match, $replace, $this->_footer1XML);
                    if ($limit == 1)
                    {
                        break;
                    }
                }
            }

            preg_match_all('/\{[^}]+\}/', $this->_footer2XML, $matches);
            foreach ($matches[0] as $k => $match)
            {
                $no_tag = strip_tags($match);
                if ($no_tag == $search)
                {
                    $match = '{' . $match . '}';
                    $this->_footer2XML = preg_replace($match, $replace, $this->_footer2XML);
                    if ($limit == 1)
                    {
                        break;
                    }
                }
            }

            preg_match_all('/\{[^}]+\}/', $this->_footer3XML, $matches);
            foreach ($matches[0] as $k => $match)
            {
                $no_tag = strip_tags($match);
                if ($no_tag == $search)
                {
                    $match = '{' . $match . '}';
                    $this->_footer3XML = preg_replace($match, $replace, $this->_footer3XML);
                    if ($limit == 1)
                    {
                        break;
                    }
                }
            }



        }

        /**
         * @param string $search
         * @param int    $numberOfClones
         */
        public function cloneRow($search, $numberOfClones)
        {

            if ('{' !== substr($search, 0, 1) && '}' !== substr($search, -1))
            {
                $search = '{' . $search . '}';
            }
            //var_dump($search);
            $tagPos = strpos($this->_documentXML, $search);
            //var_dump($tagPos);
            // exit;
            if (!$tagPos)
            {

                throw new \Exception("Can not clone row, template variable not found or variable contains markup. ( " . $search . " )");
            }

            $rowStart = $this->findRowStart($tagPos);
            $rowEnd = $this->findRowEnd($tagPos);
            $xmlRow = $this->getSlice($rowStart, $rowEnd);

            // Check if there's a cell spanning multiple rows.
            if (preg_match('#<w:vMerge w:val="restart"/>#', $xmlRow))
            {
                // $extraRowStart = $rowEnd;
                $extraRowEnd = $rowEnd;
                while (true)
                {
                    $extraRowStart = $this->findRowStart($extraRowEnd + 1);
                    $extraRowEnd = $this->findRowEnd($extraRowEnd + 1);

                    // If extraRowEnd is lower then 7, there was no next row found.
                    if ($extraRowEnd < 7)
                    {
                        break;
                    }

                    // If tmpXmlRow doesn't contain continue, this row is no longer part of the spanned row.
                    $tmpXmlRow = $this->getSlice($extraRowStart, $extraRowEnd);
                    if (!preg_match('#<w:vMerge/>#', $tmpXmlRow) && !preg_match('#<w:vMerge w:val="continue" />#', $tmpXmlRow))
                    {
                        break;
                    }
                    // This row was a spanned row, update $rowEnd and search for the next row.
                    $rowEnd = $extraRowEnd;
                }
                $xmlRow = $this->getSlice($rowStart, $rowEnd);
            }

            $result = $this->getSlice(0, $rowStart);
            for ($i = 1; $i <= $numberOfClones; $i++)
            {
                $result .= preg_replace('/\{(.*?)\}/', '{\\1#' . $i . '}', $xmlRow);

            }
            //var_dump($result);exit;
            $result .= $this->getSlice($rowEnd);

            $this->_documentXML = $result;
        }


        /**
         * Find the start position of the nearest table row before $offset.
         *
         * @param integer $offset
         *
         * @return integer
         *
         * @throws \PhpOffice\PhpWord\Exception\Exception
         */
        protected function findRowStart($offset)
        {
            $rowStart = strrpos($this->_documentXML, '<w:tr ', ((strlen($this->_documentXML) - $offset) * -1));

            if (!$rowStart)
            {
                $rowStart = strrpos($this->_documentXML, '<w:tr>', ((strlen($this->_documentXML) - $offset) * -1));
            }
            if (!$rowStart)
            {
                throw new Exception('Can not find the start position of the row to clone.');
            }

            return $rowStart;
        }

        /**
         * Find the end position of the nearest table row after $offset.
         *
         * @param integer $offset
         *
         * @return integer
         */
        protected function findRowEnd($offset)
        {
            return strpos($this->_documentXML, '</w:tr>', $offset) + 7;
        }

        /**
         * Get a slice of a string.
         *
         * @param integer $startPosition
         * @param integer $endPosition
         *
         * @return string
         */
        protected function getSlice($startPosition, $endPosition = 0)
        {
            if (!$endPosition)
            {
                $endPosition = strlen($this->_documentXML);
            }

            return substr($this->_documentXML, $startPosition, ($endPosition - $startPosition));
        }

        /**
         * Save Template
         *
         * @param string $strFilename
         */
        public function saveAs($fileName)
        {
            if (file_exists($fileName))
            {
                unlink($fileName);
            }

            $this->_objZip->addFromString('word/document.xml', $this->_documentXML);
            $this->_objZip->addFromString('word/header1.xml', $this->_header1XML);

            
            $this->_objZip->addFromString('word/footer1.xml', $this->_footer1XML); 
            $this->_objZip->addFromString('word/footer2.xml', $this->_footer2XML); 
            $this->_objZip->addFromString('word/footer3.xml', $this->_footer3XML); 

            $this->_objZip->addFromString('word/_rels/document.xml.rels', $this->_rels); #erap 07/07/2015
            $this->_objZip->addFromString('[Content_Types].xml', $this->_types); #erap 07/07/2015
            // Close zip file

            $this->_objZip->close();


            rename($this->_tempFileName, $fileName);
        }

        public function replaceImage($path, $imageName)
        {
            $this->_objZip->deleteName('word/media/' . $imageName);
            $this->_objZip->addFile($path, 'word/media/' . $imageName);
        }

        public function replaceStrToImg($strKey, $arrImgPath)
        {
            //289x108
            $search = $strKey;
            $strKey = '{' . $strKey . '}';

            $relationTmpl = '<Relationship Id="RID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/IMG"/>';
            $imgTmpl = '<w:pict><v:shape type="#_x0000_t75" style="width:WIDpx;height:HEIpx"><v:imagedata r:id="RID" o:title=""/></v:shape></w:pict>';
            $typeTmpl = ' <Override PartName="/word/media/IMG" ContentType="image/EXT"/>';
            $toAdd = $toAddImg = $toAddType = '';
            $aSearch = array('RID', 'IMG');
            $aSearchType = array('IMG', 'EXT');

            $imgSrc = explode('.', $arrImgPath['src']);

            $imgExt = array_pop($imgSrc);
            if (in_array($imgExt, array('png', 'PNG')))
            {
                $imgExt = 'png';
            }
            $imgName = 'img' . $this->_countRels . '.' . $imgExt;
            $rid = 'rId' . $this->_countRels++;

            $this->_objZip->addFile($arrImgPath['src'], 'word/media/' . $imgName);

            if (isset($arrImgPath['size']))
            {
                $w = $arrImgPath['size']['width'];
                $h = $arrImgPath['size']['height'];
            }
            else
            {
                $w = 289;
                $h = 108;
                // $w=150;
                // $h=35;
            }

            $toAddImg .= str_replace(array('RID', 'WID', 'HEI'), array($rid, $w, $h), $imgTmpl);
            if (isset($img['dataImg']))
            {
                $toAddImg .= '<w:br/><w:t>' . $this->limpiarString($img['dataImg']) . '</w:t><w:br/>';
            }

            $aReplace = array($imgName, $imgExt);
            $toAddType .= str_replace($aSearchType, $aReplace, $typeTmpl);

            $aReplace = array($rid, $imgName);
            $toAdd .= str_replace($aSearch, $aReplace, $relationTmpl);


            if (substr($search, 0, 1) !== '{' && substr($search, -1) !== '}')
            {
                $search = '{' . $search . '}';
            }

            preg_match_all('/\{[^}]+\}/', $this->_documentXML, $matches);
            foreach ($matches[0] as $k => $match)
            {
                $no_tag = strip_tags($match);
                if ($no_tag == $search)
                {
                    $match = '{' . $match . '}';
                    $this->_documentXML = preg_replace($match, $toAddImg, $this->_documentXML);
                }
            }


            //  $this->_documentXML = str_replace('<w:t>' . $strKey . '</w:t>', $toAddImg, $this->_documentXML);
            $this->_types = str_replace('</Types>', $toAddType, $this->_types) . '</Types>';
            $this->_rels = str_replace('</Relationships>', $toAdd, $this->_rels) . '</Relationships>';
        }


    }
