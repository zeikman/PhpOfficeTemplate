<?php

use PhpOffice\PhpWord\TemplateProcessor;

class PhpWordTemplateProcessor extends TemplateProcessor
{
  /**
   * Content of document rels (in XML format) of the temporary document.
   *
   * @var string
   */
  private $temporaryDocumentRels;

  /**
   * @param string $documentTemplate The fully qualified template filename
   */
  public function __construct($documentTemplate)
  {
    parent::__construct($documentTemplate);

    $this->temporaryDocumentRels = $this->zipClass->getFromName($this->getRelationsName($this->getMainPartName()));
  }

  /**
   * Replace image by name
   *
   * @param string $search
   * @param mixed $replace
   */
  public function replaceImage($search, $replace): void
  {
    $fileNameStart = strpos($this->temporaryDocumentRels, $search);
    $fileName = strstr(substr($this->temporaryDocumentRels, $fileNameStart), '"', true);

    if ($fileNameStart !== false)
      $this->zipClass->addFromString("word/media/" . $fileName, $replace);
  }
}

?>