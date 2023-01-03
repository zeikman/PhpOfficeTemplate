<?php

require '../vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Settings as WordSettings;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;

/**
 * PhpWordTemplate
 */
class PhpWordTemplate
{
  public const DEFAULT_RENDERER = 'tcpdf';

  private $word_obj;
  private $target_dir;
  private $file_name;
  private $file_prefix;
  private $file_post;
  private $pdf_renderer;
  private $relative_file_path;
  private $empty_if_var_unfound;

  function __construct()
  {
    $args = func_get_arg(0);

    $this->target_dir   = $args['target_dir'];
    $this->file_name    = $args['file_name'];
    $this->file_prefix  = $args['file_prefix'];
    $this->file_post    = $args['file_post'];

    $this->pdf_renderer = $args['pdf_renderer']
      ? strtolower($args['pdf_renderer'])
      : self::DEFAULT_RENDERER;

    $this->empty_if_var_unfound = gettype($args['enable_empty_space']) == 'boolean'
      ? $args['enable_empty_space']
      : false;

    $this->relative_file_path = $this->target_dir . $this->file_name;

    $template_path = $this->file_post
      ? $this->file_post
      : $this->relative_file_path;

    $this->word_obj = new TemplateProcessor($template_path);
  }

  public function getPhpWord()
  {
    return $this->word_obj;
  }

  /**
   * Substitute variable in cell with value (Word)
   *
   * @param enableEmptyValueIfUnfound - enable empty space substitution if variable unfound
   */
  public function substituteCell($data) {
    if (count($data) > 0) {
      $this->word_obj->setValues($data);

      // replace with empty space [' '] if variable unfound in data pool
      if ($this->empty_if_var_unfound) {
        $unfoundVarList = [];

        foreach ($this->word_obj->getVariables() as $key => $v)
          $unfoundVarList[$v] = ' ';

        $this->word_obj->setValues($unfoundVarList);
      }

      // var_dump($this->word_obj->getVariables()); exit;
    }
  }

  /**
   * Change pdf renderer class
   *
   * @param renderer - PDF renderer class name
   */
  public function setPdfRenderer($renderer = null)
  {
    $this->pdf_renderer = $renderer
      ? strtolower($renderer)
      : self::DEFAULT_RENDERER; // default
  }

  private function _saveTempFile()
  {
    $pos = strrpos($this->file_name, "/");

    $temp_file_name = $pos > -1
      ? substr($this->file_name, $pos + 1)
      : $this->file_name;

    $temp_file_path = $this->target_dir . "/temp_$temp_file_name";

    // [$this->word_obj] is PhpWord template processor
    $this->word_obj->saveAs($temp_file_path);

    return $temp_file_path;
  }

  private function _setPdfRenderer()
  {
    // mpdf renderer need >> chmod 775 pathinfo/mpdf
    if ($this->pdf_renderer == 'tcpdf')
      WordSettings::setPdfRenderer(WordSettings::PDF_RENDERER_TCPDF, "../vendor/tecnickcom/tcpdf");

    if ($this->pdf_renderer == 'mpdf')
      WordSettings::setPdfRenderer(WordSettings::PDF_RENDERER_MPDF, "../vendor/mpdf/mpdf");

    if ($this->pdf_renderer == 'dompdf')
      WordSettings::setPdfRenderer(WordSettings::PDF_RENDERER_DOMPDF, "../vendor/dompdf/dompdf");
  }

  private function _createWriterPDF()
  {
    $temp_file_path = self::_saveTempFile();

    self::_setPdfRenderer();

    $phpWord    = WordIOFactory::load($temp_file_path);
    $objWriter  = WordIOFactory::createWriter($phpWord, 'PDF');

    return [$temp_file_path, $objWriter];
  }

  /**
   * Output to browser the temporary pdf file
   */
  public function displayPDF($output_file_name, $unlink = false)
  {
    [$temp_file_path, $objWriter] = self::_createWriterPDF();

    $output_file_name = strpos($output_file_name, '.pdf') > -1
      ? $output_file_name
      : "$output_file_name.pdf";

    // PDF header configuration
    header('Content-type: application/pdf');
    header('Content-Disposition: inline; filename="' . $output_file_name . '"');
    header('Cache-Control: max-age=0');

    $objWriter->save('php://output');

    // remove temp file
    unlink($temp_file_path);

    if (!$this->file_post && $unlink)
      unlink($this->relative_file_path);
  }

  /**
   * Output to browser as attachment for download
   *
   * @param output_file_name  - download file name
   * @param downloadAs        - download type
   */
  public function download($output_file_name, $download_as = 'pdf')
  {
    if ($download_as == 'pdf') {
      [$temp_file_path, $objWriter] = self::_createWriterPDF();

      $output_file_name = strpos($output_file_name, '.pdf') > -1
        ? $output_file_name
        : "$output_file_name.pdf";

      header('Content-type: application/pdf');
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Cache-Control: max-age=0');

      $objWriter->save('php://output');

      // remove temp file
      unlink($temp_file_path);
    }

    if ($download_as == 'docx') {
      $output_file_name = strpos($output_file_name, '.docx') > -1
        ? $output_file_name
        : "$output_file_name.docx";

      header("Content-Description: File Transfer");
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      header('Content-Transfer-Encoding: binary');
      header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      header('Expires: 0');

      $this->word_obj->saveAs("php://output");
    }

    if ($download_as == 'doc') {
      $output_file_name = strpos($output_file_name, '.doc') > -1
        ? $output_file_name
        : "$output_file_name.doc";

      header("Content-Description: File Transfer");
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      header('Content-Transfer-Encoding: binary');
      header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      header('Expires: 0');

      $this->word_obj->saveAs("php://output");
    }

    if ($download_as == 'odt') {
      // header("Content-Description: File Transfer");
      // header('Content-Disposition: attachment; filename="'.$filename.'.odt"');
      // header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      // header('Content-Transfer-Encoding: binary');
      // header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      // header('Expires: 0');

      // $this->word_obj->saveAs("php://output");
      // exit;
    }
  }

  public function save() // save to server
  {
    # code...
  }
}

?>