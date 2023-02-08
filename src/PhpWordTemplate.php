<?php

require '../vendor/autoload.php';

// use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Settings as WordSettings;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;

use NcJoes\OfficeConverter\OfficeConverter;

/**
 * resolve Special Characters (ampersand) issue
 *
 *  - https://github.com/PHPOffice/PHPWord/issues/401
 */
\PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);

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
  private $enable_empty_space;
  private $enable_office_convertor;
  private $orientation;
  private $force_unlink;

  function __construct()
  {
    $args = func_get_arg(0);

    $this->target_dir   = $args['target_dir'];
    $this->file_name    = $args['file_name'];
    $this->file_prefix  = $args['file_prefix'];
    $this->file_post    = $args['file_post'];

    $this->pdf_renderer = isset($args['pdf_renderer'])
      ? strtolower($args['pdf_renderer'])
      : self::DEFAULT_RENDERER;

    $this->enable_empty_space = gettype($args['enable_empty_space']) == 'boolean'
      ? $args['enable_empty_space']
      : false;

    $this->enable_office_convertor = gettype($args['enable_office_convertor']) == 'boolean'
      ? $args['enable_office_convertor']
      : false;

    $this->relative_file_path = $this->target_dir . $this->file_name;

    $template_path = $this->file_post
      ? $this->file_post
      : $this->relative_file_path;

    if (pathinfo($this->file_name)['extension'] == 'doc') {
      $temporary_file_docx = self::_convertDocToDocx();

      if ($temporary_file_docx) {
        $this->file_name = $temporary_file_docx;
        $this->file_post = '';

        $this->relative_file_path = $this->target_dir . $this->file_name;

        $template_path = $this->relative_file_path;

        $this->word_obj = new TemplateProcessor($template_path);

        $this->force_unlink = true;

      } else {
        die(nl2br("PhpOfficeTemplate Error:\nMessage: Unsupported file type > doc."));
        exit;
      }
    }
    else
      $this->word_obj = new TemplateProcessor($template_path);
  }

  /**
   * Convert .doc to .docx if OfficeConverter found before passing to TemplateProcessor
   */
  private function _convertDocToDocx()
  {
    if (class_exists(OfficeConverter::class)) {
      if ($this->file_post) {
        $destination = $this->target_dir . 'temp_source_' . $this->file_name;

        if (move_uploaded_file($this->file_post, $destination)) {
          $temp_docx_file = str_replace('.doc', '.docx', 'temp_result_' . $this->file_name);

          $converter = new OfficeConverter($destination);

          $converter->convertTo($temp_docx_file);
          // var_dump($destination);
          // var_dump($temp_docx_file);

          unlink($destination);

          return $temp_docx_file;
        }

        return '';

      } else {
        // NOTE: file in server
        // 1. try convert file to docx
      }

    } else {
      return '';
    }
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
    if ($this->word_obj && count($data) > 0) {
      // TODO: set image value
      // https://stackoverflow.com/questions/71717015/phpword-not-able-to-replace-existing-image-in-the-docx-file
      // https://stackoverflow.com/questions/24018003/how-to-add-set-images-on-phpoffice-phpword-template

      if ($data['image']) {
        $image_data = $data['image'];

        unset($data['image']);

        foreach ($image_data as $var => $value) {
          // var_dump($var);
          // var_dump($value);
          // var_dump(file_exists($value));

          if (file_exists($value)) {
            // if (pathinfo($var)['extension'] == 'jpg') {
            //   // var_dump("word/media/$var"); exit;

            //   $this->word_obj->zip()->addFromString("word/media/1.jpg", file_get_contents($value));
            //   // $this->word_obj->zip()->addFromString("word/media/image1.jpg", file_get_contents($value));
            //   // $this->word_obj->zip()->addFromString("word/media/$var", file_get_contents($value));
            //   // $this->word_obj->zip()->AddFromString("word/media/$var", file_get_contents($value));
            //   // $this->word_obj->zip()->pclzipAddFromString("word/media/$var", file_get_contents($value));

            //   // var_dump(file_get_contents($value));

            // } else {
              $this->word_obj->setImageValue($var, $value);
            // }
          }

          if ((is_array($value) || is_callable($value)) && $value['path']) {
            $this->word_obj->setImageValue($var, $value);
          }
        }

        // exit;
      }

      $this->word_obj->setValues($data);

      // replace with empty space [' '] if variable unfound in data pool
      if ($this->enable_empty_space) {
        $unfoundVarList = [];

        foreach ($this->word_obj->getVariables() as $key => $v)
          $unfoundVarList[$v] = ' ';

        $this->word_obj->setValues($unfoundVarList);
      }

      // var_dump($this->word_obj->getVariables()); exit;
    }
  }

  /**
   * Change page orientation
   *
   * @param orientation - page orientation
   */
  public function setOrientation($orientation = 'portrait')
  {
    $this->orientation = $orientation;

    // $pageSetup = $this->word_obj->getActiveSheet()->getPageSetup();

    // if ($orientation == 'landscape')
    //   $pageSetup->setOrientation(SpreadsheetPageSetup::ORIENTATION_LANDSCAPE);

    // if ($orientation == 'portrait')
    //   $pageSetup->setOrientation(SpreadsheetPageSetup::ORIENTATION_PORTRAIT);
  }

  private function _getFileName($file_path) {
    $pos = strrpos($file_path, "/");

    return $pos > -1
      ? substr($file_path, $pos + 1)
      : $file_path;
  }

  private function _getTemporaryFilePath()
  {
    $temp_file_name = self::_getFileName($this->file_name);

    return $this->target_dir . "temp_$temp_file_name";

  }

  private function _saveTemplateProcessor()
  {
    $temp_file_path = self::_getTemporaryFilePath();

    // [$this->word_obj] is PhpWord template processor
    $this->word_obj->saveAs($temp_file_path);

    return $temp_file_path;
  }

  /**
   * Change pdf renderer class
   *
   * @param renderer - PDF renderer class name
   */
  public function setPdfRenderer($renderer = null)
  {
    $pdf_available = ['tcpdf', 'mpdf', 'dompdf'];

    if (in_array($renderer, $pdf_available)) {
      $this->pdf_renderer = $renderer
        ? strtolower($renderer)
        : self::DEFAULT_RENDERER; // default
    }
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

  private function _createWriter($writer_type = 'pdf')
  {
    // currently support: 'ODText', 'RTF', 'Word2007', 'HTML', 'PDF'

    if ($writer_type == 'pdf') {
      self::_setPdfRenderer();

      $temp_file_path = self::_saveTemplateProcessor();

      $phpWord    = WordIOFactory::load($temp_file_path);
      $writer_obj = WordIOFactory::createWriter($phpWord, 'PDF');

      return [$temp_file_path, $writer_obj];
    }

    if ($writer_type == 'docx') {
      $temp_file_path = self::_saveTemplateProcessor();

      $phpWord    = WordIOFactory::load($temp_file_path);
      $writer_obj = WordIOFactory::createWriter($phpWord, 'Word2007');

      return [$temp_file_path, $writer_obj];
    }

    if ($writer_type == 'doc') {
      $temp_file_path = self::_saveTemplateProcessor();

      $phpWord    = WordIOFactory::load($temp_file_path);
      $writer_obj = WordIOFactory::createWriter($phpWord, 'Word2007');

      return [$temp_file_path, $writer_obj];
    }

    if ($writer_type == 'odt') {
      $temp_file_path = self::_saveTemplateProcessor();

      $phpWord    = WordIOFactory::load($temp_file_path);
      $writer_obj = WordIOFactory::createWriter($phpWord, 'ODText');

      return [$temp_file_path, $writer_obj];
    }

    return null;
  }

  /**
   * TODO: docx to pdf using phpword is worst
   *
   * Solution :
   * 1. people suggest to use libreoffice using exec()
   * 2. Need to install LibreOffice on system, download LibreOffice thru website with tar.gz
   *    > extract file : tar -xvf LibreOffice.tar.gz
   *    > move into file : cd LibreOffice_file/RPMS
   *    > install file : su -c 'yum install *.rpm'
   *
   * DocX => pdf styles missing
   * https://github.com/PHPOffice/PHPWord/issues/1139
   * https://stackoverflow.com/questions/54616086/no-styling-when-converting-docx-into-pdf-with-phpword/54660038#54660038
   */
  private function _createOfficeConvertor()
  {

    $temp_file_path = self::_saveTemplateProcessor();

    /*
     * Need to edit exec() function with following reference for the command
     * https://stackoverflow.com/questions/10169042/unable-to-run-oowriter-as-web-user
     */
    $converter = new OfficeConverter($temp_file_path);

    return [$temp_file_path, $converter];
  }

  private function _checkFileExtension($file_name, $extension)
  {
    return strpos($file_name, ".$extension") > -1
      ? $file_name
      : "$file_name.$extension";
  }

  /**
   * Output to browser the temporary pdf file
   */
  public function displayPDF($output_file_name, $unlink = false)
  {
    $output_file_name = self::_checkFileExtension($output_file_name, 'pdf');

    // https://stackoverflow.com/questions/44143604/php-check-if-use-a-valid-class

    // Office Convertor by ncjoes
    if ($this->enable_office_convertor && class_exists(OfficeConverter::class)) {
      [$temp_file_path, $converter] = self::_createOfficeConvertor();

      $pdf_path = str_replace('.docx', '.pdf', $temp_file_path);

      $temp_pdf_name = self::_getFileName($pdf_path);

      $converter->convertTo($temp_pdf_name);

      header('Content-type: application/pdf');
      header('Content-Disposition: inline; filename="' . $output_file_name . '"');
      header('Content-Transfer-Encoding: binary');
      header('Accept-Ranges: bytes');

      @readfile($pdf_path);

      // remove temp file
      unlink($temp_file_path);
      unlink($pdf_path);

      if (!$this->file_post && $unlink && file_exists($this->relative_file_path))
        unlink($this->relative_file_path);

      if ($this->force_unlink)
        unlink($this->relative_file_path);

    }
    // PHPWord PDF Writer
    else {
      header('Content-type: application/pdf');
      header('Content-Disposition: inline; filename="' . $output_file_name . '"');
      header('Cache-Control: max-age=0');

      [$temp_file_path, $writer_obj] = self::_createWriter('pdf');

      $writer_obj->save('php://output');

      // remove temp file
      unlink($temp_file_path);

      if (!$this->file_post && $unlink && file_exists($this->relative_file_path))
        unlink($this->relative_file_path);
    }
  }

  /**
   * Output to browser as attachment for download
   *
   * @param output_file_name  - download file name
   * @param downloadAs        - download type
   */
  public function download($output_file_name, $download_as = 'pdf')
  {
    $output_file_name = self::_checkFileExtension($output_file_name, strtolower($download_as));

    if ($download_as == 'pdf') {
      if ($this->enable_office_convertor && class_exists(OfficeConverter::class)) {
        [$temp_file_path, $converter] = self::_createOfficeConvertor();

        $pdf_path = str_replace('.docx', '.pdf', $temp_file_path);

        $temp_pdf_name = self::_getFileName($pdf_path);

        $converter->convertTo($temp_pdf_name);

        header('Content-type: application/pdf');
        header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
        header('Cache-Control: max-age=0');
        header('Content-Transfer-Encoding: binary');
        header('Accept-Ranges: bytes');

        @readfile($pdf_path);

        // remove temp file
        unlink($temp_file_path);
        unlink($pdf_path);

      } else {
        header('Content-type: application/pdf');
        header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
        header('Cache-Control: max-age=0');

        [$temp_file_path, $writer_obj] = self::_createWriter('pdf');

        $writer_obj->save('php://output');

        // remove temp file
        unlink($temp_file_path);
      }
    }

    if ($download_as == 'docx') {
      header("Content-Description: File Transfer");
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      header('Content-Transfer-Encoding: binary');
      header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      header('Expires: 0');

      $this->word_obj->saveAs("php://output");
    }

    if ($download_as == 'doc') {
      header("Content-Description: File Transfer");
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      header('Content-Transfer-Encoding: binary');
      header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      header('Expires: 0');

      $this->word_obj->saveAs("php://output");
    }

    if ($download_as == 'odt') {
      header("Content-Description: File Transfer");
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      header('Content-Transfer-Encoding: binary');
      header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      header('Expires: 0');

      $this->word_obj->saveAs("php://output");
    }

    if ($this->force_unlink)
      unlink($this->relative_file_path);
  }

  // save to server
  public function save($output_file_name, $save_as = 'pdf')
  {
    if (in_array($save_as, ['pdf', 'docx', 'doc', 'odt'], true)) {
      $output_file_name = self::_checkFileExtension($output_file_name, $save_as);

      $file_save_name = $output_file_name
        ? $output_file_name
        : $this->file_prefix . '_' . mt_rand(1, 100000) . '.pdf';

      $file_save_path = $this->target_dir . $file_save_name;

      if ($save_as == 'pdf') {
        if ($this->enable_office_convertor && class_exists(OfficeConverter::class)) {
          [$temp_file_path, $converter] = self::_createOfficeConvertor();

          $converter->convertTo($file_save_name);

        } else {
          [$temp_file_path, $writer_obj] = self::_createWriter($save_as);

          $writer_obj->save($file_save_path);
        }

        // remove temp file
        unlink($temp_file_path);

        if ($this->force_unlink)
          unlink($this->relative_file_path);

        return $file_save_path;

      } else {
        $temp_file_path = self::_saveTemplateProcessor();

        rename($temp_file_path, $file_save_path);

        if ($this->force_unlink)
          unlink($this->relative_file_path);

        return $file_save_path;
      }
    }

    return null;
  }
}
