<?php

/**
 * Filling system data into document/spreadsheet templates.
 *
 * Created  : 2022-12-25
 * @author    Tan Yong Xiang (zeikmanoffice@gmail.com)
 * @copyright Copyright (C) 2022 Tan Yong Xiang
 *
 * TODO: thing to specified in README.md for demo purpose
 * ### Things To Do :
 *
 * 1. create folder 'uploaded_template' inside demo/
 *
 * 2. change demo/uploaded_template owner
 *  # chown <your-user>:apache demo/uploaded_template
 *
 * 3. change demo/uploaded_template mod
 *  # chmod 775 demo/uploaded_template
 *
 * Install composer dependencies
 *  # composer install --ignore-platform-reqs
 */

/**
 * Prerequisite Libraries
 *
 * https://github.com/PHPOffice/PhpSpreadsheet
 * https://github.com/PHPOffice/PHPWord
 * https://github.com/mpdf/mpdf
 * https://github.com/ncjoes/office-converter
 * https://github.com/PHPOffice/PHPExcel
 */

// require '../vendor/autoload.php';
// require_once 'vendor/phpoffice/phpword/bootstrap.php';

include_once 'PhpSpreadsheetTemplate.php';
include_once 'PhpWordTemplate.php';
// include_once '../lib/PHPExcel/Classes/PHPExcel.php';
// include_once '../lib/PDFMerger/PDFMerger.php';

use PDFMerger\PDFMerger;

// use Mpdf\Mpdf;

// use NcJoes\OfficeConverter\OfficeConverter;

/**
 * PhpOfficeTemplate
 */
class PhpOfficeTemplate
{
  public $file_name   = null;
  public $sheet_name  = null;
  public $file_post   = null; // for direct process without upload file to server
  public $target_dir  = null;
  public $data_main   = null;
  public $data_pool   = null;
  public $file_prefix = 'Template';

  private $file_type  = null;
  private $php_obj    = null;

  private $enable_empty_space = false;
  private $enable_office_convertor = false;

  private $_now_datetime  = null;
  private $_now_date      = null;
  private $_now_time      = null;
  private $_now_day       = null;
  private $_now_month     = null;
  private $_now_year      = null;
  private $_now_hour      = null;
  private $_now_minute    = null;
  private $_now_second    = null;

  function __construct()
  {
    // init datetime properties
    $this->_now_datetime  = getdate(date('U'));
    $this->_now_day       = $this->_now_datetime['mday'];
    $this->_now_month     = str_pad($this->_now_datetime['mon'], 2, 0, STR_PAD_LEFT);
    $this->_now_year      = $this->_now_datetime['year'];
    $this->_now_hour      = str_pad($this->_now_datetime['hours'], 2, 0, STR_PAD_LEFT);
    $this->_now_minute    = str_pad($this->_now_datetime['minutes'], 2, 0, STR_PAD_LEFT);
    $this->_now_second    = str_pad($this->_now_datetime['seconds'], 2, 0, STR_PAD_LEFT);
    $this->_now_date      = "$this->_now_day-$this->_now_month-$this->_now_year";
    $this->_now_time      = "$this->_now_hour-$this->_now_minute-$this->_now_second";

    $args = func_get_arg(0);
    // var_dump($args); exit;

    $this->target_dir = isset($args['target_dir'])
      ? $args['target_dir']
      : '';

    // file name String
    $this->file_name = isset($args['file_name'])
      ? $args['file_name']
      : '';

    // PHP $_FILES Object
    $this->file_post = isset($args['file_post'])
      ? $args['file_post']
      : '';

    $this->file_prefix = isset($args['file_prefix'])
      ? $args['file_prefix']
      : 'Template';

    $this->sheet_name = isset($args['sheet_name'])
      ? $args['sheet_name']
      : 'template';

    $this->enable_empty_space = gettype($args['enable_empty_space']) == 'boolean'
      ? $args['enable_empty_space']
      : false;

    $this->enable_office_convertor = gettype($args['enable_office_convertor']) == 'boolean'
      ? $args['enable_office_convertor']
      : false;

    // $this->data_main = $args['main']
    //   ? $args['main']
    //   : [];

    $this->data_pool = $args['data']
      ? $args['data']
      : []; // in form [${variable_name} => 'data']
    // var_dump($this->data_pool); exit;

    $this->setFilenamePrefix($this->file_prefix);

    if ($this->file_name) {
      // check template existence
      $file_located = $this->target_dir . $this->file_name;

      if (!$this->file_post && !file_exists($file_located)) {
        die(nl2br("PhpOfficeTemplate Error:\nMessage: Template file not found : $file_located"));
        exit;

      } else {
        // auto identify file type
        $spreadsheet_extension  = ['.xlsx', '.xls', '.ods'];
        $word_extension         = ['.docx', '.doc', '.odt'];

        if ($this->_contains($this->file_name, $spreadsheet_extension))
          $this->file_type = 'spreadsheet';

        if ($this->_contains($this->file_name, $word_extension))
          $this->file_type = 'word';

        if ($this->file_type == 'spreadsheet') {
          if (!$this->file_name)
            die(nl2br("PhpOfficeTemplate Error\nMessage: No filename specified."));

          if (!$this->sheet_name)
            die(nl2br("PhpOfficeTemplate Error:\nMessage: No sheetname specified."));

          /*/
          // PhpExcel - fall back to utilise PhpExcel as it can render embed image
          $this->_initPhpExcel();
          /*/
          // PhpSpreadsheet
          // $this->_initPhpSpreadsheet();

          // $php_obj = new PhpSpreadsheetTemplate([
          $this->php_obj = new PhpSpreadsheetTemplate([
            'target_dir'  => $this->target_dir,
            'file_name'   => $this->file_name,
            'file_prefix' => $this->file_prefix,
            'file_post'   => $this->file_post,
            'sheet_name'  => $this->sheet_name,
            'enable_empty_space' => $this->enable_empty_space,
          ]);

          // $this->php_obj = $php_obj->getPhpSpreadsheet();
          //*/

          if (count($this->data_pool) > 0)
            $this->php_obj->substituteCell($this->data_pool);

        } else if ($this->file_type == 'word') {
          $this->php_obj = new PhpWordTemplate([
            'target_dir'  => $this->target_dir,
            'file_name'   => $this->file_name,
            'file_prefix' => $this->file_prefix,
            'file_post'   => $this->file_post,
            'sheet_name'  => $this->sheet_name,
            'enable_empty_space'      => $this->enable_empty_space,
            'enable_office_convertor' => $this->enable_office_convertor,
          ]);

          // $this->php_obj = $php_obj->getPhpWord();

          if (count($this->data_pool) > 0)
            $this->php_obj->substituteCell($this->data_pool);

        } else {
          die(nl2br("PhpOfficeTemplate Error:\nMessage: Unsupported file type."));
          exit;
        }
      }
    }
  }

  /* ========== ========== ========== ========== ========== ========== ========== ==========
   * Private Methods
   */

  // https://stackoverflow.com/questions/13795789/check-if-string-contains-word-in-array
  private function _contains($str, array $arr)
  {
    foreach ($arr as $a) {
      if (stripos($str, $a) !== false)
        return true;
    }

    return false;
  }

  /**
   * PhpExcel
   */
  private function _initPhpExcel()
  {
    /*/
    // Specify form template path
    $fileType = PHPExcel_IOFactory::identify($this->file_name);
    $objReader = PHPExcel_IOFactory::createReader($fileType);
    $objReader->setLoadSheetsOnly($this->sheet_name);

    $this->php_obj = $objReader->load($this->file_name);
    $sheet = $this->php_obj->getActiveSheet();

    if ($sheet == NULL) {
      die(nl2br("_initPhpExcel Error:\nMessage: Worksheet '$this->sheet_name' not found."));
      exit;
    }

    $sheet->setTitle('template_' . time() . rand(1, 1000));

    $sheet
      ->getPageSetup()
      ->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
    //*/
  }

  private function _getKeysInCell()
  {
    # code...
  }

  private function _substituteCellPhpWord()
  {
    # code...
  }

  private function _createCheckList()
  {
    # code...
  }

  /* ========== ========== ========== ========== ========== ========== ========== ==========
   * Public Methods
   */

  public function addWorkSheet()
  {
    # code...
  }

  /**
   * Return PhpOfficeTemplate object
   */
  public function getPhpOfficeObject()
  {
    return $this->php_obj;
  }

  /**
   * Set/Update PhpOfficeTemplate object
   *
   * @param phpOfficeObject - phpoffice object
   */
  public function setPhpOfficeObject($phpOfficeObject)
  {
    $this->php_obj = $phpOfficeObject;
  }

  public function cloneWorkSheet()
  {
    # code...
  }

  public function addSheet()
  {
    # code...
  }

  public function getCreationDateTime()
  {
    # code...
  }

  public function fillCellValue()
  {
    # code...
  }

  public function fillForm()
  {
    # code...
  }

  /**
   * Change pdf renderer class
   *
   * @param renderer - PDF renderer class name
   */
  public function setPdfRenderer($renderer = null)
  {
    $this->php_obj->setPdfRenderer($renderer);
  }

  /**
   * Change page orientation
   *
   * @param orientation - page orientation
   */
  public function setOrientation($orientation = 'portrait')
  {
    // landscape
    $this->php_obj->setOrientation($orientation);
  }

  /**
   * Set file name prefix
   *
   * @param prefix - file name prefix
   */
  public function setFilenamePrefix()
  {
    # code...
  }

  /**
   * Get file name prefix
   */
  public function getFilenamePrefix()
  {
    # code...
  }

  public function mergePDF()
  {
    # code...
  }

  /**
   * Merge multiple PDF files into single PDF file
   *
   * NOTE: Purely for PDF combination only
   *
   * @param outputTo    - Output result to [browser, file, download, string]
   * @param pdfFileArr  - array of temporary files
   */
  // static function mergePDF($output_as = 'browser')
  // static function mergeFiles($output_as = 'browser', $pdf_files_arr = [], $output_name = 'pdf_merge')
  // {
  //   // $filename = $this->file_prefix . '_' . $this->_now_date . '_' . $this->_now_time . '.pdf';
  //   $output_name = strpos($output_name, ".pdf") > -1
  //     ? $output_name
  //     : "$output_name.pdf";

  //   $pdf_merger = new PDFMerger;

  //   if (count($pdf_files_arr) > 0) {
  //     // add PDF only if exists
  //     foreach ($pdf_files_arr as $key => $pdf) {
  //       if (pathinfo($pdf)['extension'] == 'pdf' && file_exists($pdf))
  //         $pdf_merger->addPDF($pdf, 'all');
  //     }

  //     // output to browser As
  //     if (in_array($output_as, ['file', 'download', 'string', 'browser'], true))
  //       $pdf_merger->merge($output_as, $output_name);
  //     else
  //       return null;
  //   }

  //   return null;
  // }

  /**
   * Output the result
   *
   * @param method  - output method
   * @param type    - method type
   */
  public function output()
  {
    $get_args = func_get_args()
      ? func_get_args()[0]
      : [];

    $default_args = [
      'method'  => 'default',
      'type'    => 'pdf',
      'unlink'  => false
    ];

    $args = array_merge($default_args, $get_args);

    $method = $args['method'];
    $type   = $args['type'];
    $unlink = $args['unlink'];

    $output_file_name = $args['output_file_name'];

    $filename = $output_file_name
      ? $output_file_name
      : $this->file_prefix . '_' . $this->_now_date . '_' . $this->_now_time;

    switch ($method) {
      case 'browser':
        if ($this->php_obj)
          $this->php_obj->displayPDF($filename, $unlink);
        break;

      case 'download':
        if ($this->php_obj)
          $this->php_obj->download($filename, $type);
        break;

      case 'server':
        if ($this->php_obj)
          return $this->php_obj->save($filename, $type);
        break;

      default:
        if ($this->php_obj)
          $this->php_obj->displayPDF($filename, $unlink);
        break;
    }

    exit;
  }
}