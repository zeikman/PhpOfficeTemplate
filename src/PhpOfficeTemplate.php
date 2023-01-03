<?php

/**
 * Filling system data into document/spreadsheet templates.
 *
 * Created  : 2022-12-25
 * @author    Tan Yong Xiang (zeikmanoffice@gmail.com)
 * @copyright Copyright (C) 2022 Tan Yong Xiang
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

require '../vendor/autoload.php';
// require_once 'vendor/phpoffice/phpword/bootstrap.php';

include_once 'PhpSpreadsheetTemplate.php';
include_once 'PhpWordTemplate.php';
include_once '../lib/PHPExcel/Classes/PHPExcel.php';

use Mpdf\Mpdf;

use NcJoes\OfficeConverter\OfficeConverter;

/**
 * resolve Special Characters (ampersand) issue
 *
 *  - https://github.com/PHPOffice/PHPWord/issues/401
 */
\PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);

/**
 * PhpOfficeTemplate
 */
class PhpOfficeTemplate {

  public $filename        = null;
  public $sheetname       = null;
  public $file_post       = null; // for direct process without upload file to server
  public $target_dir      = null;
  public $data_main       = null;
  public $data_pool       = null;
  public $autoload_path   = null;
  public $filename_prefix = 'Template';

  private $file_type      = null;
  private $php_obj        = null;
  private $empty_if_var_unfound = false;

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

    // assign public variables
    $this->target_dir = $args['target_dir']
      ? $args['target_dir']
      : '';

    $this->filename = $args['file_name']
      ? $args['file_name']
      : '';

    $this->file_post = $args['file_post']
      ? $args['file_post']
      : '';

    $this->filename_prefix = $args['file_prefix']
      ? $args['file_prefix']
      : 'Template';

    $this->sheetname = $args['sheet_name']
      ? $args['sheet_name']
      : 'template';

    $this->empty_if_var_unfound = gettype($args['enable_empty_space']) == 'boolean'
      ? $args['enable_empty_space']
      : false;

    // $this->data_main = $args['main']
    //   ? $args['main']
    //   : [];

    $this->data_pool = $args['data']
      ? $args['data']
      : []; // in form [${variable_name} => 'data']
    // var_dump($this->data_pool); exit;

    // $this->autoload_path = $args['path']
    //   ? $args['path']
    //   : '../../../';

    $this->setFilenamePrefix($this->filename_prefix);

    if ($this->filename) {
      // check template existence
      $file_located = $this->target_dir . $this->filename;

      if (!$this->file_post && !file_exists($file_located)) {
        die(nl2br("PhpOfficeTemplate Error:\nMessage: Template file not found : $file_located"));
        exit;

      } else {
        // auto identify file type
        $spreadsheet_extension  = ['.xlsx', '.xls', '.ods'];
        $word_extension         = ['.docx', '.doc', '.cdt'];

        if ($this->_contains($this->filename, $spreadsheet_extension))
          $this->file_type = 'spreadsheet';

        if ($this->_contains($this->filename, $word_extension))
          $this->file_type = 'word';

        if ($this->file_type == 'spreadsheet') {
          if (!$this->filename)
            die(nl2br("PhpOfficeTemplate Error\nMessage: No filename specified."));

          if (!$this->sheetname)
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
            'file_name'   => $this->filename,
            'file_prefix' => $this->filename_prefix,
            'file_post'   => $this->file_post,
            'sheet_name'  => $this->sheetname,
            'enable_empty_space' => $this->empty_if_var_unfound
          ]);

          // $this->php_obj = $php_obj->getPhpSpreadsheet();
          //*/

          if (count($this->data_pool) > 0)
            $this->php_obj->substituteCell($this->data_pool);

        } else if ($this->file_type == 'word') {
          $this->php_obj = new PhpWordTemplate([
            'target_dir'  => $this->target_dir,
            'file_name'   => $this->filename,
            'file_prefix' => $this->filename_prefix,
            'file_post'   => $this->file_post,
            'sheet_name'  => $this->sheetname,
            'enable_empty_space' => $this->empty_if_var_unfound
          ]);

          // $this->php_obj = $php_obj->getPhpWord();

          if (count($this->data_pool) > 0)
            $this->php_obj->substituteCell($this->data_pool);

        } else {
          die(nl2br("PhpOfficeTemplate Error:\nMessage: Unsupported file type."));
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
    $fileType = PHPExcel_IOFactory::identify($this->filename);
    $objReader = PHPExcel_IOFactory::createReader($fileType);
    $objReader->setLoadSheetsOnly($this->sheetname);

    $this->php_obj = $objReader->load($this->filename);
    $sheet = $this->php_obj->getActiveSheet();

    if ($sheet == NULL) {
      die(nl2br("_initPhpExcel Error:\nMessage: Worksheet '$this->sheetname' not found."));
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
   * Return PhpOffice object
   */
  public function getPhpOfficeObject()
  {
    return $this->php_obj;
  }

  /**
   * Set/Update PhpOffice object
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
  public function setOrientation()
  {
    # code...
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

  public function combinePDF()
  {
    # code...
  }

  /**
   * Output the result
   *
   * @param method  - output method
   * @param type    - method type
   */
  public function output()
  {
    $args = [
      ...[
        'method'  => 'default',
        'type'    => 'pdf',
        'unlink'  => false
      ],
      ...func_get_args()
        ? func_get_args()[0]
        : []
    ];

    $method = $args['method'];
    $type   = $args['type'];
    $unlink = $args['unlink'];

    $filename = $this->filename_prefix . '_' . $this->_now_date . '_' . $this->_now_time;

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
          $this->php_obj->save();
        break;

      default:
        if ($this->php_obj)
          $this->php_obj->displayPDF($filename, $unlink);
        break;
    }

    exit;
  }
}

?>