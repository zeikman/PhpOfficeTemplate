<?php

require '../vendor/autoload.php';

// use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory as SpreadsheetIOFactory;
// use PhpOffice\PhpSpreadsheet\Settings as SpreadsheetSettings;
// use PhpOffice\PhpSpreadsheet\Calculation\Logical\Boolean;
use PhpOffice\PhpSpreadsheet\Cell\DataType as SpreadsheetDataType;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup as SpreadsheetPageSetup;

// use PhpOffice\PhpSpreadsheet\Reader\Csv as SpreadsheetReaderCsv;
// use PhpOffice\PhpSpreadsheet\Reader\Xls as SpreadsheetReaderXls;
// use PhpOffice\PhpSpreadsheet\Reader\Xlsx as SpreadsheetReaderXlsx;

// use PhpOffice\PhpSpreadsheet\Writer\Xlsx as SpreadsheetWriterXlsx;
use PhpOffice\PhpSpreadsheet\Writer\Pdf as SpreasheetPdf;
// use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf as SpreadsheetMpdf;

/**
 * PhpSpreadsheetTemplate
 */
class PhpSpreadsheetTemplate
{
  public const DEFAULT_RENDERER = 'mpdf';
  public const VAR_PATTERN      = '${'; // ${variable_name}

  private $spreadsheet_obj;
  private $target_dir;
  private $file_name;
  private $file_prefix;
  private $file_post;
  private $sheet_name;
  private $pdf_renderer;
  private $relative_file_path;
  private $enable_empty_space;
  private $setLoadSheetsOnly = true; // TODO: enhance : load multiple worksheet

  function __construct()
  {
    $args = func_get_arg(0);

    $this->target_dir   = $args['target_dir'];
    $this->file_name    = $args['file_name'];
    $this->file_prefix  = $args['file_prefix'];
    $this->file_post    = $args['file_post'];

    $this->sheet_name = $args['sheet_name']
      ? $args['sheet_name']
      : 'template';

    $this->pdf_renderer = isset($args['pdf_renderer'])
      ? strtolower($args['pdf_renderer'])
      : self::DEFAULT_RENDERER;

    $this->enable_empty_space = gettype($args['enable_empty_space']) == 'boolean'
      ? $args['enable_empty_space']
      : false;

    $this->relative_file_path = $this->target_dir . $this->file_name;

    // auto identify file type and load file
    // https://www.youtube.com/watch?v=p6ELMxvMyyE
    if ($this->file_post) {
      $php_reader = SpreadsheetIOFactory::createReader(SpreadsheetIOFactory::identify($this->file_post));
      $php_check  = $php_reader->load($this->file_post);
    } else {
      $php_reader = SpreadsheetIOFactory::createReader(SpreadsheetIOFactory::identify($this->relative_file_path));
      $php_check  = $php_reader->load($this->relative_file_path);
    }

    if ($php_check->getSheetByName($this->sheet_name) == null)
      die(nl2br("PhpSpreadsheetTemplate Error:\nMessage: Worksheet '$this->sheet_name' not found."));

    if ($this->setLoadSheetsOnly)
      $php_reader->setLoadSheetsOnly($this->sheet_name);

    $template_path = $this->file_post
      ? $this->file_post
      : $this->relative_file_path;

    $this->spreadsheet_obj = $php_reader->load($template_path);

    $this->spreadsheet_obj
      ->getActiveSheet()
      ->setTitle('template_' . time() . random_int(1, 1000));

    self::_setDefaultMargin();
  }

  private function _setDefaultMargin()
  {
    if ($this->pdf_renderer == 'mpdf') {
      $sheet = $this->spreadsheet_obj
        ->getActiveSheet();

      $sheet
        ->getPageSetup()
        ->setPaperSize(SpreadsheetPageSetup::PAPERSIZE_A4);

      // TODO: enhance so that user can set margin
      $margin = 0.75;
      $margin = 0.50;
      // $margin = 0.16;

      $sheet
        ->getPageMargins()
        // ->setLeft(0.2)
        // ->setRight($margin)
        ->setTop($margin)
        ->setBottom($margin);

      /* A4 default margins for PHPSpreadsheet
        private 'left' => float 0.16
        private 'right' => float 0.16
        private 'top' => float 1
        private 'bottom' => float 1
        private 'header' => float 0.51180555555556
        private 'footer' => float 0.51180555555556
      */

      /* A4 default margins for PHPExcel
        private 'left' => float 0.75
        private 'right' => float 0.75
        private 'top' => float 1
        private 'bottom' => float 1
        private 'header' => float 0.51180555555556
        private 'footer' => float 0.51180555555556
      */
    }
  }

  public function getPhpSpreadsheet()
  {
    return $this->spreadsheet_obj;
  }

  /**
   * Get all words begin with $ (dollar sign)
   *
   * @param subject - cell value
   */
  function _getKeysInCell($subject/* string */, $pattern = null)
  {
    $subject = $subject
      ? $subject
      : '';

    // detect '${'
    $pattern = $pattern
      ? $pattern
      : '/\$\{\w+/';

    preg_match_all($pattern, $subject, $matches);

    return $matches[0];

    // $testMsg = 'Good Morning, Mr. $myName, how are you $period ?';
    // preg_match_all('/\$\w+/', $testMsg, $matches);
    // var_dump($matches[0]);

    // $userinfo = "Name: <b>John Poul</b> <br> Title: <b>PHP Guru</b>";
    // preg_match_all ("/<b>(.*)<\/b>/U", $userinfo, $pat_array);
    // print $pat_array[0][0]." <br> ".$pat_array[0][1]."\n";
  }

  /**
   * Substitute variable in cell with value (Spreadsheet)
   *
   * @param enableEmptyValueIfUnfound - enable empty space substitution if variable unfound
   */
  public function substituteCell($data)
  {
    $sheet = $this->spreadsheet_obj->getActiveSheet();

    // NOTE: read & convert cell data into array

    // read all
    // $xlsData = $sheet->toArray(null, true, true, true);

    // get writable area
    $cellRange = $sheet->calculateWorksheetDimension();

    // read based on cell range
    $xlsData = $sheet->rangeToArray($cellRange, null, true, true, true);

    // number of row
    $nr = count($xlsData);

    for ($i = 1; $i <= $nr; $i++) {
      foreach ($xlsData[$i] as $cellCol => $cellData) {

        $newVarCount = count($this->_getKeysInCell($cellData));

        if ($cellData && $newVarCount > 0) {
          $cellCoordinate = $cellCol . $i;

          // replace with original value if no variable found
          $newValue = $cellData;

          if ($cellData && strpos($cellData, self::VAR_PATTERN) > -1)
            $newValue = strtr($cellData, $data);
          // var_dump('$newValue : '.$newValue);

          // replace with empty space [' '] if variable unfound in data pool
          if ($this->enable_empty_space && $cellData && strpos($cellData, self::VAR_PATTERN) > -1) {
            $unfoundVarList = [];

            // $search = $this->_getKeysInCell($newValue, '/\$\{\w+\}$/');
            $search = $this->_getKeysInCell($newValue, '/\$\{\w+\}/');
            // var_dump($search);

            foreach ($search as $key => $v)
              $unfoundVarList[$v] = ' ';

            $newValue = strtr($newValue, $unfoundVarList);
          }

          $sheet->setCellValueExplicit($cellCoordinate, $newValue, SpreadsheetDataType::TYPE_STRING);
        }
      }
    }

    // var_dump($this->enable_empty_space); exit;
  }

  /**
   * Set cell value
   *
   * @param cells - array of excel-like column row field
   * @param value - cell new value
   * @param sheetNumb - worksheet number
   */
  function fillCellValue($cell/* string */, $value/* any */, $sheetNumb = 0/* xxcel_sheet_number */, $dataType = SpreadsheetDataType::TYPE_STRING)
  {
    $this->spreadsheet_obj
      ->setActiveSheetIndex($sheetNumb);

    $this->spreadsheet_obj
      ->getActiveSheet()
      ->setCellValueExplicit($cell, $value, $dataType);
  }

  /**
   * Change page orientation
   *
   * @param orientation - page orientation
   */
  public function setOrientation($orientation = 'portriat')
  {
    $pageSetup = $this->spreadsheet_obj->getActiveSheet()->getPageSetup();

    if ($orientation == 'landscape')
      $pageSetup->setOrientation(SpreadsheetPageSetup::ORIENTATION_LANDSCAPE);

    if ($orientation == 'portriat')
      $pageSetup->setOrientation(SpreadsheetPageSetup::ORIENTATION_PORTRAIT);
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
    /*/
    $classPdf = My_Spreadsheet_Mpdf::class;

    $classPdf = My_Spreadsheet_Mpdf::class; // TODO: try using external Mpdf (install thru composer)
    $classPdf = My_Custom_MPDF::class; // TODO: try using external Mpdf (install thru composer)
    // $classPdf = OriginalMpdf::class; // TODO: try using external Mpdf (install thru composer)
    //*/

    if ($this->pdf_renderer == 'mpdf')
      $classPdf = SpreasheetPdf\Mpdf::class;

    if ($this->pdf_renderer == 'tcpdf')
      $classPdf = SpreasheetPdf\Tcpdf::class;

    if ($this->pdf_renderer == 'dompdf')
      $classPdf = SpreasheetPdf\Dompdf::class;

    SpreadsheetIOFactory::registerWriter('Pdf', $classPdf);
  }

  private function _createWriterPDF()
  {
    self::_setPdfRenderer();

    $objWriter = SpreadsheetIOFactory::createWriter($this->spreadsheet_obj, 'Pdf');

    return $objWriter;
  }

  /**
   * Output to browser the temporary pdf file
   */
  public function displayPDF($output_file_name, $unlink = false)
  {
    $objWriter = self::_createWriterPDF();

    $output_file_name = strpos($output_file_name, '.pdf') > -1
      ? $output_file_name
      : "$output_file_name.pdf";

    // PDF header configuration
    header('Content-type: application/pdf');
    header('Content-Disposition: inline; filename="' . $output_file_name . '"');
    header('Cache-Control: max-age=0');

    $objWriter->save('php://output');

    if (!$this->file_post && $unlink && file_exists($this->relative_file_path))
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
    $download_as = strtolower($download_as);
    $filename = random_int(1, 100000) . '.pdf';

    if ($download_as == 'pdf') {
      $objWriter = self::_createWriterPDF();

      $output_file_name = strpos($output_file_name, '.pdf') > -1
        ? $output_file_name
        : "$output_file_name.pdf";

      // Redirect output to a client’s web browser (PDF)
      header('Content-type: application/pdf');
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Cache-Control: max-age=0');

      $objWriter->save('php://output');
    }

    if ($download_as == 'xlsx') {
      $output_file_name = strpos($output_file_name, '.xlsx') > -1
        ? $output_file_name
        : "$output_file_name.xlsx";

      // Redirect output to a client’s web browser (Xlsx)
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Cache-Control: max-age=0');

      $objWriter = SpreadsheetIOFactory::createWriter($this->spreadsheet_obj, 'Xlsx');
      $objWriter->save('php://output');
    }

    if ($download_as == 'xls') {
      $output_file_name = strpos($output_file_name, '.xls') > -1
        ? $output_file_name
        : "$output_file_name.xls";

      // Redirect output to a client’s web browser (Xls)
      header('Content-Type: application/vnd.ms-excel');
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Cache-Control: max-age=0');

      $objWriter = SpreadsheetIOFactory::createWriter($this->spreadsheet_obj, 'Xls');
      $objWriter->save('php://output');
    }

    if ($download_as == 'ods') {
      $output_file_name = strpos($output_file_name, '.ods') > -1
        ? $output_file_name
        : "$output_file_name.ods";

      // Redirect output to a client’s web browser (Ods)
      header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');
      header('Content-Disposition: attachment; filename="' . $output_file_name . '"');
      header('Cache-Control: max-age=0');

      $objWriter = SpreadsheetIOFactory::createWriter($this->spreadsheet_obj, 'Ods');
      $objWriter->save('php://output');
    }
  }

  public function save() // save to server
  {
    # code...
  }
}

?>