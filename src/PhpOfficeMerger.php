<?php

/**
 * Created  : 2023-1-21
 * @author    Tan Yong Xiang (zeikmanoffice@gmail.com)
 * @copyright Copyright (C) 2023 Tan Yong Xiang
 */

include_once '../lib/PDFMerger/PDFMerger.php';

use PDFMerger\PDFMerger;

class PhpOfficeMerger /* extends AnotherClass implements Interface */
{
  // function __construct()
  // {
  //   //
  // }

  static function mergeFiles($file_list)
  {
    # code...
  }

  // https://www.php.net/manual/en/features.file-upload.multiple.php
  static function reArrangeArrayFiles($file_post)
  {
    foreach ($file_post as $key => $all) {
      foreach ($all as $i => $val) {
        $file_ary[$i][$key] = $val;
      }
    }

    return $file_ary;
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
  static function mergePDF($output_as = 'browser', $pdf_files_arr = [], $output_name = 'pdf_merge')
  {
    // $filename = $this->file_prefix . '_' . $this->_now_date . '_' . $this->_now_time . '.pdf';
    $output_name = strpos($output_name, ".pdf") > -1
      ? $output_name
      : "$output_name.pdf";

    $pdf_merger = new PDFMerger;

    if (count($pdf_files_arr) > 0) {
      // add PDF only if exists
      foreach ($pdf_files_arr as $key => $pdf) {
        if (pathinfo($pdf)['extension'] == 'pdf' && file_exists($pdf))
          $pdf_merger->addPDF($pdf, 'all');
      }

      // output to browser As
      if (in_array($output_as, ['file', 'download', 'string', 'browser'], true))
        $pdf_merger->merge($output_as, $output_name);
      else
        return null;
    }

    return null;
  }
}

?>