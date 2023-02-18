<?php

// Error reporting
$debug = TRUE;
$debug = FALSE;

error_reporting(E_ALL);

ini_set('display_errors', $debug);
ini_set('display_startup_errors', $debug);

// Set PHP script load waiting time limit
// set_time_limit(0);

// https://stackoverflow.com/questions/49116886/uncaught-mpdf-mpdfexception-the-html-code-size-is-larger-than-pcre-backtrack-li/49126984#49126984
ini_set("pcre.backtrack_limit", "5000000");

// Include PhpOfficeTemplate
include_once '../src/PhpOfficeTemplate.php';
include_once '../src/PhpOfficeMerger.php';
// include_once '../lib/number2word.php';

/*/
$user_agent = $_SERVER['HTTP_USER_AGENT'];
$browser = get_browser(null, true);

var_dump("$user_agent<br /><br />");
// var_dump($browser);
// print_r($browser);
//*/

// https://stackoverflow.com/questions/8754080/how-to-get-exact-browser-name-and-version
function getBrowser()
{
  $u_agent  = $_SERVER['HTTP_USER_AGENT'];
  $bname    = 'Unknown';
  $platform = 'Unknown';
  $version  = "";

  //First get the platform?
  if (preg_match('/linux/i', $u_agent)) {
    $platform = 'linux';

  } elseif (preg_match('/macintosh|mac os x/i', $u_agent)) {
    $platform = 'mac';

  } elseif (preg_match('/windows|win32/i', $u_agent)) {
    $platform = 'windows';
  }

  // Next get the name of the useragent yes seperately and for good reason
  if (preg_match('/MSIE/i', $u_agent) && !preg_match('/Opera/i', $u_agent)) {
    $bname = 'Internet Explorer';
    $ub = "MSIE";

  } elseif (preg_match('/Edg/i', $u_agent)) {
    $bname = 'Microsoft Edge';
    $ub = "Edg";

  } elseif (preg_match('/Firefox/i', $u_agent)) {
    $bname = 'Mozilla Firefox';
    $ub = "Firefox";

  } elseif (preg_match('/Chrome/i', $u_agent)) {
    $bname = 'Google Chrome';
    $ub = "Chrome";

  } elseif (preg_match('/Safari/i', $u_agent)) {
    $bname = 'Apple Safari';
    $ub = "Safari";

  } elseif (preg_match('/Opera/i', $u_agent)) {
    $bname = 'Opera';
    $ub = "Opera";

  } elseif (preg_match('/Netscape/i', $u_agent)) {
    $bname = 'Netscape';
    $ub = "Netscape";
  }

  // finally get the correct version number
  $known = array('Version', $ub, 'other');
  $pattern = '#(?<browser>' . join('|', $known) . ')[/ ]+(?<version>[0-9.|a-zA-Z.]*)#';

  if (!preg_match_all($pattern, $u_agent, $matches)) {
    // we have no matching number just continue
  }

  // see how many we have
  $i = count($matches['browser']);

  if ($i != 1) {
    // we will have two since we are not using 'other' argument yet
    // see if version is before or after the name
    $version = strripos($u_agent, "Version") < strripos($u_agent, $ub)
      ? $matches['version'][0]
      : $matches['version'][1];

  } else {
    $version = $matches['version'][0];
  }

  // check if we have a number
  if ($version == null || $version == "")
    $version = "?";

  return array(
    'userAgent' => $u_agent,
    'name'      => $bname,
    'version'   => $version,
    'platform'  => $platform,
    'pattern'   => $pattern
  );
}

/*/
// now try it
$user_agent = $_SERVER['HTTP_USER_AGENT'];
print_r("$user_agent<br /><br />");

$ua = getBrowser();
$yourbrowser = "Your browser: " . $ua['name'] . "<br />Version: " . $ua['version'] . "<br />Platform: " . $ua['platform'] . "<br />User Agent: " . $ua['userAgent'];
print_r($yourbrowser);
exit;
//*/

// https://www.php.net/manual/en/features.file-upload.multiple.php
function reArrayFiles(&$file_post)
{
  $isMulti = is_array($file_post['name']);
  $file_count = $isMulti
    ? count($file_post['name'])
    : 1;
  $file_keys = array_keys($file_post);

  $file_ary = array();
  // $file_count = count($file_post['name']);
  // $file_keys = array_keys($file_post);

  for ($i = 0; $i < $file_count; $i++) {
    foreach ($file_keys as $key) {
      if ($isMulti)
        $file_ary[$i][$key] = $file_post[$key][$i];
      else
        $file_ary[$i][$key] = $file_post[$key];
    }
  }

  return $file_ary;
}

// function reArrangeArrayFiles($file_post)
// {
//   foreach ($file_post as $key => $all) {
//     foreach ($all as $i => $val) {
//       $file_ary[$i][$key] = $val;
//     }
//   }

//   return $file_ary;
// }

// var_dump($_POST); exit;

$filename_option = isset($_POST['filename_option'])
  ? $_POST['filename_option']
  : 'My Output';

$output_option = isset($_POST['output_option'])
  ? $_POST['output_option']
  : 'browser';

$download_option = isset($_POST['download_option'])
  ? $_POST['download_option']
  : 'pdf';

$orientation_option = isset($_POST['orientation_option'])
  ? $_POST['orientation_option']
  : 'default';

$pdf_option = isset($_POST['pdf_option'])
  ? $_POST['pdf_option']
  : 'default' ;

$emptyspace_option = isset($_POST['emptyspace_option']) && $_POST['emptyspace_option'] == 'true'
  ? true
  : false;

$offconverter_option = isset($_POST['offconverter_option']) && $_POST['offconverter_option'] == 'true'
  ? true
  : false;

// file name
$file_name = isset($_POST['filename'])
  ? $_POST['filename']
  : '';

// file name when download
$file_prefix = isset($_POST['fileprefix'])
  ? $_POST['fileprefix']
  : '';

// directory to store template file
$target_dir = isset($_POST['targetdir'])
  ? $_POST['targetdir']
  : 'uploaded_template/';

// prepare data
$data = [
  '${cp_name}'        => 'Testing Company',
  '${cp_roc}'         => 'Testing ROC',
  '${cp_ssm}'         => 'Testing SSM',
  '${my_name}'        => 'Testing Name',
  '${my_dob}'         => '1 Jan 2023',
  '${my_age}'         => '18',
  '${my_city}'        => 'Akihabara',
  '${my_state}'       => 'Tokyo',
  '${my_country}'     => 'Japan',
  '${my_id}'          => 'Testing ID',
  '${my_contact}'     => 'Testing Contact',
  '${my_addr}'        => 'Testing Address
Testing Address Line 2',
  '${my_passport}'    => 'Testing Passport',
  '${my_marital}'     => 'Single',
  '${my_job}'         => 'Testing Job',
  '${my_gender_code}' => 'M',
  '${my_gender}'      => 'Male',

  // NOTE: only for phpword
  'replaceimage' => [
    // TODO: test replace image
    // 'phpoffice.jpg' => 'template_files/phpword.jpg',

    // 'image1' => 'https://legacy.gscdn.nl/archives/images/HassVivaCatFight.jpg',
    'image1' => 'template_files/200x200-w.png',
    'image2' => 'template_files/phpoffice.jpg',
    'image3' => 'template_files/200x200.png',
  ],

  'image' => [
    'img_phpoffice' => 'template_files/phpoffice.jpg',
    'img_phpoffice_2' => 'template_files/phpoffice.jpg',
    'img_phpoffice_3' => 'template_files/phpoffice.jpg',
    'img_phpoffice_4' => 'template_files/phpoffice.jpg',
    'img_phpoffice_5' => 'template_files/phpoffice.jpg',
    'img_phpoffice_6' => 'template_files/phpoffice.jpg',
    'img_phpoffice_7' => 'template_files/phpoffice.jpg',

    'img_phpword' => 'template_files/phpword.jpg',
    'img_phpword_2' => 'template_files/phpword.jpg',
    'img_phpword_3' => array(
      'path'    => 'https://legacy.gscdn.nl/archives/images/HassVivaCatFight.jpg',
      'width'   => 100,
      // 'height'  => 100,
      // 'height'  => '',
      // 'height'  => 50,
      // 'ratio'   => false, // set to false if don't want resize in ratio
    ),

    // // https://stackoverflow.com/questions/25772821/cannot-use-object-of-type-closure-as-array
    // 'img_phpword_2' => call_user_func_array(function () {
    //   return array(
    //     'path'    => 'template_files/phpword.png',
    //     'width'   => 50,
    //     'height'  => 50,
    //     'ratio'   => false,
    //   );
    // }, []),

    // invalid syntax
    // '${img_phpword}' => 'template_files/phpword.png',

    // invalid image
    // 'img_phpword' => 'template_files/phpword.svg',
  ],
];

// var_dump($_FILES); exit;

$upload_data = isset($_FILES['upload_data'])
  ? $_FILES['upload_data']
  : null;

if ($upload_data && $upload_data['tmp_name']) {
  $json = file_get_contents($upload_data['tmp_name']);
  $data = json_decode($json, true);
}

// var_dump($data); exit;

// merge PDF
if ($target_dir == 'merge_files') {
  // var_dump(get_defined_vars()); exit;

  //*/
  $file_ary = reArrayFiles($_FILES['upload_multifile']);
  // $file_ary = reArrangeArrayFiles($_FILES['upload_multifile']);

  if (count($file_ary) == 1 && $file_ary[0]['name'] == '') {
    die(nl2br("template-output Error:\nMessage: Please select at least one file."));
    exit;
  }

  $target_dir = 'uploaded_template/';
  $pdf_fils_list = [];

  foreach ($file_ary as $key => $file) {
    // print 'File Name: ' . $file['name'] . '<br />';
    // print 'File Name: ' . $file['tmp_name'] . '<br />';
    // print 'File Type: ' . $file['type'] . '<br />';
    // print 'File Size: ' . $file['size'] . '<br /><br />';

    if (pathinfo($file['name'])['extension'] == 'pdf') {
      $destination = $target_dir . $file['name'];

      if (move_uploaded_file($file['tmp_name'], $destination))
        array_push($pdf_fils_list, $destination);

    } else {
      $template = new PhpOfficeTemplate([
        'target_dir'  => $target_dir,
        'file_name'   => $file['name'],
        'file_post'   => $file['tmp_name'],
        // 'file_prefix' => $file_prefix,
        // 'sheet_name'  => 'template',
        'data'        => $data,

        'enable_empty_space'      => $emptyspace_option,
        'enable_office_convertor' => $offconverter_option,
      ]);

      $template->setPdfRenderer($pdf_option);

      $pdf_result = $template->output([
        'output_file_name'  => "pdf_file_$key",
        'method'            => 'server',
      ]);

      array_push($pdf_fils_list, $pdf_result);
    }
  }

  $pdf_fils_list = count($pdf_fils_list) > 0
    ? $pdf_fils_list
    : [
      'template_files/PhpOfficeTemplate_Excel.xlsx',
      'uploaded_template/My Output 1.pdf',
      'uploaded_template/My Output 2.pdf',
      'uploaded_template/My Output 3.pdf',
    ];
  // var_dump($pdf_fils_list); exit;

  //*/
  $output_as = 'browser';
  // $output_as = 'download';

  // file | download | string | browser(default)
  $output_as = $output_option == 'server'
    ? 'file'
    // ? 'string'
    : $output_option;

  PhpOfficeMerger::mergePDF($output_as, $pdf_fils_list, 'PDF_MERGE');

  foreach ($pdf_fils_list as $key => $file) {
    unlink($file);
  }
  //*/

  exit;
  //*/
}

if (!is_dir($target_dir) && $target_dir != 'stackoverflow') {
  $dir_created = mkdir($target_dir, 0775, true);

  if (!$dir_created) {
    var_dump("Failed to create directory : '$target_dir'.");
    var_dump("You may need to change your project owner/mod, or create the folder manually.");
    exit;
  }
}

if ($target_dir == 'stackoverflow') {
  $file = 'template_files/Template_Word_Portrait_Image.docx';
  $file_exist = file_exists($file);

  if ($file_exist) {
    $tset_save = new \PhpOffice\PhpWord\TemplateProcessor($file);
    $tset_save->saveAs('uploaded_template/test.docx');
    var_dump('done');
  }

  exit();

  $target_dir = 'uploaded_template/';
  $target_dir = 'template_files/';

  $file_name = 'test.docx';

  $file_name = 'test_stackoverflow.ods';
  $file_name = 'test_stackoverflow.csv';
  $file_name = 'test_stackoverflow.xlsx';

  $file_loc = $target_dir . $file_name;

  //*/
  // https://stackoverflow.com/questions/75358767/its-possible-to-access-sheet-in-import-class-of-laravel-excel-or-get-cell-styl
  $file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_loc);
  $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);
  $spreadsheet = $reader->load($file_loc);

  $worksheet = $spreadsheet->getActiveSheet();

  // var_dump($worksheet);
  // var_dump($worksheet->getCell('A1')->getStyle()->getFont()->getColor()->getARGB());

  var_dump($worksheet->getCell('A1')->getStyle()->exportArray());
  // var_dump($worksheet->getCell('A1')->getStyle()->exportArray()['fill']['startColor']);
  // var_dump($worksheet->getCell('B1')->getStyle()->exportArray()['fill']['startColor']);
  // var_dump($worksheet->getCell('C1')->getStyle()->exportArray()['fill']['startColor']);

  exit();
  //*/

  /*/
  // https://stackoverflow.com/questions/74404431/phpspreadsheet-load-csv-float-losing-precision
  $file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_loc);
  $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);
  $spreadsheet = $reader->load($file_loc);

  $worksheet = $spreadsheet->getActiveSheet();

  $worksheet->getCell('B1')->setValue(2);

  $value1 = $worksheet->getCell('A1')->getValue();
  $value2 = $worksheet->getCell('A2')->getValue();
  $value3 = $worksheet->getCell('A3')->getValue();
  $value4 = $worksheet->getCell('A4')->getValue();

  var_dump($value1);
  echo '<br>';
  var_dump($value2);
  echo '<br>';
  var_dump($value3);
  echo '<br>';
  var_dump($value4);
  echo '<br>';

  var_dump(4.02020325142409);
  echo '<br>';
  var_dump(3.90812005382548);
  echo '<br>';
  var_dump(4.55605765112764);
  echo '<br>';
  var_dump(4.4730378939229);
  echo '<br>';

  $worksheet->getCell('B1')->setValue($value1);
  $worksheet->getCell('B2')->setValue($value2);
  $worksheet->getCell('B3')->setValue($value3);
  $worksheet->getCell('B4')->setValue($value4);

  $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Csv');
  $writer->save($target_dir . 'result_' . $file_name);

  exit();
  //*/

  /*/
  // https://stackoverflow.com/questions/74611606/phpspreadsheet-write-10-000-record-is-too-slow
  $process_time = microtime(true);

  $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
  $spreadsheet = $reader->load($file_loc);
  $row_count = 10000;
  $col_count = 50;

  for ($r = 1; $r <= $row_count; $r++) {
    $rowArray = [];

    for ($c = 1; $c <= $col_count; $c++) {
      $rowArray[] = $r . ".Content " . $c;
    }

    $spreadsheet->getActiveSheet()->fromArray(
      $rowArray,
      NULL,
      'A' . $r
    );
  }

  $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
  $writer->save($target_dir . 'result_' . $file_name);

  unset($reader);
  unset($writer);
  $spreadsheet->disconnectWorksheets();
  unset($spreadsheet);

  $process_time = microtime(true) - $process_time;
  echo $process_time."\n";
  exit;
  // end of https://stackoverflow.com/questions/74611606/phpspreadsheet-write-10-000-record-is-too-slow
  //*/

  // $file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_loc);
  // $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);
  // $spreadsheet = $reader->load($file_loc);

  // $worksheet = $spreadsheet->getActiveSheet();

  // // $worksheet->getCell('A1')->setValue('John');
  // // $worksheet->getCell('A2')->setValue('Smith');

  // $row_number = 5;
  // $check_value = $worksheet->getCell('C' . $row_number)->getCalculatedValue();
  // // var_dump($check_value);

  // if ($check_value == '0')
  //   // $worksheet->setCellValue("C$row_number", 100);
  //   $worksheet->removeRow($row_number, 1);

  // // $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
  // $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Ods');
  // $writer->save($target_dir . 'result_' . $file_name);

  // // // die;
  // // // to download file
  // // header('Content-Type: application/vnd.ms-excel');
  // // header("Content-Length:" . filesize($filename));
  // // header("Content-Disposition: attachment;filename=$filename");
  // // header('Cache-Control: max-age=0');
  // // $writer->save('php://output');

  // exit();
}

// default excel sheet name
$sheetname = 'template';

// Init PhpOfficeTemplate
$config = [
  'target_dir'  => $target_dir,
  'file_name'   => $file_name,
  'file_prefix' => $file_prefix,
  'sheet_name'  => $sheetname,
  'data'        => $data,

  'enable_empty_space'      => $emptyspace_option,
  'enable_office_convertor' => $offconverter_option,
];

$upload_file = isset($_FILES['upload_file'])
  ? $_FILES['upload_file']
  : null;

if ($upload_file) {
  $config['file_name'] = $upload_file['name'];
  $config['file_post'] = $upload_file['tmp_name'];
}

// if ($file_name == 'test.docx')
//   $config['enable_empty_space'] = true;

//*/
// run template
$template = new PhpOfficeTemplate($config);

// set orientation : portriat | landscape
$template->setOrientation($orientation_option);

// set PDF renderer : mpdf | tcpdf | dompdf
// default for spreadsheet | success render img on PhpWord
$template->setPdfRenderer($pdf_option);

if ($file_name == 'test.docx') {
  // var_dump($template->getPhpOfficeObject()->getPhpWord()); exit;

  /*/
  $template
    ->getPhpOfficeObject()
    ->getPhpWord()
    ->setValue('customer_name', 'my name');
  //*/

  /*/
  $template
    ->getPhpOfficeObject()
    ->getPhpWord()
    ->cloneBlock('block_name', 3, true, true);
  //*/

  /*/
  $replacements = array(
    array('customer_name' => 'Batman', 'customer_address' => 'Gotham City'),
    array('customer_name' => 'Superman', 'customer_address' => 'Metropolis'),
  );

  $template
    ->getPhpOfficeObject()
    ->getPhpWord()
    ->cloneBlock('block_name', 0, true, false, $replacements);
  //*/

  /*/
  // https://github.com/PHPOffice/PHPWord/issues/838
  // https://github.com/PHPOffice/PHPWord/issues/268
  $content = '${block_var} ${var1} ${/block_var}';
  $content = '${block_var}\n${var2}\n${block_var}';
  $replace = $content;

  $template
    ->getPhpOfficeObject()
    ->getPhpWord()
    ->setValue('line1', $replace);
  //*/

  //*/
  // https://stackoverflow.com/questions/56620136/include-a-line-break-in-a-value
  $template
    ->getPhpOfficeObject()
    ->getPhpWord()
    ->setValues([
      'line1' => '${block_var}',
      'line2' => '${var1}',
      'line3' => '${/block_var}',
    ]);
  //*/

  $template->output([
    'output_file_name' => 'test2',
    'method'  => 'server',  // browser | download | server | default:''
    'type'    => 'docx',    // xlsx | xls | ods | docx | doc | odt | default:pdf
  ]);

  $config['file_name'] = 'test2.docx';
  $template = new PhpOfficeTemplate($config);

  $replacements = array(
    array('var1' => 'value1'),
    array('var1' => 'value2'),
  );

  $template
    ->getPhpOfficeObject()
    ->getPhpWord()
    ->cloneBlock('block_var', 0, true, false, $replacements);

  // exit;
  //*/
}

$result = $template->output([
  'output_file_name'  => $filename_option,  // 'My Output',
  'method'            => $output_option,    // browser | download | server | default:''
  'type'              => $download_option,  // xlsx | xls | ods | docx | doc | odt | default:pdf
  // 'unlink'            => true,              // true to remove uploaded template, default:false
]);

if ($output_option == 'server') {
  echo "File save successfully in $result";
}
/*/
// merge PDF
$output_as = 'browser';
$output_as = 'download';

PhpOfficeTemplate::mergeFiles($output_as, [
  'template_files/PhpOfficeTemplate_Excel.xlsx',
  'uploaded_template/My Output.pdf',
  'uploaded_template/My Output 2.pdf',
]);
//*/

exit;

?>