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
include_once '../lib/number2word.php';

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
  : 'uploaded_template';

// var_dump($_FILES);
// exit;

/*/
$file_ary = reArrayFiles($_FILES['upload_multifile']);
$file_ary = reArrangeArrayFiles($_FILES['upload_multifile']);

// foreach ($file_ary as $file) {
//   print 'File Name: ' . $file['name'];
//   print 'File Type: ' . $file['type'];
//   print 'File Size: ' . $file['size'];
// }

var_dump($file_ary);
exit;
//*/

if ($target_dir == 'merge_files') {
  // merge PDF
  $output_as = 'browser';
  // $output_as = 'download';

  PhpOfficeMerger::mergePDF($output_as, [
    'template_files/PhpOfficeTemplate_Excel.xlsx',
    'uploaded_template/My Output.pdf',
    'uploaded_template/My Output 2.pdf',
    'uploaded_template/My Output 3.pdf',
  ]);

  exit;
}

if (!is_dir($target_dir)) {
  $dir_created = mkdir($target_dir, 0775, true);

  if (!$dir_created) {
    var_dump("Failed to create directory : '$target_dir'.");
    var_dump("You may need to change your project owner/mod, or create the folder manually.");
    exit;
  }
}

// default excel sheet name
$sheetname = 'template';

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
  '${my_addr}'        => 'Testing Address',
  '${my_passport}'    => 'Testing Passport',
  '${my_marital}'     => 'Single',
  '${my_job}'         => 'Testing Job',
  '${my_gender_code}' => 'M',
  '${my_gender}'      => 'Male',
];
// var_dump($data);

$upload_data = isset($_FILES['upload_data'])
  ? $_FILES['upload_data']
  : null;

if ($upload_data && $upload_data['tmp_name']) {
  $json = file_get_contents($upload_data['tmp_name']);
  $data = json_decode($json, true);
}

// var_dump($data);
// exit;

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

//*/
// run template
$template = new PhpOfficeTemplate($config);

// set orientation : portriat | landscape
$template->setOrientation($orientation_option);

// set PDF renderer : mpdf | tcpdf | dompdf
// default for spreadsheet | success render img on PhpWord
$template->setPdfRenderer($pdf_option);

$result = $template->output([
  'output_file_name' => 'My Output',
  'method'  => $output_option,    // browser | download | server | default:''
  'type'    => $download_option,  // xlsx | xls | ods | docx | doc | odt | default:pdf
  // 'unlink'  => true,              // true to remove uploaded template, default:false
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