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

$output_option    = $_POST['output_option'];
$download_option  = $_POST['download_option'];

$file_name        = $_POST['filename'];   // file name
$file_prefix      = $_POST['fileprefix']; // file name when download
$target_dir       = $_POST['targetdir'];  // directory to store template file
$sheetname        = 'template';           // default excel sheet name

// prepare data
$data = [
  '${cp_company_name}'  => 'My Company',
  '${m_ext_sisco_id}'   => 'My Sisko ID',
  '${wk_nm}'            => 'Worker Name',
  '${wk_contact}'       => '0123456789',
];

// Init PhpOfficeTemplate
$config = [
  'target_dir'          => $target_dir,
  'file_name'           => $file_name,
  'file_prefix'         => $file_prefix,
  'sheet_name'          => $sheetname,
  'data'                => $data,
  'enable_empty_space'  => true,
];

$upload_file = $_FILES['upload_file'];

if (isset($upload_file)) {
  $config['file_name'] = $upload_file['name'];
  $config['file_post'] = $upload_file['tmp_name'];
}

// var_dump($_FILES);
// var_dump($_POST);

$template = new PhpOfficeTemplate($config);

// default for spreadsheet | success render img on PhpWord
// $template->setPdfRenderer('mpdf');
// $template->setPdfRenderer('tcpdf');
// $template->setPdfRenderer('dompdf');

// set orientation
// TODO: turn into settings
// $template->setOrientation('portriat');
// $template->setOrientation('landscape');

$template->output([
  'method'  => $output_option,    // browser | download | server | default:null
  'type'    => $download_option,  // xlsx | xls | ods | docx | doc | odt | default:pdf
  // 'unlink'  => true,              // true to remove uploaded template, default:false
]);

// echo json_encode([
//   'success' => 1
// ]);

?>