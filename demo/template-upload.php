<?php

// Error reporting
$debug = TRUE;
// $debug = FALSE;
error_reporting(E_ALL);
ini_set('display_errors', $debug);
ini_set('display_startup_errors', $debug);

$target_dir   = $_POST['targetdir'];
$upload_file  = $_FILES['upload_file'];
$filename     = $upload_file['name'];

if (isset($upload_file)) {

  $target_file = $filename;
  $destination = $target_dir . $target_file;

  if (move_uploaded_file($upload_file['tmp_name'], $destination)) {
    $message = $filename . ' uploaded !';

    echo json_encode([
      'status'      => 200,
      'readyState'  => 4,
      'message'     => $message
    ]);

  } else {
    $message = $destination . ' upload failed ...';

    print_r($_FILES);

    echo json_encode([
      'status'  => 0,
      'message' => $message
    ]);
  }

  exit;
}

?>