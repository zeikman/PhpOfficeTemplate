<?php

// https://stackoverflow.com/questions/33011956/escape-special-character-from-string-in-php
// https://stackoverflow.com/questions/14114411/remove-all-special-characters-from-a-string

function clean($string) {
  $string = str_replace(' ', '-', $string); // Replaces all spaces with hyphens.

  return preg_replace('/[^A-Za-z0-9\-]/', '', $string); // Removes special chars.
}

$dir = __DIR__ . '/demo/template_files';

// $file = clean($_GET['file']);
$filename = $_GET['file'];
$file = $dir . "/$filename";

// echo $file;

// exit;

if (file_exists($file)) {
  header('Content-Description: File Transfer');
  header('Content-Type: application/octet-stream');
  header('Content-Disposition: attachment; filename="'.basename($file).'"');
  header('Content-Transfer-Encoding: binary');
  header('Expires: 0');
  header('Cache-Control: must-revalidate');
  header('Pragma: public');
  header('Content-Length: ' . filesize($file));

  ob_clean();

  flush();

  readfile($file);

  exit;
}

?>