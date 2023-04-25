<?php

// https://stackoverflow.com/questions/720751/how-to-read-a-list-of-files-from-a-folder-using-php
// var_dump(__DIR__);
// var_dump(getcwd());

$dir = __DIR__ . '/demo/template_files';

$filelist = [];

/*/
if (is_dir($dir) && $dh = opendir($dir)) {
  while (($file = readdir($dh)) !== false) {
    array_push($filelist, $file);
  }

  closedir($dh);
}
/*/
if (is_dir($dir)) {
  // https://www.php.net/manual/en/function.str-starts-with.php
  if (!function_exists('str_starts_with')) {
    function str_starts_with($str, $start) {
      return (@substr_compare($str, $start, 0, strlen($start))==0);
    }
  }

  $files = scandir($dir, 0);

  for ($i = 2; $i < count($files); $i++) {
    if (str_starts_with($files[$i], 'Template') || $files[$i] == 'data.json')
      array_push($filelist, $files[$i]);
  }
}
//*/

echo json_encode($filelist);
exit;

?>