<?php
header('Content-Type: text/plain; charset=UTF-8');
date_default_timezone_set('Europe/Warsaw');
echo "PROF-INSTAL DIAG ".date('Y-m-d H:i:s')."\n\n";

$keys = [
 'REQUEST_URI','REQUEST_METHOD','HTTP_HOST','SERVER_NAME','SERVER_ADDR','SERVER_PORT','REQUEST_SCHEME',
 'HTTPS','HTTP_X_FORWARDED_PROTO','HTTP_X_FORWARDED_SSL','HTTP_X_FORWARDED_FOR',
 'HTTP_X_REAL_IP','HTTP_CF_VISITOR','HTTP_X_FORWARDED_HOST','HTTP_X_FORWARDED_PORT'
];
foreach ($keys as $k) {
  $v = isset($_SERVER[$k]) ? $_SERVER[$k] : '(not set)';
  echo str_pad($k, 24).'= '.$v."\n";
}

if (function_exists('getallheaders')) {
  echo "\n--- ALL HEADERS ---\n";
  foreach (getallheaders() as $k=>$v) echo $k.": ".$v."\n";
}
