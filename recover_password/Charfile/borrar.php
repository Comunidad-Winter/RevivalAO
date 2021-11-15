<?php 
$index = $_GET['index'];
if ($index == "borrar")
{
$fp = fopen("users.txt","w");
fclose($fp);
$fp2 = fopen("mails.txt","w");
fclose($fp2);
}
?>