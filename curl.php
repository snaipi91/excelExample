<?php
/**
 * Created by PhpStorm.
 * User: Andrey
 * Date: 16.12.2015
 * Time: 2:31
 */
header("Content-Type:application/vnd.ms-excel");
header("Content-Disposition:attachment;filename='test.xls'");

$token = "TOKEN";

// делаем запрос к API Football
$ch = curl_init("http://football-api.com/api/?Action=fixtures&OutputType=JSON&APIKey=" . $token . "&from_date=19.07.2015&to_date=15.12.2015");

curl_setopt($ch, CURLOPT_HEADER, 0);
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);

$obj = curl_exec($ch);

curl_close($ch);

$obj = json_decode($obj, true);
$player = $obj['matches'];

require('excel.php');

?>

<pre>
    <? echo print_r($player); ?>
</pre>
