<?php
include 'Broadstreet.php';

$network_id    = '5210'; // Something you have access to
$advertiser_id = '88651'; // And advertiser under that network
$access_token  = 'b5f2a01cca1621a6e2b6c4a23523f20ae2c3e7ca8daa400ff1ad6cbc9fbf38c9'; // see https://my.broadstreetads.com/access-token

try
{
    $client = new Broadstreet($access_token);
    
    /* Create an ad */
    $ad = $client->createAdvertisement($network_id, $advertiser_id, 'New Ad!', 'text', array (
        'default_text' => 'This is the message'
    ));
    
    /* Print ad code */
    echo $ad->html;
}
catch(Exception $ex)
{
    echo "Whoops, there was a problem connecting to Broadstreet:" . $ex->__toString();
}