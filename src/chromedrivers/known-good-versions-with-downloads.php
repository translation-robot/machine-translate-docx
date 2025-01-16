<?php

// URL of the JSON file
$json_url = 'https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json';

// Download the JSON file
$json_content = file_get_contents($json_url);

// Check if download was successful
if ($json_content === false) {
    die('Failed to download JSON file.');
}

// Get the domain from the request headers
$domain = isset($_SERVER['HTTP_HOST']) ? $_SERVER['HTTP_HOST'] : '';

// Modify the JSON content with the dynamic domain
$modified_json = str_replace('https://storage.googleapis.com/', "https://$domain/proxy.php?url=https://storage.googleapis.com/", $json_content);

// Set header for JSON response
header('Content-Type: application/json');

// Output the modified JSON
echo $modified_json;

?>
