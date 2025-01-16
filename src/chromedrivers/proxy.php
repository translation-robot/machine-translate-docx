<?php

// Check if URL parameter is provided
if (isset($_GET['url'])) {
    $target_url = $_GET['url'];

    // Validate URL format and domain
    if (filter_var($target_url, FILTER_VALIDATE_URL) &&
        strpos($target_url, 'https://storage.googleapis.com/') === 0) {

        // Extract filename from the URL
        $filename = basename($target_url);

        // Validate filename against regex pattern
        if (preg_match('/^chromedriver-.*\.zip$/', $filename)) {
            // Initialize cURL session
            $ch = curl_init();

            // Set the URL to be retrieved
            curl_setopt($ch, CURLOPT_URL, $target_url);

            // Other options you may want to set
            curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true); // Follow redirects
            curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false); // Disable SSL verification (not recommended for production)

            // Set option to return the transfer as a string
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);

            // Execute cURL session
            $response = curl_exec($ch);

            // Check for errors
            if (curl_errno($ch)) {
                header("HTTP/1.1 500 Internal Server Error");
                echo 'Error: ' . curl_error($ch);
            } else {
                // Set headers for binary download with dynamic filename
                header('Content-Type: application/octet-stream');
                header('Content-Disposition: attachment; filename="' . $filename . '"');

                // Output the content to return to the client
                echo $response;
            }

            // Close cURL session
            curl_close($ch);
            exit(); // Stop script execution after downloading
        } else {
            header("HTTP/1.1 403 Forbidden");
            echo "Filename does not match the required pattern.";
        }
    } else {
        header("HTTP/1.1 403 Forbidden");
        echo "Invalid or unauthorized URL.";
    }
} else {
    header("HTTP/1.1 400 Bad Request");
    echo "URL parameter is missing.";
}

?>
