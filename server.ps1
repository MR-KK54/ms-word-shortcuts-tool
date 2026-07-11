# ====================================================================
# WordKeys Hub: Pure PowerShell Web Server
# Serves the static site locally on http://localhost:8080/ without Node/Python
# ====================================================================

$port = 8080
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

Write-Output "=========================================="
Write-Output "WordKeys Hub Server is running!"
Write-Output "URL: http://localhost:$port/"
Write-Output "=========================================="
Write-Output "Press Ctrl+C in this terminal to stop."

# Stop listening if listener is stopped or interrupted
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $urlPath = $request.Url.LocalPath
        if ($urlPath -eq "/") {
            $urlPath = "/index.html"
        }

        # Resolve local file path
        $filePath = Join-Path $PSScriptRoot $urlPath.TrimStart('/')

        if (Test-Path $filePath -PathType Leaf) {
            $content = [System.IO.File]::ReadAllBytes($filePath)
            
            # Match mime type
            $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
            $contentType = "text/plain"
            if ($extension -eq ".html") { $contentType = "text/html; charset=utf-8" }
            elseif ($extension -eq ".css") { $contentType = "text/css" }
            elseif ($extension -eq ".js") { $contentType = "application/javascript" }
            elseif ($extension -eq ".json") { $contentType = "application/json" }

            $response.ContentType = $contentType
            $response.ContentLength64 = $content.Length
            $response.OutputStream.Write($content, 0, $content.Length)
        } else {
            $response.StatusCode = 404
            $msg = [System.Text.Encoding]::UTF8.GetBytes("404 File Not Found")
            $response.ContentLength64 = $msg.Length
            $response.OutputStream.Write($msg, 0, $msg.Length)
        }
        $response.Close()
    }
}
catch {
    # Clean exit on Ctrl+C or port conflicts
}
finally {
    $listener.Stop()
    $listener.Close()
}
