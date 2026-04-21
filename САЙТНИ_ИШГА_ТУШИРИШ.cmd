@echo off
chcp 65001 > nul
title Навоий ер платформаси
echo.
echo ========================================
echo   Навоий вилояти - Ер платформаси
echo ========================================
echo.
echo Сервер ишга туширилмоқда...
echo.

start "" "http://localhost:8765/navoi_map_platform.html"

powershell -NoProfile -NoLogo -ExecutionPolicy Bypass -Command ^
  "$listener = New-Object System.Net.HttpListener; ^
   $listener.Prefixes.Add('http://localhost:8765/'); ^
   $listener.Start(); ^
   Write-Host 'Сайт очиқ: http://localhost:8765/navoi_map_platform.html'; ^
   Write-Host 'Тўхтатиш учун: Ctrl+C'; ^
   $root = (Get-Item -LiteralPath '%~dp0').FullName; ^
   $mime = @{'.html'='text/html; charset=utf-8';'.htm'='text/html; charset=utf-8';'.js'='application/javascript; charset=utf-8';'.css'='text/css; charset=utf-8';'.json'='application/json; charset=utf-8';'.png'='image/png';'.jpg'='image/jpeg';'.svg'='image/svg+xml';'.ico'='image/x-icon';'.woff'='font/woff';'.woff2'='font/woff2'}; ^
   while ($listener.IsListening) { ^
     try { ^
       $ctx = $listener.GetContext(); ^
       $p = [System.Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath); ^
       if ($p -eq '/' -or $p -eq '') { $p = '/navoi_map_platform.html' }; ^
       $full = Join-Path $root ($p.TrimStart('/').Replace('/', '\')); ^
       if (Test-Path -LiteralPath $full -PathType Leaf) { ^
         $ext = [System.IO.Path]::GetExtension($full).ToLower(); ^
         $ct = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { 'application/octet-stream' }; ^
         $b = [System.IO.File]::ReadAllBytes($full); ^
         $ctx.Response.ContentType = $ct; ^
         $ctx.Response.ContentLength64 = $b.Length; ^
         $ctx.Response.Headers.Add('Cache-Control','no-store'); ^
         $ctx.Response.OutputStream.Write($b, 0, $b.Length); ^
       } else { ^
         $ctx.Response.StatusCode = 404; ^
       }; ^
       $ctx.Response.Close(); ^
     } catch {} ^
   }"

pause
