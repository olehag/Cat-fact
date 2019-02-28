Add-Type -AssemblyName System.Speech
$SpeechSynth = New-Object System.Speech.Synthesis.SpeechSynthesizer
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$CatFact = (ConvertFrom-Json (Invoke-WebRequest -Uri 'https://catfact.ninja/fact')).fact
Write-Host "Did you know? ...."
$SpeechSynth.Speak("did you know?")
Write-Host ""
Write-Host ""
Write-Host $CatFact
$SpeechSynth.Speak($CatFact)