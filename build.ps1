Install-Module ps2exe -Force -Scope CurrentUser

Invoke-ps2exe `
    -inputFile "RealtimeSyncStarter.ps1" `
    -outputFile "RealtimeSyncStarter.exe" `
    -noConsole `
    -iconFile ""
