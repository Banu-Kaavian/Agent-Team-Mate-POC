# Quick test script to join a Teams meeting
# Usage: .\test-join.ps1

$baseUrl = "http://localhost:5000"

# 1. Health check
Write-Host "`n=== HEALTH CHECK ===" -ForegroundColor Cyan
$health = Invoke-RestMethod -Uri "$baseUrl/" -Method Get
Write-Host "Status: $($health.status)" -ForegroundColor Green
Write-Host "Time: $($health.time)"

# 2. Join meeting
Write-Host "`n=== JOINING MEETING ===" -ForegroundColor Cyan
$body = @{
    meetingId = "273 866 171 404 758"
    passcode  = "5GG7JL7j"
} | ConvertTo-Json

try {
    $result = Invoke-RestMethod -Uri "$baseUrl/api/join" `
        -Method Post `
        -ContentType "application/json" `
        -Body $body

    Write-Host "Call ID : $($result.callId)" -ForegroundColor Green
    Write-Host "State   : $($result.state)" -ForegroundColor Green
    Write-Host "Media   : $($result.media)" -ForegroundColor Green
}
catch {
    Write-Host "JOIN FAILED: $($_.Exception.Message)" -ForegroundColor Red
    $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
    Write-Host $reader.ReadToEnd() -ForegroundColor Red
}

Write-Host "`n=== BOT IS NOW IN THE MEETING ===" -ForegroundColor Yellow
Write-Host "1. Open your Teams meeting (ID: 273 866 171 404 758)" -ForegroundColor White
Write-Host "2. You should see 'Agent Team Mate' join" -ForegroundColor White
Write-Host "3. Say: 'Agent Nova, what is today's date?'" -ForegroundColor White
Write-Host "4. Wait 5-10 seconds for the response" -ForegroundColor White
Write-Host ""
Write-Host "Watch the bot terminal for:" -ForegroundColor White
Write-Host "  - 'SPEECH RECOGNIZED'" -ForegroundColor Gray
Write-Host "  - 'AGENT NOVA INVOCATION DETECTED'" -ForegroundColor Gray
Write-Host "  - 'AGENT TEAM MATE RESPONSE'" -ForegroundColor Gray
Write-Host "  - 'PLAYING AI RESPONSE IN TEAMS'" -ForegroundColor Gray
