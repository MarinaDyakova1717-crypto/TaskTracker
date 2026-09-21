# Требуемые команды из методички — вывод в консоль для анализа
$base = "C:\Practice\TaskTracker_AuditProject"
Set-Location $base

Write-Host "=== Get-Location ==="
Get-Location

Write-Host "`n=== Docs (all .md, .txt, exclude .git) ==="
Get-ChildItem -Recurse -File -Include *.md,*.txt | Where-Object { $_.FullName -notmatch '\\.git\\' } | Select-Object -Property FullName

Write-Host "`n=== TaskTrackerApp dir ==="
Get-ChildItem -Recurse -Directory | Where-Object { $_.Name -eq 'TaskTrackerApp' } | Select-Object -Property FullName

Write-Host "`n=== .NET mentions in docs ==="
Get-ChildItem -Recurse -File -Include *.md,*.txt | Select-String -Pattern '\.NET' | Select-Object -Property Path, Line

Write-Host "`n=== TaskTracker mentions (md,txt,cs,csproj) ==="
Get-ChildItem -Recurse -File -Include *.md,*.txt,*.cs,*.csproj | Select-String -Pattern 'TaskTracker' | Select-Object -Property Path, Line

Write-Host "`n=== Installed Files (tree-like via dir /s) ==="
Get-ChildItem -Recurse -File | Select-Object -Property FullName
