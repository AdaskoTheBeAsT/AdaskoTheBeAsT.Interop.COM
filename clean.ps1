Get-ChildItem .\ -Include bin,obj -Recurse | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -Recurse }
