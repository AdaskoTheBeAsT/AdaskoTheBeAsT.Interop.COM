@echo off
set TLBIMP="C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe"

REM One MSIL-neutral interop assembly works for both x86 and x64 host processes.
%TLBIMP% .\x86\Debug\NativeCOM.dll /out:.\x86\Debug\Interop.NativeCOM.dll /namespace:NativeCOM /machine:Agnostic
%TLBIMP% .\x64\Debug\NativeCOM.dll /out:.\x64\Debug\Interop.NativeCOM.dll /namespace:NativeCOM /machine:Agnostic