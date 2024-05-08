<#
Pealkiri: 32-bit PowerShell sessiooni vahetamine 64-bit vastu
Autor: Arjan Vroege
Kuup�ev: 07.02.2018
Saadavus: https://web.archive.org/web/20230628202424/https://www.vroege.biz/?p=3970
#>
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    write-warning "WARNING: 32bit PowerShell, starting x64 PowerShell session and continue"
    if ($myInvocation.Line) {
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
    }else{
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" $args
    }
exit $lastexitcode
}



<#
Pealkiri: Remote Desktop Services API-st funktsiooni importimine
Autor: boxdog
Kuup�ev: 25.01.2019
Saadavus: https://stackoverflow.com/questions/54357542/msgbox-in-powershell-script-run-from-task-scheduler-not-working/54362903#54362903
#>
$typeDefinition = @"
using System;
using System.Runtime.InteropServices;

public class WTSMessage {
    [DllImport("wtsapi32.dll", SetLastError = true)]
    public static extern bool WTSSendMessage(
        IntPtr hServer,
        [MarshalAs(UnmanagedType.I4)] int SessionId,
        String pTitle,
        [MarshalAs(UnmanagedType.U4)] int TitleLength,
        String pMessage,
        [MarshalAs(UnmanagedType.U4)] int MessageLength,
        [MarshalAs(UnmanagedType.U4)] int Style,
        [MarshalAs(UnmanagedType.U4)] int Timeout,
        [MarshalAs(UnmanagedType.U4)] out int pResponse,
        bool bWait
     );

     static int response = 0;

     public static int SendMessage(int SessionID, String Title, String Message, int Timeout, int MessageBoxType) {
        WTSSendMessage(IntPtr.Zero, SessionID, Title, Title.Length, Message, Message.Length, MessageBoxType, Timeout, out response, true);

        return response;
     }

}
"@
Add-Type -TypeDefinition $typeDefinition



<#
Pealkiri: Aktiivse kasutaja ID p�rimine
Autor: VVG
Kuup�ev: 22.05.2019
Saadavus: https://stackoverflow.com/questions/56239473/powershell-interaction-between-system-user-and-logged-on-users
#>
$RawOuput = (quser) -replace '\s{2,}', ',' | ConvertFrom-Csv
$sessionID = $null

Foreach ($session in $RawOuput) {  
	if($session.STATE -eq "Active"){    
		$sessionID = $session.ID
	}                               
}



<#
Pealkiri: Arvuti v�ljal�limine, kasutajapoolse peatamisv�imalusega
Autor: Marti Orav
Kuup�ev: 27.04.2024
#>
# P�hjus sulgemiseks
$reason = "Energia kokkuhoid"

# Kui pole aktiivset kasutajat, siis kohene v�ljal�limine.
if($sessionID -eq $null) {
    shutdown /s /f /t 0 /c $reason /d p:0:0
}else{
    # Kui on aktiivne kasutaja, siis ajalise viitega v�ljal�limine
    shutdown /s /f /t 600 /c $reason /d p:0:0
    
    # Kasutajalt sisendi k�simine, kas l�limine peatada
    $vastus = [WTSMessage]::SendMessage($sessionID, "Sulgumisteade", "Arvuti sulgub 10 minuti jooksul. Kas soovid sulgumist peatada?", 60, 52)

    # Jaatava vastuse korral l�limine peatatakse
    if($vastus -eq 6){
        shutdown /a
    }
}