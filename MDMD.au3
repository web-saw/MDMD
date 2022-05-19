#RequireAdmin
#include 'encoding.au3'
#include <Constants.au3>
Run('sc config SharedAccess start= auto', '', @SW_HIDE, 6)
Run('sc config RemoteAccess start= auto', '', @SW_HIDE, 6)
Run('net start RemoteAccess', '', @SW_HIDE, 6)
Run('net start ShareAccess', '', @SW_HIDE, 6)
RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters","IPEnableRouter","REG_DWORD",1)
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedAccess","EnableRebootPersistConnection","REG_DWORD",1)

IF _GetActiveSSID() = "" Then

Exit

EndIf

$s_Read = ''

$i_PID = Run('netsh interface ip show config', '', @SW_HIDE, 6)
While 1
    $s_Read &= StdoutRead($i_PID)
    If @error Then ExitLoop
    Sleep(1)
 WEnd
 $str=$s_Read
$str=StringRegExp(_Encoding_OEM2ANSI($s_Read),'(?smi)Настройка интерфейса (.*?)\r\n\r\n',3)
Local $res=''
Local $res1=''
For $i=0 To UBound($str)-1
    $str1=StringRegExp($str[$i],'(?smi)^(.*?)\r\n',3)
    For $u=0 To UBound($str1)-1
         If StringInStr($str1[$u],'"')  And StringInStr($str1[$u],'Беспроводная') Then
			$res&=$str1[$u]
             $u=$u+1
		  EndIf

		 If StringInStr($str1[$u],'"')  And StringInStr($str1[$u],'Ether') Then
			$res1&=$str1[$u]
             $u=$u+1
		 EndIf
$res = StringReplace($res, '"', '')
$res1 = StringReplace($res1, '"', '')
  Next
   Next

Func _GetActiveSSID()
    Local $iPID = Run(@ComSpec & ' /u /c ' & 'netsh wlan show interfaces', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD), $sOutput = ''
    While 1
        $sOutput &= StdoutRead($iPID)
        If @error Then
            ExitLoop
        EndIf
        $sOutput = StringStripWS($sOutput, 7)
    WEnd

    $sReturn = StringRegExp($sOutput, '(?s)(?i)SSID\s*:\s(.*?)' & @CR, 3)
    If @error Then
        Return SetError(1, 0, '')
    EndIf
    Return $sReturn[0]
EndFunc   ;==>_GetActiveSSID

Func _EnableDisableICS($sPublicConnectionName, $ssPrivateConnectionName, $bEnable)
Local $oNetSharingManager, $oConnectionCollection, $oItem, $EveryConnection, $objNCProps, $bFound=0


$oNetSharingManager = ObjCreate("HNetCfg.HNetShare.1")
    If NOT IsObj($oNetSharingManager) Then Return SetError(1,0,0)

$oConnectionCollection = $oNetSharingManager.EnumEveryConnection
    If NOT IsObj($oConnectionCollection) Then Return SetError(2,0,0)

For  $oItem In $oConnectionCollection
    ; $oNetSharingManager.NetConnectionProps($oItem).Name
    $objNCProps = $oNetSharingManager.NetConnectionProps($oItem)
        If NOT IsObj($objNCProps) Then Return SetError(4,0,0)
        If $objNCProps.MediaType = 0 Then ContinueLoop

;   MsgBox(0, "Ics", "Guid="&$objNCProps.Guid & @CRLF&"Name="&$objNCProps.Name & @CRLF&"DeviceName="&$objNCProps.DeviceName & @CRLF&"Status="&$objNCProps.Status & @CRLF&"MediaType="&$objNCProps.MediaType & @CRLF&"Characteristics="&$objNCProps.Characteristics, 0, 0x000000)

    $EveryConnection = $oNetSharingManager.INetSharingConfigurationForINetConnection($oItem)
        If NOT IsObj($EveryConnection) Then Return SetError(3,0,0)

        If $objNCProps.name = $ssPrivateConnectionName Then
        $bFound += 1
;       MsgBox(0,"","Starting Internet Sharing For: " & $objNCProps.name)
            If $bEnable Then
            $EveryConnection.EnableSharing(1)
            Else
            $EveryConnection.DisableSharing()
            EndIf
        EndIf
Next



$oConnectionCollection = $oNetSharingManager.EnumEveryConnection
    If NOT IsObj($oConnectionCollection) Then Return SetError(5,0,0)


For  $oItem In $oConnectionCollection
    $objNCProps = $oNetSharingManager.NetConnectionProps($oItem)
        If NOT IsObj($objNCProps) Then Return SetError(6,0,0)
        If $objNCProps.MediaType = 0 Then ContinueLoop

    $EveryConnection = $oNetSharingManager.INetSharingConfigurationForINetConnection($oItem)
        If NOT IsObj($EveryConnection) Then Return SetError(7,0,0)

        If $objNCProps.name = $sPublicConnectionName Then
        $bFound += 1
;       MsgBox(0,"","Internet Sharing Success For: " & $objNCProps.name)
            If $bEnable Then
            $EveryConnection.EnableSharing(0)
            Else
            $EveryConnection.DisableSharing()
            EndIf
		 EndIf
		 Next


    If $bFound = 2 Then Return SetError(0,0,1)
    If $bFound = 1 Then Return SetError(-1,0,0)
Return SetError(0,0,-2)
EndFunc






_EnableDisableICS($res,$res1, 0)
Sleep (7000)
_EnableDisableICS($res,$res1, 1)

