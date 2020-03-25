#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Images\hba.ico
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;*****************************************
;Get-HBAWin.au3 by jacky
;ISN AutoIt Studio 版本 v. 1.10
;*****************************************

;Set Res_HiDpi, if not compiled
If Not @Compiled Then DllCall("User32.dll", "bool", "SetProcessDPIAware")

#Region AutoIt Stuff
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", 802)
Opt("GUICloseOnESC", 1) ;Can close every GUI with ESC
#include <Misc.au3>
_Singleton("Get-HBAWin.exe")
#EndRegion AutoIt Stuff


#Region AutoIt Includes
; AutoIt Include
#include <Array.au3>

; Customer Include
#include "Forms\Get-HBAWin.isf"

#EndRegion AutoIt Includes

Opt("MustDeclareVars", 1) ; 变量必须先定义，此语句放在包含isf语句后面。

#Region Main Code
; 定义变量

; 控件赋值
GUICtrlSetFont($edtMemo,10,400,0,"Consolas")
GUICtrlSetStyle($edtMemo, $ES_READONLY)
GUICtrlSetData($btnClose, "Close")
GUISetIcon(@SystemDir& "\shell32.dll", 177)


; 显示窗体
GUISetState(@SW_SHOW)

Local $aHBA =  Get_HBAWin()
Local $iCol = UBound($aHBA, 2)
If $iCol = 1 Then 
    MemoWrite("Not Found Fibre Channel Host Bus Adapters" )
Else
    While $iCol > 1
        MemoWrite(_ArrayToString($aHBA, ": ", -1, -1, @CRLF, 0, 1))
        MemoWrite(@CR)
        _ArrayColDelete($aHBA, 1)
        $iCol = UBound($aHBA, 2)
    WEnd 
EndIf

#EndRegion Main Code


#Region OnEvent
GUISetOnEvent($GUI_EVENT_CLOSE, _Exit)
GUICtrlSetOnEvent($btnClose, _Exit)
GUICtrlSetOnEvent($txtWebsite, Website)
#EndRegion OnEvent



#Region While Loop
While 1

	Sleep(250) ;Idle
WEnd
#EndRegion While Loop

#Region Customer Functions
; #FUNCTION# ====================================================================================================================
; Name ..........: _Exit
; Description ...: 使用此函数退出脚本
; Syntax ........: _Exit()
; Parameters ....: None
; Return values .: None
; Author ........: Jacky
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Exit()
	Exit
EndFunc

Func Get_HBAWin($host = ".")
    Local $wmiObject
    Local $sComputerName = $host
    Local $sNameSpace = "ROOT\WMI"
    Local $sClass =  "MSFC_FCAdapterHBAAttributes"
    Local Enum $Model, $ModelDescription, $DriverVersion, $NodeWWN, $FirmwareVersion, $DriverName, $Active
    Dim $aHBA[7][1]
    $aHBA[$Model][0] = "Model"
    $aHBA[$ModelDescription][0] = "ModelDescription"
    $aHBA[$DriverVersion][0] = "DriverVersion"
    $aHBA[$NodeWWN][0] = "NodeWWN"
    $aHBA[$FirmwareVersion][0] = "FirmwareVersion"
    $aHBA[$DriverName][0] = "DriverName"
    $aHBA[$Active][0] = "Active"
    $wmiObject = ObjGet("WINMGMTS:\\"& $sComputerName & "\" & $sNameSpace )
    If IsObj($wmiObject) And (Not @error) Then 
        ; Get instances of MSFC_FCAdapterHBAAttributes 
        Local $Instances = $wmiObject.InstancesOf($sClass)
        ; Enumerate instances
        Local $i = 1
        For $Instance In $Instances
            ReDim $aHBA[7][$i + 1]
            $aHBA[$Model][$i] =  $Instance.Model
            $aHBA[$ModelDescription][$i] =  $Instance.ModelDescription
            $aHBA[$DriverVersion][$i] =  $Instance.DriverVersion
            Local $sNodeWWN, $sWWN = Null
            For $sNodeWWN In $Instance.NodeWWN
                If $sWWN = Null Then 
                    $sWWN = Hex($sNodeWWN, 2)
                Else 
                    $sWWN = $sWWN & ":" & Hex($sNodeWWN, 2)
                EndIf 
            Next 
            $aHBA[$NodeWWN][$i] = $sWWN
            $aHBA[$FirmwareVersion][$i] =  $Instance.FirmwareVersion
            $aHBA[$DriverName][$i] =  $Instance.DriverName
            $aHBA[$Active][$i] =  $Instance.Active
            $i = $i + 1
        Next
        Return $aHBA
    Else 
        Return 0
    EndIf    
EndFunc

Func MemoWrite($string)
    GUICtrlSetData($edtMemo, $string & @CRLF, 1)
EndFunc

Func Website()
    ShellExecute("http://www.puhua.net/get_hbawin")
EndFunc

#EndRegion Customer Functions






