Attribute VB_Name = "Global_Offline"
Option Explicit

'///////////////////////////////////////
'Created by George Selos
'Unable to find this API anywhere on the Internet
'Used MSDN Library to find necessary function
'Any comments email me at gchelos@hotmail.com
'This is Freeware
'///////////////////////////////////////

Private Type INTERNET_CONNECTED_INFO
  dwConnectedState As Long
  dwFlags As Long
End Type

Const INTERNET_STATE_CONNECTED = &H1
Const INTERNET_STATE_DISCONNECTED_BY_USER = &H10
Const ISO_FORCE_DISCONNECTED = &H1
Const INTERNET_OPTION_CONNECTED_STATE = 50

Declare Function _
InternetSetOption Lib "Wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, _
ByVal dwOption As Long, _
ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long
                                     
Public Sub GoOffline(mOffline As Boolean)

Dim mInt As Long
Dim iso As INTERNET_CONNECTED_INFO

If mOffline = True Then
 iso.dwConnectedState = INTERNET_STATE_DISCONNECTED_BY_USER
 iso.dwFlags = ISO_FORCE_DISCONNECTED
Else
 iso.dwConnectedState = INTERNET_STATE_CONNECTED
End If

mInt = InternetSetOption(0&, INTERNET_OPTION_CONNECTED_STATE, iso, Len(iso))
End Sub
