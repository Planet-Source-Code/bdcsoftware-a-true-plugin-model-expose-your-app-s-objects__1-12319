VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type GUIDs
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Declares needed to register object in the ROT
Private Const ACTIVEOBJECT_STRONG = 0
Private Const ACTIVEOBJECT_WEAK = 1
Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As Long, rclsid As GUIDs) As Long
Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32.dll" (ByVal pUnk As IUnknown, rclsid As GUIDs, ByVal dwFlags As Long, pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32.dll" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long
Dim OLEInstance As Long
  
Private mImageData() As Long 'Array of Image Points
Public Event ImageHasChanged() 'Event, that will be raised once the plugin is ready to transfer data back to the main program

Private Sub Class_Initialize()
'The magic happens here
'This code is responsible for creating the entry for "MainApp.MyImage" in the rot
'Replace "MainApp.MyImage" with your own class
    Dim mGUID As GUIDs
    Dim lp As Long
    OLEInstance = 0
    lp = CLSIDFromProgID(StrPtr("MainApp.MyImage"), mGUID)
    If lp = 0 Then
        lp = RegisterActiveObject(Me, mGUID, ACTIVEOBJECT_WEAK, OLEInstance)
    End If
End Sub

Friend Sub Quit()
    'Once we are done with the main program, lets clean up the rot
    'by removing the entry for our ActiveX Server
    
    If OLEInstance <> 0 Then
        RevokeActiveObject OLEInstance, 0
    End If
    CoDisconnectObject Me, 0
End Sub

Private Sub Class_Terminate()
    RevokeActiveObject OLEInstance, 0
End Sub

'Ok I know, an Array is not the best Idea to deal with images
' but for sake of simplicity I did not wanted to deal with DIBs

Property Get ImageData() As Variant
    ImageData = mImageData
End Property

Property Let ImageData(pImageData As Variant)
    mImageData = pImageData
    DoEvents
    RaiseEvent ImageHasChanged  'This just notifys the main app that
                                'Data has been modified (plugin or else)
End Property


