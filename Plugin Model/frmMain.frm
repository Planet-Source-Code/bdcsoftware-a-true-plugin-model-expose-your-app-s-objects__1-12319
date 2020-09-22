VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MainApp"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   447
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   616
      TabIndex        =   1
      Top             =   600
      Width           =   9270
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1110
         Top             =   5940
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open File"
         Filter          =   "Pictures (*.bmp)|*.bmp"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   450
         Top             =   5895
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0454
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0D30
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1058
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open  "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Run Plugin"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit  "
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mImgData() As Long
'You have to declare the class withevents for the plugin to be able
'to "raise events" (fake)
Dim WithEvents mImg As MyImage
Attribute mImg.VB_VarHelpID = -1
Dim Loaded As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Sub Form_Unload(Cancel As Integer)
    'Lets just clean up
    mImg.Quit
    Set mImg = Nothing
End Sub

Private Sub mImg_ImageHasChanged()
    'Copy the ImageData back into the picture box
    Dim x As Integer
    Dim y As Integer
    Dim mData() As Long
    
    mData = mImg.ImageData
    
    If Loaded = False Then
        'MsgBox "Changed"
        For x = 0 To UBound(mData, 1)
            For y = 0 To UBound(mData, 2)
                'Debug.Print mData(x, y)
                Picture1.PSet (x, y), mData(x, y)
            Next y
        Next x
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        LoadImage
    ElseIf Button.Index = 5 Then
        Unload Me
    ElseIf Button.Index = 3 Then
        Loaded = False
        
        'We could have used an ActiveX Dll/EXE for our plugin
        'but I fellt the example would be simpler with just one Activex EXE
        
        'Start Plugin
        Call ShellExecute(0, vbNullString, App.Path & "\Plugin\Plugin.exe", "", App.Path & "\Plugin\", 0)
    End If
End Sub

Private Sub LoadImage()
    CD1.InitDir = App.Path
    CD1.ShowOpen
    If CD1.FileName <> "" Then
        Picture1.Picture = LoadPicture(CD1.FileName)
    End If
    
    Loaded = True
    
    CopyImageToArray
    
    'mImg will be created, it will register itself with the ROT (class.Initialize)
    Set mImg = New MyImage
    mImg.ImageData = mImgData

    Toolbar1.Buttons(3).Enabled = True
End Sub


Private Sub CopyImageToArray()
    Dim x As Integer
    Dim y As Integer
    
    ReDim mImgData(Picture1.ScaleWidth, Picture1.ScaleHeight) As Long
    For x = 0 To Picture1.ScaleWidth
        For y = 0 To Picture1.ScaleHeight
            mImgData(x, y) = Picture1.Point(x, y)
        Next
    Next
End Sub
