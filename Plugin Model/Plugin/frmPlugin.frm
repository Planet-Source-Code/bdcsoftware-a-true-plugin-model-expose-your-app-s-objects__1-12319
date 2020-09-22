VERSION 5.00
Begin VB.Form frmPlugin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plugin"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Close Plugin && Transfer Data"
      Height          =   435
      Left            =   3750
      TabIndex        =   3
      Top             =   90
      Width           =   2925
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   15
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   591
      TabIndex        =   1
      Top             =   660
      Width           =   8895
   End
   Begin VB.CommandButton cmdModifyImage 
      Caption         =   "Modify Image"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1935
      TabIndex        =   0
      Top             =   90
      Width           =   1665
   End
End
Attribute VB_Name = "frmPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOBJ As Object
Dim mData() As Long

Private Sub cmdGetData_Click()
    
    Dim x As Integer
    Dim y As Integer
    
    
    Set mOBJ = GetObject(, "MainApp.MyImage")
    

    If IsArray(mOBJ.imagedata) Then
        mData = mOBJ.imagedata
        For x = 0 To UBound(mData, 1)
            For y = 0 To UBound(mData, 2)
                'Debug.Print mData(x, y)
                Picture1.PSet (x, y), mData(x, y)
            Next y
        Next x
    End If
    
    cmdModifyImage.Enabled = True
    DoEvents
End Sub

Private Sub cmdModifyImage_Click()
    
    drawtext Picture1.hdc, "Take a Look at the 'modify' routine by Rang3r.", 0, 170, vbWhite, 0.6, "Arial", 36
    drawtext Picture1.hdc, "It's very cool.", 0, 210, vbWhite, 0.2, "Arial", 82
    
End Sub

Private Sub cmdUnload_Click()
    Me.MousePointer = vbHourglass
    CopyImageToArray
    Me.Hide
    mOBJ.imagedata = mData
    Me.MousePointer = 0
    Unload Me
End Sub

Private Sub CopyImageToArray()
    Dim x As Integer
    Dim y As Integer
    
    ReDim mData(Picture1.ScaleWidth, Picture1.ScaleHeight) As Long
    For x = 0 To Picture1.ScaleWidth
        For y = 0 To Picture1.ScaleHeight
            mData(x, y) = Picture1.Point(x, y)
        Next
    Next
End Sub

Public Sub drawtext(hdc As Long, text As String, xpos As Long, ypos As Long, color As Long, opacity As Double, fontname As String, fontsize As Long)
    Dim size                                  As DWORD
    Dim ret                                   As Long
    Dim ndc                                   As Long
    Dim nbmp                                  As Long
    Dim hjunk
    Dim font                                  As LOGFONT
    Dim hfont                                 As Long
    Dim pixels()                              As RGBQUAD
    Dim npixels()                             As RGBQUAD
    Dim bgpixels()                            As RGBQUAD
    Dim rgbcol(3)                             As Byte
    Dim x, y, yy
    Dim bminfo                                As BITMAPINFO
    Dim tmp                                   As Double
    Dim alpha                                 As Double
    With font
        .lfHeight = -(fontsize * 20) / Screen.TwipsPerPixelY ' set font size
        .lfFaceName = fontname & Chr(0) 'apply font name
        .lfWeight = 0   'this is how bold the font is .. apply a in param if you want
    End With
    
    '-----------------------------------------
    'create a dc for our backbuffer
    ndc = CreateCompatibleDC(hdc)
    'create a bitmap for our backbuffer
    nbmp = CreateCompatibleBitmap(hdc, 1, 1) 'make a temp bitmap so we can get the size of the text
    'attach our bitmap to our backbuffer
    hjunk = SelectObject(ndc, nbmp)
    'apply the font to our backbuffer
    hfont = CreateFontIndirect(font)
    SelectObject ndc, hfont
    
    'get size of the text we want to draw
    ret = GetTabbedTextExtent(ndc, text, Len(text), 0, 0)
    
    'delete our temp bmp
    DeleteObject hfont
    DeleteObject ndc
    DeleteObject nbmp
    'this part was only to measure the size of the text
    '----------------------------------------
    'now lets draw the text...
    
    
    'split our color value to a byte array
    'this is my own invention ... pretty nice (?)
    CopyMemoryLong VarPtr(rgbcol(0)), VarPtr(color), 4
    'split the return value from gettextextent into two integers
    CopyMemoryLong VarPtr(size), VarPtr(ret), 4
    
    ypos = ypos - size.high / 2
    'create a dc for our backbuffer
    ndc = CreateCompatibleDC(hdc)
    'create a bitmap for our backbuffer
    nbmp = CreateCompatibleBitmap(hdc, size.low, size.high)
    'attach our bitmap to our backbuffer
    hjunk = SelectObject(ndc, nbmp)
    'apply the font to our backbuffer
    hfont = CreateFontIndirect(font)
    SelectObject ndc, hfont
    'set black background coloy
    SetBkColor ndc, 0
    'set white forecolor
    SetTextColor ndc, vbWhite
    'write the text to our backbuffer
    TabbedTextOut ndc, 0, 0, text, Len(text), 0, 0, 0
    'resize the arrays to the same size as the bbuffer
    ReDim pixels(size.low - 1, size.high - 1)
    ReDim npixels(size.low - 1, size.high - 1)
    ReDim bgpixels(size.low - 1, size.high - 1)
    
    'set the bitmap info (so we can get the gfx data in and out of our arrays
    With bminfo.bmiHeader
        .biSize = Len(bminfo.bmiHeader)
        .biWidth = size.low
        .biHeight = size.high
        .biPlanes = 1
        .biBitCount = 32
    End With
    'store the drawn text in our "pixels" array
    GetDIBits ndc, nbmp, 0, size.high, pixels(0, 0), bminfo, 1
    'get the bg graphics into our "bgpixels" array
    BitBlt ndc, 0, 0, size.low, size.high, hdc, xpos, ypos, vbSrcCopy
    GetDIBits ndc, nbmp, 0, size.high, bgpixels(0, 0), bminfo, 1
    yy = Int(size.high / 2)
    npixels = bgpixels
    For x = 0 To size.low - 2 Step 2
        For y = 0 To size.high - 2 Step 2
            'alpha is the average of the color of 2*2 pixels /255
            'now we have a value between 0 and 1
            '0 is transparent
            '1 is soild white
            'now multiply alpha with the opacity factor
            'ie if opacity is 0.5 ...  aplha will be max 0.5
            'since we draw our text with white . we only need to check the strength of one color (in this case blue)
            'coz red and green will always be the same as the blue
            alpha = (((0 + (pixels(x + 0, y + 0).rgbBlue) + (pixels(x + 1, y + 0).rgbBlue) + (pixels(x + 0, y + 1).rgbBlue) + (pixels(x + 1, y + 1).rgbBlue)) / 4) / 255) * opacity
            'alpha is now the opacity factor 0-1
            'calculate amount of blue to apply
            'and how much of the background that is going to be seen
            tmp = (alpha * rgbcol(2)) + bgpixels(x / 2, y / 2).rgbBlue * (1 - alpha)
            'never go higher than 255
            If tmp > 255 Then tmp = 255
            'store the result at x/2 and y/2 (the new picture is only 0.5 times as high and wide
            npixels(x / 2, y / 2).rgbBlue = tmp
            'calculate amount of red to apply
            'and how much of the background that is going to be seen
            tmp = (alpha * rgbcol(0)) + bgpixels(x / 2, y / 2).rgbRed * (1 - alpha)
            'never go higher than 255
            If tmp > 255 Then tmp = 255
            npixels(x / 2, y / 2).rgbRed = tmp
            'calculate amount of green to apply
            'and how much of the background that is going to be seen
            tmp = (alpha * rgbcol(1)) + bgpixels(x / 2, y / 2).rgbGreen * (1 - alpha)
            'never go higher than 255
            If tmp > 255 Then tmp = 255
            npixels(x / 2, y / 2).rgbGreen = tmp
        Next
    Next
    'apply the new picture to our bbuffer-dc
    SetDIBits ndc, nbmp, 0, size.high, npixels(0, 0), bminfo, 1
    'blit our bbuffer-dc to the screen
    BitBlt hdc, xpos, ypos, size.low, size.high, ndc, 0, 0, vbSrcCopy
    'clean up
    DeleteObject hfont
    DeleteObject ndc
    DeleteObject nbmp
End Sub
