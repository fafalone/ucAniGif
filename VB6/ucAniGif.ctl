VERSION 5.00
Begin VB.UserControl ucAniGif 
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   ToolboxBitmap   =   "ucAniGif.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   600
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   144
      Y1              =   0
      Y2              =   112
   End
   Begin VB.Line Line1 
      X1              =   144
      X2              =   0
      Y1              =   0
      Y2              =   112
   End
End
Attribute VB_Name = "ucAniGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ucAniGif v1.0.3
'by Jon Johnson (fafalone)
'
'Licensed under the MIT license
'
'Project home: https://github.com/fafalone/ucAniGif
'
'Updates:
'
'v1.0.3 (17 May 2024) - Back color not saved properly; now defaults to default UF BackColor.
'v1.0.2 (17 May 2024) - Back color wasn't initialized properly.
'

Option Explicit

Private pFactory As IShellImageDataFactory
Private pImage As IShellImageData

Private mInit As Boolean

Private mLoop As Boolean 'Unimplemented
Private Const mDefLoop As Boolean = True

Private mFile As String

Private mBk As Long
Private Const mDefBk As Long = &H8000000F

Private mAuto As Boolean
Private Const mDefAuto As Boolean = False

Private mSize As Boolean
Private Const mDefSize As Boolean = False

Private mPlaying As Boolean
Private mLoaded As Boolean


Private mAnim As Boolean 'IsAnimated; if not, don't try playing it.
Private mDelay As Long 'Frame delay

Private mCXY As SIZE

Private Sub UserControl_Initialize() 'Handles UserControl.Initialize
    Set pFactory = New ShellImageDataFactory
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag) 'Handles UserControl.ReadProperties
    mFile = PropBag.ReadProperty("File", "")
    'mLoop = PropBag.ReadProperty("Loop", mDefLoop)
    mAuto = PropBag.ReadProperty("Autoplay", mDefAuto)
    mSize = PropBag.ReadProperty("SizeToFit", mDefSize)
    mBk = PropBag.ReadProperty("BackColor", mDefBk)
    If mInit = False Then
        Init
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag) 'Handles UserControl.WriteProperties
    PropBag.WriteProperty "File", mFile, ""
    'PropBag.WriteProperty "Loop", mLoop, mDefLoop
    PropBag.WriteProperty "Autoplay", mAuto, mDefAuto
    PropBag.WriteProperty "SizeToFit", mSize, mDefSize
    PropBag.WriteProperty "BackColor", mBk, mDefBk
End Sub

Private Sub UserControl_InitProperties() 'Handles UserControl.InitProperties
    mLoop = mDefLoop
    mAuto = mDefAuto
    mSize = mDefSize
    mBk = mDefBk
    UserControl.BackColor = mBk
End Sub

Private Sub UserControl_Show() 'Handles UserControl.Show
    If mInit = False Then Init
End Sub

Private Sub Init()
    mInit = True
    If mFile <> "" Then
        If LoadImageFromFile(mFile) Then
            Line1.Visible = False
            Line2.Visible = False
            sidRedraw
            If Ambient.UserMode Then
                If mAuto Then Play
                Exit Sub
            End If
        End If
    End If
    Line1.X1 = 0
    Line1.X2 = UserControl.Width
    Line1.Y1 = 0
    Line1.Y2 = UserControl.Height
    Line2.X1 = 0
    Line2.X2 = UserControl.Width
    Line2.Y1 = UserControl.Height
    Line2.Y2 = 0
    UserControl.BackColor = mBk
End Sub

Private Sub UserControl_Resize() 'Handles UserControl.Resize
    If mLoaded Then
        sidRedraw
    Else
        Line1.X1 = 0
        Line1.Y1 = 0
        Line1.X2 = UserControl.ScaleWidth
        Line1.Y2 = UserControl.ScaleHeight
        
        Line2.X1 = 0
        Line2.Y1 = UserControl.ScaleHeight
        Line2.X2 = UserControl.ScaleWidth
        Line2.Y2 = 0
    End If
End Sub

Public Property Get BackColor() As OLE_COLOR: BackColor = mBk: End Property
Public Property Let BackColor(ByVal clr As OLE_COLOR)
    mBk = clr
    UserControl.BackColor = mBk
    If mLoaded Then
        sidRedraw
    End If
End Property

Public Property Get Autoplay() As Boolean: Autoplay = mAuto: End Property
Public Property Let Autoplay(ByVal bValue As Boolean): mAuto = bValue: End Property

Public Property Get SizeToFit() As Boolean: SizeToFit = mSize: End Property
Public Property Let SizeToFit(ByVal bValue As Boolean)
    If bValue <> mSize Then
        mSize = bValue
        If mLoaded Then sidRedraw
    End If
End Property
Public Property Get File() As String: File = mFile: End Property
Public Property Let File(ByVal sPath As String)
    mFile = sPath
    If LoadImageFromFile(mFile) Then
        Line1.Visible = False
        Line2.Visible = False
        sidRedraw
    End If
End Property

Public Sub Play()
    If mLoaded = False Then
        If LoadImageFromFile(mFile) Then
            Line1.Visible = False
            Line2.Visible = False
            Timer1.Interval = mDelay
            Timer1.Enabled = True
        End If
    Else
        Timer1.Interval = mDelay
        Timer1.Enabled = True
    End If
End Sub
Public Sub Pause()
    Timer1.Enabled = False
End Sub
Public Sub StopPlaying()
    Timer1.Enabled = False
    UserControl.Cls
    UserControl.Refresh
End Sub

Private Function LoadImageFromFile(sPath As String) As Boolean
    On Error GoTo e0
    Debug.Print "LoadImageFromFile->Entry, mFile=" & mFile & ", sPath=" & sPath
    Dim hr As Long
    mLoaded = False
    pFactory.CreateImageFromFile StrPtr(sPath), pImage
    If (pImage Is Nothing) = False Then
        hr = pImage.Decode(SHIMGDEC_DEFAULT, UserControl.ScaleWidth, UserControl.ScaleHeight)
        If SUCCEEDED(hr) Then
            Debug.Print "Loaded and decoded image..."
            If pImage.IsAnimated() = S_OK Then
                Debug.Print "Recognized animated gif..."
                mAnim = True
                pImage.GetDelay mDelay
            End If
            pImage.GetSize mCXY
            Debug.Print "LoadImageFromFile cx=" & mCXY.cx & ",cy=" & mCXY.cy
            mLoaded = True
            LoadImageFromFile = True
        Else
            Debug.Print "Failed to decode file, hr=" & hr ' & ": " & GetSystemErrorString(hr)
        End If
    Else
        Debug.Print "Failed to open file."
    End If
    
    Exit Function
e0:
    Debug.Print "Error loading file, " & Err.Number & ": " & Err.Description
End Function

Private Sub Timer1_Timer() 'Handles Timer1.Timer
    pImage.NextFrame
    sidRedraw
End Sub
Private Sub sidRedraw()
    UserControl.Cls
    Dim rcS As RECT, rcD As RECT
    rcS.Right = mCXY.cx
    rcS.Bottom = mCXY.cy
    If mSize Then
        rcD.Right = UserControl.ScaleWidth
        rcD.Bottom = UserControl.ScaleHeight
    Else
        rcD = rcS
    End If
    pImage.Draw UserControl.hDC, rcD, rcS
End Sub

#If TWINBASIC = 0 Then
    '[Description("Indicates whether an HRESULT value represents a successful operation (>= 0)")]
    Public Function SUCCEEDED(hr As Long) As Boolean
        SUCCEEDED = (hr >= 0)
    End Function
#End If
 
