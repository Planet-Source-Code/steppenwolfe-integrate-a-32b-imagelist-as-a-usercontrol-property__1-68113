VERSION 5.00
Begin VB.UserControl ucTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demo UC"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   690
   End
End
Attribute VB_Name = "ucTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'/* based on the vbaccelerator imagelist control, with rewrites and additions:
'/* http://www.vbaccelerator.com/home/VB/Code/Controls/ImageList/vbAccelerator_Image_List_Control/article.asp

'/* How to use this /*
'-> Step 1
'/~ Add the clsImageList and ppgImages files to your usercontrol project
'/~ Instance the imagelist classes in the UserControl_Initialize routine
'-> Step 2
'/~ Add the GetImlObj, and  imagecount properties to the usercontrol
'-> Step 3
'/~ Set up the usercontrol ReadProperties and WriteProperties as demonstrated.
'-> Step 4
'/~ Open the usercontrol. Go to Tools -> Proceedure Attributes, and click 'Advanced'.
'/~ Select the image property, ex. 'LargeImages'. In the 'Use this page in Property
'/~ Browser' combo box, add the ppgImages property page. Repeat for each exposed image property.
'-> Step 5 (optional)
'/~ To use icons with alpha channels, link your project to the v6 (xp and up) version
'/~ of ComCtl32.DLL by adding a manifest. Note that render styles differ between com versions.

'/~ Thats it..

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Private m_lhMod             As Long
Private m_cSmallImages      As clsImageList
Private m_cLargeImages      As clsImageList

Public Property Get GetImlObj(Optional ByVal sPropName As String) As Long
Attribute GetImlObj.VB_MemberFlags = "400"
'/* add this property
'/* note that 'SmallImages' and 'LargeImages' are the uc
'/* property names you want to expose, and can be changed.
'/* If you want to use different property names, then you must
'/* change this routine, and the references in then PropertyPage_EditProperty
'/* event in the ppgImages property page to reflect the new property names
    If (sPropName = "SmallImages") Then
        GetImlObj = ObjPtr(m_cSmallImages)
    ElseIf (sPropName = "LargeImages") Then
        GetImlObj = ObjPtr(m_cLargeImages)
    End If
End Property

Private Sub UserControl_Initialize()
'/* you only need to instantiate the imagelists here

    m_lhMod = LoadLibrary("shell32.dll")
    Set m_cSmallImages = New clsImageList
    With m_cSmallImages
        .IconSizeX = 16
        .IconSizeY = 16
        .Create
    End With

    Set m_cLargeImages = New clsImageList
    With m_cLargeImages
        .IconSizeX = 32
        .IconSizeY = 32
        .Create
    End With
    
End Sub

'/* optional expose external methods
Public Property Get SmallImageCount() As Long
    SmallImageCount = m_cSmallImages.ImageCount
End Property

Public Property Get SmallImageX() As Long
    SmallImageX = m_cSmallImages.IconSizeX
End Property

Public Property Get SmallImageY() As Long
    SmallImageY = m_cSmallImages.IconSizeY
End Property

Public Property Get LargeImageCount() As Long
    LargeImageCount = m_cLargeImages.ImageCount
End Property

Public Property Get LargeImageX() As Long
    LargeImageX = m_cLargeImages.IconSizeX
End Property

Public Property Get LargeImageY() As Long
    LargeImageY = m_cLargeImages.IconSizeY
End Property

Public Sub Draw(ByVal lHdc As Long, _
                ByVal lIndex As Long, _
                ByVal lX As Long, _
                ByVal lY As Long, _
                ByVal lState As Long, _
                ByVal lDither As Long, _
                ByVal bLarge As Boolean)

    If bLarge Then
        m_cLargeImages.DrawImage lHdc, lIndex, lX, lY, lState, lDither
    Else
        m_cSmallImages.DrawImage lHdc, lIndex, lX, lY, lState, lDither
    End If
    
End Sub

'/* add the imagecount properties
Public Property Get SmallImages() As Long
Attribute SmallImages.VB_ProcData.VB_Invoke_Property = "ppgImages"
    If Not (m_cSmallImages Is Nothing) Then
        SmallImages = m_cSmallImages.ImageCount
    End If
End Property

Public Property Let SmallImages(ByVal PropVal As Long)
    If Not (m_cSmallImages Is Nothing) Then
        m_cSmallImages.ImageCount = PropVal
        PropertyChanged "SmallImageCount"
    End If
End Property

Public Property Get LargeImages() As Long
    If Not (m_cLargeImages Is Nothing) Then
        LargeImages = m_cLargeImages.ImageCount
    End If
End Property

Public Property Let LargeImages(ByVal PropVal As Long)
Attribute LargeImages.VB_ProcData.VB_Invoke_PropertyPut = "ppgImages"
    If Not (m_cLargeImages Is Nothing) Then
        m_cLargeImages.ImageCount = PropVal
        PropertyChanged "LargeImageCount"
    End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'/* set up your read write properties like this

Dim bIcons() As Byte
Dim bKeys() As Byte

    With PropBag
        '/* read small iml
        If Not (m_cSmallImages Is Nothing) Then
            ReDim bIcons(0)
            ReDim bKeys(0)
            m_cSmallImages.IconSizeX = .ReadProperty("SmallIconSizeX", 16)
            m_cSmallImages.IconSizeY = .ReadProperty("SmallIconSizeY", 16)
            m_cSmallImages.ColourDepth = .ReadProperty("SmallColourDepth", &H18)
            m_cSmallImages.ImageCount = .ReadProperty("SmallImageCount", 0)
            On Error Resume Next
            bIcons = .ReadProperty("SmallImages", "")
            bKeys = .ReadProperty("SmallKeys", "")
            If (UBound(bIcons) > 0) Then
                m_cSmallImages.RestoreIcons bKeys, bIcons
            End If
            On Error GoTo 0
        End If
        Erase bKeys
        Erase bIcons
        ReDim bIcons(0)
        ReDim bKeys(0)
        '/* read large iml
        If Not (m_cLargeImages Is Nothing) Then
            m_cLargeImages.IconSizeX = .ReadProperty("LargeIconSizeX", 32)
            m_cLargeImages.IconSizeY = .ReadProperty("LargeIconSizeY", 32)
            m_cLargeImages.ColourDepth = .ReadProperty("LargeColourDepth", &H18)
            m_cLargeImages.ImageCount = .ReadProperty("LargeImageCount", 0)
            On Error Resume Next
            bIcons = .ReadProperty("LargeImages", "")
            bKeys = .ReadProperty("LargeKeys", "")
            If (UBound(bIcons) > 0) Then
                m_cLargeImages.RestoreIcons bKeys, bIcons
            End If
            On Error GoTo 0
        End If
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Dim bIcons() As Byte
Dim bKeys() As Byte

    With PropBag
        '/* save small iml
        If Not (m_cSmallImages Is Nothing) Then
            .WriteProperty "SmallIconSizeX", m_cSmallImages.IconSizeX, 16
            .WriteProperty "SmallIconSizeY", m_cSmallImages.IconSizeY, 16
            .WriteProperty "SmallColourDepth", m_cSmallImages.ColourDepth, &H18
            .WriteProperty "SmallImageCount", m_cSmallImages.ImageCount, 0
            ReDim bIcons(0)
            ReDim bKeys(0)
            On Error Resume Next
            m_cSmallImages.SaveIcons bKeys, bIcons
            If (UBound(bIcons) > 0) Then
                .WriteProperty "SmallImages", bIcons, ""
            End If
            If (UBound(bKeys) > 0) Then
                .WriteProperty "SmallKeys", bKeys, ""
            End If
            On Error GoTo 0
        End If
        Erase bKeys
        Erase bIcons
        ReDim bIcons(0)
        ReDim bKeys(0)
        '/* save large iml
        If Not (m_cLargeImages Is Nothing) Then
            .WriteProperty "LargeIconSizeX", m_cLargeImages.IconSizeX, 32
            .WriteProperty "LargeIconSizeY", m_cLargeImages.IconSizeY, 32
            .WriteProperty "LargeColourDepth", m_cLargeImages.ColourDepth, &H18
            .WriteProperty "LargeImageCount", m_cLargeImages.ImageCount, 0
            ReDim bIcons(0)
            ReDim bKeys(0)
            On Error Resume Next
            m_cLargeImages.SaveIcons bKeys, bIcons
            If (UBound(bIcons) > 0) Then
                .WriteProperty "LargeImages", bIcons, ""
            End If
            If (UBound(bKeys) > 0) Then
                .WriteProperty "LargeKeys", bKeys, ""
            End If
            On Error GoTo 0
        End If
    End With

End Sub

Private Sub UserControl_Terminate()
    If Not (m_lhMod = 0) Then
        FreeLibrary m_lhMod
        m_lhMod = 0
    End If
    Set m_cSmallImages = Nothing
    Set m_cLargeImages = Nothing
End Sub
