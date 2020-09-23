VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPictureLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
''Copyright 2002 Roger Gilchrist
''Inspired by
''Rocky Mountain Computer Consulting, Inc.
''http://www.vb-helper.com/HowTo/pasterch.zip
''USAGE
''Setup
''MINIMUM requirement: A Form with 2 picture boxes and a second Form with a Richtextbox and CommonDialog
''picturebox Form must resize itself to show one picturebox the other should be Visible=False
''both boxes should be same settings for all Read-Only at run Time Properties.
''Class takes care of all RunTime resets
'' To use class you will need to provide a toolbar or buttons/menus to access the various functions
''See frmPictureLoader in demo for more details

''Place the following in a Module because the class is called from both the FrmPictureLoader and the main form

'Public PL As New ClsPictureLoader

''In the Form_Load event of each form add this giving full Form.Control names as the controls are on diffenrent forms

'PL.AssignControls Form1.CommonDialog1, frmPictureLoader.picPictureLoader, Form1.RichTextBox1

''See frmPictureLoader in demo for details of calling

''PUBLIC members
''About --              Displays a simple description of the class
''AssignControls --     set controls for class to use
''ClipBoardClear --     Clear clipboard (use optionally in unload event for program)
''ClipBoardFormat--     Detect what if anything is in the Clipboard (Data might not be from your Program!)
''Crop --               Crop image
''DoBothScaling--       Boolean scale both ways
''DoHeightScaling--     Boolean scale height only
''DoWidthScaling--      Boolean scale width only
''FlipHorizontal--      Do horizontal flip
''FlipHStatus --        Read only used to reset tool buttons
''FlipVertical--        Do vertical flip
''FlipVStatus --        Read only used to reset tool buttons
''HeightScaling--       read/write width scale
''DocFolder --          set/get the Document folder Default = App.Path You might like to change it to use My Documents folder
''Paste --              set picture to ClipBoard then paste to RichTextBox
''PictureFolder --      set/get the picture folder Initial Default = App.Path resets to where every you load a picture from. You might like to change it to use My Pictures folder for initial value
''PictureIsAvailable--  boolean detector use for enable/disable menus/buttons etc test picture is in class
''PictureIsLoaded --    boolean detector use for enable/disable menus/buttons etc test picture is in RTB
''PictureLoad --        load new picture. Reset all settings
''PictureLoadToScale -- load picture with scaling,flipping and cropping
''RTFPictureScale --    Wrapper for RTF_SelRTFScale and RTF_textRTFScale
''SavePictureFromRTB -- Save selected picure from RTB to disk
''ScaleSize--           Read/write scale for current scale mode
''TrackPictureFolder -- if True PictureFolder changes to match latest picture source folder
''                      The matching TrackDocumentFolder is not used in this demo as the file handling is not
''                      part of this demo class.
''UseAppFolder --       Set DocFolder and PictureFolder to App.Path
''UseStandardFolders -- Set DocFolder and PictureFolder to MyDocuments and MyPictures folders
''WidthScaling--        read/write width scale
'
''PRIVATE members
''CheckOrientation--    Redo any fliping need for re-sizing
''CopyHiddenBack--      copy hidden picture box to display box
''CropMarksClear--      clear crop marks from image
''MyPictures --         Returns MyPictures path uses ReadKey
''MyDocuments --        Returns MyDocuments path uses ReadKey
''ReadKey --            Read standard folders
''ResizeHidden --       Resize the hidden box to hold images
''RTF_DoScaling --      Engine for RTF scaling
''RTF_SelRTFScale --    Call RTF_DoScaling for SelRTF
''RTF_textRTFScale --   Call RTF_DoScaling for whole text

''WithEvents members
''m_Pic_MouseDown--     activate cropping
''m_Pic_MouseMove--     show cropping box (not very weell)
''m_Pic_MouseUp         create final crop box

Private m_CDlg As CommonDialog
Private WithEvents m_Pic As PictureBox
Attribute m_Pic.VB_VarHelpID = -1
Private m_RTB As RichTextBox
Private HiddenPic As PictureBox
Private PicFilename As String
Private m_PictureFolder As String
Private m_TrackPictureFolder As Boolean ' change m_PictureFolder to match latest source folder DEFAULT True
Private m_DocFolder As String
Private m_PictureIsLoaded  As Boolean
Private FlipVState As Boolean
Private FlipHState As Boolean
Private m_Scale As Single
Private m_WidthScale As Single
Private m_HeightScale As Single
Private m_DoWidthScale As Boolean
Private m_DoHeightScale As Boolean
Private m_DoBothScale As Boolean
Private m_Croppable As Boolean
Private ClipXOrigin As Single
Private ClipYOrigin As Single
Private ClipXOld As Single
Private ClipYOld As Single
Private ClipXEnd As Single
Private ClipYEnd As Single
Private m_AutoReCrop As Boolean
Private ClipBorderColor As Long
Public Enum FlipMode
    FlipNone
    FlipH
    FlipV
    FlipHV
End Enum
#If False Then
Rem Skip On
'this line keeps Ulli's Code Formatter happy
'about this trick for preserving Enums case when you use them
Dim FlipNone
Dim FlipH
Dim FlipV
Dim FlipHV
#End If
Rem Skip Off
Private Const GraphicsFilter As String = "Graphic Files|*.bmp;*.ico;*.cur;*.rle;*.wmf;*.emf;*.gif;*.jpg|Bitmap images(.bmp)|*.bmp|BitMap images(RLE)|*.rle|Icons(.ico)|*.ico|Cursors(.cur)|*.cur|Windows MetaFiles images(.wmf)|*.wmf|Enhanced MetaFiles images(.emf)|*.emf|GIF images(.gif)|*.gif|JPEG images(.jpg)|*.jpg"
Private Const SaveFilter As String = "Bitmap images(.bmp)|*.bmp"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_PASTE As Long = &H302

Public Sub About()

  Dim msg As String

    'THis is an abreviated version of the help message designed to fit on a MsgBox
    msg = "A class inserts pictures in RichTextBox." & vbCr & _
          "Big images may load slowly or cause memory errors." & vbCr & _
          "Note: Image data is stored in RTF code; document size increases rapidly with many or big pictures." & vbCr & _
          "3 main actions of class." & vbCr & _
          "A. Pre-Insert" & vbCr & _
          " Files: bmp, ico, cur, rle, wmf, emf, gif & jpg" & vbCr & _
          " Resize images before pasting to document(smaller images reduce document size). " & vbCr & _
          " You can also Flip and Crop at this time." & vbCr & _
          "B. Post-Insert" & vbCr & _
          "Shrink & Expand buttons & Combo manipulate image in RTF code. " & vbCr & _
          "Allows fine-tuning of image but doesn't change data size in the document." & vbCr & _
          "No selection = all images re-sized else selected image only changes." & vbCr & _
          "Possible memory problems/crash for large images." & vbCr & _
          "C. Save image from RTF." & vbCr & _
          " Select image, press the 'Save a Picture' button." & vbCr & _
          " Saving Formats: .bmp, .rle, .wmf,  .gif or .jpg" & vbCr & _
          " Limitations: Only works if selection is picture only, file size is loaded size not RTF code size." & vbCr & _
          "" & vbCr & _
          "Note images can be cut and pasted." & vbCr & _
          "Copyright 2002 Roger Gilchrist"
    MsgBox msg, , "ClsPictureLoader"

End Sub

Public Sub AssignControls(CmDlg As CommonDialog, Pic As PictureBox, pic2 As PictureBox, RTB As RichTextBox)

  'Set the controls which will be used

    Set m_Pic = Pic
    Set HiddenPic = pic2
    Set m_CDlg = CmDlg
    Set m_RTB = RTB
    With m_Pic
        .BackColor = m_RTB.BackColor ' just in case the image has transparent bits (icons for example)
        .AutoSize = True
        .ScaleMode = vbTwips
    End With 'M_PIC
    With HiddenPic 'Match hidden PictureBox to display box
        .BackColor = m_Pic.BackColor
        .AutoSize = m_Pic.AutoSize
        .ScaleMode = m_Pic.ScaleMode
    End With 'HIDDENPIC

End Sub

Private Sub CheckOrientation()

    If FlipHState Then
        FlipHState = Not FlipHState 'turn off so it will turn on againg
        FlipVertical
    End If
    If FlipVState Then
        FlipVState = Not FlipVState
        FlipHorizontal
    End If

End Sub

Private Sub Class_Initialize()

    UseAppFolder
    ''or if you want to be more standard
    'UseStandardFolders
    TrackPictureFolder = True
    m_PictureIsLoaded = False
    m_DoBothScale = True
    m_HeightScale = 1
    m_WidthScale = 1
    m_Scale = 1

End Sub

Public Sub Clear()

  'reset all image manipulations to off

    m_Pic.Picture = LoadPicture()
    m_PictureIsLoaded = False
    m_AutoReCrop = False
    m_Pic.AutoRedraw = True
    m_Croppable = False
    FlipVState = False
    FlipHState = False
    m_Scale = 1
    m_DoBothScale = True

End Sub

Public Sub ClipBoardClear(Optional Auto As Boolean = True)

    If ClipBoardFormat <> "Empty" Then
        If Not Auto Then
            If MsgBox("You have data in the clipboard." & vbNewLine & "Do you want to keep it for use in other programs?", vbInformation + vbYesNo, App.Title) = vbYes Then
                Exit Sub '>---> Bottom
            End If
        End If
        DoEvents
        Clipboard.Clear
    End If

End Sub

Public Function ClipBoardFormat() As String

  'This is just a test function to check what is in the
  'ClipBoard while developing things
  'Also called by ClipBoardClear to see if anything is in ClipBoard

    If Clipboard.GetFormat(vbCFText) Then
        ClipBoardFormat = "vbCFText"
      ElseIf Clipboard.GetFormat(vbCFBitmap) Then 'CLIPBOARD.GETFORMAT(VBCFTEXT) = FALSE
        ClipBoardFormat = "vbCFBitmap"
      ElseIf Clipboard.GetFormat(vbCFRTF) Then 'CLIPBOARD.GETFORMAT(VBCFBITMAP) = FALSE
        ClipBoardFormat = "vbCFRTF"
      ElseIf Clipboard.GetFormat(vbCFLink) Then 'CLIPBOARD.GETFORMAT(VBCFRTF) = FALSE
        ClipBoardFormat = "vbCFLink"
      ElseIf Clipboard.GetFormat(vbCFMetafile) Then 'CLIPBOARD.GETFORMAT(VBCFLINK) = FALSE
        ClipBoardFormat = "vbCFMetafile"
      ElseIf Clipboard.GetFormat(vbCFDIB) Then 'CLIPBOARD.GETFORMAT(VBCFMETAFILE) = FALSE
        ClipBoardFormat = "vbCFDIB"
      ElseIf Clipboard.GetFormat(vbCFPalette) Then 'CLIPBOARD.GETFORMAT(VBCFDIB) = FALSE
        ClipBoardFormat = "vbCFPalette"
      ElseIf Clipboard.GetFormat(vbCFEMetafile) Then 'CLIPBOARD.GETFORMAT(VBCFPALETTE) = FALSE
        ClipBoardFormat = "vbCFEMetaFile"
      ElseIf Clipboard.GetFormat(vbCFFiles) Then 'CLIPBOARD.GETFORMAT(VBCFEMETAFILE) = FALSE
        ClipBoardFormat = "vbCFFiles"
      Else 'CLIPBOARD.GETFORMAT(VBCFFILES) = FALSE
        ClipBoardFormat = "Empty"
    End If

End Function

Private Sub CopyHiddenBack()

  'copy hidden picture back to visible m_pic picture

    m_Pic.PaintPicture HiddenPic.Image, _
                       0, 0, HiddenPic.Width, HiddenPic.Height, _
                       0, 0, HiddenPic.Width, HiddenPic.Height

End Sub

Public Sub Crop(Optional ReSize As Boolean = False)

  Dim X1 As Single, Y1 As Single
  Dim NWid As Single
  Dim NHigh As Single
  Dim ReScaleWid As Single
  Dim ReScaleHigh As Single

    ReScaleWid = IIf(ReSize, m_WidthScale, 1)
    ReScaleHigh = IIf(ReSize, m_HeightScale, 1)
    If m_Croppable Then
        CropMarksClear
        m_Pic.AutoRedraw = True
        NWid = Abs(ClipXOld - ClipXOrigin) * ReScaleWid
        NHigh = Abs(ClipYOld - ClipYOrigin) * ReScaleHigh
        'this is where autoCrop goes wrong I think
        X1 = IIf(ClipXOld > ClipXOrigin, ClipXOrigin, ClipXOld) * ReScaleHigh
        Y1 = IIf(ClipYOld > ClipYOrigin, ClipYOrigin, ClipYOld) * ReScaleWid

        HiddenPic.Width = NWid
        HiddenPic.Height = NHigh
        HiddenPic.PaintPicture m_Pic.Image, 0, 0, NWid, NHigh, X1, Y1, NWid, NHigh
        m_Pic.Picture = LoadPicture() 'clear old picture
        m_Pic.Width = NWid            'resize it
        m_Pic.Height = NHigh          '
        CopyHiddenBack
        m_AutoReCrop = True
        'm_Croppable = False
    End If

End Sub

Private Sub CropMarksClear()

  'remove cropping marks

    m_Pic.AutoRedraw = False
    m_Pic.Cls

End Sub

Public Property Let DoBothScaling(ByVal DoScale As Boolean)

    m_DoBothScale = DoScale
    If DoScale Then
        DoWidthScaling = False
        DoHeightScaling = False
    End If

End Property

Public Property Get DoBothScaling() As Boolean

    DoBothScaling = m_DoBothScale

End Property

Public Property Get DocFolder() As String

    DocFolder = m_DocFolder

End Property

Public Property Let DocFolder(ByVal DFolder As String)

  'allows you to set the document folder
  'defaults to App.Path

    m_DocFolder = DFolder

End Property

Public Property Let DoHeightScaling(ByVal DoScale As Boolean)

    m_DoHeightScale = DoScale
    If DoScale Then
        DoWidthScaling = False
        DoBothScaling = False
    End If

End Property

Public Property Get DoHeightScaling() As Boolean

    DoHeightScaling = m_DoHeightScale

End Property

Public Property Get DoWidthScaling() As Boolean

    DoWidthScaling = m_DoWidthScale

End Property

Public Property Let DoWidthScaling(ByVal DoScale As Boolean)

    m_DoWidthScale = DoScale
    If DoScale Then
        DoHeightScaling = False
        DoBothScaling = False
    End If

End Property

Public Sub FlipHorizontal()

  'horizontally flip visible m_pic picture to hidden picture

    ResizeHidden
    HiddenPic.PaintPicture m_Pic.Image, _
                           0, 0, m_Pic.Width, m_Pic.Height, _
                           m_Pic.Width, 0, -m_Pic.Width, m_Pic.Height
    CopyHiddenBack
    FlipVState = Not FlipVState

End Sub

Public Property Get FlipHStatus() As Boolean

    FlipHStatus = FlipHState

End Property

Public Sub FlipVertical()

  'vertically flip visible m_pic picture to hidden picture

    ResizeHidden
    HiddenPic.PaintPicture m_Pic.Image, _
                           0, 0, m_Pic.Width, m_Pic.Height, _
                           0, m_Pic.Height, m_Pic.Width, -m_Pic.Height
    CopyHiddenBack
    FlipHState = Not FlipHState

End Sub

Public Property Get FlipVStatus() As Boolean

    FlipVStatus = FlipVState

End Property

Public Property Let HeightScaling(ByVal ScaleSize As Single)

    m_HeightScale = ScaleSize

End Property

Public Property Get HeightScaling() As Single

    HeightScaling = m_HeightScale

End Property

Public Sub m_Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then

        m_Pic.DrawMode = vbXorPen
        ClipBorderColor = m_Pic.Point(X, Y)
        If ClipXOld <> ClipXOrigin Then
            m_Pic.Line (ClipXOrigin, ClipYOrigin)-(ClipXOld, ClipYOld), , B
        End If
        m_Croppable = True
        CropMarksClear
        m_AutoReCrop = False
        ClipXOrigin = X
        ClipYOrigin = Y
        ClipXOld = X
        ClipYOld = Y
        m_Pic.DrawWidth = 1

        m_Pic.MousePointer = vbSizePointer

    End If
    If Button = 2 Then
        m_Pic.Cls
    End If

End Sub

Public Sub m_Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'This should redraw the cropping rectangle as you resize the box but doesn't
  'let me know if you know why

    If Button = 1 Then
        If ClipXOld <> X And ClipYOld <> Y Then
            CropMarksClear
            m_Pic.DrawStyle = vbDot
            m_Pic.Line (ClipXOrigin, ClipYOrigin)-(ClipXOld, ClipYOld), ClipBorderColor, B
        End If
        m_Pic.AutoRedraw = True
        m_Pic.Line (ClipXOrigin, ClipYOrigin)-(X, Y), , B
        ClipXOld = X
        ClipYOld = Y

    End If

End Sub

Public Sub m_Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        m_Pic.DrawWidth = 1
        m_Pic.DrawStyle = vbDot
        CropMarksClear
        m_Pic.DrawMode = vbCopyPen
        m_Pic.DrawWidth = 1
        m_Pic.Line (ClipXOrigin, ClipYOrigin)-(X, Y), , B
        ClipXOld = X
        ClipYOld = Y
        m_Pic.MousePointer = vbDefault
    End If

End Sub

Private Function MyDocuments() As String

  ''these routines are from
  ''http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=40728&lngWId=1
  '' and allow you to get common default folders

    MyDocuments = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")

End Function

Private Function MyPictures() As String

  ''these routines are from
  ''http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=40728&lngWId=1
  '' and allow you to get common default folders

    MyPictures = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures")

End Function

Public Sub Paste()

  'this inserts a picture into the clipboard then dumps the clipboard to the RichtextBox
  'NOTE the image is still in the clipboard and can be pasted in again
  'with [Ctrl-V] or [Shift-Ins] or any other Paste method you may be using

    CropMarksClear
    Clipboard.Clear
    Clipboard.SetData m_Pic.Image
    SendMessage m_RTB.hwnd, WM_PASTE, 0, 0
    m_PictureIsLoaded = True

End Sub

Public Property Let PictureFolder(ByVal PFolder As String)

  'allows you to set the picture folder
  'defaults to App.Path

    m_PictureFolder = PFolder

End Property

Public Property Get PictureFolder() As String

    PictureFolder = m_PictureFolder

End Property

Public Property Get PictureIsAvailable() As Boolean

  'allows you to test that a picture was loaded

    PictureIsAvailable = m_PictureIsLoaded

End Property

Public Property Get PictureIsLoaded() As Boolean

  'allows you to test that a picture was loaded

    PictureIsLoaded = InStr(m_RTB.TextRTF, "{\pict") 'm_PictureIsLoaded

End Property

Public Sub PictureLoad()

    With m_CDlg
        .Filter = GraphicsFilter
        .InitDir = m_PictureFolder
        .FilterIndex = 1
        .ShowOpen
        If Len(.FileName) Then
            PicFilename = .FileName
            If m_TrackPictureFolder Then
                'change the picture folder to the last picture folder selected
                'assumes that if you are going to load another picture it is
                'probably in the same place as last one
                m_PictureFolder = Left$(PicFilename, InStrRev(PicFilename, "\") - 1)
            End If

            Clear
            PictureLoadToScale m_Scale

        End If
    End With 'M_CDLG

End Sub

Public Sub PictureLoadToScale(Optional Scaler As Single = 1)

  'can load picture pre-scaled
  'so that the RTF code is smaller

  Dim OWidth As Single
  Dim OHeight As Single

    If Len(PicFilename) Then

        m_Scale = Scaler ' reset scale to new value if needed
        m_Scale = IIf(m_Scale = 0, 1, m_Scale)
        m_HeightScale = IIf(m_HeightScale = 0, 1, m_HeightScale)
        m_WidthScale = IIf(m_WidthScale = 0, 1, m_WidthScale)
        If m_DoBothScale Then
            m_HeightScale = m_Scale
            m_WidthScale = m_Scale
          ElseIf m_DoHeightScale Then 'M_DOBOTHSCALE = FALSE
            m_HeightScale = Scaler
          ElseIf m_DoWidthScale Then 'M_DOHEIGHTSCALE = FALSE
            m_WidthScale = Scaler
        End If
        With m_Pic
            .Visible = False
            .Picture = LoadPicture(PicFilename) 'Load original to size picturebox properly
            m_PictureIsLoaded = Len(PicFilename)
            OWidth = .Width
            OHeight = .Height
            '        .Height = .Height
            '        If m_HeightScale Then
            .Height = .Height * m_HeightScale
            '        End If
            '        .Width = .Width '* Scaler
            '        If m_WidthScale Then
            .Width = .Width * m_WidthScale
            '        End If
            '       If Scaler <> 1 Then ' do nothing if scaler is 1
            .Picture = LoadPicture()
            .PaintPicture LoadPicture(PicFilename), _
                          0, 0, .Width, .Height, _
                          0, 0, OWidth, OHeight
            '       End If
            If m_AutoReCrop Then

            End If
            CheckOrientation
            Crop True
            .Visible = True
            .Refresh

        End With 'M_PIC
    End If

End Sub



Private Function ReadKey(Value As String) As String

  ''these routines are from
  ''http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=40728&lngWId=1
  '' and allow you to get common default folders
  '' slightly modified

  Dim B As Object

    On Error Resume Next
        Set B = CreateObject("wscript.shell")
        ReadKey = B.RegRead(Value)

    On Error GoTo 0

End Function

Private Sub ResizeHidden()

    With HiddenPic
        .Picture = LoadPicture()
        .Top = m_Pic.Top
        .Left = m_Pic.Left
        .Width = m_Pic.Width
        .Height = m_Pic.Height
    End With 'HIDDENPIC

End Sub

Private Function RTF_DoScaling(RTFCode As String, Scaler As Double, Optional OneOnly As Boolean = False) As String

  'picwgoal3089\pichgoal2025
  'this keeps picture in proportion
  'scaler 0 to 1 shrinks; 1 > ? expands 'maximum size is integer limit of 32767
  'This does not work for very large pictures because the RTFCode string is far too large for memory
  'NOTE
  'Because of the way it works each rescale is based on the current size of image.
  'If you scale up by 200% you have to scale down by 50% to get back to original size
  'It is just too complex to keep an array of the initial size of each image in the RTF
  'to allow you to automatically re-normaalise the resize values.
  'You would have to maintain hidden copies of each images size and location in the document and cope with new inserts, deletions and cut and paste
  'The picture loader routines cdo this by remembering the manipulation states, reloading the image and applying as needed
  
  'Remember you can also simply drag the selection handles on pictures to resize them in RichTextBoxes

  Dim picWidth As Single, picWidthStart As Single, picWidthEnd As Single
  Dim picHeight As Single, picHeightStart As Single, picHeightEnd As Single
  Dim i As Integer
  Dim OldSize As String, NewSize As String, NW As Long, NH As Long

    picWidthStart = 1

    On Error Resume Next
        Do While InStr(picWidthStart, RTFCode, "picwgoal")
            picWidthStart = InStr(picWidthStart, RTFCode, "picwgoal")

            If picWidthStart Then
                picWidthStart = picWidthStart + 8
                picWidthEnd = InStr(picWidthStart, RTFCode, "\") - picWidthStart
                picWidth = Mid$(RTFCode, picWidthStart, picWidthEnd)
                picHeightStart = InStr(picWidthStart, RTFCode, "pichgoal")
                picHeightStart = picHeightStart + 8
                picHeightEnd = InStr(picHeightStart, RTFCode, " ") - picHeightStart
                picHeight = Mid$(RTFCode, picHeightStart, picHeightEnd)

                OldSize$ = Mid$(RTFCode, picWidthStart, InStr(picWidthStart, RTFCode, " ") - picWidthStart)
                'this stuff is not needed unless you want to try and limit the scaling of your picture
                'picwgoal and pichgoal are integers
                '                If (picHeight * Scaler) < 32767 And (picWidth * Scaler) < 32767 Then
                '                    If (picHeight * Scaler) > 1 And (picWidth * Scaler) > 1 Then

                NewSize$ = OldSize$ & "\picscalex" & (100 * Scaler) & "\picscaley" & (100 * Scaler)
                'this stuff is not needed
                'NB the picscale[x|y]## does not appear in RTF code it resets picwgoal and pichgoal to percentage
                '                      Else 'NOT (PICHEIGHT...
                '                        NewSize$ = OldSize$
                '                        MsgBox " minimum picture size reached"
                '                    End If
                '                  Else 'NOT (PICHEIGHT...
                '                    NewSize$ = OldSize$
                '                    MsgBox " maximim picture size reached"
                '                End If

                RTFCode = Replace(RTFCode, OldSize$, NewSize$, 1)
                picWidthStart = picWidthStart + Len(NewSize) 'force past last location
            End If
            If OneOnly Then
                Exit Do '>---> Loop
            End If
            DoEvents
        Loop
        RTF_DoScaling = RTFCode
    On Error GoTo 0

Exit Function

woops:
    MsgBox Error(Err)

End Function

Private Sub RTF_SelRTFScale(Scaler As Double)

  'Make public if you wish to use directly
  'scale selected picture in RichTextBox
  'use the wrapper Sub RTFPictureScale rather than this because it assumes that you have selected a picture

    m_RTB.SelRTF = RTF_DoScaling(m_RTB.SelRTF, Scaler)

End Sub

Private Sub RTF_textRTFScale(Scaler As Double)

  'Make public if you wish to use directly
  'scale all pictures in RichTextBox
  'use the wrapper Sub RTFPictureScale rather than this because it assumes that you have no selection

    m_RTB.TextRTF = RTF_DoScaling(m_RTB.TextRTF, Scaler)

End Sub

Public Sub RTFPictureScale(Scaler As Double)

  'wrapper for two different ways to scale pictures (see above)
  ' this is safer than calling directly as you can't call the wrong one

    If Scaler <> 100 Then '100 = no change so ignore call
        m_RTB.MousePointer = vbHourglass
        If m_RTB.SelLength Then
            'scale only selected picture
            RTF_SelRTFScale Scaler
          Else 'M_RTB.SELLENGTH = FALSE
            'scale all pictures
            RTF_textRTFScale Scaler
        End If
        m_RTB.MousePointer = vbDefault
        m_RTB.SetFocus
    End If

End Sub

Public Sub SavePictureFromRTB()

  ' NOT FULLY DEVELOPED YET
  'NEEDS TO BE ABLE TO TEST THAT SELECTION IS ONLY A PICTURE

  ' only hit if is selection and contains a picture

    If m_RTB.SelLength Then
        If InStr(Left$(m_RTB.SelRTF, 2000), "{\pict") Then
            'If your picture doesn't hit try increasing the 2000 above
            'this test on only part of the SelRTF allows very large pictures to be inserted
            'increase should not be needed unless you have enormous font and colour tables
            'WARNING this is a crash point if the program is not visible
            'but is needed to make sure that SendKeys causes selection to go to ClipBoard
            'ANY SUGGESTIONS ON HOW TO AVOID THIS WILL BE APPRECIATED
            m_RTB.SetFocus      'shift focus back to RichTextBox
            SendKeys "^c", True ' Send Selection to Clipboard
            With m_CDlg
                .InitDir = m_PictureFolder
                .FileName = ""
                .Filter = SaveFilter
                .FilterIndex = 1
                .ShowSave
            End With 'M_CDLG

            'if there is a file name then process it
            If Len(m_CDlg.FileName) Then
                'set the picture to the picturebox
                With m_Pic
                    .Cls
                    .Picture = Clipboard.GetData()
                    Clipboard.Clear
                    .Refresh
                    'then save it
                    SavePicture .Image, m_CDlg.FileName
                End With 'M_PIC
            End If
        End If
    End If

End Sub

Public Property Let ScaleSize(ByVal SValue As Single)

    m_Scale = SValue

End Property

Public Property Get ScaleSize() As Single

    ScaleSize = m_Scale

End Property

Public Property Get TrackPictureFolder() As Boolean

    TrackPictureFolder = m_TrackPictureFolder

End Property

Public Property Let TrackPictureFolder(ByVal TrackPic As Boolean)

    m_TrackPictureFolder = TrackPic

End Property

Public Sub UseAppFolder()

  ' set the App.Path as initial directory for CommonDialog

    m_PictureFolder = App.Path
    m_DocFolder = App.Path

End Sub

Public Sub UseStandardFolders()

  ' set the MyPictures and MyDocuments as initial directories for CommonDialog

    m_PictureFolder = MyPictures
    m_DocFolder = MyDocuments

End Sub

Public Property Get WidthScaling() As Single

    WidthScaling = m_WidthScale

End Property

Public Property Let WidthScaling(ByVal ScaleSize As Single)

    m_WidthScale = ScaleSize

End Property

':) Ulli's VB Code Formatter V2.13.6 (4/12/2002 1:21:30 PM) 87 + 773 = 860 Lines (Skipped 5)
