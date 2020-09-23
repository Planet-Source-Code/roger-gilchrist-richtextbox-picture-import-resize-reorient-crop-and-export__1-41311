VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPictureLoader 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Picture loader -  Copyright 2002 Roger Gilchrist"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pctHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":0000
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":0542
            Key             =   "close"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":0A94
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":0FD6
            Key             =   "mirror"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":10E8
            Key             =   "flip"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":11FA
            Key             =   "width"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":1514
            Key             =   "height"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":182E
            Key             =   "both"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":1B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPictureLoader.frx":1E62
            Key             =   "crop"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "percentage resize"
      Top             =   60
      Width           =   855
   End
   Begin VB.PictureBox picPictureLoader 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open image file"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste Picture  to Document"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crop"
            Object.ToolTipText     =   "crop to selection frame (Resize before cropping)"
            ImageKey        =   "crop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fliph"
            Object.ToolTipText     =   "Swap Left && Right"
            ImageKey        =   "mirror"
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "flipv"
            Object.ToolTipText     =   "Swap Tob && Bottom"
            ImageKey        =   "flip"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "width"
            Object.ToolTipText     =   "Horizontal stretch only"
            ImageKey        =   "width"
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "height"
            Object.ToolTipText     =   "Vertical  stretch only"
            ImageKey        =   "height"
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "both"
            Object.ToolTipText     =   "Proportional  stretch"
            ImageKey        =   "both"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scalespacer"
            Style           =   4
            Object.Width           =   1000
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            Object.ToolTipText     =   "Close tool"
            ImageKey        =   "close"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPictureLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
''DON'T FORGET place this line in a Module somewhere
'Public PL As New ClsPictureLoader
Private NoClick As Boolean

Private Sub combo1_Click()

  Dim Tval As Single

    PL.PictureLoadToScale CSng(Val(Combo1.List(Combo1.ListIndex))) / 100
    picPictureLoader_Resize

End Sub

Private Sub Form_Activate()

  'picPictureLoader = LoadPicture()

    PL.Clear
    ToolBarButtonStatus

End Sub

Private Sub Form_Load()

  Dim i As Integer

    PL.AssignControls Form1.CommonDialog1, frmPictureLoader.picPictureLoader, frmPictureLoader.pctHidden, Form1.RichTextBox1
    picPictureLoader_Resize
    Combo1.Clear ' as you may be reloading this form many times clear and reset
    For i = 10 To 1000 Step 10
        Combo1.AddItem i
    Next i
    Combo1.Text = 100 ' set to 100%

End Sub

Private Sub picPictureLoader_Resize()

  'draw minimum form size to show picture

  Dim MinWidth As Long

    With picPictureLoader
        MinWidth = Toolbar1.Buttons("close").Left + Toolbar1.Buttons("close").Width + .Left
        If .Visible Then ' only if it is visible
            .Left = 120
            .Top = Toolbar1.Height
            If .Width > MinWidth Then
                Me.Width = (.Width + .Left * 3) '* 2' DEBUG Allows you to see pcthidden
                'pctHidden.Left = .Width
                'pctHidden.Visible = True
              Else 'NOT .WIDTH...
                Me.Width = MinWidth + .Left
            End If
            Me.Height = .Height + Toolbar1.Height * 2.5
            DoEvents
            .Refresh
            ToolBarButtonStatus
        End If
    End With 'PICPICTURELOADER

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim ComboReset As Boolean

    Select Case Button.Key
      Case "open"
        ComboReset = True
        PL.PictureLoad
        picPictureLoader_Resize
      Case "paste"
        PL.Paste
        Me.Hide
      Case "crop"
        PL.Crop
      Case "fliph"
        PL.FlipHorizontal
      Case "flipv"
        PL.FlipVertical
      Case "both"
        PL.DoBothScaling = Not PL.DoBothScaling
      Case "height"
        PL.DoHeightScaling = Not PL.DoHeightScaling
      Case "width"
        PL.DoWidthScaling = Not PL.DoWidthScaling
      Case "close"
        Me.Hide
    End Select
    ToolBarButtonStatus ComboReset

End Sub

Private Sub ToolBarButtonStatus(Optional SizeTo100 As Boolean)

  'keep the buttons in correct states
  'loading a new picture turns off all image manipulations so buttons and combo need to be reset

    With Toolbar1
        .Buttons("paste").Enabled = PL.PictureIsAvailable
        .Buttons("crop").Enabled = PL.PictureIsAvailable
        .Buttons("fliph").Enabled = PL.PictureIsAvailable
        .Buttons("flipv").Enabled = PL.PictureIsAvailable
        .Buttons("flipv").Value = IIf(PL.FlipHStatus, tbrPressed, tbrUnpressed)
        .Buttons("fliph").Value = IIf(PL.FlipVStatus, tbrPressed, tbrUnpressed)

        .Buttons("both").Enabled = PL.PictureIsAvailable
        .Buttons("width").Enabled = PL.PictureIsAvailable
        .Buttons("height").Enabled = PL.PictureIsAvailable
    End With 'TOOLBAR1
    If SizeTo100 Then
        Combo1.Text = 100
        Toolbar1.Buttons("both").Value = tbrPressed
    End If
    Combo1.Enabled = PL.PictureIsAvailable

End Sub

':) Ulli's VB Code Formatter V2.13.6 (4/12/2002 1:21:36 PM) 5 + 117 = 122 Lines
