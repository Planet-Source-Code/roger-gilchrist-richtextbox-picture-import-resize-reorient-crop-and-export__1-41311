VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "RTB Picture Pasteing Testbed Copyright 2002 Roger Gilchrist "
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                    "
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2790
      TabIndex        =   1
      ToolTipText     =   "Change picture size by X%"
      Top             =   50
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5953
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"picloadtest.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":0082
            Key             =   "Opendoc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":0194
            Key             =   "Savedoc"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":02A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":06F8
            Key             =   "Savepic"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":0B4A
            Key             =   "shrink"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":0F9C
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":13EE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":1A42
            Key             =   "new"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "picloadtest.frx":1F84
            Key             =   "Openpic"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newdoc"
            Object.ToolTipText     =   "Clear doc"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Opendoc"
            Object.ToolTipText     =   "Open an RTF document"
            ImageKey        =   "Opendoc"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Savedoc"
            Object.ToolTipText     =   "Save  an RTF document"
            ImageKey        =   "Savedoc"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "openpic"
            Object.ToolTipText     =   "Open a picture file"
            ImageKey        =   "Openpic"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "savepic"
            Object.ToolTipText     =   "Save  a picture file"
            ImageKey        =   "Savepic"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "shrink"
            Object.ToolTipText     =   "Decrease Picture Size 1/2 (RTF)"
            ImageKey        =   "shrink"
            Object.Width           =   1000
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "expand"
            Object.ToolTipText     =   "Increase  Picture Size X2 (RTF)"
            ImageKey        =   "expand"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1000
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileOpt 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "Save &As..."
         Index           =   3
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PL As New ClsPictureLoader
Public CurFileName As String

Private Sub combo1_Click()

    PL.RTFPictureScale CSng(Val(Combo1.List(Combo1.ListIndex))) / 100
    Combo1_getFocus

End Sub

Private Sub Combo1_getFocus()

  'resets combo to 100% after each scaling operation

    Combo1.Text = 100

End Sub

Private Sub Command1_Click(index As Integer)

    Select Case index
      Case 0
        frmPictureLoader.Show vbModal, Me
      Case 1 'Shrink
        PL.RTFPictureScale 0.5
      Case 2 'expand
        PL.RTFPictureScale 2
    End Select

End Sub

Private Sub Command3_Click()

    PL.SavePictureFromRTB

End Sub

Private Sub Form_Load()

  Dim i As Integer

    StatusBar1.SimpleText = "ClipBoard " & PL.ClipBoardFormat
    For i = 10 To 400 Step 10
        Combo1.AddItem i
    Next i
    Combo1.Text = 100 'This triggers the combo_Click event but 100% size is ignored by scaling routine

    CommonDialog1.InitDir = App.Path
    PL.AssignControls Form1.CommonDialog1, frmPictureLoader.picPictureLoader, frmPictureLoader.pctHidden, Form1.RichTextBox1
    PL.UseAppFolder ' use this for experimental stuff
    '  PL.UseStandardFolders 'use this for real applications
    RichTextBox1_SelChange ' reset buttons on toolbar
End Sub

Private Sub Form_Resize()

    With RichTextBox1
        .Top = Toolbar1.Height
        .Left = 50
        .Width = Form1.Width - 200
        .Height = Form1.ScaleHeight - StatusBar1.Height - Toolbar1.Height

    End With 'RICHTEXTBOX1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    PL.ClipBoardClear False
    End

End Sub

Private Sub mnuAbout_Click()

    PL.About

End Sub

Private Sub mnufileOpt_Click(index As Integer)

    With CommonDialog1
        .Filter = "Rich Text File|*.Rtf"
        .FilterIndex = 1
        .InitDir = PL.DocFolder
        .FileName = ""
    End With 'COMMONDIALOG1
    Select Case index
      Case 0
        RichTextBox1.Text = ""
        CurFileName = ""
      Case 1
        CommonDialog1.ShowOpen
        If Len(CommonDialog1.FileName) Then
            CurFileName = CommonDialog1.FileName
            RichTextBox1.FileName = CurFileName
            PL.DocFolder = Left$(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\") - 1)
        End If
      Case 2
        If Len(CurFileName) Then
            RichTextBox1.SaveFile CurFileName
          Else 'LEN(CURFILENAME) = FALSE
            CommonDialog1.ShowSave
            If Len(CommonDialog1.FileName) Then
                RichTextBox1.SaveFile CommonDialog1.FileName
                CurFileName = CommonDialog1.FileName
            End If

        End If

      Case 3
        CommonDialog1.ShowSave
        If Len(CommonDialog1.FileName) Then
            RichTextBox1.SaveFile CommonDialog1.FileName
            CurFileName = CommonDialog1.FileName
        End If

      Case 5
        PL.ClipBoardClear False
        End
    End Select

End Sub

Private Sub mnuhelp_Click()

    frmHelp.Show , Me

End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 9 Then 'QandD way to make insert sensible
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If

End Sub

Private Sub RichTextBox1_SelChange()

  Dim isPic As Boolean

    'this is a choke point disable it and RTB may be faster

    StatusBar1.SimpleText = "ClipBoard " & PL.ClipBoardFormat ' This is really only for DEBUG purposes
    'These are only for appearance sake,if there are no pictures the class prevents anything from happening
    isPic = PL.PictureIsLoaded
    Toolbar1.Buttons("shrink").Enabled = isPic
    Toolbar1.Buttons("expand").Enabled = isPic
    Toolbar1.Buttons("savepic").Enabled = isPic
    Combo1.Enabled = isPic

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
        Select Case Button.Key
          Case "newdoc"
            mnufileOpt_Click 0
          Case "Opendoc"
            mnufileOpt_Click 1
          Case "Savedoc"
            mnufileOpt_Click 2
          Case "openpic"
            frmPictureLoader.Show vbModal, Me
          Case "savepic"
            PL.SavePictureFromRTB
          Case "shrink"
            PL.RTFPictureScale 0.5
          Case "expand"
            PL.RTFPictureScale 2
          Case "help"
            mnuhelp_Click
        End Select
    On Error GoTo 0

End Sub

':) Ulli's VB Code Formatter V2.13.6 (4/12/2002 1:21:39 PM) 3 + 179 = 182 Lines
