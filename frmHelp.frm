VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   8145
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4683
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      FileName        =   "C:\Program Files\Microsoft Visual Studio\VB98\QND Programs\ExtendedRTF\picloaderTest\picloaderhelp.Rtf"
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   5040
      TabIndex        =   0
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    frmHelp.Hide

End Sub

Private Sub Form_Load()

  Dim msg As String

    ' uncomment this if you lose the help file for some reason.
    ''
    ''msg = "The class was designed to add small pictures to RichTextBoxes." & vbCr
    ''msg = msg & "Large images (25.6 MG original reduced to 30% size worked on my system) may load slowly or cause memory overload." & vbCr
    ''msg = msg & "Note that the data for the image is stored internally in the RTF code so documents rapidly increase in size if you add many or big pictures." & vbCr
    ''msg = msg & "The class offers two ways to re-size images; before pasteing to RTB and after pasteing." & vbCr & vbCr
    ''msg = msg & "Pre-Insertion Image Manipulation:" & vbCr
    ''msg = msg & "Loadable Formats: *.bmp,  *.ico, *.cur, *.rle, *.wmf, *.emf, *.gif and *.jpg" & vbCr
    ''msg = msg & "NEW Flip image: Flip the image by pressing the toggle buttons"
    ''msg = msg & "Top to Bottom" & vbCr
    ''msg = msg & "Left to Right " & vbCr
    ''msg = msg & "or both. " & vbCr & vbCr
    ''
    ''msg = msg & "NEW Resize Modes: " & vbCr
    ''msg = msg & "Resize allows you to load smaller versions of big images thus reducing the size of your document. " & vbCr
    ''msg = msg & "Resize whole image in 3 ways" & vbCr
    ''msg = msg & "1. Proportionally " & vbCr
    ''msg = msg & "2. Height only (Good for making graphic dividers)" & vbCr
    ''msg = msg & "3. Width only " & vbCr & vbCr
    ''
    ''msg = msg & "Resizing Inserted Images" & vbCr
    ''msg = msg & "The Shrink and Expand buttons and Combo on the Main form manipulate the image in RTF code. " & vbCr
    ''msg = msg & "This allows you to fine tune the image but doesn't change the size of the data in the document." & vbCr
    ''msg = msg & "If no image is selected then all images are re-scaled otherwise the selected image only is selected." & vbCr
    ''msg = msg & " May cause memory problems for large images." & vbCr
    ''msg = msg & vbCr
    ''msg = msg & "Saving an image from the RTB." & vbCr
    ''msg = msg & "   To do this you need to select the image first, then press the SaveImage button." & vbCr
    ''msg = msg & vbCr
    ''msg = msg & "Formats: .bmp, .rle, .wmf,  .gif and *.jpg (ico and cur could be saved but are not real icon/cursors)" & vbCr
    ''msg = msg & "Limitations: You can only save one image at a time." & vbCr
    ''msg = msg & "             Text in the selection results in a blank image." & vbCr
    ''msg = msg & "             The image is saved at the size set in the loading/pre-pasting operations." & vbCr
    ''msg = msg & vbCr
    ''msg = msg & "Note you can also simply cut and paste the image to a graphics program (or to an new spot on this document or another RTF document)."
    ''
    '' 'RichTextBox1.Text = msg

End Sub

Private Sub Form_Resize()

    With Command1
        .Top = frmHelp.Height - 850
        .Left = (frmHelp.Width - Command1.Width) / 2
    End With 'COMMAND1
    With RichTextBox1
        .Top = 0
        .Left = 0
        .Width = frmHelp.ScaleWidth
        .Height = frmHelp.Height - 900
    End With 'TEXT1'RICHTEXTBOX1
    Refresh

End Sub

':) Ulli's VB Code Formatter V2.13.6 (4/12/2002 1:21:17 PM) 1 + 68 = 69 Lines
