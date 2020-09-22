VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComDia 
   Caption         =   "Common Dialog Controls  [example]"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdColor 
      Caption         =   "&Color Dialog"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "&Font Dialog"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Dialog"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save &As Dialog"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Dialog"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   120
      MaxLength       =   90
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "One 2000"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmComDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdColor_Click()
    On Error GoTo ColorErr 'catches error when user hits cancel
    'selects black when loads
    dlgCommon.Flags = cdlCCRGBInit
    'shows current color
    dlgCommon.Color = txtEdit.ForeColor
    '==================
    dlgCommon.ShowColor
    '==================
    'sets the color user selected to textbox
    txtEdit.ForeColor = dlgCommon.Color
ColorErr:
    Exit Sub

End Sub

Private Sub cmdFont_Click()
    On Error GoTo FontErr 'catches error when user hits cancel
    'loads the fonts
    dlgCommon.Flags = cdlCFScreenFonts
    'shows what the current settings are in in the Commondialog
    dlgCommon.FontName = txtEdit.FontName
    dlgCommon.FontBold = txtEdit.FontBold
    dlgCommon.FontItalic = txtEdit.FontItalic
    dlgCommon.FontSize = txtEdit.FontSize
    '=================
    dlgCommon.ShowFont
    '=================
    'changes the settings according to the commondialog changes
    txtEdit.FontName = dlgCommon.FontName
    txtEdit.FontBold = dlgCommon.FontBold
    txtEdit.FontItalic = dlgCommon.FontItalic
    txtEdit.FontSize = dlgCommon.FontSize
FontErr:
    Exit Sub
    
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo OpenErr 'catches error when user hits cancel
    dlgCommon.Filter = "Text Files (*.txt)|*.txt" 'sets the file type
    dlgCommon.FileName = "" 'default filename
    dlgCommon.ShowOpen
    Open dlgCommon.FileName For Input As #1 'opens file and gets txt of file
    txtEdit.Text = Input(LOF(1), 1) 'loads txt
    Close #1 'closes the file
OpenErr:
    
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo PrintErr 'catches error when user hits cancel
    dlgCommon.Flags = cdlPDHidePrintToFile + cdlPDNoPageNums 'hides the print to file option and page #
    dlgCommon.ShowPrinter
    'centers the txt horizontally to be printed on the paper
    Printer.ScaleLeft = -((Printer.Width - txtEdit.Width) / 2)
    '==========================
    Printer.ForeColor = txtEdit.ForeColor 'sets the color to be printed
    Printer.Print txtEdit.Text 'prints the one line of text
    Printer.EndDoc 'tells the printer only to print one line
PrintErr:
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveErr 'catches error when user hits cancel
    dlgCommon.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist 'sets flags to overwrite file and pathmustexist
    dlgCommon.Filter = "Text Files (*.txt)|*.txt" 'sets the file type
    dlgCommon.ShowSave
    Open dlgCommon.FileName For Output As #1 'gets and opens the txt file
    Print #1, txtEdit.Text 'saves file
    Close #1 'done with file and closes file
SaveErr:
    
End Sub

Private Sub Form_Load()
    dlgCommon.CancelError = True 'catches errors that occur when the user hits cancel

End Sub
