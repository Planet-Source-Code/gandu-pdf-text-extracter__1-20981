VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "PDF Text"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGetText 
      Caption         =   "Get Text"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtPDFText 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1320
      Width           =   4500
   End
   Begin VB.CommandButton cmdGetFile 
      Caption         =   "Select PDF File"
      Height          =   375
      Left            =   100
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   100
      TabIndex        =   0
      Top             =   480
      Width           =   4500
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "File :"
      Height          =   285
      Left            =   100
      TabIndex        =   1
      Top             =   200
      Width           =   800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myPDF As Object
Private myPDFPage As Object
Private myPageHilite As Object
Private pageSelect As Object
    
Private pdfData As String
Private myPDFPageCount As Object
Private openResult As Boolean
Private closeResult As Boolean
Private hiliteResult As Boolean
Private pageCount As Integer

Private filelocation As String
Private pagenumber As Integer

Private Sub cmdGetFile_Click()
    
    'open the commondialog control
    CommonDialog1.ShowOpen
    
    'puts the filename selected inot the text box
    txtFileName.Text = CommonDialog1.FileName
End Sub


Private Sub cmdGetText_Click()
    
    'clean up
    txtPDFText.Text = ""
    pdfData = ""
    Label2.Caption = ""
    
    getBodyTextPDF
    txtPDFText.Text = pdfData
    
    MsgBox "Finished ; )", vbInformation, "PDF Text Data"
End Sub

Private Sub getBodyTextPDF()

    'instantiate the adobe object that we are going to use
    'we are using this object b/c this is the only object i
    'could find that had a function that returned the number of
    'pages in a prf file.  the number of pages is important later on
    
    Set myPDFPageCount = CreateObject("acroexch.pddoc")
    
    
    'when we open the file it will return true/false

    filelocation = txtFileName.Text
    openResult = myPDFPageCount.Open(filelocation)
    
    'little but of error handling, if we cannot open the file properly
    If openResult = False Then
        Set myPDFPageCount = Nothing
        MsgBox "Error opening file"
        Exit Sub
    End If
    
    'get the number of pages
    pageCount = myPDFPageCount.GetNumPages
    
    'when we close the file it will return truw/false
    closeResult = myPDFPageCount.Close
    
    'little but of error handling, if we cannot open the file properly
    If closeResult = False Then
        Set myPDFPageCount = Nothing
        MsgBox "Error closing file"
        Exit Sub
    End If

    'destroy the object we do not need it anymore
    Set myPDFPageCount = Nothing
    
    'i could only figure out how to get text from one page at a time
    'so i decided to run a loop that would get the text from a file
    'one page at a time. (adobe counts the first page)


    'instantiate object that we are going to use to get the text
    Set myPDF = CreateObject("acroexch.pddoc")
    
    'once again open the file
    openResult = myPDF.Open(filelocation)


    For pagenumber = 0 To pageCount - 1
        DoEvents
        getPDFTextFromPage pagenumber
        Label2.Caption = "Extracting : " & pagenumber + 1 & " of " & pageCount
    Next
    
    Set myPDF = Nothing

End Sub

Private Sub getPDFTextFromPage(pagenumber As Integer)

    'create pdf page object, with a specified page
    Set myPDFPage = myPDF.AcquirePage(pagenumber)

    'create a hilite object, this hilite object is what we will use to extract
    'the text, if you can hilite text then you can pull it out of the pdf file.
    Set myPageHilite = CreateObject("acroexch.hilitelist")
    
    'returns true/false, we are setting the parameters of the hilite object,
    'we are telling the hilite object that when you are called hilite the
    'entire page (0-9000)
    hiliteResult = myPageHilite.Add(0, 9000)

    'we are now going to hilite the page specified
    Set pageSelect = myPDFPage.CreatePageHilite(myPageHilite)
    
    'when pdf hilites it breaks up the page into little pieces so when we try
    'to extract that data from the hilite we ger it in little chuncks so have to loop
    'the data togther and append it together.
    
    'we can also use the same string  (pdfData to append all the pages together)
    Dim i As Integer
    For i = 0 To pageSelect.GetNumText - 1
        DoEvents
        pdfData = pdfData & pageSelect.GetText(i)
    Next

    'clean up

    Set myPDFPage = Nothing
    Set myPageHilite = Nothing
    Set pageSelect = Nothing
    
End Sub

Private Sub Form_Load()
    Label2.Caption = ""
End Sub
