Attribute VB_Name = "basBinHex"
' ---------------------------------------------------------
' Microsoft Word 97 document format.
' Require Microsoft's Word 8.0 Object Library
' ---------------------------------------------------------

Option Explicit


Public Sub Insert2Word()

    Dim I As Integer
    Dim ColNames As String   'Col Headers
    Dim Wrd As Object        ' Word Object
    Dim TextData As String
    Dim TheRange As Object    'Word range

    ' ---------------------------------------------------------
    'Setup Col Headers
    ' ---------------------------------------------------------
    For I = 0 To 5
        ColNames = ColNames & "Column " & I + 1 & Chr(9)
    Next
    ColNames = ColNames & vbCrLf
    ' ---------------------------------------------------------
    ' clear objects
    ' ---------------------------------------------------------
      Set Wrd = Nothing
      Set TheRange = Nothing
      
    ' ---------------------------------------------------------
    ' Create new document to hold data
    ' ---------------------------------------------------------
    Set Wrd = CreateObject("Word.Basic")
    Wrd.FileNewDefault

    ' ---------------------------------------------------------
    ' setup the layout of this document
    ' ---------------------------------------------------------
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.5)
        .BottomMargin = InchesToPoints(0.5)
        .LeftMargin = InchesToPoints(0.75)
        .RightMargin = InchesToPoints(0.75)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalCenter
        .SuppressEndnotes = False
        .MirrorMargins = False
    End With
    DoEvents

    ' ---------------------------------------------------------
    ' Load the data into Word
    ' ---------------------------------------------------------
    With Wrd
        Dim Rows As Integer, Cols As Integer
        'Column Names
        .Insert ColNames
        TextData = ""
        For Rows = 0 To 53 '
            For Cols = 0 To 5
                'Create data line to insert
                TextData = TextData & "Row " & Rows & " Col " & Cols & Chr(9)
            Next
            'Insert into word
            .Insert TextData & vbCrLf
            TextData = ""
         Next
    End With
  
    ' ---------------------------------------------------------
    ' Define the font for this document
    ' ---------------------------------------------------------
    Set TheRange = ActiveDocument.Range(Start:=0, End:=0) ' all
    With TheRange
         .WholeStory
         .Font.Name = "Courier New"
         .Font.Size = 9
         .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    ' ---------------------------------------------------------
    'Uncomment and apply path for default save
    ' ---------------------------------------------------------
    'ActiveDocument.SaveAs ("D:\TestWord.Doc")
    'ActiveDocument.Application.Quit
    ' ---------------------------------------------------------
    ' Display the Word application
    ' ---------------------------------------------------------
    Wrd.AppShow
  
End Sub

