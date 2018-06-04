Attribute VB_Name = "modPrinter"
Option Explicit
' +-----------------------------------+
' |                                   |
' |     M A R K   M E R C E R ' S     |
' |    P R I N T E R   M O D U L E    |
' |  WEB: www.woodjoint.ca/mark       |
' |  EMAIL: mark@creatiworks.ca       |
' |  VERSION: 1.8.3                   |
' |  DATE: May 26, 2009               |
' |                                   |
' +-----------------------------------+

' ****************NOTE****************

' THIS MODULE REQUIRES THE FOLLOWING
' COMPONENTS, PLEASE ADD THEM IN YOUR
' COMPONENTS LIST.

' Microsoft Windows Common Controls 6.0 (SP6)
' Microsoft Windows Common Controls-2 6.0
' Microsoft Chart Control 6.0 (SP4)
' Microsoft FlexGrid Control 6.0 (SP3)



' .: P A T C H   N O T E S :.
'-----------------------------------------------
'previously: v0.0:   big mess
'2007-01-12: v1.0:   Started tracking the function by version
'2007-01-15: v1.1:   Added twips scalemode check
'                    Added optionbox and checkbox
'                    synchronized with other versions to add MSFlexgrid, Bezier, and MSChart Linegraph, Bargraph, and piechart
'                    Controls now calculate their absolute position by inheriting relative positions from their containers (added this in general for all control types, not just line)
'2007-02-01: v1.2:   Added auto word wrap to printed labels and textboxes using setRect and drawTextEx api calls... issues later
'2007-03-08: v1.2.1: Fixed double && in button names
'2007-03-09: v1.3:   Removed dependancy on frmMain.Picture1
'                    Added color to labels and text boxes
'                    fixed detection of non-viewable controls like imagelists and timers
'                    Controls now inherit visibility from their containers
'2007-12-19: v1.4:   Added the LOGFONT type and other declarations necessary for printRotated (these were giving compile errors)
'2008-01-03: v1.5:   fixed offsets for MSchart controls
'                    fixed shape background transparency via fillstyle
'                    consolidated the printTextColor and printText functions allowing for transparent backgrounds
'                    fixed label coloring and background transparency
'2008-06-19: v1.6:   fixed issues with drawtextex, when printing on other printers it had major scaling issues.  now using printer.scalex and .scaley
'                    also always remember that api calls print things in the vbPixel scalemode... so always switch scalemodes for this.
'                    add 'SetRectEmpty Rec' at the end of the the function before Rec is destroyed... this was causing serious crashes / problems with unsupported automation types!
'                    fixed double or missing &'s in labels and multiline textboxes.
'                    printListView - If multiPage and lv.ListItems.Count > pageLength * p Then Printer.EndDoc (added the '* p') so that it doesnt enddoc at the last page
'2008-06-23: v1.6.1: fixed scaling issues with checkbox control
'                    Added the "Frame" control as a printable type.
'2008-07-07: v1.6.2: Fixed printFlexGrid function to not print columns with zero width
'                    Fixed printTextRect function (untested)
'2008-07-08: v1.6.3: printListView now vertically positions text based on twips rather than the font height 'where it lies' method
'2008-07-08: v1.6.4: Fixed printListView text was not aligning properly in multipage listview's as a result of last update.
'2008-12-04: v1.7:   Fixed font matching for DrawTextEx for multiline TextBoxes and Labels (bold, fontname, fontsize)
'2009-01-27: v1.8:   Extension of fixes from v1.4.  The LOGFONT i had setup was not working plus the printRotated function still had a dependancy on a physical picturebox.
'                    copying from the picturebox to printer object was proving to be difficult, slow, and poor quality.  So i looked up online how to print rotated text directly to the printer.
'                    I implemented a new printRotated method and abolished the old picturebox.
'                    printBarGraph: I changed the printer.enddoc inside the loop to .newpage and put an .enddoc at the end of the function...
'                    this allows the program to print to the Adobe PDF printer and have all pages in the same pdf file.
'2009-02-20: v1.8.1: Added italics to labels, textboxes, and buttons... some fontNAMES automatically set italics to true.
'                    Added sizing for form backgrounds so that it doesnt print a full page of grey background when the form is only small...
'2009-03-18: v1.8.2: corrected the sizing for form backgrounds, now uses scalex and scaley and f.picture if a picture has been specified for the form and f.image otherwise.
'2009-05-26: v1.8.3: added default values for horizOffset and vertiOffset
'2013-06-03: v1.8.4: added the framePrint function to print just a single frame from a form


' ************************************

Public Const DT_CENTER = &H1
Public Const DT_WORDBREAK = &H10
Public Const LF_FACESIZE = 32
Public Const PI = 3.14159265359

Private Type RECT
  Left   As Long
  Top   As Long
  Right   As Long
  Bottom   As Long
End Type

Public Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type

Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, ByVal pJob As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long ' or Boolean

Public Declare Function DrawTextEx Lib "user32.dll" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, ByRef lpRect As RECT, ByVal un As Long, ByRef lpDrawTextParams As Any) As Long
Public Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32.dll" (ByRef lpRect As RECT) As Long


Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long _
        ) _
        As Long
        



Sub selectPrinter(ByVal pName As String)
    Dim P As Printer
    For Each P In Printers
        If P.DeviceName = pName Then
            Set Printer = P
        End If
    Next P
End Sub

Function greyScale(color As Long) As Long
Dim r, g, b As Long
r = (color And &H10000FF)
g = (color And &H100FF00) / &H100
b = (color And &H1FF0000) / &H10000

r = Int((r + g + b) / 3#)
greyScale = r + r * &H100 + r * &H10000

End Function




Sub printBarGraph(ByRef frm As Form, ByRef printObj As Printer, ByRef Chart As MSChart, ByVal Left As Double, ByVal Top As Double, ByVal width As Double, ByVal height As Double, ByVal BarsPerPage As Long, ByVal header As String, ByVal firstPageOnly As Boolean, Optional ByRef status_string As Variant)
    Dim i As Byte
    Dim col As Long
    Dim shadCol As Long
    Dim cat As Long
    Dim barLeft As Double
    Dim barHeight As Double
    Dim barWidth As Double
    Dim graphTop As Double
    Dim graphBottom As Double
    Dim graphLeft As Double
    Dim graphRight As Double
    Dim scaleHeightConst As Double
    Dim pages As Byte
    Dim currentPage As Byte
    Dim M, n As Double
    Dim twips_per_inch  As Double
    
    Dim rot_x As Long
    Dim rot_y As Long
    
    twips_per_inch = 1440
    
    Chart.row = 1
    Chart.column = 1
    graphTop = Top * twips_per_inch
    graphBottom = (height + Top) * twips_per_inch
    graphLeft = Left * twips_per_inch
    graphRight = (width + Left) * twips_per_inch
    scaleHeightConst = Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum / (graphBottom - graphTop)
    barWidth = (graphRight - graphLeft) * 0.98 / BarsPerPage
    'col = &HFF
    shadCol = &H555555 ' &H660000
    
    If Chart.RowCount Mod BarsPerPage > 0 Then
        pages = Int(Chart.RowCount / BarsPerPage) + 1
    Else
        pages = CLng(Chart.RowCount / BarsPerPage)
    End If
    
    For currentPage = 1 To pages
        printObj.ScaleMode = vbTwips
        printObj.Orientation = vbPRORLandscape
        'printObj.Cls
        
        printObj.FontBold = False
        printObj.FontItalic = True
        printObj.fontname = "Verdana"
        printObj.fontsize = 10
        printObj.CurrentX = 1 * twips_per_inch
        printObj.CurrentY = 0.2 * twips_per_inch
        printObj.Print Date
        printObj.CurrentX = 1 * twips_per_inch
        printObj.CurrentY = 0.5 * twips_per_inch
        printObj.Print header
        printObj.CurrentX = 9 * twips_per_inch
        printObj.CurrentY = 0.2 * twips_per_inch
        printObj.Print "Page " & currentPage & " of " & pages
        printObj.FontItalic = False
    
        printObj.Line (graphLeft, graphTop)-(graphLeft, graphBottom + 0.02 * twips_per_inch), 0
        printObj.Line (graphRight, graphTop)-(graphRight, graphBottom + 0.02 * twips_per_inch), 0
        printObj.Line (graphLeft, graphTop)-(graphRight, graphTop), 0
        printObj.Line (graphLeft, (graphTop * 3 + graphBottom) / 4)-(graphRight, (graphTop * 3 + graphBottom) / 4), 0
        printObj.Line (graphLeft, (graphTop + graphBottom) / 2)-(graphRight, (graphTop + graphBottom) / 2), 0
        printObj.Line (graphLeft, (graphTop + graphBottom * 3) / 4)-(graphRight, (graphTop + graphBottom * 3) / 4), 0
        printObj.Line (graphLeft, graphBottom + 0.02 * twips_per_inch)-(graphRight, graphBottom + 0.02 * twips_per_inch), 0
        printObj.fontname = "Times New Roman"
        printObj.FontBold = False
        printObj.fontsize = 8
        printObj.CurrentX = graphLeft - 0.2 * twips_per_inch
        printObj.CurrentY = graphBottom - 0.1 * twips_per_inch
        printObj.Print "0"
        printObj.CurrentX = graphLeft - 0.5 * twips_per_inch
        printObj.CurrentY = graphTop - 0.1 * twips_per_inch
        printObj.Print Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum
        printObj.CurrentX = graphLeft - 0.5 * twips_per_inch
        printObj.CurrentY = (graphTop + graphBottom) / 2 - 0.1 * twips_per_inch
        printObj.Print CLng(Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum / 2)
        printObj.CurrentX = graphLeft - 0.5 * twips_per_inch
        printObj.CurrentY = (graphTop * 3 + graphBottom) / 4 - 0.1 * twips_per_inch
        printObj.Print CLng(Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum * 0.75)
        printObj.CurrentX = graphLeft - 0.5 * twips_per_inch
        printObj.CurrentY = (graphTop + graphBottom * 3) / 4 - 0.1 * twips_per_inch
        printObj.Print CLng(Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum * 0.25)
        
        printObj.FontBold = True
        printObj.fontsize = 10
            
        For i = 1 + (currentPage - 1) * BarsPerPage To Chart.RowCount
            If i > currentPage * BarsPerPage Then Exit For
            If TypeOf status_string Is Label Then
                status_string.Caption = Chart.RowLabel
            Else
                status_string = Chart.RowLabel
            End If
            DoEvents
            Chart.row = i
            barLeft = graphLeft - 0.2 * twips_per_inch + barWidth * ((CDbl(i - 1) Mod BarsPerPage) + 1)
            barHeight = val(Chart.Data) / scaleHeightConst
            
            DoEvents
            
            '_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
            'Printing rotated text
            rot_x = Printer.ScaleX((1 + ((i - 1) Mod BarsPerPage)) * barWidth + 280, vbTwips, vbPixels)
            rot_y = Printer.ScaleY(graphBottom + 1500, vbTwips, vbPixels)
            printRotated Chart.RowLabel, 70, rot_x, rot_y, "Arial", 16
            DoEvents
            
            'just make the color of the bars blue by default
            col = vbBlue
            
            frm.Picture1.Cls
            printObj.Line (barLeft + 0.04 * twips_per_inch, graphBottom)-(barLeft + 0.04 * twips_per_inch + barWidth * 0.64, graphBottom - barHeight - 0.03 * twips_per_inch), shadCol, BF
            printObj.Line (barLeft, graphBottom)-(barLeft + barWidth * 0.64, graphBottom - barHeight), col, BF
            printObj.Line (barLeft - 0.03 * twips_per_inch + barWidth, graphBottom + 0.02 * twips_per_inch)-(barLeft - 0.42 * twips_per_inch + barWidth, graphBottom + 1.1 * twips_per_inch), 0
        Next i
        DoEvents
        printObj.NewPage
        If firstPageOnly Then
            Exit For ' use this line when testing, prints only page 1
        End If
    Next currentPage
    printObj.FontBold = False
    printObj.EndDoc
    DoEvents
End Sub


Function printRotated(ByVal OutString As String, ByVal angle_deg As Double, ByVal x As Long, ByVal y As Long, ByVal fontname As String, ByVal fontsize As Double)
      Dim lf As LOGFONT
      Dim result As Long
      Dim hOldfont As Long
      Dim hPrintDc As Long
      Dim hFont As Long
      
      hPrintDc = Printer.hdc
      lf.lfEscapement = Int(angle_deg * 10#)
      lf.lfHeight = (fontsize * -20) / Printer.TwipsPerPixelY
      lf.lfFaceName = fontname
      hFont = CreateFontIndirect(lf)
      hOldfont = SelectObject(hPrintDc, hFont)
      result = TextOut(hPrintDc, x, y, OutString, Len(OutString))
      result = SelectObject(hPrintDc, hOldfont)
      result = DeleteObject(hFont)
End Function


Sub printPieChart(ByRef printObj As Variant, ByRef Chart As MSChart, ByVal xCenter As Double, ByVal yCenter As Double, ByVal rad As Double, ByVal legendColumnWidthMult As Double)
    Dim i As Byte
    Dim Total As Long
    Dim angleStart As Double
    Dim angleEnd As Double
    Dim col As Long
    Dim x, y As Double
    Dim centerslicex, centerslicey As Double
    Dim backupPrintObjFillStyle As Long
    'printObj.Circle (xCenter, yCenter), rad, 0 '&HFF
    Total = 0
    For i = 1 To Chart.ColumnCount
        Chart.column = i
        Total = Total + Chart.Data
    Next i
    angleStart = 0
    angleEnd = 0
    printObj.ScaleMode = vbInches
    printObj.forecolor = vbBlack
    printObj.FontBold = True
    printObj.fontname = "Times New Roman"
    printObj.fontsize = 10
    printObj.CurrentX = xCenter - rad
    printObj.CurrentY = yCenter - rad * 1.2
    printObj.Print Chart.RowLabel
    printObj.FontBold = False
    printObj.fontsize = 8
    For i = 1 To Chart.ColumnCount
        Chart.column = i
        angleEnd = angleStart + (Chart.Data / Total) * 360#
        
        backupPrintObjFillStyle = printObj.FillStyle
        printObj.FillStyle = vbFSSolid
        col = CLng(Chart.Plot.SeriesCollection(i).Pen.VtColor.Red) + CLng(Chart.Plot.SeriesCollection(i).Pen.VtColor.Green) * 256 + CLng(Chart.Plot.SeriesCollection(i).Pen.VtColor.Blue) * 65536
        If angleEnd > angleStart + 0.5 Then
            printObj.FillColor = col
            printObj.Circle (xCenter, yCenter), rad, col, ((720 + angleEnd - 90) Mod 360 - 360) / 180 * PI, ((720 + angleStart - 90) Mod 360 - 360) / 180 * PI
        End If
        
        If CLng((Chart.Data / Total) * 100) >= 5 Then
            centerslicex = xCenter + Cos(((angleStart + angleEnd) / 2# - 90) / 180# * PI) * rad * 0.65
            centerslicey = yCenter + Sin(((angleStart + angleEnd) / 2# - 90) / 180# * PI) * rad * 0.65
            printObj.CurrentX = centerslicex - (0.5 * ((Len(Left$(Chart.ColumnLabel, 12)) + 1) * 0.05))
            printObj.CurrentY = centerslicey - 0.1
            printObj.forecolor = vbWhite 'line not tested as of aug 9, 2004
            printObj.Print Left$(Chart.ColumnLabel, 12) & "."
        End If
        angleStart = angleEnd
        
        If i > Chart.ColumnCount / 2 Then
            x = xCenter + rad * 0.2 * legendColumnWidthMult
            y = yCenter + rad * 1.2 + (i - Chart.ColumnCount / 2) * rad * 0.2
        Else
            x = xCenter - rad * 0.9 * legendColumnWidthMult
            y = yCenter + rad * 1.2 + i * rad * 0.2
        End If
        printObj.forecolor = vbBlack
        printObj.Line (x, y)-(x + rad * 0.15, y + rad * 0.15), col, BF
        printObj.CurrentX = x + rad * 0.2
        printObj.CurrentY = y
        printObj.Print Chart.ColumnLabel & " " & CLng((Chart.Data / Total) * 100) & "%"
    Next i
    printObj.FillStyle = backupPrintObjFillStyle

End Sub


Public Function PrinterBusy(ByVal PrinterDeviceName As String) As Boolean
Dim hPrinter        As Long
Dim BytesNeeded     As Long
Dim JobsReturned    As Long

'Get the printer Handle
OpenPrinter PrinterDeviceName, hPrinter, ByVal 0&

'Get the Printer active jobs
EnumJobs hPrinter, 0, 127, 1, ByVal 0&, 0, BytesNeeded, JobsReturned

If BytesNeeded = 0 Then
    PrinterBusy = False
Else
    ReDim TempBuff(BytesNeeded - 1) As Byte
    EnumJobs hPrinter, 0, 127, 1, TempBuff(0), BytesNeeded, BytesNeeded, JobsReturned
    PrinterBusy = (JobsReturned > 0)
End If
'Close printer
ClosePrinter hPrinter
End Function



Sub printTextColor(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal forecolor As Long, ByVal backcolor As Long, ByVal textwidth As Double, ByVal fontname As String, ByVal fontsize As Byte, ByVal bold As Boolean, ByVal justify As Byte)
    'prints text by specifying foreground and background colors
    '(negative values for backcolor makes it transparent)
    
    Dim textlength As Double
    
    Printer.forecolor = forecolor
    Printer.Font = fontname
    Printer.FontBold = bold
    Printer.fontsize = fontsize
        
    'determine the width of the text
    textlength = Printer.textwidth(s)
    
    If backcolor >= 0 Then ' for transparent background pass in -1
        Printer.FillStyle = vbFSSolid
        Printer.FillColor = backcolor
        Printer.Line (x, y)-(x + textwidth, y + (22 * fontsize) + 15), backcolor, BF
        Printer.Line (x, y)-(x + textwidth, y + (22 * fontsize) + 15), backcolor, B
    End If
    
    If justify = 1 Then
        Printer.CurrentX = x + textwidth - textlength
        Printer.CurrentY = y
    ElseIf justify = 2 Then
        Printer.CurrentX = x + (textwidth / 2#) - (textlength / 2#)
        Printer.CurrentY = y
    Else
        Printer.CurrentX = x
        Printer.CurrentY = y
    End If
    
    Printer.Print s
    
End Sub


Sub printTextRect(ByVal Text As String, ByVal Left As Long, ByVal Top As Long, ByVal width As Long, ByVal height As Long)
    Dim rec As RECT
    Printer.ScaleMode = vbPixels
    
    Text = Replace(Text, " & ", " && ")
    SetRect rec, Printer.ScaleX(Left, vbTwips, Printer.ScaleMode), Printer.ScaleY(Top, vbTwips, Printer.ScaleMode), Printer.ScaleX(Left + width, vbTwips, Printer.ScaleMode), Printer.ScaleY(Top + height, vbTwips, Printer.ScaleMode)
    DrawTextEx Printer.hdc, Text, Len(Text), rec, DT_WORDBREAK, ByVal 0&
    
    Printer.ScaleMode = vbTwips
    
    'DoEvents
    SetRectEmpty rec
    'Let Rec = Null 'Nothing
End Sub


Sub printTextWB(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal textwidth As Double, ByVal fontname As String, ByVal fontsize As Byte, ByVal bold As Boolean, ByVal justify As Byte)
    'prints black text on white background
    printTextColor s, x, y, vbBlack, vbWhite, textwidth, fontname, fontsize, bold, justify
End Sub

Sub printText(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal textwidth As Double, ByVal fontname As String, ByVal fontsize As Byte, ByVal bold As Boolean, ByVal justify As Byte)
    'prints black text on transparent background
    printTextColor s, x, y, vbBlack, -1, textwidth, fontname, fontsize, bold, justify

End Sub


Sub printListView(ByVal lv As ListView, ByVal pageLength As Byte, ByVal LeftMargin As Double, ByVal TopMargin As Double, ByVal columnScale As Double, ByVal multiPage As Boolean)
    Dim PrinterName     As String
    Dim ColumnX(25)     As Double
    Dim numColumns      As Integer
    'Dim sourceForm      As Form
    Dim y               As Double
    Dim i, j, P, count  As Integer
    Dim temp, temp2     As String
    Dim pageStart       As Integer
    Dim pageEnd         As Integer
    Dim q               As Integer
    Dim twips_per_inch  As Double
    Dim twips_per_char  As Double
    'I set rowScale to 6.735 on March 1, 2005
    Const rowScale = 6.735 '6.647 '6.784
    twips_per_inch = 1440
    twips_per_char = twips_per_inch / 12#
    
    With Printer
        .fontname = "Courier New"
        '.Orientation = 1
        .ScaleMode = vbTwips
        .FontBold = False
        .fontsize = 10
        PrinterName = .DeviceName
    End With
    
    numColumns = lv.ColumnHeaders.count 'LV.ListItems(0).ListSubItems.Count
    ColumnX(0) = LeftMargin
    For i = 1 To numColumns
        ColumnX(i) = ColumnX(i - 1) + (lv.ColumnHeaders(i).width * columnScale)
    Next i
        
    For P = 1 To Int(lv.ListItems.count / pageLength) + 1
        If multiPage Then
            printTextWB "Page " & P & " of " & (Int(lv.ListItems.count / pageLength) + 1), 10400, 14700, 2000, "arial", 10, False, 0
            With Printer
                .fontname = "Courier New"
                .FontBold = False
                .fontsize = 10
            End With
        End If
        ' init pages
        If lv.ListItems.count <= pageLength Then
            pageStart = 1
            pageEnd = lv.ListItems.count
        Else
            If P = 1 Then
                pageStart = 1
                pageEnd = pageLength
            ElseIf P = Int(lv.ListItems.count / pageLength) + 1 Then
                pageStart = pageStart + pageLength
                pageEnd = lv.ListItems.count
            Else
                pageStart = pageStart + pageLength
                pageEnd = pageEnd + pageLength
            End If
        End If
        y = TopMargin
        
        Printer.FontBold = True
        For i = 0 To numColumns - 1
            If lv.ColumnHeaders(i + 1).width > 100 Then
                Printer.CurrentX = ColumnX(i)
                Printer.CurrentY = y - 320
                Printer.Print Left$(Trim(lv.ColumnHeaders(i + 1).Text), Int(lv.ColumnHeaders(i + 1).width * columnScale / twips_per_char))
            End If
        Next i
        Printer.FontBold = False
        'j = pageStart
        'MsgBox j
        'MsgBox pageLength
        'MsgBox j Mod pageLength
        'MsgBox TopMargin - 240 + (j Mod pageLength) * twips_per_inch / rowScale
        For j = pageStart To pageEnd
        
            y = TopMargin - 240 + (((j - 1) Mod pageLength) + 1) * twips_per_inch / rowScale
            For i = 0 To numColumns - 1
                If lv.ColumnHeaders(i + 1).width > 100 Then
                    Printer.CurrentX = ColumnX(i)
                    Printer.CurrentY = y
                    If i = 0 Then
                        temp = lv.ListItems(j).Text
                    Else
                        temp = lv.ListItems(j).SubItems(i)
                    End If
                    temp = Left$(Trim(temp), Int(lv.ColumnHeaders(i + 1).width * columnScale / twips_per_char))
                    temp2 = ""
                    If lv.ColumnHeaders(i + 1).Alignment = lvwColumnRight Then
                        For count = 1 To Int(lv.ColumnHeaders(i + 1).width * columnScale / twips_per_char) - Len(temp)
                            temp2 = " " + temp2
                        Next
                        Printer.Print temp2 & temp
                    ElseIf lv.ColumnHeaders(i + 1).Alignment = lvwColumnCenter Then
                        For count = 1 To Int((Int(lv.ColumnHeaders(i + 1).width * columnScale / twips_per_char) - Len(temp)) / 2)
                            temp2 = " " + temp2
                        Next
                        Printer.Print temp2 & temp
                    Else
                        Printer.Print temp
                    End If
                End If
            Next i
            'y = Printer.CurrentY
        Next j
        Printer.Line (LeftMargin - 30, TopMargin - 445)-(ColumnX(numColumns) - 30, TopMargin - 445)
        Printer.Line (LeftMargin - 30, TopMargin - 75)-(ColumnX(numColumns) - 30, TopMargin - 75)
        For i = 0 To numColumns
            Printer.Line (ColumnX(i) - 30, TopMargin - 445)-(ColumnX(i) - 30, pageLength * twips_per_inch / rowScale + TopMargin)
        Next i
        For i = 1 To pageLength + 1 Step 5
            Printer.Line (LeftMargin - 30, TopMargin - 240 + i * twips_per_inch / rowScale)-(ColumnX(numColumns) - 30, TopMargin - 240 + i * twips_per_inch / rowScale)
        Next i
        Printer.Line (LeftMargin - 30, TopMargin + pageLength * twips_per_inch / rowScale)-(ColumnX(numColumns) - 30, TopMargin + pageLength * twips_per_inch / rowScale)
        If multiPage Then
            If lv.ListItems.count > pageLength * P Then Printer.NewPage
            'MsgBox "end of page " & p
        Else
            If lv.ListItems.count > pageLength Then Exit Sub
        End If
    Next P
    'Set sourceForm = Nothing
    Set lv = Nothing
    'Printer.EndDoc
End Sub

Sub printFlexGrid(ByRef printObj As Variant, ByVal fg As MSFlexGrid, ByVal LeftMargin As Double, ByVal TopMargin As Double, ByVal columnScale As Double)
    Dim PrinterName     As String
    Dim r               As Long
    Dim c               As Long
    Dim x               As Long
    Dim y               As Long
    Dim twips_per_inch  As Double
    Dim twips_per_line  As Double
    
    'I changed rowScale to 6.5 on June 21, 2006
    'For this procedure rowScale is not so vitally precise as it is in the listview procedure
    Const rowScale = 6.5 '6.735
    twips_per_inch = 1440
    twips_per_line = twips_per_inch / rowScale
        
    With printObj
        .fontname = "Courier New"
        .ScaleMode = vbTwips
        .FontBold = False
        .fontsize = 10
        PrinterName = .DeviceName
    End With
    
    printObj.Line (LeftMargin, TopMargin)-(LeftMargin + fg.width * columnScale, TopMargin + fg.height), vbWhite, BF
    printObj.Line (LeftMargin, TopMargin)-(LeftMargin + fg.width * columnScale, TopMargin + fg.height), 0, B
    x = LeftMargin
    For c = 0 To fg.Cols - 1
        x = x + fg.ColWidth(c) * columnScale
        printObj.Line (x, TopMargin)-(x, TopMargin + fg.height), 0
    Next c
    
    y = TopMargin '+ 45 'twips_per_line - 45
    For r = 0 To fg.rows - 1
        x = LeftMargin + 45
        fg.row = r
        For c = 0 To fg.Cols - 1
            fg.col = c
            If fg.ColWidth(c) > 0 Then
                printObj.CurrentY = y
                printObj.CurrentX = x
                printObj.Print fg.Text
            End If
            
            x = x + fg.ColWidth(c) * columnScale
        Next c
        y = y + twips_per_line
        printObj.Line (LeftMargin, y)-(LeftMargin + fg.width * columnScale, y), 0
        If y >= TopMargin + fg.height Then Exit For
    Next r
End Sub
Sub printLineGraph(ByRef printObj As Variant, ByRef Chart As MSChart, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double)
    Dim i, j As Long
    Dim column As Long
    Dim row As Long
    Dim rowWidth As Long
    Dim col As Long
    Dim max As Double
    
    max = Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum
    
    printObj.ScaleMode = vbTwips
    printObj.forecolor = vbBlack
    printObj.FontBold = False
    printObj.fontname = "Times New Roman"
    printObj.fontsize = 8
    
    printObj.CurrentX = x - 400
    printObj.CurrentY = y - 70
    printObj.Print Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum
    printObj.CurrentX = x - 400
    printObj.CurrentY = y + (height / 2) - 70
    printObj.Print Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum / 2
    printObj.CurrentX = x - 400
    printObj.CurrentY = y + height - 70
    printObj.Print "0"
    
    rowWidth = width / Chart.RowCount
    printObj.FillColor = vbWhite
    printObj.Line (x, y)-(x + width, y + height), vbWhite, BF
    printObj.Line (x, y)-(x + width, y + height), 0, B
    printObj.Line (x, y + (height / 2))-(x + width, y + (height / 2)), 0
    
    For i = 0 To Chart.RowCount - 1
        Chart.row = i + 1
        printObj.CurrentX = x + ((i + 0.25) * rowWidth)
        printObj.CurrentY = y + height + 100
        printObj.Print Chart.RowLabel
        printObj.Line (x + i * rowWidth, y)-(x + i * rowWidth, y + height), 0
    Next i
    
    printObj.DrawWidth = 5
    For i = 1 To Chart.ColumnCount
        Chart.column = i
        col = CLng(Chart.Plot.SeriesCollection(i).Pen.VtColor.Red) + CLng(Chart.Plot.SeriesCollection(i).Pen.VtColor.Green) * 256 + CLng(Chart.Plot.SeriesCollection(i).Pen.VtColor.Blue) * 65536
        j = 0
        Chart.row = 1
        printObj.PSet (x + ((j + 0.5) * rowWidth), y + height - height * (Chart.Data / max)), col
        For j = 1 To Chart.RowCount - 1
            Chart.row = j + 1
            printObj.Line -(x + ((j + 0.5) * rowWidth), y + height - height * (Chart.Data / max)), col
        Next j
    Next i
    printObj.DrawWidth = 1
End Sub



Sub formPrint(f As Form, Optional ByVal horizOffset As Long = 50, Optional ByVal vertiOffset As Long = 50)

' +-----------------------------------+
' |                                   |
' |     M A R K   M E R C E R ' S     |
' |      F O R M   P R I N T E R      |
' |  WEB: www.woodjoint.ca/mark       |
' |  EMAIL: mark@creatiworks.ca       |
' |  DOCUMENTATION: Under Main Module |
' |                                   |
' +-----------------------------------+

    If f.ScaleMode <> vbTwips Then
        MsgBox "Form scalemode is not twips, please change this before printing"
        Exit Sub
    End If
    
    Dim c As Control
    Dim cont As Control
    Dim forecol As Long ' forecolor
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim invisible As Boolean
    Dim rows As Long
    Dim s As String
    Dim str As String
    Dim line As Long
    Dim r As Double ' radius of pie chart
    Dim L As Long ' Master Left location of a control
    Dim T As Long ' Master Top location of a control
    Dim adjust As Long ' can be used to adjust the size of items in a control.
    Dim rec As RECT
    Dim txt As String
    
    Printer.ScaleMode = vbTwips
    Printer.forecolor = 0
    If f.Picture = 0 Then
        Printer.PaintPicture f.Image, horizOffset, vertiOffset, Printer.ScaleX(f.ScaleWidth, vbTwips, Printer.ScaleMode), Printer.ScaleY(f.ScaleHeight, vbTwips, Printer.ScaleMode)
    Else
        Printer.PaintPicture f.Picture, horizOffset, vertiOffset, Printer.ScaleX(f.ScaleWidth, vbTwips, Printer.ScaleMode), Printer.ScaleY(f.ScaleHeight, vbTwips, Printer.ScaleMode)
    End If
    Printer.FillStyle = vbFSSolid
    ' the following two ifs clip a forms bkg picture
    'If f.height < f.Picture.height Then
    '    Printer.Line (-30 + horizOffset, f.height + vertiOffset)-(f.Picture.width + 30 + horizOffset, f.Picture.height + 30 + vertiOffset), vbWhite, BF
    'End If
    'If f.width < f.Picture.width Then
    '    Printer.Line (f.width + horizOffset, -30 + vertiOffset)-(f.Picture.width + 30 + horizOffset, f.Picture.height + 30 + vertiOffset), vbWhite, BF
    'End If
    
    For Each c In f.Controls
        invisible = False
        If Not TypeOf c Is Timer And _
            Not TypeOf c Is ImageList And _
            Not TypeOf c Is CommonDialog And _
            Not TypeOf c Is Menu Then
            Printer.FillColor = vbWhite
            
            If TypeOf c Is line Then
                L = 0
                T = 0
            Else
                L = c.Left
                T = c.Top
            End If
            If c.Container.name <> f.name Then
                Set cont = c.Container
                Do
                    If cont.Visible Then
                        L = L + cont.Left
                        T = T + cont.Top
                        If cont.Container.name = f.name Then
                            Exit Do
                        Else
                            Set cont = cont.Container
                        End If
                    Else
                        invisible = True
                        Exit Do
                    End If
                Loop
                
            End If
        
            If Not invisible Then
                If TypeOf c Is Frame Then
                    If c.Visible Then
                        Printer.Font = c.Font
                        Printer.fontsize = c.fontsize
                        Printer.Line (L + horizOffset, T + vertiOffset + 30)-(c.width + L + horizOffset, c.height + T + vertiOffset - 15), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset + (c.Font.Size * 15))-(c.width + L + horizOffset, c.height + T + vertiOffset - 15), 0, B
                        printTextWB c.Caption, L + horizOffset + 60, T + vertiOffset + 30, Printer.textwidth(c.Caption) + 30, c.Font.name, c.Font.Size, c.Font.bold, 0
                    End If
                End If
                If TypeOf c Is textBox Then
                    If c.Visible Then
                        Printer.FontItalic = c.FontItalic
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        If c.BorderStyle = 1 Then
                            Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        End If
                        
                        forecol = c.forecolor
                        If forecol < 0 Then forecol = vbBlack
                        
                        Printer.Font = c.Font
                        Printer.fontsize = c.fontsize
                        If c.MultiLine Then
                            Printer.ScaleMode = vbPixels
                            
                            Printer.forecolor = forecol
                            Printer.FontBold = c.FontBold
                            Printer.fontname = c.fontname
                            Printer.fontsize = c.fontsize
                            
                            txt = Replace(c.Text, " & ", " && ")
                            'SetRect Rec, (horizOffset + L) / RECT_SCALE + 15, (vertiOffset + T) / RECT_SCALE + 15, (horizOffset + L + c.width) / RECT_SCALE, (vertiOffset + T + c.height) / RECT_SCALE
                            SetRect rec, Printer.ScaleX((horizOffset + L) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleX((horizOffset + L + c.width), vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T + c.height), vbTwips, Printer.ScaleMode)
                            DrawTextEx Printer.hdc, txt, Len(txt), rec, DT_WORDBREAK, ByVal 0&
                            
                            Printer.ScaleMode = vbTwips
                        Else
                            printTextColor c.Text, L + horizOffset + 15, T + vertiOffset + 15, forecol, vbWhite, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                        End If
                    End If
                ElseIf TypeOf c Is ComboBox Then
                    If c.Visible Then
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        printTextWB c.Text, L + horizOffset + 15, T + vertiOffset + ((c.height - 30) / 2#) - (c.Font.Size * 7.5), c.width - 30, c.Font.name, c.Font.Size, c.Font.bold, 0
                    End If
                ElseIf TypeOf c Is CommandButton Then
                    If c.Visible Then
                        Printer.FontItalic = c.FontItalic
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        printTextWB Replace(c.Caption, "&&", "&"), L + horizOffset + 15, T + vertiOffset + (c.height / 2#) - (c.Font.Size * 7.5), c.width - 30, c.fontname, c.Font.Size, c.Font.bold, 2
                    End If
                ElseIf TypeOf c Is Shape Then
                    If c.Shape <> 4 And c.Visible Then
                        'Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), greyScale(c.FillColor), BF ' greyscale version for non color printing
                        Printer.FillStyle = c.FillStyle
                        If c.FillStyle = 0 Then '(solid)
                            Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), c.FillColor, BF
                        Else
                            Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), c.BorderColor, B
                        End If
                        Printer.FillStyle = vbFSSolid
                    End If
                ElseIf TypeOf c Is line Then
                    If c.Visible Then
                        Printer.DrawWidth = 5
                        Printer.Line (L + c.x1 + horizOffset, T + c.Y1 + vertiOffset)-(L + c.x2 + horizOffset, T + c.Y2 + vertiOffset), c.BorderColor
                        Printer.DrawWidth = 1
                    End If
                ElseIf TypeOf c Is Label Then
                    If c.Visible Then
                        Printer.FontItalic = c.FontItalic
                        forecol = c.forecolor
                        If forecol < 0 Then forecol = vbBlack
                        
                        Printer.Font = c.Font
                        Printer.fontsize = c.fontsize
                        
                        txt = Replace(c.Caption, " && ", " & ")
                        If Printer.textwidth(txt) > c.width Then 'multiple lines
                            Printer.ScaleMode = vbPixels
                            Printer.forecolor = forecol
                            Printer.FontBold = c.FontBold
                            Printer.fontname = c.fontname
                            SetRect rec, Printer.ScaleX((horizOffset + L) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleX((horizOffset + L + c.width), vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T + c.height), vbTwips, Printer.ScaleMode)
                            DrawTextEx Printer.hdc, txt, Len(txt), rec, DT_WORDBREAK, ByVal 0&
                            Printer.ScaleMode = vbTwips
                        Else
                            If c.BackStyle = 1 Then ' opaque background
                                printTextColor txt, L + horizOffset + 15, T + vertiOffset, c.forecolor, c.backcolor, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                            Else ' transparent background
                                printTextColor txt, L + horizOffset + 15, T + vertiOffset, c.forecolor, -1, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                            End If
                            'printText txt, L + horizOffset + 15, T + vertiOffset + (c.height / 2#) - (c.Font.size * 7.5) - 15, c.width - 30, c.fontname, c.Font.size, c.Font.bold, c.Alignment
                        End If
                    End If
                ElseIf TypeOf c Is OptionButton Then
                    If c.Visible Then
                        Printer.FillColor = vbWhite
                        Printer.Circle (L + horizOffset + 100, T + vertiOffset + (c.height / 2#)), 90, vbBlack
                        If c.value Then
                            Printer.FillColor = vbBlack
                            Printer.Circle (L + horizOffset + 100, T + vertiOffset + (c.height / 2#)), 50, vbBlack
                        End If
                        printText c.Caption, L + horizOffset + 220, T + vertiOffset + (c.height / 2#) - (c.Font.Size * 7.5) - 15, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                    End If
                ElseIf TypeOf c Is CheckBox Then
                    If c.Visible Then
                        adjust = 80 ' half of the width (or) height of the box
                        Printer.Line (L + horizOffset, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) + adjust), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) + adjust), vbBlack, B
                        'Printer.Line (L + horizOffset + adjust, T + vertiOffset + adjust)-(L + horizOffset + c.height - adjust, T + vertiOffset + c.height - adjust), vbBlack, B
                        If c.value = 1 Then
                            Printer.Line (L + horizOffset, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) + adjust), vbBlack
                            Printer.Line (L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset, T + vertiOffset + (c.height / 2#) + adjust), vbBlack
                            'Printer.Line (L + horizOffset + adjust, T + vertiOffset + adjust)-(L + horizOffset + c.height - adjust, T + vertiOffset + c.height - adjust), vbBlack
                            'Printer.Line (L + horizOffset + c.height - adjust, T + vertiOffset + adjust)-(L + horizOffset + adjust, T + vertiOffset + c.height - adjust), vbBlack
                        End If
                        printText c.Caption, L + horizOffset + adjust * 2 + 60, T + vertiOffset + (c.height / 2#) - (c.Font.Size * 7.5) - 15, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, 0
                    End If
                ElseIf TypeOf c Is PictureBox Then
                    If c.Visible Then
                        Printer.PaintPicture c.Image, L + horizOffset, T + vertiOffset, c.width, c.height
                    End If
                ElseIf TypeOf c Is ListView Then
                    If c.Visible Then
                        rows = Int(c.height / 225) - 1
                        If rows < 1 Then rows = 1
                        printListView c, rows, L + horizOffset, 450 + T + vertiOffset, 1, False
                    End If
                    'printListView c, 27, 1000, 1000, 1, False
                ElseIf TypeOf c Is dtPicker Then
                    If c.Visible Then
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        printText Format(c.value, "mmm dd, yyyy"), L + horizOffset + 15, T + vertiOffset + ((c.height - 30) / 2#) - (c.Font.Size * 7.5), c.width - 30, "arial", c.Font.Size, c.Font.bold, 0
                    End If
                ElseIf TypeOf c Is MSChart Then
                    If c.Visible And c.chartType = 3 Then 'line
                        If c.width > 850 And c.height > 600 Then
                            printLineGraph Printer, c, L + horizOffset + 650, T + vertiOffset + 250, c.width - 850, c.height - 600
                        End If
                    ElseIf c.Visible And c.chartType = 1 Then 'bar
                        If c.width > 850 And c.height > 600 Then
                            printBarGraph f, Printer, c, L + horizOffset + 650, T + vertiOffset + 250, c.width - 850, c.height - 600, 30, "Bar Graph", True
                        End If
                    ElseIf c.Visible And c.chartType = 14 Then 'pie
                        If c.width > 850 And c.height > 600 Then
                            r = (c.width * 3 + c.height) / 8#
                            printPieChart Printer, c, L + horizOffset + r, T + vertiOffset + r, r, 1
                        End If
                    End If
                ElseIf TypeOf c Is MSFlexGrid Then
                    If c.Visible Then
                        printFlexGrid Printer, c, L + horizOffset, T + vertiOffset, 1
                    End If
                'ElseIf TypeOf c Is Bezier Then
                '    If c.Visible Then
                '        MsgBox "About to print a bezier graph... pause and see if enddoc has been called"
                '        c.drawGraph True
                '        MsgBox "Just Finished printing a bezier graph... pause and see if enddoc has been called"
                '    End If
                End If
            End If ' inherited invisiblity
        End If ' non-visible controls
    Next c ' each control
    
    Printer.EndDoc
    
    Set c = Nothing
    Set cont = Nothing
    SetRectEmpty rec
    'Set Rec = Nothing
End Sub

Sub framePrint(f As Form, FR As Frame, Optional ByVal horizOffset As Long = 50, Optional ByVal vertiOffset As Long = 50)

' +-----------------------------------+
' |                                   |
' |     M A R K   M E R C E R ' S     |
' |      F O R M   P R I N T E R      |
' |  WEB: www.woodjoint.ca/mark       |
' |  EMAIL: mark@creatiworks.ca       |
' |  DOCUMENTATION: Under Main Module |
' |                                   |
' +-----------------------------------+

'Oct 17, 2012 - Added funtionality to break down printing into pages.  If controls are below a certain point it holds them until the subsequent page.

    If f.ScaleMode <> vbTwips Then
        MsgBox "Form scalemode is not twips, please change this before printing"
        Exit Sub
    End If
    
    Dim c As Control
    Dim cont As Control
    Dim forecol As Long ' forecolor
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim invisible As Boolean
    Dim rows As Long
    Dim s As String
    Dim str As String
    Dim line As Long
    Dim r As Double ' radius of pie chart
    Dim L As Long ' Master Left location of a control
    Dim T As Long ' Master Top location of a control
    Dim b As Long ' Master Bottom location of the control
    Dim y As Long ' Actual Print Coordinate including the passed in offset variable.
    Dim adjust As Long ' can be used to adjust the size of items in a control.
    Dim rec As RECT
    Dim txt As String
    Dim skip As Boolean
    Dim pageheight As Long
    Dim page As Long
    Dim left_to_print As Boolean
    Dim print_this_control As Boolean
    
    Printer.ScaleMode = vbTwips
    Printer.forecolor = 0
    pageheight = Printer.ScaleHeight - 400
    page = 1
    
    If f.Picture = 0 Then
        Printer.PaintPicture f.Image, horizOffset, vertiOffset, Printer.ScaleX(f.ScaleWidth, vbTwips, Printer.ScaleMode), Printer.ScaleY(f.ScaleHeight, vbTwips, Printer.ScaleMode)
    Else
        Printer.PaintPicture f.Picture, horizOffset, vertiOffset, Printer.ScaleX(f.ScaleWidth, vbTwips, Printer.ScaleMode), Printer.ScaleY(f.ScaleHeight, vbTwips, Printer.ScaleMode)
    End If
    Printer.FillStyle = vbFSSolid
    ' the following two ifs clip a forms bkg picture
    'If f.height < f.Picture.height Then
    '    Printer.Line (-30 + horizOffset, f.height + vertiOffset)-(f.Picture.width + 30 + horizOffset, f.Picture.height + 30 + vertiOffset), vbWhite, BF
    'End If
    'If f.width < f.Picture.width Then
    '    Printer.Line (f.width + horizOffset, -30 + vertiOffset)-(f.Picture.width + 30 + horizOffset, f.Picture.height + 30 + vertiOffset), vbWhite, BF
    'End If
    
    Do
    left_to_print = False
    For Each c In f.Controls
        invisible = False
        If Not TypeOf c Is Timer And _
            Not TypeOf c Is ImageList And _
            Not TypeOf c Is CommonDialog And _
            Not TypeOf c Is Menu Then
            Printer.FillColor = vbWhite
            
            skip = False
            If TypeOf c Is line Then
                L = 0
                T = 0
                skip = True
            Else
                L = c.Left
                T = c.Top
            End If
            
            If c.Container.name <> f.name Then
                Set cont = c.Container
                Do
                    If cont.Container.name = f.name Then
                        skip = True
                        Exit Do
                    Else
                        If cont.Visible Then
                            If cont.name = FR.name Then
                                Exit Do
                            Else
                                L = L + cont.Left
                                T = T + cont.Top
                                Set cont = cont.Container
                            End If
                        Else
                            invisible = True
                            Exit Do
                        End If
                    End If
                Loop
            Else
                skip = True
            End If
            
            
            
            
            
            y = T + vertiOffset
            If Not skip Then b = y + c.height
            
            'If Not invisible And Not skip Then print_this_control = True
            If (y >= pageheight * (page - 1) And b < pageheight * page) Or (y < 0 And page = 1) Then print_this_control = True
            If invisible Then print_this_control = False
            If skip Then print_this_control = False
            If b >= pageheight * page And Not invisible And Not skip Then
                left_to_print = True
            End If
            If y > pageheight Then y = y Mod pageheight
            'If B >= pageheight * page Then MsgBox c.name
            T = y - vertiOffset
            
            
            If print_this_control Then
                If TypeOf c Is Frame Then
                    If c.Visible Then
                        Printer.Font = c.Font
                        Printer.fontsize = c.fontsize
                        Printer.Line (L + horizOffset, T + vertiOffset + 30)-(c.width + L + horizOffset, c.height + T + vertiOffset - 15), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset + (c.Font.Size * 15))-(c.width + L + horizOffset, c.height + T + vertiOffset - 15), 0, B
                        printTextWB c.Caption, L + horizOffset + 60, T + vertiOffset + 30, Printer.textwidth(c.Caption) + 30, c.Font.name, c.Font.Size, c.Font.bold, 0
                    End If
                End If
                If TypeOf c Is textBox Then
                    If c.Visible Then
                        Printer.FontItalic = c.FontItalic
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        
                        forecol = c.forecolor
                        If forecol < 0 Then forecol = vbBlack
                        
                        Printer.Font = c.Font
                        Printer.fontsize = c.fontsize
                        If c.MultiLine Then
                            Printer.ScaleMode = vbPixels
                            
                            Printer.forecolor = forecol
                            Printer.FontBold = c.FontBold
                            Printer.fontname = c.fontname
                            Printer.fontsize = c.fontsize
                            
                            txt = Replace(c.Text, " & ", " && ")
                            'SetRect Rec, (horizOffset + L) / RECT_SCALE + 15, (vertiOffset + T) / RECT_SCALE + 15, (horizOffset + L + c.width) / RECT_SCALE, (vertiOffset + T + c.height) / RECT_SCALE
                            SetRect rec, Printer.ScaleX((horizOffset + L) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleX((horizOffset + L + c.width), vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T + c.height), vbTwips, Printer.ScaleMode)
                            DrawTextEx Printer.hdc, txt, Len(txt), rec, DT_WORDBREAK, ByVal 0&
                            
                            Printer.ScaleMode = vbTwips
                        Else
                            printTextColor c.Text, L + horizOffset + 15, T + vertiOffset + 15, forecol, vbWhite, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                        End If
                    End If
                ElseIf TypeOf c Is ComboBox Then
                    If c.Visible Then
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        printTextWB c.Text, L + horizOffset + 15, T + vertiOffset + ((c.height - 30) / 2#) - (c.Font.Size * 7.5), c.width - 30, c.Font.name, c.Font.Size, c.Font.bold, 0
                    End If
                ElseIf TypeOf c Is CommandButton Then
                    If c.Visible Then
                        Printer.FontItalic = c.FontItalic
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        printTextWB Replace(c.Caption, "&&", "&"), L + horizOffset + 15, T + vertiOffset + (c.height / 2#) - (c.Font.Size * 7.5), c.width - 30, c.fontname, c.Font.Size, c.Font.bold, 2
                    End If
                ElseIf TypeOf c Is Shape Then
                    If c.Shape <> 4 And c.Visible Then
                        'Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), greyScale(c.FillColor), BF ' greyscale version for non color printing
                        Printer.FillStyle = c.FillStyle
                        If c.FillStyle = 0 Then '(solid)
                            Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), c.FillColor, BF
                        Else
                            Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), c.BorderColor, B
                        End If
                        Printer.FillStyle = vbFSSolid
                    End If
                ElseIf TypeOf c Is line Then
                    If c.Visible Then
                        Printer.DrawWidth = 5
                        Printer.Line (L + c.x1 + horizOffset, T + c.Y1 + vertiOffset)-(L + c.x2 + horizOffset, T + c.Y2 + vertiOffset), c.BorderColor
                        Printer.DrawWidth = 1
                    End If
                ElseIf TypeOf c Is Label Then
                    If c.Visible Then
                        Printer.FontItalic = c.FontItalic
                        forecol = c.forecolor
                        If forecol < 0 Then forecol = vbBlack
                        
                        Printer.Font = c.Font
                        Printer.fontsize = c.fontsize
                        
                        txt = Replace(c.Caption, " && ", " & ")
                        If Printer.textwidth(txt) > c.width Then 'multiple lines
                            Printer.ScaleMode = vbPixels
                            Printer.forecolor = forecol
                            Printer.FontBold = c.FontBold
                            Printer.fontname = c.fontname
                            SetRect rec, Printer.ScaleX((horizOffset + L) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T) + 15, vbTwips, Printer.ScaleMode), Printer.ScaleX((horizOffset + L + c.width), vbTwips, Printer.ScaleMode), Printer.ScaleY((vertiOffset + T + c.height), vbTwips, Printer.ScaleMode)
                            DrawTextEx Printer.hdc, txt, Len(txt), rec, DT_WORDBREAK, ByVal 0&
                            Printer.ScaleMode = vbTwips
                        Else
                            If c.BackStyle = 1 Then ' opaque background
                                printTextColor txt, L + horizOffset + 15, T + vertiOffset, c.forecolor, c.backcolor, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                            Else ' transparent background
                                printTextColor txt, L + horizOffset + 15, T + vertiOffset, c.forecolor, -1, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                            End If
                            'printText txt, L + horizOffset + 15, T + vertiOffset + (c.height / 2#) - (c.Font.size * 7.5) - 15, c.width - 30, c.fontname, c.Font.size, c.Font.bold, c.Alignment
                        End If
                    End If
                ElseIf TypeOf c Is OptionButton Then
                    If c.Visible Then
                        Printer.FillColor = vbWhite
                        Printer.Circle (L + horizOffset + 100, T + vertiOffset + (c.height / 2#)), 90, vbBlack
                        If c.value Then
                            Printer.FillColor = vbBlack
                            Printer.Circle (L + horizOffset + 100, T + vertiOffset + (c.height / 2#)), 50, vbBlack
                        End If
                        printText c.Caption, L + horizOffset + 220, T + vertiOffset + (c.height / 2#) - (c.Font.Size * 7.5) - 15, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, c.Alignment
                    End If
                ElseIf TypeOf c Is CheckBox Then
                    If c.Visible Then
                        adjust = 80 ' half of the width (or) height of the box
                        Printer.Line (L + horizOffset, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) + adjust), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) + adjust), vbBlack, B
                        'Printer.Line (L + horizOffset + adjust, T + vertiOffset + adjust)-(L + horizOffset + c.height - adjust, T + vertiOffset + c.height - adjust), vbBlack, B
                        If c.value = 1 Then
                            Printer.Line (L + horizOffset, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) + adjust), vbBlack
                            Printer.Line (L + horizOffset + adjust * 2, T + vertiOffset + (c.height / 2#) - adjust)-(L + horizOffset, T + vertiOffset + (c.height / 2#) + adjust), vbBlack
                            'Printer.Line (L + horizOffset + adjust, T + vertiOffset + adjust)-(L + horizOffset + c.height - adjust, T + vertiOffset + c.height - adjust), vbBlack
                            'Printer.Line (L + horizOffset + c.height - adjust, T + vertiOffset + adjust)-(L + horizOffset + adjust, T + vertiOffset + c.height - adjust), vbBlack
                        End If
                        printText c.Caption, L + horizOffset + adjust * 2 + 60, T + vertiOffset + (c.height / 2#) - (c.Font.Size * 7.5) - 15, c.width - 30, c.fontname, c.Font.Size, c.Font.bold, 0
                    End If
                ElseIf TypeOf c Is PictureBox Then
                    If c.Visible Then
                        Printer.PaintPicture c.Image, L + horizOffset, T + vertiOffset, c.width, c.height
                    End If
                ElseIf TypeOf c Is ListView Then
                    If c.Visible Then
                        rows = Int(c.height / 225) - 1
                        If rows < 1 Then rows = 1
                        printListView c, rows, L + horizOffset, 450 + T + vertiOffset, 1, False
                    End If
                    'printListView c, 27, 1000, 1000, 1, False
                ElseIf TypeOf c Is dtPicker Then
                    If c.Visible Then
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), vbWhite, BF
                        Printer.Line (L + horizOffset, T + vertiOffset)-(c.width + L + horizOffset, c.height + T + vertiOffset), 0, B
                        printText Format(c.value, "mmm dd, yyyy"), L + horizOffset + 15, T + vertiOffset + ((c.height - 30) / 2#) - (c.Font.Size * 7.5), c.width - 30, "arial", c.Font.Size, c.Font.bold, 0
                    End If
                ElseIf TypeOf c Is MSChart Then
                    If c.Visible And c.chartType = 3 Then 'line
                        If c.width > 850 And c.height > 600 Then
                            printLineGraph Printer, c, L + horizOffset + 650, T + vertiOffset + 250, c.width - 850, c.height - 600
                        End If
                    ElseIf c.Visible And c.chartType = 1 Then 'bar
                        If c.width > 850 And c.height > 600 Then
                            printBarGraph f, Printer, c, L + horizOffset + 650, T + vertiOffset + 250, c.width - 850, c.height - 600, 30, "Bar Graph", True
                        End If
                    ElseIf c.Visible And c.chartType = 14 Then 'pie
                        If c.width > 850 And c.height > 600 Then
                            r = (c.width * 3 + c.height) / 8#
                            printPieChart Printer, c, L + horizOffset + r, T + vertiOffset + r, r, 1
                        End If
                    End If
                ElseIf TypeOf c Is MSFlexGrid Then
                    If c.Visible Then
                        printFlexGrid Printer, c, L + horizOffset, T + vertiOffset, 1
                    End If
                'ElseIf TypeOf c Is Bezier Then
                '    If c.Visible Then
                '        MsgBox "About to print a bezier graph... pause and see if enddoc has been called"
                '        c.drawGraph True
                '        MsgBox "Just Finished printing a bezier graph... pause and see if enddoc has been called"
                '    End If
                End If
            End If ' inherited invisiblity
        End If ' non-visible controls
    Next c ' each control
    
    page = page + 1
    Printer.NewPage
    Loop Until left_to_print = False
    
    Printer.EndDoc
    
    Set c = Nothing
    Set cont = Nothing
    SetRectEmpty rec
    'Set Rec = Nothing
End Sub






