'==============================================================
' Title: Power Point Generator
' Goal: To generate PowerPoint Presentations from Excel Spreadsheets
' Author: Paulo Ramos
' Date: Jun 2014
'==============================================================

'==============================================================
Sub Run_Batch()
    ORIG$ = "PARAMETERS"
    
    li = Sheets(ORIG$).Cells(4, 3).Value
    lf = Sheets(ORIG$).Cells(5, 3).Value
    
    For i = li To lf
     Call select_range(i)
    Next i
    
    Sheets(ORIG$).Select
    Sheets(ORIG$).Cells(1, 1).Select
End Sub

'==============================================================
Sub select_range(ct)
    '-- Goal: To select the range or the Chart(Graphic) to be copied --
    ORIG$ = "PARAMETERS"
    SS$ = Sheets(ORIG$).Cells(ct, 2).Text ' Name of the Spreadsheet
    RG$ = Sheets(ORIG$).Cells(ct, 3).Text ' Range
    GF$ = Sheets(ORIG$).Cells(ct, 4).Text  ' Graphic name
    
    t1 = Sheets(ORIG$).Cells(ct, 11).Value   ' Top
    l1 = Sheets(ORIG$).Cells(ct, 12).Value   ' left
    h1 = Sheets(ORIG$).Cells(ct, 13).Value   ' Height
    w1 = Sheets(ORIG$).Cells(ct, 14).Value   ' Width
    
    '----- Update Parameters ----------
    C1$ = Sheets(ORIG$).Cells(ct, 6).Text  ' cell #1
    P1$ = Sheets(ORIG$).Cells(ct, 7).Text  ' Parameter #1
    
    C2$ = Sheets(ORIG$).Cells(ct, 8).Text  ' cell #2
    P2$ = Sheets(ORIG$).Cells(ct, 9).Text  ' Parameter #2
    
    If C1$ <> "" Then
        Sheets(SS$).Select
        Range(C1$ & ":" & C1$) = P1$
    End If
    If C2$ <> "" Then
        Sheets(SS$).Select
        Range(C2$ & ":" & C2$) = P2$
    End If
    
    '----------------------------------
    Title$ = Sheets(ORIG$).Cells(ct, 15).Text   ' Width
    
    If RG$ <> "" Then
        ' Copying data from Excel
    Sheets(SS$).Select
        Range(RG$).Select
        Selection.Copy
        Call CreateSlide(t1, l1, h1, w1, Title$, ct)
    Else
        ' If it is a chart (graphic)
        If GF$ <> "" Then
            Sheets(SS$).Select
            ActiveSheet.ChartObjects(GF$).Activate
            Application.CutCopyMode = False
            ActiveChart.ChartArea.Copy
            Call CreateSlide(t1, l1, h1, w1, Title$, ct)
        End If
    End If
End Sub

'==============================================================

Sub CreateSlide(t1, l1, h1, w1, Title$, lx)
    
    Dim PowerPointConn As Object
    Dim CurrentChart As Excel.Chart
    Dim SlideNumber As Integer, INI_File As String
    
    
    On Error Resume Next
    Set PowerPointConn = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If PowerPointConn Is Nothing Then Set PowerPointConn = CreateObject("PowerPoint.Application")
    If PowerPointConn.Presentations.Count = 0 Then PowerPointConn.Presentations.Add msoTrue
    PowerPointConn.Visible = msoTrue
    
    SlideNumber = PowerPointConn.ActivePresentation.Slides.Count + 1
    
    PowerPointConn.ActivePresentation.Slides.Add Index:=SlideNumber, Layout:=11 ' ppLayoutTitleOnly = 11
    
    
    PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(1).TextFrame.TextRange.Text = Title$
    
    
    PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes.PasteSpecial (2)
    
    
    ' *********** Location ****************
    PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(2).Top = Int(t1)
    PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(2).Left = Int(l1)
    ' *********** Size ********************
    
    PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(2).Width = Int(w1)
    PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(2).Height = Int(h1)
    
    
    Call batch_create_hyperlink(lx, SlideNumber)
    
    Call format_title(Title$, SlideNumber)
    
    Set PowerPointConn = Nothing

End Sub



Sub batch_create_hyperlink(lx, SlideNumber)
'lx = 11

    Dim hpl(4) As String
    ORIG$ = "PARAMETERS"
    
    hp_column = 24 ' To check in the parameters spreadsheet
    
    
    dsl = Val(Sheets(ORIG$).Cells(8, hp_column).Value)
    
    
    cx = 3
    For i = hp_column To hp_column + 50
    
        hpr$ = Sheets(ORIG$).Cells(lx, i).Text
        If hpr$ = "" Then GoTo out_hp
        hpr$ = hpr$ & "#"
        '==================Extracting Data ===========================
        For j = 0 To 4 ' cleaning fields
            hpl(j) = ""
        Next j
        
        ct = 1
        For j = 1 To Len(hpr$)
            A$ = Right$(Left$(hpr$, j), 1)
            If A$ = "#" Then
                A$ = ""
                hpl(ct) = B$
                ct = ct + 1
                B$ = ""
            End If
            B$ = B$ & A$
        Next j
        'MsgBox (hpl(1) & " - " & hpl(2) & " - " & hpl(3) & " - " & hpl(4))
        
        hp_name$ = hpl(1)
        hp_link = Val(hpl(2)) + dsl
        hp_top = Val(hpl(3))
        hp_left = Val(hpl(4))
        Call Create_Hiperlink(SlideNumber, hp_name$, hp_link, hp_top, hp_left, cx)
        
        '==============================================================
        cx = cx + 1 ' control of shapes
    Next i
    
out_hp:
    Call batch_create_textbox(lx, SlideNumber, cx)
    
End Sub

Sub batch_create_textbox(lx, SlideNumber, cx)
    'lx = 11
    
    Dim hpl(4) As String
    ORIG$ = "PARAMETERS"
    
    hp_column = 17 ' To check in the parameters spreadsheet
    
    
    dsl = Val(Sheets(ORIG$).Cells(8, hp_column).Value)
    
    
    'cx = 3
    For i = hp_column To hp_column + 4
    
        hpr$ = Sheets(ORIG$).Cells(lx, i).Text
        If hpr$ = "" Then Exit Sub
        hpr$ = hpr$ & "#"
        '==================Extracting Data ===========================
        For j = 0 To 4 ' cleaning fields
            hpl(j) = ""
        Next j
        
        ct = 1
        For j = 1 To Len(hpr$)
            A$ = Right$(Left$(hpr$, j), 1)
            If A$ = "#" Then
                A$ = ""
                hpl(ct) = B$
                ct = ct + 1
                B$ = ""
            End If
            B$ = B$ & A$
        Next j
        'MsgBox (hpl(1) & " - " & hpl(2) & " - " & hpl(3) & " - " & hpl(4))
        
        hp_name$ = hpl(1)
        hp_link = 0
        hp_top = Val(hpl(2))
        hp_left = Val(hpl(3))
        Call Create_TextBox(SlideNumber, hp_name$, hp_link, hp_top, hp_left, cx)
        
        '==============================================================
        cx = cx + 1 ' control of shapes
    Next i

End Sub


Sub Create_Hiperlink(SlideNumber, hp_name$, hp_link, hp_top, hp_left, cx)

        Sheets("PARAMETERS").Select
        ActiveSheet.Shapes.Range(Array("CaixaDeTexto 13")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = hp_name$
        
        
        ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="file:///#" & Str$(hp_link)
        ActiveSheet.Shapes.Range(Array("CaixaDeTexto 13")).Select
        Selection.Copy


        '===================== Power Point Activation =================
        Dim PowerPointConn As Object
        Dim CurrentChart As Excel.Chart
        Set PowerPointConn = GetObject(, "PowerPoint.Application")
        PowerPointConn.Visible = msoTrue

        PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes.PasteSpecial (0)


        ' *********** Location ****************
        PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(cx).Top = Int(hp_top)
        PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(cx).Left = Int(hp_left)
        ' *********** Size ********************

        'PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(cx).Width = Int(w1)
        'PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(cx).Height = Int(h1)

End Sub
Sub Create_TextBox(SlideNumber, hp_name$, hp_link, hp_top, hp_left, cx)

        Sheets("PARAMETERS").Select
        ActiveSheet.Shapes.Range(Array("Rectangle 31")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = hp_name$
        
        
        ActiveSheet.Shapes.Range(Array("Rectangle 31")).Select
        Selection.Copy


        '===================== Power Point Activation =================
        Dim PowerPointConn As Object
        Dim CurrentChart As Excel.Chart
        Set PowerPointConn = GetObject(, "PowerPoint.Application")
        PowerPointConn.Visible = msoTrue

        PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes.PasteSpecial (0)

        ' *********** Location ****************
        PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(cx).Top = Int(hp_top)
        PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(cx).Left = Int(hp_left)

End Sub

Sub format_title(Title$, SlideNumber)

        Dim PowerPointConn As Object
        Dim CurrentChart As Excel.Chart
        Set PowerPointConn = GetObject(, "PowerPoint.Application")
        PowerPointConn.Visible = msoTrue

ORIG$ = "PARAMETERS"
If Title$ = " " Or Title$ = "" Then
   Exit Sub
End If
    
' =========== Getting new line ================
subtitle_flag = False
nl = 0
B$ = ""
For i = 1 To Len(Title$)
        A$ = Right$(Left$(Title$, i), 1)
        If A$ = Chr$(10) Or A$ = Chr$(13) Then
            subtitle_flag = True
            nl = i
            Exit For
        End If
Next i

nl = i

'msg$ = "row 1 : Start= 1 End=" & nl & Chr$(10) & "Row 2: Start=" & nl & " End=" & Len(Title$)
'MsgBox (msg$ & Chr$(10) & Title$ & Chr$(10) & B$)

' ============ Getting formats ==================
Sheets(ORIG$).Cells(4, 15).Select
    font_name = Selection.Font.Name
    font_color = Selection.Font.Color
    font_bold = Selection.Font.Bold
    'font_underline = Selection.Font.Underline
    font_italic = Selection.Font.Italic
    
    font_size = Sheets(ORIG$).Cells(4, 14).Value
'================================================



PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(1).TextFrame.TextRange.Text = Title$
    
ci = 1
cf = nl
cl = cf - ci
    With PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(1).TextFrame.TextRange.Characters(Start:=1, Length:=cl).Font
        .Name = font_name
        .Color = font_color
        .Size = font_size
        .Bold = font_bold
        .Italic = font_italic
        '.Underline = font_underline
        
    End With
    
    
If subtitle_flag = True Then
 ci = nl + 1
 cf = Len(Title$)
 cl = cf - ci + 1
 
' ============ Getting formats ==================
Sheets(ORIG$).Cells(5, 15).Select
    font_name = Selection.Font.Name
    font_color = Selection.Font.Color
    font_bold = Selection.Font.Bold
    'font_underline = Selection.Font.Underline
    font_italic = Selection.Font.Italic
    font_size = Sheets(ORIG$).Cells(5, 14).Value
    
'================================================
 
    With PowerPointConn.ActivePresentation.Slides(SlideNumber).Shapes(1).TextFrame.TextRange.Characters(Start:=ci, Length:=cl).Font
        .Name = font_name
        .Color = font_color
        .Size = font_size
        .Bold = font_bold
        .Italic = font_italic
        '.Underline = font_underline
    End With
End If
End Sub


