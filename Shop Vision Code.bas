Attribute VB_Name = "ShopVision"
Sub main()

'Main program of Shop Vision. This program intends to run through query of labor activity and update the shapes in the user interface
    
    Dim wb As Workbook
    Dim Floor As Worksheet
    Dim ws As Worksheet
    Dim LaborData As Variant
    Dim Resource As String
    Dim Delim As String
    Dim ShapeExist As Shape
    Dim Image As String
    Dim Status As String
    Dim Info As String
    Dim ReqQty As String
    Dim Job As String
    Dim Progress As String
    Dim PartNum As String
    Dim JobNum As String
    Dim Employee As String
    Dim ProdQty As String
    Dim PercentComplete As Variant
    Dim TodaysEstimate As Variant
    Dim TimeLeft As Variant

    
    Set wb = Workbooks("Shop Vision.xlsm")
    Set ws = wb.Sheets("LaborData")
    Set Floor = wb.Sheets("First Floor")

    
    Call ResetData(Floor)
    
    LaborData = ws.Range("A1").CurrentRegion.Offset(1, 0)                                        'Convert LaborData to 2D Array
    rw = 1                                                                                       'First row of Labor Data
    
    Do Until LaborData(rw, 7) = ""                                                               'Loop through Labor Data
        Resource = LaborData(rw, 7)                                                              'Resource Item

            'Check to make sure shape exists
            On Error Resume Next
            Set ShapeExist = Nothing
            Set ShapeExist = wb.Sheets("First Floor").Shapes(Resource)
            On Error GoTo 0
            If Not ShapeExist Is Nothing Then
            
                ''''''''''''Update User Interface''''''''''''''''''''''''''
                PartNum = LaborData(rw, 4)
                JobNum = LaborData(rw, 1)
                LaborType = CStr(LaborData(rw, 10))
                Employee = LaborData(rw, 5)
                ProdQty = LaborData(rw, 12)
                PercentComplete = Round(LaborData(rw, 15), 2)
                TodaysEstimate = Round(LaborData(rw, 16), 2)
                TimeLeft = Round(ProdQty - LaborData(rw, 3) - LaborData(rw, 14), 0)
                If TimeLeft < 0 Then
                    TimeLeft = 0
                Else
                    If LaborData(rw, 11) <= 0 Then
                        TimeLeft = "??"
                    Else
                        TimeLeft = Round(TimeLeft / LaborData(rw, 11), 0)
                    End If
                End If
                
                Call PartImage(Floor, Resource, PartNum, LaborType)
                Call PartStatus(Floor, Resource, LaborType)
                Call ProductionInfo(Floor, Resource, PartNum, Employee)
                Call ProductionQty(Floor, Resource, ProdQty)
                Call JobInfo(Floor, Resource, JobNum)
                Call StatusBar(Floor, Resource, PercentComplete, TodaysEstimate, TimeLeft, LaborType)
            End If

            rw = rw + 1
        
    Loop
    ws.Protect (UserInterfaceOnly = True)
End Sub

Sub PartImage(ws, ShapeName As String, PartNum As String, LaborType As Variant)

    'Modify PartImage Shape to have the appropriate fill. If a image file exists, replace the fill with an image, else have a standard blank image

    Dim Pic As String
    Dim Image As Variant

    Set Image = ws.Shapes.Range(Array("Image_" & UCase(ShapeName)))                              'Create Shape object
    
    Select Case LaborType
    
    Case "P", "S"
    
    If Dir("S:\Engineering\Josh Zastrow\SMC images\" & PartNum & ".png") <> "" Then
        Pic = "S:\Engineering\Josh Zastrow\SMC images\" & PartNum & ".png"
        With Image.Fill                                                         'Insert Pic
            .Visible = msoTrue
            .UserPicture Pic
            .TextureTile = msoFalse
            .RotateWithObject = msoTrue
        End With
        
        With Image.TextFrame2.TextRange                                         'Reset Text to nothing
            .Characters.Text = ""
        End With
        
    Else
    
        With Image.Fill                                                         'Change Background Image to white
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .Solid
        End With
        
        With Image.TextFrame2.TextRange                                         'Change text to "No Image"
            .Characters.Text = "No" & Chr(13) & "Image"
            .Font.Bold = msoTrue
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        
    End If
    
    Case ""
    
            With Image.Fill                                                         'Change Background Image to white
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .Solid
        End With
        
        With Image.TextFrame2.TextRange                                         'Change text to "No Image"
            .Characters.Text = "IDLE"
            .Font.Bold = msoTrue
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End Select
End Sub


Sub PartStatus(ws As Worksheet, ShapeName As String, LaborType As Variant)

'Modifies Status button based on labor type. Red means no work, yellow is setup, green is production

    Dim Image As Variant
    Dim Pic As String
    
    Set Image = ws.Shapes.Range("Status_" & UCase(ShapeName))                                      'Create Shape object

    Select Case LaborType                                                      'Modify Status Color
    
        Case ""
        
            With Image.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 0, 0)
                .Transparency = 0
                .Solid
            End With
            
        Case "S"
        
            With Image.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 192, 0)
                .Transparency = 0
                .Solid
            End With
                        
        Case "P"
        
            With Image.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 176, 80)
                .Transparency = 0
                .Solid
            End With
        
        End Select
End Sub


Sub ProductionInfo(ws As Worksheet, ShapeName As String, PartNum As String, Employee As String)

'Modifies the Production Info Box

    Dim Image As Variant

    Set Image = ws.Shapes.Range(Array("Info_" & UCase(ShapeName)))                              'Create Shape object
    
    With Image.TextFrame2.TextRange                                        'Change text to "No Image"
        .Characters.Text = PartNum & Chr(13) & Employee
    End With
        
End Sub

Sub ProductionQty(ws As Worksheet, ShapeName As String, ProdQty As String)

'Modifies production quantity

    Set Image = ws.Shapes.Range(Array("ReqQty_" & UCase(ShapeName)))                              'Create Shape object
    Image.TextFrame2.TextRange.Characters.Text = ProdQty

End Sub
Sub JobInfo(ws As Worksheet, ShapeName As String, JobNum As String)

'Modifies production job number
 
    Set Image = ws.Shapes.Range(Array("JobNum_" & UCase(ShapeName)))                              'Create Shape object
    Image.TextFrame2.TextRange.Characters.Text = JobNum

End Sub

Sub StatusBar(ws As Worksheet, ShapeName As String, GradStop1 As Variant, GradStop2 As Variant, TimeLeft As Variant, LaborType As Variant)

'Modifies Status Bar

    Set Image = ws.Shapes.Range(Array("Progress_" & UCase(ShapeName)))                'Create Shape object
    
    Select Case LaborType
    
    Case "P"
        
        'Make sure GradStops don't exceed 100%
        If (GradStop1 + GradStop2 + 0.02) >= 1 Then
            If GradStop1 >= 1 Then
                GradStop1 = 0.97
                GradStop2 = 0.01
            Else
                GradStop2 = 1 - 0.03 - GradStop1
            End If
        End If
        
        With Image.Fill
            .ForeColor.RGB = RGB(255, 255, 255)
            .OneColorGradient msoGradientVertical, 1, 1
            .GradientStops.Insert vbGreen, 0
            .GradientStops.Insert vbGreen, GradStop1
            .GradientStops.Insert RGB(155, 155, 155), GradStop1 + 0.01
            .GradientStops.Insert RGB(155, 155, 155), GradStop1 + GradStop2
            .GradientStops.Insert vbWhite, GradStop1 + GradStop2 + 0.01
        End With
        
        Image.TextFrame2.TextRange.Characters.Text = TimeLeft & " Hours Left"
    
    Case "S"
    
        Image.Fill.ForeColor = vbYellow
        Image.TextFrame2.TextRange.Characters.Text = "Setup Time unknown"
    Case ""
        
        Image.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Image.Fill.OneColorGradient msoGradientVertical, 1, 1
        Image.TextFrame2.TextRange.Characters.Text = ""
    End Select
End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  
  Bool = False
  For j = 1 To UBound(arr)
  If arr(j, 1) = stringToBeFound Then
    Bool = True
    Exit For
  End If
  Next
  IsInArray = Bool
End Function

Sub ResetData(ws As Worksheet)

    'Check to make sure shape exists
    
    Dim Resources() As Variant
    Dim ShapeExist As Variant
    Dim wb As Workbook
    
    Set wb = Workbooks("Shop Vision.xlsm")
    Resources = wb.Sheets("Resources").Range("A1:A65").Value
    

    For k = 1 To UBound(Resources)
    
        On Error Resume Next
        Set ShapeExist = Nothing
        Set ShapeExist = wb.Sheets("First Floor").Shapes(CStr(Resources(k, 1)))
        On Error GoTo 0
        If Not ShapeExist Is Nothing Then
                Call PartImage(ws, CStr(Resources(k, 1)), " ", "")
                Call PartStatus(ws, CStr(Resources(k, 1)), "")
                Call ProductionInfo(ws, CStr(Resources(k, 1)), " ", " ")
                Call ProductionQty(ws, CStr(Resources(k, 1)), " ")
                Call JobInfo(ws, CStr(Resources(k, 1)), " ")
                Call StatusBar(ws, CStr(Resources(k, 1)), 0.01, 0.01, "", "")
        End If
    Next k
End Sub
