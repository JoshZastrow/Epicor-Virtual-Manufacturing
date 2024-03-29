Attribute VB_Name = "ShopVision"
Sub main()

'Main program of Shop Vision. This program intends to run through query of labor activity and update the shapes in the user interface
    
    Dim wb As Workbook
    Dim Floor As Worksheet
    Dim Data As Worksheet
    Dim Resource As String
    Dim Delim As String
    Dim ShapeExist As Shape
    Dim Image As String
    Dim Status As String
    Dim OpNum As String
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
    Dim LaborData As ListObject
    
    Set wb = Workbooks("Shop Vision.xlsm")                                                          'Define Excel Objects
    Set Data = wb.Sheets("LaborData")
    Set Floor = wb.Sheets("ShopFloor")
    Floor.Unprotect
    
    Call ResetData(Floor)
    
    Set LaborData = Data.ListObjects("ActiveLabor_Table")                                           'Convert LaborData to 2D Array
    rw = 1                                                                                          'First row of Labor Data
     
    For rw = 1 To LaborData.ListRows.Count                                                          'Loop through Labor Data
        Resource = LaborData.DataBodyRange(rw, 1)                                                   'Resource Item

        'Check to make sure shape exists and Activity is present
        On Error Resume Next
        Set ShapeExist = Nothing
        Set ShapeExist = Floor.Shapes(Resource)
        On Error GoTo 0

        JobNum = LaborData.ListColumns("JobNum").DataBodyRange(rw, 1).Value
        
        If Not ShapeExist Is Nothing And JobNum <> "" Then                                          'Check for activity and shape
            
            OpNum = "(" & LaborData.ListColumns("OpNum").DataBodyRange(rw, 1).Value & ")"
            PartNum = LaborData.ListColumns("PartNum").DataBodyRange(rw, 1).Value                   'Dictionary values....
            LaborType = CStr(LaborData.ListColumns("LaborType").DataBodyRange(rw, 1).Value)
            Employee = LaborData.ListColumns("Employee").DataBodyRange(rw, 1).Value
            ProdQty = LaborData.ListColumns("ProdQty").DataBodyRange(rw, 1).Value
            LaborRate = LaborData.ListColumns("ProdStandard").DataBodyRange(rw, 1).Value
            PercentComplete = LaborData.ListColumns("PercentComplete").DataBodyRange(rw, 1).Value
            TodaysEstimate = LaborData.ListColumns("PercentToday").DataBodyRange(rw, 1).Value
            ProdQty = LaborData.ListColumns("ProdQty").DataBodyRange(rw, 1).Value
            ProducedAlready = LaborData.ListColumns("TotalLabor").DataBodyRange(rw, 1).Value
            ProducedToday = LaborData.ListColumns("ProducedToday").DataBodyRange(rw, 1).Value
            QtyLeft = Round(ProdQty - ProducedAlready - ProducedToday, 0)
            LaborRate = LaborData.ListColumns("ProdStandard").DataBodyRange(rw, 1).Value
            SetupTime = LaborData.ListColumns("SetupTime").DataBodyRange(rw, 1).Value
            TimeElapsed = LaborData.ListColumns("TimeElapsed").DataBodyRange(rw, 1).Value
            TimeLeft = EstimateProd(LaborType, TimeElapsed, QtyLeft, LaborRate)

            Call PartImage(Floor, Resource, PartNum, LaborType)
            Call PartStatus(Floor, Resource, LaborType)
            Call ProductionInfo(Floor, Resource, PartNum, Employee, OpNum)
            Call ProductionQty(Floor, Resource, ProdQty)
            Call JobInfo(Floor, Resource, JobNum, OpNum)
            Call StatusBar(Floor, Resource, PercentComplete, TodaysEstimate, TimeLeft, LaborType)
        
        End If

    Next rw
    Floor.Protect
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
            .Characters.Text = "No" & Chr(13) & "Img"
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


Sub ProductionInfo(ws As Worksheet, ShapeName As String, PartNum As String, Employee As String, Op As String)

'Modifies the Production Info Box

    Dim Image As Variant

    Set Image = ws.Shapes.Range(Array("Info_" & UCase(ShapeName)))                              'Create Shape object
    
    With Image.TextFrame2.TextRange                                        'Change text to "No Image"
        .Characters.Text = PartNum & Op & Chr(13) & Employee
    End With
        
End Sub

Sub ProductionQty(ws As Worksheet, ShapeName As String, ProdQty As String)

'Modifies production quantity

    Set Image = ws.Shapes.Range(Array("ReqQty_" & UCase(ShapeName)))                              'Create Shape object
    Image.TextFrame2.TextRange.Characters.Text = ProdQty

End Sub
Sub JobInfo(ws As Worksheet, ShapeName As String, JobNum As String, Op As String)

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
        
        Image.TextFrame2.TextRange.Characters.Text = TimeLeft
    
    Case "S"
    
        Image.Fill.ForeColor.RGB = RGB(255, 255, 0)
        Image.Fill.Solid
        Image.TextFrame2.TextRange.Characters.Text = TimeLeft
    Case ""
        
        Image.Fill.ForeColor.RGB = RGB(255, 0, 0)
        Image.Fill.Solid
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
    
    Dim ShapeExist As Variant
    Dim wb As Workbook
    
    Set wb = Workbooks("Shop Vision.xlsm")
    
    Set Resources = wb.Sheets("LaborData").ListObjects("ActiveLabor_Table").ListColumns("Resource ID").DataBodyRange
    For k = 1 To Resources.Rows.Count
        
        On Error Resume Next
        Set ShapeExist = Nothing
        Set ShapeExist = ws.Shapes(Resources(k, 1).Value)
        On Error GoTo 0
        If Not ShapeExist Is Nothing Then
                Call PartImage(ws, CStr(Resources(k, 1).Value), " ", "")
                Call PartStatus(ws, CStr(Resources(k, 1).Value), "")
                Call ProductionInfo(ws, CStr(Resources(k, 1).Value), " ", " ", "")
                Call ProductionQty(ws, CStr(Resources(k, 1).Value), " ")
                Call JobInfo(ws, CStr(Resources(k, 1).Value), " ", "")
                Call StatusBar(ws, CStr(Resources(k, 1).Value), 0.01, 0.01, "", "")
        End If
    Next k
End Sub
Function EstimateProd(ByVal LaborType As String, ByVal Time As Variant, ByVal QtyLeft As Integer, LaborRate As Variant) As String

Select Case LaborType

Case "P"
            If QtyLeft <= 0 Then                                                                    'Adjust Time for inaccurate estimates
                EstimateProd = "Past Complete"
            Else
                If LaborRate <= 0 Then                                                              'Message for missing labor rates
                    EstimateProd = "No Prod. Rate"
                Else
                    If (QtyLeft / LaborRate) >= 1.5 Then                                            'Message for more than 1 hour left
                        EstimateProd = CStr(Round(QtyLeft / LaborRate, 0)) & " Hours Left"
                    Else
                        EstimateProd = CStr(Round(QtyLeft / LaborRate, 0)) & " Hour Left"           'Message for 1 hour left (singular)
                    End If
                End If
            End If
            
Case "S"
            If Hour(Time) >= 0.5 And Hour(Time) < 1.5 Then EstimateProd = Hour(Time) & " hr " & Minute(Time) & " min"
            If Hour(Time) < 0.5 Or Hour(Time) >= 1.5 Then EstimateProd = Hour(Time) & " hrs " & Minute(Time) & " min"
End Select

End Function
Sub copyshape()
'
' Copies GUI machine info shape group, creates a new version for each resource ID in the labor data query
'
'Preconditions: Must have a shape group labeled as one of your resource IDs. Rename I in code below to be that ID


I = InputBox("Enter Resource ID(shape name) of first shape: ")                                      'Initial Shape Name (The first Resource ID shape)

R_Col = Workbooks("Shop Vision.xlsm").Sheets("LaborData").ListObjects("ActiveLabor_Table").ListColumns("Resource ID").DataBodyRange

For Each R In R_Col
    Debug.Print R
    ActiveSheet.Shapes.Range(Array(I)).Select
    Selection.Copy
    ActiveSheet.Paste
    ActiveSheet.Shapes.Range(Array(I)).Name = R
    ActiveSheet.Shapes.Range(Array("ReqQty_" & I)).Name = "ReqQty_" & R
    ActiveSheet.Shapes.Range(Array("Status_" & I)).Name = "Status_" & R
    ActiveSheet.Shapes.Range(Array("JobNum_" & I)).Name = "JobNum_" & R
    ActiveSheet.Shapes.Range(Array("Image_" & I)).Name = "Image_" & R
    ActiveSheet.Shapes.Range(Array("Info_" & I)).Name = "Info_" & R
    ActiveSheet.Shapes.Range(Array("Progress_" & I)).Name = "Progress_" & R
Next R

End Sub

