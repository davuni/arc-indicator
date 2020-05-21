Sub Indicator()
' bydukenuke@newsmth.net
' 2010/7/11 06:13
'
' Update by oicu#lsxk.org
' 2010/9/12 20:44
'
' Arc shape ported by D. Squirryl
' 2020/5/14 11:42

    Dim mySlides As Slides
    Dim pageBar As ShapeRange
    Dim pageIndicator As Shape
    Dim pageWidth, pageHeight, pageStep
    Dim MyArray() As Variant
    Dim i, j, k
    j = 0
    k = 0

    Set mySlides = Application.ActivePresentation.Slides
    pageWidth = Application.ActivePresentation.SlideMaster.Width
    pageHeight = Application.ActivePresentation.SlideMaster.Height
    
' Parameters
    Dim x, y, w, h, r, color, colorbg, colorfocus, trans, transbg, transfocus
    x = pageWidth - 90
    y = 40
    w = 50
    h = 50
    r = 0.1
    color = RGB(255, 255, 255)
    colorbg = RGB(255, 255, 255)
    colorfocus = RGB(255, 255, 255)
    trans = 0.8
    transbg = 0.95
    transfocus = 0.75
' Parameters

    ReDim MyArray(mySlides.Count, 0)
    For i = 1 To mySlides.Count
        If mySlides.Item(i).SlideShowTransition.Hidden = True Then
            j = j + 1
            MyArray(i, 0) = 1
        Else
            MyArray(i, 0) = 0
        End If
    Next
    If mySlides.Count - j > 0 Then
        pageStep = 360 / (mySlides.Count - j)
    Else
        pageStep = 0
    End If
    On Error Resume Next
    For i = 1 To mySlides.Count
        k = k + MyArray(i, 0)
        If IsNull(pageBar) Or pageBar.Count = 0 Then GoTo newIndicator
        Set pageIndicator = pageBar.Item(1)
        GoTo nextPage
        
newIndicator:
        Set pageIndicatorBG = mySlides.Item(i).Shapes.AddShape(msoShapeBlockArc, x, y, w, h)
        Set pageIndicator = mySlides.Item(i).Shapes.AddShape(msoShapeBlockArc, x, y, w, h)
        Set pageIndicatorFocus = mySlides.Item(i).Shapes.AddShape(msoShapeBlockArc, x, y, w, h)
        
        pageIndicator.Fill.Transparency = trans
        pageIndicatorBG.Fill.Transparency = transbg
        pageIndicatorFocus.Fill.Transparency = transfocus
        pageIndicator.Fill.ForeColor.RGB = color
        pageIndicatorBG.Fill.ForeColor.RGB = colorbg
        pageIndicatorFocus.Fill.ForeColor.RGB = colorfocus
        pageIndicator.Line.Visible = msoFalse
        pageIndicatorBG.Line.Visible = msoFalse
        pageIndicatorFocus.Line.Visible = msoFalse
        pageIndicator.Adjustments.Item(1) = -90
        pageIndicator.Adjustments.Item(2) = pageStep - 90
        pageIndicator.Adjustments.Item(3) = r
        pageIndicatorBG.Adjustments.Item(1) = 0
        pageIndicatorBG.Adjustments.Item(2) = 360
        pageIndicatorBG.Adjustments.Item(3) = r
        pageIndicatorFocus.Adjustments.Item(1) = -90
        pageIndicatorFocus.Adjustments.Item(2) = pageStep - 90
        pageIndicatorFocus.Adjustments.Item(3) = r

nextPage:
        pageIndicator.Adjustments.Item(2) = i * pageStep - 90
        pageIndicatorFocus.Adjustments.Item(1) = (i - 1) * pageStep - 90
        pageIndicatorFocus.Adjustments.Item(2) = i * pageStep - 90
        
    Next
End Sub
