Sub ImportPictures()
   ' How to use this function:
   '
   ' Preparation:
   ' Set browser size to 1260x1024 using the chrome extension "windows resizer"
   ' Look at the website in question and make a screenshot using for example extension "Screen Capture"
   '
   ' Animation:
   ' Most front pages these days have some kind of rotation going on. If you want to, you can capture a number of screenshots and
   ' name additional screenshots with the same name as prefix followed by a sequence number.
   ' For example "mainsite.png", "mainsite2.png", "mainsite3.png" etc.
   '
   ' Create a blank slide in powerpoint in the appropriate industry section and run this macro
   ' You will be asked for the name of the file(s) and how many there are.
   '
   ' When that is done the macro will insert these screenshots, size them appropriately and add 2 second automatic fade animation
   ' between each of them.
   '
   ' Test it by pressing "Shift-F5" in powerpoint.
   
   ' -----
   
   
   
   Dim oSlide As Slide
   Dim oPicture As Shape
   Dim desktop As String
   Dim myTrigger As String
   Dim myAnimation As String
   Dim myDuraction, myFirstDuration As String
   
   
   myDesktop = "C:\Users\tb\Desktop\"
   
   ' Trigger Options:
   '   msoAnimTriggerAfterPrevious
   '   msoAnimTriggerMixed
   '   msoAnimTriggerNone
   '   msoAnimTriggerOnPageClick
   '   msoAnimTriggerOnShapeClick
   '   msoAnimTriggerWithPrevious

   myTrigger = msoAnimTriggerAfterPrevious
   myAlternativeTrigger = msoAnimTriggerOnPageClick
   
   ' Animation options:
   '   Source: http://msdn.microsoft.com/en-us/library/office/aa157283(v=office.10).aspx
   
   myAnimation = msoAnimEffectFade
   myAlternativeAnimation = msoAnimEffectAppear
   
   ' Slide animation timing:
   myFirstDuration = 0.5
   myDuration = 1.75
   MyDelay = 2
   MyFirstDelay = 0
   
   ' Change slide index position to the first slide
   '  ActiveWindow.View.GotoSlide 1

   ' Set oSlide to the active slide in the normal view.
   ' Assuming, of course, the user has created a new blank slide with notes from ther reference update and this is now
   ' the currently active slide
   
   Set oSlide = ActiveWindow.Presentation.Slides(ActiveWindow.Selection.SlideRange(1).SlideIndex)
   
   fname = InputBox("Enter the generic sreen shot name:", "Filename")
   
   'Handle multiple screenshots
   Answer = MsgBox("Is this the only screenshot?", vbQuestion + vbYesNo, "Just one?")
   If Answer = vbNo Then
      fnumber = InputBox("How many screenshots are there:", "Number?")
      If Not IsNumeric(fnumber) Then
         Debug.Print "Assuming just one screenshot"
         fnumber = 0
      End If
      Answer = MsgBox("Do you want automatic transitions?", vbQuestion + vbYesNo, "Auto build?")
      If Answer = vbNo Then
         myTrigger = myAlternativeTrigger
      End If
      Answer = MsgBox("Do you want fading transitions?", vbQuestion + vbYesNo, "Fade?")
      If Answer = vbNo Then
         myAnimation = myAlternativeAnimation
      End If
   Else
      fnumber = 1
      myAnimation = msoAnimEffectAppear
   End If
    
   ' Maybe check Existance of file later?
   ' For now assume there is a single screenshot with no number
   ' optionally followed by same name with a suffix from 2 and up
       
   For i = 1 To fnumber
       If i = 1 Then
          FileName = fname & ".png"
       Else
          FileName = fname & Str(i) & ".png"
       End If
       
       Debug.Print "Inserting: " & FileName
       Set oPicture = oSlide.Shapes.AddPicture(myDesktop & FileName, _
           msoFalse, msoTrue, 1, 1, -1, -1)
          
       ' MsgBox (ActivePresentation.PageSetup.SlideWidth) tells me it is 720 pixels
     
       ' Now scale the picture to full size, with "Relative to original
       ' picture size" set to true for both height and width.
       oPicture.ScaleHeight 1, msoFalse
       oPicture.ScaleWidth 1, msoFalse
   

       ' Move the picture to the center of the slide. Select it.
       With ActivePresentation.PageSetup
          oPicture.Left = (.SlideWidth \ 2) - (oPicture.Width \ 2)
          oPicture.Top = (.SlideHeight \ 2) - (oPicture.Height \ 2)
          oPicture.Select
       End With
       
       ' Add animation to the screenshot
       ' Make sure the first screenshot is trigged automatically - otherwise the screen will be blank
       
       thisTrigger = myTrigger
       If i = 1 Then
          thisTrigger = msoAnimTriggerWithPrevious
       End If
       
       Set Effect = oSlide.TimeLine.MainSequence _
                      .AddEffect(Shape:=oPicture, _
                      EffectID:=myAnimation, _
                      trigger:=thisTrigger)

       'Sets the duration of the animation if automatic transition
       If myTrigger = msoAnimTriggerAfterPrevious Then
          If i = 1 Then
             If myAnimation <> msoAnimEffectAppear Then
                Effect.Timing.Duration = myFirstDuration
                Effect.Timing.TriggerDelayTime = MyFirstDelay
             End If
          Else
             Effect.Timing.Duration = myDuration
             Effect.Timing.TriggerDelayTime = MyDelay
          End If
       End If
       
   Next i

End Sub