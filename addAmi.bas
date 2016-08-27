Sub AnimationRandomizor()

    Const NumEffects As Byte = 77

    Dim SlideObject As Slide
    Dim ShapeObject As Shape
    Dim HoldRandomValue As Byte
    Dim EffectList(1 To NumEffects) As Long
    Dim TotalChanges As Long

    ' Used for error trapping.
    On Error Resume Next
    Err.Clear

    ' Initialize the counters.
    TotalChanges = 0

    ' Set up the Effect list.
    SetUpEffectList EffectList()
    Dim x(100) As Shape

    ' Outer loop goes through every slide in the Active presentation.
    SlideCount = 0
    For Each SlideObject In Application.ActivePresentation.Slides

        ' Inner loop goes through every shape in the presentation.
        Count = 0
        For Each ShapeObject In SlideObject.Shapes
            Count = Count + 1
            Set x(Count) = SlideObject.Shapes(Count)
            'MsgBox (x(Count).Top)
        Next ShapeObject
        'MsgBox (Count)
        For i = 1 To Count - 1
            For j = i + 1 To Count
                If x(i).Top > x(j).Top Then
                    Set t = x(i)
                    Set x(i) = x(j)
                    Set x(j) = t
                End If
            Next j
        Next i
        For i = 1 To Count
            'Set ooo = x(i)
            With x(i).AnimationSettings

                ' This property must be set to True for any of the
                ' other properties of the AnimationSettings object
                ' to take effect.
                .Animate = msoTrue

                ' Assign a random animation to the object.
                'Randomize
                'HoldRandomValue = Int((NumEffects * Rnd) + 1)

                ' Assign a random animatation to the object.
                .EntryEffect = EffectList(1)
                If Err.Number <> 0 Then
                    'MsgBox "An error occured. Try runnning the macro " _
                    '& "again.", vbCritical, "Error"
                End If

                ' Increment the object count.
                TotalChanges = TotalChanges + 1
            End With
        Next i
        SlideCount = SlideCount + 1
        If (SlideCount Mod 10) = 0 Then
            'MsgBox (SlideCount)
        End If

    Next SlideObject

    ' See whether any objects were changed.
    If TotalChanges = 0 Then
        MsgBox "No objects available. No changes were made " _
            & "to the presentation.", vbInformation, "No Objects"
    Else

        ' Set up the message box.
        If TotalChanges = 1 Then
            MsgBox "One object was given a random custom animation.", _
                vbInformation, "Random Custom Animation"
        Else
            MsgBox "Suscessfully applied a custom animation to all " _
                & "objects.", vbInformation, TotalChanges _
                & " Objects Animated"
        End If

    End If

End Sub

' Assign effect constants to the List Array.

Sub SetUpEffectList(ByRef List() As Long)

    ' Appear
    List(1) = ppEffectAppear

    ' Fly Effects
    List(2) = ppEffectFlyFromBottom
    List(3) = ppEffectFlyFromBottomLeft
    List(4) = ppEffectFlyFromBottomRight
    List(5) = ppEffectFlyFromLeft
    List(6) = ppEffectFlyFromRight
    List(7) = ppEffectFlyFromTop
    List(8) = ppEffectFlyFromTopLeft
    List(9) = ppEffectFlyFromTopRight

    ' Blinds Effects
    List(10) = ppEffectBlindsHorizontal
    List(11) = ppEffectBlindsVertical

    ' Box Effects
    List(12) = ppEffectBoxIn
    List(13) = ppEffectBoxOut

    ' Checkerboard Effects
    List(14) = ppEffectCheckerboardAcross
    List(15) = ppEffectCheckerboardDown

    ' Crawl Effects
    List(16) = ppEffectCrawlFromDown
    List(17) = ppEffectCrawlFromLeft
    List(18) = ppEffectCrawlFromRight
    List(19) = ppEffectCrawlFromUp

    ' Dissolve
    List(20) = ppEffectDissolve

    ' Flash Effect
    List(21) = ppEffectFlashOnceFast
    List(22) = ppEffectFlashOnceMedium
    List(23) = ppEffectFlashOnceSlow

    ' Peek Effect
    List(24) = ppEffectPeekFromDown
    List(25) = ppEffectPeekFromLeft
    List(26) = ppEffectPeekFromRight
    List(27) = ppEffectPeekFromUp

    ' Random Effects
    List(28) = ppEffectRandomBarsHorizontal
    List(29) = ppEffectRandomBarsVertical

    ' Spiral
    List(30) = ppEffectSpiral

    ' Split Effects
    List(31) = ppEffectSplitHorizontalIn
    List(32) = ppEffectSplitHorizontalOut
    List(33) = ppEffectSplitVerticalIn
    List(34) = ppEffectSplitVerticalOut

    ' Stretch Effects
    List(35) = ppEffectStretchAcross
    List(36) = ppEffectStretchDown
    List(37) = ppEffectStretchLeft
    List(38) = ppEffectStretchRight
    List(39) = ppEffectStretchUp

    ' Strips Effects
    List(40) = ppEffectStripsDownLeft
    List(41) = ppEffectStripsDownRight
    List(42) = ppEffectStripsLeftDown
    List(43) = ppEffectStripsLeftUp
    List(44) = ppEffectStripsRightDown
    List(45) = ppEffectStripsRightUp
    List(46) = ppEffectStripsUpLeft
    List(47) = ppEffectStripsUpRight

    ' Swivel
    List(48) = ppEffectSwivel

    ' Wipe Effects
    List(49) = ppEffectWipeDown
    List(50) = ppEffectWipeLeft
    List(51) = ppEffectWipeRight
    List(52) = ppEffectWipeUp

    ' Zoom Effects
    List(53) = ppEffectZoomBottom
    List(54) = ppEffectZoomCenter
    List(55) = ppEffectZoomIn
    List(56) = ppEffectZoomInSlightly
    List(57) = ppEffectZoomOut
    List(58) = ppEffectZoomOutSlightly

    ' The following effects may not work.

    ' Uncover Effects
    List(59) = ppEffectUncoverDown
    List(60) = ppEffectUncoverLeft
    List(61) = ppEffectUncoverLeftDown
    List(62) = ppEffectUncoverLeftUp
    List(63) = ppEffectUncoverRight
    List(64) = ppEffectUncoverRightDown
    List(65) = ppEffectUncoverRightUp
    List(66) = ppEffectUncoverUp

    ' Cover Effects
    List(67) = ppEffectCoverDown
    List(68) = ppEffectCoverLeft
    List(69) = ppEffectCoverLeftDown
    List(70) = ppEffectCoverLeftUp
    List(71) = ppEffectCoverRight
    List(72) = ppEffectCoverRightDown
    List(73) = ppEffectCoverRightUp
    List(74) = ppEffectCoverUp

    ' Cut Effects
    List(75) = ppEffectCut
    List(76) = ppEffectCutThroughBlack

    ' Fade
    List(77) = ppEffectFade

End Sub