Attribute VB_Name = "modAgent"


Sub SetBalloonStyleOptions()
'------------------------------------------------
'-- This subroutine sets the check boxes for the
'-- the word balloon settings
'------------------------------------------------

'-- Check to see if the balloon is on

If Character.Balloon.Style And BalloonOn Then
    'BalloonStyleOption(0).Value = 1
Else
    'BalloonStyleOption(0).Value = 0
End If

'-- Check to see if Auto-Hide is on

If Character.Balloon.Style And AutoHide Then
    'BalloonStyleOption(1).Value = 1
Else
    'BalloonStyleOption(1).Value = 0
End If

'-- Check to see if Auto-Pace is on

If Character.Balloon.Style And AutoPace Then
    'BalloonStyleOption(2).Value = 1
Else
    'BalloonStyleOption(2).Value = 0
End If

'-- Check to see if Size-To-Text is on

If Character.Balloon.Style And SizeToText Then
    'BalloonStyleOption(3).Value = 1
Else
    'BalloonStyleOption(3).Value = 0
End If


'-- Set the controls based on Advanced Character Options

If Not Character.Balloon.Enabled Then
    'BalloonStyleOption(0).Enabled = False
    'BalloonStyleOption(1).Enabled = False
    'BalloonStyleOption(2).Enabled = False
    'BalloonStyleOption(2).Enabled = False
Else
    'BalloonStyleOption(0).Enabled = True
    'BalloonStyleOption(1).Enabled = True
    'BalloonStyleOption(2).Enabled = True
    'BalloonStyleOption(2).Enabled = True
End If

End Sub
