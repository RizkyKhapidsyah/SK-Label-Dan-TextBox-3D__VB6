Attribute VB_Name = "Mesin"
Option Explicit

Public Sub PaintControl3D(frm As Form, Ctl As Control)
    ' This Sub draws lines around controls to make them 3d
    
    ' darkgrey, upper - horizontal
    frm.Line (Ctl.Left, Ctl.Top - 15)-(Ctl.Left + _
          Ctl.Width, Ctl.Top - 15), &H808080, BF
    ' darkgrey, left - vertical
    frm.Line (Ctl.Left - 15, Ctl.Top)-(Ctl.Left - 15, _
         Ctl.Top + Ctl.Height), &H808080, BF
    ' white, right - vertical
    frm.Line (Ctl.Left + Ctl.Width, Ctl.Top)- _
      (Ctl.Left + Ctl.Width, Ctl.Top + Ctl.Height), &HFFFFFF, BF
    ' white, lower - horizontal
    frm.Line (Ctl.Left, Ctl.Top + Ctl.Height)- _
    (Ctl.Left + Ctl.Width, Ctl.Top + Ctl.Height), &HFFFFFF, BF

End Sub

Public Sub PaintForm3D(frm As Form)
    ' This Sub draws lines around the Form to make it 3d
    
    ' white, upper - horizontal
    frm.Line (0, 0)-(frm.ScaleWidth, 0), &HFFFFFF, BF
    ' white, left - vertical
    frm.Line (0, 0)-(0, frm.ScaleHeight), &HFFFFFF, BF
    ' darkgrey, right - vertical
    frm.Line (frm.ScaleWidth - 15, 0)-(frm.ScaleWidth - 15, _
       frm.Height), &H808080, BF
    ' darkgrey, lower - horizontal
    frm.Line (0, frm.ScaleHeight - 15)-(frm.ScaleWidth, _
      frm.ScaleHeight - 15), &H808080, BF

End Sub

