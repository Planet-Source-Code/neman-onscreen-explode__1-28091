VERSION 5.00
Object = "{34F681D0-3640-11CF-9294-00AA00B8A733}#1.0#0"; "DANIM.DLL"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "OnScreenX"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   ControlBox      =   0   'False
   Icon            =   "OnScreenX2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7125
   Begin DirectAnimationCtl.DAViewerControl DAControl 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DAControl_Click()
    End
End Sub

Private Sub Form_Load()


Form1.Left = 200 * Screen.TwipsPerPixelX
Form1.Top = -1 * Screen.TwipsPerPixelY

MsgBox "Just click on animation to stop"

    lHandle = SetForegroundWindow(Form1.hwnd)


DoEvents

Form1.Visible = False

BitBlt Form1.hDC, 0, -1, Form1.Width, Form1.Height, GetDC(0), 200, -1, vbSrcCopy

Form1.Visible = True


Set m = DAControl.MeterLibrary


   
   geoBase = CurDir$ & "\"
   
  
  
 Set swGeo = m.ModifiableBehavior(m.EmptyGeometry)
 Set swImg = m.ModifiableBehavior(m.EmptyImage)
 
 Set rawImg1 = m.ImportImage(geoBase + "earth.jpg")
 
 Set xf = CreateObject("DX3DTransform.Microsoft.Shapes")
 xf.Shape = "SPHERE"
 'xf.DoubleSided = False
   
 Set resulta = m.ApplyDXTransform(xf, Array(rawImg1), Null)
 Set rawGeo = resulta.OutputBvr

 
 Set xf1 = CreateObject("DX3DTransform.Microsoft.Explode")
  
  xf1.PositionJump = 1
  xf1.MaxRotations = 20
  xf1.FinalVelocity = 50
  'xf1.DecayTime = 0.31
  'xf1.Tumble = False
  
 Set holdTime = m.DANumber(0).Duration(1)
 Set forward = m.Interpolate(0, 0.1, 12)
 Set back = m.Interpolate(0.1, 0, 4)
 Set evaluator = m.Sequence(holdTime, m.Sequence(forward, back)).RepeatForever()
 
      
 Set result = m.ApplyDXTransform(xf1, Array(rawGeo), evaluator)
 Set realGeo = result.OutputBvr
 Set realTransScale = m.Scale3Uniform(0.01)
 Set realTransRotY = m.Rotate3RateDegrees(m.YVector3, 25)
 Set realTransF = m.Compose3Array(Array(realTransScale, realTransRotY))
 Set realGeo = realGeo.Transform(realTransF)
  swGeo.SwitchTo (realGeo)
 Set camera = m.PerspectiveCamera(0.6, 0.59)
 Set light = m.UnionGeometry(m.AmbientLight, m.DirectionalLight.Transform(m.Rotate3Degrees(m.XVector3, -90)))
 Set lightAndGeo = m.UnionGeometry(swGeo, light)
  swImg.SwitchTo (lightAndGeo.Render(camera))
  

  
 DAControl.Image = swImg
 DAControl.Start

End Sub


Private Sub Form_LostFocus()

    BlitAgain

End Sub

Private Sub Form_Resize()

Form1.DAControl.Height = Form1.Height
Form1.DAControl.Width = Form1.Width

End Sub

Private Function BlitAgain()

Form1.Visible = False

DoEvents

BitBlt Form1.hDC, 0, -1, Form1.Width, Form1.Height, GetDC(0), 200, -1, vbSrcCopy

    lHandle = SetForegroundWindow(Form1.hwnd)
    SetActiveWindow Form1.hwnd
    

Form1.Visible = True


End Function
