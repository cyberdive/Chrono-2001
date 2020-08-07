Attribute VB_Name = "Module1"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub WindowOnTop(frm As Form)
  
  x% = frm.Left
  y% = frm.Top
  w% = frm.Width
  h% = frm.Height
  TwipsToPixel x%, y%
  TwipsToPixel w%, h%
  SetWindowPos frm.hwnd, -1, x%, y%, w%, h%, &H50

End Sub

Public Sub TwipsToPixel(xValeur%, yValeur%)
  
  ux% = Screen.TwipsPerPixelX
  xValeur% = xValeur% / ux%
  uy% = Screen.TwipsPerPixelY
  yValeur% = yValeur% / uy%

End Sub


