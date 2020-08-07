VERSION 5.00
Begin VB.Form frmdep 
   BorderStyle     =   0  'None
   ClientHeight    =   570
   ClientLeft      =   13485
   ClientTop       =   585
   ClientWidth     =   1395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdgo 
      Caption         =   "Go"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   120
   End
   Begin VB.Label lblheure 
      Caption         =   "0:0:0"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnupopu 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnudirpath 
         Caption         =   "dirpath"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmdep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vheures As Integer
Dim vminutes As Integer
Dim vsecondes As Integer
Private Sub cmdgo_Click()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
    cmdgo.Caption = "Go"
Else
    Timer1.Enabled = True
    cmdgo.Caption = "Stop"
End If
End Sub


Private Sub Form_Load()

WindowOnTop Me


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
frmdep.PopupMenu mnupopu, , x, y
End Sub

Private Sub Form_Unload(Cancel As Integer)
If exist <> "Faux" Then
temp = App.Path & "\" & "temp.tmp"
Open temp For Input As #1
    fichier = Input(LOF(1) - 2, 1)
Close #1
'MsgBox fichier
    
    Open fichier For Input As #1
        a = Input(LOF(1), 1)
    Close #1

    Open fichier For Output As #1
    'msgBox Date
    'MsgBox Time
    'MsgBox lblheure.Caption
        'Print #1,
        
        Print #1, Date; Tab; Time; Tab; lblheure.Caption
        Print #1, a
    Close #1
    Else: MsgBox fichier & exist
End If

End Sub

Private Sub mnuabout_Click()
MsgBox "Bonjour, ce petit programme a été conçu par Patrick TELLIER (uin 12481992) merci", vbOKOnly, "A propos"
End Sub

Private Sub mnudirpath_Click()
MsgBox App.Path
End Sub

Private Sub mnuquit_Click()
Unload Me
End
End Sub

Private Sub Timer1_Timer()
lblheure.Caption = vheures & ":" & vminutes & ":" & vsecondes
vsecondes = Val(vsecondes) + 1
If vsecondes > 59 Then
    vsecondes = 0
    vminutes = vminutes + 1
    If App.Path & "\" & "fichier.pat" <> "" Then
            Open App.Path & "\" & "fichier.tmp" For Output As #1
                Print #1, Date; Tab(1); Time; Tab(1); lblheure
            Close #1
    End If
    
    If vminutes > 59 Then
        vminutes = 0
        vheures = Val(vheures) + 1
    End If
End If

End Sub
