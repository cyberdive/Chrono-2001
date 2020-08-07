VERSION 5.00
Begin VB.Form frmnom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sut quel client voulez vous travailler ?"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Client :"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdmoins 
         Caption         =   "Supprimer un client"
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdplus 
         Caption         =   "Ajouter un client"
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   120
         Pattern         =   "*.clt"
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmnom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmoins_Click()
Dim nom As String
Dim x
nom = InputBox("Quel client voulez vous supprimer ?", "Supprimer un client")

message = "Etes-vous sur de vouloir supprimer le client" & nom
atuer = App.Path & "\" & nom & ".clt"
titre = "Suppression d'un client"
x = MsgBox(message, vbYesNo, titre)
If x = vbYes Then Kill (atuer)
File1.Refresh


End Sub

Private Sub cmdplus_Click()
Dim nom As String
nom = InputBox("Entrez le nom que vous voulez donner à votre client : ", "Nom de client")
nom = App.Path & "\" & nom & ".clt"
Open nom For Output As #1
Close #1
File1.Refresh

End Sub

Private Sub File1_DblClick()
fichier = App.Path & "\" & File1.filename
'MsgBox fichier
'client = InputBox("Veuillez entrez le nom de votre client :", "Nom du client")
'fichier = App.Path & "\" & client & ".clt"
'MsgBox fichier & exist
'On Error GoTo erreur
Open fichier For Input As #1
Close #1

temp = App.Path & "\" & "temp.tmp"
Open temp For Output As #1
    Print #1, fichier
Close #1

Unload Me
frmdep.Show

Exit Sub
erreur:

If MsgBox("Votre client n'existe pas, voulez vous le créer ?", vbYesNo, "client inconnu") = vbYes Then
        Open App.Path & "\" & client & ".clt" For Output As #1
        Close #1
Else: exist = "Faux"
End If


End Sub

Private Sub Form_Load()
File1.Path = App.Path
End Sub
