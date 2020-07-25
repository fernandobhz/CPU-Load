VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPU Load"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CPU Load"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Este programa irá fazer calculos em loop infinito para gerar carga para a CPU."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Command1.Caption <> "Cancelar" Then
    Command1.Caption = "Cancelar"
Else
    Command1.Caption = "CPU Load"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Command1.Caption = "Cancelar" Then
    MsgBox "Primeiro cancele a operação", vbCritical
    Cancel = 1
End If
End Sub

Private Sub Timer1_Timer()
Dim i As Long
i = 0

Dim a As Double
Do While Command1.Caption = "Cancelar"
    
    a = Sqr(10 * 10 * 10) * Sqr(10 * 10 * 10)
    a = a * a

    DoEvents
Loop

End Sub
