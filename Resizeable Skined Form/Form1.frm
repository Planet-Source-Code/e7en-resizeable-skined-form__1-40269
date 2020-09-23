VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Press Enter to quit"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Info 
      BackStyle       =   0  'Transparent
      Height          =   2115
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================================
'======================================================================================
'                              Resizeable, Skinable Form
'                              ----------  -------- ----
'                                  Created By §e7eN
'
'   Notes: Yes this is not the greatest Method but effective none the less.
'          This is not the final result. The final result will support button
'          Pictures wich have a Up, Down and Over state, Propper Move Form Method
'          plus more. If anyone would be interested in Purchasing the Final result
'          or Custom Designed software then Email me at Hate_114@hotmail.com.
'
'          Hope This Has been usefull for you,
'                                               §e7eN
'
'======================================================================================
'======================================================================================

Dim Skin As ClsSkinForm

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()

Set Skin = New ClsSkinForm
Info.Caption = "Here is a little Demo." & vbCrLf & "If you like this please remember to Vote" & vbCrLf & vbCrLf & "Enjoy, §e7eN"

Me.Show
Skin.SkinForm Me, App.Path & "\Skin\Settings.ini"

End Sub

Private Sub Form_Resize()
Skin.GFX
End Sub


