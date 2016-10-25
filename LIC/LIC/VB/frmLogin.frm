VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   4170
   ClientLeft      =   15375
   ClientTop       =   8355
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   278
   ScaleLeft       =   500
   ScaleMode       =   0  'User
   ScaleTop        =   500
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdok_Click()
    'check for correct password
    If txtPassword = "password" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

