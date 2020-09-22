VERSION 5.00
Begin VB.Form fLogin 
   BorderStyle     =   3  'Pevný dialog
   Caption         =   "Login"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "fLogon.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Cenrovat vlastník
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox tServer 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "."
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox tUser 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox tPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Doprava
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Doprava
      Caption         =   "Admin Account"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Doprava
      Caption         =   "IP address"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "fLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objSWbemLocator As WbemScripting.SWbemLocator
Dim objSWbemServices As WbemScripting.SWbemServices

Private Sub cmdOK_Click()
   On Error GoTo err_login
   Screen.MousePointer = vbHourglass
   Dim objSWbemLocator
   Dim objSWbemServices
   
   Set objSWbemLocator = New WbemScripting.SWbemLocator
   Set objSWbemServices = objSWbemLocator.ConnectServer(tServer.Text, "root\cimv2", tUser.Text, tPass.Text)

   Screen.MousePointer = vbNormal


   fMain.sServer = tServer.Text
   fMain.sAccount = tUser.Text
   fMain.sPassword = tPass.Text
   Unload Me
   fMain.cMenu.ListIndex = 23
   Call fMain.Command1_Click
   

   Exit Sub
err_login:
   MsgBox "Chyba: " & Err.Number & Err.Description

End Sub

