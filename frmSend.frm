VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmSend 
   BackColor       =   &H00400040&
   Caption         =   "Send"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "Send"
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSend_Click()
Dim strEMail As String
    Dim strMessage As String

    frmSend.MAPIMessages1.MsgIndex = -1
    
    'The next line is to set the subject
    frmSend.MAPIMessages1.MsgSubject = "Test Mail"
    
    frmSend.MAPIMessages1.RecipIndex = 0
    frmSend.MAPIMessages1.RecipType = 1
    
    'Put Mail address at end of next line
    strEMail = InputBox("Please insert E-Mail address that will receive this test", "E-Mail Address")
    If strEMail = "" Then
        GoTo EndOfSub
    End If
    frmSend.MAPIMessages1.RecipDisplayName = strEMail
    
    
    'Put Message in strMessage ready for next line
    strMessage = "This Is A Test..."
    frmSend.MAPIMessages1.MsgNoteText = strMessage
    
    
    frmSend.MAPIMessages1.RecipType = mapToList
    frmSend.MAPISession1.NewSession = False
    frmSend.MAPISession1.DownLoadMail = False
    
    
    frmSend.MAPISession1.SignOn
    frmSend.MAPIMessages1.SessionID = frmSend.MAPISession1.SessionID
    frmSend.MAPIMessages1.AddressResolveUI = True
    
    frmSend.MAPIMessages1.ResolveName
    
   
    frmSend.MAPIMessages1.Send
  
    frmSend.MAPISession1.SignOff
EndOfSub:
End Sub
