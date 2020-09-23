VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Auto E-mail using MAPI Controls"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Text            =   "Attch. Name"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Attachment path"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Subject"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter e-mail address"
      Top             =   1080
      Width           =   2655
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   840
      Top             =   5280
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
      Left            =   240
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "by littlegreenrussian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "For an e-mail to be sent automatically, simply edit the code to replace the Text1, Text2 etc with your own entries."
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Auto - Email using MAPI Controls."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'user clicks send
        On Error GoTo mailerr: 'go to the error handling bit If there is an error
            MAPISession1.SignOn 'sign on


    If MAPISession1.SessionID <> 0 Then 'signed on


        With MAPIMessages1
                .SessionID = MAPISession1.SessionID
                .Compose 'start a new message
            .AttachmentName = Text5 'attachment name
                .AttachmentPathName = Text4 ' attachment path (get this from the text box or a default dirrectory)
                .RecipAddress = Text1 'set the receiver's email To the one they specified (again, text box or a default address)
            .MsgSubject = Text2 'set the subject
            .MsgNoteText = Text3 'message text
                .Send False 'don't display a dialog saying it was sent
                    
                    End With
                        Exit Sub
                            End If
mailerr:                 'error handling
                    MsgBox "Error " & Err.Description
                End Sub


Private Sub Command2_Click() 'clear clicked
    
    Text1.Text = "" 'clear the text boxes
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""

End Sub
