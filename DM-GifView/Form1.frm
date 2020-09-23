VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "DM Gif View"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   390
      Left            =   1800
      TabIndex        =   19
      Top             =   3780
      Width           =   1650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&About"
      Height          =   390
      Left            =   75
      TabIndex        =   18
      Top             =   3780
      Width           =   1650
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   5490
      Max             =   1000
      TabIndex        =   16
      Top             =   3300
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   6795
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Gif"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Gif File"
      Height          =   390
      Left            =   90
      TabIndex        =   12
      Top             =   3270
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   390
      Left            =   3570
      TabIndex        =   10
      Top             =   3270
      Width           =   1650
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2625
      Top             =   4425
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   105
      TabIndex        =   1
      Top             =   150
      Width           =   7395
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3120
         TabIndex        =   11
         Top             =   1935
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Index           =   3
         Left            =   3900
         TabIndex        =   9
         Top             =   1185
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Frames"
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   8
         Top             =   1185
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Index           =   2
         Left            =   3900
         TabIndex        =   7
         Top             =   930
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Index           =   1
         Left            =   3900
         TabIndex        =   6
         Top             =   660
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Index           =   0
         Left            =   3900
         TabIndex        =   5
         Top             =   405
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   4
         Top             =   930
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   3
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Verision"
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   405
         Width           =   555
      End
      Begin VB.Image GifView 
         Height          =   1890
         Left            =   225
         Top             =   330
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   390
      Left            =   1830
      TabIndex        =   0
      Top             =   3270
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   4185
      Picture         =   "Form1.frx":0000
      Top             =   5205
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblnum 
      AutoSize        =   -1  'True
      Caption         =   "."
      Height          =   195
      Left            =   6810
      TabIndex        =   17
      Top             =   3285
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Change Delay"
      Height          =   195
      Left            =   5625
      TabIndex        =   15
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   930
      TabIndex        =   14
      Top             =   2820
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Filename"
      Height          =   195
      Left            =   135
      TabIndex        =   13
      Top             =   2835
      Width           =   630
   End
   Begin VB.Image GifSrc 
      Height          =   1035
      Index           =   1
      Left            =   360
      Top             =   4515
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Hello eveyone this my attempt at playing Animated Gifs in vb this is quit
' a cheap way of doing it but it does seems to work quite well
' How it works ok from the code below you will see that I am Extraacting each
' Gif from a Animated gif by adding a header and the end then loading them back into
' a image control array and useing a timer to play them back very simple but it works

'Anyway I hope you like my code please vote if you do

Private Type Gif_Stuff
    ID As String * 3
    Verision As String * 3
    GifWidth As Integer
    GifHeight As Integer
    
End Type

Dim TGif As Gif_Stuff
Dim Gif_Count As Integer
Dim IsLoaded As Boolean


Function AddSlash(LzPath As String) As String
    If Right(LzPath, 1) <> "\" Then AddSlash = LzPath & "\" Else AddSlash = LzPath
    
End Function
Public Sub LoadGif(lzFilename As String)
Dim GifData As String, GifHead As String, GifEnd As String, _
NewGifData As String, ExtPath As String


Dim nPos, lPos, I As Integer
Dim TFile1 As Long, TFile2 As Long


    GifEnd = Chr(0) & "!ù"
    TFile1 = FreeFile
    
    If Dir(lzFilename) = "" Then ' Checks if the file is there
        MsgBox "File " & lzFilename & " not found", vbCritical, "File Not Found"
        Exit Sub
    Else
        ' I use this to just get some info from the gif file
        Open lzFilename For Binary As #1
            Get #1, , TGif
        Close #1
    End If
    
        ' Just does a check to see if the gif file is inviald or not
        ' it must return teh string "GIF" to work
        If TGif.ID <> "GIF" Then
            MsgBox "Inviald Gif Format", vbCritical, "Error"
            Exit Sub
        Else
            Open lzFilename For Binary As #TFile1
                GifData = Space(LOF(1)) ' Create some room for our file
                Get #TFile1, , GifData   ' Get all the data in one go
            Close #TFile1
        End If
        
        nPos = InStr(1, GifData, GifEnd) + Len(GifEnd) - 2 ' Find end of first gif
        GifHead = Left(GifData, nPos) ' Get the header for the new gif
        lPos = nPos + 2
        Gif_Count = 1 ' Set gif file counter to 1
        
        Do
            ' Start Extracting the files of the files we need
            nPos = InStr(lPos, GifData, GifEnd) + Len(GifEnd)
            If nPos > Len(GifEnd) Then
            
            ' Check and create a Tmp folder to extract all the gifs to
            If Dir(AddSlash(App.Path) & "giftmp", vbDirectory) = "" Then MkDir AddSlash(App.Path) & "giftmp"
                ExtPath = AddSlash(App.Path) & AddSlash("giftmp")
                TFile2 = FreeFile
                
                Open ExtPath & Gif_Count & ".gif" For Binary As #TFile2
                    NewGifData = GifHead & Mid(GifData, lPos - 1, nPos - lPos)
                    Put #TFile2, , NewGifData
                Close #TFile2
                
                lPos = nPos
                Gif_Count = Gif_Count + 1
            End If
        Loop Until nPos = Len(GifEnd)
        ' Just clear some stuff we don't need any more
        GifData = ""
        NewGifData = ""
        GifHead = ""
        GifEnd = ""
        lPos = 0
        nPos = 0
        
        On Error Resume Next
        For I = 1 To Gif_Count
            Load GifSrc(I)
            ' Now we load all our new files back into a image array
            GifSrc(I).Picture = LoadPicture(ExtPath & I & ".gif")
            Kill ExtPath & I & ".gif" ' Remove all the files
            RmDir ExtPath   ' And also remove the tmp folder to
        Next
        I = 1
        GifView.Picture = Image1.Picture
        
        ' Lets just show some info about the gif file
        Label2(0).Caption = TGif.Verision
        Label2(1).Caption = TGif.GifWidth
        Label2(2).Caption = TGif.GifHeight
        Label2(3).Caption = Gif_Count - 1
        Gif_Count = Gif_Count - 1 ' Take of the extra one
        IsLoaded = True
        
        
End Sub

Private Sub Command1_Click()
    If IsLoaded Then
        Timer1.Enabled = True
        Exit Sub
    Else
        Timer1.Enabled = False
    End If

End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    
End Sub

Private Sub Command3_Click()
Dim FExt As String
    CDialog.ShowOpen
    CDialog.DialogTitle = "Open Animated Gif"
    FExt = Right(UCase(CDialog.FileName), 3)
    If Len(FExt) = 0 Then Exit Sub
    If FExt <> "GIF" Then
        MsgBox "This program will only open Animated Gof Files", vbInformation, "Inviald File"
        Exit Sub
    Else
        LoadGif CDialog.FileName
        Label5.Caption = CDialog.FileName
    End If
    FExt = ""
    
End Sub

Private Sub Command4_Click()
    MsgBox "DM Animated Gif Viewer by Ben Jones" & _
    vbCrLf & vbCrLf & Space(8) & "Please Vote for me", vbInformation, "About....."
    
End Sub

Private Sub Form_Load()
    HScroll1.Value = 100
    GifView.Picture = Image1.Picture
    
End Sub

Private Sub HScroll1_Change()
    If HScroll1.Value <= 10 Then HScroll1.Value = 10
    lblnum.Caption = HScroll1.Value
    Timer1.Interval = HScroll1.Value
    
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
    
End Sub

Private Sub Timer1_Timer()
Static Cnt As Integer
On Error Resume Next
    Cnt = Cnt + 1
    If Cnt = Gif_Count Then Cnt = 1
    GifView.Picture = GifSrc(Cnt).Picture
    Label3.Caption = "Current Frame = " & Cnt
End Sub
