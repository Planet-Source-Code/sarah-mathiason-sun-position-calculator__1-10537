VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFlash.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H009C9C9C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sun Position Calculator"
   ClientHeight    =   7545
   ClientLeft      =   2970
   ClientTop       =   3450
   ClientWidth     =   5970
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraHelp 
      BackColor       =   &H009C9C9C&
      Height          =   7365
      Left            =   105
      TabIndex        =   69
      Top             =   105
      Visible         =   0   'False
      Width           =   5790
      Begin VB.CheckBox Check7 
         BackColor       =   &H0098A8A8&
         Caption         =   "Azimuth"
         DownPicture     =   "Form2.frx":030A
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   3570
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H0098A8A8&
         Caption         =   "Altitude"
         DownPicture     =   "Form2.frx":2FAE
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   3255
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H0098A8A8&
         Caption         =   "Declination"
         DownPicture     =   "Form2.frx":5C52
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2940
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0098A8A8&
         Caption         =   "Distance"
         DownPicture     =   "Form2.frx":88F6
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2625
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0098A8A8&
         Caption         =   "Coordinates"
         DownPicture     =   "Form2.frx":B59A
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2310
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0098A8A8&
         Caption         =   "Solar Time"
         DownPicture     =   "Form2.frx":E23E
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1995
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0098A8A8&
         Caption         =   "Day Number"
         DownPicture     =   "Form2.frx":10EE2
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1680
         Value           =   2  'Grayed
         Width           =   1400
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0098A8A8&
         Height          =   4425
         Left            =   1575
         MultiLine       =   -1  'True
         TabIndex        =   87
         Top             =   945
         Width           =   4110
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   210
         ScaleHeight     =   1275
         ScaleWidth      =   2745
         TabIndex        =   83
         Top             =   5460
         Width           =   2745
         Begin VB.Label Label17 
            BackColor       =   &H009C9C9C&
            BackStyle       =   0  'Transparent
            Caption         =   "ICQ: 11649134"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   225
            Left            =   1365
            TabIndex        =   85
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label Label16 
            BackColor       =   &H009C9C9C&
            BackStyle       =   0  'Transparent
            Caption         =   "Written by Andrew Mathiason"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   105
            TabIndex        =   84
            Top             =   210
            Width           =   2535
         End
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWave1 
         Height          =   1275
         Left            =   2835
         TabIndex        =   82
         Top             =   5460
         Width           =   2745
         _cx             =   4199146
         _cy             =   4196553
         Movie           =   "e:\vb projects\sunangleiii\email.swf"
         Src             =   "e:\vb projects\sunangleiii\email.swf"
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   -1  'True
         BGColor         =   ""
         SWRemote        =   ""
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Done"
         Height          =   300
         Left            =   2310
         Picture         =   "Form2.frx":13B86
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   6930
         Width           =   1400
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Distance"
         DownPicture     =   "Form2.frx":1682A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":194CE
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   4095
         Width           =   1400
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4725
         Picture         =   "Form2.frx":1C172
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   525
         Width           =   1000
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3675
         Picture         =   "Form2.frx":1EE16
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   525
         Width           =   1000
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Longitude"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         Picture         =   "Form2.frx":21ABA
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   525
         Width           =   1000
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Latitude"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1575
         Picture         =   "Form2.frx":2475E
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   525
         Width           =   1000
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H009C9C9C&
         Caption         =   "Calculate"
         DownPicture     =   "Form2.frx":27402
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         MaskColor       =   &H009C9C9C&
         Picture         =   "Form2.frx":2A0A6
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   5040
         Width           =   1400
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Output Info:"
         DownPicture     =   "Form2.frx":2CD4A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":2F9EE
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1365
         Width           =   1400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export Motion"
         DownPicture     =   "Form2.frx":32692
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":35336
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   4620
         Width           =   1400
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Day Calc"
         DownPicture     =   "Form2.frx":37FDA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":3AC7E
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   945
         Width           =   1400
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Time Calc"
         DownPicture     =   "Form2.frx":3D922
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":405C6
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   525
         Width           =   1400
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   5775
         Y1              =   6825
         Y2              =   6825
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help:  Click on the replica of the button to read the function"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.CheckBox chkN 
      BackColor       =   &H0098A8A8&
      Caption         =   "Day Number"
      DownPicture     =   "Form2.frx":4326A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1575
      Width           =   1400
   End
   Begin VB.CheckBox chkLSoT 
      BackColor       =   &H0098A8A8&
      Caption         =   "Solar Time"
      DownPicture     =   "Form2.frx":45F0E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1890
      Width           =   1400
   End
   Begin VB.CheckBox chkXYZ 
      BackColor       =   &H0098A8A8&
      Caption         =   "Coordinates"
      DownPicture     =   "Form2.frx":48BB2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2205
      Width           =   1400
   End
   Begin VB.CheckBox chkDistance 
      BackColor       =   &H0098A8A8&
      Caption         =   "Distance"
      DownPicture     =   "Form2.frx":4B856
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   1400
   End
   Begin VB.CheckBox chkD 
      BackColor       =   &H0098A8A8&
      Caption         =   "Declination"
      DownPicture     =   "Form2.frx":4E4FA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2835
      Width           =   1400
   End
   Begin VB.CheckBox chkAl 
      BackColor       =   &H0098A8A8&
      Caption         =   "Altitude"
      DownPicture     =   "Form2.frx":5119E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3150
      Width           =   1400
   End
   Begin VB.CheckBox chkAz 
      BackColor       =   &H0098A8A8&
      Caption         =   "Azimuth"
      DownPicture     =   "Form2.frx":53E42
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3465
      Width           =   1400
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
      Height          =   1065
      Left            =   105
      TabIndex        =   86
      Top             =   6405
      Width           =   1380
      _cx             =   4196738
      _cy             =   4196183
      Movie           =   "e:\vb projects\sunangleiii\orbit.swf"
      Src             =   "e:\vb projects\sunangleiii\orbit.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      DownPicture     =   "Form2.frx":56AE6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Picture         =   "Form2.frx":5978A
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   840
      Width           =   1400
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Motion"
      DownPicture     =   "Form2.frx":5C42E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Picture         =   "Form2.frx":5F0D2
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   5460
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.TextBox txtMF 
      BackColor       =   &H00505050&
      ForeColor       =   &H0000FFFF&
      Height          =   6945
      Left            =   5985
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   61
      Top             =   525
      Width           =   3480
   End
   Begin VB.Frame fraLong 
      BackColor       =   &H009C9C9C&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   1680
      TabIndex        =   17
      Top             =   525
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CommandButton cmdLongCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   2310
         Picture         =   "Form2.frx":61D76
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1785
         Width           =   1400
      End
      Begin VB.TextBox txtLL 
         Alignment       =   2  'Center
         BackColor       =   &H006C6C6C&
         ForeColor       =   &H00BFFFFF&
         Height          =   300
         Left            =   420
         TabIndex        =   32
         Text            =   "106"
         Top             =   735
         Width           =   1400
      End
      Begin VB.CommandButton cmdLongDone 
         Caption         =   "Done"
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":64A1A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1785
         Width           =   1400
      End
      Begin VB.OptionButton opnEW 
         BackColor       =   &H009CA89C&
         Caption         =   "West"
         Height          =   300
         Index           =   1
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   945
         Value           =   -1  'True
         Width           =   1400
      End
      Begin VB.OptionButton opnEW 
         BackColor       =   &H009CA89C&
         Caption         =   "East"
         Height          =   300
         Index           =   0
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   630
         Width           =   1400
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3780
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set Longitude Degrees"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   15
         TabIndex        =   36
         Top             =   0
         Width           =   3795
      End
   End
   Begin VB.Frame fraTime 
      BackColor       =   &H009C9C9C&
      Height          =   3585
      Left            =   1680
      TabIndex        =   44
      Top             =   525
      Visible         =   0   'False
      Width           =   3795
      Begin VB.ComboBox cboGMT 
         BackColor       =   &H006C6C6C&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BFFFFF&
         Height          =   345
         Left            =   1200
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "G.M.T. -6:00"
         Top             =   870
         Width           =   1450
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H009C9C9C&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1155
         TabIndex        =   50
         Top             =   2490
         Width           =   1515
         Begin VB.OptionButton opnYN 
            BackColor       =   &H00A89C9C&
            Caption         =   "No"
            Height          =   330
            Index           =   1
            Left            =   780
            Style           =   1  'Graphical
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   650
         End
         Begin VB.OptionButton opnYN 
            BackColor       =   &H00A89C9C&
            Caption         =   "Yes"
            Height          =   330
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   0
            Width           =   650
         End
      End
      Begin VB.TextBox txtMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006C6C6C&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BFFFFF&
         Height          =   330
         Left            =   1965
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1590
         Width           =   700
      End
      Begin VB.TextBox txtHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006C6C6C&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BFFFFF&
         Height          =   330
         Left            =   1215
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1590
         Width           =   700
      End
      Begin VB.CommandButton cmdTimeCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   2310
         Picture         =   "Form2.frx":676BE
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3150
         Width           =   1400
      End
      Begin VB.CommandButton cmdTimeDone 
         Caption         =   "Done"
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":6A362
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3150
         Width           =   1400
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H009C9C9C&
         Caption         =   "Time Zone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   57
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H009C9C9C&
         Caption         =   ":"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1830
         TabIndex        =   56
         Top             =   1635
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H009C9C9C&
         Caption         =   "Time:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1215
         TabIndex        =   55
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H009C9C9C&
         Caption         =   "Daylight Savings Time?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   54
         Top             =   1950
         Width           =   1455
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   3780
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set Time Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   3795
      End
   End
   Begin VB.TextBox txtDist 
      Alignment       =   2  'Center
      BackColor       =   &H006C6C6C&
      ForeColor       =   &H00BFFFFF&
      Height          =   300
      Left            =   105
      TabIndex        =   43
      Text            =   "100"
      Top             =   4725
      Width           =   1400
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4830
      Picture         =   "Form2.frx":6D006
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1000
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H009C9C9C&
      Height          =   4425
      Left            =   1680
      TabIndex        =   37
      Top             =   525
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CommandButton cmdDateCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   2310
         Picture         =   "Form2.frx":6FCAA
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3990
         Width           =   1400
      End
      Begin VB.CommandButton cmdDateDone 
         Caption         =   "Done"
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":7294E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3990
         Width           =   1400
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2610
         Left            =   420
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   630
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   4604
         _Version        =   393216
         ForeColor       =   12582911
         BackColor       =   10263708
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthBackColor  =   10263708
         ShowToday       =   0   'False
         StartOfWeek     =   22740993
         TitleBackColor  =   10263708
         TrailingForeColor=   8421504
         CurrentDate     =   36730
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   3780
         Y1              =   3885
         Y2              =   3885
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006C6C6C&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BFFFFF&
         Height          =   300
         Left            =   1155
         TabIndex        =   40
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   10
         TabIndex        =   39
         Top             =   0
         Width           =   3795
      End
   End
   Begin VB.CommandButton cmdLong 
      Caption         =   "Longitude"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      Picture         =   "Form2.frx":755F2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1000
   End
   Begin VB.Frame fraLat 
      BackColor       =   &H009C9C9C&
      Height          =   2220
      Left            =   1680
      TabIndex        =   26
      Top             =   525
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CommandButton cmdLatCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   2310
         Picture         =   "Form2.frx":78296
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1785
         Width           =   1400
      End
      Begin VB.CommandButton cmdLatDone 
         Caption         =   "Done"
         Height          =   300
         Left            =   105
         Picture         =   "Form2.frx":7AF3A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1785
         Width           =   1400
      End
      Begin VB.OptionButton opnNS 
         BackColor       =   &H009CA89C&
         Caption         =   "North"
         Height          =   300
         Index           =   0
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   630
         Value           =   -1  'True
         Width           =   1400
      End
      Begin VB.OptionButton opnNS 
         BackColor       =   &H009CA89C&
         Caption         =   "South"
         Height          =   300
         Index           =   1
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   945
         Width           =   1400
      End
      Begin VB.TextBox txtLat 
         Alignment       =   2  'Center
         BackColor       =   &H006C6C6C&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BFFFFF&
         Height          =   300
         Left            =   420
         TabIndex        =   27
         Text            =   "53"
         Top             =   735
         Width           =   1400
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set Latitude Degrees"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   15
         TabIndex        =   34
         Top             =   0
         Width           =   3795
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3780
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.CommandButton cmdLat 
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      Picture         =   "Form2.frx":7DBDE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1000
   End
   Begin VB.CheckBox chkSN 
      BackColor       =   &H0098A8A8&
      Caption         =   "Solar Noon"
      DownPicture     =   "Form2.frx":80882
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2835
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3780
      Picture         =   "Form2.frx":83526
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1000
   End
   Begin VB.PictureBox picHold 
      Height          =   495
      Left            =   735
      Picture         =   "Form2.frx":861CA
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   25
      Top             =   11130
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcAllDay 
      Appearance      =   0  'Flat
      BackColor       =   &H009C9C9C&
      Caption         =   "Calculate"
      DownPicture     =   "Form2.frx":9660E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaskColor       =   &H009C9C9C&
      Picture         =   "Form2.frx":992B2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5880
      Width           =   1400
   End
   Begin VB.CommandButton cmdCalcIndTime 
      Caption         =   "Calculate"
      DownPicture     =   "Form2.frx":9BF56
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      Picture         =   "Form2.frx":9EBFA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.OptionButton opnCalc 
      BackColor       =   &H00A8A898&
      Caption         =   "Time Calc."
      DownPicture     =   "Form2.frx":A189E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   105
      Picture         =   "Form2.frx":A4542
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   1400
   End
   Begin VB.OptionButton opnCalc 
      BackColor       =   &H00A8A898&
      Caption         =   "Day Calc."
      DownPicture     =   "Form2.frx":A71E6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   105
      Picture         =   "Form2.frx":A9E8A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Value           =   -1  'True
      Width           =   1400
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10875
      Top             =   675
   End
   Begin VB.TextBox txtElev 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10875
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "1392"
      Top             =   1035
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picAZ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H007C7C7C&
      Height          =   1800
      Left            =   1575
      ScaleHeight     =   1740
      ScaleWidth      =   2745
      TabIndex        =   59
      Top             =   525
      Width           =   2800
      Begin VB.Line Line12 
         BorderColor     =   &H006C6C6C&
         X1              =   0
         X2              =   2730
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line Line11 
         BorderColor     =   &H006C6C6C&
         X1              =   0
         X2              =   2730
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line10 
         BorderColor     =   &H006C6C6C&
         X1              =   467
         X2              =   467
         Y1              =   1680
         Y2              =   0
      End
      Begin VB.Line Line9 
         BorderColor     =   &H006C6C6C&
         X1              =   1867
         X2              =   1867
         Y1              =   1680
         Y2              =   0
      End
      Begin VB.Line Line8 
         BorderColor     =   &H006C6C6C&
         X1              =   2333
         X2              =   2333
         Y1              =   1680
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H006C6C6C&
         X1              =   933
         X2              =   933
         Y1              =   1680
         Y2              =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H006C6C6C&
         X1              =   1400
         X2              =   1400
         Y1              =   1680
         Y2              =   0
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   2730
         Y1              =   900
         Y2              =   900
      End
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5355
      Picture         =   "Form2.frx":ACB2E
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2520
      Width           =   150
   End
   Begin VB.TextBox txtScroll 
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5070
      Left            =   1575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2415
      Width           =   4230
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4830
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkH 
      BackColor       =   &H0098A8A8&
      Caption         =   "Hour Angle"
      DownPicture     =   "Form2.frx":BCF72
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2835
      Width           =   1400
   End
   Begin VB.CheckBox chkET 
      BackColor       =   &H0098A8A8&
      Caption         =   "Eqn of Time"
      DownPicture     =   "Form2.frx":BFC16
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1400
   End
   Begin VB.Label lblTimePos 
      Alignment       =   2  'Center
      BackColor       =   &H009C9C9C&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BFFFFF&
      Height          =   285
      Left            =   4410
      TabIndex        =   67
      Top             =   1890
      Width           =   1365
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H008C8C8C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4410
      TabIndex        =   66
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H008C8C8C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4410
      TabIndex        =   65
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label lblyPos 
      Alignment       =   2  'Center
      BackColor       =   &H009C9C9C&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BFFFFF&
      Height          =   285
      Left            =   4410
      TabIndex        =   64
      Top             =   945
      Width           =   1365
   End
   Begin VB.Label Label10 
      BackColor       =   &H006C6C6C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is a Lightwave 5.6 motion file."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   5985
      TabIndex        =   63
      Top             =   210
      Width           =   3270
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H008C8C8C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Output Info:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   24
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008C8C8C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   23
      Top             =   4410
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Elevation: "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10875
      TabIndex        =   21
      Top             =   795
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Andrew Mathiason
'ICQ:11649134
'bullhorn@home.com
'August 8 / 2000

'there are 2 shockwave flash controls
'you need to change the path to the appropriate spot on your harddrive
'where the file email.swf and orbit.swf

Dim Month(11) ' holds month names
Dim LST 'local standard time
Dim LSoT 'local solar time
Dim Days(11) 'holds the amount of days in each month
Dim N ' day number
Dim Rad ' degrees to radians
Dim ET ' value from equation of time
Dim LSTM ' local standard time meridian
Dim B ' a temp variable
Dim LL ' latitude
Dim H ' hour angle
Dim Temp1 ' temp variable
Dim Temp2 'temp variable
Dim D ' declination
Dim Al 'altitude angle
Dim Az 'azimuth angle
Dim L ' longitude
Dim SS
Dim SR
Dim PrevDay As Boolean
Dim LSoT2
Dim Temp3
Dim Temp4
Dim Xco
Dim Yco
Dim Zco
Dim S1
Dim Distance
Dim Q
Dim XXX
Dim YYY
Dim ZZZ
Dim OldLst
Dim AllDAy As Boolean
Dim Frm As Integer

Private Sub cmdCalcAllDay_Click()
cmdExport.Visible = False
Frm = -1
picAZ.Cls
AllDAy = True
Q = 0
L = Val(txtLat.Text)
LL = Val(txtLL.Text)
If opnEW(1).Value = True Then LL = LL * -1
txtScroll.Text = txtScroll.Text & vbNewLine & "-----------------------------------------------------------------"
txtScroll.Text = txtScroll.Text & vbNewLine & "Date : " & lblDate.Caption
For i = 0 To 23
    LST = i * 60
    If opnYN(0).Value = True Then
        LST = LST - 60
    End If
    EquationOfTime
    ConvertToSolarTime
    HourAngle
    Declination
    AltitudeAngle
    AzimuthAngle
picAZ.ScaleWidth = 1400
picAZ.ScaleHeight = Val(txtDist.Text) * 2
    CalcCoords
    MotionFile
    txtScroll.Text = txtScroll.Text & vbNewLine & Format(Int(LST / 60), "00") & ":" & Format(Int(LST Mod 60), "00")
    UpdateInfo
Next
txtScroll.SelStart = Len(txtScroll.Text)
txtScroll.SelLength = 0
End Sub

Private Sub cmdCalcIndTime_Click()
cmdExport.Visible = False
If AllDAy = True Then
    AllDAy = False
picAZ.Cls
End If
Q = 0
PrevDay = False
L = Val(txtLat.Text)
LL = Val(txtLL.Text)
If opnEW(1).Value = True Then LL = LL * -1
txtScroll.Text = txtScroll.Text & vbNewLine & "-----------------------------------------------------------------"
txtScroll.Text = txtScroll.Text & vbNewLine & "Date : " & lblDate.Caption
LST = Val(txtHour.Text) * 60 + Val(txtMin.Text)
txtScroll.Text = txtScroll.Text & vbNewLine & " - Local Standard Time : " & Format(txtHour.Text, "00") & ":" & Format(txtMin.Text, "00") & " = " & LST & " Minutes"
If opnYN(0).Value = True Then
    LST = LST - 60
    txtScroll.Text = txtScroll.Text & vbNewLine & " -  Adjusting for D.S.T.  : " & Format(Val(txtHour.Text) - 1, "00") & ":" & Format(txtMin.Text, "00") & " = " & LST & " Minutes"
End If
EquationOfTime
ConvertToSolarTime
HourAngle
Declination
AltitudeAngle
AzimuthAngle
picAZ.ScaleWidth = 1400
picAZ.ScaleHeight = Val(txtDist.Text) * 2
CalcCoords
UpdateInfo
End Sub
Private Sub UpdateInfo()
If chkN.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Day Number : " & N
End If
If chkLSoT.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Local Solar Time: " _
    & Format(Temp1, "00") & ":" & Format(Temp2, "00")
End If
If chkET.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Equation of Time Result = " & Format(ET, "###.00")
End If
If chkH.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Hour Angle = " & Format(H, "###.00") & "°"
End If
If chkD.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Declination = " & Format(D, "###.00") & "°"
End If
If chkAl.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Altitude Angle = " & Format(Al, "###.00") & "°"
End If
If chkAz.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Azimuth Angle = " & Format(Az, "###.00") & "°"
End If
If chkSN.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Solar Noon = " & Format(Temp3, "00") & ":" & Format(Temp4, "00") & " = " & Int(LSoT2) & " Minutes"
End If
If chkDistance.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - Distance to Sun = " & Val(txtDist.Text)
End If
If chkXYZ.Value = vbChecked Then
    txtScroll.Text = txtScroll.Text & vbNewLine & " - (x,y,z) = ( " & Format(Xco, "######.00") & ", " & Format(Yco, "######.00") & ", " & Format(Zco, "######.00") & ")"
End If

txtScroll.SelStart = Len(txtScroll.Text)
txtScroll.SelLength = 0
End Sub
Private Sub CalcCoords()

If AllDAy = True Then
    Q = Q + 1
End If
Dim XTemp
Distance = Val(txtDist.Text)
Yco = Sin(Al / Rad) * Distance
XTemp = (Cos(Al / Rad) * Distance)
Zco = Sin(Az / Rad) * XTemp
Xco = Cos(Az / Rad) * XTemp
If Q = 0 Then
    picAZ.PSet (LST, (-1 * Xco + picAZ.ScaleHeight / 2)), RGB(252, 63, 63)
    picAZ.PSet (LST, (-1 * Yco + picAZ.ScaleHeight / 2)), RGB(0, 126, 0)
    picAZ.PSet (LST, (-1 * Zco + picAZ.ScaleHeight / 2)), RGB(0, 63, 126)
End If
If Q > 1 Then
    picAZ.Line (LST, (-1 * Xco + picAZ.ScaleHeight / 2))-(OldLst, (-1 * XXX + picAZ.ScaleHeight / 2)), RGB(252, 63, 63)
    picAZ.Line (LST, (-1 * Yco + picAZ.ScaleHeight / 2))-(OldLst, (-1 * YYY + picAZ.ScaleHeight / 2)), RGB(0, 126, 0)
    picAZ.Line (LST, (-1 * Zco + picAZ.ScaleHeight / 2))-(OldLst, (-1 * ZZZ + picAZ.ScaleHeight / 2)), RGB(0, 63, 126)
End If
OldLst = LST
ZZZ = Zco
YYY = Yco
XXX = Xco
End Sub

Private Sub AzimuthAngle()
X = (Sin(Al / Rad) * Sin(L / Rad) - Sin(D / Rad)) / (Cos(Al / Rad) * Cos(L / Rad))
Az = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
Az = Az * Rad
If H < 0 Then Az = Az * -1
End Sub

Private Sub AltitudeAngle()
X = Cos(L / Rad) * Cos(D / Rad) * Cos(H / Rad) + Sin(L / Rad) * Sin(D / Rad)
Al = Atn(X / Sqr(-X * X + 1))
Al = Al * Rad
End Sub

Private Sub Declination()
D = 23.45 * Sin((360 / 365 * (284 + N)) / Rad)
End Sub

Private Sub EquationOfTime()
B = 360 * (N - 81) / 365
ET = 9.87 * Sin(2 * B / Rad) - 7.53 * Cos(B / Rad) - 1.5 * Sin(B / Rad)
End Sub

Private Sub ConvertToSolarTime()
If cboGMT.Text = "G.M.T." Then
    LSTM = 0
End If
For i = 1 To 13
    If cboGMT.Text = "G.M.T. +" & i & ":00" Then
        LSTM = i
    End If
Next
For i = 1 To 12
    If cboGMT.Text = "G.M.T. -" & i & ":00" Then
        LSTM = -1 * i
    End If
Next
LSoT = LST + 4 * (LL - (LSTM * 15)) + ET
Temp1 = Int((LSoT / 60))
Temp2 = LSoT Mod 60
If Int(LSoT) = 719 Then
    Temp1 = 12
    Temp2 = 0
End If
If LSoT < 0 Then
    LSoT = (1440 + LSoT)
    Temp1 = Int((LSoT / 60))
    Temp2 = LSoT Mod 60
    PrevDay = True
End If
End Sub

Private Sub HourAngle()
Dim MinNoon
Dim Noon
Noon = 60 * 12
MinNoon = Int(Noon - LSoT)
H = MinNoon / 4
End Sub

Private Sub cmdClear_Click()
txtScroll.Text = ""
End Sub

Private Sub cmdDate_Click()
fraTime.Visible = False
fraLong.Visible = False
fraLat.Visible = False
fraDate.Visible = True
End Sub

Private Sub cmdDateCancel_Click()
fraDate.Visible = False
End Sub

Private Sub cmdDateDone_Click()
fraDate.Visible = False
End Sub

Private Sub cmdExport_Click()
'save the motion file using common dialog
    Dim fltxt As String
        fltxt = txtMF.Text
    CommonDialog1.Filter = "Lightwave Motion Files|*.mot"
    On Error GoTo stp
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Open (CommonDialog1.FileName) For Output As #1
            Print #1, fltxt
        Close #1
    End If
stp:
End Sub

Private Sub cmdHelp_Click()
fraHelp.Visible = True
End Sub

Private Sub cmdLat_Click()
fraTime.Visible = False
fraDate.Visible = False
fraLong.Visible = False
fraLat.Visible = True
End Sub

Private Sub cmdLatCancel_Click()
fraLat.Visible = False
End Sub

Private Sub cmdLatDone_Click()
fraLat.Visible = False
End Sub

Private Sub cmdLong_Click()
fraTime.Visible = False
fraDate.Visible = False
fraLat.Visible = False
fraLong.Visible = True
End Sub

Private Sub cmdLongCancel_Click()
fraLong.Visible = False
End Sub

Private Sub cmdLongDone_Click()
fraLong.Visible = False
End Sub

Private Sub cmdTime_Click()
fraDate.Visible = False
fraLong.Visible = False
fraLat.Visible = False
fraTime.Visible = True
End Sub

Private Sub cmdTimeCancel_Click()
fraTime.Visible = False
End Sub

Private Sub cmdTimeDone_Click()
fraTime.Visible = False
End Sub
'**************************************************************
'everything in the section beneath is for the help information
'**************************************************************
Private Sub Command1_Click()
Text1.Text = "Individual Time Calculation:" & vbNewLine & vbNewLine & "1.) Set Location and Time Information using the green buttons along the top." & vbNewLine & vbNewLine & _
"2.) Choose which information you wish to be shown using the 'Output Info' section." & vbNewLine & vbNewLine & _
"3.) Click 'Calculate' and see the results printed in both text and graphical format."
End Sub

Private Sub Command10_Click()
Text1.Text = "Distance to Sun:" & vbNewLine & vbNewLine & "1.) For use only when requesting the Coordinates option." & vbNewLine & vbNewLine & _
"2.) This set the distance to the sun to a number you chose for placement within Lightwave." & vbNewLine & vbNewLine & _
"3.) Note that Sqr(X^2+Y^2+Z^2) = the distance you choose." & vbNewLine & vbNewLine & _
"4.) For best results in Lightwave make sure the distance is much larger than the animation dimensions."
End Sub

Private Sub Command11_Click()
fraHelp.Visible = False
End Sub

Private Sub Command2_Click()
Text1.Text = "Entire Day Calculation:" & vbNewLine & vbNewLine & "1.) Set Location information using the green buttons along the top." & vbNewLine & vbNewLine & _
"2.) Choose which information you wish to be shown using the 'Output Info' section." & vbNewLine & vbNewLine & _
"3.) Click 'Calculate' and see the results printed in both text and graphical format." & vbNewLine & vbNewLine & _
"4.) The results can be exported to a Lightwave Motion File using the 'Export Motion' button."

End Sub

Private Sub Command3_Click()
Text1.Text = "Export Motion:" & vbNewLine & vbNewLine & _
"1.) This function is only available when you choose a Day Calculation." & vbNewLine & vbNewLine & _
"2.) The coordinate output of the calculation is saved to a .mot file" & vbNewLine & vbNewLine & _
"3.) This file can be used by Lightwave to avoid a lot of manual position entry."
End Sub

Private Sub Command4_Click()
Text1.Text = "Output Information:" & vbNewLine & vbNewLine & "1.) Set up your calculation type, Day or Time." & vbNewLine & _
"2.) Choose which information you wish to be shown using the 'Output Info' section." & vbNewLine & _
"     i)   Day Number:  The day of the year (1-365)" & vbNewLine & _
"     ii)  Solar Time:  Converts the time to Solar Time" & vbNewLine & _
"     iii) Coordinates: Returns x,y,z for the sun based on the Distance you provide. Use this for positioning an object using Lightwave." & vbNewLine & _
"     iv)  Distance:    Shows the distance to the sun you chose for the x,y,z position." & vbNewLine & _
"     v)   Declination: Some astronomy number you may like to know." & vbNewLine & _
"     vi)  Altitude:    The angle in degrees of the sun above the horizon.  A negative result indicates the sun is below the horizon." & vbNewLine & _
"     vii) Azimuth:     The angle in degrees of the sun away from due South.  A negative result indicates the sun is in the West" & vbNewLine & _
"3.) Click on 'Calculate' to see only the information you selected. "
End Sub

Private Sub Command5_Click()
Text1.Text = "Calculation:" & vbNewLine & vbNewLine & _
"1.) Set up all information using the green buttons along the top." & vbNewLine & vbNewLine & _
"2.) Select the calculation type." & vbNewLine & vbNewLine & _
"3.) Select output information." & vbNewLine & vbNewLine & _
"4.) Click 'Calculate' to start the calculations."
End Sub

Private Sub Command6_Click()
Text1.Text = "Latitude:" & vbNewLine & vbNewLine & _
"1.) Enter in the latitude of your location."
End Sub

Private Sub Command7_Click()
Text1.Text = "Longitude:" & vbNewLine & vbNewLine & _
"1.) Enter in the longitude of your location."
End Sub

Private Sub Command8_Click()
Text1.Text = "Date Information:" & vbNewLine & vbNewLine & _
"1.) Find the date you wish to base your calculations on using the calendar."

End Sub

Private Sub Command9_Click()
Text1.Text = "Time Information:" & vbNewLine & vbNewLine & _
"1.) Set your time zone." & vbNewLine & vbNewLine & _
"2.) Choose whether Daylight Savings Time is currently observed at the date you have previously selected." & vbNewLine & vbNewLine & _
"3.) If a specific Time Calculation has been selected then you may enter in the time." & vbNewLine & vbNewLine & _
"4.) Time must be entered in 24 hour format with midnight being 00:00"

End Sub
'**************************************************************
'everything in the section above is for the help information
'**************************************************************

Private Sub Form_Load()
Frm = -1
PictureTile Form2, picHold
Rad = 180 / 3.14159265359 ' used to convert to radians
'the Days array hold the number of days in the year at the start of each month 0 - 11
Days(0) = 0
Days(1) = 31
Days(2) = 31 + 28
Days(3) = 31 + 28 + 31
Days(4) = 31 + 28 + 31 + 30
Days(5) = 31 + 28 + 31 + 30 + 31
Days(6) = 31 + 28 + 31 + 30 + 31 + 30
Days(7) = 31 + 28 + 31 + 30 + 31 + 30 + 31
Days(8) = 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31
Days(9) = 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30
Days(10) = 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31
Days(11) = 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30
txtHour.Text = Hour(Time)
txtMin.Text = Minute(Time)
cboGMT.AddItem "G.M.T."
For i = 1 To 13
    cboGMT.AddItem "G.M.T. +" & i & ":00"
Next
For i = 1 To 12
    cboGMT.AddItem "G.M.T. -" & i & ":00"
Next
Month(0) = "January"
Month(1) = "February"
Month(2) = "March"
Month(3) = "April"
Month(4) = "May"
Month(5) = "June"
Month(6) = "July"
Month(7) = "August"
Month(8) = "September"
Month(9) = "October"
Month(10) = "November"
Month(11) = "December"
DateCheck
End Sub

Private Sub Form_Resize()
'set limits on form resizing, note that form can be expanded beyond original size to see more.
'actually since putting this in I made it so it cant be resized.

If Form2.Height < 7995 Then Form2.Height = 7995
If Form2.Height > 8025 Then Form2.Height = 8025
If Form2.Width < 6015 Then Form2.Width = 6015
If Form2.Width > 9630 Then Form2.Width = 9630
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
DateCheck
End Sub

Private Sub DateCheck()
'set up the day number and change the label showing the date
Dim MonthTemp
Dim DateParts
MonthTemp = MonthView1.Value
DateParts = Split(MonthTemp, "/")
j = UBound(DateParts)
If j = 0 Then
    MsgBox "You need to set the date format in Control Panel to:" & vbNewLine & _
    "dd/MM/yyyy for Windows 2000." & vbNewLine & _
    "Or" & vbNewLine & _
    "MM/dd/yyyy for Windows 98.", vbCritical, "Invalid Date Format."
    End
End If
m = Val(DateParts(1)) - 1
If m < 12 Then
    N = Days(DateParts(1) - 1) + DateParts(0)
    lblDate.Caption = Month(m) & " " & DateParts(0)
End If
If m = "" Then
    MsgBox "You need to set the date format in Control Panel to:" & vbNewLine & _
    "dd/MM/yyyy for Windows 2000." & vbNewLine & _
    "Or" & vbNewLine & _
    "MM/dd/yyyy for Windows 98.", vbCritical, "Invalid Date Format."
    End
End If
If m > 11 Then
    m = Val(DateParts(0)) - 1
    N = Days(DateParts(0) - 1) + DateParts(1)
    lblDate.Caption = Month(m) & " " & DateParts(1)
End If


End Sub

Private Sub opnCalc_Click(Index As Integer)
'choose between calculation types
cmdExport.Visible = False
If opnCalc(0).Value = True Then
Label6.Enabled = False
Label7.Enabled = False
txtHour.Enabled = False
txtMin.Enabled = False
cmdCalcAllDay.Visible = True
cmdCalcIndTime.Visible = False
txtElev.Enabled = False
Label5.Enabled = False
End If
If opnCalc(1).Value = True Then
Label6.Enabled = True
Label7.Enabled = True
txtHour.Enabled = True
txtMin.Enabled = True
cmdCalcAllDay.Visible = False
cmdCalcIndTime.Visible = True
txtElev.Enabled = True
Label5.Enabled = True
End If
End Sub

Sub PictureTile(pic1 As Form, Pic As PictureBox)
'This tiles the image across the form
    Dim i As Integer
    Dim t As Integer
    pic1.AutoRedraw = True
    Pic.BorderStyle = 0
    For t = 0 To pic1.Height Step Pic.ScaleHeight
        For i = 0 To pic1.Width Step Pic.ScaleWidth
            pic1.PaintPicture Pic.Picture, i, t
        Next i
    Next t
End Sub

Private Sub MotionFile()
'this routine sets up the Lightwave Motion File
Frm = Frm + 1
If Frm = 0 Then
    txtMF.Text = ""
    txtMF.Text = "LWMO"
    txtMF.Text = txtMF.Text & vbNewLine & "1"
    txtMF.Text = txtMF.Text & vbNewLine & "9"
    txtMF.Text = txtMF.Text & vbNewLine & "24"
End If
txtMF.Text = txtMF.Text & vbNewLine & Format(Xco, "######.00") & " " & Format(Yco, "######.00") & " " & Format(Zco, "######.00") & " 45 35 0 1 1 1"
txtMF.Text = txtMF.Text & vbNewLine & Frm & " 0 0 0 0"
txtMF.SelStart = Len(txtMF.Text)
txtMF.SelLength = 0
If Frm = 23 Then cmdExport.Visible = True
End Sub

Private Sub picAZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'When you move the mouse over the graph the labels to the right are updated with the current position
lblyPos.Caption = Format((-1 * Y + picAZ.ScaleHeight / 2), "######.00")
ta1 = Int((X / 2 / 60))
ta2 = (X / 2) Mod 60
If Int(X / 2) = 719 Then
    ta1 = 12
    ta2 = 0
End If
lblTimePos.Caption = ta1 & ":" & Format(ta2, "00")

End Sub
