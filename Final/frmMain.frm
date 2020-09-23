VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form Form1 
   Caption         =   "HTML styles"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInvisible 
      Caption         =   "Invisibles"
      Height          =   2175
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtDMFont 
         Height          =   375
         Left            =   1440
         TabIndex        =   181
         Text            =   "12px Arial"
         Top             =   480
         Width           =   495
      End
      Begin MSComDlg.CommonDialog dlgDMFont 
         Left            =   1680
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer tmrtempFile7 
         Interval        =   1
         Left            =   2280
         Top             =   1440
      End
      Begin MSComDlg.CommonDialog dlgPopupMenu 
         Left            =   1560
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer tmrMsgBox2 
         Interval        =   100
         Left            =   3120
         Top             =   960
      End
      Begin VB.PictureBox picMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   0
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   1110
         ScaleWidth      =   405
         TabIndex        =   17
         Top             =   360
         Width           =   435
      End
      Begin VB.Timer tmrTempFile6 
         Interval        =   1
         Left            =   2640
         Top             =   960
      End
      Begin VB.TextBox txtAMainFont 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Text            =   "12px  Arial"
         Top             =   480
         Width           =   495
      End
      Begin VB.Timer tmrTempFile3 
         Interval        =   1
         Left            =   2160
         Top             =   960
      End
      Begin VB.TextBox txtHoverFont 
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Text            =   "12px  Arial"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtScrollColor8 
         Height          =   375
         Left            =   720
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "FFFFFF"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor7 
         Height          =   375
         Left            =   720
         MaxLength       =   6
         TabIndex        =   13
         Text            =   "0E5BB0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor6 
         Height          =   375
         Left            =   720
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "0E5BB0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor2 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "0E5BB0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor3 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "008FF0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor4 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "008FF0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor5 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "008FF0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor1 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "008FF0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Timer tmrTempFile2 
         Interval        =   1
         Left            =   1680
         Top             =   960
      End
      Begin VB.Timer tmrMain 
         Interval        =   1
         Left            =   240
         Top             =   960
      End
      Begin MSComDlg.CommonDialog dlgButton1 
         Left            =   480
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog dlgButton2 
         Left            =   360
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtCursor 
         Height          =   405
         Left            =   240
         TabIndex        =   6
         Text            =   "hand"
         Top             =   480
         Width           =   495
      End
      Begin VB.Timer tmrTempFile1 
         Interval        =   1
         Left            =   1200
         Top             =   960
      End
      Begin VB.Timer tmrButton 
         Interval        =   1
         Left            =   720
         Top             =   960
      End
      Begin VB.TextBox txtButtonColor1 
         Height          =   405
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0E5BB0"
         Top             =   480
         Width           =   735
      End
      Begin MSComDlg.CommonDialog dlgButton3 
         Left            =   240
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox txtGenCode2 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":3081
      End
      Begin VB.TextBox txtButtonColor3 
         Height          =   435
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "008FF0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtButtonColor2 
         Height          =   405
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "FFFFFF"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFont 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Text            =   "10px  Arial"
         Top             =   480
         Width           =   495
      End
      Begin MSComDlg.CommonDialog dlgColorAll 
         Left            =   960
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   6015
      Left            =   5160
      ScaleHeight     =   5955
      ScaleWidth      =   4515
      TabIndex        =   32
      Top             =   240
      Width           =   4575
      Begin VB.Frame fraButton 
         BorderStyle     =   0  'None
         Caption         =   "Button Style"
         Height          =   6015
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   4335
         Begin VB.Frame fraImage 
            Caption         =   "Image"
            ForeColor       =   &H80000002&
            Height          =   4335
            Left            =   360
            TabIndex        =   34
            Top             =   480
            Width           =   3615
            Begin VB.TextBox txtImage2 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   40
               Text            =   "ImageFolder/1.gif"
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox txtImage1 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   39
               Text            =   "ImageFolder/2.gif"
               Top             =   600
               Width           =   2655
            End
            Begin VB.TextBox txtImageWidth 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   38
               Text            =   "150"
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox txtImageHeight 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1920
               TabIndex        =   37
               Text            =   "150"
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox txtUrl2 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   36
               Text            =   "index.htm"
               Top             =   2760
               Width           =   2655
            End
            Begin VB.TextBox txtImageCur 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   35
               Text            =   "hand"
               Top             =   3720
               Width           =   2655
            End
            Begin VB.Label Label30 
               Caption         =   "Main image (same folder): "
               Height          =   255
               Left            =   480
               TabIndex        =   46
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label Label31 
               Caption         =   "Mouseover image(same folder):"
               Height          =   255
               Left            =   480
               TabIndex        =   45
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label32 
               Caption         =   "Image width:"
               Height          =   255
               Left            =   480
               TabIndex        =   44
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label33 
               Caption         =   "Image height:"
               Height          =   255
               Left            =   1920
               TabIndex        =   43
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label34 
               Caption         =   "URL:"
               Height          =   255
               Left            =   480
               TabIndex        =   42
               Top             =   2520
               Width           =   375
            End
            Begin VB.Label Label35 
               Caption         =   "Cursor: Example: auto | crosshair | default | move | text | wait | help ,etc.."
               Height          =   375
               Left            =   480
               TabIndex        =   41
               Top             =   3240
               Width           =   2775
            End
         End
         Begin VB.OptionButton optImage 
            Caption         =   "Image"
            Height          =   375
            Left            =   1920
            TabIndex        =   67
            Top             =   5280
            Width           =   735
         End
         Begin VB.TextBox txtUrl 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2280
            TabIndex        =   66
            Text            =   "index.htm"
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CommandButton cmdButtonCur 
            Caption         =   "Edit"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   2640
            Width           =   615
         End
         Begin VB.CommandButton cmdButtonBorder 
            Caption         =   "Edit"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox txtButLeft 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1560
            TabIndex        =   63
            Text            =   "10"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtButTop 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            TabIndex        =   62
            Text            =   "10"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtButWidth 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2280
            TabIndex        =   61
            Text            =   "120"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtButHeight 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3000
            TabIndex        =   60
            Text            =   "120"
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton cmdButtonBGMO 
            Caption         =   "Select"
            Height          =   375
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton cmdButtobBG 
            Caption         =   "Selcet"
            Height          =   375
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton cmdTextColor 
            Caption         =   "Select"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Select"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1800
            Width           =   615
         End
         Begin VB.OptionButton optAlpha 
            Caption         =   "Alpha"
            Height          =   375
            Left            =   840
            TabIndex        =   55
            Top             =   4920
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optWave 
            Caption         =   "Wave"
            Height          =   375
            Left            =   1920
            TabIndex        =   54
            Top             =   4920
            Width           =   855
         End
         Begin VB.OptionButton optBlur 
            Caption         =   "Blur"
            Height          =   375
            Left            =   3000
            TabIndex        =   53
            Top             =   4920
            Width           =   855
         End
         Begin VB.OptionButton optNormal 
            Caption         =   "Normal"
            Height          =   375
            Left            =   840
            TabIndex        =   52
            Top             =   5280
            Width           =   855
         End
         Begin VB.Frame fraButtonBlur 
            Caption         =   "CSS filter : BLUR"
            ForeColor       =   &H80000002&
            Height          =   1575
            Left            =   1200
            TabIndex        =   47
            Top             =   3240
            Width           =   2055
            Begin VB.TextBox txtBlurstr 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1200
               TabIndex        =   49
               Text            =   "90"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtBlurDir 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   240
               TabIndex        =   48
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label16 
               Caption         =   "Direction (Angle)"
               Height          =   495
               Left            =   240
               TabIndex        =   51
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label15 
               Caption         =   "Strenght 0-1000"
               Height          =   615
               Left            =   1200
               TabIndex        =   50
               Top             =   480
               Width           =   735
            End
         End
         Begin VB.Frame fraButtonWave 
            Caption         =   "CSS filter : WAVE"
            ForeColor       =   &H80000002&
            Height          =   1575
            Left            =   840
            TabIndex        =   77
            Top             =   3240
            Width           =   2775
            Begin VB.TextBox txtWaveStr 
               BackColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   240
               TabIndex        =   80
               Text            =   "1"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtWaveFreq 
               BackColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   1080
               TabIndex        =   79
               Text            =   "899"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtWavelStr 
               BackColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   1920
               TabIndex        =   78
               Text            =   "30"
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label14 
               Caption         =   "Light- strength: 0-100"
               Height          =   615
               Left            =   1920
               TabIndex        =   83
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Freq: 0-900"
               Height          =   375
               Left            =   1080
               TabIndex        =   82
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label12 
               Caption         =   "Strength: 0-100"
               Height          =   495
               Left            =   240
               TabIndex        =   81
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.Frame fraButtonAlpha 
            Caption         =   "CSS filter : ALPHA"
            ForeColor       =   &H80000002&
            Height          =   1575
            Left            =   840
            TabIndex        =   70
            Top             =   3240
            Width           =   2775
            Begin VB.TextBox txtAlphaStyle 
               BackColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   1920
               TabIndex        =   73
               Text            =   "2"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtAlphaFinOp 
               BackColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   1080
               TabIndex        =   72
               Text            =   "0"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtAlphaOp 
               BackColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   240
               TabIndex        =   71
               Text            =   "100"
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "Style: 0-3"
               Height          =   255
               Left            =   1800
               TabIndex        =   76
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "Finish Opacity: 0-100"
               Height          =   615
               Left            =   1080
               TabIndex        =   75
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Opacity: 0-100"
               Height          =   375
               Left            =   240
               TabIndex        =   74
               Top             =   480
               Width           =   735
            End
         End
         Begin VB.Frame fraButtonNormal 
            Caption         =   "Normal"
            ForeColor       =   &H80000002&
            Height          =   1575
            Left            =   840
            TabIndex        =   68
            Top             =   3240
            Width           =   2775
            Begin VB.Label Label49 
               Caption         =   "NORMAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   495
               Left            =   720
               TabIndex        =   69
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.Label Label56 
            Caption         =   "[ Button Style ]"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   1560
            TabIndex        =   152
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "URL:"
            Height          =   255
            Left            =   2280
            TabIndex        =   94
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "BG Color"
            Height          =   375
            Left            =   2280
            TabIndex        =   93
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Mouseover BG Color"
            Height          =   375
            Left            =   3000
            TabIndex        =   92
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Text Color"
            Height          =   375
            Left            =   1560
            TabIndex        =   91
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Fonts"
            Height          =   255
            Left            =   840
            TabIndex        =   90
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Cursor"
            Height          =   255
            Left            =   1560
            TabIndex        =   89
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Border"
            Height          =   255
            Left            =   840
            TabIndex        =   88
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Left"
            Height          =   255
            Left            =   1680
            TabIndex        =   87
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Top"
            Height          =   255
            Left            =   840
            TabIndex        =   86
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Width"
            Height          =   255
            Left            =   2280
            TabIndex        =   85
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Height"
            Height          =   255
            Left            =   3000
            TabIndex        =   84
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame fraScroll 
         BorderStyle     =   0  'None
         Caption         =   "Scroll bar"
         Height          =   6015
         Left            =   0
         TabIndex        =   134
         Top             =   0
         Width           =   4455
         Begin VB.CommandButton cmdScroolFC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   142
            Top             =   3960
            Width           =   1095
         End
         Begin VB.CommandButton cmdScrollAC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdScrollBC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   140
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton cmdScrollDC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   139
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton cmdScrollTC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   138
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdScroll3DC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton cmdScrollSC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   136
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdScrollHC 
            Caption         =   "Select"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label57 
            Caption         =   "[ Scroll Bar ]"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   1680
            TabIndex        =   153
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label36 
            Caption         =   "Scrollbar-face-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   150
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Label Label28 
            Caption         =   "Scrollbar-arrow-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   149
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Label27 
            Caption         =   "Scrollbar-base-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   148
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label label26 
            Caption         =   "Scrollbar-darkshadow-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   147
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label label25 
            Caption         =   "Scrollbar-track-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   146
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label23 
            Caption         =   "Scrollbar-3Dlight-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   145
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Scrollbar-shadow-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   144
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Scrollbar-highlight-color:"
            Height          =   255
            Left            =   720
            TabIndex        =   143
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame fraHover 
         BorderStyle     =   0  'None
         Caption         =   "Hover"
         Height          =   6015
         Left            =   0
         TabIndex        =   108
         Top             =   0
         Width           =   4455
         Begin VB.TextBox txtHHeight 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2640
            TabIndex        =   121
            Text            =   "20"
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdHoverBorder 
            Caption         =   "Border"
            Height          =   375
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   4440
            Width           =   1215
         End
         Begin VB.TextBox txtHWidth 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            TabIndex        =   119
            Text            =   "173"
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdAMainFont 
            Caption         =   "Font"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdHoverFont 
            Caption         =   "Font"
            Height          =   375
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox txtVC 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2640
            TabIndex        =   116
            Text            =   "FFFFFF"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtAAC 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            TabIndex        =   115
            Text            =   "FFFFFF"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtLC 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2640
            TabIndex        =   114
            Text            =   "FFFFFF"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtAC 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2640
            TabIndex        =   113
            Text            =   "000000"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtABC 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            TabIndex        =   112
            Text            =   "008FF0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtHBC 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            TabIndex        =   111
            Text            =   "0E5BB0"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdAColor 
            Caption         =   "Generate  color"
            Height          =   375
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   5400
            Width           =   1455
         End
         Begin VB.TextBox txtAcolor 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            MaxLength       =   6
            TabIndex        =   109
            Text            =   "FF6600"
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Label Label58 
            Caption         =   "[ Hover ]"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   1800
            TabIndex        =   154
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label50 
            Caption         =   "Height"
            Height          =   255
            Left            =   2640
            TabIndex        =   133
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label45 
            Caption         =   "Border:"
            Height          =   255
            Left            =   1680
            TabIndex        =   132
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label43 
            Caption         =   "Width"
            Height          =   255
            Left            =   840
            TabIndex        =   131
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label41 
            Caption         =   "Edit main font:"
            Height          =   255
            Left            =   840
            TabIndex        =   130
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label48 
            Caption         =   "Edit hover font:"
            Height          =   255
            Left            =   2640
            TabIndex        =   129
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label47 
            Caption         =   "Generate HTML color:"
            Height          =   255
            Left            =   1560
            TabIndex        =   128
            Top             =   5040
            Width           =   1695
         End
         Begin VB.Label Label46 
            Caption         =   "Visited text color:"
            Height          =   255
            Left            =   2640
            TabIndex        =   127
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "Active text color:"
            Height          =   255
            Left            =   840
            TabIndex        =   126
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "Link text color:"
            Height          =   255
            Left            =   2640
            TabIndex        =   125
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label40 
            Caption         =   "Main BG color:"
            Height          =   255
            Left            =   840
            TabIndex        =   124
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label39 
            Caption         =   "Main text color:"
            Height          =   255
            Left            =   2640
            TabIndex        =   123
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label37 
            Caption         =   "Hover BG color:"
            Height          =   375
            Left            =   840
            TabIndex        =   122
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame fraDropDown 
         BorderStyle     =   0  'None
         Caption         =   "Drop Down"
         Height          =   6015
         Left            =   0
         TabIndex        =   168
         Top             =   0
         Width           =   4575
         Begin VB.CommandButton cmdDMcolorgen 
            Caption         =   "Generate color"
            Height          =   375
            Left            =   2400
            TabIndex        =   202
            Top             =   5400
            Width           =   1335
         End
         Begin VB.TextBox txtDMColorGen 
            Height          =   375
            Left            =   840
            MaxLength       =   6
            TabIndex        =   201
            Text            =   "FFFFFF"
            Top             =   5400
            Width           =   1455
         End
         Begin VB.CommandButton cmdDMDel 
            Caption         =   "Remove"
            Height          =   375
            Left            =   3360
            TabIndex        =   190
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtDMDel 
            Height          =   375
            Left            =   1320
            TabIndex        =   189
            Text            =   "Click the list box"
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CommandButton cmdDMAdd 
            Caption         =   "Add"
            Height          =   855
            Left            =   3360
            TabIndex        =   188
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtDMName 
            Height          =   375
            Left            =   1320
            TabIndex        =   187
            Text            =   "Index"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox txtDMURL 
            Height          =   375
            Left            =   1320
            TabIndex        =   186
            Text            =   "index.htm"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.ListBox lstDropDown 
            Height          =   900
            IntegralHeight  =   0   'False
            ItemData        =   "frmMain.frx":312F
            Left            =   720
            List            =   "frmMain.frx":3131
            TabIndex        =   185
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtDMHeight 
            Height          =   375
            Left            =   3360
            TabIndex        =   184
            Text            =   "37"
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox txtDMWidth 
            Height          =   375
            Left            =   2400
            TabIndex        =   183
            Text            =   "180"
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox txtDmBG 
            Height          =   375
            Left            =   1560
            TabIndex        =   182
            Text            =   "ImageFolder/4.jpg"
            Top             =   4320
            Width           =   1695
         End
         Begin VB.CommandButton cmdDMFont 
            Caption         =   "Font"
            Height          =   375
            Left            =   480
            TabIndex        =   180
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox txtDMColor 
            Height          =   375
            Left            =   1440
            TabIndex        =   179
            Text            =   "FFFFFF"
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox txtDMBGcolor 
            Height          =   375
            Left            =   480
            TabIndex        =   178
            Text            =   "0953B0"
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label80 
            Caption         =   "Generate HTML color:"
            Height          =   255
            Left            =   1560
            TabIndex        =   203
            Top             =   5040
            Width           =   1695
         End
         Begin VB.Label Label79 
            Caption         =   "Background URL:"
            Height          =   255
            Left            =   1560
            TabIndex        =   200
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Label Label78 
            Caption         =   "Edit font:"
            Height          =   255
            Left            =   480
            TabIndex        =   199
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label77 
            Caption         =   "Height"
            Height          =   255
            Left            =   3360
            TabIndex        =   198
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label76 
            Caption         =   "Width:"
            Height          =   255
            Left            =   2400
            TabIndex        =   197
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label75 
            Caption         =   "Text color:"
            Height          =   255
            Left            =   1440
            TabIndex        =   196
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label74 
            Caption         =   "BG color:"
            Height          =   255
            Left            =   480
            TabIndex        =   195
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label73 
            Caption         =   "Remove:"
            Height          =   255
            Left            =   480
            TabIndex        =   194
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label72 
            Caption         =   "Link name:"
            Height          =   255
            Left            =   360
            TabIndex        =   193
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label71 
            Caption         =   "Link URL:"
            Height          =   255
            Left            =   480
            TabIndex        =   192
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label70 
            Caption         =   "Menu list:"
            Height          =   255
            Left            =   600
            TabIndex        =   191
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label69 
            Caption         =   "[ Drop Down Menu ]"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   1560
            TabIndex        =   177
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame fraPopupMenu 
         BorderStyle     =   0  'None
         Caption         =   "Menu Bar"
         Height          =   6015
         Left            =   0
         TabIndex        =   95
         Top             =   0
         Width           =   4575
         Begin VB.TextBox txtSMleft 
            Height          =   375
            Left            =   3120
            TabIndex        =   172
            Text            =   "28"
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox txtSMtop 
            Height          =   375
            Left            =   2280
            TabIndex        =   171
            Text            =   "28"
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox txtMMleft 
            Height          =   375
            Left            =   1440
            TabIndex        =   170
            Text            =   "10"
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox txtMMtop 
            Height          =   375
            Left            =   600
            TabIndex        =   169
            Text            =   "10"
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox txtSMHover 
            Height          =   375
            Left            =   3600
            TabIndex        =   166
            Text            =   "FFFFFF"
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox txtSMBGColor 
            Height          =   375
            Left            =   2760
            TabIndex        =   164
            Text            =   "F9F8C0"
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox txtMMBGColor 
            Height          =   375
            Left            =   1080
            TabIndex        =   161
            Text            =   "53B2F0"
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox txtSMtextcolor 
            Height          =   375
            Left            =   1920
            TabIndex        =   159
            Text            =   "000000"
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox txtMMtextColor 
            Height          =   375
            Left            =   240
            TabIndex        =   157
            Text            =   "000000"
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox txtMenuColorGen 
            Height          =   375
            Left            =   960
            MaxLength       =   6
            TabIndex        =   156
            Text            =   "FFFFFF"
            Top             =   5400
            Width           =   1455
         End
         Begin VB.CommandButton cmdMenuColorGen 
            Caption         =   "Generate color"
            Height          =   375
            Left            =   2520
            TabIndex        =   155
            Top             =   5400
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelSm 
            Caption         =   "Remove"
            Height          =   375
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   2760
            Width           =   735
         End
         Begin VB.CommandButton cmdAddSm 
            Caption         =   "Add "
            Height          =   855
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtDelSm 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1080
            TabIndex        =   100
            Text            =   "Click list in the list box."
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtSmName 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1080
            TabIndex        =   99
            Text            =   "Yahoo"
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtSmUrl 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1080
            TabIndex        =   98
            Text            =   "http://www.yahoo.com"
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox txtMMenu 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1080
            TabIndex        =   97
            Text            =   "Menu"
            Top             =   600
            Width           =   3255
         End
         Begin VB.ListBox lstPopupMenu 
            BackColor       =   &H00FFFFFF&
            Height          =   660
            IntegralHeight  =   0   'False
            ItemData        =   "frmMain.frx":3133
            Left            =   1080
            List            =   "frmMain.frx":313A
            TabIndex        =   96
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label68 
            Caption         =   "Sub left:"
            Height          =   255
            Left            =   3120
            TabIndex        =   176
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label67 
            Caption         =   "Sub top:"
            Height          =   255
            Left            =   2280
            TabIndex        =   175
            Top             =   4200
            Width           =   855
         End
         Begin VB.Label Label61 
            Caption         =   "Main left:"
            Height          =   255
            Left            =   1440
            TabIndex        =   174
            Top             =   4200
            Width           =   855
         End
         Begin VB.Label Label59 
            Caption         =   "Main top:"
            Height          =   255
            Left            =   600
            TabIndex        =   173
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label66 
            Caption         =   "Sub hover BG color:"
            Height          =   375
            Left            =   3600
            TabIndex        =   167
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Sub BG color:"
            Height          =   375
            Left            =   2760
            TabIndex        =   165
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label64 
            Caption         =   "Main BG color:"
            Height          =   480
            Left            =   1080
            TabIndex        =   163
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label63 
            Caption         =   "Generate HTML color:"
            Height          =   255
            Left            =   1680
            TabIndex        =   162
            Top             =   5040
            Width           =   1695
         End
         Begin VB.Label Label62 
            Caption         =   "Sub text color:"
            Height          =   495
            Left            =   1920
            TabIndex        =   160
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label60 
            Caption         =   "Main text color:"
            Height          =   495
            Left            =   240
            TabIndex        =   158
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "[ Menu  Bar ]"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   1800
            TabIndex        =   151
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Remove:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   107
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label54 
            Caption         =   "Link name:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label53 
            Alignment       =   2  'Center
            Caption         =   " Link url:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label52 
            Alignment       =   2  'Center
            Caption         =   "Main menu:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label51 
            Caption         =   "Sub menu:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   960
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   5160
      ScaleHeight     =   915
      ScaleWidth      =   4515
      TabIndex        =   26
      Top             =   6480
      Width           =   4575
      Begin VB.OptionButton optButton 
         Caption         =   "Button style"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optScroll 
         Caption         =   "Scroll Bar"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optHover 
         Caption         =   "Hover "
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton optPopupMenu 
         Caption         =   "Popup menu"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton optDropDown 
         Caption         =   "Drop Down Menu"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   4755
      TabIndex        =   18
      Top             =   6480
      Width           =   4815
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swfb2 
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Width           =   1275
         _cx             =   2249
         _cy             =   661
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ExactFit"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swfb3 
         Height          =   375
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   1275
         _cx             =   2249
         _cy             =   661
         FlashVars       =   ""
         Movie           =   " "
         Src             =   " "
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ExactFit"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swfb1 
         Height          =   375
         Left            =   240
         TabIndex        =   204
         Top             =   240
         Width           =   1275
         _cx             =   2249
         _cy             =   661
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ExactFit"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   4755
      TabIndex        =   21
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtGenCode1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2535
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Top             =   360
         Width           =   4335
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2535
         Left            =   240
         TabIndex        =   22
         Top             =   3240
         Width           =   4335
         ExtentX         =   7646
         ExtentY         =   4471
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Label Label21 
         Caption         =   "[ Generated Code ]"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "[ Preview ]"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Xc As Integer
Private Yc As Integer
Dim MoveForm As Boolean
Dim TempFile1 As String
Dim TempFile2 As String
Dim TempFile3 As String
Dim TempFile4 As String
Dim TempFile5 As String
Dim TempFile6 As String
Dim TempFile7 As String
Dim prevwalp As Boolean
Dim cur As String

'html color generator ;hover part
Private Sub cmdAColor_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtAcolor.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'add list to list box( popup menu)
Private Sub cmdAddSm_Click()
On Error Resume Next
    lstPopupMenu.AddItem "<tr><td><a href='" & txtSmUrl.Text & "'>" & txtSmName.Text & "</a><br></td></tr>"
End Sub

'<A></A> fonts
Private Sub cmdAMainFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
    dlgButton1.Flags = 3
    dlgButton1.FontName = "Arial"
    dlgButton1.ShowFont
    If dlgButton1.FontItalic = True Then
        itl = "italic"
    Else
        itl = ""
    End If
    If dlgButton1.FontBold = True Then
        bld = "bold"
    Else
    bld = ""
    End If
    txtAMainFont.Text = itl & " " & bld & " " & dlgButton1.FontSize & "px" & " " & dlgButton1.FontName
End Sub

'button border
Private Sub cmdButtonBorder_Click()
    Form2.Show
    Form1.Enabled = False
End Sub
'button cursor
Private Sub cmdButtonCur_Click()
    cur = InputBox("Cursor: ( auto | crosshair | default | move | e-resize | ne-resize | nw-resize | n-resize | se-resize | sw-resize | w-resize | text | wait | help )", "Cursor Editor", "hand")
    txtCursor.Text = cur
End Sub

'del list in list box(popup menu)
Private Sub cmdDelSm_Click()
On Error Resume Next
lstPopupMenu.RemoveItem lstPopupMenu.ListIndex
End Sub

'generate html color ;dropdown menu part
Private Sub cmdDMcolorgen_Click()
On Error Resume Next
    dlgDMFont.ShowColor
    txtDMColorGen.Text = Right(StrReverse(Hex(dlgDMFont.Color)), Len(Hex(dlgDMFont.Color)) - 1) & "000000"
End Sub

Private Sub cmdDMDel_Click()
On Error Resume Next
lstDropDown.RemoveItem lstDropDown.ListIndex
End Sub

Private Sub cmdDMFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
    dlgDMFont.Flags = 3
    dlgDMFont.FontName = "Arial"
    dlgDMFont.ShowFont
    If dlgDMFont.FontItalic = True Then
        itl = "italic"
    Else
        itl = ""
    End If
    If dlgDMFont.FontBold = True Then
        bld = "bold"
    Else
        bld = ""
    End If
    txtDMFont.Text = itl & " " & bld & " " & dlgDMFont.FontSize & "px" & " " & dlgDMFont.FontName
End Sub

Private Sub cmdHoverBorder_Click()
    Form2.Show
    Form1.Enabled = False
End Sub

'popup menu html color generator
Private Sub cmdMenuColorGen_Click()
On Error Resume Next
    dlgPopupMenu.ShowColor
    txtMenuColorGen.Text = Right(StrReverse(Hex(dlgPopupMenu.Color)), Len(Hex(dlgPopupMenu.Color)) - 1) & "000000"
End Sub

'scroll 3d color
Private Sub cmdScroll3DC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor1.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll arrow color
Private Sub cmdScrollAC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor7.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll base color
Private Sub cmdScrollBC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor6.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll Dark color
Private Sub cmdScrollDC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor5.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll highlight color
Private Sub cmdScrollHC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor1.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll shadow color
Private Sub cmdScrollSC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor2.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll track color
Private Sub cmdScrollTC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor4.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll face color
Private Sub cmdScroolFC_Click()
On Error Resume Next
    dlgColorAll.ShowColor
    txtScrollColor8.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'button text color
Private Sub cmdTextColor_Click()
On Error Resume Next
    dlgButton3.ShowColor
    txtButtonColor2.Text = Right(StrReverse(Hex(dlgButton3.Color)), Len(Hex(dlgButton3.Color)) - 1) & "000000"
End Sub

'button BG color
Private Sub cmdButtobBG_Click()
On Error Resume Next
    dlgButton2.ShowColor
    txtButtonColor3.Text = Right(StrReverse(Hex(dlgButton2.Color)), Len(Hex(dlgButton2.Color)) - 1) & "000000"
End Sub
'button Mouseover BG color
Private Sub cmdButtonBGMO_Click()
On Error Resume Next
    dlgButton1.ShowColor
    txtButtonColor1.Text = Right(StrReverse(Hex(dlgButton1.Color)), Len(Hex(dlgButton1.Color)) - 1) & "000000"
End Sub

'button fonts
Private Sub cmdFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
    dlgButton1.Flags = 3
    dlgButton1.FontName = "Arial"
    dlgButton1.ShowFont
    If dlgButton1.FontItalic = True Then
        itl = "italic"
    Else
        itl = ""
    End If
    If dlgButton1.FontBold = True Then
        bld = "bold"
    Else
        bld = ""
    End If
    txtFont.Text = itl & " " & bld & " " & dlgButton1.FontSize & "px" & " " & dlgButton1.FontName
End Sub



'hover fonts
Private Sub cmdHoverFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
    dlgButton1.Flags = 3
    dlgButton1.FontName = "Arial"
    dlgButton1.ShowFont
    If dlgButton1.FontItalic = True Then
        itl = "italic"
    Else
        itl = ""
    End If
    If dlgButton1.FontBold = True Then
        bld = "bold"
    Else
        bld = ""
    End If
    txtHoverFont.Text = itl & " " & bld & " " & dlgButton1.FontSize & "px" & " " & dlgButton1.FontName
End Sub

'add drop down list
Private Sub cmdDMAdd_Click()
On Error Resume Next
    lstDropDown.AddItem "<option value='" & txtDMURL.Text & "'>" & txtDMName.Text & "</option>"
End Sub

'drag form code
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveMe = True
    Xc = X
    Yc = Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MoveForm = True Then
        Form1.Left = Form1.Left + (X - Xc)
        Form1.Top = Form1.Top + (Y - Yc)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form1.Left = Form1.Left + (X - Xc)
    Form1.Top = Form1.Top + (Y - Yc)
    MoveForm = False

End Sub
'end of drag form code

Private Sub Form_Load()
On Error Resume Next
    
    prevwalp = True
    Form2.bordercolortxt.MaxLength = 6
    fraInvisible.Visible = False
    'load flash buttons
    swfb1.Movie = App.Path & "\ImageFolder\1.swf"
    swfb2.Movie = App.Path & "\ImageFolder\2.swf"
    swfb3.Movie = App.Path & "\ImageFolder\3.swf"
    'create temp file
    TempFile1 = App.Path & "\1.html"
    TempFile2 = App.Path & "\2.html"
    TempFile3 = App.Path & "\3.html"
    TempFile4 = App.Path & "\4.html"
    TempFile5 = App.Path & "\5.html"
    TempFile6 = App.Path & "\6.html"
    TempFile7 = App.Path & "\7.html"

End Sub

'temp file delete and close application
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Sample, how to use Msgbox2.
MsgBox2 Form1.hwnd, "        Thank you for using HTML Style II.     ", "Thank you...", vbOKOnly, 0, 0, 30, 75
    
    Unload Form2
    If TempFile1 <> "" Then Kill TempFile1
    If TempFile2 <> "" Then Kill TempFile2
    If TempFile3 <> "" Then Kill TempFile3
    If TempFile4 <> "" Then Kill TempFile4
    If TempFile5 <> "" Then Kill TempFile5
    If TempFile6 <> "" Then Kill TempFile6
    If TempFile7 <> "" Then Kill TempFile7
    End
End Sub

'drop down menu list
Private Sub lstDropDown_Click()
    txtDelSm.Text = lstDropDown.List(lstDropDown.ListIndex)
End Sub

'popup menu list
Private Sub lstPopupMenu_Click()
    txtDelSm.Text = lstPopupMenu.List(lstPopupMenu.ListIndex)
End Sub


'generate code
Private Sub swfb1_GotFocus()
If optButton.Value = True Then
    If optAlpha.Value = True Then
        txtGenCode2.FileName = TempFile1
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optWave.Value = True Then
        txtGenCode2.FileName = TempFile2
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optBlur.Value = True Then
        txtGenCode2.FileName = TempFile3
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optNormal.Value = True Then
        txtGenCode2.FileName = TempFile4
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optImage.Value = True Then
        txtGenCode2.FileName = TempFile5
        txtGenCode1.Text = txtGenCode2.Text
    End If
End If
If optScroll.Value = True Then
        txtGenCode2.FileName = TempFile2
        txtGenCode1.Text = txtGenCode2.Text
End If
If optHover.Value = True Then
        txtGenCode2.FileName = TempFile3
        txtGenCode1.Text = txtGenCode2.Text
End If
If optPopupMenu.Value = True Then
        txtGenCode2.FileName = TempFile6
        txtGenCode1.Text = txtGenCode2.Text
End If
If optDropDown.Value = True Then
        txtGenCode2.FileName = TempFile7
        txtGenCode1.Text = txtGenCode2.Text
End If


End Sub

'preview html
Private Sub swfb2_GotFocus()
If optButton.Value = True Then
    If optAlpha.Value = True Then
        WebBrowser1.Navigate TempFile1
    End If
    If optWave.Value = True Then
        WebBrowser1.Navigate TempFile2
    End If
    If optBlur.Value = True Then
        WebBrowser1.Navigate TempFile3
    End If
    If optNormal.Value = True Then
        WebBrowser1.Navigate TempFile4
    End If
    If optImage.Value = True Then
        WebBrowser1.Navigate TempFile5
    End If
End If
If optHover.Value = True Then
     WebBrowser1.Navigate TempFile3
End If
If optScroll.Value = True Then
     WebBrowser1.Navigate TempFile2
End If
If optPopupMenu.Value = True Then
     WebBrowser1.Navigate TempFile6
End If
If optDropDown.Value = True Then
     WebBrowser1.Navigate TempFile7
End If


End Sub


'preview in internet explorer
Private Sub swfb3_GotFocus()
If optButton.Value = True Then
    If optAlpha.Value = True Then
        Shell "explorer.exe " & TempFile1, 0
    End If
    If optBlur.Value = True Then
        Shell "explorer.exe " & TempFile3, 0
    End If
    If optWave.Value = True Then
        Shell "explorer.exe " & TempFile2, 0
    End If
    If optNormal.Value = True Then
        Shell "explorer.exe " & TempFile4, 0
    End If
    If optImage.Value = True Then
        Shell "explorer.exe " & TempFile5, 0
    End If
End If
If optHover.Value = True Then
Shell "explorer.exe " & TempFile3, 0
End If
If optScroll.Value = True Then
Shell "explorer.exe " & TempFile2, 0
End If
If optPopupMenu.Value = True Then
Shell "explorer.exe " & TempFile6, 0
End If
If optDropDown.Value = True Then
Shell "explorer.exe " & TempFile7, 0
End If
swfb2.SetFocus
End Sub




'main timer loop for button,scroll and hover
Private Sub tmrMain_Timer()
If optButton.Value = True Then
    fraButton.Visible = True
    tmrTempFile1.Enabled = True
Else
    fraButton.Visible = False
    tmrTempFile1.Enabled = False
End If
If optScroll.Value = True Then
    fraScroll.Visible = True
    tmrTempFile2.Enabled = True
Else
    fraScroll.Visible = False
    tmrTempFile2.Enabled = False
End If
If optHover.Value = True Then
    fraHover.Visible = True
    tmrTempFile3.Enabled = True
Else
    fraHover.Visible = False
    tmrTempFile3.Enabled = False
End If
If optPopupMenu.Value = True Then
    fraPopupMenu.Visible = True
    tmrTempFile6.Enabled = True
Else
    fraPopupMenu.Visible = False
    tmrTempFile6.Enabled = False
End If

If optDropDown.Value = True Then
    fraDropDown.Visible = True
    tmrtempFile7.Enabled = True
Else
    fraDropDown.Visible = False
    tmrtempFile7.Enabled = False
End If

End Sub


Private Sub tmrMsgBox2_Timer()
Dim cRet As Long
Dim BCap(0 To 9) As String
Dim setBCap(0 To 9) As String
Dim i As Integer
Dim j As Integer

'call the code in module that put picture in message box
    Call MsgBoxWPicture
    
    
'sample of how change caption of buttons in message box(The one with "Welcome..." Thank you).
BCap(0) = "OK"
BCap(1) = "NO"
BCap(2) = "&No"
BCap(3) = "&Yes"
BCap(4) = "CANCEL"
BCap(5) = "&Abort"
BCap(6) = "&Retry"
BCap(7) = "&Ignore"
BCap(8) = "&Apply"
setBCap(0) = "[ OK ]"
setBCap(1) = "[ NO ]"
setBCap(2) = "[ &No ]"
setBCap(3) = "[ &Yes ]"
setBCap(4) = "[ Cancel ]"
setBCap(5) = "[ &Abort ]"
setBCap(6) = "[ &Retry ]"
setBCap(7) = "[ &Ignore ]"
setBCap(8) = "[ &Apply ]"
               
cRet = FindWindow("#32770", "Thank you...")
'find buttons
For i = 0 To 8
    bwnd = FindWindowEx(cRet, ByVal 0&, vbNullString, BCap(i))
    'change  message box button text
                    
    If CBool(cRet) = True Then
        SetWindowText bwnd, setBCap(i)
    Else
        Exit Sub
    End If
Next i
   
End Sub

'button temp files timer loop
'write/create temp file
Private Sub tmrTempFile1_Timer()
On Error Resume Next

Open TempFile1 For Output As #1
    Print #1, "<html>"
    Print #1, "<body>"
    Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text & ";"
    Print #1, "filter:Alpha(opacity=" & txtAlphaOp.Text & ",finishopacity=" & txtAlphaFinOp.Text & ",style=" & txtAlphaStyle.Text & ");font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
    Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
    Print #1, "Button Style</button>"
    Print #1, "</body>"
    Print #1, "</html>"
Close #1
 
Open TempFile2 For Output As #1
    Print #1, "<html>"
    Print #1, "<body>"
    Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text & ";"
    Print #1, "filter:wave(strength=" & txtWaveStr.Text & ",freq=" & txtWaveFreq.Text & ",lightstrength=" & txtWavelStr.Text & ");font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
    Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
    Print #1, "Button Style</button>"
    Print #1, "</body>"
    Print #1, "</html>"
Close #1
 
 
Open TempFile3 For Output As #1
    Print #1, "<html>"
    Print #1, "<body>"
    Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text & ";"
    Print #1, "filter:blur(direction=" & txtBlurDir.Text & ",strength=" & txtBlurstr.Text & ");font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
    Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
    Print #1, "Button Style</button>"
    Print #1, "</body>"
    Print #1, "</html>"
Close #1
 
Open TempFile4 For Output As #1
    Print #1, "<html>"
    Print #1, "<body>"
    Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text
    Print #1, ";font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
    Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
    Print #1, "Button Style</button>"
    Print #1, "</body>"
    Print #1, "</html>"
Close #1
 
On Error Resume Next
Open TempFile5 For Output As #1
    Print #1, "<html>"
    Print #1, "<body>"
    Print #1, "<img  name=Im src='" & txtImage1.Text & "' width=" & txtImageWidth.Text & " height=" & txtImageHeight.Text
    Print #1, " style='cursor:" & txtImageCur.Text & "' onClick=window.open('" & txtUrl2.Text & "') "
    Print #1, " onmouseover=Im.src='" & txtImage2.Text & "' onmouseout=Im.src='" & txtImage1.Text & "' >"
    Print #1, "</body>"
    Print #1, "</html>"
 Close #1
   
  
If prevwalp = True Then
  WebBrowser1.Navigate TempFile1
  txtGenCode2.FileName = TempFile1
  txtGenCode1.Text = txtGenCode2.Text
       prevwalp = False
    Exit Sub
 End If
   
 End Sub

'button timer loop; option: alpha,wave,normal,image and blur.
Private Sub tmrButton_Timer()
If optAlpha.Value = True Then
    fraButtonAlpha.Visible = True
Else
    fraButtonAlpha.Visible = False
End If
If optWave.Value = True Then
    fraButtonWave.Visible = True
Else
    fraButtonWave.Visible = False
End If
If optBlur.Value = True Then
    fraButtonBlur.Visible = True
Else
    fraButtonBlur.Visible = False
End If
If optNormal.Value = True Then
    fraButtonNormal.Visible = True
Else
    fraButtonNormal.Visible = False
End If
If optImage.Value = True Then
    fraImage.Visible = True

Else
    fraImage.Visible = False
End If
End Sub



'scroll temp file timer loop
'write/create temp file
Private Sub tmrTempFile2_Timer()

Open TempFile2 For Output As #1
    Print #1, "<html>"
    Print #1, "<style>"
    Print #1, "body"
    Print #1, "{SCROLLBAR-HIGHLIGHT-COLOR:#" & txtScrollColor1.Text & ";"
    Print #1, "SCROLLBAR-SHADOW-COLOR:#" & txtScrollColor2.Text & ";"
    Print #1, "SCROLLBAR-3DLIGHT-COLOR#:"; txtScrollColor3.Text & ";"
    Print #1, "SCROLLBAR-TRACK-COLOR#:" & txtScrollColor4.Text & ";"
    Print #1, "SCROLLBAR-DARKSHADOW-COLOR:#" & txtScrollColor5.Text & ";"
    Print #1, "SCROLLBAR-BASE-COLOR:#" & txtScrollColor6.Text & ";"
    Print #1, "SCROLLBAR-ARROW-COLOR:#" & txtScrollColor7.Text & ";"
    Print #1, "SCROLLBAR-FACE-COLOR:#" & txtScrollColor8.Text & ";}</style>"
    Print #1, "<body>"
    Print #1, "Spaces below, so you can preview it well. "
    Print #1, "<br><br><br><br><br><br>"
    Print #1, "<br><br><br><br><br><br>"
    Print #1, "<br><br><br><br><br><br>"
    Print #1, "<br><br><br><br><br><br>"
    Print #1, "</body>"
    Print #1, "</html>"
Close #1

End Sub

'hover temp file timer loop
'write/create temp file
Private Sub tmrTempFile3_Timer()
Open TempFile3 For Output As #1
    Print #1, "<html><head>"
    Print #1, "<style type=text/css>"
    Print #1, "a {width:" & txtHWidth.Text & ";height:" & txtHHeight.Text & ";text-decoration:none;border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "; background:#" & txtABC.Text & ";color:#" & txtAC.Text & ";font:" & txtAMainFont.Text & "};"
    Print #1, "a:hover {background:#" & txtHBC.Text & ";font:" & txtHoverFont.Text & "};"
    Print #1, "a:link {color:" & txtLC.Text & "};"
    Print #1, "a:ative {color:" & txtAAC.Text & "};"
    Print #1, "a:visited {color:" & txtVC.Text & "};"
    Print #1, "</style></head>"
    Print #1, "<body>"
    Print #1, "<table><tr><td align=center><a href=c:\>Press Preview</a></td></tr></table>"
    Print #1, "</body>"
    Print #1, "</html>"
Close #1
 
 
End Sub

'popup menu timer loop
'write/create temp file
Private Sub tmrTempFile6_Timer()
On Error Resume Next
Dim a As Integer
Open TempFile6 For Output As #1
    Print #1, "<HTML><HEAD>"
    Print #1, "<style type=text/css>"
    Print #1, "a{text-decoration:none;width:130};"
    Print #1, "a:hover{background:#" & txtSMHover & "};"
    Print #1, "a:link{color:#" & txtSMtextcolor.Text & "};"
    Print #1, "a:visited{color:#" & txtSMtextcolor.Text & "};"
    Print #1, "a:activate{color:#" & txtSMtextcolor.Text & "};"
    Print #1, ".tables{"
    Print #1, "border:1 solid black;"
    Print #1, "font:12px Arial;"
    Print #1, "background:#" & txtSMBGColor.Text & ";"
    Print #1, "width:130}</style>"
    Print #1, "<script language=JavaScript>"
    Print #1, "function showall() {"
    Print #1, "document.all.menu2.style.visibility='Visible';}"
    Print #1, "function hidemenu() {"
    Print #1, "document.all.menu2.style.visibility='hidden';}</script></HEAD>"
    Print #1, "<body>"
    Print #1, "<div id='menu' align=center style='position:absolute;top:" & txtMMtop.Text & ";left:" & txtMMleft.Text & ";width:100;"
    Print #1, "background:#" & txtMMBGColor.Text & ";color:#" & txtMMtextColor.Text & ";border:1 solid black;font:16px Arial' "
    Print #1, "onMouseOver='showall()' onMouseOut='hidemenu()'>"
    Print #1, txtMMenu.Text & "</div>"
    Print #1, "<div id='menu2' align=justify style='position:absolute;top:" & txtSMtop.Text & ";left:" & txtSMleft.Text & ";"
    Print #1, "visibility:Hidden' onMouseOver='showall()' onMouseOut='hidemenu()'>"
    Print #1, "<table class='tables'>"
        For a = 1 To lstPopupMenu.ListCount
            Print #1, lstPopupMenu.List(a - 1)
        Next a
    Print #1, "</table></div>"
    Print #1, "</body>"
    Print #1, "</HTML>"
Close #1
End Sub


'dropdown menu timer loop
'write/create temp file
Private Sub tmrtempFile7_Timer()
On Error Resume Next
Dim a2 As Integer

Open TempFile7 For Output As #1

    Print #1, "<html>"
    Print #1, "<style type=text/css>"
    Print #1, ".dropmenu{"
    Print #1, "background:" & txtDMBGcolor.Text & ";"
    Print #1, "color:#" & txtDMColor & ";"
    Print #1, "font:" & txtDMFont.Text & ";}"
    Print #1, ".tds{"
    Print #1, "border:1 solid #000000;"
    Print #1, "background:url(" & txtDmBG.Text & ");"
    Print #1, "height:" & txtDMHeight.Text & ";"
    Print #1, "width:" & txtDMWidth.Text & ";}</style>"
    Print #1, "<body>"
    Print #1, "<table style='position:absolute;left:10;top:10'>"
    Print #1, "<tr><td align=center class='tds'>"
    Print #1, "<select class='dropmenu' name='URls'>"
        For a2 = 1 To lstDropDown.ListCount
            Print #1, lstDropDown.List(a2 - 1)
        Next a2
    Print #1, "</select>"
    Print #1, "<INPUT TYPE=SUBMIT VALUE='Go' onclick=self.location=(URls.options[URls.selectedIndex].value)"
    Print #1, "style='color:#" & txtDMColor.Text & ";background:" & txtDMBGcolor.Text & ";font:" & txtDMFont.Text & "'></td></tr>"
    Print #1, "</table>"
    Print #1, "<body></html>"
Close #1

End Sub

