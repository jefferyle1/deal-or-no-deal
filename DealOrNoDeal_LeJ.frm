VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000014&
   Caption         =   "Deal Or No Deal By Jeffery Le"
   ClientHeight    =   13635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   25305
   BeginProperty Font 
      Name            =   "Bahnschrift"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   13635
   ScaleWidth      =   25305
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   200
      Left            =   1440
      Top             =   9960
   End
   Begin VB.Frame fraGame 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8850
      Left            =   12720
      TabIndex        =   65
      Top             =   -1200
      Visible         =   0   'False
      Width           =   12615
      Begin VB.CommandButton cmdNoDeal 
         Caption         =   "NO DEAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         TabIndex        =   67
         Top             =   6840
         Width           =   2415
      End
      Begin VB.CommandButton cmdDeal 
         Caption         =   "DEAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   66
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label lblNumberToOpen 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   71
         Top             =   5160
         Width           =   12375
      End
      Begin VB.Label lblOffer 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   70
         Top             =   4320
         Width           =   12375
      End
      Begin VB.Label lblMenuCasesRemaining 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   69
         Top             =   3480
         Width           =   12375
      End
      Begin VB.Label lblMenuCases 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   68
         Top             =   2640
         Width           =   12375
      End
   End
   Begin VB.Frame fraEndGame 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8850
      Left            =   12720
      TabIndex        =   72
      Top             =   6720
      Width           =   12615
      Begin VB.CommandButton cmdNo 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8160
         TabIndex        =   76
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "YES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   75
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label lblNewHighScore 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Bahnschrift"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   81
         Top             =   4320
         Width           =   12375
      End
      Begin VB.Label lblLastOffer 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   3720
         Visible         =   0   'False
         Width           =   12375
      End
      Begin VB.Label lblReservedCase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   77
         Top             =   3000
         Width           =   12375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Would you like to play again?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   74
         Top             =   5040
         Width           =   12375
      End
      Begin VB.Label lblCongrats 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Top             =   2280
         Width           =   12375
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdSelected 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         TabIndex        =   64
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2520
         TabIndex        =   61
         Top             =   7200
         Width           =   7455
         Begin VB.Label lblBriefCase 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Bahnschrift"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   7215
         End
         Begin VB.Label lblRound 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Bahnschrift"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   7215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8640
         TabIndex        =   58
         Top             =   0
         Width           =   3735
         Begin VB.Label lblRemaining 
            Caption         =   "Remaining:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label lblOpened 
            Caption         =   "Opened: "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame fraCases 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   2520
         TabIndex        =   29
         Top             =   1320
         Width           =   7455
         Begin VB.CommandButton cmdNumber 
            Caption         =   "26"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   25
            Left            =   6360
            TabIndex        =   55
            Top             =   4560
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   24
            Left            =   120
            TabIndex        =   54
            Top             =   4560
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   23
            Left            =   6360
            TabIndex        =   53
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   22
            Left            =   5040
            TabIndex        =   52
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   21
            Left            =   3720
            TabIndex        =   51
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   20
            Left            =   2520
            TabIndex        =   50
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   19
            Left            =   1320
            TabIndex        =   49
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   18
            Left            =   120
            TabIndex        =   48
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   17
            Left            =   6360
            TabIndex        =   47
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   16
            Left            =   5040
            TabIndex        =   46
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   15
            Left            =   3720
            TabIndex        =   45
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   14
            Left            =   2520
            TabIndex        =   44
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   13
            Left            =   1320
            TabIndex        =   43
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   12
            Left            =   120
            TabIndex        =   42
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   11
            Left            =   6360
            TabIndex        =   41
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   10
            Left            =   5040
            TabIndex        =   40
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   9
            Left            =   3720
            TabIndex        =   39
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   8
            Left            =   2520
            TabIndex        =   38
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   7
            Left            =   1320
            TabIndex        =   37
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   6
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   5
            Left            =   6360
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   4
            Left            =   5040
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   3
            Left            =   3720
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   2520
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNumber 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   10200
         TabIndex        =   15
         Top             =   1320
         Width           =   2175
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   23
            Top             =   2640
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   21
            Top             =   3600
            Width           =   1905
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   20
            Top             =   4080
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   19
            Top             =   4560
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   120
            TabIndex        =   18
            Top             =   5040
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   120
            TabIndex        =   17
            Top             =   5520
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   25
            Left            =   120
            TabIndex        =   16
            Top             =   6000
            Width           =   1900
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   2175
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   2160
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Top             =   2640
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   3120
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   7
            Top             =   3600
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   6
            Top             =   4080
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   5
            Top             =   4560
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   4
            Top             =   5040
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   3
            Top             =   5520
            Width           =   1900
         End
         Begin VB.Label lblMoney 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   2
            Top             =   6000
            Width           =   1900
         End
      End
      Begin VB.Label lblHighScore 
         Alignment       =   2  'Center
         Caption         =   "This Session's Highscore: $0.00 "
         BeginProperty Font 
            Name            =   "Bahnschrift"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   80
         Top             =   8400
         Width           =   12375
      End
      Begin VB.Label Label2 
         Caption         =   "Your Briefcase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Deal Or No Deal                  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   63
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2400
         TabIndex        =   56
         Top             =   120
         Width           =   6135
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Scroll to the right to see the round and end game interfaces -------->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      TabIndex        =   78
      Top             =   9240
      Visible         =   0   'False
      Width           =   8055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Programmer: Jeffery Le

    Dim Money(0 To 25) As Single
    Dim RoundNumber As Integer
    Dim NumberOfCases As Integer
    Dim SelectedCaseVal As Integer
    Dim Bank As Single
    Dim HighScore As Single
    Dim Offer As Single
    
Private Sub Form_Load()

    Dim X As Integer
    Dim Y As Integer
        
    'Code that moves frames into position and reduces size of the form
    'Needed so that all frames can be seen during design time while also making the form proper size during run time
    frmMain.Width = 12915
    frmMain.Height = 9400
    fraGame.Move 50, 0
    fraEndGame.Move 50, 0
    
       
    'Code that is needed if the user decides to restart the game
    Bank = 0
    fraEndGame.Visible = False
    cmdSelected.Visible = False
    For Y = 0 To 25
        lblMoney(Y).BackColor = &H80FF80
        cmdNumber(Y).Visible = True
        cmdNumber(Y).Enabled = True
    Next Y
    frmMain.fraCases.Enabled = True
    frmMain.lblReservedCase.Visible = True
    frmMain.lblLastOffer.Visible = False
    frmMain.lblNewHighScore.Visible = False
    
    'Initializing the money array
    For X = 0 To 25
        Money(X) = 0
    Next X
    
    'Code to setup information that will be seen by user when program loads
    MoneyReader Money(), Bank
    RoundNumber = 0
    NumberOfCases = 26
    lblRound.Caption = "The Beginning"
    lblBriefCase.Caption = "Choose your reserved case"
    frmMain.lblOpened.Caption = "Opened: " & 0
    frmMain.lblRemaining.Caption = "Remaining: " & 26
    
End Sub
    

Private Sub cmdYes_Click()
    'Code that is executed if user wants to play again
    Form_Load
End Sub

Private Sub cmdNo_Click()
    'Code that is executed if user doesn't want to play again
    End
End Sub

Private Sub cmdNoDeal_Click()
    
    'Code that is executed when user chooses no deal
    frmMain.fraGame.Visible = False
    frmMain.fraCases.Enabled = True
End Sub

Private Sub cmdDeal_Click()
    
    'Code that is executed when user chooses deal
    frmMain.lblCongrats.Caption = "Congratulations, you've won " & Format$(Offer, "Currency") & "!"
    
    'If the player wins more money than the session's highscore, the highscore is replaced
    If Offer > HighScore Then
        HighScore = Offer
        lblNewHighScore.Visible = True
        lblNewHighScore.Caption = "You've acquired a new highscore of " & Format$(HighScore, "Currency") & " for this session!"
        lblHighScore.Caption = "This Session's Highscore: " & Format$(HighScore, "Currency")
    End If
    
    frmMain.fraEndGame.Visible = True
    frmMain.fraGame.Visible = False
End Sub

Private Sub cmdNumber_Click(Index As Integer)

    Dim X As Integer
    Dim Y As Integer
    
    'If statement needed so number of cases opened is accurate
    If RoundNumber <> 0 Then
        NumberOfCases = NumberOfCases - 1
    End If
    
    'Removes case from board when it is selected
    cmdNumber(Index).Visible = False
    
    'For loop statement that will unhighlight a dollar value on the side if the case that has that value is selected
    For X = 0 To 25
        If lblMoney(X).Caption = Format$(Money(Index), "Currency") And RoundNumber <> 0 Then
            lblMoney(X).BackColor = &H8000000F
        End If
    Next X
    
    'Code for selecting a case in the first round and putting it in the top left
    If RoundNumber = 0 Then
        cmdSelected.Visible = True
        cmdNumber(Index).Visible = False
        cmdSelected.Caption = cmdNumber(Index).Caption
        SelectedCaseVal = Index
        RoundNumber = 1
        'Reads the reserved case money so that it can be used in calculation
        Bank = Bank + Money(Index)
        frmMain.lblReservedCase.Caption = "Your reserved case contained " & Format$(Money(Index), "Currency") & "."
    End If
    
    'Code for rounds 1 to 10
    Select Case RoundNumber
        Case 1 To 10
            Bank = Bank - Money(Index)
            RoundUpdater NumberOfCases, RoundNumber, Bank, Offer
            Status NumberOfCases, RoundNumber
        End Select
    
    'Disables all of the command button except for the reserved case which is moved back onto the board
    If RoundNumber = 10 Then
        For Y = 0 To 25
            cmdNumber(Y).Enabled = False
        Next Y
        cmdNumber(SelectedCaseVal).Enabled = True
        cmdNumber(SelectedCaseVal).Visible = True
        cmdSelected.Visible = False
    End If
    
    'Not actually a round but instead states how much money player receives
    If RoundNumber = 11 Then
        frmMain.lblLastOffer.Visible = True
        frmMain.lblReservedCase.Visible = False
        frmMain.lblCongrats.Caption = "You have selected your reserved case and have won " & Format$(Money(Index), "Currency") & " from it!"
        frmMain.fraEndGame.Visible = True
        
        'If the player wins more money than the session's highscore, the highscore is replaced
        If Money(Index) > HighScore Then
            HighScore = Money(Index)
            lblNewHighScore.Visible = True
            lblNewHighScore.Caption = "You've acquired a new highscore of " & Format$(HighScore, "Currency") & " for this session!"
            lblHighScore.Caption = "This Session's Highscore: " & Format$(HighScore, "Currency")
        End If
        
    End If
    
End Sub

Private Sub tmrTime_Timer()

    Dim Title As String
    
    'Code that allows for the scrolling title
    Title = frmMain.lblTitle.Caption
    Title = Mid$(Title, 2, Len(Title)) + Left(Title, 1)
    frmMain.lblTitle.Caption = Title
    
End Sub
