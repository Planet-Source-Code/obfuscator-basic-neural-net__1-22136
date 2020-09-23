VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Neural Network"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BNNForm.frx":0000
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help && Info"
      Height          =   465
      Left            =   795
      TabIndex        =   56
      Top             =   2775
      Width           =   1005
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset Weights"
      Height          =   465
      Left            =   795
      TabIndex        =   55
      Top             =   2205
      Width           =   1005
   End
   Begin MSComctlLib.Slider sldHidden1 
      Height          =   270
      Index           =   1
      Left            =   5055
      TabIndex        =   49
      Top             =   3675
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      LargeChange     =   1
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Train 1000 times"
      Height          =   465
      Left            =   795
      TabIndex        =   48
      Top             =   1155
      Width           =   1005
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test"
      Height          =   525
      Left            =   4800
      TabIndex        =   47
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton btnXOR 
      Caption         =   "XOR"
      Height          =   345
      Left            =   2625
      TabIndex        =   23
      Top             =   6375
      Width           =   660
   End
   Begin VB.CommandButton btnOR 
      Caption         =   "OR"
      Height          =   345
      Left            =   1830
      TabIndex        =   22
      Top             =   6375
      Width           =   660
   End
   Begin VB.CommandButton btnAND 
      Caption         =   "AND"
      Height          =   345
      Left            =   1020
      TabIndex        =   21
      Top             =   6375
      Width           =   660
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   2550
      TabIndex        =   12
      Top             =   5850
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   1905
      TabIndex        =   11
      Top             =   5865
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1380
      TabIndex        =   10
      Top             =   5865
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   2550
      TabIndex        =   9
      Top             =   5490
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1905
      TabIndex        =   8
      Top             =   5490
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1380
      TabIndex        =   7
      Top             =   5490
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2550
      TabIndex        =   6
      Top             =   5115
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1905
      TabIndex        =   5
      Top             =   5115
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1380
      TabIndex        =   4
      Top             =   5115
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2550
      TabIndex        =   3
      Top             =   4725
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1905
      TabIndex        =   2
      Top             =   4725
      Width           =   375
   End
   Begin VB.TextBox TrainData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Top             =   4725
      Width           =   375
   End
   Begin VB.CommandButton btnTrain 
      Caption         =   "Train"
      Height          =   465
      Left            =   795
      TabIndex        =   0
      Top             =   555
      Width           =   1005
   End
   Begin MSComctlLib.Slider sldHidden2 
      Height          =   270
      Index           =   1
      Left            =   4620
      TabIndex        =   50
      Top             =   4695
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldHidden1 
      Height          =   270
      Index           =   2
      Left            =   4605
      TabIndex        =   51
      Top             =   5820
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldHidden2 
      Height          =   270
      Index           =   2
      Left            =   4860
      TabIndex        =   52
      Top             =   6870
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldOutput 
      Height          =   270
      Index           =   1
      Left            =   8670
      TabIndex        =   53
      Top             =   4335
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldOutput 
      Height          =   270
      Index           =   2
      Left            =   8670
      TabIndex        =   54
      Top             =   6150
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldBiasO 
      Height          =   270
      Left            =   7560
      TabIndex        =   65
      Top             =   5205
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldBias1 
      Height          =   270
      Left            =   6975
      TabIndex        =   69
      Top             =   3615
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin MSComctlLib.Slider sldBias2 
      Height          =   270
      Left            =   7140
      TabIndex        =   72
      Top             =   6915
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      _Version        =   393216
      Min             =   -10
      SelectRange     =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Input 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   3735
      TabIndex        =   77
      Top             =   3900
      Width           =   780
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Input 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   3810
      TabIndex        =   76
      Top             =   6360
      Width           =   780
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   9630
      TabIndex        =   75
      Top             =   4815
      Width           =   780
   End
   Begin VB.Label lblBias2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7920
      TabIndex        =   74
      Top             =   6915
      Width           =   900
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Caption         =   "HN(2).Bias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7230
      TabIndex        =   73
      Top             =   6690
      Width           =   1455
   End
   Begin VB.Label lblBias1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7770
      TabIndex        =   71
      Top             =   3615
      Width           =   900
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Caption         =   "HN(1).Bias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7095
      TabIndex        =   70
      Top             =   3405
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Hidden 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   7320
      TabIndex        =   68
      Top             =   4230
      Width           =   915
   End
   Begin VB.Label lblBiasO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8355
      TabIndex        =   67
      Top             =   5205
      Width           =   900
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Caption         =   "ON.Bias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7665
      TabIndex        =   66
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Caption         =   "ON.Weight(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8775
      TabIndex        =   64
      Top             =   5925
      Width           =   1455
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   "ON.Weight(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8790
      TabIndex        =   63
      Top             =   4110
      Width           =   1455
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "HN(2).Weight(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4980
      TabIndex        =   62
      Top             =   6630
      Width           =   1455
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "HN(1).Weight(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4740
      TabIndex        =   61
      Top             =   5595
      Width           =   1455
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "HN(2).Weight(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4710
      TabIndex        =   60
      Top             =   4455
      Width           =   1455
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "HN(1).Weight(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5160
      TabIndex        =   59
      Top             =   3435
      Width           =   1455
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Output Data Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6420
      TabIndex        =   58
      Top             =   645
      Width           =   2265
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Training Data Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   615
      TabIndex        =   57
      Top             =   3795
      Width           =   2265
   End
   Begin VB.Label lblNetOut 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   7290
      TabIndex        =   46
      Top             =   2490
      Width           =   1350
   End
   Begin VB.Label lblNetOut 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   7290
      TabIndex        =   45
      Top             =   2190
      Width           =   1350
   End
   Begin VB.Label lblNetOut 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   7290
      TabIndex        =   44
      Top             =   1875
      Width           =   1350
   End
   Begin VB.Label lblNetOut 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   7290
      TabIndex        =   43
      Top             =   1575
      Width           =   1350
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6855
      TabIndex        =   42
      Top             =   2490
      Width           =   300
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6330
      TabIndex        =   41
      Top             =   2490
      Width           =   300
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6855
      TabIndex        =   40
      Top             =   2175
      Width           =   300
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6330
      TabIndex        =   39
      Top             =   2175
      Width           =   300
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6855
      TabIndex        =   38
      Top             =   1875
      Width           =   300
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6330
      TabIndex        =   37
      Top             =   1875
      Width           =   300
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6330
      TabIndex        =   36
      Top             =   1575
      Width           =   300
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6855
      TabIndex        =   35
      Top             =   1575
      Width           =   300
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Inputs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6420
      TabIndex        =   34
      Top             =   1005
      Width           =   675
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7350
      TabIndex        =   33
      Top             =   1005
      Width           =   675
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6345
      TabIndex        =   32
      Top             =   1290
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6870
      TabIndex        =   31
      Top             =   1290
      Width           =   300
   End
   Begin VB.Label lblOutput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   9465
      TabIndex        =   30
      Top             =   6150
      Width           =   900
   End
   Begin VB.Label lblOutput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   9465
      TabIndex        =   29
      Top             =   4335
      Width           =   900
   End
   Begin VB.Label lblHidden2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   5655
      TabIndex        =   28
      Top             =   6870
      Width           =   900
   End
   Begin VB.Label lblHidden2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   5400
      TabIndex        =   27
      Top             =   4695
      Width           =   900
   End
   Begin VB.Label lblHidden1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   5385
      TabIndex        =   26
      Top             =   5820
      Width           =   900
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Hidden 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   7485
      TabIndex        =   25
      Top             =   5985
      Width           =   915
   End
   Begin VB.Label lblHidden1 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   5865
      TabIndex        =   24
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Train 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   20
      Top             =   5895
      Width           =   795
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Train 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   19
      Top             =   5520
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Train 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   18
      Top             =   5130
      Width           =   795
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Train 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   17
      Top             =   4740
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1935
      TabIndex        =   16
      Top             =   4455
      Width           =   300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1410
      TabIndex        =   15
      Top             =   4455
      Width           =   300
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   14
      Top             =   4170
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Inputs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1485
      TabIndex        =   13
      Top             =   4170
      Width           =   675
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Basic Neural Network (with FAR too many comments)

' James Lewis, 2001

' Based on original code by Richard Gardner.

' This is a basic neural network, programmed and commented with
' the beginner to neural networks in mind.

' I'm just a neural net learner at this stage, but after
' reading up on how neurons and neural nets work in practice,
' I couldn't find much code that explained clearly and basically
' how a neural net works in practice - and how to put the theories
' into code.

' So after downloading Richard Gardner's code, I decided to change it
' around a bit, comment it, and add a few bits and bobs to give the
' new neural netter a kick start.

' The way this has been programmed is based on Richard's original code
' I have basically taken his code, changed variable names
' to make things a little more self explanatory, and added comments
' to describe the process and the math.
' Most of the credit for the code here should go directly to Richard.

' I hope you learn something from this. There may be errors in here,
' I may have learned some concepts wrongly, and there may be omissions
' that are very relevant. I give no guarantess, this is simply put
' together because I spent a hell of a long time looking for something
' like this to help me learn about neural networks and back prop.
' Please direct any comments, questions or suggestions to
' misterbodega@hotmail.com

' I'm assuming that if you are looking at this code, you have a reasonable
' background knowledge on what a neural network is, and basically how
' neurons work. If you don't, here are some basic definitions (if
' you don't know already the basic theory behind a neuron, go away and
' read about it first, then come back) :

' -------- The Neuron --------
' A neuron in its simplest form takes a number of inputs, multiplies
' each one individually by a 'weight', sums all the results of the
' inputs x weights, and applies an activation function on that figure
' to give an output.

' ----- A Neural Network -----
' A neural network is one or more neurons, connected together. A NN
' has one or more inputs, and one or more outputs. The basic idea of
' a neural network is that it after training it with data, it should
' be able to 'learn' a pattern to the data, and when subsequently
' provided with only the inputs, provide the required outputs.

' - A Backwards Propogation NN -
' Backwards propogation (or 'back-prop') refers to the method of working
' out how far wrong the output of a neural network is in its current
' state (the 'error'), and propogating a change in weights
' backwards through the
' network to correct this error

' -- XOR, AND, OR --
' Logical operations, go read up.

' Notes on the Code

' I've defined a Neuron data type, which I've then made an array of
' for the hidden layer, and there's only one neuron in the output layer.

' While I write the notes, I've been using XOR as the training input
' Some things may be specific to how the network looks with XOR training.

Const LEARNING_RATE As Double = 0.5

' The Learning Rate is a measure of how much the weights are changed
' in each training cycle. See the code.

Dim HiddenNeuron(1 To 2) As Neuron

' The Hidden layer Neurons (two of them)

Dim OutputNeuron As Neuron

' The Output layer Neuron

Option Explicit

Private Sub Train(Input1 As Double, Input2 As Double, Target As Double)

Dim k As Integer

' This takes two inputs and a target output and 'trains' the network.
' The basic idea behind this is to calculate the output of the network
' in its current state,
' have information provided on what it *should* have output (the target)
' and then
' propogate changes backwards through the network to correct the
' difference.

' See introduction for the basics on what a neuron is and how it works

HiddenNeuron(1).Output = Activation(HiddenNeuron(1).Bias + Input1 * HiddenNeuron(1).Weights(1) + Input2 * HiddenNeuron(1).Weights(2))
HiddenNeuron(2).Output = Activation(HiddenNeuron(2).Bias + Input1 * HiddenNeuron(2).Weights(1) + Input2 * HiddenNeuron(2).Weights(2))

' Find the current output for the Hidden Layer Neurons (sum of the bias
' plus the Inputs multiplied by the respective weights for each Neuron
' Note: HiddenNeuron(n).Weight(1) is the weight for Input1,
' HiddenNeuron(n).Weight(2) is the weight for Input2, and for the

OutputNeuron.Output = Activation(OutputNeuron.Bias + HiddenNeuron(1).Output * OutputNeuron.Weights(1) + HiddenNeuron(2).Output * OutputNeuron.Weights(2))

' In this case, the output neuron takes as its input the output from
' the two hidden layer Neurons. So for the output neuron weight(1) is the
' weight from HiddenNeuron(1), and weight(2) is the weight for
' HiddenNeuron(2)

' The next few stages are (basically) as follows:

' 1) Take a measure of how far wrong each neuron is.
' 2) Use this measure to correct the weights in the network.

' To make a little more sense of the code, bear in mind that
' we don't know how far wrong the Hidden Layer Neurons are
' until we know how wrong the output Neuron is - hence the Delta
' for the Hidden Neurons is calculated FROM the Delta of the
' Output Neuron. (see the backwards propogation idea? - you have to
' calculate the error change backwards because the information you
' have on what is correct propogates from the output back to the
' input)

OutputNeuron.Delta = OutputNeuron.Output * (1 - OutputNeuron.Output) * (Target - OutputNeuron.Output)

' The 'delta' is the measure of how far wrong the output from each
' neuron is.

' Once we have the delta, it allows us to make an alteration to the
' weights in the network. The bigger the Delta, the larger the error
' in the network, and so the larger we want to alter the weights.

' The above calculation of OutputNeuron.Delta first multiplies the
' output by (1- output). This has the effect of providing a larger
' figure when the output is at 0.5, and a minimum figure when the out
' put is at either 1 or 0 (do the math to confirm this). I.E. The Delta
' will be bigger, and so we're going to adjust the weight MORE when
' the current output is in the middle of the range (i.e. near 0.5). If
' the output is at either end of the range (i.e. at 1 or 0) then the
' Delta will come out smaller, and so we want to adjust the weight LESS.
' This simply has the effect of moving the weights more quickly if
' the current output from the Neuron is around 0.5 - the weight will
' be moved less if the neuron output is near 0 or near 1. (Bear in
' mind usually you'll want to get a more definite answer from
' a neural network - you want it to say 'Yes' or 'No' (i.e. 1 or 0)
' 0.5 corresponds to 'Maybe', which isn't a very useful answer.

' This figure is then multiplied by (Target - OutputNeuron.Output)
' This has the effect of making the delta LARGER if the error of the
' Neuron is larger.

' So overall this math says 'The Delta will be larger the nearer the
' Neuron output is to 1 or 0, and it will be larger the more wrong
' the Neuron is'.

HiddenNeuron(1).Delta = HiddenNeuron(1).Output * (1 - HiddenNeuron(1).Output) * (OutputNeuron.Weights(1) * OutputNeuron.Delta)
HiddenNeuron(2).Delta = HiddenNeuron(2).Output * (1 - HiddenNeuron(2).Output) * (OutputNeuron.Weights(2) * OutputNeuron.Delta)

' These deltas are the ones for the Hidden Layer. The math is similar
' here except for the last factor. Remember the Delta for each Neuron
' is how much we want to correct it by, but for the hidden layer, we
' don't have a specific figure of precisely what we want the output
' to be, so the Delta has to be calculated by how wrong the Output
' Neuron was (which is its Delta) and the current weight from the
' Hidden Neuron to the Output one. As far as I can see, the current
' weight is included as a factor here to reflect how 'important' that
' current weight is - the more important it is - i.e. the more its
' going to affect the Output Neuron, the more it should be altered.

' So now we have the delta for each Neuron - how much we want to change
' each Neuron's weights. So we'll use them to update the weights:

For k = 1 To 2

    HiddenNeuron(k).Bias = HiddenNeuron(k).Bias + LEARNING_RATE * 1 * HiddenNeuron(k).Delta
    HiddenNeuron(k).Weights(1) = HiddenNeuron(k).Weights(1) + LEARNING_RATE * Input1 * HiddenNeuron(k).Delta
    HiddenNeuron(k).Weights(2) = HiddenNeuron(k).Weights(2) + LEARNING_RATE * Input2 * HiddenNeuron(k).Delta

Next k

' See above how the Weight is altered by the Delta multiplied by the
' Learning rate - the larger the delta, and the larger the learning
' rate (which is a constant) - the more we're going to change each
' weight. But - the important part here is that we alter the weight
' of the Neuron also in terms of the INPUT. The larger the input was
' the more important this weight is to alter and so the more we're
' going to alter it by. - Bear this in mind when you look at how
' the weights for two neurons can start moving in the same direction
' initially and then change to moving in opposite directions - this
' is because of the Delta mainly being applied to a weight when
' there is a high input on that weight.

OutputNeuron.Bias = OutputNeuron.Bias + LEARNING_RATE * 1 * OutputNeuron.Delta
OutputNeuron.Weights(1) = OutputNeuron.Weights(1) + LEARNING_RATE * HiddenNeuron(1).Output * OutputNeuron.Delta
OutputNeuron.Weights(2) = OutputNeuron.Weights(2) + LEARNING_RATE * HiddenNeuron(2).Output * OutputNeuron.Delta

' And the same for the Output Neuron

End Sub

Public Function Activation(Value As Double)

    Activation = (1 / (1 + Exp(Value * -1)))
    
' The activation function gives us an output for each Neuron from 0
' to 1. Useful to understand the above - mainly in terms of how it
' looks on a graph.

End Function

Private Sub btnAND_Click()

' Set the training data to an AND operation

SetSeries
TrainData(2) = 0
TrainData(5) = 0
TrainData(8) = 0
TrainData(11) = 1

End Sub

Private Sub btnHelp_Click()

MsgBox ("This is a basic Neural Net. It begins with randomised weights throughout the network and the inputs set for XOR. You can change the input information if you want. Click the Train button to train the network once (which is actually once for each row in the training data table, or click the Train x 1000 to train 1000 times. You'll see the weights update as the network is training, and you can click Test at any point during or after training to test each of the four possible inputs. 'Reset Weights' randomises all the weights. Dragging the sliders will set a weight to somewhere inbetween -1 (far left) and 1 (far right). Note that training can give weights values outside this range. The XOR, OR and AND buttons simply put some preset training data into the training table. Note: the sliders may look like they go from -10 to 10, they don't - they slide from -1 to 1. Have fun, look at the code (at www.planetsourcecode.com), and email me with any questions or queries: misterbodega@hotmail.com")

End Sub

Private Sub btnOR_Click()

' Set the training data to an OR operation

SetSeries
TrainData(2) = 0
TrainData(5) = 1
TrainData(8) = 1
TrainData(11) = 1

End Sub

Private Sub btnReset_Click()

RandomiseWeights

UpdateWeightView

End Sub

Private Sub btnXOR_Click()

' Set the training data to an XOR operation - most interesting (IMO)

SetSeries
TrainData(2) = 0
TrainData(5) = 1
TrainData(8) = 1
TrainData(11) = 0

End Sub


Private Sub btnTest_Click()

' Run the network, and put the results in the output bit.

' (CDec(CSng) just formats so that it'll fit in the box

RunNetwork 0, 0

lblNetOut(0).Caption = CDec(CSng(OutputNeuron.Output))

RunNetwork 0, 1

lblNetOut(1).Caption = CDec(CSng(OutputNeuron.Output))

RunNetwork 1, 0

lblNetOut(2).Caption = CDec(CSng(OutputNeuron.Output))

RunNetwork 1, 1

lblNetOut(3).Caption = CDec(CSng(OutputNeuron.Output))

End Sub

Private Function RunNetwork(Input1 As Double, Input2 As Double) As Double

' Run network - No need really to pass the output back as the Neurons are
' all public, but what the hey.

' Should be obvious how this works if you know how a Neuron works
' - Take the activation function of the sum of all the inputs multiplied
' by their respective weights.

HiddenNeuron(1).Output = Activation(HiddenNeuron(1).Bias + HiddenNeuron(1).Weights(1) * Input1 + HiddenNeuron(1).Weights(2) * Input2)
HiddenNeuron(2).Output = Activation(HiddenNeuron(2).Bias + HiddenNeuron(2).Weights(1) * Input1 + HiddenNeuron(2).Weights(2) * Input2)
OutputNeuron.Output = Activation(OutputNeuron.Bias + OutputNeuron.Weights(1) * HiddenNeuron(1).Output + OutputNeuron.Weights(2) * HiddenNeuron(2).Output)

' Pass it back.

RunNetwork = OutputNeuron.Output

End Function

Private Sub btnTrain_Click()

' Train the 'net once for each of the four training lines

Dim k As Integer

For k = 0 To 11 Step 3

Train TrainData(k), TrainData(k + 1), TrainData(k + 2)

Next k

UpdateWeightView
DoEvents

End Sub

Private Sub SetSeries()
TrainData(0) = 0
TrainData(1) = 0
TrainData(3) = 0
TrainData(4) = 1
TrainData(6) = 1
TrainData(7) = 0
TrainData(9) = 1
TrainData(10) = 1
End Sub

Private Sub RandomiseWeights()
Dim k As Integer

For k = 1 To 2
    HiddenNeuron(k).Weights(1) = jRnd(2) - 1
    HiddenNeuron(k).Weights(2) = jRnd(2) - 1
    HiddenNeuron(k).Bias = Rnd(2) - 1
Next k

OutputNeuron.Weights(1) = jRnd(2) - 1
OutputNeuron.Weights(2) = jRnd(2) - 1
OutputNeuron.Bias = jRnd(2) - 1

UpdateWeightView

End Sub

Private Sub UpdateWeightView()

' Update all the sliders and their labels.

sldHidden1(1).Value = HiddenNeuron(1).Weights(1) * 10
sldHidden1(2).Value = HiddenNeuron(1).Weights(2) * 10

sldHidden2(1).Value = HiddenNeuron(2).Weights(1) * 10
sldHidden2(2).Value = HiddenNeuron(2).Weights(2) * 10

sldOutput(1).Value = OutputNeuron.Weights(1) * 10
sldOutput(2).Value = OutputNeuron.Weights(2) * 10

lblHidden1(1).Caption = CDec(CSng(HiddenNeuron(1).Weights(1)))
lblHidden1(2).Caption = CDec(CSng(HiddenNeuron(1).Weights(2)))
lblHidden2(1).Caption = CDec(CSng(HiddenNeuron(2).Weights(1)))
lblHidden2(2).Caption = CDec(CSng(HiddenNeuron(2).Weights(2)))
lblOutput(1).Caption = CDec(CSng(OutputNeuron.Weights(1)))
lblOutput(2).Caption = CDec(CSng(OutputNeuron.Weights(2)))

sldBias1.Value = HiddenNeuron(1).Bias * 10
sldBias2.Value = HiddenNeuron(2).Bias * 10
sldBiasO.Value = OutputNeuron.Bias * 10

lblBias1.Caption = CDec(CSng(HiddenNeuron(1).Bias))
lblBias2.Caption = CDec(CSng(HiddenNeuron(2).Bias))
lblBiasO.Caption = CDec(CSng(OutputNeuron.Bias))

End Sub

Private Sub Command1_Click()
Dim k As Integer
For k = 1 To 1000
btnTrain_Click
Next k

End Sub

Private Sub Form_Load()

' ReDim the Neurons to have the correct number of weights

' This is just written like this so that its flexible code
' Would be easy (for example) to increase the number of weights,
' Or add another OutputNeuron if you wanted to.

ReDim HiddenNeuron(1).Weights(2)
ReDim HiddenNeuron(2).Weights(2)
ReDim OutputNeuron.Weights(2)

ScaleMode = vbPixels
Me.Width = ScaleX(694 + 8, vbPixels, vbTwips)
Me.Height = ScaleY(492 + 28, vbPixels, vbTwips)

RandomiseWeights

btnXOR_Click

End Sub

Private Function jRnd(iRnd As Integer) As Double

' return a random value from 0 to iRnd

jRnd = iRnd * Rnd(1974)

End Function

Private Sub sldHidden1_Scroll(Index As Integer)

HiddenNeuron(1).Weights(Index) = sldHidden1(Index).Value / 10

UpdateWeightView

End Sub

Private Sub sldHidden2_Scroll(Index As Integer)

HiddenNeuron(2).Weights(Index) = sldHidden2(Index).Value / 10

UpdateWeightView

End Sub

Private Sub sldOutput_Scroll(Index As Integer)

OutputNeuron.Weights(Index) = sldOutput(Index).Value / 10

UpdateWeightView

End Sub

Private Sub sldBias1_Scroll()

HiddenNeuron(1).Bias = sldBias1.Value / 10

UpdateWeightView

End Sub

Private Sub sldBias2_Scroll()

HiddenNeuron(2).Bias = sldBias2.Value / 10

UpdateWeightView

End Sub

Private Sub sldBiasO_Scroll()

OutputNeuron.Bias = sldBiasO.Value / 10

UpdateWeightView

End Sub

