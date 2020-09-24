VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form area 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Area Calculator"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   240
      TabIndex        =   13
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton Command10 
         Caption         =   "Re-Calculate"
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   1890
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         TabIndex        =   15
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Rectangle Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Square Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Circle Area "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Area Unit Measurement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4895
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Circle"
      TabPicture(0)   =   "area.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Square"
      TabPicture(1)   =   "area.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Rectangle"
      TabPicture(2)   =   "area.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Display - Rectangle"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   5655
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   480
            TabIndex        =   9
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2880
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Area = HW (Height * Width)"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2760
            Width           =   5415
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   1695
            Left            =   1200
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   1320
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Display - Circle"
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5655
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3720
            TabIndex        =   5
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Area = Pie(3.14) * radius to the second power"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2760
            Width           =   5415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   6
            Top             =   1200
            Width           =   735
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   1695
            Left            =   960
            Shape           =   3  'Circle
            Top             =   600
            Width           =   3495
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   2640
            X2              =   3360
            Y1              =   1560
            Y2              =   1560
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Display - Square"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   5655
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Area = Side X * 4"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2760
            Width           =   5415
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Side X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   3
            Top             =   1320
            Width           =   615
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   1695
            Left            =   1080
            Shape           =   1  'Square
            Top             =   720
            Width           =   3495
         End
      End
   End
End
Attribute VB_Name = "area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
 If Combo1.Text = "" Then Combo1.Text = "Centimetres"
 If Text3.Text <> "" Then Text4.Text = (3.14 * (Val(Text3.Text) * Val(Text3.Text)))
 If Text5.Text <> "" Then Text6.Text = (Val(Text5.Text) * (Text5.Text))
 If Text1.Text <> "" And Text2.Text <> "" Then Text7.Text = (Text1.Text * Text2.Text)
 
 If Combo1.Text = "Centimetres" And Text4.Text <> "" Then Text4.Text = Text4.Text + " cm²"
 If Combo1.Text = "Centimetres" And Text6.Text <> "" Then Text6.Text = Text6.Text + " cm²"
 If Combo1.Text = "Centimetres" And Text7.Text <> "" Then Text7.Text = Text7.Text + " cm²"
 '²
 If Combo1.Text = "Millimetres" And Text4.Text <> "" Then Text4.Text = Text4.Text + " mm²"
 If Combo1.Text = "Millimetres" And Text6.Text <> "" Then Text6.Text = Text6.Text + " mm²"
 If Combo1.Text = "Millimetres" And Text7.Text <> "" Then Text7.Text = Text7.Text + " mm²"
   
End Sub

Private Sub Form_Load()
 Combo1.AddItem "Centimetres"
 Combo1.AddItem "Millimetres"
End Sub
