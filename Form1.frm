VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Payment Of Salary"
   ClientHeight    =   7980
   ClientLeft      =   11415
   ClientTop       =   1515
   ClientWidth     =   8175
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   12
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListScreen1 
      Height          =   645
      Left            =   1890
      TabIndex        =   39
      Top             =   1980
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1138
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Currency"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "x"
         Text            =   "x"
         Object.Width           =   609
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Posotita"
         Text            =   "Quantity"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "="
         Text            =   "="
         Object.Width           =   660
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Synolo"
         Text            =   "Total"
         Object.Width           =   3351
      EndProperty
   End
   Begin MSComctlLib.ListView ListScreen2 
      Height          =   645
      Left            =   90
      TabIndex        =   35
      Top             =   1980
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1138
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Currency"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "x"
         Text            =   "x"
         Object.Width           =   609
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Posotita"
         Text            =   "Quantity"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "="
         Text            =   "="
         Object.Width           =   660
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Synolo"
         Text            =   "Total"
         Object.Width           =   3351
      EndProperty
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   15
      Left            =   2970
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "0"
      Top             =   1260
      Width           =   645
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   14
      Left            =   3780
      MaxLength       =   6
      TabIndex        =   32
      Top             =   2700
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   13
      Left            =   3780
      MaxLength       =   6
      TabIndex        =   31
      Top             =   1620
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   12
      Left            =   3780
      MaxLength       =   6
      TabIndex        =   30
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   11
      Left            =   3150
      MaxLength       =   6
      TabIndex        =   29
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   10
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   28
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   9
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   27
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   8
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   26
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   7
      Left            =   630
      MaxLength       =   6
      TabIndex        =   25
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   6
      Left            =   0
      MaxLength       =   6
      TabIndex        =   24
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   5
      Left            =   3150
      MaxLength       =   6
      TabIndex        =   23
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   4
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   22
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   3
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   21
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   2
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   20
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   1
      Left            =   630
      MaxLength       =   6
      TabIndex        =   19
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox TxtBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Index           =   0
      Left            =   0
      MaxLength       =   6
      TabIndex        =   18
      Top             =   540
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   14
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   13
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1170
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   12
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   11
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3150
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   10
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3150
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   9
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3150
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   8
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3150
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   7
      Left            =   630
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3150
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3150
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   5
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   4
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   3
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   2
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   1
      Left            =   630
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   500
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Value           =   2  'Grayed
      Width           =   600
   End
   Begin VB.CommandButton cmdDebug 
      Appearance      =   0  'Flat
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   34
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox ChkDebug 
      Caption         =   "Debug mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1890
      TabIndex        =   33
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdShow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4410
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6570
      Width           =   4245
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6570
      Width           =   4245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total"
      Height          =   285
      Left            =   90
      TabIndex        =   40
      Top             =   1620
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Section"
      Height          =   285
      Left            =   1890
      TabIndex        =   41
      Top             =   1620
      Width           =   1725
   End
   Begin VB.Label lblEmp 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      TabIndex        =   38
      Top             =   1260
      Width           =   285
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   " Salary:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1890
      TabIndex        =   17
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Employed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   37
      Top             =   1260
      Width           =   1365
   End
   Begin VB.Shape Ypoloipo 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1725
      Left            =   0
      Top             =   1080
      Width           =   3705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GotFocusArxPoso As Boolean
Dim TxtPoint, SumPl1, SumPl2 As Integer
Dim poso, Plithos(14), sum, sumd, Nomisma(14) As Single
Dim text1Change As Boolean
Dim InputText As String
Dim SumPlithos(14) As Single
Dim x As Integer
Dim PosSum As Single
Dim AllNull As Boolean
Dim EntCount As Integer
Dim ShowMoney As Boolean
Sub ArrangeObjects()

For i = 0 To 14
 If Check(i).Value = Unchecked Then
  If TxtBox(15).Text <> "" Then
   cmdShow.Enabled = True
  Else
  cmdShow.Enabled = False
  End If
   Exit For
 Else
   cmdShow.Enabled = False
 End If
Next i
End Sub

Sub ShowSumResults()
ListScreen2.ListItems.Clear
synolo = 0
 For i = 0 To 14
 If SumPlithos(i) > 0 Then
 EntCount = ListScreen2.ListItems.Count + 1
 ListScreen2.ListItems.Add EntCount, , Nomisma(i)
 ListScreen2.ListItems(EntCount).SubItems(1) = "x"
 ListScreen2.ListItems(EntCount).SubItems(2) = Int(SumPlithos(i))
 ListScreen2.ListItems(EntCount).SubItems(3) = "="
 ListScreen2.ListItems(EntCount).SubItems(4) = Format((Nomisma(i) * SumPlithos(i)), "Currency")
 End If
 synolo = synolo + (Nomisma(i) * SumPlithos(i))
 Next i
 EntCount = ListScreen2.ListItems.Count + 1
 ListScreen2.ListItems.Add EntCount, , ""
 ListScreen2.ListItems(EntCount).SubItems(4) = "*" & Format(synolo, "Currency") & ""

End Sub

Sub Process()
Dim Ypoloipo As Single
On Error Resume Next
Call Form_Resize
If ChkDebug.Value = vbChecked Then
 DebugForm.Show
 DebugForm.Cls
 ListScreen1.ListItems.Clear
 ListScreen1.HideColumnHeaders = False
 ListScreen1.Font = courier
 EntCount = ListScreen1.ListItems.Count + 1
 ListScreen1.ListItems.Add EntCount, , "*Debug mode"
 ListScreen1.ListItems(EntCount).SubItems(1) = "-"
 ListScreen1.ListItems(EntCount).SubItems(2) = "Testing edition"
 ListScreen1.ListItems(EntCount).SubItems(3) = "-"
 ListScreen1.ListItems(EntCount).SubItems(4) = "kwstarikanos@gmail.com"
 EntCount = ListScreen1.ListItems.Count + 1
 ListScreen1.ListItems.Add EntCount, , "i) Nomisma(i)"
 ListScreen1.ListItems(EntCount).SubItems(1) = "Plithos(i)"
 ListScreen1.ListItems(EntCount).SubItems(2) = "sum"
 ListScreen1.ListItems(EntCount).SubItems(3) = "Ypoloipo"
 ListScreen1.ListItems(EntCount).SubItems(4) = "Plithos(i) > TxtBox(i)"
Else
 DebugForm.Hide
 ListScreen1.HideColumnHeaders = True
 ListScreen1.Font = "Simplified Arabic"
End If

If ChkDebug.Value = vbChecked Then
EntCount = ListScreen1.ListItems.Count + 1
ListScreen1.ListItems.Add EntCount, , i & "= "
ListScreen1.ListItems(EntCount).SubItems(3) = "(" & Ypoloipo & ")"
End If

Dim n, iRemain, CurrentSum As Integer
Ypoloipo = TxtBox(15) * 100
Ypoloipo = Int(Ypoloipo)
sum = 0

For i = 0 To 14
'Arxikopoihsh____________________________________________________________________________
n = 0
iRemain = 0
Plithos(i) = 0
DebugMessage = ""
Next1Coin = False

'Starting________________________________________________________________________________
If Check(i).Value = Unchecked Then
DebugForm.Print "--|Start|--" & Nomisma(i), "Τιμή εισόδου <- "; Ypoloipo, " | "; Ypoloipo / 100
Plithos(i) = Int(Ypoloipo) \ Nomisma(i) * 100

If Int(Ypoloipo) Mod Nomisma(i) * 100 = 0 Then
 DebugForm.Print "Το " & Nomisma(i) & " χωράει ακριβώς " & Ypoloipo \ Nomisma(i) * 100 & " φορές στο " & Ypoloipo / 100 & ""
 Ypoloipo = Ypoloipo Mod Nomisma(i) * 100
Else
DebugForm.Print "Το " & Nomisma(i) & " δεν χωράει ακριβώς φορές στο  " & Int(Ypoloipo) / 100 & " και αφήνει υπόλοιπο: "; (Int(Ypoloipo) Mod Nomisma(i) * 100) / 100
CurrentSum = Plithos(i) * Nomisma(i)
iRemain = Mid(Ypoloipo Mod CurrentSum * 100, 1, 1)
DebugForm.Print "" & Int(Ypoloipo) / 100 & " Mod " & Plithos(i) * Nomisma(i) & " = "; Int(Ypoloipo) / 100 Mod Plithos(i) * Nomisma(i); ""

DebugForm.Print "iRemain = "; iRemain


If Mid(Nomisma(i) * 100, 1, 1) = 5 Or Mid(Nomisma(i) * 100, 1, 1) = 2 Then
  If Mid(Nomisma(i + 1) * 100, 1, 1) = 1 And Check(i + 1).Value = Unchecked Then
  Next1Coin = True
  End If
  If Mid(Nomisma(i + 2) * 100, 1, 1) = 1 And Check(i + 2).Value = Unchecked Then
  Next1Coin = True
  End If
End If

  If Plithos(i) > 0 And Next1Coin = False Then
   Select Case iRemain
     Case 1, 3
DebugForm.Print "Plithos(" & i & ") = "; Plithos(i) & " - 1 = "; Plithos(i) - 1
     Plithos(i) = Plithos(i) - 1
     End Select
  End If


CurrentSum = Plithos(i) * Nomisma(i)
Ypoloipo = Ypoloipo - CurrentSum * 100
End If

SumPlithos(i) = SumPlithos(i) + Plithos(i)
sum = sum + Plithos(i) * Nomisma(i)
DebugForm.Print , "Τιμή εξόδου - >"; , Ypoloipo, " | "; Ypoloipo / 100
DebugForm.Print "--|End|" & "___________________________________________________________"
DebugForm.Print ""
End If

  If Plithos(i) > 0 Then
    If ChkDebug.Value = vbChecked Then
     EntCount = ListScreen1.ListItems.Count + 1
     ListScreen1.ListItems.Add EntCount, , i & ") " & Nomisma(i)
     ListScreen1.ListItems(EntCount).SubItems(1) = " x " & Plithos(i) & " = " & Format(Plithos(i) * Nomisma(i), "Currency")
     ListScreen1.ListItems(EntCount).SubItems(2) = Format(sum, "Currency")
     ListScreen1.ListItems(EntCount).SubItems(3) = Ypoloipo
     ListScreen1.ListItems(EntCount).SubItems(4) = Plithos(i) > TxtBox(i)
    End If
  End If
Next i

 PosSum = 0
 For i = 0 To 14
 PosSum = PosSum + Plithos(i) * Nomisma(i)
 Next i
 ShowMoney = False
End Sub

Sub ShowResults()
'If ChkDebug.Value = vbChecked Then

'Else
ListScreen1.Font.Size = 12
TxtBox(15).Text = ""
TxtBox(15).SetFocus
cmdShow.Enabled = False
lblEmp.Caption = x
EntCount = ListScreen1.ListItems.Count + 1

If ChkDebug.Value = Checked Then
ListScreen1.ListItems.Add EntCount, , ""
lblEmp.Caption = ""
Else
ListScreen1.ListItems.Add EntCount, , "x (" & x & ")"
End If

For i = 0 To 14
 If Plithos(i) > 0 Then
 EntCount = ListScreen1.ListItems.Count + 1
 ListScreen1.ListItems.Add EntCount, , Nomisma(i)
 ListScreen1.ListItems(EntCount).SubItems(1) = "x"
 ListScreen1.ListItems(EntCount).SubItems(2) = Int(Plithos(i))
 ListScreen1.ListItems(EntCount).SubItems(3) = "="
 ListScreen1.ListItems(EntCount).SubItems(4) = Format((Nomisma(i) * Plithos(i)), "Currency")
 End If
Next i
 EntCount = ListScreen1.ListItems.Count + 1
 ListScreen1.ListItems.Add EntCount, , ""
 ListScreen1.ListItems(EntCount).SubItems(4) = "*" & Format(sum, "Currency")
 sum = 0
'End If
End Sub

Private Sub Check_Click(Index As Integer)
Call ArrangeObjects


'Colors
If Check(Index).Value = vbChecked Then
Check(Index).ForeColor = &H8000000C
Check(Index).BackColor = &HE0E0E0
TxtBox(Index).Text = "" 'Otan den einai tsekarismeno na midenistei
TxtBox(Index).BackColor = &H8000000B
Else
Check(Index).ForeColor = vbBlack
Select Case Index
Case 0
Check(Index).BackColor = &HFFC0FF
Case 1
Check(Index).BackColor = &HC0FFFF
Case 2
Check(Index).BackColor = &HC0FFC0
Case 3
Check(Index).BackColor = &HC0E0FF
Case 4
Check(Index).BackColor = &HFFC0C0
Case 5
Check(Index).BackColor = &HC0C0FF
Case 6
Check(Index).BackColor = &HC0C000
Case 7 To 8
Check(Index).BackColor = &H80000018
Case 9 To 11
Check(Index).BackColor = &H80FFFF
Case 12 To 14
Check(Index).BackColor = &H8080FF
End Select

' If TxtBox(Index).Text = "" Then
'  InputText = InputBox("Παρακαλώ εισάγεται πλήθος για το νόμισμα: " & Format(Nomisma(Index), "Currency"), "Εισαγωγή πλήθους για νόμισμα: " & Format(Nomisma(Index), "Currency") & "", "0")
'  If InputText = "" Then InputText = 0
'  TxtBox(Index).Text = InputText
' End If
End If
End Sub










Private Sub cmdReset_Click()
ListScreen1.ListItems.Clear
ListScreen2.ListItems.Clear
x = 0
PosSum = 0
Ypoloipo = 0
AllNull = True



For i = 0 To 14
Plithos(i) = 0
SumPlithos(i) = 0
Next i
lblEmp.Caption = ""
TxtBox(15).Text = ""
TxtBox(15).SetFocus
End Sub

Private Sub cmdShow_Click()
On Error Resume Next

If TxtBox(15).Text = "00" Then
ChkDebug.Value = vbChecked

ElseIf TxtBox(15).Text = "0" Then
ChkDebug.Value = vbUnchecked
cmdReset_Click

ElseIf TxtBox(15).Text = "000" Then
ListScreen1.Visible = False
ListScreen2.Visible = False
Form1.Height = 7600
Form1.FontSize = 12
Form1.ForeColor = vbGreen
Form1.BackColor = vbBlack
Form1.Font = courier
Form1.FontBold = False
Form1.Cls
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print " For i = 0 To 14", , , "Nomisma(0) = 500"
Print " If Check(i).Value = Unchecked Then", , "Nomisma(1) = 200"
Print "", , , , "Nomisma(2) = 100"
Print "   Plithos(i) = Int(Ypoloipo) \ Nomisma(i) * 100", "Nomisma(3) = 50"
Print "   If Plithos(i) > Int(TxtBox(i).Text) Then", , "Nomisma(4) = 20"
Print "     Plithos(i) = Int(TxtBox(i).Text)", , "Nomisma(5) = 10"
Print "     Ypoloipo = Ypoloipo - Plithos(i) * Nomisma(i)", "Nomisma(6) = 5"
Print "     If Ypoloipo < 0 Then Ypoloipo = 0", , "Nomisma(7) = 2"
Print "   Else", , , , "Nomisma(8) = 1"
Print "     Ypoloipo = Ypoloipo Mod Nomisma(i) * 100", "Nomisma(9) = 0.50"
Print "  End If", , , , "Nomisma(10) = 0.20"
Print "     sum = sum + Plithos(i) * Nomisma(i)", , "Nomisma(11) = 0.10"
Print " Else", , , , "Nomisma(12) = 0.05"
Print "  Plithos(i) = 0", , , "Nomisma(13) = 0.02"
Print "", , , , "Nomisma(14) = 0.01"
Print " End If"
Print " Next i"

Else
ListScreen1.Visible = True
ListScreen2.Visible = True
End If

Call Process

 x = x + 1
 If Format(PosSum, "Currency") <> Format(TxtBox(15).Text, "Currency") Then
   x = x - 1
   If Format(PosSum, "Currency") < Format(TxtBox(15).Text, "Currency") Then
    MsgBox "Το ποσό των χρημάτων που επιλέξατε για τον " & x + 1 & " εργαζόμενο δεν επαρκεί, η διαφορά είναι: " & Format(TxtBox(15).Text - PosSum, "Currency"), vbInformation
   Else
    MsgBox "Το ποσό των χρημάτων που επιλέξατε για τον " & x + 1 & " εργαζόμενο είναι παραπάνω, η διαφορά είναι: " & Format(PosSum - TxtBox(15).Text, "Currency"), vbInformation
   End If
  For i = 0 To 14
   SumPlithos(i) = SumPlithos(i) - Plithos(i)
   Plithos(i) = 0
  Next i
 Else
  Call ShowResults
  Call ShowSumResults
 End If

TxtBox(15).SetFocus
End Sub






Private Sub Form_Load()
For i = 0 To 14
 Check(i).Value = vbUnchecked
Next i


Nomisma(0) = 500
Nomisma(1) = 200
Nomisma(2) = 100
Nomisma(3) = 50
Nomisma(4) = 20
Nomisma(5) = 10
Nomisma(6) = 5
Nomisma(7) = 2
Nomisma(8) = 1
Nomisma(9) = 0.5
Nomisma(10) = 0.2
Nomisma(11) = 0.1
Nomisma(12) = 0.05
Nomisma(13) = 0.02
Nomisma(14) = 0.01
x = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
Ypoloipo.Top = 0
Ypoloipo.Height = 1905
Ypoloipo.Left = 0
Ypoloipo.Width = Form1.Width

Check(0).Top = 0
Check(0).Height = 500
Check(0).Left = 0
Check(0).Width = Form1.Width \ 15 - 35

TxtBox(0).Top = 540
TxtBox(0).Height = 500
TxtBox(0).Left = 0
TxtBox(0).Width = Check(0).Width

For i = 1 To 14
 Check(i).Top = 0
 Check(i).Height = 500
 Check(i).Left = Check(i - 1).Left + Check(i - 1).Width + 20
 Check(i).Width = Check(i - 1).Width
 
 TxtBox(i).Top = 540
 TxtBox(i).Height = 500
 TxtBox(i).Left = TxtBox(i - 1).Left + TxtBox(i - 1).Width + 20
 TxtBox(i).Width = TxtBox(i - 1).Width
Next i

Label1.Top = 1100
Label1.Height = 375
Label1.Left = 0
Label1.Width = 1365

lblEmp.Top = 1100
lblEmp.Height = 375
lblEmp.Left = 1350
lblEmp.Width = Form1.Width / 2 - 1365

lbl3.Top = 1100
lbl3.Height = 375
lbl3.Left = Form1.Width / 2
lbl3.Width = Form1.Width / 2 / 1.5

TxtBox(15).Top = 1080
TxtBox(15).Height = 390
TxtBox(15).Left = lbl3.Left + lbl3.Width + 40
TxtBox(15).Width = Form1.Width - TxtBox(15).Left - 240
Label2.Top = 1530
Label2.Height = 285
Label2.Left = 0
Label2.Width = Form1.Width / 2
Label3.Top = 1530
Label3.Height = 285
Label3.Left = Form1.Width / 2
Label3.Width = Form1.Width / 2

'ListScreens
ListScreen2.Top = 1890
ListScreen2.Height = Form1.Height - 1890 - 570
ListScreen2.Left = 0
ListScreen2.Width = Form1.Width / 2
ListScreen2.ColumnHeaders(1).Width = ListScreen2.Width \ 5
ListScreen2.ColumnHeaders(2).Width = ListScreen2.ColumnHeaders(1).Width - 400
ListScreen2.ColumnHeaders(3).Width = ListScreen2.ColumnHeaders(1).Width + 200
ListScreen2.ColumnHeaders(4).Width = ListScreen2.ColumnHeaders(1).Width - 400
ListScreen2.ColumnHeaders(5).Width = ListScreen2.ColumnHeaders(1).Width + 500

 ListScreen1.Top = 1890
 ListScreen1.Height = Form1.Height - 1890 - 570
If ChkDebug.Value = vbChecked Then
 ListScreen1.Left = 0
 ListScreen1.Width = Form1.Width
 ListScreen1.ColumnHeaders(1).Width = ListScreen1.Width \ 5
 ListScreen1.ColumnHeaders(2).Width = ListScreen1.ColumnHeaders(1).Width - 900
 ListScreen1.ColumnHeaders(3).Width = ListScreen1.ColumnHeaders(1).Width + 200
 ListScreen1.ColumnHeaders(4).Width = ListScreen1.ColumnHeaders(1).Width - 1000
 ListScreen1.ColumnHeaders(5).Width = ListScreen1.ColumnHeaders(1).Width + 1400
Else
 ListScreen1.Left = Form1.Width / 2
 ListScreen1.Width = Form1.Width / 2 - 235
 ListScreen1.ColumnHeaders(1).Width = ListScreen1.Width \ 5
 ListScreen1.ColumnHeaders(2).Width = ListScreen1.ColumnHeaders(1).Width - 400
 ListScreen1.ColumnHeaders(3).Width = ListScreen1.ColumnHeaders(1).Width + 200
 ListScreen1.ColumnHeaders(4).Width = ListScreen1.ColumnHeaders(1).Width - 400
 ListScreen1.ColumnHeaders(5).Width = ListScreen1.ColumnHeaders(1).Width + 540
End If








End Sub

Private Sub TxtBox_Change(Index As Integer)
'On Error Resume Next
 'Call Process
 'For i = 0 To 14
 'If Plithos(i) > 0 Then
 'Check(i).Value = Unchecked
 'Plithos(i) = 0
 'Else
 'Check(i).Value = Checked
 'End If
 'Next i



If Index <> 15 Then
  If TxtBox(Index).Text <> "" Then
    If TxtBox(Index).Text > 0 Then
    Check(Index).Value = vbUnchecked
    Else
    Check(Index).Value = Checked
    End If
  End If
Else

End If


Call ArrangeObjects
End Sub

Private Sub TxtBox_GotFocus(Index As Integer)
On Error Resume Next
If TxtBox(Index).Text = 0 Then TxtBox(Index).Text = ""
TxtBox(Index).BackColor = vbWhite
TxtPoint = Index
Call ArrangeObjects
End Sub


Private Sub TxtBox_KeyPress(Index As Integer, KeyAscii As Integer)
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 44 Then
  KeyAscii = 0
  End If
  If KeyAscii = 46 Then
  KeyAscii = 44
  End If
  If TxtBox(Index).Text = "" And KeyAscii = 44 Then
  KeyAscii = 0
  End If
  If (KeyAscii = 46 Or KeyAscii = 44) And InStr(TxtBox(Index).Text, ",") Then
  KeyAscii = 0
  End If
  'Eisagogi mono 2 psifiwn meta tin ypodiastoli
  If KeyAscii <> 8 Then
    If InStr(TxtBox(Index).Text, ",") Then
      If Len(TxtBox(Index).Text) - InStr(TxtBox(Index).Text, ",") >= 2 Then
      KeyAscii = 0
      End If
   End If
  End If
If KeyAscii = 0 Then Beep
End Sub

Private Sub TxtBox_LostFocus(Index As Integer)

If Index < 15 Then
If TxtBox(Index).Text = "" Then
TxtBox(Index).Text = ""
TxtBox(Index).BackColor = &H8000000B
Check(Index).Value = Checked
End If

If TxtBox(Index).Text = "0" Then
TxtBox(Index).Text = ""
TxtBox(Index).BackColor = &H8000000B
End If

Else
If TxtBox(Index).Text = "" Then
TxtBox(Index).Text = ""
TxtBox(Index).BackColor = &H8000000B
End If

End If


End Sub

Sub Diathesima()
sum = 0
SumPl1 = 0

Print ""
For i = 0 To 14
If TxtBox(i).Text = "" Then TxtBox(i).BackColor = &H8000000B
If TxtBox(i).Text > 0 Then
Print Nomisma(i), "x", Int(TxtBox(i)), Format(Nomisma(i) * Int(TxtBox(i)), "Currency")
End If

SumPl1 = SumPl1 + TxtBox(i)
sum = sum + (Nomisma(i) * Int(TxtBox(i)))
Next i

'TxtBox(15) = sum
Print "_______________________________________________________________"
Print "Διαθέσιμα:", , "(" & SumPl1 & ")", Format(sum, "Currency")
End Sub


