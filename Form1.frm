VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ToggleBox checkbox control"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   3915
      _extentx        =   6906
      _extenty        =   476
      value           =   -1  'True
      shadowline      =   0   'False
      backstyle       =   2
      caption         =   "Earth Protection Shield"
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   3915
      _extentx        =   6906
      _extenty        =   476
      value           =   0   'False
      shadowline      =   0   'False
      backstyle       =   2
      caption         =   "Ugly Toggle Box"
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   3915
      _extentx        =   6906
      _extenty        =   476
      value           =   -1  'True
      shadowline      =   0   'False
      backstyle       =   1
      caption         =   "Additional Back style mode"
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   735
      _extentx        =   1296
      _extenty        =   476
      value           =   -1  'True
      shadowline      =   0   'False
      backstyle       =   1
      caption         =   ""
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   4
      Left            =   2580
      TabIndex        =   4
      Top             =   1200
      Width           =   735
      _extentx        =   1296
      _extenty        =   476
      value           =   0   'False
      shadowline      =   0   'False
      backstyle       =   1
      caption         =   ""
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   5
      Left            =   180
      TabIndex        =   6
      Top             =   1560
      Width           =   3915
      _extentx        =   6906
      _extenty        =   476
      value           =   -1  'True
      shadowline      =   -1  'True
      backstyle       =   2
      caption         =   "Earth Protection Shield"
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   6
      Left            =   180
      TabIndex        =   7
      Top             =   1860
      Width           =   3915
      _extentx        =   6906
      _extenty        =   476
      value           =   0   'False
      shadowline      =   -1  'True
      backstyle       =   2
      caption         =   "Ugly Toggle Box"
   End
   Begin Project1.OptionBox OptionBox1 
      Height          =   270
      Index           =   7
      Left            =   180
      TabIndex        =   8
      Top             =   2160
      Width           =   3915
      _extentx        =   6906
      _extenty        =   476
      value           =   -1  'True
      shadowline      =   -1  'True
      backstyle       =   1
      caption         =   "Additional Back style mode"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Caption"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   1260
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

