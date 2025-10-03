VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Kriptonian 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GCC - Documentos Privados"
   ClientHeight    =   11685
   ClientLeft      =   5100
   ClientTop       =   3330
   ClientWidth     =   19350
   Icon            =   "mdAesCtr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11685
   ScaleWidth      =   19350
   Begin VB.Frame Frames 
      Caption         =   "Base De Datos"
      Height          =   10815
      Index           =   4
      Left            =   9240
      TabIndex        =   30
      Top             =   600
      Width           =   9855
      Begin MSDataGridLib.DataGrid DataGrid 
         Height          =   10215
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   18018
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   27
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight Condensed"
            Size            =   18
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Listado"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frames 
      Caption         =   "Base De Datos"
      Height          =   1575
      Index           =   3
      Left            =   5160
      TabIndex        =   23
      Top             =   2280
      Width           =   3735
      Begin VB.CommandButton UpdateRecordDB 
         Caption         =   "Actualiza Cambios En Registro"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton CreateRecordDB 
         Caption         =   "Nueva Entrada"
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
         Left            =   1800
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton ReadDB 
         Caption         =   "Abrir BD"
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
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton CreateTable 
         Caption         =   "Crear Tabla"
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
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton CreateDB 
         Caption         =   "Crear BD"
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frames 
      Caption         =   "Seguridad"
      Height          =   1455
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   600
      Width           =   8535
      Begin VB.CommandButton cmdXport 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Xport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Exporta la BD"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClearControls 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Limpia el contenido de todos los controles"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton save 
         Caption         =   "&Guarda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   27
         ToolTipText     =   "Guarda el texto cifrado a un archivo"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton load 
         Caption         =   "&Abre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   26
         ToolTipText     =   "Abre el archivo con el texto cifrado"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Desencripta 
         Caption         =   "&Desencripta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Encripta 
         Caption         =   "&Encripta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdPanicButton 
         BackColor       =   &H000000FF&
         Height          =   465
         Left            =   7935
         MaskColor       =   &H000000FF&
         Picture         =   "mdAesCtr.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Botón De Pánico"
         Top             =   255
         Width           =   465
      End
      Begin VB.TextBox Password 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   720
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7920
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Bahnschrift"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frames 
      Caption         =   "Encriptado"
      Height          =   3015
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   8400
      Width           =   8655
      Begin VB.TextBox Encriptado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Cascadia Code SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2535
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Text            =   "mdAesCtr.frx":0BD4
         Top             =   360
         Width           =   8175
      End
   End
   Begin VB.Frame Frames 
      Caption         =   "Texto"
      Height          =   4215
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   3960
      Width           =   8655
      Begin VB.TextBox Texto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Text            =   "mdAesCtr.frx":0BDE
         Top             =   360
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "Col"
      Height          =   360
      Index           =   33
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Pone color a la selección "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "Fnt"
      Height          =   360
      Index           =   32
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   " Cambia la fuente "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "ST"
      Height          =   360
      Index           =   31
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Tacha la selección "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "U"
      Height          =   360
      Index           =   30
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Subraya la selección "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "I"
      Height          =   360
      Index           =   29
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   " Pone cursiva a la selección "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "N"
      Height          =   360
      Index           =   28
      Left            =   2640
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " Pone negrita a la selección "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "D"
      Height          =   360
      Index           =   27
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " Justifica por la derecha "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "C"
      Height          =   360
      Index           =   26
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Centra el texto "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "I"
      Height          =   360
      Index           =   25
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Justifica por la izquierda "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "P"
      Height          =   360
      Index           =   24
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Imprimir "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "S"
      Height          =   360
      Index           =   23
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Guardar "
      Top             =   11760
      Width           =   360
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "L"
      Height          =   360
      Index           =   22
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Abrir "
      Top             =   11760
      Width           =   360
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift Light Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift Light Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   12120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"mdAesCtr.frx":0BE9
   End
   Begin MSComDlg.CommonDialog cmdlgAgenda 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFCC&
      Caption         =   "El archivo se descifra al cargarlo y se cifra al guardarlo, no hay modo de recuperarlo si olvidas el password."
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   19815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Categoría"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "Kriptonian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'To resize form
Private Type ControlInfo_type
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type
Dim ControlInfos() As ControlInfo_type
'To resize form




Dim lastIndex As Long
Dim searchText As String
Dim foundPositions() As Long
Dim currentPosition As Integer

Public MyDBGrdRs As ADODB.Recordset

Sub BackupDB()
    'FileCopy App.Path & "\Kripto.accdb", App.Path & "\Kripto.accdb_" & Mid(Date, 4, 2) & Left(Date, 2) & Right(Date, 4) & "_" & Left(Time, 8)
End Sub
Sub SetupDg()
    DataGrid.Columns(0).Width = 500
    DataGrid.Columns(1).Width = 4000
    DataGrid.Columns(2).Width = 4000
    DataGrid.Columns(3).Width = 0
    DataGrid.Columns(4).Width = 0 '1400
    DataGrid.Columns(5).Width = 0 '800
    DataGrid.Columns(6).Width = 0 '800

    DataGrid.Columns(0).Alignment = dbgCenter
    DataGrid.Columns(1).Alignment = dbgLeft
    DataGrid.Columns(2).Alignment = dbgLeft
    'DataGrid.Columns(4).Alignment = dbgCenter
    'DataGrid.Columns(5).Alignment = dbgCenter

    'DataGrid.Columns(4).NumberFormat = "DD MM YY"
End Sub
Private Sub UpdtUsr()
    'Update Table
    With Chazaro
        .Execute "UPDATE Agents " _
                & "SET Agent = 'LordMaster' " _
                & ", Codex = 'U2FsdGVkX1+SoWaZ0O0DxdylpXrG0o5wZK0a7A=='" _
                & ", Created = '" & Date & "'" _
                & ", Status = True"
    End With
End Sub

Private Sub TestEncrypt()
'example
    Dim sPass       As String
    Dim sText       As String
    Dim sEncr       As String
    
    sPass = "password123"
    sText = "this is a test"
    sEncr = AesEncryptString(sText, sPass)
    Debug.Assert sText = AesDecryptString(sEncr, sPass)
    
    Debug.Print "Result (Base64): " & sEncr
    Debug.Print "Raw byte-array:  " & StrConv(FromBase64Array(sEncr), vbUnicode)
    Debug.Print "Decrypted:       " & AesDecryptString(sEncr, sPass)
End Sub
    
Private Sub TestHmac()
'example
    Dim baEncr()    As Byte
    Dim baHmacEncr(0 To 31) As Byte
    Dim baHmacDecr(0 To 31) As Byte
    
    baEncr = ToUtf8Array("test message")
    baHmacEncr(0) = 0           '--- 0 -> generate hash before encrypting
    AesCryptArray baEncr, ToUtf8Array("pass"), Hmac:=baHmacEncr
    baHmacDecr(0) = 1           '--- 1 -> decrypt and generate hash after that
    AesCryptArray baEncr, ToUtf8Array("pass"), Hmac:=baHmacDecr
    Debug.Assert InStrB(1, baHmacDecr, baHmacEncr) = 1
    
    Debug.Print "baHmacDecr: " & StrConv(baHmacDecr, vbUnicode)
    Debug.Print "baHmacEncr: " & StrConv(baHmacEncr, vbUnicode)
End Sub

Private Sub cmdClearControls_Click()
    Dim Chzr As Integer
    Password.Text = ""
    Texto.Text = ""
    Encriptado.Text = ""
    For Chzr = 0 To 1
        txtFields(Chzr).Text = ""
    Next
    Password.PasswordChar = "*"
End Sub
Private Sub cmdCommands_Click(Index As Integer)
Select Case Index
            Case Is = 22 'abre
2300            cmdlgAgenda.Filter = "Rich Text Format files|*.rtf"
2310            cmdlgAgenda.ShowOpen
2320            RTB.LoadFile cmdlgAgenda.FileName, rtfRTF
2330        Case Is = 23 'guarda
2340            cmdlgAgenda.ShowSave
2350            RTB.SaveFile (cmdlgAgenda.FileName) ', textRTF)
2360        Case Is = 24 'imprime
2370            cmdlgAgenda.Flags = cdlPDReturnDC + cdlPDNoPageNums
2380            If RTB.SelLength = 0 Then
2390                cmdlgAgenda.Flags = cmdlgAgenda.Flags + cdlPDAllPages
2400            Else
2410                cmdlgAgenda.Flags = cmdlgAgenda.Flags + cdlPDSelection
2420            End If
2430            cmdlgAgenda.ShowPrinter
2440            'Printer.Print ""
2450            RTB.SelPrint cmdlgAgenda.hDC
2460            Printer.EndDoc
2470        Case Is = 25: RTB.SelAlignment = rtfLeft 'vbAlignLeft
2480        Case Is = 26: RTB.SelAlignment = rtfCenter 'vbCenter
2490        Case Is = 27: RTB.SelAlignment = rtfRight 'vbAlignRight
2500        Case Is = 28
2510            If RTB.SelBold = True Then
2520                RTB.SelBold = False
2530            Else
2540                RTB.SelBold = True
2550            End If
2560        Case Is = 29
2570            If RTB.SelItalic = True Then
2580                RTB.SelItalic = False
2590            Else
2600                RTB.SelItalic = True
2610            End If
2620        Case Is = 30
2630            If RTB.SelUnderline = True Then
2640                RTB.SelUnderline = False
2650            Else
2660                RTB.SelUnderline = True
2670            End If
2680        Case Is = 31
2690            If RTB.SelStrikeThru = True Then
2700                RTB.SelStrikeThru = False
2710            Else
2720                RTB.SelStrikeThru = True
2730            End If
2740        Case Is = 32 'fuentes
2750            cmdlgAgenda.Flags = cdlCFBoth
2760            cmdlgAgenda.ShowFont
2770            With RTB
2780                .SelFontName = cmdlgAgenda.FontName
2790                .SelFontSize = cmdlgAgenda.FontSize
2800                .SelBold = cmdlgAgenda.FontBold
2810                .SelItalic = cmdlgAgenda.FontItalic
2820                .SelStrikeThru = cmdlgAgenda.FontStrikethru
2830                .SelUnderline = cmdlgAgenda.FontUnderline
2840            End With
2850        Case Is = 33 'color
2860            cmdlgAgenda.ShowColor
2870            RTB.SelColor = cmdlgAgenda.Color
2880        Case Is = 34: MsgBox "Intenta con F1", vbOKOnly, "¿No puedes?" 'ayuda
    End Select
End Sub

Private Sub cmdPanicButton_Click()
    On Error Resume Next
    Set DataGrid.DataSource = Nothing
    Chazaro.Close
    cmdClearControls_Click
    End
End Sub

Private Sub cmdXport_Click()
' Requires reference to Microsoft ActiveX Data Objects
Dim Rs As ADODB.Recordset
Dim FSO As Object
Dim FileOut As Object
Dim SqlOut As Object
Dim DecryptedText As String
Dim LineOut As String
Dim rawdata As Variant
Dim SqlLine As String

' Defensive: validate connection object
If Chazaro Is Nothing Then
    MsgBox "Connection object is not initialised.", vbCritical
    Exit Sub
End If

Set Rs = New ADODB.Recordset
Rs.Open "SELECT Category, Title, Content, Data FROM TopSecret", Chazaro, adOpenStatic, adLockOptimistic

' Defensive: validate recordset
If Rs.EOF Then
    MsgBox "No records found.", vbInformation
    Rs.Close
    Exit Sub
End If

' Create Markdown file
Set FSO = CreateObject("Scripting.FileSystemObject")

Set FileOut = FSO.CreateTextFile("TopSecretDump.md", True, True) ' Unicode mode
Set SqlOut = FSO.CreateTextFile("TopSecretDump.sql", True)

Do Until Rs.EOF
    rawdata = Rs("Content").Value
    If Not IsNull(rawdata) Then
        DecryptedText = AesDecryptString(CStr(rawdata), "3r937hK0g10n")
        Texto.Text = Texto.Text & vbCr & DecryptedText
        SqlLine = "INSERT INTO SciPhr (Category, Title, Data, Cipher, reg_date, Status) VALUES (" & _
                  "'" & Replace(SafeText(Rs("Category").Value), "'", "''") & "', " & _
                  "'" & Replace(SafeText(Rs("Title").Value), "'", "''") & "', " & _
                  "'" & Replace(SafeText(DecryptedText), "'", "''") & "', " & _
                  "'" & Replace(SafeText(Rs("Content").Value), "'", "''") & "', " & _
                  "'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "', " & _
                  "1);"
    Else
        DecryptedText = "[NULL]"
    End If

    LineOut = "## " & SafeText(Rs("Category").Value) & vbCrLf
    LineOut = LineOut & "**" & SafeText(Rs("Title").Value) & "**" & vbCrLf & vbCrLf
    LineOut = LineOut & SafeText(DecryptedText) & vbCrLf & vbCrLf

    FileOut.WriteLine LineOut
    SqlOut.WriteLine SqlLine
    Rs.MoveNext
Loop

FileOut.Close
Rs.Close

MsgBox "Markdown file created successfully.", vbInformation
End Sub
Function SafeText(ByVal V As Variant) As String
    If IsNull(V) Then
        SafeText = ""
    Else
        SafeText = CStr(V)
    End If
End Function
Private Sub DataGrid_HeadClick(ByVal ColIndex As Integer)
    'On Error Resume Next
    'Sort
    On Error Resume Next
    Dim sortField As String
    Dim sortString As String

    sortField = DataGrid.Columns(ColIndex).Caption
    If InStr(MyDBGrdRs.Sort, "Asc") Then
        sortString = sortField & " Desc"
    Else
        sortString = sortField & " Asc"
    End If
    MyDBGrdRs.Sort = sortString
End Sub
Private Sub DataGrid_OnAddNew()
    'DataGrid.Columns(3).text = 0
    'DataGrid.Columns(4).text = Date
End Sub
Private Sub DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtFields(0).Text = DataGrid.Columns(1).Text
    txtFields(1).Text = DataGrid.Columns(2).Text
    Encriptado.Text = DataGrid.Columns(3).Text
    Texto.Text = AesDecryptString(Encriptado.Text, Password.Text)
End Sub
Private Sub Encripta_Click()
    If Password.Text = "" Then MsgBox "Por seguridad debes introducir un password para el cifrado; si lo dejas en blanco el cifrado será menos seguro.", vbOKOnly, "Introduce un password"
    Encriptado.Text = AesEncryptString(Texto.Text, Password.Text)
End Sub
Private Sub Desencripta_Click()
    If Password.Text = "" Then MsgBox "No has introducido un password, el rpoceso seguirá sin uno.", vbOKOnly, "Introduce el password"
    Texto.Text = AesDecryptString(Encriptado.Text, Password.Text)
    RTB.Text = AesDecryptString(Encriptado.Text, Password.Text)
End Sub
Private Sub CreateRecordDB_Click()
    On Error GoTo ErrorHandler:
    If Chazaro.State = 0 Then MsgBox "Base De Datos Cerrada", vbOKOnly, "La BD no ha sido abierta"

    If Password.Text = "" Then MsgBox "Por seguridad debes introducir un password para el descifrado.", vbOKOnly, "Introduce un password": Exit Sub
    If Encriptado.Text = "" Then MsgBox "Debes encriptar el texto", vbOKOnly, "Missing Data": Exit Sub
    If txtFields(0).Text = "" Then MsgBox "Debes poner una categoría", vbOKOnly, "Missing Data [CATEGORÍA]": Exit Sub
    If txtFields(1).Text = "" Then MsgBox "Debes poner un título", vbOKOnly, "Missing Data [TÍTULO]": Exit Sub
    'Create Record
    With Chazaro
        .Execute "INSERT INTO TopSecret (Category, Title, Content, Created, Status ) " _
                & "VALUES ('" & txtFields(0).Text & "' " _
                & ", '" & txtFields(1).Text & "'" _
                & ", '" & Encriptado.Text & "'" _
                & ", '" & Date & "'" _
                & ", '1')"


'        .Execute "INSERT INTO TopSecret (Category, Title, Content, Created, Status ) " _
                & "VALUES ('Categoría'" _
                & ", 'Título'" _
                & ", 'Contenido'" _
                & ", '" & Date & "'" _
                & ", '1')"

        '.Execute "INSERT INTO Agents (Agent, Codex, Created, Status ) " _
                & "VALUES ('LordMaster'" _
                & ", 'U2FsdGVkX1+SoWaZ0O0DxdylpXrG0o5wZK0a7A=='" _
                & ", '" & Date & "'" _
                & ", '1')"
    End With
    MyDBGrdRs.Requery
    SetupDg

    'BackupDB
ErrorHandler:
End Sub
Private Sub CreateTable_Click()

    Dim myAnswer As Integer
    myAnswer = MsgBox("SI EXISTEN, ESTA OPERACIÓN ELIMINARÁ LAS TABLAS ANTERIORES", vbOKCancel, "PRECAUCIÓN - ¡P E L I G R O !")

    If myAnswer = vbOK Then
        Dim UserName As String
        Dim Psswrd As String
        UserName = InputBox("Nombre del usuario", "Introduce un nombre de usuario", "James Bond")
        If UserName = "" Or UserName = "James Bond" Then MsgBox "¡Ash!", vbOKOnly, "Escribe lo que se te pide": Exit Sub
        Psswrd = InputBox("Password", "Introduce un password", "Abrakadabra")
        If Psswrd = "" Or Psswrd = "Abrakadabra" Then MsgBox "¡Ash!", vbOKOnly, "Escribe lo que se te pide": Exit Sub

        Psswrd = AesEncryptString(Psswrd, Password.Text)

        Main Password.Text 'Opens DB
        'Create Table
        With Chazaro
            On Error Resume Next
            .Execute "DROP TABLE TopSecret"
            
            If Err.Number <> 0 Then
                MsgBox "No se pudo eliminar la tabla 'TopSecret'. Puede que no exista. Crearemos una nueva", vbExclamation, "Error"
                Err.Clear
            Else
                MsgBox "Tabla 'TopSecret' eliminada correctamente.", vbInformation, "Éxito"
            End If
            On Error GoTo 0
            .Execute "CREATE TABLE TopSecret(" _
                    & "Id IDENTITY CONSTRAINT PK_UID PRIMARY KEY," _
                    & "Category TEXT(25) WITH COMPRESSION NOT NULL," _
                    & "Title TEXT(50) WITH COMPRESSION NOT NULL," _
                    & "Content MEMO WITH COMPRESSION NOT NULL," _
                    & "Created DATETIME NOT NULL," _
                    & "Status YESNO DEFAULT False)", , _
                      adCmdText Or adExecuteNoRecords

            On Error Resume Next
            .Execute "DROP TABLE Agents"
            
            If Err.Number <> 0 Then
                MsgBox "No se pudo eliminar la tabla 'Agents'. Puede que no exista. Crearemos una nueva", vbExclamation, "Error"
                Err.Clear
            Else
                MsgBox "Tabla 'TopSecret' eliminada correctamente.", vbInformation, "Éxito"
            End If
            On Error GoTo 0

            .Execute "CREATE TABLE Agents(" _
                    & "Id IDENTITY CONSTRAINT PK_UID PRIMARY KEY," _
                    & "Agent TEXT(25) WITH COMPRESSION NOT NULL," _
                    & "Codex TEXT(40) WITH COMPRESSION NOT NULL," _
                    & "Created DATETIME NOT NULL," _
                    & "Status YESNO DEFAULT False)", , _
                      adCmdText Or adExecuteNoRecords
    
            '.Execute "INSERT INTO Agents (Agent, Codex, Created, Status ) " _
                    & "VALUES ('LordMaster'" _
                    & ", 'U2FsdGVkX1+SoWaZ0O0DxdylpXrG0o5wZK0a7A=='" _
                    & ", '" & Date & "'" _
                    & ", '1')"
    
            .Execute "INSERT INTO Agents (Agent, Codex, Created, Status ) " _
                    & "VALUES ('" & UserName & "' " _
                    & ", '" & Psswrd & "'" _
                    & ", '" & Date & "'" _
                    & ", '1')"
        End With
        MsgBox "Las Tablas han sido Creadas."
    ElseIf myAnswer = vbCancel Then
        MsgBox "El proceso ha sido Cancelado."
    End If

End Sub
Private Sub CheckCredentials()
    On Error GoTo ErrorHandler
    If Password.Text = "" Then MsgBox "Por seguridad debes introducir un password para el descifrado.", vbOKOnly, "Introduce un password": Exit Sub
    Dim Psswrd As String
    Psswrd = InputBox("Password", "Introduce password de usuario", "Abrakadabra")
    If Psswrd = "" Or Psswrd = "Abrakadabra" Then MsgBox "¡Ash!", vbOKOnly, "Escribe lo que se te pide": Exit Sub
    Main Password.Text

    Dim DBPassword                  As String
    Dim UserDecriptedPassword       As String
    Dim UserPassword                As String

    'check for correct password
    Set MyADORs = New ADODB.Recordset
    MyADORs.Open "SELECT Agent, Codex " _
    & "FROM Agents ", Chazaro, adOpenStatic, adLockOptimistic
    DBPassword = MyADORs!Codex

    UserDecriptedPassword = AesDecryptString(DBPassword, Psswrd)
    Password.Text = UserDecriptedPassword

    If Password.Text = UserDecriptedPassword Then
        MyADORs.Close: Set MyADORs = Nothing
    Else
        MsgBox "For security reasons, this app is ending", vbOKOnly, "Hacked!"
        End
    End If

    Exit Sub

ErrorHandler:

MsgBox Err.Number & " | " & Err.Description, vbOKOnly, "ERROR"

End Sub
Private Sub Form_DblClick()
    CheckCredentials
End Sub

Private Sub ModifyFieldsDB()
Exit Sub
Stop
Chazaro.Execute "ALTER TABLE Agents ALTER COLUMN Agent TEXT(25) NOT NULL"
Chazaro.Execute "ALTER TABLE TopSecret ALTER COLUMN Category text(25) NOT NULL"
Chazaro.Execute "ALTER TABLE TopSecret ALTER COLUMN Title text(50) NOT NULL"
'conn.Execute "ALTER TABLE TopSecret ALTER COLUMN Category DECIMAL(5,2) NOT NULL"

End Sub

Private Sub Form_Resize()
    'To resize form
    Dim ThisControl As Control, HorizRatio As Single, VertRatio As Single, Iter As Integer

    If Me.WindowState = vbMinimized Then Exit Sub

    HorizRatio = Me.Width / ControlInfos(0).Width
    VertRatio = Me.Height / ControlInfos(0).Height

    Iter = 0
    For Each ThisControl In Me.Controls
      Iter = Iter + 1
      On Error Resume Next  ' hack to bypass controls
      With ThisControl
        .Left = ControlInfos(Iter).Left * HorizRatio
        .Top = ControlInfos(Iter).Top * VertRatio
        .Width = ControlInfos(Iter).Width * HorizRatio
        .Height = ControlInfos(Iter).Height * VertRatio
        .FontSize = ControlInfos(Iter).FontSize * HorizRatio
      End With
      On Error GoTo 0
    Next
    'To resize form
End Sub

Private Sub lblPassword_DblClick()
    If Password.PasswordChar = "" Then Password.PasswordChar = "*": Exit Sub
    If Password.PasswordChar = "*" Then Password.PasswordChar = "": Exit Sub
End Sub


Private Sub Password_DblClick()
    If Chazaro.State = 0 Then Main Password.Text
End Sub
Private Sub ReadDB_Click()
    On Error GoTo ErrorHandler:
    If Password.Text = "" Then MsgBox "Por seguridad debes introducir un password para el descifrado.", vbOKOnly, "Introduce un password": Exit Sub
    If Chazaro.State = 0 Then MsgBox "Base De Datos Cerrada", vbOKOnly, "La BD no ha sido abierta"

    Screen.MousePointer = vbHourglass
    Set MyDBGrdRs = New ADODB.Recordset

    'Open Table
    With Chazaro
        MyDBGrdRs.Open "SELECT * FROM TopSecret", Chazaro, adOpenStatic, adLockOptimistic
        Set DataGrid.DataSource = MyDBGrdRs
        DataGrid.MarqueeStyle = dbgHighlightRow
        DataGrid.AllowUpdate = True
        DataGrid.AllowArrows = True
        DataGrid.AllowAddNew = True
        'DataGrid.Columns(3).Locked = True
        'DataGrid.Columns(4).Locked = True
    End With
    SetupDg
    Texto.Text = AesDecryptString(Encriptado.Text, Password.Text)
    RTB.Text = AesDecryptString(Encriptado.Text, Password.Text)
ErrorHandler:
    Screen.MousePointer = vbDefault
End Sub
Private Sub CreateDB_Click()
    If Password.Text = "" Then MsgBox "Por seguridad debes introducir un password para el descifrado.", vbOKOnly, "Introduce un password": Exit Sub

    Dim myAnswer As Integer
    Dim DBPath As String
    DBPath = App.Path & "\Kripto.accdb"
    myAnswer = MsgBox("SI EXISTE, ESTA OPERACIÓN ELIMINARÁ LA BASE DE DATOS ANTERIOR", vbOKCancel, "PRECAUCIÓN - ¡P E L I G R O !")
    If myAnswer = vbOK Then
        On Error Resume Next
        Kill DBPath
        If Err.Number <> 0 Then
            MsgBox "No se pudo eliminar el archivo. Puede que no exista o esté en uso. Crearemos uno nuevo", vbExclamation, "Error"
            Err.Clear
        Else
            MsgBox "Archivo eliminado correctamente.", vbInformation, "Éxito"
        End If
        On Error GoTo 0
    End If

    If myAnswer = vbOK Then
        Dim catDB As Object
        Set catDB = CreateObject("ADOX.Catalog")

        'createDB
        With catDB
            '.Create "Provider=Microsoft.ACE.OLEDB.12.0;" _
                  & "Data Source= " & App.Path & "\Kripto.accdb" & ";" _
                  & "Jet OLEDB:Database Password = 'ErgethKoglon';"
            .Create "Provider=Microsoft.ACE.OLEDB.12.0;" _
                  & "Data Source= " & App.Path & "\Kripto.accdb" & ";" _
                  & "Jet OLEDB:Database Password = '" & Password.Text & "';"
        End With

        Set catDB = Nothing
        MsgBox "La BD ha sido Creada."
    ElseIf myAnswer = vbCancel Then
        MsgBox "El proceso ha sido Cancelado."
    End If

End Sub
Private Sub Form_Load()
    'To resize form
    Dim ThisControl As Control
    ReDim Preserve ControlInfos(0 To 0)
    ControlInfos(0).Width = Me.Width
    ControlInfos(0).Height = Me.Height
    For Each ThisControl In Me.Controls
      ReDim Preserve ControlInfos(0 To UBound(ControlInfos) + 1)
      On Error Resume Next  ' hack to bypass controls with no size or position properties
      With ControlInfos(UBound(ControlInfos))
        .Left = ThisControl.Left
        .Top = ThisControl.Top
        .Width = ThisControl.Width
        .Height = ThisControl.Height
        .FontSize = ThisControl.FontSize
      End With
      On Error GoTo 0
    Next
    'To resize form


'Sample usage

'Just copy/paste mdAesCtr.bas from Source code section below to your project and will be able to strongly encrypt a user-supplied text with a custom password by calling AesEncryptString like this

''encrypted = AesEncryptString(userText, password)

'To decrypt the original text use AesDecryptString function with the same password like this

''origText = AesDecryptString(encrypted, password)

'These functions use sane defaults for salt and cipher strength that you don't have to worry about. These also encode/expect the string in encrypted in base-64 format so it can be persisted/mailed/transported as a simple string.

'Advanced usage

'Both string functions above use AesCryptArray worker function to encrypt/decrypt UTF-8 byte-arrays of the original strings. You can directly call AesCryptArray if you need to process binary data or need to customize AES salt and/or AES key length (strength) parameters.

'Function AesCryptArray also allows calculating detached HMAC-SHA256 on the input/output data ("detached" means hashes has to be stored separately, supports both encrypt-then-MAC and MAC-then-encrypt) when used like this

''AesCryptArray baEncr, ToUtf8Array("pass"), Hmac:=baHmacEncr

'(See More samples section below)

'Stream usage

'When contents to be encrypted does not fit in (32-bit) memory you can expose private pvCryptoAesCtrInit/Terminate/Crypt functions so these can be used to implement read/process/write loop on paged original content.

'implementation

'This implementation used to be based on WinZip AES-encrypted archives as implemented in ZipArchive project but now is compatible with openssl enc command when using aes-256-ctr cipher.
    lastIndex = 0
End Sub
Private Sub load_Click()
    Dim iFile As Integer
    Encriptado.Text = ""
    Texto.Text = ""

    iFile = FreeFile
    'FileName = InputBox("Nombre del archivo", "Introduce un nombre de archivo", File)
    cmdlgAgenda.Filter = "Cryptic Files|*.Krypt"
    cmdlgAgenda.ShowOpen
    If cmdlgAgenda.FileName <> "" Then
        Open cmdlgAgenda.FileName For Input As #iFile
        Encriptado.Text = Input(LOF(iFile), iFile)
        Close #iFile
    Else
        MsgBox "No seleccionaste archivo", vbOKOnly, "Faltan Datos"
    End If
End Sub
Private Sub save_Click()
    Dim iFile As Integer
    Dim FileName As String
    iFile = FreeFile
'    FileName = InputBox("Nombre del archivo", "Introduce un nombre de archivo", "Top Secret")
'    If FileName = "" Then FileName = "Top Secret"
'    Open Dir & "\" & FileName & ".Krypt" For Output As #iFile
    cmdlgAgenda.ShowSave
    Open cmdlgAgenda.FileName & ".Krypt" For Output As #iFile
    Print #iFile, Encriptado.Text
    Close #iFile
End Sub

Private Sub Texto_DblClick()
    With Texto
        .SelStart = 0
        .SelLength = Len(.Text)
   End With
End Sub

Private Sub UpdateRecordDB_Click()
    If Password.Text = "" Then MsgBox "Por seguridad debes introducir un password para el descifrado.", vbOKOnly, "Introduce un password": Exit Sub
    If Encriptado.Text = "" Then MsgBox "Debes encriptar el texto", vbOKOnly, "Missing Data": Exit Sub
    If txtFields(0).Text = "" Then MsgBox "Debes seleccionar el registro a modificar", vbOKOnly, "Missing Data [ID]": Exit Sub
    If Chazaro.State = 0 Then MsgBox "Base De Datos Cerrada", vbOKOnly, "La BD no ha sido abierta"

    'Update Table
    With Chazaro
        .Execute "UPDATE TopSecret " _
                & "SET Category = '" & txtFields(0).Text & "' " _
                & ", Title = '" & txtFields(1).Text & "'" _
                & ", Content = '" & Encriptado.Text & "'" _
                & ", Created = '" & Date & "'" _
                & ", Status = True " _
                & "WHERE ID = " & DataGrid.Columns(0).Text & " "
    End With
    MyDBGrdRs.Requery
    SetupDg
    'BackupDB
End Sub
