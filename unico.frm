VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assistente de Backup"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6165
   Icon            =   "unico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      Picture         =   "unico.frx":5719A
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame frameAjuste 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ajustes do Destino"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   4095
      Begin VB.CheckBox checkFull 
         BackColor       =   &H80000005&
         Caption         =   "Cópia Full ( Completo )"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtOrigem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "c:\jean\marcelo"
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtDestino 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\BACKUP\"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Origem:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Destino:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   6255
      Begin VB.CommandButton btnExecutar 
         Caption         =   "&Executar"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton btnSair 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label labelStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Insira o caminho em que deseja fazer o BACKUP"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label labelA 
      BackStyle       =   0  'Transparent
      Caption         =   "Assistente de Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ScriptControl1_Error()

End Sub

'Wellisson 30-11-2016
Private Sub btnExecutar_Click()
    frameAjuste = True
    labelA.Caption = "Aguarde..."
    labelStatus.Caption = "Copiando os arquivos..."
    labelA.Refresh
    labelStatus.Refresh
    Screen.MousePointer = vbHourglass
    
    Dim FSO As New FileSystemObject
        If checkFull.Value = 0 Then
            If Not FSO.FolderExists(txtDestino.Text) = True Then
                FSO.CreateFolder txtDestino.Text
            End If
         
            If FSO.FolderExists(txtOrigem.Text) = True Then
                FSO.CopyFile txtOrigem.Text & "\bdmarc.mdb", txtDestino.Text, True
                FSO.CopyFile txtOrigem.Text & "\bdconfig.mdb", txtDestino.Text, True
                FSO.CopyFile txtOrigem.Text & "\bdfiscal.mdb", txtDestino.Text, True
                MsgBox "Backup Concluído com êxito!", vbInformation, "Backup do Sistema"
            Else
                MsgBox "Nenhum arquivo encontra, por favor verifique!", vbInformation, "Backup do Sistema"
            End If
        Else
            If FSO.FolderExists(txtOrigem.Text) = True Then
                FSO.CopyFolder txtOrigem.Text, txtDestino.Text & "FULL-BACKUP-" & Format(Date, "DD-MM-YYYY"), True
                MsgBox "Backup realizado com sucesso!", vbInformation, "Backup do sistema"
            Else
                MsgBox "O diretório não foi encontrado, por favor verifique!", vbInformation, "Backup do sistema"
            End If
        End If
    End
erro:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Backup do Sistema"
End Sub
Private Sub btnSair_Click()
    End
End Sub
