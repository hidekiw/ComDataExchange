VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cliente VB6 - Comunicação COM"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3720
      Width           =   7935
   End
   Begin VB.CommandButton btnProcessData 
      Caption         =   "Processar Dados"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Text            =   "dados para processar"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calcular Soma"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtNumber2 
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Text            =   "25"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtNumber1 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "10"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton btnSetMessage 
      Caption         =   "Enviar para C#"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "Mensagem do VB6"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton btnGetMessage 
      Caption         =   "Receber do C#"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Resultados:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Dados para processar:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Números para somar:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Mensagem para enviar:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declaração do objeto COM
Private comObj As Object

Private Sub Form_Load()
    ' Inicializar o componente COM do C#
    On Error GoTo ErrorHandler
    
    Set comObj = CreateObject("MyApp.DataExchange")
    txtResult.Text = "Componente COM inicializado com sucesso!" & vbCrLf
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao inicializar componente COM: " & Err.Description, vbCritical
    txtResult.Text = "ERRO: Componente COM não encontrado. Certifique-se de que o componente C# foi registrado."
End Sub

Private Sub btnGetMessage_Click()
    On Error GoTo ErrorHandler
    
    If comObj Is Nothing Then
        MsgBox "Componente COM não inicializado!", vbExclamation
        Exit Sub
    End If
    
    Dim message As String
    message = comObj.GetMessage()
    
    txtResult.Text = txtResult.Text & "Mensagem recebida do C#: " & message & vbCrLf
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao receber mensagem: " & Err.Description, vbCritical
End Sub

Private Sub btnSetMessage_Click()
    On Error GoTo ErrorHandler
    
    If comObj Is Nothing Then
        MsgBox "Componente COM não inicializado!", vbExclamation
        Exit Sub
    End If
    
    comObj.SetMessage txtMessage.Text
    txtResult.Text = txtResult.Text & "Mensagem enviada para C#: " & txtMessage.Text & vbCrLf
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao enviar mensagem: " & Err.Description, vbCritical
End Sub

Private Sub btnCalculate_Click()
    On Error GoTo ErrorHandler
    
    If comObj Is Nothing Then
        MsgBox "Componente COM não inicializado!", vbExclamation
        Exit Sub
    End If
    
    Dim num1 As Integer, num2 As Integer, result As Integer
    num1 = CInt(txtNumber1.Text)
    num2 = CInt(txtNumber2.Text)
    
    result = comObj.CalculateSum(num1, num2)
    
    txtResult.Text = txtResult.Text & "Cálculo: " & num1 & " + " & num2 & " = " & result & vbCrLf
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao calcular: " & Err.Description, vbCritical
End Sub

Private Sub btnProcessData_Click()
    On Error GoTo ErrorHandler
    
    If comObj Is Nothing Then
        MsgBox "Componente COM não inicializado!", vbExclamation
        Exit Sub
    End If
    
    Dim processed As String
    processed = comObj.ProcessData(txtInput.Text)
    
    txtResult.Text = txtResult.Text & "Dados processados: " & processed & vbCrLf
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao processar dados: " & Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Limpar referência do objeto COM
    Set comObj = Nothing
End Sub