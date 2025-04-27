VERSION 5.00
Begin VB.Form frmConsultaCEP 
   Caption         =   "Consulta de CEP"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5010
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar CEP"
      Height          =   495
      Left            =   3300
      TabIndex        =   1
      Top             =   300
      Width           =   1455
   End
   Begin VB.TextBox txtCEP 
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   2895
   End
   Begin VB.Label lblRua 
      Caption         =   "Rua:"
      Height          =   375
      Left            =   300
      TabIndex        =   2
      Top             =   1000
      Width           =   4455
   End
   Begin VB.Label lblBairro 
      Caption         =   "Bairro:"
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   1500
      Width           =   4455
   End
   Begin VB.Label lblCidade 
      Caption         =   "Cidade:"
      Height          =   375
      Left            =   300
      TabIndex        =   4
      Top             =   2000
      Width           =   4455
   End
   Begin VB.Label lblEstado 
      Caption         =   "Estado:"
      Height          =   375
      Left            =   300
      TabIndex        =   5
      Top             =   2500
      Width           =   4455
   End
End
Attribute VB_Name = "frmConsultaCEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdConsultar_Click()
    Dim httpRequest As Object
    Dim jsonResponse As New hJsonBag
    Dim url As String
    Dim cleanedCEP As String
        
    ' Limpa o CEP
    cleanedCEP = Replace(txtCEP.Text, "-", "")
    cleanedCEP = Replace(cleanedCEP, ".", "")
    cleanedCEP = Trim(cleanedCEP)
    
    If Len(cleanedCEP) <> 8 Then
        MsgBox "Digite um CEP válido (8 dígitos)", vbExclamation
        Exit Sub
    End If
    
    ' Monta URL
    url = "https://brasilapi.com.br/api/cep/v1/" & cleanedCEP
    
    ' Faz a requisição
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.send
    
    If httpRequest.Status = 200 Then
        ' Lê o JSON
        jsonResponse.Parse ((httpRequest.responseText))
        
        ' Preenche os Labels
        lblRua.Caption = "Rua: " & jsonResponse("street")
        lblBairro.Caption = "Bairro: " & jsonResponse("neighborhood")
        lblCidade.Caption = "Cidade: " & jsonResponse("city")
        lblEstado.Caption = "Estado: " & jsonResponse("state")
        
    ElseIf httpRequest.Status = 404 Then
        MsgBox "CEP não encontrado.", vbInformation
    Else
        MsgBox "Erro ao consultar CEP: " & httpRequest.Status, vbCritical
    End If
    
    ' Libera objetos
    Set httpRequest = Nothing
    Set jsonResponse = Nothing
End Sub
