VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Auto Entregador"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4740
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Versão 1.26122022"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public verva As New ADODB.Recordset
 Public usando As New ADODB.Recordset
Public usando1 As New ADODB.Recordset
Public gravamensaN1 As New ADODB.Recordset
Public buscamensaN1 As New ADODB.Recordset
   Public tipop As New ADODB.Recordset
   Public formap As New ADODB.Recordset
Public cnn8 As New ADODB.Connection
Public cnn As New ADODB.Connection
Public qualdaslinhas As String
Public ondeta As String
Public qualnroid As String
Public contadorpassagem As Double
Public desligar As String
Public diario As String
Public mensal As String
Public contadordiario As Double
Public contadormensal As Double
Public geral As Double

Public esconder As Integer


Public notempo As Date
Public qualnoroid As Integer




Private Sub Command3_Click()
 
End Sub

Public Sub cnn8_abrir()

On Error GoTo abertura
c_cnn8 = FreeFile
Open "c:\Geral\c_cnn8.txt" For Input As c_cnn8
Line Input #c_cnn8, Linha  'caminho
Line Input #c_cnn8, linha1  ' porta
Line Input #c_cnn8, linha2  ' nomebanco
Line Input #c_cnn8, linha3  ' usuario
Line Input #c_cnn8, linha4  ' senha
Close #c_cnn8
On Error GoTo abertura1
If cnn8.State = 0 Then
'abreinterno = "banco1617.mysql.uhserver.com"
abreinterno = Linha
'Porta = "" & "3306"
Porta = "" & linha1
'meubanco = "banco1617"
meubanco = linha2
'meuusuario = "a1617"
meuusuario = linha3
'minhasenha = "aires132+"
minhasenha = linha4
If UCase(linha4) = "PADRAO" Then
minhasenha = "pctrim"
End If
cnn8.ConnectionTimeout = 1 ' coloquei 1 mas o que mandam colocar é 5
cnn8.CommandTimeout = 1 ' coloquei 1 mas o que mandam colocar é 400
cnn8 = "Driver={MySQL ODBC 5.1 Driver};SERVER=" & abreinterno
cnn8 = cnn8 & ";port=" & Porta
cnn8 = cnn8 & "; DATABASE=" & meubanco
cnn8 = cnn8 & ";Uid=" & meuusuario
cnn8 = cnn8 & ";Pwd=" & minhasenha
cnn8.CursorLocation = adUseClient
cnn8.Open
End If
Exit Sub
abertura:
Close #c_cnn8
c_cnn8 = FreeFile
Open "c:\Geral\c_cnn8.txt" For Output As c_cnn8
'Print #c_cnn8, "banco1617.mysql.uhserver.com" 'caminho
Print #c_cnn8, "www.comdadoinf.ddns.com.br" 'caminho
Print #c_cnn8, "3308"  ' porta
'Print #c_cnn8, "banco1617"  ' nomebanco
Print #c_cnn8, "suportecliente"  ' nomebanco
Print #c_cnn8, "root"  ' usuario
Print #c_cnn8, "padrao"  ' senha
Close #c_cnn8
Exit Sub
abertura1:

End Sub

 
Public Sub lerservidortxt()
sup1 = FreeFile
Open "C:\geral\servidor1.Txt" For Input As sup1
Line Input #sup1, Linha0
Line Input #sup1, linha1
Line Input #sup1, linha2
Line Input #sup1, linha3
Line Input #sup1, linha4
Line Input #sup1, linha5
Line Input #sup1, linha6
Line Input #sup1, linha7
Line Input #sup1, linha8
Line Input #sup1, linha9
Line Input #sup1, linha10
Line Input #sup1, linha11
Line Input #sup1, Linha12
'Line Input #sup1, linha13

Close #nf2

comprimento = Len(Linha0)
contacomprimento = 1
Do While contacomprimento <= comprimento
qualcara = Mid(Linha0, contacomprimento, 1)
If qualcara <> "=" Then
ondeta = ondeta & qualcara
Else
contacomprimento = 2000
End If
contacomprimento = contacomprimento + 1
Loop


comprimento = Len(Linha12)
contacomprimento = 1
Do While contacomprimento <= comprimento
qualcara = Mid(Linha12, contacomprimento, 1)
If qualcara <> "=" Then
qualnroid = qualnroid & qualcara
Else
contacomprimento = 2000
End If
contacomprimento = contacomprimento + 1
Loop

ondeta = Trim(ondeta)

qualnroid = Trim(qualnroid)


End Sub

 

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Label3 = ""
If Text1 = "" Then
Text1 = "56545566"
End If
If usando.State <> 0 Then
usando.Close
End If
 
usando.Open "Select * From deliverypendente where comanda=" & Text1 & "", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
If usando.EOF = False Then
nahorade = Date & " " & Time
horadodespacho = usando!horadodespacho
lalao = DateDiff("n", horadodespacho, nahorade)
If lalao <= 15 Then
Text1 = ""
Label3 = lalao
Exit Sub
End If
Else
Text1 = ""
Label3 = "Inexistente"
Exit Sub
End If


teminadoprocesso = Format(Date, "yyyy-mm-dd") & " " & Time
Dinheiro = 0
CartaDebito = 0
CartaoCredito = 0
RecFiado = 0
Ticket = 0
temdados = ""
If usando.State <> 0 Then
usando.Close
End If
usando.Open "Select * From contador", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
relativo = usando!relativoaodia
If Trim(relativo) = "" Then
relativo = Date
usando!relativoaodia = relativo
usando.Update
End If
usando.Close
If usando.State <> 0 Then
usando.Close
End If
 
usando.Open "Select * From deliverypendente where comanda=" & Text1 & "", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
Do While usando.EOF = False
temdados = "Sim"
comanda1 = usando!comanda
telefone1 = usando!telefone
cliente1 = usando!cliente
endereco1 = usando!Endereco
bairro1 = usando!Bairro
cep1 = usando!cep
observacao1 = usando!observacao
nomeproduto1 = usando!nomeproduto
nroproduto1 = usando!nroproduto
codigoproduto1 = usando!codigoproduto
pediutroco = usando!pediutroco
comopaga = Trim(usando!FormaPgto)
 totaldacomanda = usando!totaldacomanda
VALORPRODUTO1 = usando!valorproduto
pediucontravale1 = usando!pediucontravale
classificacao1 = usando!CLASSIFICACAO
entregador1 = usando!entregador
horadopedido = usando!horadopedido
horadodespacho = usando!horadodespacho


If codigoproduto1 = "TXENTREGA" Then
taxaentregador = VALORPRODUTO1
End If
pagarimposto1 = usando!pagarimposto
If usando1.State <> 0 Then
usando1.Close
End If

If tipop.State <> 0 Then
tipop.Close
End If
tipop.Open "Select * From tipospagamento WHERE MOEDA =""" & comopaga & """", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
If tipop.EOF = False Then
comopaga = tipop!tratar_como
Else
comopaga = "DINHEIRO"
End If

 
 
usando1.Open "Select * From comandas", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
usando1.AddNew
usando1!comandanro = comanda1
usando1!telefone = telefone1
usando1!cliente = cliente1
usando1!Endereco = endereco1
usando1!quantos = nroproduto1
usando1!produtocodigo = codigoproduto1
usando1!mercadoria = nomeproduto1
usando1!valor = VALORPRODUTO1
usando1!totaldacomanda = totaldacomanda
usando1!ficacao = classificacao1
usando1!entregador = entregador1
usando1!apagarimposto = pagarimposto1
usando1!obs = "SIM"
usando1!imprecontravale = pediucontravale1
usando1!hpedidofeito = horadopedido
usando1!hpedidodespachado = horadodespacho

usando1!Data = Format(relativo, "yyyy-mm-dd")
If Trim(comopaga) = "DINHEIRO" Then
usando1!Dinheiro = totaldacomanda
Else
usando1!Dinheiro = 0
End If
If Trim(comopaga) = "CARTAO DEBITO" Then
usando1!cartaodebito = totaldacomanda
Else
usando1!cartaodebito = 0
End If
If Trim(comopaga) = "CARTAO CREDITO" Then
usando1!CartaoCredito = totaldacomanda
Else
usando1!CartaoCredito = 0
End If
If Trim(comopaga) = "VOUCHER" Then
usando1!Ticket = totaldacomanda
Else
usando1!Ticket = 0
End If
usando1!caixinha = 0
usando1!desconto = 0
'If comopaga = "RECFUTURO" Then
usando1!recfuturo = 0
'End If
usando1!cheque = 0
usando1!Fechamento = 0
usando1!contravale = 0
usando1.Update
usando1.Close
usando.MoveNext
Loop

If comopaga = "DINHEIRO" Then
Dinheiro = totaldacomanda
Else
Dinheiro = 0
End If

If comopaga = "CHEQUE" Then
cheque = totaldacomanda
Else
cheque = 0
End If

If comopaga = "CARTAO DEBITO" Then
cartaodebito = totaldacomanda
Else
cartaodebito = 0
End If

If comopaga = "CARTAO CREDITO" Then
CartaoCredito = totaldacomanda
Else
CartaoCredito = 0
End If

If comopaga = "VOUCHER" Then
Ticket = totaldacomanda
Else
Ticket = 0
End If

contravalerecebido = 0
recfuturo = 0
desconto = 0
contravaleimpresso = 0
If totaldacomanda <> Empty Then
If formap.State <> 0 Then
formap.Close
End If
   formap.Open "insert into formapagamento(Dinheiro,Cheque,CartaoDebito,CartaoCredito,Ticket,TotalComanda,NroComanda,usuario) values (" & Replace(Dinheiro, ",", ".") & "," & Replace(cheque, ",", ".") & "," & Replace(cartaodebito, ",", ".") & "," & Replace(CartaoCredito, ",", ".") & "," & Replace(Ticket, ",", ".") & "," & Replace(totaldacomanda, ",", ".") & "," & Text1 & ",'AUTOMATICO')", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
End If
   'formap.Open "insert into formapagamento(Dinheiro,cheque,cartaodebito,cartaocredito,ticket,contravalerecebido,recfuturo,desconto,contravaleimpresso,nrocomanda,totalcomanda,usuario) values (" & dinheiro & "," & cheque & ", " & cartaodebito & "," & cartaocredito & "," & ticket & ",0,0,0,0," & Replace(Text1, ",", ".") & "," & Replace(totaldacomanda, ",", ".") & ",'MOTOQUEIRO_AUTO')", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
If temdados = "Sim" Then
fimfim = Format(Date, "yyyy-mm-dd") & " " & Time
If usando.State <> 0 Then
usando.Close
End If
usando.Open "insert automotoqueiro (nrocomanda,entregador,finalizado,valorcomanda,valortroco,valortaxa) values (" & comanda1 & ",""" & entregador1 & """,""" & fimfim & """," & Replace(totaldacomanda, ",", ".") & "," & Replace(pediutroco, ",", ".") & "," & Replace(taxaentregador, ",", ".") & ")", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
End If
'usando.Open "insert  From automotoqueiro (nrocomanda,entregador,finalizado) values (" & comanda1 & ",""" & entregador1 & """,""" & fimfim & """)", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
If usando.State <> 0 Then
usando.Close
End If
usando.Open "delete  From deliverypendente where comanda=" & Text1 & "", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
If usando.State <> 0 Then
usando.Close
End If
usando.Open "update statuspedidos set situacao ='PEDIDO FINALIZADO',finalizado=""" & teminadoprocesso & """  where comandanro=" & Text1 & "", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
Text1 = ""
Text1.SetFocus
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
ondeta = ""
Call lerservidortxt
Label2 = ondeta
DoEvents
If cnn.State <> 0 Then
cnn.Close
End If
cnn = "Driver={MySQL ODBC 5.1 Driver};SERVER=" & ondeta
cnn = cnn & ";port=" & "3308"
cnn = cnn & "; DATABASE=" & "modulomesa" & qualnroid
cnn = cnn & ";Uid=" & "root"
cnn = cnn & ";Pwd=" & "pctrim"
cnn.Open

Timer1.Enabled = False
Text1.Locked = False
End Sub
