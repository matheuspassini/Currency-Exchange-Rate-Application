VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCotacaoDolar 
   Caption         =   "UserForm1"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7740
   OleObjectBlob   =   "dolar3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCotacaoDolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dolarAtual As String

Private Sub btmAtualizaDolar_Click()


    Dim site, html, resumoHTML As String
    Dim requisicao As Object
    
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    
    site = "https://dolarhoje.com/"
    requisicao.Open "GET", site, False
    requisicao.send
    html = requisicao.ResponseText
    
    resumoHTML = Mid(html, InStr(html, "cotMoeda nacional"), 100)
    resumoHTML = Mid(resumoHTML, InStr(resumoHTML, "value="), 11)
    dolarAtual = Right(resumoHTML, 4)
    
    
    lbReal.Caption = dolarAtual
    TextBox1.Value = VBA.Format(1, "###,###,##0.00")
    
    Set requisicao = Nothing
    site = vbNullString
    html = vbNullString
    resumoHTML = vbNullString
    
End Sub


Private Sub TextBox1_Change()
    If TextBox1.TextLength > 0 And IsNumeric(TextBox1) Then
        lbReal.Caption = VBA.Format(TextBox1.Value * dolarAtual, "###,###,##0.00")
    Else
        lbReal.Caption = dolarAtual
    End If
    
End Sub

Private Sub UserForm_Initialize()
    Call btmAtualizaDolar_Click
End Sub
