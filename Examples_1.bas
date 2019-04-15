Attribute VB_Name = "Examples_1"
Option Explicit




Public Sub VaAoGooglePesquiseWikipedia()
    'O nome da sub ja diz o que o exemplo faz ! rsrsrs
    Dim ie As New ieRV
    With ie
        Debug.Print "Quantos ies abertos ? "; .howMuchIesIsOpened
        'inicia ie Invisivel
        Call .initIE(SW_HIDE, True, InternetExplorer)
        Debug.Print "Quantos ies abertos ? "; .howMuchIesIsOpened
        'navega , dighita wikipedia e pesquisa
        Call .NavigateTo("www.google.com.br", SW_HIDE)
        Call .getElement(20, "tagname", "input", "title", "pesquisar").setAttribute("innerText", "Wikipedia")
        Call .getElement(20, "tagname", "input", "value", "*pesquisa*", "parentNode.tagname", "CENTER").Click
        .ie.visible = True
    End With
End Sub


 
 
Public Sub TabelaParaRangeDoExcel()
    Const HTMLtabelaExemplo As String _
        = "<table>" & _
              "<tr> " & _
                "<th>Month</th>" & _
                "<th>Savings</th>" & _
              "</tr>" & _
              "<tr>" & _
               " <td>January</td>" & _
              "  <td>$100</td>" & _
             " </tr>" & _
             "<tr>" & _
               " <td>Feb</td>" & _
              "  <td>$400</td>" & _
             " </tr>" & _
            "</table>"
    Dim ie As New ieRV
    With ie
        Call .initIE
        Call .NavigateTo("")
        .ie.Document.body.innerHTML = HTMLtabelaExemplo
        Call .TableToRange(.getElement(5, "tagname", "table"), ThisWorkbook.Sheets(1).Cells(6, 6))
    End With


End Sub


Public Sub PropsIe()
    Dim ie As New ieRV
    With ie
        .initIE
        .NavigateTo ("about:blank")
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=False)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=True, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=True, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=True, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=True, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=True, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=True, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .bringIeToFront
        Call .setPropertiesIES
    End With
End Sub


Public Sub AlterIERegistrys()
    Dim ie As New ieRV
    '\/ take a look in this function nice ie registrys
    Call ie.RegistryIE
End Sub


Public Sub Using_The_Ie()
    Dim ie As New ieRV
    With ie
        .initIE
        .NavigateTo ("https://www.linkedin.com/in/ronan-vico/")
        .initIE SW_SHOWNORMAL, False
        .NavigateTo ("https://github.com/RonanVico/")
    End With
End Sub




