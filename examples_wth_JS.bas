Attribute VB_Name = "examples_wth_JS"
Option Explicit

'I used to program more automations through the javascript / jquery of pages than
'using the same elements inside the vba, since it was easier to use the logs window and dev
'within browsers than programming directly in vba,
'for example, when using .ExecScript, it was possible to use commands in Jquery,
'If I'm speaking Greek to you, ignore probably not the time for you to know it!


'The same example in module examples_1, but different, in comment is how it was done
'without the javascript, and how it was done using javascript

Public Sub goToGoogleAndSeachWikiPediaWithJS()
    Dim ie As New ieRV
    With ie
        Call .initIE
        Call .NavigateTo("www.google.com.br")
        Call .waitElem("document.getElementsByName('q').item(0)", ".innerText = 'Wikipedia'", 20)
        'Call ie.getElement(20, "tagname", "input", "title", "pesquisar").setAttribute("innerText", "Wikipedia")
        Call .waitElem("document.getElementsByName('btnK').item(0)", ".click()", 20)
        'Call ie.getElement(20, "tagname", "input", "value", "*pesquisa*", "parentNode.tagname", "CENTER").Click
        Stop
    End With
End Sub


Public Sub AcceptOneAlert()
    'LEIA OS COMENTARIOS ANTES DE DAR f5 Cabeçudo
    Dim ie As New ieRV
    With ie
        Call .initIE(noAddOns:=True)
        Call .NavigateTo("about:blank")
        'Show a popup
        'Call .execScript("setTimeout(""alert('ESSE É UM ERRO QUE O RONAN VICO CRIOU!');"", 1)")
        Call .ExecScriptAssync("alert('ESSE É UM ERRO QUE O RONAN VICO CRIOU!')", 1)
        Call .wait(2000)
        'This line accepts the popup, throwing an error with the message contained in the popup.
        'You can also Send False and not receive the error, just accepting the alert eg: Call .acceptAlert (false)
        'Error only comes up when you have the alert, the error is 12345 and can be handled.
        Call .acceptAlert(True)
    End With
    Stop
End Sub
