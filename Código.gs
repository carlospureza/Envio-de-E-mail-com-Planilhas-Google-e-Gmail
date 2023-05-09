function onOpen(){

    A1.setValue("Matrícula");
    B1.setValue("Nome");
    C1.setValue("Status");
    D1.setValue("E-mail");
    
    var ui = SpreadsheetApp
    .getUi()
    .createMenu("Enviar E-mails")
    .addItem("Enviar", "enviarEmail")
    .addToUi();
    
    var aviso = SpreadsheetApp
    .getUi()
    .alert("Importe a planilha conforme o modelo.");
    
    
    }