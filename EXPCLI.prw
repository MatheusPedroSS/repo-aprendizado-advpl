#include 'protheus.ch'
#include 'parmtype.ch'
#include 'TopConn.ch'

/*/{Protheus.doc}
    Fonte com o objetivo de exportar um simples relatorio de cliente para o excel.
    @author Matheus Pedro
    @since 12/02/2021
    @version undefined
    /*/
User Function EXPCLI()
    
Private cCliente1 := Space(6)
Private cCliente2 := Space(6)

getParam()

FWMsgRun(, { |oSay| Processa() }, "Buscando informações de Clientes", "Gerando Excel")

Return

/*/{Protheus.doc} Processa
    @type  Static Function
    @author Matheus Pedro
    @since 12/02/2021
    @version undefined
    /*/
Static Function Processa()

Local cQuery    := ""
Local cAlias    := GetNextAlias()
Local aDados    := {}

Sleep(5000)

cQuery  := "SELECT A1_COD, A1_LOJA, A1_NOME, A1_END, A1_MUN, A1_EST, A1_DTCAD   " + CRLF
cQuery  += "FROM " + RetSQLName("SA1") + " SA1                                  " + CRLF
cQuery  += "WHERE SA1.D_E_L_E_T_ = ''                                            " + CRLF
cQuery  += "AND A1_COD BETWEEN '"+ cCliente1 + "' AND '" + cCliente2 + "'       " + CRLF

TCQUERY cQuery NEW ALIAS (cAlias)

(cAlias)->(DbGoTop())

Do while !(cAlias)->(Eof())
    aAdd(aDados, {(cAlias)->A1_COD,;
        (cAlias)->A1_LOJA,;
        (cAlias)->A1_NOME,;
        (cAlias)->A1_END,;
        (cAlias)->A1_MUN,;
        (cAlias)->A1_EST,;
        (cAlias)->A1_DTCAD,;
        .F.})

    (cAlias)->(DbSkip())

endDo

(cAlias)->(DbCloseArea())

geraexcel(aDados)

Return 


Static Function geraexcel(aDados)

    Local oExcel    := FWMSExcel():New()
    Local oExcelApp := Nil
    Local cAba      := "Cadastro de Clientes"
    Local cTabela   := "Cadastro de Clientes"
    Local cArquivo  := "Cadastro de Clientes" + dToS( msDate() ) + "_" + strtran(time(), ":", "") + ".XLS"
    Local cPath     := "C:\TEMP\"
    Local cDefPath  := GetSrvProfString( "StartPath", "\system\")
    Local i

    if len(aDados) >0
        if !ApOleClient("MSExcel")
            MsgAlert("Microsoft Excel não instalado!")
            Return
        Endif

        oExcel:AddWorkSheet(cAba)
        oExcel:AddTable(cAba, cTabela)

        oExcel:AddColumn(cAba, cTabela, "CÓDIGO"            , 1, 1, .F.) 
        oExcel:AddColumn(cAba, cTabela, "LOJA"              , 1, 1, .F.) 
        oExcel:AddColumn(cAba, cTabela, "NOME"              , 1, 1, .F.) 
        oExcel:AddColumn(cAba, cTabela, "ENDEREÇO"          , 1, 1, .F.) 
        oExcel:AddColumn(cAba, cTabela, "MUNICIPIO"         , 1, 1, .F.) 
        oExcel:AddColumn(cAba, cTabela, "ESTADO"            , 1, 1, .F.) 
        oExcel:AddColumn(cAba, cTabela, "DATA DO CADASTRO"  , 1, 1, .F.) 

        For i := 1 to len(aDados)

            oExcel:AddRow(cAba,;
                cTabela,;
                {aDados[i][1]    ,;
                    aDados[i][2]    ,;
                    aDados[i][3]    ,;
                    aDados[i][4]    ,;
                    aDados[i][5]    ,;
                    aDados[i][6]    ,;
                    aDados[i][7]    })

        Next i

        if !Empty(oExcel:aWorkSheet)

            oExcel:Activate()
            oExcel:GetXMLFile(cArquivo)

            CpyS2T(cDefPath+cArquivo, cPath)

            oExcelApp   := MsExcel():New()
            oExcelApp:WorkBooks:Open(cPath+cArquivo)
            oExcelApp:SetVisible(.T.)
        Endif

    Else
        Alert("Não existem dados para emitir o relatório.")
    Endif

Return

/*/{Protheus.doc} getParam
    @type  Static Function
    @author Matheus Pedro
    @since 12/02/2021
    @version undefined
    /*/
Static Function getParam(param_name)

Local alParamBox := {}
Local clTitulo   := "Parametros"
Local alButtons  := {}
Local llCentered := .T.
Local nlPosx     := Nil
Local nlPosy     := Nil
Local clLoad     := ""
Local llCanSave  := .T.
Local llUserSave := .T.
Local llRet      := .T.
Local blOk       
Local alParams   := {}

AADD(alParamBox, {1, "Cliente de: "     ,Space(6)   ,"@!",".T."  ,"SA1"  ,"",25,.F.})
AADD(alParamBox, {1, "Cliente ate: "     ,Space(6)   ,"@!",".T."  ,"SA1"  ,"",25,.F.})

llRet := ParamBox(alParamBox, clTitulo, alParams, blOk, alButtons, llCentered, nlPosx, nlPosy,, clLoad, llCanSave, llUserSave)

if ( llRet )
    cCliente1   := alParams[1]
    cCliente2   := alParams[2]
Endif

Return 
