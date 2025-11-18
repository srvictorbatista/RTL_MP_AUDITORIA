@echo off & chcp 65001 >nul & COLOR 0A & setlocal enabledelayedexpansion

:: ####################################################################################################################################################################
:: TXT_TO_CSV.CMD  [RTL em Windows Batch]  Versão 0.0.1
:: Repositório original no GitHub: https://github.com/srvictorbatista/RTL_MP_AUDITORIA
:: ####################################################################################################################################################################
::
:: Este script foi desenvolvido para processar arquivos .txt exportados manualmente da página de movimentações do Mercado Livre
:: (https://www.mercadopago.com.br/activities/), onde o conteúdo é simplesmente copiado e colado em um arquivo de texto puro.
:: O objetivo é estruturar esses dados em um arquivo .csv devidamente formatado para análise e auditoria interna.
::
:: Funcionamento geral:
::   • Para cada arquivo *.txt presente no diretório atual, o script cria uma cópia correspondente em formato .csv.
::   • No começo do .csv, grava BOM UTF-8 para compatibilidade com Excel, seguida da linha de cabeçalho:
::     "MOV TIPO";"VALOR";"TRANSAÇÃO ID";"HORÁRIO";;;;"TOTAL VENDAS:";"=SOMA(B2:B2000)";;"COMISSÃO:";"=I1/100*5";;"FATURA:";"=I1-L1";
::   • Em seguida, analisa o conteúdo do .txt registrando os campos conforme encontrados na sequência do relatório:
::       1) A linha contendo o tipo da operação (ex.: "Venda") inicia um novo registro.
::       2) A presença de "R$" indica que a próxima linha numérica representa a parte inteira do valor.
::       3) Uma linha contendo apenas "," indica que a linha numérica seguinte é a parte fracionária (centavos).
::       4) Linhas contendo apenas dígitos são utilizadas como parte inteira ou fracionária do valor, conforme o estado atual.
::       5) A linha contendo "Transa" (case-insensitive) define o ID da transação; vírgulas residuais são removidas.
::       6) A linha que contém "h" (ex.: "14h40" ou "14h") especifica o horário e, ao ser encontrada, completa o registro e grava no .csv.
::   • Se o arquivo terminar antes que o horário seja encontrado, qualquer registro parcialmente montado ainda assim será gravado.
::
:: Normalização, tolerância e comportamentos especiais:
::   • Caso não exista parte fracionária, assume centavos "00".
::   • Caso apareçam sinais '+' ou '-' em linhas isoladas, são preservados.
::   • Espaços excedentes no início ou fim de cada linha são removidos automaticamente.
::   • Linhas vazias são ignoradas.
::
:: Observações importantes:
::   • Totalmente compatível com Windows CMD (cmd.exe).
::   • Mantém o comportamento lógico equivalente ao script Bash e ao original em BAT, corrigindo a montagem do campo VALOR.
::   • O arquivo CSV gerado utiliza separador ';' e todos os campos são envelopados por aspas, garantindo importação limpa no Excel.
::
:: Nota técnica:
::   • O campo TRANSAÇÃO recebe ao final um caractere NBSP (U+00A0), exatamente como a intenção original, forçando o Excel a interpretá-lo como texto.
::
:: Estrutura assumida nas linhas originais do relatório:
::   Tipo  →  "R$"  →  parte inteira  →  ","  →  parte fracionária  →  "Transação <id>"  →  "<hh>h<mm>"
::   Quando algum elemento estiver ausente, o parser aplica tolerância (ex.: sem fracionário → assume "00").
::
:: ####################################################################################################################################################################



rem Loop por todos os .txt na pasta
for %%F in (*.txt) do (

    set "outfile=%%~dpnF.csv"

    rem Escreve BOM UTF-8 e cabeçalho
    > "!outfile!" (
        <nul set /p "=﻿"
    )
    >> "!outfile!" echo "MOV TIPO";"VALOR";"TRANSAÇÃO ID";"HORÁRIO"; ;;;"TOTAL VENDAS:";"=SOMA(B2:B200000)"; ;"COMISSÃO:";"=I1/100*5"; ;"FATURA:";"=I1-L1";

    rem inicializa estado
    set "tipo="
    set "valint="
    set "valfrac="
    set "trx="
    set "hh="
    set "mm="
    set "inRecord=0"
    set "sawR=0"
    set "sawComma=0"
    set "sign="

    rem lê arquivo linha a linha
    for /f "usebackq delims=" %%L in ("%%F") do (
        set "ln=%%L"

        rem trim leading
        for /f "tokens=* delims= " %%x in ("!ln!") do set "ln=%%x"

        rem trim trailing (loop simples)
        for /l %%i in (1,1,200) do if defined ln if "!ln:~-1!"==" " set "ln=!ln:~0,-1!"

        if defined ln (

            rem detecta linha de tempo contendo "h"
            set "tmp=!ln:h=!"
            if not "!tmp!"=="!ln!" (

                rem parse hora e minuto
                for /f "tokens=1,2 delims=h" %%H in ("!ln!") do (
                    set "hh=%%H"
                    set "mm=%%I"
                )

                if not defined mm set "mm="
                if defined mm if "!mm:~1!"=="" set "mm=0!mm!"

                if not defined valint set "valint=0"
                if not defined valfrac set "valfrac=00"

                rem aplica sinal
                if defined sign if not "!sign!"=="" set "valint=!sign!!valint!"

                rem grava registro já com hora única
                call :write_row "!tipo!" "!valint!" "!valfrac!" "!trx!" "!hh!" "!mm!" "!outfile!"

                rem reset
                set "tipo="
                set "valint="
                set "valfrac="
                set "trx="
                set "hh="
                set "mm="
                set "inRecord=0"
                set "sawR=0"
                set "sawComma=0"
                set "sign="

            ) else (

                if "!inRecord!"=="0" (
                    rem início de novo registro
                    set "tipo=!ln!"
                    set "inRecord=1"
                    set "valint="
                    set "valfrac="
                    set "trx="
                    set "hh="
                    set "mm="
                    set "sawR=0"
                    set "sawComma=0"
                    set "sign="
                ) else (

                    rem dentro de registro
                    set "tmpline=!ln!"

                    echo !tmpline! | findstr /i "Transa" >nul 2>&1
                    if not errorlevel 1 (
                        rem captura o ÚLTIMO token da linha como ID da transação
                        rem aceita linhas como "Transação 133825690412" e "Transação de exemplo 133825690412"
                        set "last="
                        for %%T in (!ln!) do set "last=%%~T"
                        set "trx=!last!"
                        set "trx=!trx:,=!"
                    ) else (
                        echo !ln! | find "R$" >nul 2>&1
                        if not errorlevel 1 (
                            set "sawR=1"
                        ) else if "!ln!"=="," (
                            set "sawComma=1"
                        ) else if "!ln!"=="-" (
                            set "sign=-"
                        ) else if "!ln!"=="+" (
                            set "sign=+"
                        ) else (
                            rem tenta interpretar como número
                            set /a tryval=!ln! >nul 2>&1
                            if !errorlevel! EQU 0 (
                                if "!sawComma!"=="1" (
                                    if not defined valfrac set "valfrac=!ln!"
                                ) else (
                                    if not defined valint set "valint=!ln!"
                                )
                            )
                        )
                    )
                )
            )
        )
    )

    rem registro pendente
    if defined inRecord if defined valint (
        if not defined valfrac set "valfrac=00"
        if defined sign if not "!sign!"=="" set "valint=!sign!!valint!"
        call :write_row "!tipo!" "!valint!" "!valfrac!" "!trx!" "" "" "!outfile!"
    )

)

echo Conversão concluída.
exit /b


:write_row
rem %1 tipo
rem %2 valint
rem %3 valfrac
rem %4 trx
rem %5 hh
rem %6 mm
rem %7 outfile

setlocal enabledelayedexpansion
set "tipo=%~1"
set "valint=R$%~2"
set "valfrac=%~3"
set "trx=%~4 "
set "hh=%~5"
set "mm=%~6"
set "outfile=%~7"

if not defined valint set "valint=0"
if not defined valfrac set "valfrac=00"
if not defined trx set "trx="

rem normaliza
for /f "tokens=* delims= " %%A in ("!tipo!") do set "tipo=%%A"
for /f "tokens=* delims= " %%A in ("!trx!") do set "trx=%%A"

rem monta horário em uma coluna
if defined hh (
    if not defined mm set "mm=00"
    set "hora_final=!hh!:!mm!"
) else (
    set "hora_final="
)

>> "!outfile!" echo "!tipo!";"!valint!,!valfrac!";"!trx!";"!hora_final!"
endlocal
goto :eof
