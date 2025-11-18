#!/usr/bin/env bash
######################################################################################################################################################################
# TXT_TO_CSV.SH [RTL em script bash] Versão 0.0.1 
######################################################################################################################################################################
#
# Este script foi concebido para processar arquivos .txt oriundos da página de movidemntações do MercadoLivre (https://www.mercadopago.com.br/activities/2): 
# Tendo o conteudo do relatorio, literalmente sido copiado e colado em um arquivo de bloco de notas simples. Para fins de analise e/ou auditoria interna.
#
# Funcionamento geral:
# - Para cada arquivo *.txt no diretório atual cria um arquivo com o mesmo nome e extensão .csv.
# - Grava no início do .csv a BOM UTF-8 (para compatibilidade com Excel) e a linha de cabeçalho:
#   "MOV TIPO";"VALOR";"TRANSAÇÃO ID";"HORÁRIO"; ;;;"TOTAL VENDAS:";"=SOMA(B2:B2000)"; ;"COMISSÃO:";"=I1/100*5"; ;"FATURA:";"=I1-L1";
# - Analisa o .txt linha a linha, montando registros com o seguinte fluxo tolerante:
#     1) Linha com o tipo (ex.: "Venda") inicia um novo registro.
#     2) Linha contendo "R$" marca que o próximo(s) número(s) representam parte inteira do valor.
#     3) Linha contendo apenas "," indica que os próximos números são a parte fracionária (centavos).
#     4) Linhas numéricas (apenas dígitos) são usadas para preencher parte inteira ou fracionária conforme o estado.
#     5) Linha contendo "Transa" (case-insensitive) captura o ID da transação (removendo possíveis vírgulas).
#     6) Linha contendo "h" (ex.: "14h40" ou "14h") captura hora e minuto e, ao encontrá-la, escreve o registro completo no .csv.
# - Se ao final do arquivo houver um registro pendente (sem linha de hora registrada), ele também será escrito.
# - O campo TRANSAÇÃO recebe um caractere NBSP (U+00A0) ao final, para forçar interpretação textual no Excel, seguindo a intenção original.
#
# Detalhes de tolerância e normalização:
# - Se não houver parte fracionária, assume-se "00".
# - Sinais "+" ou "-" são respeitados quando aparecem como linhas separadas.
# - Espaços em branco nas extremidades das linhas são removidos.
# - Linhas vazias são ignoradas.
#
# Observações de implementação:
# - Compatível com bash em Ubuntu (testado com bash 5.x).
# - Mantém comportamento funcional equivalente ao script original (desenvolvido em CMD/BAT), corrigindo a montagem do campo VALOR no output.
# - Produz CSV separado por ponto-e-vírgula e campos entre aspas, pronto para importação no Excel configurado por separadores ; .
#
# Notas adicionais:
# O script supõe que as linhas num .txt seguem o padrão observado no exemplo: tipo → linha com R$ → parte inteira → linha , → parte fracionária → linha Transação <id> → linha <hh>h<mm> (ou <hh>h). Se algum arquivo variar, o parser tenta ser tolerante (p.ex. fracionário ausente assume 00).
#
# O caractere especial adicionado ao final do campo TRANSAÇÃO é um NBSP (U+00A0), igual à intenção original, para forçar interpretação textual no Excel.
######################################################################################################################################################################

set -euo pipefail
IFS=$'\n\t'
shopt -s nullglob

write_row() {
    local tipo="$1"
    local valint="$2"
    local valfrac="$3"
    local trx="$4"
    local hh="$5"
    local mm="$6"
    local outfile="$7"

    [[ -z "$valint" ]] && valint="0"
    [[ -z "$valfrac" ]] && valfrac="00"
    [[ -z "$trx" ]] && trx=""

    tipo="$(printf "%s" "$tipo" | sed -e 's/^[ \t]*//' -e 's/[ \t]*$//')"
    trx="$(printf "%s" "$trx" | sed -e 's/^[ \t]*//' -e 's/[ \t]*$//')"

    local hora_final=""
    if [[ -n "$hh" ]]; then
        [[ -z "$mm" ]] && mm="00"
        hora_final="${hh}:${mm}"
    fi

    trx="${trx}"$'\u00A0'

    printf '"%s";"R$%s,%s";"%s";"%s"\n' \
        "$tipo" "$valint" "$valfrac" "$trx" "$hora_final" >> "$outfile"
}

for file in *.txt; do
    outfile="${file%.*}.csv"

    printf '\xEF\xBB\xBF' > "$outfile"

    echo '"MOV TIPO";"VALOR";"TRANSAÇÃO ID";"HORÁRIO"; ;;;"TOTAL VENDAS:";"=SOMA(B2:B200000)"; ;"COMISSÃO:";"=I1/100*5"; ;"FATURA:";"=I1-L1";' >> "$outfile"

    tipo=""
    valint=""
    valfrac=""
    trx=""
    hh=""
    mm=""
    inRecord=0
    sawR=0
    sawComma=0
    sign=""

    while IFS= read -r line || [[ -n "$line" ]]; do

        ln="${line#"${line%%[![:space:]]*}"}"
        ln="${ln%"${ln##*[![:space:]]}"}"

        [[ -z "$ln" ]] && continue

        # Horário (ex.: 14h40)
        if [[ "$ln" == *h* ]]; then
            hh="${ln%%h*}"
            mm="${ln#*h}"
            mm="${mm//[[:space:]]/}"

            [[ -z "$mm" ]] && mm="0"
            [[ ${#mm} -eq 1 ]] && mm="0$mm"

            [[ -z "$valint" ]] && valint="0"
            [[ -z "$valfrac" ]] && valfrac="00"

            if [[ -n "$sign" ]]; then
                valint="${sign}${valint}"
            fi

            write_row "$tipo" "$valint" "$valfrac" "$trx" "$hh" "$mm" "$outfile"

            tipo=""
            valint=""
            valfrac=""
            trx=""
            hh=""
            mm=""
            inRecord=0
            sawR=0
            sawComma=0
            sign=""

            continue
        fi

        # Início de novo registro
        if [[ $inRecord -eq 0 ]]; then
            tipo="$ln"
            inRecord=1
            valint=""
            valfrac=""
            trx=""
            hh=""
            mm=""
            sawR=0
            sawComma=0
            sign=""
            continue
        fi

        # Transação — agora capturando o ÚLTIMO TOKEN, igual ao CMD atualizado
        if printf "%s\n" "$ln" | grep -qi "Transa"; then

            # captura o último token da linha
            last=$(printf "%s" "$ln" | awk '{print $NF}')

            # remove vírgulas residuais
            trx="${last//,/}"

            continue
        fi

        # Flag R$
        if [[ "$ln" == *"R$"* ]]; then
            sawR=1
            continue
        fi

        if [[ "$ln" == "," ]]; then
            sawComma=1
            continue
        fi

        if [[ "$ln" == "-" ]]; then
            sign="-"
            continue
        fi

        if [[ "$ln" == "+" ]]; then
            sign="+"
            continue
        fi

        # Valores numéricos
        if [[ "$ln" =~ ^[0-9]+$ ]]; then
            if [[ $sawComma -eq 1 ]]; then
                [[ -z "$valfrac" ]] && valfrac="$ln"
            else
                [[ -z "$valint" ]] && valint="$ln"
            fi
        fi

    done < "$file"

    # Registro pendente no final
    if [[ $inRecord -ne 0 && -n "$valint" ]]; then
        [[ -z "$valfrac" ]] && valfrac="00"
        if [[ -n "$sign" ]]; then
            valint="${sign}${valint}"
        fi
        write_row "$tipo" "$valint" "$valfrac" "$trx" "" "" "$outfile"
    fi

done

echo "Conversão concluída."
