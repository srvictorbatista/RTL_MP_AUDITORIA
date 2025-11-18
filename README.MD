
# **RTL DE AUDITORIA DE TRANSAÇÕES DO MERCADO PAGO/LIVRE EM CMD/BASH**

### **Descrição técnica e orientações**

Este documento descreve, de forma técnica e detalhada, o funcionamento do conjunto de scripts que compõem a RTL de auditoria de movimentações financeiras exportadas manualmente da área de atividades do Mercado Livre. O conteúdo original consiste em blocos de texto copiáveis diretamente da interface do Mercado Pago. Esses trechos, quando colados em arquivos `.txt`, são processados por dois mecanismos distintos, um em Windows CMD e outro em Bash, garantindo compatibilidade entre ambientes Windows e Linux. Embora sejam implementações diferentes, ambas compartilham o mesmo modelo lógico de parser e geram um arquivo `.csv` estruturado e adequado para análise contábil, reconciliação e auditoria interna.

---

## **1. Contexto e Objetivo Geral**

O fluxo de movimentações exibido pelo ***[Mercado Livre](https://www.mercadopago.com.br/activities/)*** não possui opção de exportação em CSV. Assim, a solução desenvolvida converte esses registros, obtidos por simples cópia e colagem, em um formato tabular.
Os scripts analisam arquivos `.txt` presentes no diretório de execução e produzem arquivos `.csv` equivalentes, aplicando regras determinísticas de interpretação de conteúdo. O CSV resultante é compatível com Excel, utiliza BOM UTF-8 e adota ponto-e-vírgula como separador.

---

## **2. Estrutura Funcional Comum aos Dois Ambientes**

A lógica interna foi projetada como RTL – Runtime Transformation Logic – com estados de leitura e captura de campos, permitindo que os scripts reconstruam a transação mesmo quando o conteúdo original está fragmentado em múltiplas linhas.
A sequência esperada de elementos dentro de cada registro é a seguinte:

1. A linha contendo o tipo da movimentação, como “Venda”, inicia um novo registro.
2. A presença do marcador “R$” sinaliza que os próximos números representam o valor da operação, começando pela parte inteira.
3. Quando surge uma linha contendo apenas uma vírgula, identifica-se o separador entre parte inteira e fracionária.
4. Linhas compostas somente por dígitos são classificadas como números úteis para completar o valor.
5. A linha que contém “Transação”, independentemente de maiúsculas ou minúsculas, fornece o identificador único da operação.
6. A linha que inclui um horário no formato “14h40” — ou uma variação como “14h” — finaliza o conjunto de dados da transação. Nesse momento, o registro é considerado completo e é escrito no CSV.

Caso algum elemento falte, o parser aplica tolerância. Por exemplo: quando não existe parte fracionária, assume-se o valor “00”. Da mesma forma, se o arquivo termina sem que o horário seja encontrado, a transação parcial é persistida no resultado. Linhas vazias são ignoradas e sinais isolados, como “+” e “−”, são preservados. Esse comportamento permite robustez em fluxos reais, nos quais irregularidades são comuns.

---

## **3. Saída Estruturada e Garantias Técnicas**

Ambos os scripts geram um CSV seguindo convenções rigorosas:

* Cada campo é delimitado por aspas duplas.
* O separador entre campos é o ponto-e-vírgula.
* O arquivo sempre inicia com a BOM UTF-8 para compatibilidade com Excel em Windows.
* O cabeçalho inclui colunas principais — tipo, valor, transação, hora e minuto — além de campos adicionais que contêm fórmulas utilizadas para cálculo automático de totais, comissões e faturamento consolidado.

Como detalhe técnico importante, o campo “Transação ID” recebe um caractere NBSP no final. Esse espaço não-quebrável impede que o Excel converta o número em formato científico ou altere sua formatação.

---

## **4. Implementação no Windows CMD**

A versão CMD foi projetada para ambiente nativo Windows, utilizando apenas recursos de batch scripting. Ela processa cada `.txt` encontrado no diretório atual, cria um `.csv` correspondente e atua com estados internos de parsing.
Todo o comportamento descrito anteriormente é preservado, incluindo normalização, remoção de espaços nas extremidades e tratamento seguro de valores.

---

## **5. Implementação em Bash**

A versão Bash mantém o mesmo comportamento lógico, voltada para ambientes Linux.
O script é compatível com Bash 5.x e segue a mesma cadência estrutural: leitura sequencial, montagem do registro, encerramento momentâneo ao encontrar a linha com horário e escrita final no CSV.
Por produzir resultado idêntico ao CMD, os dois ambientes podem ser utilizados alternadamente sem quebra de compatibilidade.

---

## **6. Considerações Técnicas Finais**

A RTL desenvolvida permite converter de forma confiável o conteúdo fragmentado e desestruturado da interface do Mercado Livre em um formato tabular útil para auditorias profissionais. Ela reduz intervenção manual, evita erros de manipulação de dados e padroniza o processo de reconciliação financeira.

Tanto a versão CMD quanto a versão Bash foram construídas com o mesmo compromisso: reprodutibilidade, tolerância a variações de entrada e geração de saída compatível com ferramentas corporativas, especialmente Excel.

Essa documentação consolida a intenção original do projeto: fornecer um mecanismo estável e auditável para transformar relatórios informais em dados consistentes, estruturados e prontos para uso técnico.
