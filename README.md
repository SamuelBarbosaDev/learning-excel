# Learning Excel

Este guia é voltado para usuários que já conhecem o básico de Excel e querem dominar funções essenciais para análise de dados, automação de relatórios e manipulação de informações.

Cada função contém:

- Descrição
- Qual problema resolve
- Sintaxe (parâmetros)
- Exemplo
- Resultado retornado

## Índice

- [Learning Excel](#learning-excel)
  - [Índice](#índice)
  - [🔢 Matemática e Estatística](#-matemática-e-estatística)
    - [1) SOMA](#1-soma)
    - [2) MÉDIA](#2-média)
    - [3) MÁXIMO](#3-máximo)
    - [4) MÍNIMO](#4-mínimo)
    - [5) CONT.NÚM](#5-contnúm)
    - [16) SOMASE](#16-somase)
    - [17) SOMASES](#17-somases)
    - [18) CONT.SE](#18-contse)
    - [19) CONT.SES](#19-contses)
    - [20) MÉDIASE](#20-médiase)
    - [21) MÉDIASES](#21-médiases)
  - [🧠 Lógicas](#-lógicas)
    - [6) SE](#6-se)
    - [7) E](#7-e)
    - [8) OU](#8-ou)
    - [9) SEERRO](#9-seerro)
  - [🔎 Procura e Referência](#-procura-e-referência)
    - [10) PROCV](#10-procv)
    - [11) PROCX](#11-procx)
    - [12) ÍNDICE](#12-índice)
    - [13) CORRESP](#13-corresp)
    - [23) FILTRO (Excel 365)](#23-filtro-excel-365)
    - [24) ÚNICO](#24-único)
    - [25) CLASSIFICAR](#25-classificar)
  - [✍ Texto](#-texto)
    - [14) CONCAT](#14-concat)
    - [26) EXT.TEXTO](#26-exttexto)
    - [27) ARRUMAR](#27-arrumar)
    - [28) LOCALIZAR](#28-localizar)
  - [📅 Data](#-data)
    - [15) HOJE](#15-hoje)
  - [🧮 Análise de Dados](#-análise-de-dados)
    - [22) SOMARPRODUTO](#22-somarproduto)
    - [29) DATA](#29-data)
    - [30) DIAS](#30-dias)
  - [🔗 Combinações de Funções Mais Usadas](#-combinações-de-funções-mais-usadas)
    - [1) ÍNDICE + CORRESP](#1-índice--corresp)
    - [2) SE + E](#2-se--e)
    - [3) SE + OU](#3-se--ou)
    - [4) SEERRO + PROCV](#4-seerro--procv)
    - [5) SOMARPRODUTO + CONDIÇÕES](#5-somarproduto--condições)
    - [6) ÍNDICE + CORRESP + CORRESP](#6-índice--corresp--corresp)
    - [7) CONCAT + TEXTO](#7-concat--texto)
    - [8) FILTRO + CLASSIFICAR](#8-filtro--classificar)
    - [9) ÚNICO + CONT.SE](#9-único--contse)
    - [10) HOJE + SE](#10-hoje--se)
    - [📌 Dica Importante](#-dica-importante)

## 🔢 Matemática e Estatística

### 1) SOMA

**Descrição**
Adiciona valores numéricos individuais, intervalos ou combinações de ambos. Ignora células vazias e textos.

**Resolve**
Totalizações rápidas como somar vendas, despesas, horas trabalhadas ou quantidades em estoque.

**Sintaxe:**

```excel

SOMA(número1; [número2]; ...)

```

**Exemplo:**

```excel

=SOMA(A1:A5)

```

Se A1:A5 = 10, 20, 30, 40, 50

**Resultado:**

```output

150

```

### 2) MÉDIA

**Descrição**
Calcula a média aritmética de valores numéricos.

**Resolve**
Avaliar desempenho médio, como notas de alunos, faturamento médio mensal ou tempo médio de atendimento.

**Sintaxe:**

```excel

MÉDIA(número1; [número2]; ...)

```

**Exemplo:**

```excel

=MÉDIA(A1:A4)

```

Valores: 6, 8, 10, 6

**Resultado:**

```output

7,5

```

### 3) MÁXIMO

**Descrição**
Retorna o maior valor dentro de um conjunto de dados.

**Resolve**
Identificar picos de vendas, maior salário, maior temperatura etc.

**Sintaxe:**

```excel

MÁXIMO(intervalo)

```

**Exemplo:**

```excel

=MÁXIMO(A1:A5)

```

Valores: 5, 12, 7, 20, 9

**Resultado:**

```output

20

```

### 4) MÍNIMO

**Descrição**
Retorna o menor valor do conjunto.

**Resolve**
Encontrar menor custo, pior nota ou menor tempo.

**Sintaxe:**

```excel

=MÍNIMO(intervalo)

```

**Resultado:** esperado: menor valor do intervalo

### 5) CONT.NÚM

**Descrição**
Conta quantas células possuem números.

**Resolve**
Descobrir quantos registros numéricos válidos existem.

**Sintaxe:**

```excel

CONT.NÚM(intervalo)

```

**Exemplo:**

A1:A5 = 10, "Texto", 5, vazio, 8

**Resultado:**

```output

3

```

### 16) SOMASE

**Descrição:**  
Soma valores com base em um critério específico.

**Resolve:**  
Somar valores filtrados por condição (ex: somar vendas de um vendedor específico).

**Sintaxe:**

```excel

SOMASE(intervalo; critério; [intervalo_soma])

```

**Exemplo:**

```excel

=SOMASE(A1:A5;">10")

```

Se A1:A5 = 5, 15, 20, 8, 12

**Resultado:**

```output

47

```

### 17) SOMASES

**Descrição:**  
Soma valores usando múltiplos critérios.

**Resolve:**  
Análises condicionais complexas (ex: vendas de João em Janeiro).

**Sintaxe:**

```excel

SOMASES(intervalo_soma; intervalo1; critério1; ...)

```

**Exemplo:**

```excel

=SOMASES(C:C;A:A;"João";B:B;"Janeiro")

```

**Resultado:**  
Soma dos valores em C que atendem ambos critérios.

### 18) CONT.SE

**Descrição:**  
Conta células que atendem um critério.

**Resolve:**  
Contar ocorrências (ex: quantos alunos passaram).

**Sintaxe:**

```excel

CONT.SE(intervalo; critério)

```

**Exemplo:**

```excel

=CONT.SE(A1:A5;">=7")

```

**Resultado:**  
Quantidade de valores ≥ 7.

### 19) CONT.SES

**Descrição:**  
Conta com múltiplos critérios.

**Resolve:**  
Análises com mais de uma condição.

**Sintaxe:**

```excel

CONT.SES(intervalo1; critério1; ...)

```

**Exemplo:**

```excel

=CONT.SES(A:A;"João";B:B;"Aprovado")

```

**Resultado:**  
Número de registros que atendem ambos critérios.

### 20) MÉDIASE

**Descrição:**  
Calcula média com base em critério.

**Resolve:**  
Média de subconjuntos de dados.

**Sintaxe:**

```excel

MÉDIASE(intervalo; critério; [intervalo_média])

```

**Exemplo:**

```excel

=MÉDIASE(A1:A5;">=7")

```

**Resultado:**  
Média apenas dos valores ≥7.

### 21) MÉDIASES

**Descrição:**  
Média com múltiplos critérios.

**Resolve:**  
Análises segmentadas.

**Sintaxe:**

```excel

MÉDIASES(intervalo_média; intervalo1; critério1; ...)

```

## 🧠 Lógicas

### 6) SE

**Descrição**
Executa um teste lógico e retorna valores diferentes dependendo **resultado:**.

**Resolve**
Automatizar decisões como aprovação/reprovação, bônus, status de pagamento etc.

**Sintaxe:**

```excel

SE(teste_lógico; valor_se_verdadeiro; valor_se_falso)

```

**Exemplo:**

```excel

=SE(A1>=7;"Aprovado";"Reprovado")

```

Se A1 = 8

**Resultado:**

```output

"Aprovado"

```

### 7) E

**Descrição**
Retorna VERDADEIRO apenas se todas as condições forem verdadeiras.

**Resolve**
Regras com múltiplos critérios obrigatórios.

**Sintaxe:**

```excel

E(condição1; condição2; ...)

```

**Exemplo:**

```excel

=E(A1>=7;B1>=75%)

```

**Resultado:**

VERDADEIRO ou FALSO

### 8) OU

**Descrição**
Retorna VERDADEIRO se pelo menos uma condição for verdadeira.

**Resolve**
Cenários com alternativas.

**Sintaxe:**

```excel

OU(condição1; ...)

```

### 9) SEERRO

**Descrição**
Captura erros em fórmulas e substitui por outro valor.

**Resolve**
Evitar #DIV/0!, #N/D e outros erros em relatórios.

**Sintaxe:**

```excel

SEERRO(valor; valor_se_erro)

```

**Exemplo:**

```excel

=SEERRO(A1/B1;0)

```

Se B1 = 0

**Resultado:**

```output

0

```

## 🔎 Procura e Referência

### 10) PROCV

**Descrição**
Busca um valor na primeira coluna de uma tabela e retorna um valor correspondente de outra coluna.

**Resolve**
Buscar preços, nomes, códigos ou dados relacionados.

**Sintaxe:**

```excel

PROCV(valor_procurado; tabela; núm_coluna; [procurar_intervalo])

```

**Exemplo:**

```excel

=PROCV("João";A2:B10;2;FALSO)

```

**Resultado:**

Retorna o valor correspondente da coluna 2.

### 11) PROCX

**Descrição**
Versão moderna e mais poderosa de busca.

**Resolve**
Limitações do PROCV.

**Sintaxe:**

```excel

PROCX(valor; matriz_procura; matriz_retorno)

```

### 12) ÍNDICE

**Descrição**
Retorna valor baseado em posição.

**Resolve**
Busca dinâmica sem depender de ordem de colunas.

**Sintaxe:**

```excel

ÍNDICE(matriz; linha; [coluna])

```

### 13) CORRESP

**Descrição**
Localiza posição de um valor.

**Resolve**
Base para buscas avançadas.

**Sintaxe:**

```excel

CORRESP(valor; matriz; 0)

```

### 23) FILTRO (Excel 365)

**Descrição:**  
Extrai dados que atendem critérios.

**Resolve:**  
Substitui filtros manuais.

**Sintaxe:**

```excel

FILTRO(matriz; incluir)

```

**Exemplo:**

```excel

=FILTRO(A1:B10;B1:B10="Aprovado")

```

**Resultado:**  
Retorna apenas linhas aprovadas.

### 24) ÚNICO

**Descrição:**  
Retorna valores sem duplicatas.

**Resolve:**  
Listas únicas automáticas.

**Sintaxe:**

```excel

ÚNICO(matriz)

```

**Resultado:**  
Lista sem repetições.

### 25) CLASSIFICAR

**Descrição:**  
Ordena dados dinamicamente.

**Resolve:**  
Ordenação automática.

**Sintaxe:**

```excel

CLASSIFICAR(matriz; [índice]; [ordem])

```

**Exemplo:**

```excel

=CLASSIFICAR(A1:A10)

```

## ✍ Texto

### 14) CONCAT

**Descrição**
Une textos.

**Resolve**
Combinar nomes, códigos e descrições.

**Sintaxe:**

```excel

CONCAT(texto1; ...)

```

**Exemplo:**

```excel

=CONCAT("Olá ";A1)

```

### 26) EXT.TEXTO

**Descrição:**  
Extrai parte do texto.

**Resolve:**  
Separar códigos e padrões.

**Sintaxe:**

```excel

EXT.TEXTO(texto; início; núm_caract)

```

**Exemplo:**

```excel

=EXT.TEXTO("ABC123";4;3)

```

**Resultado:**

```output

123

```

### 27) ARRUMAR

**Descrição:**  
Remove espaços extras.

**Resolve:**  
Limpeza de dados importados.

**Sintaxe:**

```excel

ARRUMAR(texto)

```

### 28) LOCALIZAR

**Descrição:**  
Encontra posição de texto (case-sensitive).

**Resolve:**  
Identificar padrões.

**Sintaxe:**

```excel

LOCALIZAR(texto_procurado; dentro_texto)

```

## 📅 Data

### 15) HOJE

**Descrição**
Retorna data atual do sistema.

**Resolve**
Relatórios automáticos baseados na data.

**Sintaxe:**

```excel

HOJE()

```

**Resultado:**

Ex:

```output

11/02/2026

```

## 🧮 Análise de Dados

### 22) SOMARPRODUTO

**Descrição:**  
Multiplica arrays e soma os resultados.

**Resolve:**  
Cálculos ponderados e análises sem colunas auxiliares.

**Sintaxe:**

```excel

SOMARPRODUTO(matriz1; matriz2)

```

**Exemplo:**

```excel

=SOMARPRODUTO(A1:A3;B1:B3)

```

Se A = 2,3,4 e B = 10,20,30

**Resultado:**

```output

200

```

### 29) DATA

**Descrição:**  
Cria datas válidas.

**Resolve:**  
Padronização de datas.

**Sintaxe:**

```excel

DATA(ano; mês; dia)

```

**Exemplo:**

```excel

=DATA(2026;2;11)

```

**Resultado:**

```output

11/02/2026

```

### 30) DIAS

**Descrição:**  
Calcula diferença entre datas.

**Resolve:**  
Controle de prazos.

**Sintaxe:**

```excel

DIAS(data_final; data_inicial)

```

**Exemplo:**

```excel

=DIAS("10/02/2026";"01/02/2026")

```

**Resultado:**

```output

9

```

## 🔗 Combinações de Funções Mais Usadas

Muitas soluções poderosas no Excel não vêm de uma única função, mas da combinação entre elas.  
Essas combinações permitem buscas dinâmicas, análises condicionais avançadas e modelos mais robustos.

Cada combinação abaixo mostra:

- O que faz  
- Qual problema resolve  
- Como funciona  
- Exemplo prático  

### 1) ÍNDICE + CORRESP

**O que faz:**
Busca valores em uma tabela de forma dinâmica, sem a limitação de procurar apenas da esquerda para a direita.

**Resolve:**
Supera limitações do PROCV:

- Pode buscar para qualquer direção  
- Não quebra ao inserir colunas  
- Funciona em grandes bases de dados  

**Como funciona:**
CORRESP encontra a posição.  
ÍNDICE retorna o valor nessa posição.

**Sintaxe:**

```excel

ÍNDICE(matriz_retorno; CORRESP(valor_procurado; matriz_procura; 0))

```

**Exemplo:**

```excel

=ÍNDICE(B:B; CORRESP("João"; A:A; 0))

```

Se:
A:A = nomes  
B:B = salários  

**Resultado:**
Retorna o salário de João.

### 2) SE + E

**O que faz:**
Executa uma ação apenas se múltiplas condições forem verdadeiras.

**Resolve:**
Regras de negócio com vários critérios obrigatórios.

**Sintaxe:**

```excel

SE(E(cond1; cond2); valor_se_verdadeiro; valor_se_falso)

```

**Exemplo:**

```excel

=SE(E(A1>=7; B1>=75%); "Aprovado"; "Reprovado")

```

**Resultado:**
"Aprovado" somente se nota ≥7 E frequência ≥75%.

### 3) SE + OU

**O que faz:**
Executa lógica quando pelo menos uma condição é verdadeira.

**Resolve:**
Cenários com alternativas válidas.

**Exemplo:**

```excel

=SE(OU(A1>=7; B1="Aprovado"); "Passou"; "Não passou")

```

### 4) SEERRO + PROCV

**O que faz:**
Evita que buscas retornem erros visíveis.

**Resolve:**
Relatórios mais limpos e profissionais.

**Sintaxe:**

```excel

SEERRO(PROCV(...); "Não encontrado")

```

**Exemplo:**

```excel

=SEERRO(PROCV(A1; A:B; 2; FALSO); "Não encontrado")

```

**Resultado:**
Se não achar o valor, mostra "Não encontrado" em vez de #N/D.

### 5) SOMARPRODUTO + CONDIÇÕES

**O que faz:**
Permite soma com múltiplos critérios sem SOMASES.

**Resolve:**
Análises avançadas em versões antigas do Excel.

**Exemplo:**

```excel

=SOMARPRODUTO((A1:A10="João")*(B1:B10="Jan")*(C1:C10))

```

**Resultado:**
Soma valores de João em Janeiro.

### 6) ÍNDICE + CORRESP + CORRESP

**O que faz:**
Busca em duas dimensões (linha e coluna).

**Resolve:**
Tabelas matriciais.

**Sintaxe:**

```excel

ÍNDICE(matriz;
CORRESP(valor_linha; col_linhas; 0);
CORRESP(valor_coluna; col_cabeçalho; 0))

```

**Exemplo:**
Buscar vendas de João em Março numa tabela de meses.

**Resultado:**
Valor exato na interseção.

### 7) CONCAT + TEXTO

**O que faz:**
Combina texto com números formatados.

**Resolve:**
Criação de mensagens dinâmicas.

**Exemplo:**

```excel

=CONCAT("Total: R$ "; TEXTO(A1;"0,00"))

```

**Resultado:**
"Total: R$ 150,00"

### 8) FILTRO + CLASSIFICAR

**O que faz:**
Filtra e ordena automaticamente.

**Resolve:**
Relatórios dinâmicos sem Tabela Dinâmica.

**Exemplo:**

```excel

=CLASSIFICAR(FILTRO(A2:C20; C2:C20="Aprovado"))

```

**Resultado:**
Lista apenas aprovados já ordenados.

### 9) ÚNICO + CONT.SE

**O que faz:**
Cria resumo de frequência.

**Resolve:**
Análise de ocorrências.

**Exemplo:**
Lista única:

```excel

=ÚNICO(A:A)

```

Contagem:

```excel

=CONT.SE(A:A; D1)

```

**Resultado:**
Quantas vezes cada item aparece.

### 10) HOJE + SE

**O que faz:**
Automatiza status baseado em datas.

**Resolve:**
Controle de prazos e vencimentos.

**Exemplo:**

```excel

=SE(A1<HOJE();"Vencido";"No prazo")

```

**Resultado:**
Status automático por data.

### 📌 Dica Importante

Quanto mais você combina funções:

- Menos colunas auxiliares precisa  
- Mais dinâmicas suas planilhas ficam  
- Maior é a escalabilidade do modelo  

Dominar combinações é o que diferencia usuários intermediários de avançados.
