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
    - [SOMA](#soma)
    - [MÉDIA](#média)
    - [MÁXIMO](#máximo)
    - [MÍNIMO](#mínimo)
    - [CONT.NÚM](#contnúm)
    - [SOMASE](#somase)
    - [SOMASES](#somases)
    - [CONT.SE](#contse)
    - [CONT.SES](#contses)
    - [MÉDIASE](#médiase)
    - [MÉDIASES](#médiases)
  - [🧠 Lógicas](#-lógicas)
    - [SE](#se)
    - [E](#e)
    - [OU](#ou)
    - [SEERRO](#seerro)
  - [🔎 Procura e Referência](#-procura-e-referência)
    - [PROCV](#procv)
    - [PROCX](#procx)
    - [ÍNDICE](#índice-1)
    - [CORRESP](#corresp)
    - [FILTRO (Excel 365)](#filtro-excel-365)
    - [ÚNICO](#único)
    - [CLASSIFICAR](#classificar)
  - [✍ Texto](#-texto)
    - [CONCAT](#concat)
    - [EXT.TEXTO](#exttexto)
    - [ARRUMAR](#arrumar)
    - [LOCALIZAR](#localizar)
  - [📅 Data](#-data)
    - [HOJE](#hoje)
  - [🧮 Análise de Dados](#-análise-de-dados)
    - [SOMARPRODUTO](#somarproduto)
    - [DATA](#data)
    - [DIAS](#dias)
  - [🔗 Combinações de Funções Mais Usadas](#-combinações-de-funções-mais-usadas)
    - [ÍNDICE + CORRESP](#índice--corresp)
    - [SE + E](#se--e)
    - [SE + OU](#se--ou)
    - [SEERRO + PROCV](#seerro--procv)
    - [SOMARPRODUTO + CONDIÇÕES](#somarproduto--condições)
    - [ÍNDICE + CORRESP + CORRESP](#índice--corresp--corresp)
    - [CONCAT + TEXTO](#concat--texto)
    - [FILTRO + CLASSIFICAR](#filtro--classificar)
    - [ÚNICO + CONT.SE](#único--contse)
    - [HOJE + SE](#hoje--se)
    - [📌 Dica Importante](#-dica-importante)

## 🔢 Matemática e Estatística

### SOMA

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

### MÉDIA

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

### MÁXIMO

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

### MÍNIMO

**Descrição**
Retorna o menor valor do conjunto.

**Resolve**
Encontrar menor custo, pior nota ou menor tempo.

**Sintaxe:**

```excel

=MÍNIMO(intervalo)

```

**Resultado:** esperado: menor valor do intervalo

### CONT.NÚM

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

### SOMASE

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

### SOMASES

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

### CONT.SE

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

### CONT.SES

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

### MÉDIASE

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

### MÉDIASES

**Descrição:**  
Média com múltiplos critérios.

**Resolve:**  
Análises segmentadas.

**Sintaxe:**

```excel

MÉDIASES(intervalo_média; intervalo1; critério1; ...)

```

## 🧠 Lógicas

### SE

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

### E

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

### OU

**Descrição**
Retorna VERDADEIRO se pelo menos uma condição for verdadeira.

**Resolve**
Cenários com alternativas.

**Sintaxe:**

```excel

OU(condição1; ...)

```

### SEERRO

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

### PROCV

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

### PROCX

**Descrição**
Versão moderna e mais poderosa de busca.

**Resolve**
Limitações do PROCV.

**Sintaxe:**

```excel

PROCX(valor; matriz_procura; matriz_retorno)

```

### ÍNDICE

**Descrição**
Retorna valor baseado em posição.

**Resolve**
Busca dinâmica sem depender de ordem de colunas.

**Sintaxe:**

```excel

ÍNDICE(matriz; linha; [coluna])

```

### CORRESP

**Descrição**
Localiza posição de um valor.

**Resolve**
Base para buscas avançadas.

**Sintaxe:**

```excel

CORRESP(valor; matriz; 0)

```

### FILTRO (Excel 365)

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

### ÚNICO

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

### CLASSIFICAR

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

### CONCAT

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

### EXT.TEXTO

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

### ARRUMAR

**Descrição:**  
Remove espaços extras.

**Resolve:**  
Limpeza de dados importados.

**Sintaxe:**

```excel

ARRUMAR(texto)

```

### LOCALIZAR

**Descrição:**  
Encontra posição de texto (case-sensitive).

**Resolve:**  
Identificar padrões.

**Sintaxe:**

```excel

LOCALIZAR(texto_procurado; dentro_texto)

```

## 📅 Data

### HOJE

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

### SOMARPRODUTO

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

### DATA

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

### DIAS

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

### ÍNDICE + CORRESP

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

### SE + E

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

### SE + OU

**O que faz:**
Executa lógica quando pelo menos uma condição é verdadeira.

**Resolve:**
Cenários com alternativas válidas.

**Exemplo:**

```excel

=SE(OU(A1>=7; B1="Aprovado"); "Passou"; "Não passou")

```

### SEERRO + PROCV

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

### SOMARPRODUTO + CONDIÇÕES

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

### ÍNDICE + CORRESP + CORRESP

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

### CONCAT + TEXTO

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

### FILTRO + CLASSIFICAR

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

### ÚNICO + CONT.SE

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

### HOJE + SE

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
