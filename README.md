# Atualização Automatizada de Cotas de Fundos de Investimento com Power Query

## 1. Visão geral do projeto

Este projeto tem como objetivo automatizar a extração, o tratamento e a consolidação das cotas mensais de fundos de investimento diretamente do portal da Comissão de Valores Mobiliários (CVM), utilizando **Power Query e linguagem M**.

A solução foi estruturada para consultar as informações disponibilizadas no sistema de informações diárias de fundos, selecionar os registros de cada competência mensal e consolidar os resultados em uma tabela preparada para análise no Power BI.

O processo utiliza um parâmetro para armazenar o endereço eletrônico do fundo, uma função reutilizável para consultar os dados de cada competência e uma consulta principal responsável por executar a função para os meses desejados.

## 2. Objetivos

- Automatizar a coleta mensal das cotas de fundos de investimento.
- Consultar os dados diretamente no portal da CVM.
- Reduzir a necessidade de extrações e atualizações manuais.
- Reutilizar a mesma função para diferentes competências.
- Consolidar os resultados em uma única tabela.
- Identificar o último registro disponível de cada competência.
- Preparar os dados para análises históricas no Power BI.

## 3. Arquitetura da solução

A solução foi dividida em três componentes principais:

### 3.1. Parâmetro `CaminhoVinci`

Armazena o endereço eletrônico utilizado para acessar a página do fundo no portal da CVM.

```powerquery
"https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=290965&PK_SUBCLASSE=-1"
    meta [
        IsParameterQuery = true,
        Type = "Text",
        IsParameterQueryRequired = true
    ]
```

**Finalidade:**

- Centralizar a URL de consulta.
- Facilitar a manutenção do endereço.
- Permitir que a função utilize o parâmetro sem precisar repetir a URL em seu código.

O parâmetro foi configurado como texto e definido como obrigatório.

**Observação:** o identificador presente na URL deve corresponder ao fundo ou à classe que se deseja consultar. Para utilizar a solução com outros fundos, é necessário confirmar o identificador correto no portal da CVM.

### 3.2. Função `fxVinci`

A função recebe uma competência mensal como argumento e retorna uma tabela com os registros encontrados para o período.

Assinatura:

```powerquery
(Competencia as text) =>
```

O parâmetro `Competencia` recebe o valor selecionado no portal da CVM, conforme o formato esperado pelo controle de competência.

A função utiliza o parâmetro `CaminhoVinci` para acessar a página e executar a consulta.

### 3.3. Consulta principal

A consulta principal gera a lista de competências, executa a função para cada mês, consolida as tabelas retornadas e realiza os tratamentos necessários para disponibilizar a tabela final.

O fluxo é:

**Parâmetro de URL → Função M → Consulta mensal → Consolidação → Seleção da última cota → Tipagem → Power BI**

## 4. Funcionamento da função `fxVinci`

A função é responsável pela extração dos dados da CVM e pela estruturação inicial da tabela mensal.

### 4.1. Primeiro acesso à página

```powerquery
FonteInicial =
    Web.Contents(Url),

HtmlInicial =
    Text.FromBinary(
        FonteInicial,
        TextEncoding.Utf8
    )
```

A função acessa a página informada no parâmetro e converte o conteúdo retornado em texto para permitir a leitura do HTML.

### 4.2. Extração dos campos ASP.NET

```powerquery
ViewState =
    Text.BetweenDelimiters(
        HtmlInicial,
        "id=""__VIEWSTATE"" value=""",
        """"
    ),

ViewStateGenerator =
    Text.BetweenDelimiters(
        HtmlInicial,
        "id=""__VIEWSTATEGENERATOR"" value=""",
        """"
    ),

EventValidation =
    Text.BetweenDelimiters(
        HtmlInicial,
        "id=""__EVENTVALIDATION"" value=""",
        """"
    )
```

A página utiliza campos de estado do ASP.NET para controlar as interações com o formulário.

A função extrai os valores de `__VIEWSTATE`, `__VIEWSTATEGENERATOR` e `__EVENTVALIDATION`, necessários para reproduzir o envio do formulário.

Esses valores são obtidos dinamicamente a cada consulta, em vez de serem mantidos fixos no código.

### 4.3. Envio da competência por POST

```powerquery
PostData =
    [
        __EVENTTARGET = "ddComptc",
        __EVENTARGUMENT = "",
        __VIEWSTATE = ViewState,
        __VIEWSTATEGENERATOR = ViewStateGenerator,
        __EVENTVALIDATION = EventValidation,
        ddComptc = Competencia
    ]
```

O registro `PostData` reúne os campos enviados ao servidor, incluindo a competência selecionada.

A requisição é executada com `Web.Contents`, utilizando o cabeçalho `application/x-www-form-urlencoded` e a função `Uri.BuildQueryString` para preparar os dados do formulário.

Essa etapa reproduz a interação necessária para solicitar os registros da competência desejada.

### 4.4. Extração da tabela HTML

```powerquery
Tabela =
    Html.Table(
        HtmlPost,
        {
            {
                "Dia",
                "table#dgDocDiario tr td:nth-child(1)"
            },
            {
                "Quota",
                "table#dgDocDiario tr td:nth-child(2)"
            }
        },
        [
            RowSelector = "table#dgDocDiario tr"
        ]
    )
```

A função `Html.Table` transforma o conteúdo HTML em uma tabela estruturada.

São extraídas duas colunas:

- **Dia:** dia do registro disponibilizado na tabela.
- **Quota:** valor da cota informado pela fonte.

A seleção utiliza os seletores CSS da tabela identificada como `dgDocDiario`.

### 4.5. Tratamento dos registros

```powerquery
#"Remover Cabeçalho" =
    Table.Skip(Tabela, 1),

#"Remover Vazios" =
    Table.SelectRows(
        #"Remover Cabeçalho",
        each [Dia] <> null and [Dia] <> ""
    )
```

O primeiro passo remove a linha de cabeçalho da tabela extraída.

Em seguida, são excluídas as linhas sem dia preenchido, evitando que registros vazios sejam incorporados ao resultado.

### 4.6. Identificação da competência

```powerquery
#"Adicionar Competencia" =
    Table.AddColumn(
        #"Remover Vazios",
        "Competencia",
        each Competencia
    )
```

A coluna `Competencia` identifica o período consultado.

Isso permite consolidar posteriormente os registros de diferentes meses e distinguir a origem temporal de cada conjunto de dados.

**Resultado da função:** uma tabela com as colunas `Dia`, `Quota` e `Competencia`.

## 5. Funcionamento da consulta principal

### 5.1. Geração das competências

```powerquery
Competencias =
    List.Transform(
        {1..12},
        each
            Text.PadStart(
                Text.From(_),
                2,
                "0"
            ) & "/2026"
    )
```

A lista é construída dinamicamente a partir dos números de 1 a 12, gerando as competências mensais de janeiro a dezembro de 2026.

O uso de `Text.PadStart` garante que os meses de janeiro a setembro tenham dois dígitos.

Exemplos:

- `01/2026`
- `02/2026`
- `03/2026`
- `12/2026`

O ano está definido diretamente no código. Para tornar o processo reutilizável, ele pode futuramente ser transformado em parâmetro.

### 5.2. Execução da função por competência

```powerquery
Dados =
    List.Transform(
        Competencias,
        each try fxVinci(_) otherwise null
    )
```

A função `fxVinci` é executada para cada competência.

A expressão `try ... otherwise null` impede que um erro em uma consulta mensal interrompa automaticamente toda a execução, substituindo o resultado com erro por `null`.

Essa estratégia aumenta a tolerância a falhas, mas exige atenção: um mês que falhar pode desaparecer da consolidação sem que o usuário perceba imediatamente.

### 5.3. Consolidação dos resultados

```powerquery
Validos =
    List.RemoveNulls(Dados),

Consolidado =
    Table.Combine(Validos)
```

`List.RemoveNulls` remove os resultados nulos.

Em seguida, `Table.Combine` reúne as tabelas mensais em uma única estrutura, facilitando a análise histórica das cotas.

### 5.4. Filtragem e agrupamento

A consulta remove registros com quota vazia e agrupa os dados pela competência.

O agrupamento mantém os registros de cada mês em uma tabela aninhada, permitindo selecionar posteriormente o último dia disponível de cada período.

### 5.5. Seleção do último registro mensal

```powerquery
Table.Max(
    [Tabela],
    "Dia"
)
```

A função `Table.Max` seleciona o registro com o maior valor da coluna `Dia` dentro de cada competência.

O objetivo é obter uma cotação representativa do último registro disponível no período.

**Ponto de atenção:** a coluna `Dia` é extraída inicialmente como texto. A comparação textual pode não representar corretamente a ordem cronológica ou numérica em todos os formatos.

Para garantir a seleção correta, recomenda-se converter o dia para um tipo numérico ou de data antes de aplicar a ordenação, conforme o formato efetivamente retornado pela CVM.

### 5.6. Identificador mensal e identificação do fundo

A consulta cria a coluna `Id_Mes` com um índice sequencial e adiciona a coluna `Fundo`, preenchida com o valor `"Vinci"`.

Essas colunas auxiliam na identificação dos registros e na utilização da tabela em relatórios.

O índice é gerado conforme a ordem dos registros resultantes e não deve ser tratado como um identificador permanente entre atualizações caso a quantidade ou a ordem dos registros mude.

### 5.7. Data da última atualização e tipagem

A consulta combina `Dia` e `Competencia` para construir a coluna `DataUltimaAtualizacao`.

Ao final, aplica os tipos de dados esperados:

- `Id_Mes`: número inteiro.
- `Competencia`: data.
- `Dia`: número.
- `Quota`: número.
- `Fundo`: texto.
- `DataUltimaAtualizacao`: data.

A tipagem facilita o uso dos campos em cálculos, filtros e visualizações do Power BI.

**Ponto de atenção:** a conversão da competência e da data concatenada depende do formato textual recebido. Caso haja erros de conversão, recomenda-se construir as datas explicitamente ou utilizar uma cultura apropriada.

## 6. Resultado esperado

A consulta principal produz uma tabela consolidada com uma linha por competência, contendo o último registro selecionado para cada mês.

Estrutura esperada:

| Coluna | Finalidade |
|---|---|
| `Id_Mes` | Identificador sequencial do mês no resultado |
| `Competencia` | Competência mensal |
| `Dia` | Dia do registro selecionado |
| `Quota` | Valor da cota do fundo |
| `Fundo` | Identificação do fundo |
| `DataUltimaAtualizacao` | Data construída a partir do dia e da competência |

A tabela pode ser utilizada no Power BI para acompanhar a evolução histórica das cotas, construir gráficos de desempenho e comparar períodos.

## 7. Principais tecnologias e recursos utilizados

- **Power Query:** extração, transformação e consolidação dos dados.
- **Linguagem M:** criação das consultas, funções e regras de tratamento.
- **Web.Contents:** acesso ao portal da CVM.
- **Html.Table:** extração de dados tabulares do HTML.
- **Text.BetweenDelimiters:** extração dos campos de estado do formulário ASP.NET.
- **Uri.BuildQueryString:** preparação dos parâmetros enviados por POST.
- **List.Transform:** geração das competências e execução repetida da função.
- **Table.Combine:** consolidação das tabelas mensais.
- **Table.Group:** agrupamento por competência.
- **Table.Max:** seleção do registro de maior dia dentro de cada competência.
- **Table.TransformColumnTypes:** definição dos tipos de dados.

## 8. Validações e cuidados técnicos

Para manter a confiabilidade dos resultados, recomenda-se verificar os seguintes pontos:

1. **Integridade da extração:** conferir se todas as competências previstas retornaram dados.
2. **Falhas silenciosas:** registrar ou sinalizar as competências em que a função apresentou erro, em vez de depender exclusivamente de `otherwise null`.
3. **Ordenação cronológica:** assegurar que o último registro seja selecionado com base em uma coluna numérica ou de data corretamente tipada.
4. **Conversão de tipos:** validar a interpretação dos valores de quota e das datas, considerando os formatos de origem.
5. **Atualização do portal:** verificar se os identificadores HTML e os campos ASP.NET permanecem compatíveis com a implementação.
6. **Consistência histórica:** comparar os valores extraídos com os registros publicados no portal da CVM.
7. **Limites da fonte:** considerar que a disponibilidade dos dados depende da publicação e do funcionamento do sistema consultado.

## 9. Possíveis melhorias

Como evolução da solução, podem ser implementadas:

- Parametrização do ano de referência.
- Reutilização da função para diferentes fundos.
- Registro de erros por competência.
- Tratamento explícito de valores nulos e quotas inválidas.
- Conversão de datas antes da seleção do último registro.
- Validação da quantidade de registros retornados por mês.
- Comparação automática entre atualizações para detectar alterações históricas.
- Inclusão de indicadores de rentabilidade e variação percentual das cotas.

## 10. Conclusão

O projeto demonstra a utilização do Power Query e da linguagem M para automatizar a coleta de dados públicos financeiros, estruturando um processo reutilizável de extração, transformação e consolidação.

A separação entre parâmetro, função e consulta principal facilita a manutenção e permite ampliar a solução para outros fundos e períodos.

A integração com o Power BI oferece uma base para acompanhar a evolução das cotas de fundos de investimento, com potencial para análises históricas e indicadores de desempenho.

Mais do que automatizar uma consulta, a solução estabelece um fluxo de dados documentado, reproduzível e preparado para evoluções futuras.

---
### 🚀 Projetos e Portfólio
Desenvolvo projetos práticos envolvendo análise de dados, dashboards, indicadores de desempenho, ETL e visualização analítica, conectando tecnologia e problemas reais de negócio.

- 📊 [Portfólio com análises no Medium](https://medium.com/@maironhenrique)
- 📈 [Portfólio com Dashboards](https://app.xperiun.com/in/maironhenrique)
---
### 📫 Contato
**E-mail**: maironhenrique@hotmail.com  
**LinkedIn**: https://www.linkedin.com/in/maironhenriquedeoliveira
