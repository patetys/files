# Relatório de Adaptação ao CNPJ Alfanumérico (Receita Federal)

## Objetivo

Adaptar o sistema (VB6/COM, projeto SyasMult / P46V600) para aceitar o novo
formato de CNPJ alfanumérico da Receita Federal (vigente a partir de
07/2026), **sem quebrar o funcionamento do CNPJ numérico antigo**. Estrutura
do novo formato: 14 caracteres, sendo as 12 primeiras posições alfanuméricas
(0-9, A-Z) e as 2 últimas sempre numéricas (dígitos verificadores). Máscara
visual inalterada (`XX.XXX.XXX/XXXX-XX`).

## Estratégia geral

O CNPJ/CPF, na maior parte do sistema, era tratado como número (`Double`),
o que causava três classes de problema com valores alfanuméricos:

1. **Crash imediato** ("Type mismatch") em parâmetros/campos `Double` ao
   receber uma string com letras.
2. **Truncamento silencioso** em `Format$`/`Val`/`CDbl` (perdem as letras ou
   geram erro).
3. **SQL inválido** em concatenações sem aspas (`... = " & valor`), que só
   funcionavam por coincidência com valores puramente numéricos.

A correção seguiu o mesmo padrão em todos os pontos: trocar o tipo para
`String`, validar com `IsNumeric` apenas onde necessário para preservar o
comportamento antigo byte-a-byte, e usar a nova função `gfCNPJValido` para
validação central do CNPJ.

---

## 1. Núcleo de validação e formatação (M46V999.cls)

Arquivo: **`M46V999.cls`** (classe `clsM46V999`, é o módulo realmente usado
pelo projeto — confirmado via todos os `.vbp`; o arquivo `M46V999.txt`
presente na pasta é uma cópia órfã, não referenciada por nenhum projeto, e
**não deve ser editada**).

- **Nova função `gfCNPJValido(ByVal pCNPJ As String) As Boolean`**: valida 14
  posições, sendo as 12 primeiras alfanuméricas (0-9/A-Z, case-insensitive)
  e as 2 últimas numéricas.
- **`gpForCNPJ`** (mascarar/desmascarar CNPJ): trocada a validação de
  `IsNumeric` para `gfCNPJValido`, nos dois sentidos (aplicar máscara e
  remover máscara).
- **`gfCalCodUnicidade_V1`**, **`Is_BlackList`**, **`gfFormatarCNPJ`**:
  revisadas; já eram seguras para alfanumérico ou foram corrigidas com
  branch `IsNumeric` preservando o comportamento antigo.
- **`gObtemBoletoSegregacao`** e **`mfObtemBancoBoleto`**: parâmetro
  `pCPFCNPJ` trocado de `Double` para `String` (crash garantido com letras).
- **`gfValidaInspecao`**: parâmetro `num_cnpj_cpf` trocado de `Double` para
  `String`; concatenação SQL sem aspas corrigida.
- Comparações de truthiness corrigidas (`<> 0`, `> 0`, `= 0` em campos que
  agora são `String`) em vários pontos, incluindo o `IIf(...Num_Cgc_Cpf = 0...)`
  usado na atualização de status de transmissão de inspeção (duas
  ocorrências).

## 2. Dígito verificador (A99V001.cls)

- **`DIG_MODULO11`**: reescrita para calcular o valor de cada caractere via
  `Asc(UCase(char)) - 48` (em vez de `Val(char)`), que é o algoritmo oficial
  da Receita Federal para o módulo 11 alfanumérico. Para dígitos 0-9 o
  resultado é idêntico ao anterior (nenhuma mudança de comportamento para
  CNPJ/CPF antigos); para letras A-Z, passa a calcular corretamente
  (valores 17-42).
- **`DIG_CNPJ`**: revisada (sem alteração necessária) — confirma que as
  posições de máscara (`.`, `/`, `-`) já tinham peso zero nos vetores de
  peso, compatibilizando com a nova lógica.

## 3. Estruturas de dados (structs `stcXXXXX`) — campo trocado de `Double` para `String`

| Arquivo | Campo |
|---|---|
| `A46V702B.cls` | `Num_Cgc_Cpf` |
| `A46V712B.cls` | `Num_Cgc_Cpf`, `Cpf_Cnpj_Cre` |
| `A46V735B.cls` | `Num_CPF_CNPJ` |
| `A46V102A.cls` | `Pro_CNPJ_CPF` |

Decorrente dessa troca, todos os pontos que liam esses campos do banco com
`gfTratarNulo(..., tpoNumerico)` foram ajustados para `tpoString` em
`A46V101A.cls`, `A46V702A.cls`, `A46V712A.cls` e `M46V001.cls`.

## 4. Parâmetros/campos `Double` que recebiam CNPJ/CPF (risco de crash imediato)

Varredura dedicada (`grep` por `CPFCNPJ|CNPJ_CPF|Cgc_Cpf|CPF_CNPJ ... As Double|Long|Integer|Currency|Single`)
encontrou e corrigiu (para `String`):

- `M46V999.cls`: `gObtemBoletoSegregacao`, `mfObtemBancoBoleto`, `gfValidaInspecao`
- `A46O003A.cls`: `gfAgendarInspecao`, `mfAgendarInspecaoArquivoOFFLine`
- `A46V102A.cls`: `Pro_CNPJ_CPF`
- `T47V001.frm` / `T46V176A.frm` / `T46V208A.frm`: removida conversão
  explícita `CDbl(lDoc)` ao chamar `gfValidaInspecao`.

Varredura final confirmou **zero ocorrências restantes** desse padrão em
todo o repositório.

## 5. Comparações numéricas (`<> 0`, `> 0`, `= 0`) em campos agora `String`

Corrigidas para `gM46V999.gfPreenchido(campo) And Trim(campo) <> "0"` (ou
equivalente para `= 0`) em:

`M46V999.cls`, `A46V103A.cls`, `A46V101A.cls`, `A46O003A.cls`,
`T46V101A.frm`, `T46V102A.frm`, `T47V001.frm`, `T47V004A.frm`.

(Atribuições literais como `.Num_CPF_CNPJ = 0` foram deixadas como estão —
são conversões implícitas inofensivas de `0` para `"0"`, não comparações.)

## 6. `Format$`/`Format`/`Val` numérico aplicado a CNPJ (silenciosamente zera letras)

Corrigido com o padrão "se numérico, formata como antes; senão, completa
com zeros à esquerda preservando o conteúdo alfanumérico":

```vb
If IsNumeric(campo) Then
    resultado = Format$(campo, "00000000000000")
Else
    resultado = Right$(String(14, "0") & Trim(campo), 14)
End If
```

Aplicado em: `A46F003A.cls`, `A46O003A.cls`, `T46V003A.frm`, `T46V012B.frm`,
`T46V101A.frm`, `T46V102A.frm`, `T46V112A.frm`, `T46V145A.frm`,
`T46V174A.frm`, `T47V004A.frm`, `M46V111.cls`, `M46V997.bas`.

Casos típicos corrigidos: geração de chamadas ao SisCad via `Shell`,
montagem do arquivo `.055` (log de inspeção), comparação `mAntCNPJ`/`mAtuCNPJ`
ao detectar alteração de CNPJ em renovação de apólice, `Select Case Len(...)`
para decidir máscara de CNPJ x CPF.

## 7. SQL concatenado sem aspas (sintaxe T-SQL inválida para valor alfanumérico)

Corrigido adicionando aspas simples (`'...'`) ou usando
`gM46V999.gfFormatarCamposSQL(...)` (que já detecta o tipo da variável e
aspeia automaticamente):

`T46V012A.frm`, `T46V012B.frm`, `T46V101A.frm`, `T46V102A.frm`,
`T46V207A.frm`, `T47V001.frm`, `T47V004A.frm`, `A46V601A.cls`.

## 8. Leitura de arquivo de transmissão de largura fixa (`mfRetCampo`)

Em `A46V601A.cls`, 7 chamadas que liam o campo de CNPJ/CPF de um registro
posicional usavam a opção `2` de `mfRetCampo` ("campo numérico — rejeita com
erro se não for número"), o que bloquearia qualquer CNPJ alfanumérico
recebido nesse arquivo. Trocadas para a opção `6` ("alfanumérico com
conteúdo numérico — completa com zero, sem rejeitar"), que preserva o
comportamento exato para CNPJ/CPF numérico e passa a aceitar letras.

(`A46V101A.cls`, equivalente para leitura, já usava a opção `4`
— alfanumérica — e não precisou de ajuste.)

## Limitações fora do escopo (não é código VB6 — requer migração de banco)

Os seguintes parâmetros de stored procedures/tabelas externas têm tipo
**numérico no SQL Server**. O lado VB6 agora envia o CNPJ corretamente como
string entre aspas simples, mas o valor só será aceito de fato pelo SQL
Server se essas procedures/colunas forem migradas para `varchar`:

- `P0044800.dbo.sp_InterfaceCadastro` — parâmetro `@NUM_CGC_CPF bigint`
  (chamado de `T46V101A.frm` e `T47V001.frm`).
- `[YAS].[PROC_VALIDA_REGRAS_SEGREGACAO]` — parâmetro `@NUM_CPFCNPJ_SEG`
  (chamado de `M46V999.cls`, função `mfObtemBancoBoleto`).
- `ramosdiversos.dbo.Proc_Valida_Inspecao` e a tabela `inspecao..tab_insp`
  (`num_cnpj_cpf`) — tipo não confirmado, mesma ressalva.
- `P0044800.DBO.PROC_CONSULTA_PI_BP_EMITIDO` / `PROC_COMPLEMENTA_DADOS_BP`
  (chamadas de `A46V601A.cls`) — mesma ressalva.

Essas migrações de banco devem ser tratadas separadamente pelo time de
DBA/SQL.

## Recomendação de teste

1. Compilar o projeto (`P46V600.vbp` e variantes `D46O003*.vbp`) para
   garantir que nenhuma assinatura `ByRef`/tipo ficou inconsistente.
2. Testar fluxo completo (cadastro, emissão, inspeção, boleto, transmissão
   SisCad) com:
   - CNPJ numérico antigo (regressão — deve continuar idêntico).
   - CNPJ alfanumérico novo (ex.: `12ABC34501DE35`).
   - CPF (não afetado pela mudança da Receita, deve continuar 100% numérico).
3. Validar especificamente o dígito verificador (`DIG_CNPJ`/`DIG_MODULO11`)
   com um CNPJ alfanumérico de teste oficial da Receita Federal.
