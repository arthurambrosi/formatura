# RSVP + Google Sheets

Este projeto prepara uma planilha do Google para receber as confirmacoes dos seus sites e separar tudo em duas abas:

- `Solenidade`
- `Baile`

Tambem cria:

- `Resumo`
- `Config`

## O que o script faz

- organiza as colunas da planilha
- separa os dados de `Solenidade` e `Baile`
- cria um resumo com contagens por status
- cria uma aba `Config` com tudo o que vai para os HTMLs
- deixa a planilha como `anyone with link can view`
- recebe envios por `doPost(e)` na Web App

## Observacao importante

O script consegue preparar a planilha inteira e o backend de recebimento.

O **deploy da Web App** ainda precisa ser feito manualmente uma vez no editor do Apps Script, porque esse passo normalmente e feito pelo menu do proprio Google Apps Script.

## Passo a passo

1. Abra a planilha que voce quer usar.
2. Va em `Extensoes > Apps Script`.
3. Cole o conteudo de `Code.gs`.
4. Rode a funcao `configurarSistemaRSVP`.
5. Autorize o script.
6. Confira as abas `Config`, `Resumo`, `Solenidade` e `Baile`.
7. No Apps Script, faca o deploy:

   - `Deploy > New deployment`
   - tipo: `Web app`
   - `Execute as`: voce
   - `Who has access`: `Anyone`

8. Copie a URL final que termina com `/exec`.
9. Rode a funcao `registrarUrlWebApp('COLE_AQUI_A_URL')`.
10. Veja a aba `Config`: la ficam os nomes dos campos que vao para os HTMLs.

## O que eu vou precisar para ligar os HTMLs

Depois desse passo, para eu atualizar seus sites, eu preciso de **uma unica informacao**:

- a URL da Web App terminando com `/exec`

## Seguranca

Por padrao, este setup deixa a planilha com link publico de **visualizacao**.

Isso e o recomendado.

Se voce realmente quiser liberar edicao publica da planilha, existe a funcao:

- `liberarEdicaoPublicaDaPlanilha()`

Mas isso **nao e recomendado**, porque qualquer pessoa com o link pode alterar ou apagar dados.

## Dica para os sites locais

Como seus HTMLs estao sendo usados localmente (`file:///...`), a integracao mais confiavel e usar:

- `form action="URL_DA_WEBAPP" method="POST"`

em vez de depender primeiro de `fetch`.

## Funcoes principais

- `configurarSistemaRSVP()`
- `registrarUrlWebApp(url)`
- `obterInformacoesParaHtml()`
- `liberarVisualizacaoPorLink()`
- `liberarEdicaoPublicaDaPlanilha()`
- `doGet(e)`
- `doPost(e)`
