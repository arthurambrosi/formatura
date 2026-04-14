# Campos para os HTMLs

## 1. Paginas simples

Use isso nas paginas separadas de `Solenidade` e `Baile`.

### Campos

- `nome_completo`
- `evento`
- `status`
- `origem`
- `pagina`
- `canal`

### Valores

- `evento`: `solenidade` ou `baile`
- `status`: `confirm`, `decline`, `pending`

### Exemplo

```html
<form action="COLE_AQUI_A_URL_EXEC" method="POST">
  <input type="hidden" name="evento" value="solenidade" />
  <input type="hidden" name="origem" value="Formatura" />
  <input type="hidden" name="pagina" value="solenidade.html" />
  <input type="hidden" name="canal" value="site" />

  <input type="text" name="nome_completo" />

  <input type="radio" name="status" value="confirm" />
  <input type="radio" name="status" value="decline" />
  <input type="radio" name="status" value="pending" />

  <button type="submit">Confirmar</button>
</form>
```

## 2. Pagina combinada

Use isso na pagina que confirma `Solenidade` e `Baile` no mesmo formulario.

### Campos

- `nome_completo`
- `solenidade`
- `baile`
- `origem`
- `pagina`
- `canal`

### Valores aceitos

Voce pode usar qualquer um destes padroes:

- `yes`, `no`, `maybe`
- `confirm`, `decline`, `pending`

### Exemplo

```html
<form action="COLE_AQUI_A_URL_EXEC" method="POST">
  <input type="hidden" name="origem" value="Formatura Arthur" />
  <input type="hidden" name="pagina" value="rsvp.html" />
  <input type="hidden" name="canal" value="site" />

  <input type="text" name="nome_completo" />

  <input type="radio" name="solenidade" value="yes" />
  <input type="radio" name="solenidade" value="no" />
  <input type="radio" name="solenidade" value="maybe" />

  <input type="radio" name="baile" value="yes" />
  <input type="radio" name="baile" value="no" />
  <input type="radio" name="baile" value="maybe" />

  <button type="submit">Confirmar</button>
</form>
```

## 3. O que eu preciso de voce depois

Para eu atualizar os HTMLs pra voce, eu vou precisar so disso:

- a URL da Web App publicada, terminando com `/exec`
