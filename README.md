# Top Fashion PDV

Sistema da loja Top Fashion MS: vendas com leitor de código de barras, estoque por cor e tamanho, etiquetas, caixa, lucro e crediário.

- **Telas:** este site (GitHub Pages). Funciona no computador e no celular, e pode ser instalado na tela inicial.
- **Banco de dados:** uma Planilha Google no Drive da loja, acessada pelo `Code.gs` (Google Apps Script).

Link do sistema: **https://topfashioncgms-wq.github.io/top-fashion-pdv/**

---

## Instalação (uma vez só)

### 1. Planilha e Apps Script

1. No Google Drive da loja, crie uma **Planilha Google** chamada *Top Fashion - Sistema*.
2. Na planilha, abra **Extensões > Apps Script**.
3. Apague o conteúdo do arquivo `Código.gs` e cole todo o arquivo [`Code.gs`](Code.gs) deste repositório. Salve.
4. Na barra de cima, escolha a função **configurar** e clique em **Executar**. Autorize com a conta da loja (**Avançado > Acessar... > Permitir**).
5. Clique em **Implantar > Nova implantação > App da Web**:
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
6. Copie o **URL do app da Web**. Ele termina em `/exec`.

### 2. Ligar o site à planilha

Edite o arquivo [`config.js`](config.js) e cole o link entre as aspas:

```js
window.TF_API_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

### 3. GitHub Pages

Em **Settings > Pages**, escolha *Deploy from a branch*, branch **main**, pasta **/ (root)**, e salve. Em 1 ou 2 minutos o link do sistema fica no ar.

---

## Primeiro acesso

| Quem | PIN inicial | O que pode |
|---|---|---|
| Dono | 1234 | Tudo: custo, lucro, financeiro, cancelar venda, ajustes |
| Caixa / vendedora | 4321 | Vender, dar entrada, imprimir etiqueta, crediário |

**Troque os PINs em Ajustes** no primeiro acesso, de preferência com 6 números. Depois de 10 tentativas erradas, o sistema bloqueia novos logins por 15 minutos.

**No celular:** abra o link no Chrome (Android) ou Safari (iPhone) e use **Adicionar à tela inicial**.

## Leitor de código de barras

- **Computador:** leitor USB comum. Ele digita o código e aperta Enter sozinho, sem instalar nada.
- **Celular:** toque no botão rosa da câmera, ao lado do campo de busca, e aponte para a etiqueta. Dá para ler várias peças em seguida.

## Impressora térmica de etiquetas

1. Instale o driver e configure o tamanho do papel igual ao da etiqueta (ex.: 40 x 25 mm).
2. No sistema, em **Ajustes > Etiqueta**, coloque a mesma medida.
3. Na janela de impressão: margens **Nenhuma**, escala **100**, sem cabeçalhos e rodapés.

## Atualizar o sistema

- **Telas** (`index.html`): basta mudar aqui no GitHub. O site atualiza sozinho.
- **Code.gs**: cole a nova versão no Apps Script e vá em **Implantar > Gerenciar implantações > lápis > Nova versão > Implantar**. Assim o link continua o mesmo.

## Cuidados

- Não renomeie as abas nem as colunas da planilha.
- Faça backup de vez em quando: na planilha, **Arquivo > Fazer uma cópia**.
- O sistema precisa de internet e não emite nota fiscal (NFC-e).

---

Leitor por câmera: [html5-qrcode](https://github.com/mebjas/html5-qrcode) (Apache 2.0).
