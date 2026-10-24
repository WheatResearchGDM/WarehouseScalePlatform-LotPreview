# Warehouse Scale Platform

Aplicativo operacional para pesagem de parcelas em ensaios de trigo.

## Aplicativo publicado

[Abrir o aplicativo no GitHub Pages](https://wheatresearchgdm.github.io/WarehouseScalePlatform/)

A versão do GitHub Pages salva os pesos no navegador do equipamento e atualiza o
relatório imediatamente. O progresso é local ao navegador e dispositivo; use os
arquivos exportados para backup ou transferência. O endpoint D1 permanece no código,
mas não participa do fluxo ativo de pesagem.

A conexão com a balança requer Chrome ou Edge, acesso por HTTPS e autorização do
operador para a porta serial. A configuração padrão é 9600 baud e pode ser alterada
na própria tela antes da conexão.

## Funcionalidades

- leitura de parcelas por FEID ou UUID, compatível com leitores que funcionam como teclado;
- conferência de Entity name, (OBS) Name, Block, Entry code, Row, Column e (GER) Name;
- registro e atualização de `Plot weight` ou `Seed weight`, escolhidos por sessão;
- conexão direta com balanças pela porta COM usando Web Serial, com velocidade configurável;
- fator de escala configurável para multiplicar ou dividir a leitura bruta de `10¹` a `10¹⁰`;
- preenchimento automático da variável de peso selecionada a partir da leitura serial;
- importação de arquivos Excel `.xlsx` ou `.xls` pelo operador;
- arquivo único Excel/CSV com pesos, lotes e parcelas pendentes, exportado apenas pelo Dashboard;
- Lot site e Storage em cascata, com Lot location, First order e Second order livres e continuidade opcional dos cinco campos;
- progresso em tempo real por ensaio e local, calculado pela quantidade real de parcelas importadas;
- interface responsiva com a identidade visual da GDM;
- sessões independentes em IndexedDB, restauradas no mesmo navegador.

## Dados

O aplicativo inicia sem uma base fixa. O operador deve importar a primeira aba de um Excel que contenha os cabeçalhos: `ID`, `FEID`, `UUID`, `Season year`, `Entity name`, `Trial type`, `Location`, `Row`, `Column`, `Entry code`, `Block`, `(OBS) Name`, `GID`, `(GER) Name`. `Initial plot` e `Final plot` são opcionais, inclusive seus valores. Colunas de peso são opcionais: `Plot weight` ou o alias legado `PW`, e `Seed weight` ou o alias legado `SEED.W`. Valores numéricos existentes, incluindo zero, são retomados como parcelas já pesadas para a variável selecionada na sessão. `Site` é opcional e representa apenas o local do ensaio; `Lot site` é um metadado operacional da pesagem e nunca é sugerido a partir do Excel. A planilha importada fica salva no navegador do equipamento. `Initial plot` e `Final plot` permanecem no template por compatibilidade, mas não são usados nos cálculos de progresso.

## Arquivo único e nova sessão

No Dashboard, `Export filtered Excel/CSV` inclui todas as linhas dos filtros atuais, independentemente da página. A saída começa com `Lot name`, `Lot site`, `Storage`, `Lot location`, `First order`, `Second order` e `Lot weight`. A coluna `Lot weight` é uma cópia da variável da sessão, não uma segunda pesagem. Campos vazios não bloqueiam a exportação; nomes repetidos geram aviso, sem excluir UUIDs.

Excel conserva os pesos numéricos exatos com formato decimal; CSV usa UTF-8, ponto decimal e arredondamento selecionado. Ambos podem ser reimportados, incluindo metadados de parcelas pendentes. Em uma nova sessão, um arquivo com apenas uma variável seleciona essa variável automaticamente; com ambas, prevalece a seleção do operador. Valores descritivos podem estar vazios, mas FEID e UUID precisam ser preenchidos e únicos.

`Lot site` oferece Passo Fundo e Chacabuco; `Storage` depende do local. Valores legados são identificados como `(existing)` e preservados. O local do ensaio nunca preenche automaticamente os dados operacionais.

`Start New Weighing` preserva a sessão anterior e abre uma área vazia para importar outra base. Alterações não confirmadas exigem salvar, descartar ou cancelar. Cancelar o arquivo ou recarregar mantém a área vazia; uma sessão nova só é criada após importação válida.

## Validação local

- `node scripts/validate-lot-feature.mjs`: ciclos reais Excel/CSV, precisão, metadados, catálogo e conflitos.
- `node scripts/serve-static.mjs 4197`: versão estática em `http://127.0.0.1:4197/`.
- `http://127.0.0.1:4197/test-fixtures/session-store.html`: testes reais de IndexedDB em banco temporário isolado.
- `node node_modules/typescript/bin/tsc --noEmit`, `npm run lint` e `npm run build`.

## Desenvolvimento

Requisitos: Node.js 22.13 ou superior.

```bash
npm ci
npm run db:generate
npm run build
npm run dev
```

O banco utiliza o binding D1 `DB`. Para a prévia local, aplique a migração gerada em `drizzle/` ao banco local do Wrangler antes de testar gravações.

## Estrutura principal

- `app/page.tsx`: fluxo operacional e painel de progresso;
- `app/api/weights/route.ts`: leitura e gravação dos pesos;
- `index.html` e `pages/`: versão estática publicada no GitHub Pages;
- `public/plot-import.js`: validação e transformação da planilha importada;
- `db/schema.ts`: tabela de pesagens;
- `drizzle/`: migrações do banco.
