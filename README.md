# EBD · Chamada

App de chamada para Escola Bíblica Dominical. HTML/CSS/JS puro, hospedado no GitHub Pages, com persistência no Google Sheets via Google Apps Script.

## Funcionalidades

- Lista de alunos vinda da aba `Alunos` da planilha.
- 5 status por aluno: **P** (Presente), **A** (Atrasado), **J** (Justificado), **F** (Falta), **FH** (Falta no Horário).
- Datas pré-calculadas: todos os domingos do mês escolhido.
- Contadores por status (P / A / J / F / FH / Total).
- Histórico do mês em modal (tabela alunos × domingos + % de presença).
- Adicionar novo aluno direto pelo app.
- Exportar a chamada do mês como **PNG** (com legenda e cabeçalho da igreja).
- Salvamento idempotente (sobrescreve a coluna da data, não duplica).

## Estrutura

```
ebd-chamada/
├── index.html
├── styles.css
├── app.js
├── config.js            # URL do Apps Script (editar antes de publicar)
├── apps-script/
│   └── Code.gs          # Backend (Apps Script)
└── README.md
```

## Como rodar

### 1. Preparar a planilha do Google Sheets

ID da planilha já configurado em `apps-script/Code.gs`:

```
1-SvEEVvBPRvBXspTRMA5BxhPJEbmNZA3MsDYR_BCtLU
```

- Crie/garanta uma aba chamada **`Alunos`** com os nomes na coluna **A**, a partir da linha **2** (linha 1 pode ser cabeçalho).
- As abas mensais (`AAAA-MM`, ex.: `2026-05`) são criadas automaticamente pelo backend; você não precisa criá-las à mão.

### 2. Publicar o Apps Script

1. Abra <https://script.google.com> e crie um novo projeto.
2. Cole o conteúdo de `apps-script/Code.gs` no editor.
3. Salve.
4. Clique em **Implantar → Nova implantação** (Deploy → New deployment).
5. Tipo: **Aplicativo da Web** (Web app).
6. Configurações:
   - **Executar como:** Eu (seu e-mail).
   - **Quem tem acesso:** Qualquer pessoa (Anyone).
7. Implante. Autorize quando solicitado (a planilha pertence à mesma conta).
8. Copie a URL gerada — formato: `https://script.google.com/macros/s/AKfycb.../exec`.

> Sempre que alterar o `Code.gs`, faça **Implantar → Gerenciar implantações → Editar (lápis) → Nova versão**. A URL `/exec` permanece a mesma.

### 3. Configurar o frontend

Abra `config.js` e cole a URL no campo `APPS_SCRIPT_URL`:

```js
window.APP_CONFIG = {
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbXXXXXXXXXXXXXX/exec',
  CHURCH_NAME: 'EBD'
};
```

### 4. Publicar no GitHub Pages

```bash
git init
git add .
git commit -m "EBD chamada inicial"
git branch -M main
git remote add origin https://github.com/<usuario>/<repo>.git
git push -u origin main
```

No GitHub: **Settings → Pages → Build and deployment → Source: Deploy from a branch → main / root → Save**.

A URL pública aparece em poucos minutos: `https://<usuario>.github.io/<repo>/`.

## Por que não dá erro de CORS?

O frontend envia requisições POST para o Apps Script usando `Content-Type: text/plain`, que se enquadra como **simple request** segundo o CORS — o navegador **não** dispara o preflight `OPTIONS`. O Apps Script lê o corpo via `e.postData.contents` e responde com `ContentService` (que já manda `Access-Control-Allow-Origin: *` quando publicado como "anyone").

Tentar usar `application/json` exigiria preflight `OPTIONS`, e o Apps Script não responde `OPTIONS` → seria o erro de CORS típico. Por isso evite mudar o `Content-Type`.

## Como o backend organiza os dados

- **Aba `Alunos`**: fonte da verdade para a lista de alunos.
- **Aba mensal `AAAA-MM`**:
  - Linha 1: cabeçalhos. A1 = `Aluno`. B1, C1, ... = datas (`AAAA-MM-DD`).
  - Coluna A: nomes dos alunos.
  - Células: status (`P`, `A`, `J`, `F`, `FH` ou vazio).
- A cada `getMonth` / `saveAttendance`, o backend sincroniza novos alunos para a aba do mês (sem remover quem já está lá).
- `saveAttendance` localiza a coluna da data e a sobrescreve inteira; ou cria a coluna no fim, se ainda não existir.

## Endpoints (Apps Script)

Todos retornam JSON `{ ok: true, data }` ou `{ ok: false, error }`.

| `action`          | Parâmetros                                | Descrição                                 |
| ----------------- | ----------------------------------------- | ----------------------------------------- |
| `ping`            | —                                         | Teste de conectividade.                   |
| `getStudents`     | —                                         | Lista de alunos da aba `Alunos`.          |
| `getMonth`        | `month` (`AAAA-MM`)                       | Grid de presenças do mês.                 |
| `saveAttendance`  | `month`, `date` (`AAAA-MM-DD`), `attendance` (`{nome: status}`) | Salva a coluna do dia. |
| `addStudent`      | `name`                                    | Adiciona aluno na aba `Alunos`.           |

## Acessibilidade

- Alvos de toque com mínimo 44px.
- `aria-pressed` nos botões de status e nos chips de data.
- `aria-live` no contador e no toast.
- Estados de loading e erro comunicados via toast com `role="alert"`.

## Próximos passos (ideias)

- Sincronização offline com cache em `localStorage`.
- Edição/remoção de alunos.
- Multi-classes (várias turmas na mesma planilha).
- Importar lista de alunos a partir de CSV.
