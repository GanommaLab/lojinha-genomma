# Contexto do Projeto — Lojinha Interna Genomma Lab

> Arquivo de memoria gerado automaticamente. Atualizar sempre que houver mudancas relevantes.

---

## Visao Geral

Sistema de loja interna para funcionarios da Genomma Lab comprarem produtos da empresa com desconto. Backend em Python/Flask hospedado no Render, frontend em HTML/JS estatico.

---

## Repositorio GitHub

- Organizacao: GanommaLab
- Repo: lojinha-genomma
- URL: https://github.com/GanommaLab/lojinha-genomma
- Branch principal: main

### Estrutura de pastas

```
lojinha-genomma/
├── backups/
│   ├── orders.json            ← backup de pedidos
│   └── stock.json             ← backup de estoque
└── lojinha-genomma/
    ├── app.py                 ← backend Flask principal
    ├── render.yaml
    ├── requirements.txt
    ├── COMO_USAR.md
    ├── Estoque Lojinha jun26.xlsx
    └── public/                ← frontend estatico
```

---

## Stack Tecnica

- Backend: Python 3 / Flask
- Hosting: Render (free tier) — https://lojinha-genomma.onrender.com
- Armazenamento em producao: /tmp/lojinha/ (volatil)
- Backup persistente: GitHub API
- Notificacoes: ntfy.sh + Microsoft Teams (webhook de canal via Power Automate)
- Email: smtplib (Gmail)

---

## Configuracoes do App

| Variavel | Valor |
|---|---|
| SMTP_HOST | smtp.gmail.com |
| SMTP_PORT | 587 |
| DEST_EMAIL | maycon.silva@contractor.genommalab.com |
| PORT | 3000 |
| ADMIN_USER | admin |
| ADMIN_PASS | 5827 |
| GITHUB_OWNER | GanommaLab |
| GITHUB_REPO | lojinha-genomma |
| GITHUB_BRANCH | main |
| TEAMS_WEBHOOK_URL | webhook do Workflow "Pedidos" no canal Teams da Lojinha Genomma (Power Automate). Ver `app.py` linha ~56 para a URL completa. |

---

## Notificacoes — ntfy.sh

Topico: genomma-lojinha-gl2024
URL: https://ntfy.sh/genomma-lojinha-gl2024

| Evento | Funcao | Tags |
|---|---|---|
| Novo pedido | _ntfy_order | shopping,tada / high |
| Estoque critico | _ntfy_stock_alert | warning,package / high |
| Relatorio diario | _send_daily_report | bar_chart / default |

Limite estoque critico: STOCK_LOW_THRESHOLD = 5 unidades
Relatorio diario: enviado as 08h00 pelo _daily_report_scheduler

---

## Notificacoes — Microsoft Teams

Adicionado em 29/06/2026 (commit `991498c`, apos correcao do incidente abaixo).

- Canal: "Pedidos", dentro da equipe "Lojinha Genomma" no Teams
- Mecanismo: Workflow do Power Automate (gatilho "Quando uma solicitacao HTTP for recebida"), nao e o conector legado "Incoming Webhook"
- Funcao no backend: `_send_teams_notification(order)` em `app.py`
- Disparada via `threading.Thread(..., daemon=True).start()` em dois pontos de `api_order()`: no fluxo de pedido com multiplos itens e no fluxo legado de item unico
- Formato da mensagem: Adaptive Card (schema 1.4) com nome, e-mail, vinculo, data, FactSet de produtos e valor total
- Falha no envio (ex.: webhook fora do ar) e tratada com try/except e log.warning — nunca derruba o pedido
- Variavel `TEAMS_WEBHOOK_URL` pode ser sobrescrita por env var no Render (`os.getenv('TEAMS_WEBHOOK_URL', TEAMS_WEBHOOK_URL)`), mas hoje usa o valor hardcoded no codigo

---

## Relatorio de Vendas (painel admin)

Adicionado no commit `264ff1a` ("feat: dashboard de relatorio de vendas no admin").

- Acessivel pelo botao "Relatorio" dentro de `/admin` (precisa login ADMIN_USER/ADMIN_PASS)
- Filtro por intervalo de datas (De/Ate) — usar para ver vendas por mes especifico
- Mostra: total de pedidos, valor total, ticket medio, unidades vendidas, % entregues, pendentes, top produtos por volume e por valor
- E a forma recomendada de visualizar pedidos agrupados por periodo — as notificacoes do Teams sao so alertas instantaneos e nao foram desenhadas para isso (ficam acumuladas cronologicamente no canal, como qualquer chat)

---

## Formato dos dados

### orders.json (lista)

```json
[{
  "id": "20240629143022123",
  "data_hora": "29/06/2024 14:30:22",
  "nome": "Nome do Funcionario",
  "email": "funcionario@genommalab.com",
  "tipo": "tipo_comprador",
  "items": [{"produto_code":"1008800","produto_name":"Produto","qty":2,"price":20.05,"subtotal":40.10}],
  "valor_total": 40.10,
  "comprovante": "arquivo.pdf",
  "status": "pendente"
}]
```

Status: pendente / entregue / concluido

### stock.json (dicionario)

```json
{
  "1008800": {"code":"1008800","name":"Asepxia Sabonete","stock":19,"price":20.05}
}
```

---

## Historico de mudancas

| Data | Commit | Descricao |
|---|---|---|
| 29/06/2026 | 991498c | Fix: reconstroi app.py a partir do ultimo commit estavel (264ff1a) e reaplica so a notificacao Teams, corrigindo o SyntaxError que travava o deploy |
| 29/06/2026 | 0c7d302 | (QUEBRADO) Notificacao Teams — commit subiu com sucesso mas continha SyntaxError por corrupcao na montagem manual do arquivo; deploy no Render falhou e ficou rodando a versao anterior ate a correcao acima |
| 29/06/2026 | 264ff1a | Dashboard de relatorio de vendas no admin |
| 29/06/2026 | 8464c70 | Fix: corrige SyntaxError no admin que impedia exibir pedidos |
| 29/06/2026 | ebb0f97 | Nota fiscal no admin |
| 29/06/2026 | acc59e4 | Adiciona alertas de estoque, pedido e relatorio diario |
| 29/06/2026 | 67ae689 | Backup de orders.json |
| 28/06/2026 | dcd17ba | Fix: restore app.py with ntfy instant order notification |

---

## Ambiente do Usuario

- Usuario: MAYCON
- Email pessoal: diasmaycon159@gmail.com
- Email corporativo: maycon.silva@contractor.genommalab.com
- SO: Windows x64
- Browser com senhas salvas: Browser 1 (deviceId: 4e4cca04-5416-4734-8ee4-d2093fa62f95)
- Browser secundario: Browser 2 (deviceId: 60e43fda-109d-4d54-9fbb-a7e6fd379586)

---

## Problemas conhecidos

1. GitHub API bloqueada no cloud Anthropic — usar sempre o browser para operacoes no GitHub
2. raw.githubusercontent.com funciona no browser mas nao no bash cloud
3. Render free tier dorme apos inatividade (cold start lento)
4. Arquivos em /tmp apagados a cada redeploy — backup GitHub recupera automaticamente
5. SMTP nao configurado — emails nao funcionam ate SMTP_USER/PASS serem definidos no Render
6. NUNCA editar app.py inteiro via chunking manual em sessionStorage (string-by-string) — um erro de fronteira de chunk corrompeu o arquivo silenciosamente (commit `0c7d302`), o commit no GitHub foi aceito mas o deploy no Render falhou por SyntaxError, e a producao ficou rodando a versao antiga sem avisar ninguem. Tecnica correta: buscar o conteudo-base com `fetch()` direto no browser (que tem acesso livre a internet, diferente do bash) e fazer substituicoes de string pontuais e ancoradas, validando o resultado (contagem de linhas, chaves balanceadas etc.) antes de commitar.
7. Sempre que alterar `app.py` em producao, conferir no Render (Dashboard → Events) se o deploy ficou "live" e nao "failed" antes de avisar o usuario que esta pronto

---

## Pendencias

- Configurar SMTP_USER e SMTP_PASS no Render para ativar emails
- Configurar GITHUB_TOKEN no Render para backup em producao
- Considerar cron externo (cron-job.org) para garantir relatorio diario se o servidor dormir
