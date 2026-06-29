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
- Hosting: Render (free tier)
- Armazenamento em producao: /tmp/lojinha/ (volatil)
- Backup persistente: GitHub API
- Notificacoes: ntfy.sh
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

---

## Pendencias

- Configurar SMTP_USER e SMTP_PASS no Render para ativar emails
- Configurar GITHUB_TOKEN no Render para backup em producao
- Verificar se notificacoes ntfy chegam (testar com pedido de teste)
- Considerar cron externo (cron-job.org) para garantir relatorio diario se o servidor dormir
