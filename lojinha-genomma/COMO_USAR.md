# 🛍️ Lojinha Genomma Lab — Guia de Uso

## ✅ Pré-requisitos

- **Python 3.8+** instalado no computador
- Pacotes: `flask` e `openpyxl`

## 📦 Instalação (uma única vez)

Abra o Terminal (ou Prompt de Comando) dentro da pasta `lojinha-genomma` e execute:

```bash
pip install flask openpyxl
```

## ⚙️ Configurar o Email

1. Copie o arquivo `.env.example` e renomeie para `.env`
2. Abra o `.env` com um editor de texto e preencha:

```
SMTP_USER=seu_email@gmail.com
SMTP_PASS=xxxx xxxx xxxx xxxx
```

**Como gerar a Senha de App do Gmail:**
1. Acesse https://myaccount.google.com
2. Vá em Segurança → Verificação em 2 etapas (ative se necessário)
3. Depois vá em Segurança → Senhas de app
4. Gere uma senha para "Email / Windows" e cole no SMTP_PASS

## 🚀 Iniciar o Servidor

```bash
python app.py
```

Acesse no navegador: **http://localhost:3000**

## 📋 Como funciona

1. **Comprador** acessa a página e preenche nome + email corporativo
2. Escolhe o **produto** no menu suspenso (estoque atualizado em tempo real)
3. Clica em **Fazer Pedido**
4. Informa se é **Terceirizado ou Genomma**
   - Se Genomma: aparece o CNPJ para PIX + campo de comprovante
5. Clica em **Finalizar Pedido**
6. **Você recebe um email** com todos os dados e o comprovante anexado
7. O comprador vê a mensagem de confirmação

## 🔄 Atualização de Estoque

- O estoque é atualizado automaticamente no arquivo `data/stock.json` após cada pedido
- Se quiser resetar o estoque para o Excel original, basta apagar o arquivo `data/stock.json`

## 📁 Estrutura de Arquivos

```
lojinha-genomma/
├── app.py                        ← Servidor principal (execute este!)
├── .env                          ← Configurações (crie a partir do .env.example)
├── .env.example                  ← Modelo de configuração
├── public/
│   └── index.html               ← Interface visual
├── data/
│   ├── Estoque_Lojinha_jun26.xlsx  ← Planilha original
│   └── stock.json               ← Estoque dinâmico (gerado automaticamente)
└── uploads/                     ← Comprovantes recebidos
```

## ❓ Problemas comuns

| Problema | Solução |
|---|---|
| "Connection refused" | Execute `python app.py` primeiro |
| Email não chega | Verifique `.env` e use Senha de App (não senha normal) |
| Produto sem estoque | Apague `stock.json` para recarregar do Excel |
| Porta 3000 ocupada | Adicione `PORT=3001` no `.env` |
