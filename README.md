# Gerador de Certificados — Mindworks

Sistema web para geração em lote de certificados PDF a partir de uma lista de participantes (CSV) e um modelo PowerPoint (PPTX). Roda no navegador, sem necessidade de instalar nada na máquina do usuário.

---

## Stack

| Camada | Tecnologia |
|--------|-----------|
| Backend | Python 3.11 + Flask |
| Frontend | HTML / CSS / JS puro |
| Conversão PDF | LibreOffice headless |
| Manipulação PPTX | python-pptx |
| Envio de e-mail | smtplib (SMTP interno, sem autenticação) |
| Deploy | Docker |

---

## Estrutura do projeto

```
automacao-certificados/
├── app_web.py                  # Servidor Flask — rotas e controle de jobs
├── generate_certificates.py    # Lógica de negócio (CSV, PPTX, PDF, e-mail)
├── templates/
│   └── index.html              # Interface web (single-page)
├── static/
│   ├── style.css               # Design Mindworks
│   └── app.js                  # Formulário, polling de progresso, UI
├── input_exemplo/
│   ├── participantes.CSV       # Exemplo de lista de participantes
│   └── modelo_certificado.pptx # Template de certificado de exemplo
├── docs/
│   ├── documentacao-tecnica.docx
│   └── documentacao-usuario.docx
├── Dockerfile
├── .dockerignore
├── .gitignore
└── requirements.txt
```

Diretórios criados em runtime (ignorados pelo git):

```
uploads/    # arquivos enviados pelo usuário por sessão
outputs/    # certificados gerados por sessão
logs/       # log de execução
```

---

## Formato do CSV de entrada

```
Nome;Email;Função
João Silva;joao.silva@mindworks.com.br;Analista
Maria Oliveira;maria@empresa.com;Gerente
```

- Delimitador: `;`
- Linha 0 é cabeçalho (ignorada)
- Coluna 0: nome · Coluna 1: e-mail · Coluna 2+: ignoradas
- Encodings aceitos: UTF-8 BOM, cp1252, UTF-8

---

## Placeholders no template PPTX

O sistema substitui os seguintes textos nas caixas de texto do slide:

| Placeholder | Substituído por |
|-------------|----------------|
| `[NOME DO PARTICIPANTE]` | Nome do participante |
| `[NOME]` | Nome do participante |
| `NOME DO PARTICIPANTE` | Nome do participante |
| `[EVENTO]` | Nome do evento (informado no formulário) |
| `[MINISTRANTE]` | Nome do ministrante (informado no formulário) |

> A substituição ocorre no nível de *run* (fragmento de texto dentro de um parágrafo). Se o placeholder estiver quebrado entre runs diferentes no PowerPoint, o template deve ser ajustado.

---

## API

O frontend não recarrega a página — toda comunicação é via fetch/polling.

| Método | Rota | Descrição |
|--------|------|-----------|
| `GET` | `/` | Serve `index.html` |
| `POST` | `/gerar` | Recebe `multipart/form-data` (csv, pptx, event_name, instructor_name); inicia geração em background; retorna `{"session_id"}` |
| `GET` | `/status/<id>` | Progresso e mensagens da geração (consultado pelo JS a cada 1 s) |
| `GET` | `/download/<id>` | Serve ZIP com todos os PDFs gerados |
| `POST` | `/enviar/<id>` | Inicia envio de e-mails em background; retorna `{"email_session_id"}` |
| `GET` | `/email-status/<id>` | Progresso do envio de e-mails (consultado pelo JS a cada 1 s) |

---

## Configuração de e-mail

| Parâmetro | Valor |
|-----------|-------|
| SMTP host | `10.181.30.1` |
| SMTP port | `25` |
| Autenticação | Nenhuma |
| Remetente | `automacao@mindworks.com.br` |
| Domínio permitido | `@mindworks.com.br` |

Participantes com e-mail em outro domínio são ignorados automaticamente no envio.

---

## Deploy

*(A definir)*

---

## Limitações conhecidas

- **Estado em memória:** jobs são perdidos se o processo reiniciar durante uma geração (aceitável para uso mensal).
- **Sem autenticação:** não expor à internet sem proteção adicional (VPN ou reverse proxy com auth básica).
- **Fidelidade do PDF:** conversão via LibreOffice pode diferir levemente da renderização do PowerPoint.
