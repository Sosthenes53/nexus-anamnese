# NEXUS CLIN — Automação Clínica (Google Apps Script)

Sistema de automação para recepção de anamneses, cadastro de pacientes no Feegow,
geração de requisições laboratoriais em PDF e envio de e-mails personalizados via
Google Apps Script.

---

## Pré-requisitos

- Conta Google: `nexusclinpb@gmail.com` *(o script deve ser criado e autorizado por ela)*
- Token de API do Feegow
- Chave de API da Anthropic (Claude)
- Formulário HTML de anamnese (`index.html`)

---

## 1. Autenticação

> ⚠️ **Execute este passo obrigatoriamente logado como `nexusclinpb@gmail.com`.**
> Todos os e-mails saem automaticamente dessa conta via `GmailApp`.

1. Acesse [script.google.com](https://script.google.com) **logado com `nexusclinpb@gmail.com`**
2. Clique em **"Novo projeto"**
3. Renomeie para `NEXUS CLIN Automação`
4. Apague todo o conteúdo do editor e cole o conteúdo completo de `Code.gs`
5. Clique em **"Salvar"** (`Ctrl+S`)
6. No menu superior, clique em **"Executar"** → selecione a função `doGet`
7. Uma janela de **autorização de permissões** será exibida:
   - Clique em **"Revisar permissões"**
   - Selecione a conta `nexusclinpb@gmail.com`
   - Se aparecer "App não verificado", clique em **"Avançado"** → **"Acessar NEXUS CLIN Automação (não seguro)"**
   - Autorize **todas** as permissões solicitadas:
     - ✅ Gmail (enviar e-mails)
     - ✅ Google Drive (criar/deletar arquivos)
     - ✅ Google Docs (criar documentos)
     - ✅ Conexões externas (Feegow API, Anthropic API)
8. Verifique o log em **"Execuções"** — deve mostrar a execução concluída sem erros

---

## 2. Constantes a substituir

Edite as linhas no topo do `Code.gs`:

### Token Feegow

```javascript
const FEEGOW_TOKEN = 'SEU_TOKEN_FEEGOW';
```

**Como obter:**
1. Acesse [app.feegow.com](https://app.feegow.com)
2. Vá em **Configurações → API**
3. Clique em **"Liberar token para o usuário master"**
4. Copie o token gerado e substitua `SEU_TOKEN_FEEGOW`

### Chave Anthropic

```javascript
const ANTHROPIC_KEY = 'SUA_CHAVE_ANTHROPIC';
```

**Como obter:**
1. Acesse [console.anthropic.com](https://console.anthropic.com)
2. Vá em **API Keys → Create Key**
3. Dê um nome como `nexus-clin-apps-script`
4. Copie a chave e substitua `SUA_CHAVE_ANTHROPIC`

> ⚠️ **Segurança:** Nunca compartilhe essas chaves. O script fica privado na sua conta Google.

---

## 3. Publicar como Web App

1. No editor, clique em **"Implantar"** → **"Nova implantação"**
2. Clique no ícone de engrenagem ⚙️ e selecione **"Web App"**
3. Configure:
   - **Descrição:** `NEXUS CLIN v1.0`
   - **Executar como:** `Eu (nexusclinpb@gmail.com)`
   - **Quem tem acesso:** `Qualquer pessoa`
4. Clique em **"Implantar"**
5. **Copie a URL** exibida — ela tem o formato:
   ```
   https://script.google.com/macros/s/XXXXXXXXXXXXXXXXX/exec
   ```
6. Para testar, abra a URL no navegador — deve exibir a tela de status verde ✅

> **Atenção:** Sempre que editar o `Code.gs`, você deve criar uma **nova implantação**
> (ou selecionar "Gerenciar implantações" → editar → "Nova versão") para que as
> alterações entrem em vigor.

---

## 4. Atualizar o formulário

No arquivo `index.html`, localize a constante:

```javascript
const SCRIPT_URL = 'COLE_AQUI_A_URL_DO_SEU_ENDPOINT';
```

Substitua pela URL copiada no passo anterior:

```javascript
const SCRIPT_URL = 'https://script.google.com/macros/s/XXXXXXXXXXXXXXXXX/exec';
```

---

## 5. Testar via curl

```bash
curl -L -X POST "SUA_URL_AQUI" \
  -H "Content-Type: application/json" \
  -d '{
    "nome": "Maria Teste Silva",
    "cpf": "000.000.000-00",
    "data_nascimento": "1990-05-15",
    "celular": "(83) 9 9999-9999",
    "email": "teste@email.com",
    "sexo": "Feminino",
    "cep": "58000-000",
    "endereco": "Rua Teste",
    "numero_endereco": "100",
    "cidade": "Livramento",
    "estado": "PB",
    "objetivo": "Perda de peso e mais energia",
    "nivel_energia": "4",
    "qualidade_sono": "5",
    "sint_gerais": "Cansaço | Queda de cabelo",
    "sint_emocionais": "Ansiedade"
  }'
```

---

## 6. Resultado esperado do teste

- ✅ Resposta JSON: `{ "success": true, "paciente_id": "...", "emails_enviados": true }`
- ✅ E-mail chega em `teste@email.com` com PDF anexo (Painel Metabólico NEXUS)
- ✅ Alerta interno chega em `sosthenes53@gmail.com` com dados completos do paciente
- ✅ Paciente cadastrado no Feegow (ou ID existente recuperado se CPF duplicado)
- ✅ Log completo disponível em **Apps Script → Execuções**

---

## Fluxo de execução

```
Formulário HTML  →  POST /exec
                        │
                        ├─ 1. cadastrarNoFeegow()   → Feegow API
                        ├─ 2. gerarResumoClaude()   → Anthropic API
                        ├─ 3. gerarPDFRequisicao()  → Google Docs → PDF
                        └─ 4. enviarEmails()        → GmailApp
                                │
                                ├─ E-mail paciente (com PDF anexo)
                                └─ Alerta interno → sosthenes53@gmail.com
```

---

## Tratamento de erros

| Situação | Comportamento |
|----------|---------------|
| CPF já cadastrado no Feegow | Recupera ID existente e continua normalmente |
| Falha na API Claude | Usa texto de boas-vindas padrão (fallback) |
| Falha no alerta interno | Log silencioso, não interrompe o fluxo |
| Falha no e-mail do paciente | Propaga o erro (crítico) |

---

## Snippet JS para o formulário (submitForm)

Cole este código no `index.html` substituindo a função `submitForm` existente:

```javascript
const SCRIPT_URL = 'COLE_AQUI_A_URL';

async function submitForm(event) {
  event?.preventDefault?.();
  if (isSubmitting) return;
  if (!validateSection(currentSection)) return;
  isSubmitting = true;
  setSubmitState(true);
  const data = collectData();
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const json = await res.json();
    if (json.success) {
      localStorage.removeItem(STORAGE_KEY);
      document.querySelectorAll('.section')
        .forEach(s => s.classList.remove('active'));
      document.getElementById('successScreen')
        .classList.add('show');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } else {
      showSubmitFeedback(
        'Erro ao enviar: ' + (json.error || 'tente novamente.')
      );
    }
  } catch(err) {
    showSubmitFeedback(
      'Não foi possível enviar. Verifique sua conexão.'
    );
  } finally {
    isSubmitting = false;
    setSubmitState(false);
  }
}
```

---

*NEXUS CLIN — Centro de Performance Metabólica & Longevidade | Livramento-PB*
