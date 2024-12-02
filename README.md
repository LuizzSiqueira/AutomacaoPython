# Envio de E-mails em Massa com Python

Este projeto consiste em um script Python para envio de e-mails em massa utilizando a biblioteca `smtplib` para conexões SMTP e `pandas` para manipulação de dados de uma planilha Excel contendo os endereços de e-mail.

## Requisitos

- Python 3.x
- Bibliotecas `smtplib`, `pandas`, `openpyxl`

## Configuração

### Configurar o Gmail

#### 1.1 Ativar a Verificação em Duas Etapas:
1. Acesse sua Conta Google: [Minha Conta Google](https://myaccount.google.com/).
2. No painel esquerdo, clique em **Segurança**.
3. Em **Como você faz login no Google**, ative a **Verificação em duas etapas**.
4. Siga as instruções para concluir a configuração.

#### 1.2 Gerar a Senha de App:
1. Na mesma seção **Segurança**, procure por **Senhas de App**.
2. Selecione **Outro (personalizado)** e digite um nome (exemplo: "Envio de E-mail").
3. Clique em **Gerar**. Uma senha de 16 caracteres será exibida.
4. Copie a senha gerada.

### Configurar o Script Python

Substitua o valor de `password` no código pela senha de app gerada:
```python
password = 'SUA_SENHA_DE_APP'  # Coloque a senha gerada aqui
