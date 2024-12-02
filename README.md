# Envio de E-mails em Massa com Python

Este projeto consiste em um script Python para envio de e-mails em massa utilizando a biblioteca `smtplib` para conexões SMTP e `pandas` para manipulação de dados de uma planilha Excel contendo os endereços de e-mail.

## Requisitos

- Python 3.x
- Bibliotecas `smtplib`, `pandas`, `openpyxl`

## Configuração

Configurar o Gmail
1.1 Ativar a Verificação em Duas Etapas:
Acesse sua Conta Google: Minha Conta Google.

No painel esquerdo, clique em Segurança.

Em Como você faz login no Google, ative a *Verificação em duas etapas.

Siga as instruções para concluir a configuração.

1.2 Gerar a Senha de App:
Na mesma seção Segurança, procure por *Senhas de App.

Selecione Outro (personalizado) e digite um nome (exemplo: "Envio de E-mail").

Clique em Gerar. Uma senha de 16 caracteres será exibida.

Copie a senha gerada.
Substitua o valor de password no código pela senha de app gerada
