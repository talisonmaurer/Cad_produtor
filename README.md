# Cadastro de Produtor e Declaração de Rebanho

Aplicativo em Google Apps Script para cadastrar usuários, emitir, consultar, cancelar e excluir declarações complementares de rebanho, com geração de PDF.

## Recursos

- login próprio com senha armazenada como hash com salt;
- cadastro pendente e aprovação administrativa;
- permissões separadas para emitir, cancelar e excluir;
- recuperação de senha por e-mail;
- numeração anual automática protegida contra concorrência;
- geração e compartilhamento controlado de PDFs;
- histórico de usuários e operações;
- CPF/CNPJ validado no navegador e no servidor.

## Instalação

1. Crie uma planilha e as abas utilizadas em `CONFIG`.
2. Crie uma pasta no Google Drive para os PDFs.
3. Crie um projeto no Google Apps Script com `Code.gs`, `Index.html` e `Print.html`.
4. Troque os IDs e e-mails de exemplo em `CONFIG` pelos recursos da nova instalação.
5. Configure a planilha conforme os cabeçalhos esperados pelo código.
6. Autorize o projeto e implante-o como aplicativo da Web executado pelo proprietário.

## Segurança e privacidade

- não publique IDs do ambiente real, cópias da planilha, PDFs ou cadastros;
- mantenha a pasta de PDFs restrita aos usuários autorizados;
- nunca grave senhas em texto simples;
- cumpra a LGPD e as regras locais de retenção de dados.

Este repositório contém somente uma versão sanitizada do código. Ele não está conectado ao aplicativo municipal em produção.

## Licença

MIT. Consulte [LICENSE](LICENSE).

