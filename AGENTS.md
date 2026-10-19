# AGENTS.md

Este documento descreve as regras, arquitetura e comandos deste projeto para orientar o trabalho de Agentes de IA.

## Estrutura do Projeto
Este é um projeto **Google Apps Script (Backend Puro)** configurado com TypeScript e empacotado via Vite.

## Comandos de Build
- `pnpm run build`: Compila o código TypeScript em um arquivo `.gs` pronto para nuvem.
- `pnpm run deploy`: Executa o build e utiliza o Clasp para fazer upload para o Google Apps Script.
- `pnpm run types`: Gera os arquivos de tipagem (`.d.ts`) usando o pacote `clasp-types`.

## Regras de Código
- Use `pnpm` para gerenciar dependências.
- Funções que precisam ser acessadas externamente ou pelo menu do Google Docs/Sheets devem ser injetadas globalmente.
- Mantenha o arquivo `appsscript.json` atualizado ao incluir novas permissões OAuth ou bibliotecas (Services).

