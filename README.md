Desafio VB6 - Sistema de Transações
---

Tecnologias utilizadas

- Visual Basic 6 (VB6)
- Microsoft SQL Server
- ADODB (conexão com banco de dados)
- Microsoft DataGrid
- Exportação para CSV

---

Funcionalidades

- Cadastro de transações com:
  - Número do Cartão (16 dígitos)
  - Valor
  - Data da Transação
  - Descrição
  - Status (`Aprovada`, `Pendente`, `Cancelada`)
- Filtros por:
  - Número do cartão
  - Valor
  - Data
  - Status
- Edição e exclusão de transações
- Impede edição de transações `Aprovadas`
- Exportação dos dados filtrados
- Validações para campos obrigatórios

---
Estrutura do banco de dados

A estrutura do banco é criada automaticamente via o script:

```bash
banco/criar_banco_completo.sql
