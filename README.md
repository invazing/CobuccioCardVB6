
# 💳 CobuccioCard - Grupo Adriano Cobuccio

Sistema completo para **cadastro e gerenciamento de transações de cartão de crédito**, desenvolvido em **VB.NET CORE 8.0** com **SQL Server**.

---

## ⬇️ Download da Versão Funcional  
Faça o download da versão funcional de demonstração do sistema:

📦 O pacote inclui:

- Arquivos .exe compilados
- Script SQL completo

---

## ⚙️ Funcionalidades Principais

### ✅ CRUD de Transações
- Criar, consultar, editar (exceto transações aprovadas) e excluir transações.
- Filtros inteligentes por número do cartão, data, valor e status.
- Exibição em `DataGrid` com paginação otimizada.
- Validação de campos:
  - `Id_Transacao`: gerado automaticamente
  - `Numero_Cartao`: 16 dígitos
  - `Valor_Transacao`: decimal positivo
  - `Data_Transacao`: data e hora
  - `Descricao`: até 255 caracteres
  - `Status_Transacao`: Aprovada, Pendente ou Cancelada

### 🛡️ Tratamento de Erros e Logs
- Mensagens claras e amigáveis para o usuário final.
- Log de exceções detalhado salvo automaticamente em: `%programdata%/Cobucciocard/cobucciocard.log`

### 🗃️ Integração com SQL Server
- **Stored Procedures**: totalização de transações por período e status.
- **Scalar Function**: categorização de valores por tipo.
- **TVFs (Table-Valued Functions)**: retorno de transações categorizadas.
- **Views**: consolidação de dados para relatórios financeiros.

### 📤 Exportação para Excel
- Geração de planilhas Excel com os dados das transações em períodos filtrados.

---

## 🧩 Requisitos

- Microsoft Visual Basic 6 
- SQL Server 2016 ou superior

---

## 🚀 Como Executar

### 1. Clonar o repositório

```bash
git clone https://github.com/invazing/CobuccioCardVB6
```

## 2. Criar e configurar o banco de dados

Restaure o banco de dados o script `Backup_Full_CobuccioCard.bak` para criar:

- Tabelas
- Stored Procedures
- Functions
- Views
- Dados de exemplo

## 3. Ajustar a string de conexão

Na classe ModConnection altere a String de conexão ConnectionString `"Server=;Database=;User Id=;Password=;"`, configure a string de conexão para seu ambiente SQL Server.

## 4. Rodar o projeto

Abra a solução no Microsoft Visual Basic 6 / 98, compile e pressione F5 para executar.

---

## 📁 **Estrutura do Projeto**
- `/Data`         → Camada de acesso a dados e conexão com SQL Server
- `/Forms`        → Interfaces e lógicas dos formulários (login, CRUD, consultas)
- `/Utils`       → Utilitários e helpers (ex: log de erros)
- `/Models`      → Modelo / Relatórios gerados

---

## 🧪 Exemplo de Uso

- Cadastrar e consultar transações com filtros dinâmicos.
- Exportar dados para Excel com um clique.
- Pode gerar dados de input no banco de dados via `GERAR DADOS` no `MENU` do sistema.

<img width="546" height="324" alt="image" src="https://github.com/user-attachments/assets/989154f3-cc3a-4b46-b1c3-6feddd263ab4" />

---

## 📄 **Licença**

Este projeto foi desenvolvido exclusivamente para fins de avaliação técnica.  
Distribuição comercial ou uso externo não são permitidos.

