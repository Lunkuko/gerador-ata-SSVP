# ✝️ Gerador de Ata Digital - SSVP

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-App-red)
![Status](https://img.shields.io/badge/Status-Concluído-success)

Este projeto é uma aplicação web desenvolvida para modernizar e facilitar a gestão das Conferências da **Sociedade de São Vicente de Paulo (SSVP)**. 

O sistema automatiza a redação das atas, realiza cálculos financeiros, controla a frequência dos membros e gera documentos oficiais (PDF e Word) prontos para impressão e assinatura.

---

## 🚀 Funcionalidades

### 🔐 Segurança e Acesso
- **Login Seguro:** Sistema de autenticação com níveis de acesso (Admin e Editor).
- **Gestão de Usuários:** Painel administrativo para criar novos usuários e senhas.
- **Proteção de Dados:** Senhas armazenadas com criptografia (Hash) no banco de dados.

### 📝 Gestão de Atas
- **Preenchimento Automático:** Carrega dados da última ata (saldo anterior, número da ata).
- **Chamada Inteligente:** Lista de presença e justificativas de ausência integradas.
- **Financeiro Automático:** Calcula o saldo final com base nas receitas, despesas e décima.
- **Histórico e Correção:** Permite buscar atas antigas e realizar correções/atualizações.

### 🖨️ Geração de Documentos
- **PDF Profissional:** Gera ata em PDF com texto justificado e lauda de assinaturas (linhas em branco para todos os presentes).
- **Word Editável:** Gera arquivo `.docx` caso seja necessário algum ajuste manual posterior.

---

## 🛠️ Tecnologias Utilizadas

- **[Streamlit](https://streamlit.io/):** Interface web interativa.
- **[Google Sheets API](https://developers.google.com/sheets/api):** Banco de dados na nuvem (gratuito e acessível).
- **[Streamlit Authenticator](https://github.com/mkhorasani/Streamlit-Authenticator):** Gestão de segurança e cookies.
- **[FPDF2](https://pyfpdf.github.io/fpdf2/):** Geração de relatórios PDF.
- **[Python-Docx](https://python-docx.readthedocs.io/):** Geração de documentos Word.

---

## 🗂️ Estrutura do Banco de Dados (Google Sheets)

Para que o sistema funcione, sua planilha no Google deve conter as seguintes abas (respeitando as maiúsculas/minúsculas):

| Aba | Colunas Necessárias | Descrição |
| :--- | :--- | :--- |
| **Config** | `Chave`, `Valor` | Configurações gerais (Nome da conferência, Última ata, Cidade, etc). |
| **Membros** | `Nome` | Lista de nomes para a chamada. |
| **Anos** | `Ano` | Lista de Anos Temáticos para seleção. |
| **Usuarios** | `username`, `name`, `password`, `role` | Credenciais de acesso. `role` pode ser 'admin' ou 'editor'. |
| **Historico** | `Numero`, `Data`, `Presidente`, `Secretario`, `Saldo`, ... | Armazena todas as atas geradas. |

---

## ⚙️ Instalação e Execução Local

### 1. Pré-requisitos
Certifique-se de ter o [Python](https://www.python.org/) instalado.

### 2. Clonar o Repositório
```bash
git clone [https://github.com/seu-usuario/gerador-ata-ssvp.git](https://github.com/seu-usuario/gerador-ata-ssvp.git)
cd gerador-ata-ssvp