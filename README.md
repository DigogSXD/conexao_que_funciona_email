# 🤖 Automação de Consulta Corporativa (Selenium)

Este script automatiza a consulta de dados de colaboradores (especificamente a "Área" de lotação) em um sistema web interno, utilizando uma lista de e-mails fornecida via Excel.

Desenvolvido em **Python** utilizando **Selenium** para navegação web e **Pandas** para manipulação de dados.

## 📋 Funcionalidades

-   **Leitura de Dados:** Importa uma lista de e-mails de um arquivo Excel (`emails.xlsx`).
-   **Navegação Automatizada:** Acessa o portal interno e insere os filtros de busca automaticamente.
-   **Extração de Dados:** Captura a informação da "Área" do colaborador via XPath.
-   **Tratamento de Erros:** Identifica quando um usuário não é encontrado e registra no relatório.
-   **Exportação:** Gera uma nova planilha (`resultado_consulta.xlsx`) com os dados consolidados.
-   **Segurança:** Utiliza variáveis de ambiente para proteger a URL do sistema alvo.

## 🛠️ Tecnologias Utilizadas

* [Python 3.x](https://www.python.org/)
* [Selenium WebDriver](https://www.selenium.dev/)
* [Pandas](https://pandas.pydata.org/)
* [OpenPyXL](https://openpyxl.readthedocs.io/) (Engine Excel)
