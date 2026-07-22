# Automação de Cadastro de SMP

Automação desenvolvida em Python para realizar o cadastro de SMPs em um sistema web a partir de uma planilha Excel, reduzindo atividades manuais e aumentando a confiabilidade do processo.

## Objetivo

Eliminar um processo manual e repetitivo de cadastro de SMPs, realizando automaticamente:

- Leitura de uma planilha Excel;
- Identificação de veículos sem cadastro;
- Login no sistema;
- Preenchimento automático dos formulários;
- Configuração dos pontos da rota;
- Geração da rota dinâmica;
- Configuração da SMP Agendada;
- Tratamento de erros durante a execução.

## Tecnologias

- Python
- Selenium
- Pandas
- WebDriver Manager
- OpenPyXL

## Funcionalidades

- Leitura de dados em Excel.
- Identificação de registros pendentes.
- Login automático no sistema.
- Preenchimento automático de formulários.
- Geração de rotas.
- Tratamento de exceções e recuperação de sessão.

## Dependências

```
selenium
webdriver-manager
pandas
openpyxl
python-dotenv
```

---

## Segurança

Este projeto foi adaptado para publicação.

Foram removidos:

- Credenciais
- URLs internas
- Caminhos de rede
- Informações de clientes
- Dados operacionais
- Nomes da empresa

---

## Resultados

A automação possibilitou:

- redução significativa do tempo gasto no processo;
- eliminação de tarefas repetitivas;
- diminuição de erros de digitação;
- maior confiabilidade no cadastro;
- aumento da produtividade da equipe.

---

## Aprendizados

Durante o desenvolvimento foram aplicados conceitos como:

- Automação Web com Selenium;
- Manipulação de dados utilizando Pandas;
- Esperas explícitas (`WebDriverWait`);
- Tratamento de exceções;
- Estruturação de código em Python;
- Automação de processos (RPA).

---

## Autor

Desenvolvido por **Letícia Reis**.
