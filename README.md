# 📊 Automação de JSON para Excel

Este projeto consiste em uma **aplicação de automação** desenvolvida para **consumir grandes volumes de dados em formato JSON a partir de uma API REST** e convertê-los de forma eficiente em **planilhas Excel (.xlsx)**.

A solução foi criada a partir de uma **necessidade real**, onde a exportação manual de relatórios com aproximadamente **40 mil registros**, cada um representado por um objeto JSON, demandava muito tempo e processamento. Com este sistema, o processo foi significativamente otimizado, tornando a geração do relatório muito mais rápida e confiável.

---

## 🚀 Visão Geral

A aplicação permite que o usuário:

* Informe a **URL de uma API REST** que retorna dados em JSON
* Escolha o **local onde o arquivo Excel será salvo**
* Inicie a conversão automática dos dados

A partir disso, o sistema realiza uma **requisição HTTP do tipo GET**, processa e **normaliza os dados recebidos** e gera automaticamente um **arquivo Excel estruturado**, pronto para análise e uso.

Durante todo o processo, a interface exibe uma **barra de progresso com feedback em tempo real**, garantindo uma melhor experiência ao usuário, especialmente em operações que envolvem grande volume de dados.

---

## 🔧 Funcionalidades Principais

* Consumo de API REST via requisição HTTP GET
* Suporte a grandes volumes de dados (ex.: ~40 mil registros)
* Processamento e normalização de objetos JSON
* Conversão automática para planilha Excel (.xlsx)
* Seleção do diretório de salvamento do arquivo
* Barra de progresso com atualização em tempo real
* Tratamento de erros e mensagens amigáveis ao usuário

---

## 🎯 Problema Resolvido

Antes da criação desta aplicação, a exportação de relatórios completos exigia longos períodos de espera devido ao alto volume de dados retornados pela API, o que impactava diretamente a produtividade da equipe.

Este projeto foi **desenvolvido inicialmente para uso interno em uma empresa**, onde havia a necessidade de exportar aproximadamente **40 mil registros**, cada um representado por um objeto JSON.

Com a solução implementada, foi possível:

* Reduzir significativamente o tempo de exportação
* Automatizar um processo manual e repetitivo
* Garantir consistência e confiabilidade nos dados exportados
* Melhorar a produtividade no uso e análise das informações

Após a validação da solução, a aplicação foi **empacotada e distribuída como um executável**, permitindo que usuários finais utilizassem o sistema sem a necessidade de configuração de ambiente ou conhecimento técnico.

---

## 🧠 Objetivo do Projeto

O principal objetivo deste projeto é **automatizar e otimizar o processo de exportação de dados JSON para Excel**, aplicando conceitos como:

* Consumo eficiente de APIs REST
* Processamento de grandes volumes de dados
* Manipulação e transformação de dados estruturados
* Experiência do usuário com feedback visual
* Boas práticas de organização e automação

---

## 🛠️ Tecnologias Utilizadas

* **Python** – Linguagem principal da aplicação
* **PySide 6** – Interface gráfica desktop (GUI)
* **API REST** – Consumo de dados via HTTP
* **Pandas** – Processamento, normalização e exportação dos dados para Excel

---

## 📘 Como Utilizar o Sistema

### 1️⃣ Pré-requisitos

A aplicação foi distribuída em formato de **executável**, portanto:

* ✔️ **Não é necessário instalar Python**
* ✔️ **Não é necessário configurar ambiente**
* ✔️ O usuário final precisa apenas de um **sistema operacional compatível (Windows)**

> Caso utilize a versão em código-fonte, siga as instruções abaixo.

---

### 2️⃣ Utilizando o Executável

1. Abra o executável da aplicação
2. Informe a **URL da API REST** que retorna os dados em formato JSON
3. Selecione o **diretório onde o arquivo Excel será salvo**
4. Clique em **Iniciar Exportação**
5. Acompanhe o progresso através da **barra de carregamento em tempo real**
6. Ao final do processo, o arquivo `.xlsx` será gerado automaticamente

---

### 3️⃣ Executando via Código-Fonte (Modo Desenvolvedor)

#### Instalação das dependências

Certifique-se de ter o **Python 3.9 ou superior** instalado.

```bash
pip install pyside6 pandas requests
```

#### Executar a aplicação

#### Executar a aplicação

```bash
python nomeprojeto.py
```

---

## 📈 Observações Importantes

* O sistema foi projetado para **grandes volumes de dados** (ex.: ~40 mil registros)
* O tempo de execução pode variar de acordo com:

  * velocidade da API
  * quantidade de dados retornados
  * desempenho da máquina
* O processamento ocorre de forma otimizada para evitar travamentos da interface

---

**Autor:** [Mário Gomes](https://www.linkedin.com/in/m%C3%A1rio-gomes-7b59b71b9/)

