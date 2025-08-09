# 🚀 Preenchedor Automático de Planilhas Excel

Este é um programa em **Java** desenvolvido para automatizar o preenchimento de formulários web usando dados extraídos diretamente de uma planilha Excel. Criado para acelerar tarefas repetitivas, o sistema utiliza o navegador **Google Chrome** para realizar a automação de forma eficiente.

⚠️ **Importante:** Este projeto é um modelo base. Para usá-lo em seu cenário específico, será necessário adaptar o código-fonte, conforme explicado na seção de customização.

---

## ✅ Pré-requisitos

Antes de rodar o programa, certifique-se de que seu computador atenda aos seguintes requisitos:

- **Sistema Operacional:** Windows  
- **Google Chrome:** Instalado no local padrão do sistema  
- **Java Development Kit (JDK):** Versão 17 ou superior

---

## ⚙️ Como Usar

### 1. Baixar o Programa

O arquivo executável `.jar` está disponível no diretório do projeto:

out/artifacts/automocao_planilha_jar/
Baixe o arquivo `automocao_planilha.jar` dessa pasta.

### 2. Preparar o Ambiente

Organize os arquivos conforme a estrutura abaixo para garantir que o programa funcione corretamente:

- Crie uma nova pasta (exemplo: na Área de Trabalho)  
- Copie para essa pasta o arquivo `.jar` baixado  
- Coloque na mesma pasta a planilha Excel com os dados (`.xlsx` ou `.xls`)  

⚠️ O programa processa **apenas um arquivo Excel** por execução. Evite ter múltiplas planilhas na mesma pasta para prevenir erros.

Exemplo de estrutura:

/MinhaPastaDeTrabalho
├── automocao_planilha.jar
└── meus_dados.xlsx

### 3. Executar o Programa

Dê um duplo clique no arquivo `automocao_planilha.jar` para iniciar a automação.

---

## ⚠️ Customização para Desenvolvedores

Este projeto é um **template**. Para adequá-lo ao seu caso específico, será necessário editar o código-fonte, principalmente a classe `PreenchedorDeFormulario.java`.

### Passos principais de customização:

- **Alterar a URL do Formulário:**  
  Modifique a constante `FORM_URL` para a URL do formulário que deseja automatizar.

```java
// Em PreenchedorDeFormulario.java
private static final String FORM_URL = "https://SEU-SITE-AQUI.com";
Ajustar a lógica de preenchimento:
No método preencherFormulario(...), altere os seletores dos elementos HTML para os que correspondem ao seu formulário (usando By.id, By.name, By.xpath, etc.).

Também ajuste o mapeamento das células da planilha para os campos do formulário, garantindo que os dados sejam inseridos corretamente.

Exemplo Java:

// Exemplo de preenchimento de campo, altere conforme seu formulário e planilha
preencherCampoSeguro(wait, By.name("seletor-do-seu-elemento-1"), formatter.formatCellValue(row.getCell(1)));
preencherCampoSeguro(wait, By.id("seletor-do-seu-elemento-2"), formatter.formatCellValue(row.getCell(2)));
Após essas modificações, recompile o projeto para gerar um novo .jar com suas adaptações.

📄 Licença
Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para detalhes completos.
