# Backlog


## Épicos, Features e User Stories
|Épicos | Features | Histórias de Usuário | Etapas de produção | Prioridade |
| -- | -- | -- | -- | --
| <div >1. Acesso ao sistema</div> | Login Usuário | [US01](#us01) | MVP | 1 |
|  | Login Admin | [US05](#us05) | Incremento | 13 |
|2. Preenchimento da planilha | Seleção de projeto | [US07](#us07) | MVP |2 |
|  | Seleção de template | [US02](#us02) | MVP |3 |
|3. Exportação do projeto | Exportação xslx | [US09](#us09)| MVP | 4|
|  | Exportação PDF | [US10](#us10)|Incremento |7 |
|4. Visualização dos projetos | Lista de projetos | [US14](#us14) |Incremento |8 |
|5. Pesquisa por projeto | Ordenação | [US03](#us13)|Incremento |10 |
|  | Filtragem | [US04](#us04)|Incremento |11 |
|  | Busca | [US07](#us07)|Incremento | 9|
|6. Detalhamento de informações | Informações do projeto | [US11](#us11)| Incremento|12 |
|7. Cadastro de template | Novo template | [US06](#us06)|Incremento |14 |
|8. Rastreamento | Notificação, | [US08](#us08)| Incremento| 5 |
|  | Histórico | [US12](#us12) |Incremento | 15|
|  | Reportar | [US13](#us13) |Incremento |6 |

## User Stories

#### Personas
* <b>Administrador:</b> Responsável por adicionar tabelas novas e configurar suas regras, garantindo que a solução atenda às necessidades da Finatec.
* <b>Usuário:</b> Indiviudo que realizará automação para gerar relatórios precisos e analisar dados financeiros, visando otimizar os processos de prestação de contas . 
  
### **US01**
**Eu, como usuário, quero ser capaz de acessar o sistema com login e senha para acessar a aplicação.**

Critérios de aceitação:

  -  Os campos a serem preenchidos devem ser: Usuário e Senha;
  - O sistema deve verificar se o usuário inserido está cadastrado no sistema;
  - O sistema deve verificar se a senha está correta e seja associada ao usuário inserido:
    - Caso negativo: Exibir mensagem de erro;
    - Caso positivo: O usuário deve ser autenticado e direcionado para a tela principal do sistema;

### **US02 **
**Eu, como usuário, quero poder selecionar o template** 

Critérios de aceitação:

  - Deve ser possível selecionar o template
  - O sistema deve impedir a seleção do template se o não houver um projeto selecionado;
  - O sistema deve possuir uma lista com todos os templates.

### **US03**
**Eu, como usuário, quero ordenar a lista de projetos por diferentes critérios para orientar e otimizar minha busca.**

Critérios de aceitação:

  -  Ordenar por: Nome do financiador; Nome do projeto; Período de vigência; Data de Assinatura; Data de Encerramento; Prazo Final.
  - As ordenações devem ser tanto em ordem crescente quanto decrescente;
  - As ordenações das datas devem ser possível pela ultima data de atualização ?

### **US04**
**Eu, como usuário, quero filtrar a lista de projetos através de diferentes parâmetros para facilitar na organização.**

Critérios de aceitação:

  - Filtrar por: Nome do financiador; Período de vigência; Data de Assinatura (mês); Data de Encerramento; Período de despesa.

### **US05**
**Eu, como administrador, quero ser capaz de acessar o sistema, através de login e senha específico para obter privilégios de administração.**

Critérios de aceitação:

  - Os campos a serem preenchidos devem ser: Usuário e Senha;
  - O sistema deve verificar se o usuário inserido está cadastrado no sistema como administrador;
  - O sistema deve verificar se a senha está correta e seja associada ao usuário inserido:
    - Caso negativo: Exibir mensagem de erro;
    - Caso positivo: O usuário deve ser autenticado e direcionado para a tela principal do sistema;

### **US06**
**Eu, como administrador, quero ser capaz de adicionar novos templates para tornar o sistema escalável**

Critérios de aceitação:

  - O sistema deve permitir a importação e cadastro de novos templates por meio do login de administrador;
  - O administrador deverá informar as células a serem preenchidas, de acordo com os dados existentes;

### **US07** 
**Eu, como usuário, quero buscar os projetos por texto digitado para encontrar o que preciso de forma rápida**

Critérios de aceitação:
  
  - O sistema deve permitir que a busca seja feita por id do projeto;
  - O sistema deve permitir que a busca seja feita pelo nome do projeto;
  - O sistema deve permitir que a busca seja feita pelo financiador do projeto;
  - O sistema deverá validar os campos de entrada, garantindo que não estejam errados ou vazios, onde:
  - Deverá exibir uma mensagem de erro caso tenha algo errado

### **US08**
**Eu, como usuário, quero saber todos os dados não preenchidos, indicando quais são os campos vazios para saber o que necessita ser preenchido manualmente**

Critérios de aceitação:

  - O sistema deve verificar se há campos vazios: Deve ser listado quais páginas e os campos vazios;

### **US09** 
**Eu, como usuário, quero exportar a planilha no formato xlsx para compartilhar ou manipular os dados em programas externos**

Critérios de aceitação:

  - O sistema deve permitir a exportação da planilha no formato xlsx;
  - O sistema deve exibir uma animação que foi exportado com sucesso
  - O sistema deve exibir uma animação de carregamento enquanto a planilha é preenchida.

### **US10**
**Eu, como usuário, quero poder exportar a planilha no formato pdf para fins de documentação ou apresentação**

Critérios de aceitação:

  - O sistema deve permitir a exportação da planilha no formato PDF;
  - O sistema deve exibir uma animação que foi exportado com sucesso;
  - O sistema deve exibir uma animação de carregamento enquanto o pdf é gerado.

### **US11** 
- **Eu, como usuário, quero visualizar os detalhes de um projeto selecionado prestes a ser exportado para entender melhor seu conteúdo e propósito**
- Critérios de aceitação:
  - O sistema deve detalhar algumas informações principais como: ID; Nome do projeto; Nome do financiador; Período; Datas; Coordenador; 
  - Deverá ter uma tela específica para mostrar as informações.

### **US12**
**Eu, como administrador, quero visualizar o histórico de exportações para ter o controle passado de exportações realizadas**

Critérios de aceitação:

- O sistema deve mostrar uma lista com todas as exportações indicando: Data que foi feito; ID do projeto; Nome de quem realizou;

### **US13**
**Eu, como usuário, quero ter uma maneira de reportar erros identificados para garantir a integridade dos templates/projetos**

Critérios de aceitação:

  - Deve rastrear  o usuário que reportou;
  - Deve possuir obrigatoriamente um campo para indicar o assunto;
  - Deve possuir obrigatoriamente  um campo para detalhar o problema encontrado;
  - Deve possuir um campo para indicar o template;
  - Deve possuir um campo indicando o projeto;

### **US14**
**Eu, como usuário, quero visualizar a lista de projetos para que eu possa selecionar o desejado**

Critérios de aceitação:

  - Deve apresentar todos os projetos na página inicial listados (como tabela)
  - Dever apresentar informações básicas como id, nome, coordenador, e datas.
  - Deve ter os botões de exportação diminuindo as etapas para exportar um projeto

## Histórico de Versão

| Data | Versão | Descrição | Autor | Revisor | Issue |
| --- | --- | --- | --- | --- | --- |
| 12/08/2023 | 1.0 | Criação do documento |  [Isaac](https://github.com/IsaacLusca), [Hemanoel](https://github.com/hemanoelbritoF) | [Raquel](https://github.com/raqueleucaria) |[#35](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/issues/35)|
| 12/08/2023 | 2.0 | Revisão do documento com as requisições no review do PR |  [Isaac](https://github.com/IsaacLusca), [Hemanoel](https://github.com/hemanoelbritoF) | [Raquel](https://github.com/raqueleucaria) |[#35](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/issues/35)|
| 15/08/2023 | 3.0 | Revisão do documento com os novos requisitos |  [Isaac](https://github.com/IsaacLusca), [Hemanoel](https://github.com/hemanoelbritoF) | [Raquel](https://github.com/raqueleucaria) |[#35](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/issues/35)|
| 23/08/2023 | 4.0 | Revisão do documento com os novos requisitos |  [Isaac](https://github.com/IsaacLusca), [Hemanoel](https://github.com/hemanoelbritoF) | [Raquel](https://github.com/raqueleucaria) |[#35](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/issues/35)|
