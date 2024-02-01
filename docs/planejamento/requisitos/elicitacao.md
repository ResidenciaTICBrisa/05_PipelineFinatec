# Elicitação de Requisitos

## Introdução

A elicitação de requisitos é o processo de coleta, identificação, compreensão e documentação das necessidades, expectativas e restrições dos stakeholders (partes interessadas) em relação a um sistema ou software a ser desenvolvido. Ela é uma etapa crucial no ciclo de vida do desenvolvimento de software, pois estabelece a base para a criação de um sistema que atenda aos objetivos e requisitos dos usuários, clientes e demais envolvidos. Dessa forma, este documento registra os requisitos elicitados para o sistema de prestação de contas desenvolvido para a FINATEC.

## Metodologia

A elicitação de requisitos por meio da introspecção envolve a compreensão das propriedades que um sistema deve possuir para alcançar o sucesso desejado. Nesse método, o Engenheiro de Requisitos é desafiado a imaginar as suas próprias preferências e necessidades ao desempenhar uma determinada tarefa, considerando os recursos e equipamentos disponíveis.

Os requisitos forma previamente definidos e posteriormente discutidos em grupo, assim, registrando e os listando.

## Requisitos

| ID | Requisito | Tipo |
| --- | --- | --- |
| 1 | Deve ser possível que os usuários realizem login  | RF |
| 2 | Deve ser possível que os usuários visualizem a lista de template | RF |
| 3 | Deve ser possível que os usuários filtrem a lista de planilhas por financiador | RF |
| 4 | Deve ser possível que os usuários ordenem a lista de planilhas por prazo | RF |
| 5 | Deve ser possível que os usuários ordenem por nome do projeto (a-z ou z-a) | RF |
| 6 | Deve ser possível que os usuários ordenem por financiador (a-z ou z-a) | RF |
| 7 | Deve ser possível que os usuários visualizem a planilha desejada | RF |
| 8 | Deve ser possível que os usuários busquem as planilhas por texto digitado (código, nome projeto, financiador…) | RF |
| 9 | Deve ser possível os usuários requisitem o preenchimento da planilha | RF |
| 10 | Deve ser possível o usuário tenha uma pré-visualização das informações da planilha | RF |
| 11 | Deve ser possível o usuário tenha a visualização completa das informações da planilha | RF |
| 12 | Deve ser possível acessar como administrador  | RF  |
| 13 | Deve ser possível como administrador cadastrar novos templates | RF |
| 15 | Deve ser possível alertar que a planilha não tem todos os dados preenchidos | RF |
| 16 | Deve ser possível que os usuários exportem a planilha no formato xlsx | RF |
| 17 | Deve ser possível que os usuários exportem a planilha no formato pdf | RF |
| 18 | A aplicação deve ser capaz de lidar com pelo menos 10 usuários simultâneos sem comprometer o desempenho ou a segurança | RNF |
| 19 | A aplicação deve seguir as diretrizes de acessibilidade, garantindo que seja acessível para pessoas com deficiência visual, auditiva ou física ** | RNF |
| 20 | O software desenvolvido será para ambiente desktop | RNF |
| 21 | O software desenvolvido será para ambiente dispositivos móveis ** | RNF |
| 22 | A interface do usuário deve ser intuitiva e de fácil navegação | RNF |
| 23 | Quando houver uma queda na conexão de internet do usuário, a aplicação deve ser capaz de fornecer um modo de degradação, permitindo a execução de funções básicas offline, como visualização de template já requisitado ** | RNF |
| 24 | O sistema deve ser capaz de realizar as transformações necessárias para gerar as 15 tabelas de saídas diferentes | RNF |
| 25 | O sistema deve ser capaz de conectar o banco de dados para importar os dados de prestação de contas | RNF |
| 26 | O software deve ser desenvolvido em python, flask, Django, SQL Server, HTML, CSS, JS | RNF |
| 27 | O software deve ser desenvolvido utilizando funções | RNF |
| 28 | A interação com o usuário deverá ser feita por meio de interface gráfica | RNF |

### Legenda

- **RF**: Requisitos Funcionais
- **RNF**: Requisitos não Funcionais

## Próxima Etapa

Reunião Presencial na FINATEC com os **Stakeholders (15/08 - 13h as 17h)**

  1.  Validação;
  2.  Priorização.

## Histórico de Versão

| Data | Versão | Descrição | Autor | Revisor | Issue |
| --- | --- | --- | --- | --- | --- |
| 11/08/2023 | 1.0 | Criação do documento | [Raquel](https://github.com/raqueleucaria) | [Isaac](https://github.com/IsaacLuscaEditar) | [#31](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/issues/31) |
