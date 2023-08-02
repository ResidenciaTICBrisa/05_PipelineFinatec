
# Mkdocs 


## Introdução
O **MkDocs** é um gerador de sites estáticos rápido , simples e absolutamente lindo, voltado para a construção de documentação de projetos. Os arquivos de origem da documentação são gravados em Markdown e configurados com um único arquivo de configuração YAML.

Um **workflows** (fluxo de trabalho) é um processo automatizado configurável que executará um ou mais trabalhos. Os fluxos de trabalho são definidos por um arquivo YAML com check-in em seu repositório e serão executados quando acionados por um evento em seu repositório, ou podem ser acionados manualmente ou em um cronograma definido.
Os fluxos de trabalho são definidos no .github/workflowsdiretório de um repositório, e um repositório pode ter vários fluxos de trabalho, cada um dos quais pode executar um conjunto diferente de tarefas.


## Documentação
- [Mkdocs](https://www.mkdocs.org/user-guide/installation/)
- [Material-mkdocs](https://squidfunk.github.io/mkdocs-material/getting-started/)
- [Workflows](https://docs.github.com/en/actions/using-workflows/about-workflows#about-workf)



## Tutorial
### 1. Mkdocs
#### Pré-Instalação
- Para instalar o mkdocs é preciso ter o pacote pip do Python. 
- Verifique se possui: `python --version` e `pip --version`
- Se for preciso intalar o pyhton, recomendo o Python 3.10. Se for preciso instalar o pip pela primeira vez baixe [get-pip.py](https://bootstrap.pypa.io/get-pip.py). 
- Em seguida, execute o seguinte comando para instalá-lo: `python get-pip.py
`
#### Instalação
- Instale o mkdocspacote usando pip: `pip install mkdocs`
- Verifique a instalação:` mkdocs --version`

#### Instalação Material
- Instale os temas para Mkdocs que utilizaremos inicialmente: `pip install mkdocs-material`
- Iniciando projeto (no terminal):  `mkdocs -h`
- Criando um site: `mkdocs new [nome]`
- Abrindo o site localmente: `mkdocs serve`
- Construindo a documentação: `mkdocs build`
- Implantando a documentação nas páginas do GitHub: `mkdocs gh-deploy`

### 2. Workflows
- Criando a pasta do arquivo: `.github/workflows`
- Criando o arquivo de deploy(deploy.yml):

  ```
  name: Deploy to GitHub Pages
  on:
      push:
        branches:
          - main
      pull_request:
        branches:
          - main
  jobs:
    deploy:
      runs-on: ubuntu-latest
      steps:
        - uses: actions/checkout@v2
        - uses: actions/setup-python@v2
          with:
            python-version: 3.x
        - run: pip install mkdocs-material
        - run: mkdocs build
        - run: mkdocs gh-deploy --force
  ```


**OBS**: é preciso dar a permissão de gravação para o workflow nas configurações do repositório.
![image](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/assets/81540491/fef08195-3460-4b29-9985-ac6d1ad12111)
![image](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/assets/81540491/586682bd-caae-4f5d-b7ee-2f6480cf9843)


## Histórico de Versão
|  Data  | Versão | Descrição | Autor  |  Revisor  |Issue|
|------- | ------ |---------- | ------ | --------- |-----|
| 30/06/2023 |     1.0   | Criação do documento | [Raquel](https://github.com/cansancaojennifer)  | [Hemanoel](https://github.com/hemanoelbritoF) |[#7](https://github.com/ResidenciaTICBrisa/05_PipelineFinatec/issues/7)|