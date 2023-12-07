# Instalação do PostgreSQL no Ubuntu 20.04 para Projetos Django

## Introdução

Este guia tem como objetivo auxiliar na instalação e configuração do PostgreSQL em um sistema Ubuntu 20.04, especialmente para projetos Django. Ele segue as diretrizes do [artigo original da DigitalOcean](https://www.digitalocean.com/community/tutorials/how-to-use-postgresql-with-your-django-application-on-ubuntu-20-04).

## Instalação no Linux

Certifique-se de que seu sistema está atualizado e, em seguida, instale as dependências necessárias:

```shell
sudo apt update
sudo apt install python3-pip python3-dev libpq-dev postgresql postgresql-contrib
```

## Executando o PostgreSQL

Após a instalação, é possível acessar o PostgreSQL e realizar tarefas administrativas. Para entrar no console do PostgreSQL, utilize o seguinte comando:

```shell
sudo -u postgres psql
```
## Criando um Banco de Dados e um Usuário

Para configurar um banco de dados para o seu projeto Django e criar um usuário correspondente, siga as etapas abaixo:

Acesse o console do PostgreSQL como o superusuário "postgres" (conforme mencionado anteriormente):

```shell
sudo -u postgres psql
```
Dentro do console do PostgreSQL, crie um banco de dados para o seu projeto. Substitua meuprojeto pelo nome desejado do banco de dados:

```sql
CREATE DATABASE meuprojeto;
```
Crie um usuário e defina uma senha para ele. Substitua meuusuario e senha pelos valores desejados:

```sql
CREATE USER meuusuario WITH PASSWORD 'senha';
```

Configure as preferências de codificação, isolamento de transações e fuso horário para o usuário criado:

```sql
ALTER ROLE meuusuario SET client_encoding TO 'utf8';
ALTER ROLE meuusuario SET default_transaction_isolation TO 'read committed';
ALTER ROLE meuusuario SET timezone TO 'America/Sao_Paulo';
```

> Nota Importante: Se você encontrar erros relacionados a permissões insuficientes ao referenciar o banco de dados criado em um projeto Django, ajuste as permissões do banco de dados usando o seguinte comando:

```sql
ALTER DATABASE meuprojeto OWNER TO meuusuario;
```
## Configuração no Django

Para que o Django utilize o PostgreSQL como banco de dados padrão ou adicional, faça as seguintes alterações no arquivo settings.py do seu projeto:

Para o banco de dados padrão:

```python

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'meuprojeto',
        'USER': 'meuusuario',
        'PASSWORD': 'senha',
        'HOST': 'localhost',
        'PORT': '',
    }
}
```

Para adicionar uma referência a outro banco de dados PostgreSQL (caso não seja o banco de dados principal):

```python

    DATABASES = {
        'default': {
            ...
        },
        'meuprojeto': {
            'ENGINE': 'django.db.backends.postgresql',
            'NAME': 'meuprojeto',
            'USER': 'meuusuario',
            'PASSWORD': 'senha',
            'HOST': 'localhost',
            'PORT': '',
        }
    }
```

Lembre-se de selecionar o banco de dados correto ao executar as migrações do Django. Para especificar um banco de dados diferente ao usar o comando makemigrations, utilize a opção --database. Por exemplo:

```shell
python manage.py makemigrations --database=meuprojeto
```

Com essas configurações, você estará pronto para usar o PostgreSQL com o seu projeto Django no Ubuntu 20.04. Certifique-se de adaptar os nomes de banco de dados, usuários e senhas de acordo com as necessidades específicas do seu projeto.