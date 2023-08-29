# Introdução

# Instalação Linux

    sudo apt update
    sudo apt install python3-pip python3-dev libpq-dev postgresql postgresql-contrib

## Rodando PostgreSQL
    sudo -u postgres psql

## Criando Banco de dados e Usuário
    CREATE DATABASE meuprojeto;

    CREATE USER meuusuario WITH PASSWORD 'senha';

    ALTER ROLE meuusuario SET client_encoding TO 'utf8';
    ALTER ROLE meuusuario SET default_transaction_isolation TO 'read committed';
    ALTER ROLE meuusuario SET timezone TO 'America/Sao_Paulo';