# Generated by Django 4.2.5 on 2023-10-06 17:38

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Export',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('data_export', models.DateTimeField(auto_now_add=True)),
                ('formato', models.CharField(max_length=200)),
                ('nome_template', models.CharField(max_length=200)),
                ('nome_usuario', models.CharField(max_length=200)),
                ('id_projeto', models.CharField(max_length=200)),
            ],
        ),
        migrations.CreateModel(
            name='Mapeamento',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('CODIGO', models.CharField(max_length=70)),
                ('NOME', models.CharField(max_length=70)),
                ('SALDO', models.CharField(max_length=70)),
                ('DATA_ASSINATURA', models.CharField(max_length=70)),
                ('DATA_VIGENCIA', models.CharField(max_length=70)),
                ('DATA_ENCERRAMENTO', models.CharField(max_length=70)),
                ('TIPO_CONTRATO', models.CharField(max_length=70)),
                ('INSTITUICAO_EXECUTORA', models.CharField(max_length=70)),
                ('PROCESSO', models.CharField(max_length=70)),
                ('SUBPROCESSO', models.CharField(max_length=70)),
                ('COD_PROPOSTA', models.CharField(max_length=70)),
                ('PROPOSTA', models.CharField(max_length=70)),
                ('OBJETIVOS', models.CharField(max_length=70)),
                ('VALOR_APROVADO', models.CharField(max_length=70)),
                ('NOME_TP_CONTROLE_SALDO', models.CharField(max_length=70)),
                ('GRUPO_GESTORES', models.CharField(max_length=70)),
                ('GESTOR_RESP', models.CharField(max_length=70)),
                ('COORDENADOR', models.CharField(max_length=70)),
                ('PROCEDIMENTO_COMPRA', models.CharField(max_length=70)),
                ('TAB_FRETE', models.CharField(max_length=70)),
                ('TAB_DIARIAS', models.CharField(max_length=70)),
                ('CUSTO_OP', models.CharField(max_length=70)),
                ('NOME_FINANCIADOR', models.CharField(max_length=70)),
                ('DEPARTAMENTO', models.CharField(max_length=70)),
                ('SITUACAO', models.CharField(max_length=70)),
                ('BANCO', models.CharField(max_length=70)),
                ('AGENCIA_BANCARIA', models.CharField(max_length=70)),
                ('CONTA_BANCARIA', models.CharField(max_length=70)),
                ('CENTRO_CUSTO', models.CharField(max_length=70)),
                ('CONTA_CAIXA', models.CharField(max_length=70)),
                ('CATEGORIA_PROJETO', models.CharField(max_length=70)),
                ('COD_CONVENIO_CONTA', models.CharField(max_length=70)),
                ('COD_STATUS', models.CharField(max_length=70)),
                ('IND_SUB_PROJETO', models.CharField(max_length=70)),
                ('TIPO_CUSTO_OP', models.CharField(max_length=70)),
                ('PROJETO_MAE', models.CharField(max_length=70)),
                ('ID_COORDENADOR', models.CharField(max_length=70)),
                ('ID_FINANCIADOR', models.CharField(max_length=70)),
                ('ID_INSTITUICAO', models.CharField(max_length=70)),
                ('ID_DEPARTAMENTO', models.CharField(max_length=70)),
                ('NOME_INSTITUICAO', models.CharField(max_length=70)),
                ('ID_INSTITUICAO_EXECUTORA', models.CharField(max_length=70)),
                ('ID_TIPO', models.CharField(max_length=70)),
            ],
        ),
        migrations.CreateModel(
            name='Report',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('titulo', models.CharField(max_length=200)),
                ('descricao', models.CharField(max_length=200)),
                ('tipo_erro', models.CharField(max_length=200)),
                ('id_projeto', models.CharField(max_length=200)),
                ('nome_usuario', models.CharField(max_length=200)),
            ],
        ),
        migrations.CreateModel(
            name='Template',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nome_template', models.CharField(max_length=200)),
                ('endereco_template', models.CharField(max_length=200)),
                ('mapeamento', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='app.mapeamento')),
            ],
        ),
    ]