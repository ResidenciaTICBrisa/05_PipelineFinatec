from django.db import models

class Mapeamento(models.Model):
    id_mapeamento               = models.IntegerField(primary_key=True)
    codigo                      = models.CharField(max_length=200)
    nome                        = models.CharField(max_length=200)
    saldo                       = models.CharField(max_length=200)
    data_assinatura             = models.CharField(max_length=200)
    data_vigencia               = models.CharField(max_length=200)
    data_encerramento           = models.CharField(max_length=200)
    tipo_contrato               = models.CharField(max_length=200)
    instituicao_executora       = models.CharField(max_length=200)
    processo                    = models.CharField(max_length=200)
    subprocesso                 = models.CharField(max_length=200)
    cod_proposta                = models.CharField(max_length=200)
    proposta                    = models.CharField(max_length=200)
    objetivos                   = models.CharField(max_length=200)
    valor_aprovado              = models.CharField(max_length=200)
    nome_tp_controle_saldo      = models.CharField(max_length=200)
    grupo_gestores              = models.CharField(max_length=200)
    gestor_resp                 = models.CharField(max_length=200)
    coordenador                 = models.CharField(max_length=200)
    procedimento_compra         = models.CharField(max_length=200)
    tab_frete                   = models.CharField(max_length=200)
    tab_diarias                 = models.CharField(max_length=200)
    custo_op                    = models.CharField(max_length=200)
    nome_financiador            = models.CharField(max_length=200)
    departamento                = models.CharField(max_length=200)
    situacao                    = models.CharField(max_length=200)
    banco                       = models.CharField(max_length=200)
    agencia_bancaria            = models.CharField(max_length=200)
    conta_bancaria              = models.CharField(max_length=200)
    centro_custo                = models.CharField(max_length=200)
    conta_caixa                 = models.CharField(max_length=200)
    categoria_projeto           = models.CharField(max_length=200)
    cod_convenio_conta          = models.CharField(max_length=200)
    cod_status                  = models.CharField(max_length=200)
    ind_sub_projeto             = models.CharField(max_length=200)
    tipo_custo_op               = models.CharField(max_length=200)
    projeto_mae                 = models.CharField(max_length=200)
    id_coordenador              = models.CharField(max_length=200)
    id_financiador              = models.CharField(max_length=200)
    id_instituicao              = models.CharField(max_length=200)
    id_departamento             = models.CharField(max_length=200)
    nome_instituicao            = models.CharField(max_length=200)
    id_instituicao_executora    = models.CharField(max_length=200)
    id_tipo                     = models.CharField(max_length=200)
    
class Lancamento(models.Model):
    id_mapeamento               = models.IntegerField(primary_key=True)
    id_favorecido               = models.CharField(max_length=200)
    nome_favorecido             = models.CharField(max_length=200)        
    cnpj_favorecido             = models.CharField(max_length=200)        
    tipo_favorecido             = models.CharField(max_length=200)        
    valor_lancado               = models.CharField(max_length=200)        
    valor_pago                  = models.CharField(max_length=200)    
    data_vencimento             = models.CharField(max_length=200)        
    id_status                   = models.CharField(max_length=200)    
    status_lancamento           = models.CharField(max_length=200)            
    flag_receita                = models.CharField(max_length=200)        
    data_baixa                  = models.CharField(max_length=200)    
    his_lancamento              = models.CharField(max_length=200)        
    data_emissao                = models.CharField(max_length=200)        
    num_doc_fin                 = models.CharField(max_length=200)    
    data_cria                   = models.CharField(max_length=200)    
    data_pagamento              = models.CharField(max_length=200)        
    id_lancamento               = models.CharField(max_length=200)        
    id_projeto                  = models.CharField(max_length=200)    
    id_rubrica                  = models.CharField(max_length=200)    
    nome_rubrica                = models.CharField(max_length=200)        
    tipo_movimento              = models.CharField(max_length=200)        
    id_tp_lancamento            = models.CharField(max_length=200)            
    tipo_lancamento             = models.CharField(max_length=200)  
    
class Template(models.Model):
    nome_template               = models.CharField(max_length=200)
    endereco_template           = models.CharField(max_length=200)
    mapeamento                  = models.ForeignKey(Mapeamento, on_delete=models.CASCADE)

class Report(models.Model):
    titulo                      = models.CharField(max_length=200)
    descricao                   = models.CharField(max_length=200)
    
    TIPO_ERRO_CHOICES = (
        ('erro1', 'Falta de informação no projeto'),
        ('erro2', 'Campo preenchido incorretamente'),
        # Adicione mais opções conforme necessário
    )
    tipo_erro                   = models.CharField(max_length=20, choices=TIPO_ERRO_CHOICES)
    id_projeto                  = models.CharField(max_length=50)
    nome_usuario                = models.CharField(max_length=200)

class Export(models.Model):
    data_export                 = models.DateTimeField(auto_now_add=True)
    formato                     = models.CharField(max_length=200)
    nome_template               = models.CharField(max_length=200)
    nome_usuario                = models.CharField(max_length=200)
    id_projeto                  = models.CharField(max_length=50)

class UserActivity(models.Model):
    user_id = models.CharField(max_length=255)
    tag = models.CharField(max_length=50)
    activity = models.TextField()
    timestamp = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.user_id} - {self.activity}"