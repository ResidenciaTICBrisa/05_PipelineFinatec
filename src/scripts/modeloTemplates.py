import datetime

class ModeloTemplate():
    def fundep(self):
        modelo_fundep = {
            "Relação e despesas" : [
                # campos cadastrais
                ("C3", f"{records.NOME_INSTITUICAO_EXECUTORA}"),
                ("F3", "0000*XX"),
                ("I3", "n_acordo"),
                ("C4", f"{records.NOME_PROJETO}"),
                ("C5", f"{records.NOME_COORDENADOR}"),

                # item beneficiarios
                ("I5", "Dt_PeriodoPrestacao"),
                ("C7", "Nome_Beneficiario"),
                ("D7", "CNPJ-CPF-Baneficiario"),
                ("E7", "Rubrica-Beneficiario"),
                ("F7", "CH-OB-Beneficiario"),
                ("G7", "dt-inicial-Beneficiario"),
                ("H7", "NF-DOC.FISCAL-Beneficiario"),
                ("I7", "dt-final-Beneficiario"),
                ("J7", "valor-Beneficiario"),
            ]
        }
        return modelo_fundep
        
    def fub(self):
        modelo_fub = {
            "Receita x Despesa": [
                ("A3:J3", f"Título do Projeto: {records.NOME}"), #AJEITAR
                ("A4:J4", f"Executora:  {NOME_INSTITUICAO_EXECUTORA}"), 
                ("A5:J5", f"Partícipe: {NOME_INSTITUICAO}"),      # o que eh participe?
                ("A6:J6", f"Período de Execução Físico-Financeiro: {DATA_ASSINATURA} a {DATA_VIGENCIA}"),
                ("A7:J7", f"Período que abrange esta prestação:  {DATA_VIGE}"), #VAZIO
                ("A16:A25", datetime.datetime.strptime("2014-06-23", "%Y-%m-%d")),
                ("B16:B25", "STRINGB16B25"),# intervalo nao interfere
                ("C16:C25", "STRINGC16C25"),#intervalo nao interfere
                ("E16:E25", 200),
                #("I16:I21",223),
                ("I16", 23),
                ("I17", 213),
                ("I18", 223),
                ("I19", 233),
                ("I20", 243),
                ("I21", 253),
                ("I24", 263),
                ("I26", 23787),
                ("I27", 10),
                ("I32", 100),
                ("I33", 1000),
                ("I38", 10000),
                ("H45", f"NOME_COORDENADOR")
            ],
            "Exec. Receita e Despesa":[
                ("B16",200),
                ("B17",201),
                ("B18",202),
                ("B19",203),
                ("B20",204),
                ("B21",205),
                ("B22",206),
                ("B23",207),
                ("C16",205),
                ("C17",206),
                ("C18",207),
                ("C19",208),
                ("C20",209),
                ("C21",210),
                ("C22",211),
                ("C23",212),
                ("C24",213),
                ("C25",214),
                ("F16",2051),
                ("F17",2061),
                ("F18",2071),
                ("F19",2081),
                ("F20",2091),
                ("F21",2101),
                ("F22",2111),
                ("F23",2121),
                ("G16",20351),
                ("G17",20361),
                ("G18",20371),
                ("G19",20381),
                ("G20",20391),
                ("G21",21301),
                ("G22",21311),
                ("G23",21321),
                ("I26",2011),
                ("I28",2001),
                ("I29",2001),
                ("B26",3011),
                ("B28",3001),
                ("B29",3001),
                ("B31",3011),
                ("C26",4011),
                ("C29",4011),
                ("C31",4011),
                ("F26",5011),
                ("F28",5001),
                ("F29",5001),
                ("F31",5011),
                ("G26",6011),
                ("G28",6001),
                ("G29",6001),
                ("G31",6011)
            ],
            "Pessoa Jurídica":[
                ("B11","TESTE_NOME"),
                ("C11","TESTE_CPF"),
                ("D11","TESTE_ESPECIFICACAO"),
                ("E11","TESTE_DESCRICAO"),
                ("F11","TESTE_DESCRICAO"),
                ("F11","TESTE_RECIBO"),
                ("G11",110101),
                ("H11","TESTE_CHEQUE"),
                ("I11",98765431),
                ("J11",3000)
            
            ],
            "Conciliação Bancária":[
                ("F10",5000),
                ("F11",5000),
                ("A15",120623),
                ("A15",120623),
                ("B15",9777),
                ("C15","DOCUMENTO_TESTE"),
                ("D15","DESCRIÇÃO_TESTE"),
                ("B38",9777),
                ("B39",7878),
                ("C38","TESTESTRALEATORIOC38"),
                ("C39",'TESTESTRALEATORIOC39'),
                ("D38","TESTESTRALEATORIOD38"),
                ("D39",'TESTESTRALEATORIOD39')
            
            ],
            "Rendimento de Aplicação":[
                ("B12",2023),
                ("C12",2024),
                ("D12",2025),
                ("E12",2026),
                ("F12",2027),
                ("G12",2028),
                ("H12",2029)
            ]
        }

        return modelo_fub

    def opas(self):
        modelo_opas = {
            "Relatório Consolidado" : [
                # campos cadastrais
                ("C4","numero_loa"),
                ("C5","nome_beneficiario"),
                ("C6","nome_representante_legal"),
                ("C7","titulo_loa"),
                ("C8","periodo_despesas"),
                ("C9","montante_recebido"),
                ("C10","balanco_pagtos_anteriores"),

                # atividade 1
                ("A15","atividade_1"),
                
                # montante previsto/executado atividade 1
                ("D15","passagens_diarias_previsto"),
                ("E15","passagens_diarias_executado"),
                
                ("D16","equipamentos_previsto"),
                ("E16","equipamentos_executado"),
                
                ("D17", "construcao_previsto"),
                ("E17", "construcao_executado"),
                
                ("D18", "pessoa_fisica_previsto"),
                ("E18", "pessoa_fisica_executado"),
                
                ("D19", "servicoes_materiais_previsto"),
                ("E19", "servicoes_materiais_executado"),
                
                ("D20", "treinamento_capacitacao_previsto"),
                ("E20", "treinamento_capacitacao_executado"),
                
                # atividade 2
                ("A21","atividade_2"),

                # montante previsto/executado atividade 2
                ("D21","passagens_diarias_previsto"),
                ("E21","passagens_diarias_executado"),
                
                ("D22","equipamentos_previsto"),
                ("E22","equipamentos_executado"),
                
                ("D23", "construcao_previsto"),
                ("E23", "construcao_executado"),
                
                ("D24", "pessoa_fisica_previsto"),
                ("E24", "pessoa_fisica_executado"),
                
                ("D25", "servicoes_materiais_previsto"),
                ("E25", "servicoes_materiais_executado"),
                
                ("D26", "treinamento_capacitacao_previsto"),
                ("E26", "treinamento_capacitacao_executado"),

                # atividade 3
                ("A27","atividade_3"),

                # montante previsto/executado atividade 3
                ("D27","passagens_diarias_previsto"),
                ("E27","passagens_diarias_executado"),
                
                ("D28","equipamentos_previsto"),
                ("E28","equipamentos_executado"),
                
                ("D29", "construcao_previsto"),
                ("E29", "construcao_executado"),
                
                ("D30", "pessoa_fisica_previsto"),
                ("E30", "pessoa_fisica_executado"),
                
                ("D31", "servicoes_materiais_previsto"),
                ("E31", "servicoes_materiais_executado"),
                
                ("D32", "treinamento_capacitacao_previsto"),
                ("E32", "treinamento_capacitacao_executado"),

                # atividade 4
                ("A33","atividade_4"),

                # montante previsto/executado atividade 4
                ("D33","passagens_diarias_previsto"),
                ("E33","passagens_diarias_executado"),
                
                ("D34","equipamentos_previsto"),
                ("E34","equipamentos_executado"),
                
                ("D35", "construcao_previsto"),
                ("E35", "construcao_executado"),
                
                ("D36", "pessoa_fisica_previsto"),
                ("E36", "pessoa_fisica_executado"),
                
                ("D37", "servicoes_materiais_previsto"),
                ("E37", "servicoes_materiais_executado"),
                
                ("D38", "treinamento_capacitacao_previsto"),
                ("E38", "treinamento_capacitacao_executado"),

            ],

            "Relatório Detalhado" : [
                # campos cadastrais
                ("C4","numero_loa"),
                ("C5","nome_beneficiario"),
                ("C6","nome_representante_legal"),
                ("C7","titulo_loa"),
                ("C8","periodo_despesas"),

                # despesas
                ("B13:B42", "nome_fornecedor"),
                ("C13:C42", "categoria_despesa"),
                ("D13:D42", "data_despesa"),
                ("E13:E42", "montante_despesa"),

            ]
        }
        return modelo_opas
    
    def fap(self):
        modelo_fap = {
            "" : [],
        }
        return modelo_fap
    
    def fup(self):
        modelo_fup = {
            "" : [],
        }
        return modelo_fup
    
    def ceb(self):
        modelo_ceb = {
            "" : [],
        }
        return modelo_ceb
    
    def finep(self):
        modelo_finep = {
            "" : [],
        }
        return modelo_finep
    
    def ibitic(self):
        modelo_ibitic = {
            "Capa Finatec" : [],
            "Receita x Despesa" : [],
            "Exec. Receita e Despesa" : [],
            "Diárias" : [],
            "Passagens e Desp. Locomoção" : [],
            "Material de Consumo" : [],
            "Serviços de Terceiros PF" : [],
            "Obrig Tribut e Contributivas" : [],
            "Bolsas de Pesquisa" : [],
            "Outros Serviços de Terceiros PJ" : [],
            "Equip. Material Permanente" : [],
            "Relação de Bens" : [],
            "Rendimento de Aplicação" : [],
            "Conciliação Bancária" : [],
        }
        return modelo_ibitic
    
    def anp(self):
        modelo_anp = {
            "" : [],
        }
        return modelo_anp
