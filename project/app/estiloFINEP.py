import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment,NamedStyle,Border, Side
import os
#pegar o caminho do arquivo
def pegar_caminho(nome_arquivo):

    # Obter o caminho absoluto do arquivo Python em execução
    caminho_script = os.path.abspath(__file__)

    # Obter o diretório da pasta onde o script está localizado
    pasta_script = os.path.dirname(caminho_script)

    # Combinar o caminho da pasta com o nome do arquivo Excel
    caminho = os.path.join(pasta_script, nome_arquivo)

    return caminho

def estiloGeral(tabela,tamanho,nomeVariavel,nomeTabela):
    nomeSheet=nomeVariavel
    caminho = pegar_caminho(tabela)
    workbook = openpyxl.load_workbook(caminho)
    worksheet = workbook[nomeTabela]
    size = tamanho + 10
    cinza = "d9d9d9"
    cinza_escuro = "bfbfbf"
    azul_claro = 'cdfeff'

    borda = Border(right=Side(border_style="medium"))
    worksheet.sheet_view.showGridLines = False
    # 
    for row in worksheet.iter_rows(min_row=1, max_row=size+11,min_col=10,max_col=10):
        for cell in row:
            cell.border = borda
            

    worksheet.column_dimensions['a'].width = 25
    worksheet.column_dimensions['b'].width = 25
    worksheet.column_dimensions['c'].width = 35
    worksheet.column_dimensions['d'].width = 35#descrição
    worksheet.column_dimensions['e'].width = 65 #n do recibo ou qeuivalente
    worksheet.column_dimensions['f'].width = 25 #data de emissão
    worksheet.column_dimensions['g'].width = 25 #data de emissão
    worksheet.column_dimensions['h'].width = 25 #data de emissão
    worksheet.column_dimensions['i'].width = 25 #data de emissão
    worksheet.column_dimensions['j'].width = 25 #data de emissão


    #cabecario relação de pagamentos - outro servicoes de terceiros
    worksheet.merge_cells('A1:J2')
    if nomeSheet == "diarias":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - DIÁRIAS'
    elif nomeSheet == "pessoaFisica":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  OUTROS SERVIÇOS DE TERCEIROS - PF'
    elif nomeSheet == "pessoaJuridica":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - OUTROS SERVIÇOS DE TERCEIROS - PESSOA JURÍDICA'
    elif nomeSheet == "passagenDespLocomo":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - PASSAGENS E DESPESAS COM LOCOMOÇÃO'
    elif nomeSheet == "outrosServiçosTerceiros":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - O U T R O S  S E R V I Ç O S D E T E R C E I R O S - C E L E T I S T A S'
    elif nomeSheet == "auxilioEstudante":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  AUXÍLIO FINANCEIRO A ESTUDANTE'
    elif nomeSheet == "bolsaExtensao":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  BOLSA DE EXTENSÃO'
    elif nomeSheet == "estagiario":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  ESTAGIÁRIO'
    elif nomeSheet == "custosIndiretos":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - CUSTOS INDIRETOS - FUB'
    elif nomeSheet == "materialDeConsumo":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  MATERIAL DE CONSUMO'
    elif nomeSheet == "equipamentoMaterialPermanente":
        worksheet['A1'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  EQUIPAMENTO E MATERIAL PERMANENTE'
  

    worksheet['A1'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A1'].alignment = Alignment(horizontal="center",vertical="center")
    worksheet['A1'].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")
    
    worksheet.merge_cells('A3:F3')
    worksheet['A3'] = "='Receita x Despesa'!A3:J3"
    worksheet['A3'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A3'].alignment = Alignment(horizontal="left",vertical="center")

    worksheet.merge_cells('A4:F4')
    worksheet['A4'] = "='Receita x Despesa'!A4:J4"
    worksheet['A4'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A4'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A5:F5')
    worksheet['A5'] = "='Receita x Despesa'!A5:J5"
    worksheet['A5'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A5'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A6:F6')
    worksheet['A6'] = "='Receita x Despesa'!A6:J6"
    worksheet['A6'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A6'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A7:F7')
    worksheet['A7'] = "='Receita x Despesa'!A7:J7"
    worksheet['A7'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A7'].alignment = Alignment(horizontal="left",vertical="center")
    
    #variavel
  
    input2=f'rowStyle{nomeVariavel}'
   

    #colunas azul cabecario
    locals()[input2] = NamedStyle(name=f'{input2}')
    locals()[input2].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    locals()[input2].fill = openpyxl.styles.PatternFill(start_color=azul_claro, end_color=azul_claro, fill_type='solid')
    locals()[input2].alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    locals()[input2].border = Border(top=Side(border_style="medium")  ,bottom=Side(border_style="thin") )
    locals()[input2].height = 20
    linha_number = 9
    for row in worksheet.iter_rows(min_row=linha_number, max_row=linha_number, min_col=1, max_col=10):
        for cell in row:
            cell.style = locals()[input2]
            if cell.column == 10:
                cell.border = Border(top=Side(border_style="medium")  ,bottom=Side(border_style="thin"), right=Side(border_style="medium") )

    valores = ["ITEM","NOME","CNPJ/CPF",'ESPECIFICAÇÃO DA DESPESA','DESCRIÇÃO',"Nº DO RECIBO OU EQUIVALENTE","DATA DE EMISSÃO",'CHEQUE / ORDEM BANCÁRIA','DATA DE PGTO','Valor']
    col = 1
    for a,b in enumerate(valores):
        worksheet.cell(row=linha_number, column=col, value=b)
        col = col + 1


    #Aumentar  a altura das celulas 
    for row in worksheet.iter_rows(min_row=10, max_row=size, min_col=1, max_col=10):
        worksheet.row_dimensions[row[0].row].height = 60
    input3 = f'customNumber{nomeVariavel}'
    
    # MASCARA R$
   
    locals()[input3] = NamedStyle(name=f'{input3}')
    locals()[input3].number_format = 'R$ #,##0.00'
    locals()[input3].font = Font(name="Arial", size=12, color="000000")
    locals()[input3].alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    
    #estilocinzasimcinzanao
    value_to_stop = size  
    start_row = 10
#
    for row in range(start_row,size+1):
        cell = worksheet[f'J{row}']
        cell.style = locals()[input3]
        
    for rows in worksheet.iter_rows(min_row=10, max_row=size, min_col=1, max_col=10):
            for cell in rows:
                if cell.row % 2:
                    cell.fill = PatternFill(start_color=cinza, end_color=cinza,
                                            fill_type = "solid")
                if cell.column == 10:
                    cell.font = Font(name="Arial", size=12, color="000000")
                    cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                    cell.border = Border(top=Side(border_style="hair") ,left = Side(border_style="hair") ,right =Side(border_style="medium") ,bottom=Side(border_style="hair") )
                else:
                    cell.font = Font(name="Arial", size=12, color="000000")
                    cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                    cell.border = Border(top=Side(border_style="hair") ,left = Side(border_style="hair") ,right =Side(border_style="hair") ,bottom=Side(border_style="hair") )
                
                
    #subtotal
    stringAfinarCelula =size+2
    worksheet.row_dimensions[size+2].height = 6
    celulas_mergidas_subtotal = f"A{size+2}:I{size+2}"
    worksheet.merge_cells(celulas_mergidas_subtotal)
    left_celula_cell = f"A{size+2}"
    top_left_cell = worksheet[left_celula_cell]
    top_left_cell.value = "Sub Total1"
    top_left_cell.fill = PatternFill(start_color=cinza, end_color=cinza,fill_type = "solid")
    top_left_cell.font = Font(name="Arial", size=12, color="000000",bold=True)
    top_left_cell.alignment = Alignment(horizontal="center",vertical="center")
    top_left_cell.border = Border(top=Side(border_style="thin") ,left = Side(border_style="medium") ,right =Side(border_style="thin") ,bottom=Side(border_style="medium") )

    worksheet.row_dimensions[size+2].height = 56.25

     # FORMULATOTAL
    formula = f"=SUM(J10:J{size})"
    celula = f'J{size+2}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza, end_color=cinza,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet[celula].border = Border(top=Side(border_style="thin") ,left = Side(border_style="thin") ,right =Side(border_style="medium") ,bottom=Side(border_style="medium") )
    worksheet[celula].number_format = 'R$ #,##0.00'
    #restituições creditadas
    restituicoes = size + 3
    celula_restituicoes=f'A{restituicoes}'
    worksheet[celula_restituicoes].value = "RESTITUIÇÕES CREDITADAS"
    worksheet[celula_restituicoes].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet.row_dimensions[restituicoes].height = 30


    input4 = f'row_style_diaria_append{nomeVariavel}'
    #estilo colunas restitucoes creditadas
    locals()[input4] = NamedStyle(name=f'{input4}')
    locals()[input4].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    locals()[input4].fill = openpyxl.styles.PatternFill(start_color=azul_claro, end_color=azul_claro, fill_type='solid')
    locals()[input4].alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    locals()[input4].height = 30
    locals()[input4].border = Border(top=Side(border_style="medium") ,bottom=Side(border_style="medium") )


    row_number = size + 4
   
    for column in range(1, 11):  
        cell = worksheet.cell(row=row_number, column=column)
        cell.style = locals()[input4]
        if cell.column == 10:
            cell.border = Border(top=Side(border_style="medium") ,right =Side(border_style="medium") ,bottom=Side(border_style="medium") )



    values = ["Item","Restituidor","CNPJ/CPF",'Descrição',"Cheque equivalente","Data do Cheque",'Nº do Depósito','Data da Devolução','Valor']
    coluna = 1
    for a,b in enumerate(values):
        worksheet.cell(row=row_number, column=coluna, value=b)
        if coluna == 4:
            coluna = coluna + 1
        coluna = coluna + 1
        

    merge_formula = f'D{row_number}:E{row_number}'
    worksheet.merge_cells(merge_formula)

    
    #subtotal2
    sub_total2_row = size + 5
    subtotal_merge_cells= f'A{sub_total2_row}:I{sub_total2_row}'
    worksheet.merge_cells(subtotal_merge_cells)
    top_left_subtotal2_cell_formula = f'A{sub_total2_row}'
    top_left_subtotal2_cell = worksheet[top_left_subtotal2_cell_formula]
    top_left_subtotal2_cell.value = "Sub Total 2"
    top_left_subtotal2_cell.fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    top_left_subtotal2_cell.font = Font(name="Arial", size=12, color="000000",bold=True)
    top_left_subtotal2_cell.alignment = Alignment(horizontal="center",vertical="center")
    top_left_subtotal2_cell.border = Border(top=Side(border_style="hair") ,left = Side(border_style="medium") ,right =Side(border_style="hair") ,bottom=Side(border_style="medium") )

    sub_formula_row_celula = f'J{sub_total2_row}'
    worksheet[sub_formula_row_celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[sub_formula_row_celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet[sub_formula_row_celula].number_format = 'R$ #,##0.00'
    worksheet[sub_formula_row_celula].border = Border(top=Side(border_style="thin") ,left = Side(border_style="thin") ,right =Side(border_style="medium") ,bottom=Side(border_style="medium") )

      #total1-2
    total12_row = size + 6
    total12_merge_cells = f'A{total12_row}:I{total12_row}'
    worksheet.merge_cells(total12_merge_cells)
    top_left_total12_cell_formula = f'A{total12_row}'
    top_left_total12_cell = worksheet[top_left_total12_cell_formula]
    top_left_total12_cell.value = "Total(1-2)"
    top_left_total12_cell.fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")
    top_left_total12_cell.font = Font(name="Arial", size=12, color="000000",bold=True)
    top_left_total12_cell.alignment = Alignment(horizontal="center",vertical="center")
    top_left_total12_cell.border = Border(top=Side(border_style="medium") ,left = Side(border_style="medium") ,bottom=Side(border_style="medium") )


    #total_formula
    total_formula_row = size + 6
    total_formulaa = f'=J{size}'
    total_formula_row_celula = f'J{total_formula_row}'
    worksheet[total_formula_row_celula].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")
    worksheet[total_formula_row_celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet[total_formula_row_celula].number_format = 'R$ #,##0.00'
    worksheet[total_formula_row_celula].border = Border(top=Side(border_style="medium") ,bottom=Side(border_style="medium"),right=Side(border_style="medium") )

    worksheet.row_dimensions[total_formula_row].height = 30
    worksheet[total_formula_row_celula] = total_formulaa


    #brasilia
    brasilia_row = size + 7
    brasilia_formula = f"='Receita x Despesa'!A42:J42"
    brasilia_merge_cells = f'A{brasilia_row}:I{brasilia_row}'
    worksheet.merge_cells(brasilia_merge_cells)
    top_left_brasilia_cell_formula = f'A{brasilia_row}'
    top_left_brasilia_cell = worksheet[top_left_brasilia_cell_formula]
    top_left_brasilia_cell.value = brasilia_formula
    top_left_brasilia_cell.alignment = Alignment(horizontal="center",vertical="center")

    #DiretorFinanceiro
    diretor_row = size + 8
    diretor_cargo_row = size + 9
    diretor_cpf_row = size + 10
    diretor_nome_formula = f"='Receita x Despesa'!A45"
    diretor_cargo_formula = f"='Receita x Despesa'!A46"
    diretor_cpf_formula = f"='Receita x Despesa'!A47"
    diretor_merge_cells = f'A{diretor_row}:D{diretor_row}'
    diretor_cargo_merge_cells = f'A{diretor_cargo_row}:D{diretor_cargo_row}'
    diretor_cpf_merge_cells = f'A{diretor_cpf_row}:D{diretor_cpf_row}'
    worksheet.merge_cells(diretor_merge_cells)
    worksheet.merge_cells(diretor_cargo_merge_cells)
    worksheet.merge_cells(diretor_cpf_merge_cells)
    top_left_diretor_cell_formula = f'A{diretor_row}'
    top_left_diretor_cell_cargo_formula = f'A{diretor_cargo_row}'
    top_left_diretor_cell_cpf_formula = f'A{diretor_cpf_row}'
    top_left_diretor_cell = worksheet[top_left_diretor_cell_formula]
    top_left_diretor_cell_cargo_formula = worksheet[top_left_diretor_cell_cargo_formula]
    top_left_diretor_cell_cpf_formula = worksheet[top_left_diretor_cell_cpf_formula]
    top_left_diretor_cell.value = diretor_nome_formula
    top_left_diretor_cell_cargo_formula.value = diretor_cargo_formula
    top_left_diretor_cell_cpf_formula.value = diretor_cpf_formula
    top_left_diretor_cell.alignment = Alignment(horizontal="center",vertical="center")
    top_left_diretor_cell.font = Font(bold=True)
    top_left_diretor_cell_cargo_formula.alignment = Alignment(horizontal="center",vertical="center")
    top_left_diretor_cell_cpf_formula.alignment = Alignment(horizontal="center",vertical="center")
    #Coordenadora
    coordenadora_row = size + 8
    coordenadora_cargo_row = size + 9
    coordenadora_cpf_row = size + 10
    coordenadora_nome_formula = f"='Receita x Despesa'!H45"
    coordenadora_cargo_formula = f"='Receita x Despesa'!H46"
    coordenadora_cpf_formula = f"='Receita x Despesa'!H47"
    coordenadora_merge_cells = f'F{coordenadora_row}:J{coordenadora_row}'
    coordenadora_cargo_merge_cells = f'F{coordenadora_cargo_row}:J{coordenadora_cargo_row}'
    coordenadora_cpf_merge_cells = f'F{coordenadora_cpf_row}:J{coordenadora_cpf_row}'
    worksheet.merge_cells(coordenadora_merge_cells)
    worksheet.merge_cells(coordenadora_cargo_merge_cells)
    worksheet.merge_cells(coordenadora_cpf_merge_cells)
    top_left_coordenadora_cell_formula = f'F{coordenadora_row}'
    top_left_coordenadora_cell_cargo_formula = f'F{coordenadora_cargo_row}'
    top_left_coordenadora_cell_cpf_formula = f'F{coordenadora_cpf_row}'
    top_left_coordenadora_cell = worksheet[top_left_coordenadora_cell_formula]
    top_left_coordenadora_cell_cargo_formula = worksheet[top_left_coordenadora_cell_cargo_formula]
    top_left_coordenadora_cell_cpf_formula = worksheet[top_left_coordenadora_cell_cpf_formula]
    top_left_coordenadora_cell.value = coordenadora_nome_formula
    top_left_coordenadora_cell_cargo_formula.value = coordenadora_cargo_formula
    top_left_coordenadora_cell_cpf_formula.value = coordenadora_cpf_formula
    top_left_coordenadora_cell.alignment = Alignment(horizontal="center",vertical="center")
    top_left_coordenadora_cell.font= Font(bold = True)
    top_left_coordenadora_cell_cargo_formula.alignment = Alignment(horizontal="center",vertical="center")
    top_left_coordenadora_cell_cpf_formula.alignment = Alignment(horizontal="center",vertical="center")

    
    # borda = Border(right=Side(border_style="medium"))
    # worksheet.sheet_view.showGridLines = False
    # # 
    # for row in worksheet.iter_rows(min_row=1, max_row=coordenadora_cpf_row+1,min_col=10,max_col=10):
    #     for cell in row:
    #         cell.border = borda
            
    

    for row in worksheet.iter_rows(min_row=coordenadora_cpf_row+1, max_row=coordenadora_cpf_row+1,min_col=1,max_col=10):
        for cell in row:
            if cell.column == 10:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="medium") ,bottom=Side(border_style="medium") )
            else:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="none") ,bottom=Side(border_style="medium") )

    workbook.save(tabela)
    workbook.close()
