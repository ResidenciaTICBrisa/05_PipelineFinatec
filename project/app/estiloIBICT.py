import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment,NamedStyle,Border, Side
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
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
    size = tamanho + 16
    cinza = "d9d9d9"
    cinza_escuro = "bfbfbf"
    azul = "336394"
    azul_claro = '1c8cbc'

    borda = Border(right=Side(border_style="dashed"))
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
    worksheet.merge_cells('A7:J8')
    if nomeSheet == "diarias":
        worksheet['A7'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -DIÁRIAS E PASSAGENS'
    elif nomeSheet == "custoOperacional":
        worksheet['A7'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - CUSTO OPERACIONAL'
    elif nomeSheet == "pessoaJuridica":
        worksheet['A7'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - OUTROS SERVIÇOS DE TERCEIROS - PESSOA JURÍDICA'
    elif nomeSheet == "bolsaExtensao":
        worksheet['A7'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  BOLSA DE PESQUISA'
    elif nomeSheet == "materialDeConsumo":
        worksheet['A7'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S -  MATERIAL DE CONSUMO'
    elif nomeSheet == "evento":
        worksheet['A7'] = f'R E L A Ç Ã O   D E   P A G A M E N T O S - EVENTOS'

    
        # List of image names
    image_names = [
                'finatec.png',
                'ibict.png'
    ]

    # Path to the images
    path = 'C:\\Users\\Softex\\Desktop\\entrega29\\'
   
    # List to hold Image objects
    images = []

    # Loop through the list of image names and create Image objects
    for i, name in enumerate(image_names):
        image_path = path + name
        pil_image = PILImage.open(image_path)
        pil_image.save(image_path)
        img = Image(image_path)
        images.append(img)

    worksheet.add_image(images[1], "A1")#ibict
    worksheet.add_image(images[0], "H1")#finatec
  

    worksheet['A7'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A7'].alignment = Alignment(horizontal="center",vertical="center")
    worksheet['A7'].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")
    
    worksheet.merge_cells('A9:F9')
    worksheet['A9'] = "='Receita x Despesa'!A9:J9"
    worksheet['A9'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A9'].alignment = Alignment(horizontal="left",vertical="center")

    worksheet.merge_cells('A10:F10')
    worksheet['A10'] = "='Receita x Despesa'!A10:J10"
    worksheet['A10'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A10'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A11:F11')
    worksheet['A11'] = "='Receita x Despesa'!A11:J11"
    worksheet['A11'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A11'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A12:F12')
    worksheet['A12'] = "='Receita x Despesa'!A12:J12"
    worksheet['A12'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A12'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A13:F13')
    worksheet['A13'] = "='Receita x Despesa'!A13:J13"
    worksheet['A13'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A13'].alignment = Alignment(horizontal="left",vertical="center")
    
      #variavel
  
    input2=f'rowStyle{nomeVariavel}'
   

    #colunas azul cabecario
    locals()[input2] = NamedStyle(name=f'{input2}')
    locals()[input2].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    locals()[input2].fill = openpyxl.styles.PatternFill(start_color=azul_claro, end_color=azul_claro, fill_type='solid')
    locals()[input2].alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    locals()[input2].border = Border(top=Side(border_style="thin")  ,bottom=Side(border_style="thin") )
    locals()[input2].height = 20
    linha_number = 15
    for row in worksheet.iter_rows(min_row=linha_number, max_row=linha_number, min_col=1, max_col=10):
        for cell in row:
            cell.style = locals()[input2]
            if cell.column == 10:
                cell.border = Border(top=Side(border_style="thin")  ,bottom=Side(border_style="thin"), right=Side(border_style="thin") )

    valores = ["ITEM","NOME","CNPJ/CPF",'ESPECIFICAÇÃO DA DESPESA','DESCRIÇÃO',"Nº DO RECIBO OU EQUIVALENTE","DATA DE EMISSÃO",'CHEQUE / ORDEM BANCÁRIA','DATA DE PGTO','Valor']
    col = 1
    for a,b in enumerate(valores):
        worksheet.cell(row=linha_number, column=col, value=b)
        col = col + 1


    #Aumentar  a altura das celulas 
    for row in worksheet.iter_rows(min_row=16, max_row=size, min_col=1, max_col=10):
        worksheet.row_dimensions[row[0].row].height = 60
    input3 = f'customNumber{nomeVariavel}'
    
    # MASCARA R$
   
    locals()[input3] = NamedStyle(name=f'{input3}')
    locals()[input3].number_format = 'R$ #,##0.00'
    locals()[input3].font = Font(name="Arial", size=12, color="000000")
    locals()[input3].alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    
    #estilocinzasimcinzanao
    value_to_stop = size  
    start_row = 16
#
    for row in range(start_row,size+1):
        cell = worksheet[f'J{row}']
        cell.style = locals()[input3]
        
    for rows in worksheet.iter_rows(min_row=16, max_row=size, min_col=1, max_col=10):
            for cell in rows:
                if cell.row % 2:
                    cell.fill = PatternFill(start_color=cinza, end_color=cinza,
                                            fill_type = "solid")
                if cell.column == 10:
                    cell.font = Font(name="Arial", size=12, color="000000")
                    cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                    cell.border = Border(top=Side(border_style="hair") ,left = Side(border_style="hair") ,right =Side(border_style="dashed") ,bottom=Side(border_style="hair") )
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
    top_left_cell.border = Border(top=Side(border_style="thin") ,left = Side(border_style="thin") ,right =Side(border_style="thin") ,bottom=Side(border_style="thin") )

    worksheet.row_dimensions[size+2].height = 56.25

     # FORMULATOTAL
    formula = f"=SUM(J10:J{size})"
    celula = f'J{size+2}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza, end_color=cinza,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet[celula].border = Border(top=Side(border_style="thin") ,left = Side(border_style="thin") ,right =Side(border_style="dashed") ,bottom=Side(border_style="thin") )
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
    locals()[input4].border = Border(top=Side(border_style="thin") ,bottom=Side(border_style="thin") )


    row_number = size + 4
   
    for column in range(1, 11):  
        cell = worksheet.cell(row=row_number, column=column)
        cell.style = locals()[input4]
        if cell.column == 10:
            cell.border = Border(top=Side(border_style="thin") ,right =Side(border_style="thin") ,bottom=Side(border_style="thin") )



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
    top_left_subtotal2_cell.border = Border(top=Side(border_style="hair") ,left = Side(border_style="thin") ,right =Side(border_style="hair") ,bottom=Side(border_style="thin") )

    sub_formula_row_celula = f'J{sub_total2_row}'
    worksheet[sub_formula_row_celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[sub_formula_row_celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet[sub_formula_row_celula].number_format = 'R$ #,##0.00'
    worksheet[sub_formula_row_celula].border = Border(top=Side(border_style="thin") ,left = Side(border_style="thin") ,right =Side(border_style="thin") ,bottom=Side(border_style="thin") )

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
    top_left_total12_cell.border = Border(top=Side(border_style="thin") ,left = Side(border_style="thin") ,bottom=Side(border_style="thin") )


    #total_formula
    total_formula_row = size + 6
    total_formulaa = f'=J{size}'
    total_formula_row_celula = f'J{total_formula_row}'
    worksheet[total_formula_row_celula].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")
    worksheet[total_formula_row_celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet[total_formula_row_celula].number_format = 'R$ #,##0.00'
    worksheet[total_formula_row_celula].border = Border(top=Side(border_style="thin") ,bottom=Side(border_style="thin"),right=Side(border_style="thin") )

    worksheet.row_dimensions[total_formula_row].height = 30
    worksheet[total_formula_row_celula] = total_formulaa


    #valor
    

    #brasilia
    brasilia_row = size + 7
    brasilia_formula = f"='Receita x Despesa'!A44:J44"
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
    diretor_nome_formula = f"='Receita x Despesa'!A48"
    diretor_cargo_formula = f"='Receita x Despesa'!A49"
    diretor_cpf_formula = f"='Receita x Despesa'!A50"
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
    coordenadora_nome_formula = f"='Receita x Despesa'!H48"
    coordenadora_cargo_formula = f"='Receita x Despesa'!H49"
    coordenadora_cpf_formula = f"='Receita x Despesa'!H50"
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

    
    # borda = Border(right=Side(border_style="thin"))
    # worksheet.sheet_view.showGridLines = False
    # # 
    # for row in worksheet.iter_rows(min_row=1, max_row=coordenadora_cpf_row+1,min_col=10,max_col=10):
    #     for cell in row:
    #         cell.border = borda
            
    

    for row in worksheet.iter_rows(min_row=coordenadora_cpf_row+1, max_row=coordenadora_cpf_row+1,min_col=1,max_col=10):
        for cell in row:
            if cell.column == 10:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="dashed") ,bottom=Side(border_style="dashed") )
            else:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="none") ,bottom=Side(border_style="dashed") )

    bord = Border(right=Side(border_style="dashed"))
   
    # 
    for row in worksheet.iter_rows(min_row=1, max_row=14,min_col=10,max_col=10):
        for cell in row:
            cell.border = bord
   
    workbook.save(tabela)
    workbook.close()




def estilo_conciliacoes_bancaria(tabela,tamanho,tamanho2):
    
  
    caminho = pegar_caminho(tabela)
    workbook = openpyxl.load_workbook(caminho)
    worksheet = workbook['Conciliação Bancária']

   
    size = tamanho + 21
    #worksheet.row_dimensions[27].height = 50
    cinza = "d9d9d9"
    cinza_escuro = "bfbfbf"
    azul = "336394"
    azul_claro = '1c8cbc'

    borda = Border(right=Side(border_style="dashed"))
    worksheet.sheet_view.showGridLines = False
    # 

       

    # List of image names
    image_names = [
                'finatec.png',
                'ibict.png'
    ]

    # Path to the images
    path = 'C:\\Users\\Softex\\Desktop\\entrega29\\'
   
    # List to hold Image objects
    images = []

    # Loop through the list of image names and create Image objects
    for i, name in enumerate(image_names):
        image_path = path + name
        pil_image = PILImage.open(image_path)
        pil_image.save(image_path)
        img = Image(image_path)
        images.append(img)

    worksheet.add_image(images[1], "A1")#ibict
    worksheet.add_image(images[0], "F1")#finatec


    worksheet.column_dimensions['a'].width = 25
    worksheet.column_dimensions['b'].width = 25
    worksheet.column_dimensions['c'].width = 35
    worksheet.column_dimensions['d'].width = 35
    worksheet.column_dimensions['e'].width = 20
    worksheet.column_dimensions['f'].width = 20
   

    #cabecario relação de pagamentos - outro servicoes de terceiros
    worksheet.merge_cells('A7:F8')
    worksheet['A7'] = f'C O N C I L I A Ç Ã O   B A N C Á R I A'
    worksheet['A7'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A7'].alignment = Alignment(horizontal="center",vertical="center")
    worksheet['A7'].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")
    
    worksheet.merge_cells('A9:F9')
    worksheet['A9'] = "='Receita x Despesa'!A9:J9"
    worksheet['A9'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A9'].alignment = Alignment(horizontal="left",vertical="center")

    worksheet.merge_cells('A10:F10')
    worksheet['A10'] = "='Receita x Despesa'!A10:J10"
    worksheet['A10'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A10'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A11:F11')
    worksheet['A11'] = "='Receita x Despesa'!A11:J11"
    worksheet['A11'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A11'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A12:F12')
    worksheet['A12'] = "='Receita x Despesa'!A12:J12"
    worksheet['A12'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A12'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A13:F13')
    worksheet['A13'] = "='Receita x Despesa'!A13:J13"
    worksheet['A13'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A13'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A17:F17')
    worksheet['A17'] = '1.Saldo conforme extratos bancários na data final do período'
    worksheet['A17'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A17'].alignment = Alignment(horizontal="left",vertical="center")
    worksheet['A17'].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")

    worksheet.merge_cells('A18:E18')
    worksheet['A18'] = 'Saldo de Conta Corrente(R$)'
    worksheet['A18'].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet['A18'].alignment = Alignment(horizontal="right",vertical="center")

    worksheet.merge_cells('A19:E19')
    worksheet['A19'] = 'Saldo de Aplicações Financeiras(R$)'
    worksheet['A19'].font = Font(name="Arial", size=12, color="000000",bold=True)
    worksheet['A19'].alignment = Alignment(horizontal="right",vertical="center")

    worksheet.merge_cells('A21:F21')
    worksheet['A21'] = '2. Restituições não creditadas pelo banco até a data final do período'
    worksheet['A21'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A21'].alignment = Alignment(horizontal="left",vertical="center")
    worksheet['A21'].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")

    for i in range(21,size):
        sttring = f"D{i}:F{i}"
        worksheet.merge_cells(sttring)
        
    for i in range(size+3,size+3+tamanho2):
        sttring = f"D{i}:F{i}"
        worksheet.merge_cells(sttring)

    custom_number_format_conciliacoes = []
    # MASCARA R$
    if custom_number_format_conciliacoes!= False: 
        custom_number_format_conciliacoes = NamedStyle(name='custom_number_format_conciliacoes')
        custom_number_format_conciliacoes.number_format = 'R$ #,##0.00'
        custom_number_format_conciliacoes.font = Font(name="Arial", size=12, color="000000")
        custom_number_format_conciliacoes.alignment = Alignment(horizontal="general",vertical="bottom",wrap_text=True)
    
    #stylecinza
    start_row = 22
    for row in range(start_row,size+1):
        cell = worksheet[f'B{row}']
        cell.style = custom_number_format_conciliacoes
        
    for rows in worksheet.iter_rows(min_row=22, max_row=size, min_col=1, max_col=6):
            for cell in rows:
                if cell.row % 2:
                    cell.fill = PatternFill(start_color=cinza, end_color=cinza,
                                            fill_type = "solid")
                cell.font = Font(name="Arial", size=12, color="000000")
                cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)


    row_number = 22
    values = ["Data","Valor(R$)","Documento",'Descrição']
    coluna = 1
    for a,b in enumerate(values):
        worksheet.cell(row=row_number, column=coluna, value=b)
        
        coluna = coluna + 1

    # FORMULATOTAL
    formula = f"=SUM(B16:B{size-1})"
    celula = f'B{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #Total
    celula_total = F'A{size}'
    worksheet[celula_total] = f'TOTAL'
    worksheet[celula_total].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula_total].font = Font(name="Arial", size=12, color="000000",bold=True)
    #'3. Restituições não creditadas pelo banco até a data final do período'
    string_reituicoes_creditadas = f'A{size+2}:F{size+2}'
    row_creditadas = f'A{size+2}'
    worksheet.merge_cells(string_reituicoes_creditadas)
    
    worksheet[row_creditadas] = '3. Restituições creditadas pelo banco até a data final do período'
    worksheet[row_creditadas].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet[row_creditadas].alignment = Alignment(horizontal="left",vertical="center")
    worksheet[row_creditadas].fill = PatternFill(start_color=azul_claro, end_color=azul_claro,fill_type = "solid")

    #data valor documento descrição
    row_number = size+3
    values = ["Data","Valor(R$)","Documento",'Descrição']
    coluna = 1
    for a,b in enumerate(values):
        worksheet.cell(row=row_number, column=coluna, value=b)
        coluna = coluna + 1

    for rows in worksheet.iter_rows(min_row=15, max_row=15, min_col=1, max_col=6):  
            for cell in rows:
                cell.font = Font(name="Arial", size=12, color="000000",bold=True)
    for rows in worksheet.iter_rows(min_row=row_number, max_row=row_number, min_col=1, max_col=6):
            for cell in rows:
                cell.font = Font(name="Arial", size=12, color="000000",bold=True)
                


    for row in range(size+4,size+4+tamanho2):
        cell = worksheet[f'B{row}']
        cell.style = custom_number_format_conciliacoes
        
    for rows in worksheet.iter_rows(min_row=size+3, max_row=size+3+tamanho, min_col=1, max_col=6):
            for cell in rows:
                if cell.row % 2:
                    cell.fill = PatternFill(start_color=cinza, end_color=cinza,
                                            fill_type = "solid")
                cell.font = Font(name="Arial", size=12, color="000000")
                cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                if cell.column == 6: 
                        cell.border = Border(top=Side(border_style="hair") ,left = Side(border_style="hair") ,right =Side(border_style="dashed") ,bottom=Side(border_style="hair") )


    # FORMULATOTALrestituição
    formula = f"=SUM(B{size+4}:B{size+tamanho2+3})"
    celula = f'B{size+tamanho2+4}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #Total
    celula_total = F'A{size+tamanho2+4}'
    worksheet[celula_total] = f'TOTAL'
    worksheet[celula_total].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula_total].font = Font(name="Arial", size=12, color="000000",bold=True)
    #Saldo disponível p/ período seguinte (1 +2 - 3)
    string_saldo_disponivel = f'A{size+3+tamanho2+3}:D{size+3+tamanho2+3}'
    celula_string_saldo = f'A{size+tamanho2+6}'
    worksheet[celula_string_saldo].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet.merge_cells(string_saldo_disponivel)
    worksheet[celula_string_saldo]= f'Saldo disponível p/ período seguinte (1 + 2 - 3)'
    #total saldo diposnivel
    string_merge_saldo_disponivel = f'E{size+3+tamanho2+3}:F{size+3+tamanho2+3}'
    celula_string_total = f'E{size+tamanho2+6}'
    worksheet.merge_cells(string_merge_saldo_disponivel)
    saldodiposnivelformat_conciliacoes = NamedStyle(name='saldodiposnivelformat_conciliacoes')
    saldodiposnivelformat_conciliacoes.number_format = 'R$ #,##0.00'
    saldodiposnivelformat_conciliacoes.font = Font(name="Arial", size=12, color="000000")
    saldodiposnivelformat_conciliacoes.alignment = Alignment(horizontal="general",vertical="bottom",wrap_text=True)
    saldodiposnivelformat_conciliacoes.fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    celular = worksheet[celula_string_total]
    celular.style = saldodiposnivelformat_conciliacoes
    celular.value = f'=F10+F11+B{size} -B{size+tamanho2+4}'

     #brasilia
    brasilia_row = size + tamanho2+ 8
    brasilia_formula = f"='Receita x Despesa'!A44:I44"
    brasilia_merge_cells = f'A{brasilia_row}:F{brasilia_row}'
    worksheet.merge_cells(brasilia_merge_cells)
    top_left_brasilia_cell_formula = f'A{brasilia_row}'
    top_left_brasilia_cell = worksheet[top_left_brasilia_cell_formula]
    top_left_brasilia_cell.value = brasilia_formula
    top_left_brasilia_cell.alignment = Alignment(horizontal="center",vertical="center")

    # #DiretorFinanceiro
    diretor_row = size + 10 + tamanho2
    diretor_cargo_row = size + 11 + tamanho2
    diretor_cpf_row = size + 12 + tamanho2
    diretor_nome_formula = f"='Receita x Despesa'!A48"
    diretor_cargo_formula = f"='Receita x Despesa'!A49"
    diretor_cpf_formula = f"='Receita x Despesa'!A50"
    diretor_merge_cells = f'A{diretor_row}:B{diretor_row}'
    diretor_cargo_merge_cells = f'A{diretor_cargo_row}:B{diretor_cargo_row}'
    diretor_cpf_merge_cells = f'A{diretor_cpf_row}:B{diretor_cpf_row}'
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
    coordenadora_row = size + tamanho2 + 10
    coordenadora_cargo_row = size + 11 + tamanho2
    coordenadora_cpf_row = size + 12+ tamanho2
    coordenadora_nome_formula = f"='Receita x Despesa'!G48:H48"
    coordenadora_cargo_formula = f"='Receita x Despesa'!H49"
    coordenadora_cpf_formula = f"='Receita x Despesa'!H50"
    coordenadora_merge_cells = f'D{coordenadora_row}:F{coordenadora_row}'
    coordenadora_cargo_merge_cells = f'D{coordenadora_cargo_row}:F{coordenadora_cargo_row}'
    coordenadora_cpf_merge_cells = f'D{coordenadora_cpf_row}:F{coordenadora_cpf_row}'
    worksheet.merge_cells(coordenadora_merge_cells)
    worksheet.merge_cells(coordenadora_cargo_merge_cells)
    worksheet.merge_cells(coordenadora_cpf_merge_cells)
    top_left_coordenadora_cell_formula = f'D{coordenadora_row}'
    top_left_coordenadora_cell_cargo_formula = f'D{coordenadora_cargo_row}'
    top_left_coordenadora_cell_cpf_formula = f'D{coordenadora_cpf_row}'
    top_left_coordenadora_cell = worksheet[top_left_coordenadora_cell_formula]
    top_left_coordenadora_cell_cargo_formula = worksheet[top_left_coordenadora_cell_cargo_formula]
    top_left_coordenadora_cell_cpf_formula = worksheet[top_left_coordenadora_cell_cpf_formula]
    top_left_coordenadora_cell.value = coordenadora_nome_formula
    top_left_coordenadora_cell_cargo_formula.value = coordenadora_cargo_formula
    top_left_coordenadora_cell_cpf_formula.value = coordenadora_cpf_formula
    top_left_coordenadora_cell.alignment = Alignment(horizontal="center",vertical="center")
    top_left_coordenadora_cell.border = borda
    top_left_coordenadora_cell.font= Font(bold = True)
    top_left_coordenadora_cell_cargo_formula.alignment = Alignment(horizontal="center",vertical="center")
    top_left_coordenadora_cell_cargo_formula.border = borda
    top_left_coordenadora_cell_cpf_formula.alignment = Alignment(horizontal="center",vertical="center")
    top_left_coordenadora_cell_cpf_formula.border = borda

    for row in worksheet.iter_rows(min_row=1, max_row=size+11,min_col=6,max_col=6):
        for cell in row:
            cell.border = borda

    for row in worksheet.iter_rows(min_row=coordenadora_cpf_row+1, max_row=coordenadora_cpf_row+1,min_col=1,max_col=6):
        for cell in row:
            if cell.column == 6:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="dashed") ,bottom=Side(border_style="dashed") )
            else:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="none") ,bottom=Side(border_style="dashed") )

    workbook.save(tabela)
    workbook.close()

def estilo_rendimento_de_aplicacao(tabela,tamanho):
    caminho = pegar_caminho(tabela)
    workbook = openpyxl.load_workbook(caminho)
    worksheet = workbook['Rendimento de Aplicação']

   
    size = tamanho + 16
    worksheet.row_dimensions[10].height = 2
    worksheet.row_dimensions[9].height = 20

    cinza = "d9d9d9"
    cinza_escuro = "bfbfbf"
    azul = "336394"
    azul_claro = '0198cc'
    borda = Border(right=Side(border_style="dashed"))
    worksheet.sheet_view.showGridLines = False
    # 
    for row in worksheet.iter_rows(min_row=1, max_row=size+9,min_col=8,max_col=8):
        for cell in row:
            cell.border = borda

    

    worksheet.column_dimensions['a'].width = 20
    worksheet.column_dimensions['b'].width = 20
    worksheet.column_dimensions['c'].width = 20
    worksheet.column_dimensions['d'].width = 20
    worksheet.column_dimensions['e'].width = 20
    worksheet.column_dimensions['f'].width = 20
    worksheet.column_dimensions['g'].width = 20
    worksheet.column_dimensions['h'].width = 20
    
    worksheet.add_image('finatec.png', "A1")#ibict
    worksheet.add_image('ibict.png', "F1")#finate

    #cabecario relação de pagamentos - outro servicoes de terceiros
    worksheet.merge_cells('A7:H8')
    worksheet['A7'] = f'Demonstrativo dos Ganhos Auferidos com Aplicações Financeiras'
    worksheet['A7'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A7'].alignment = Alignment(horizontal="center",vertical="center")
    worksheet['A7'].fill = PatternFill(start_color=azul, end_color=azul,fill_type = "solid")
    
    worksheet.merge_cells('A9:H9')
    worksheet['A9'] = "='Receita x Despesa'!A9:J9"
    worksheet['A9'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A9'].alignment = Alignment(horizontal="left",vertical="center")

    worksheet.merge_cells('A10:H10')
    worksheet['A10'] = "='Receita x Despesa'!A10:J10"
    worksheet['A10'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A10'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A11:H11')
    worksheet['A11'] = "='Receita x Despesa'!A11:J11"
    worksheet['A11'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A11'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A12:H12')
    worksheet['A12'] = "='Receita x Despesa'!A12:J12"
    worksheet['A12'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A12'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A13:H13')
    worksheet['A13'] = "='Receita x Despesa'!A13:J13"
    worksheet['A13'].font = Font(name="Arial", size=12, color="000000")
    worksheet['A13'].alignment = Alignment(horizontal="left",vertical="center")
    
    worksheet.merge_cells('A14:H14')
    worksheet['A14'] = 'BB CP Corpor Àgil - CNPJ 11.351.449-0001-10L'
    worksheet['A14'].font = Font(name="Arial", size=12, color="FFFFFF",bold=True)
    worksheet['A14'].alignment = Alignment(horizontal="center",vertical="center")
    worksheet['A14'].fill = PatternFill(start_color=azul, end_color=azul,fill_type = "solid")
    

    

    #stylecinza
    start_row = 11
    for rows in worksheet.iter_rows(min_row=start_row, max_row=13, min_col=1, max_col=8):
            for cell in rows:
                if cell.row % 2:
                    cell.fill = PatternFill(start_color=azul_claro, end_color=azul_claro,
                                            fill_type = "solid")
                cell.font = Font(name="Arial", size=12, color="000000",bold=True)
                cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)


    row_number = 11
    values = ["Período","Saldo Anterior","Valor Aplicado no período",'Valor Resgatado no Período','Rendimento Bruto','Imposto de Renda / IOF','Rendimento Líquido','Saldo']
    coluna = 1
    for a,b in enumerate(values):
        worksheet.cell(row=row_number, column=coluna, value=b)
        
        coluna = coluna + 1

    for i in range(1,9):
        worksheet.merge_cells(start_row=11,end_row=13,start_column=i,end_column=i)
    #BARRAS DE DADOS
    start_row = 14
    for rows in worksheet.iter_rows(min_row=start_row, max_row=size, min_col=1, max_col=8):
            for cell in rows:
                if cell.row % 2:
                    cell.fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,
                                            fill_type = "solid")
                cell.font = Font(name="Arial", size=12, color="000000")
                cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    #MASCARA VERMELHO
    for rows in worksheet.iter_rows(min_row=start_row, max_row=size-1, min_col=6, max_col=6):
            for cell in rows:
                cell.font = Font(name="Arial", size=12, color="f90000")
    #MASCARANEGRITO
    for rows in worksheet.iter_rows(min_row=start_row, max_row=size-1, min_col=1, max_col=1):
            for cell in rows:
                cell.font = Font(name="Arial", size=12, color="000000",bold=True)
    #MASCARA AZUL
    for rows in worksheet.iter_rows(min_row=start_row, max_row=size-1, min_col=7, max_col=7):
            for cell in rows:
                cell.font = Font(name="Arial", size=12, color="141fca",bold=True)

    
    #barra de totais
    # FORMULATOTAL
     #C
    formula = f"=SUM(C14:C{size-1})"
    celula = f'C{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #D
    formula = f"=SUM(D14:D{size-1})"
    celula = f'D{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #E
    formula = f"=SUM(E14:E{size-1})"
    celula = f'E{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #F
    formula = f"=SUM(F14:F{size-1})"
    celula = f'F{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #G
    formula = f"=SUM(G14:G{size-1})"
    celula = f'G{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)
    #H
    formula = f"=SUM(H14:H{size-1})"
    celula = f'H{size}'
    worksheet[celula] = formula
    worksheet[celula].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula].font = Font(name="Arial", size=12, color="000000",bold=True)

    #Total
    celula_total = F'A{size}'
    worksheet[celula_total] = f'TOTAL'
    worksheet[celula_total].fill = PatternFill(start_color=cinza_escuro, end_color=cinza_escuro,fill_type = "solid")
    worksheet[celula_total].font = Font(name="Arial", size=12, color="000000",bold=True)    

    #brasilia
    brasilia_row = size + 2
    brasilia_formula = f"='Receita x Despesa'!A42:F42"
    brasilia_merge_cells = f'A{brasilia_row}:F{brasilia_row}'
    worksheet.merge_cells(brasilia_merge_cells)
    top_left_brasilia_cell_formula = f'A{brasilia_row}'
    top_left_brasilia_cell = worksheet[top_left_brasilia_cell_formula]
    top_left_brasilia_cell.value = brasilia_formula
    top_left_brasilia_cell.alignment = Alignment(horizontal="center",vertical="center")

    # #DiretorFinanceiro
    diretor_row = size + 6 
    diretor_cargo_row = size + 7 
    diretor_cpf_row = size + 8
    diretor_nome_formula = f"='Receita x Despesa'!A45"
    diretor_cargo_formula = f"='Receita x Despesa'!A46"
    diretor_cpf_formula = f"='Receita x Despesa'!A47"
    diretor_merge_cells = f'A{diretor_row}:C{diretor_row}'
    diretor_cargo_merge_cells = f'A{diretor_cargo_row}:C{diretor_cargo_row}'
    diretor_cpf_merge_cells = f'A{diretor_cpf_row}:C{diretor_cpf_row}'
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
    coordenadora_row = size  + 6
    coordenadora_cargo_row = size + 7 
    coordenadora_cpf_row = size + 8
    coordenadora_nome_formula = f"='Receita x Despesa'!H45"
    coordenadora_cargo_formula = f"='Receita x Despesa'!H46"
    coordenadora_cpf_formula = f"='Receita x Despesa'!H47"
    coordenadora_merge_cells = f'E{coordenadora_row}:G{coordenadora_row}'
    coordenadora_cargo_merge_cells = f'E{coordenadora_cargo_row}:G{coordenadora_cargo_row}'
    coordenadora_cpf_merge_cells = f'E{coordenadora_cpf_row}:G{coordenadora_cpf_row}'
    worksheet.merge_cells(coordenadora_merge_cells)
    worksheet.merge_cells(coordenadora_cargo_merge_cells)
    worksheet.merge_cells(coordenadora_cpf_merge_cells)
    top_left_coordenadora_cell_formula = f'E{coordenadora_row}'
    top_left_coordenadora_cell_cargo_formula = f'E{coordenadora_cargo_row}'
    top_left_coordenadora_cell_cpf_formula = f'E{coordenadora_cpf_row}'
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


    for row in worksheet.iter_rows(min_row=coordenadora_cpf_row+1, max_row=coordenadora_cpf_row+1,min_col=1,max_col=8):
        for cell in row:
            if cell.column == 4:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="dashed") ,bottom=Side(border_style="dashed") )
            else:
                cell.border = Border(top=Side(border_style="none") ,left = Side(border_style="none") ,right =Side(border_style="none") ,bottom=Side(border_style="dashed") )


    workbook.save(tabela)
    workbook.close()



tabela = pegar_caminho('IBICT.xlsx')
nomeTabela ="Evento"
tituloStyle = "evento"
workbook = openpyxl.load_workbook(tabela)
sheet2 = workbook.create_sheet(title="Evento")
workbook.save("tabelapreenchida.xlsx")
workbook.close()
tabela = pegar_caminho("tabelapreenchida.xlsx")
workbook = openpyxl.load_workbook(tabela)
sheet2 = workbook.create_sheet(title="Conciliação Bancária")
workbook.save(tabela)
workbook.close()
maior = 20
maior2 = 20
tabela2 = pegar_caminho('tabelapreenchida.xlsx')
print(tabela2)
estiloGeral(tabela2,maior,tituloStyle,nomeTabela)
estilo_conciliacoes_bancaria(tabela2,maior2,maior)