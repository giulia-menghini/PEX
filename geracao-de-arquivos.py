import pandas as pd
import numpy as np
from datetime import date, timedelta
import json
import datetime
import pydata_google_auth
import pandas_gbq as pdgbq
import io


    
    
clients_df = pd.read_excel("G:\\Drives compartilhados\\Gente e Gestão FVT\\FVTech\\13. PEX\\Base (Pós Rebate).xlsx", sheet_name='GERAÇÃO')


franchises = clients_df['Franquia'].unique() 

for franchise in franchises:

    fclients_df = clients_df[(clients_df['Franquia'] == franchise)] 
    
    julho_meta_lb = float(fclients_df[fclients_df['Franquia'] == franchise]['LB 1Q META'])
    agosto_meta_lb = float(fclients_df[fclients_df['Franquia'] == franchise]['LB 2Q META'])
    setembro_meta_lb = float(fclients_df[fclients_df['Franquia'] == franchise]['LB 3Q META'])
    
    julho_meta_ba = float(fclients_df[fclients_df['Franquia'] == franchise]['BA 1Q META'])
    agosto_meta_ba = float(fclients_df[fclients_df['Franquia'] == franchise]['BA 2Q META'])
    setembro_meta_ba = float(fclients_df[fclients_df['Franquia'] == franchise]['BA 3Q META'])
    
    julho_real_lb = float(fclients_df[fclients_df['Franquia'] == franchise]['LB 1Q REAL'])
    agosto_real_lb = float(fclients_df[fclients_df['Franquia'] == franchise]['LB 2Q REAL'])
    setembro_real_lb = float(fclients_df[fclients_df['Franquia'] == franchise]['LB 3Q REAL'])
    
    julho_real_ba = float(fclients_df[fclients_df['Franquia'] == franchise]['BA 1Q REAL'])
    agosto_real_ba = float(fclients_df[fclients_df['Franquia'] == franchise]['BA 2Q REAL'])
    setembro_real_ba = float(fclients_df[fclients_df['Franquia'] == franchise]['BA 3Q REAL'])
    
    media_lb_real = float(fclients_df[fclients_df['Franquia'] == franchise]['MÉDIA LB REAL'])
    media_lb_meta = float(fclients_df[fclients_df['Franquia'] == franchise]['MÉDIA LB META'])
    media_ba_real = float(fclients_df[fclients_df['Franquia'] == franchise]['MÉDIA BA REAL'])
    media_ba_meta = float(fclients_df[fclients_df['Franquia'] == franchise]['MÉDIA BA META'])
    
    media_real_ba = (julho_real_ba + agosto_real_ba + setembro_real_ba)/3
    
    media_real_lb = (julho_real_lb + agosto_real_lb + setembro_real_lb)/3
    
    media_meta_ba = (julho_meta_ba + agosto_meta_ba + setembro_meta_ba)/3
    
    media_meta_lb = (julho_meta_lb + agosto_meta_lb + setembro_meta_lb)/3
    
    real_ponderado = (media_real_ba*0.2) + (media_real_lb*0.8)
    meta_ponderado = (media_meta_ba*0.2) + (media_meta_lb*0.8)
    final_ponderado = round((real_ponderado - meta_ponderado),2)
    
    
    
    with pd.ExcelWriter(f"G:\\Drives compartilhados\\Gente e Gestão FVT\\FVTech\\13. PEX\\Arquivos\\{franchise} - PEX 3Q22.xlsx") as writer:

            workbook = writer.book


            campo = workbook.add_format(
                {
                    "font": "Segoe UI Emoji",
                    "font_size": 10,
                    "align": "left",
                    "font_color": "#000000",
                    "valign": "vcenter"
                }

            )

            campo2 = workbook.add_format(
                {
                    "font": "Sharon Sans Medium",
                    "font_size": 22,
                    "align": "center",
                    "font_color": "#000000",
                    "valign": "vcenter"
                }

            )

            campo3 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "align": "center",
                    "font_color": "#000000",
                    "valign": "vcenter"
                }

            )

            campo33 = workbook.add_format(
                {
                    "font": "Segoe UI Emoji",
                    "bold": True,
                    "font_size": 10,
                    "align": "center",
                    "font_color": "#000000",
                    "valign": "vcenter",
                    "border": 1,
                    "border_color": "#F2F2F2",
                    "bg_color": "#F2F2F2",
                }

            )

            campo4 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "align": "center",
                    "font_color": "#FFFFFF",
                    "valign": "vcenter",
                    "bg_color": "#00A868",
                }

            )

            campo5 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "align": "center",
                    "font_color": "#FFFFFF",
                    "valign": "vcenter",
                    "bg_color": "#BFBFBF",
                }
            )

            campo6 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "align": "center",
                    "font_color": "#000000",
                    "valign": "vcenter",
                    "border": 1,
                    "border_color": "#C0C0C0",
                }
            )

            campo7 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "bold": True,
                    "align": "center",
                    "font_color": "#000000",
                    "valign": "vcenter",
                    "bg_color": "#D9D9D9",
                }
            )

            campo8 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "align": "center",
                    "bold": True,
                    "font_color": "#000000",
                    "valign": "vcenter",
                    "border": 1,
                    "border_color": "#C0C0C0",
                }
            )        


            campo9 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 10,
                    "align": "center",
                    "bold": True,
                    "font_color": "#000000",
                    "valign": "vcenter",
                }
            )     


            campo10 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 10,
                    "align": "left",
                    "font_color": "#000000",
                    "valign": "vcenter",
                }
            )


            campo11 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 10,
                    "align": "left",
                    'text_wrap': True,
                    "font_color": "#000000",
                    "valign": "vcenter",
                }
            )


            campo12 = workbook.add_format(
                {
                    "font": "Calibri",
                    "font_size": 11,
                    "bold": True,
                    "align": "center",
                    "font_color": "#FFFFFF",
                    "bg_color": "#00B050",
                    "valign": "vcenter"
                }

            )       



            campo13 = workbook.add_format(
                {
                    "font": "Calibri",
                    "font_size": 11,
                    "bold": True,
                    "align": "center",
                    "font_color": "#FFFFFF",
                    "bg_color": "#C00000",
                    "valign": "vcenter"
                }
            )


            campo14 = workbook.add_format(
                {
                    "font": "Calibri",
                    "font_size": 11,
                    "bold": True,
                    "align": "center",
                    "font_color": "#FFFFFF",
                    "bg_color": "#6F6F6F",
                    "valign": "vcenter"
                }
            )
            
            
            porcentagem = workbook.add_format(
                {
                        'num_format': "0%",
                        "font": "Sharon Sans",
                        "font_size": 11,
                        "align": "center",
                        "font_color": "#000000",
                        "valign": "vcenter",
                        "border": 1,
                        "border_color": "#C0C0C0",

                }
            )

        
            porcentagem2 = workbook.add_format(
                {
                    'num_format': "0%",
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "bold": True,
                    "align": "center",
                    "font_color": "#000000",
                    "valign": "vcenter",
                    "bg_color": "#D9D9D9",
                }
            )
            
            porcentagem3 = workbook.add_format(
                {
                    "font": "Sharon Sans",
                    "font_size": 11,
                    "align": "center",
                    'num_format': "0",
                    "font_color": "#FFFFFF",
                    "valign": "vcenter",
                    "bg_color": "#00A868",
                }

            )

            peesimg = "G:\\Drives compartilhados\\Gente e Gestão FVT\\FVTech\\13. PEX\\peesimg.png"

            worksheet = workbook.add_worksheet('Consolidado')
            worksheet.insert_image("A4", peesimg, {'x_offset': 54, 'y_offset': 25})
            worksheet.hide_gridlines(2) 
            worksheet.set_column("A:A", 8) 
            worksheet.set_column("B:F", 20) 
            worksheet.set_column('H:XFD', None, None, {'hidden': True})
            worksheet.set_default_row(hide_unused_rows=True) 
            worksheet.set_row(0, None, None, {'hidden': True})
            worksheet.set_row(0, 10)         
            worksheet.set_row(2, 10)
            worksheet.set_row(4, 10)
            worksheet.set_row(5, 20)
            worksheet.set_row(6, 20)        
            worksheet.set_row(7, 20)
            worksheet.set_row(8, 20)
            worksheet.set_row(9, 20)
            worksheet.set_row(10, 20)
            worksheet.set_row(11, 20)
            worksheet.set_row(12, 20)
            worksheet.set_row(13, 20)
            worksheet.set_zoom(80)
            worksheet.write('A1', ' ', campo3)
            worksheet.write('A3', ' ', campo3)
            worksheet.write('A16', ' ', campo3)
            worksheet.merge_range('C2:F2', f'FRANQUIA {franchise}'.replace("POLO ",""), campo2)        
            worksheet.merge_range('C4:F4', 'ACOMPANHAMENTO Q3', campo3)
            worksheet.merge_range('C6:D6', 'Lucro Bruto', campo4)
            worksheet.merge_range('E6:F6', 'Base Ativa', campo4)
            worksheet.write('C7', 'Crescimento Meta', campo5)
            worksheet.write('D7', 'Crescimento Real', campo5)
            worksheet.write('E7', 'Crescimento Meta', campo5)        
            worksheet.write('F7', 'Crescimento Real', campo5)


            worksheet.write('B8', 'Julho', campo6)
            worksheet.write('B9', 'Agosto', campo6)
            worksheet.write('B10', 'Setembro', campo6)
            worksheet.write('C8', julho_meta_lb, porcentagem)
            worksheet.write('C9', agosto_meta_lb, porcentagem)
            worksheet.write('C10', setembro_meta_lb, porcentagem)
            worksheet.write('D8', julho_real_lb, porcentagem)
            worksheet.write('D9', agosto_real_lb, porcentagem)
            worksheet.write('D10', setembro_real_lb, porcentagem)        
            worksheet.write('E8', julho_meta_ba, porcentagem)
            worksheet.write('E9', agosto_meta_ba, porcentagem)
            worksheet.write('E10', setembro_meta_ba, porcentagem)
            worksheet.write('F8', julho_real_ba, porcentagem)
            worksheet.write('F9', agosto_real_ba, porcentagem)
            worksheet.write('F10', setembro_real_ba, porcentagem)
            worksheet.write('B11', 'Trimestre', campo7)
            worksheet.write('C11', media_lb_meta, porcentagem2)
            worksheet.write('D11', media_lb_real, porcentagem2)
            worksheet.write('E11', media_ba_meta, porcentagem2)
            worksheet.write('F11', media_ba_real, porcentagem2)
            worksheet.merge_range('C12:E12', 'Meta Ponderada', campo6)
            worksheet.merge_range('C13:E13', 'Realizado Ponderado', campo6)
            worksheet.write('F12', meta_ponderado, porcentagem)
            worksheet.write('F13', real_ponderado, porcentagem)
            worksheet.merge_range('C14:E15', 'PONTUAÇÃO PEX22', campo8)
            worksheet.merge_range('F14:F15', f'{final_ponderado*10000} pb', porcentagem3)


            worksheet2 = workbook.add_worksheet('Apoio')
            worksheet2.set_zoom(80)
            worksheet2.set_column("A:A", 50)
            worksheet2.hide_gridlines(2)
            worksheet2.set_row(0, 25)
            worksheet2.set_row(14, 25)     
            worksheet2.merge_range('A1:N1', 'Etapa Eliminatória: Métodos e Processos', campo4)
            worksheet2.merge_range('A2:A6', 'Estrutura Física', campo9)
            worksheet2.merge_range('A7:A9', 'Rotina de planejamento semanal de vendas', campo9)
            worksheet2.merge_range('A10:A11', 'Rota coach', campo9)
            worksheet2.write('A12', 'Qualidade de Logística', campo9)
            worksheet2.merge_range('B2:N2', '(a) estrutura física íntegra, com boa aparência e com seus componentes em perfeito estado de funcionamento;', campo10)
            worksheet2.merge_range('B3:N3', '(b) exposição de material de divulgação da Cultura Stone;', campo10)
            worksheet2.merge_range('B4:N4', '(c) copa para atendimento ao pessoal do polo;', campo10)      
            worksheet2.merge_range('B5:N5', '(d) TV e computadores funcionando perfeitamente;', campo10)       
            worksheet2.merge_range('B6:N6', '(e) material de escritório completo.', campo10)        
            worksheet2.merge_range('B7:N7', '(a) trazer os resultados da franquia na semana;', campo10)        
            worksheet2.merge_range('B8:N8', '(b) consolidar os aprendizados da semana anterior, direcionar o time para a semana seguinte;', campo10)  
            worksheet2.merge_range('B9:N9', '(c) trazer plano de ação recalibrado para atingir a meta do mês.', campo10)       
            worksheet2.merge_range('B10:N10', '(a) planejar e executar, ao menos, uma rota coach por agente no mês;', campo10)        
            worksheet2.merge_range('B11:N11', '(b) forneceu feedback estruturado ao agente.', campo10)        
            worksheet2.merge_range('B12:N12', '(a) garantir que, pelo menos 93% todas as OS’s foram realizadas dentro do prazo máximo e atenderam as necessidades do cliente.', campo10)                   
            worksheet2.write('A13', ' ', campo9)
            worksheet2.write('A14', ' ', campo9)
            worksheet2.merge_range('A15:N15', 'Etapa Classificatória: Metas', campo4)        
            worksheet2.merge_range('A16:A17', 'Percentual de crescimento de base ativa', campo9)        
            worksheet2.merge_range('A18:A20', 'Percentual de crescimento de lucro bruto do rebate', campo9)          
            worksheet2.merge_range('B16:N17', 'Pontuação é o percentual de atingimento da meta, que será proposta pela Stone e validada pelo franqueado. Base de comparação é o mesmo trimestre do ano anterior. Meta = (Peso Lucro Bruto * Meta Lucro Bruto) + (Peso Base Ativa * Meta Base Ativa). Peso: 20%.', campo11)
            worksheet2.merge_range('B18:N20', 'Pontuação é o percentual de atingimento da meta, que será proposta pela Stone e validada pelo franqueado. Base de comparação é o mesmo trimestre do ano anterior. (Peso Lucro Bruto * Realizado Lucro Bruto) + (Peso Base Ativa * Realizado Basee Ativa) Peso: 80%.', campo11)




    
    
    
