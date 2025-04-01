import re
import openpyxl
import datetime
import pytz

def ler_arquivo(nome_arquivo):
    try:
        with open(nome_arquivo, 'r', encoding='utf-8') as file:
            conteudo = file.read()
        return conteudo
    except FileNotFoundError:
        return f"O arquivo {nome_arquivo} não foi encontrado."
    except Exception as e:
        return f"Ocorreu um erro: {e}"


def extrair_dados(conteudo):
    linhas = conteudo.splitlines()
    
    dados_extraidos = []
    
    for i in range(0, len(linhas), 3):
        if i + 2 < len(linhas):
            nome_codigo = linhas[i].strip()
            quantidade_unidade_valor_unitario = linhas[i+1].strip()
            valor_total = linhas[i+2].strip()
            
            match_nome_codigo = re.search(r"(?P<nome>.+?)\s*\(Código:\s*(?P<código>\d+)\s*\)", nome_codigo)
            if match_nome_codigo:
                nome = match_nome_codigo.group('nome').strip()
                codigo = match_nome_codigo.group('código')
            else:
                continue
            
            match_quantidade = re.search(
                r"Qtde\.\s*:\s*(?P<quantidade>\d+[\.,]?\d*)\s+UN:\s*(?P<unidade>\S+)\s+Vl\.\s*Unit\.\s*:\s*(?P<valor_unitario>\d+[\.,]?\d{0,2})",
                quantidade_unidade_valor_unitario
            )
            if match_quantidade:
                quantidade = match_quantidade.group('quantidade')
                unidade = match_quantidade.group('unidade')
                valor_unitario = match_quantidade.group('valor_unitario')
                valor_unitario = float(valor_unitario.replace(',', '.'))
            else:
                continue

            match_valor_total = re.search(r"(?P<valor_total>\d+[\.,]?\d{0,2})", valor_total)
            if match_valor_total:
                valor_total = match_valor_total.group('valor_total')
                valor_total = float(valor_total.replace(',', '.'))
            else:
                continue
            
            dados_extraidos.append({
                'nome': nome,
                'código': codigo,
                'quantidade': quantidade,
                'unidade': unidade,
                'valor_unitario': valor_unitario,
                'valor_total': valor_total
            })
    
    return dados_extraidos

def salvar_em_xlsx(dados, nome_arquivo):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'Dados Extraídos'
    
    cabecalho = ['Nome', 'Código', 'Quantidade', 'Unidade', 'Valor Unitário', 'Valor Total']
    sheet.append(cabecalho)

    for item in dados:
        sheet.append([
            item['nome'],
            item['código'],
            item['quantidade'],
            item['unidade'],
            item['valor_unitario'],
            item['valor_total']
        ])
    
    wb.save(nome_arquivo)
    print(f"Arquivo '{nome_arquivo}' salvo com sucesso!")

if __name__ == "__main__":

    nome_arquivo = 'consulta.txt'
    conteudo = ler_arquivo(nome_arquivo)
    
    itens = extrair_dados(conteudo)

    fuso_brasilia = pytz.timezone('America/Sao_Paulo')
    agora = datetime.datetime.now(fuso_brasilia)
    timestamp_formatado = agora.strftime('%d-%m-%Y_%H-%M-%S')
    nome_arquivo_xlsx = f'dados_extraidos_{timestamp_formatado}.xlsx'
    salvar_em_xlsx(itens, nome_arquivo_xlsx)