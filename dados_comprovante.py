import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tqdm import tqdm
import threading
import subprocess

# Configuração do caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Função para converter valores para inteiro
def convert_to_int(value):
    try:
        value = str(value)
        value = value.replace('$', '').replace(' ', '').replace('.', '')
        return int(value)
    except (ValueError, AttributeError):
        return None

# Função para extrair texto de PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text += page.get_text()
    except Exception as e:
        text = f"Erro ao ler PDF: {e}"
    return text

# Função para extrair texto de imagem
def extract_text_from_image(image_path):
    text = ""
    try:
        text = pytesseract.image_to_string(Image.open(image_path))
    except Exception as e:
        text = f"Erro ao ler imagem: {e}"
    return text

# Função para remover caracteres ilegais
def remove_illegal_characters(text):
    illegal_characters_re = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]|\x7F')
    return illegal_characters_re.sub("", text)

# Função para tratar comprovantes Wise
def tratar_info_wise(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    # Expressões regulares para extrair as informações
    data_re = re.search(r'\b(?:\d{1,2}/\d{1,2}/\d{2,4}|\w+\s+\d{1,2},\s+\d{4})\b', info_comp)
    if data_re:
        data = data_re.group()
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, dayfirst=True).strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'(Total para Laise Rocha De Mesquita|Total to Laise Rocha De Mesquita)\s+([\d,\.]+)', info_comp)
    if valor_re:
        valor = valor_re.group(2).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'Transfer(ência)?\s*#(\d+)', info_comp)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(2)

    pagador_re = re.search(r'Amount paid by\s+([^\n]+)', info_comp)
    if not pagador_re:
        pagador_re = re.search(r'Valor pago por\s+([^\n]+)', info_comp)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Total to\s+([^\n]+)', info_comp)
    if not recebedor_re:
        recebedor_re = re.search(r'Total para\s+([^\n]+)', info_comp)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Itaú
def tratar_info_itau(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'(?:realizada\s*em|Realizada\s*em):?\s*(\d{2}/\d{2}/\d{4})\s*às\s*(\d{2}:\d{2}:\d{2})', info_comp, re.IGNORECASE)
    if data_re:
        data = f"{data_re.group(1)} {data_re.group(2)}"
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d/%m/%Y %H:%M:%S').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'rs\s*([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID da transa[cç][aã]o\s*(\w+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'de\s+([^\n]+)', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'para\s+([^\n]+)', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Santander
def tratar_info_santander(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'Data e hora da transa[cç][aã]o\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}:\d{2}:\d{2})', info_comp, re.IGNORECASE)
    if data_re:
        data = f"{data_re.group(1)} {data_re.group(2)}"
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d/%m/%Y %H:%M:%S').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'Valor pago\s*R\$\s*([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID/Transa[cç][aã]o\s+(\w+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Institui[cç][aã]o iniciadora do pagamento\s+([^\n]+)', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Para\s+([^\n]+)\nCNPJ', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Banco do Brasil
def tratar_info_bb(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'\b\d{2}/\d{2}/\d{4}\b', info_comp)
    if data_re:
        data = data_re.group()
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, dayfirst=True).strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'VALOR:\s*R\$\s*([\d,\.]+)', info_comp)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID:\s+(\w+)', info_comp)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'CLIENTE:\s+([^\n]+)', info_comp)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'PAGO PARA:\s+([^\n]+)', info_comp)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Sicredi
def tratar_info_sicredi(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'Realizado em:\s*(\d{2}/\d{2}/\d{4})', info_comp)
    if data_re:
        data = data_re.group(1)
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, dayfirst=True).strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'Valor:\s*R\$\s*([\d,\.]+)', info_comp)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID da transação:\s+(\w+)', info_comp)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Nome do pagador:\s+([^\n]+)', info_comp)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Nome do destinatário:\s+([^\n]+)', info_comp)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Nubank
def tratar_info_nubank(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'(\d{2} \w+ \d{4} - \d{2}:\d{2}:\d{2})', info_comp)
    if data_re:
        data = data_re.group(1)
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d %b %Y - %H:%M:%S').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'Valor R\$\s*([\d,\.]+)', info_comp)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID da transa[cç][aã]o:\s+(\w+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Nome ([^\n]+)\s+Institui[cç][aã]o NU PAGAMENTOS - IP', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Nome ([^\n]+)\s+CNPJ', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Banco Inter
def tratar_info_inter(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'Data do pagamento\s+\w+,\s+(\d{2}/\d{2}/\d{4})', info_comp)
    if data_re:
        data = data_re.group(1)
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d/%m/%Y').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'RS\s*([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID da transa[cç][aã]o\s+(\w+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Nome ([^\n]+)\s+CPF/CNPJ', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Nome ([^\n]+)\s+CPF/CNPJ', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Mercado Pago
def tratar_info_mercado_pago(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'\w+-feira, (\d{2} de \w+ de \d{4}) às (\d{2}:\d{2}:\d{2})', info_comp, re.IGNORECASE)
    if data_re:
        data = f"{data_re.group(1)} {data_re.group(2)}"
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d de %B de %Y %H:%M:%S').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'R\$\s*([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'E[^\s]+', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(0)

    pagador_re = re.search(r'e\s+([^\n]+)\n\*\*.\d{3}\.\d{3}-\*\*\nMercado Pago', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'e\s+([^\n]+)\n41[^\n]+', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes C6 Bank
def tratar_info_c6(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'(?:Data & hora da transa[cç][aã]o|Data e hora da transa[cç][aã]o)\s+(\w+-feira, \d{2} de \w+ de \d{4}, \d{2}:\d{2})', info_comp, re.IGNORECASE)
    if data_re:
        data = data_re.group(1)
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%A, %d de %B de %Y, %H:%M').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'R\$ ([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID da transa[cç][aã]o\s+([^\s]+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Contadeorigem\s+([^\n]+)', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'We Love Chile', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = "We Love Chile"

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Bradesco
def tratar_info_bradesco(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'Data e Hora:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}:\d{2}:\d{2})', info_comp)
    if data_re:
        data = f"{data_re.group(1)} {data_re.group(2)}"
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d/%m/%Y %H:%M:%S').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'Valor:\s*R\$\s*([\d,\.]+)', info_comp)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'Número de Controle:\s+(\w+)', info_comp)
    if not transacao_id_re:
        transacao_id_re = re.search(r'ID da transação\s+(\w+)', info_comp)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Nome:\s+([^\n]+)\s+CPF:', info_comp)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Nome:\s+([^\n]+)\s+CNPJ:', info_comp)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Neon
def tratar_info_neon(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'Ocorreu em (\d{2} de \w+ de \d{4}) às (\d{2}:\d{2})', info_comp, re.IGNORECASE)
    if data_re:
        data = f"{data_re.group(1)} {data_re.group(2)}"
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d de %B de %Y %H:%M').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'R\$ ([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID de transa[cç][aã]o\s+([^\s]+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'Nome\s+([^\n]+)\s+CPF / CNPJ', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'Nome\s+([^\n]+)\s+CPF / CNPJ.*\n.*\nInstitui[cç][aã]o ITAU UNIBANCO S.A.', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"

# Função para tratar comprovantes Pix por chave
def tratar_info_pix(info_comp):
    data = valor = transacao_id = pagador = recebedor = ""

    data_re = re.search(r'realizado em\s*(\d{2}/\d{2}/\d{4})', info_comp, re.IGNORECASE)
    if data_re:
        data = data_re.group(1)
        try:
            # Converter para formato DD/MM/AAAA
            data = pd.to_datetime(data, format='%d/%m/%Y').strftime('%d/%m/%Y')
        except ValueError:
            pass

    valor_re = re.search(r'R\$?\s*([\d,\.]+)', info_comp, re.IGNORECASE)
    if valor_re:
        valor = valor_re.group(1).replace('.', '').replace(',', '.')

    transacao_id_re = re.search(r'ID da transa[cç][aã]o\s*([^\s]+)', info_comp, re.IGNORECASE)
    if transacao_id_re:
        transacao_id = transacao_id_re.group(1)

    pagador_re = re.search(r'dados da conta debitada\nnome\s*([^\n]+)', info_comp, re.IGNORECASE)
    if pagador_re:
        pagador = pagador_re.group(1).strip()

    recebedor_re = re.search(r'nome do favorecido\s*([^\n]+)', info_comp, re.IGNORECASE)
    if recebedor_re:
        recebedor = recebedor_re.group(1).strip()

    return f"{data};{valor};{transacao_id};{pagador};{recebedor}"


# Função para identificar o tipo de comprovante e tratar as informações
def tratar_info(info_comp):
    if any(keyword in info_comp for keyword in ['Wise Payments Limited', 'A Wise Payments Limited', 'Wise Brasil Corretora de Câmbio Ltda']):
        return tratar_info_wise(info_comp)
    elif 'Comprovante de transferência Pix por chave' in info_comp:
        return tratar_info_pix(info_comp)
    elif 'transferência realizada' in info_comp.lower():
        return tratar_info_itau(info_comp)
    elif 'Comprovante do Pix' in info_comp or 'Pronto! Seu pagamento foi realizado' in info_comp:
        return tratar_info_santander(info_comp)
    elif 'Comprovante Pix' in info_comp and 'SISBB' in info_comp:
        return tratar_info_bb(info_comp)
    elif 'Comprovante Pix' in info_comp:
        return tratar_info_bradesco(info_comp)
    elif 'Comprovante de Pagamento PIX' in info_comp:
        return tratar_info_sicredi(info_comp)
    elif info_comp.strip().lower().startswith('nu'):
        return tratar_info_nubank(info_comp)
    elif info_comp.strip().lower().startswith('sinter'):
        return tratar_info_inter(info_comp)
    elif 'mercado' in info_comp.lower():
        return tratar_info_mercado_pago(info_comp)
    elif info_comp.strip().lower().startswith('cobank'):
        return tratar_info_c6(info_comp)
    elif info_comp.strip().lower().startswith('ncon'):
        return tratar_info_neon(info_comp)
    else:
        return "Tipo de comprovante não identificado"

def process_files(html_path, comprovantes_dir, output_dir):
    # Leitura dos dados HTML dos comprovantes
    pagos_recebidos_html = pd.read_html(html_path)
    DadosBrutos = pagos_recebidos_html[3]

    # Limpando e preparando os dados
    df = DadosBrutos.copy()
    df.drop([0, 1, 2, 3], axis=0, inplace=True)  # Removendo linhas desnecessárias
    df.drop([1, 3, 6], axis=1, inplace=True)     # Removendo colunas desnecessárias
    df.columns = df.iloc[0]  # Definindo a primeira linha como cabeçalho
    df = df[1:]  # Removendo a linha de cabeçalho duplicada
    df["Data"] = df["Data"].str.split(' ').str[1]
    df['Monto'] = df['Monto'].apply(convert_to_int)  # Convertendo valores

    # Função fictícia para simular o carregamento dos dados de login
    df_logins = pd.DataFrame({'LOGIN': ['Autor1', 'Autor2'], 'SETOR': ['Setor1', 'Setor2']})
    df = df.merge(df_logins, left_on='Autor', right_on='LOGIN', how='left')
    df.drop('LOGIN', axis=1, inplace=True)  # Removendo coluna de login redundante
    new_order = ['ID', 'Data', 'Cliente', 'Autor', 'SETOR', 'Monto', 'Forma de Pago', 'Obs']
    df = df[new_order]

    # Garantindo que a coluna 'ID' seja string e substituindo valores nulos
    df['ID'] = df['ID'].astype(str).fillna('')

    # Removendo linhas cujo ID começa com "REE-"
    df = df[~df['ID'].str.startswith("REE-")]

    # Lista para armazenar as informações dos comprovantes
    info_comp_list = []
    info_tratada_list = []

    # Formas de pagamento a serem ignoradas
    formas_de_pago_ignoradas = [
        'Tarjeta de Credito',
        'Mercado Pago',
        'Efectivo oficina Santiago',
        'Efectivo oficina Atacama'
    ]

    # Iterando sobre os IDs para processar os comprovantes
    total_files = df.shape[0]
    progress["maximum"] = total_files
    for index, row in enumerate(tqdm(df.iterrows(), total=total_files)):
        row = row[1]  # Acessa a série que representa a linha
        comp_id = row['ID']
        
        # Verificar se a forma de pagamento deve ser ignorada
        if row['Forma de Pago'] in formas_de_pago_ignoradas:
            info_comp_list.append("Forma de Pago ignorada")
            info_tratada_list.append("Forma de Pago ignorada")
            continue
        
        comp_path_pdf = os.path.join(comprovantes_dir, f'{comp_id}.pdf')
        comp_path_jpg = os.path.join(comprovantes_dir, f'{comp_id}.jpg')
        comp_path_png = os.path.join(comprovantes_dir, f'{comp_id}.png')
        comp_path_jpeg = os.path.join(comprovantes_dir, f'{comp_id}.jpeg')
        
        if os.path.exists(comp_path_pdf):
            text = extract_text_from_pdf(comp_path_pdf)
        elif os.path.exists(comp_path_jpg):
            text = extract_text_from_image(comp_path_jpg)
        elif os.path.exists(comp_path_png):
            text = extract_text_from_image(comp_path_png)
        elif os.path.exists(comp_path_jpeg):
            text = extract_text_from_image(comp_path_jpeg)
        else:
            text = f"Documento não encontrado para ID: {comp_id}"
        
        info_comp_list.append(remove_illegal_characters(text))
        info_tratada_list.append(tratar_info(text))
        progress["value"] = index + 1
        root.update_idletasks()

    # Adicionando as novas colunas ao dataframe
    df['INFO_COMP'] = info_comp_list
    df['info_tratada'] = info_tratada_list

    # Salvando o DataFrame em um arquivo Excel
    output_path = os.path.join(output_dir, 'info_comprovantes.xlsx')
    df.to_excel(output_path, index=False)
    messagebox.showinfo("Concluído", f"O arquivo Excel foi criado com sucesso em {output_path}!")
    subprocess.Popen(['start', output_path], shell=True)

def select_html_file():
    file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html"), ("All files", "*.*")])
    if file_path:
        html_entry.delete(0, tk.END)
        html_entry.insert(0, file_path)

def select_comprovantes_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        comprovantes_entry.delete(0, tk.END)
        comprovantes_entry.insert(0, folder_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)

def start_processing():
    html_path = html_entry.get()
    comprovantes_dir = comprovantes_entry.get()
    output_dir = output_entry.get()
    if not html_path or not comprovantes_dir or not output_dir:
        messagebox.showwarning("Aviso", "Por favor, selecione o arquivo HTML, a pasta dos comprovantes e a pasta de destino.")
        return
    
    threading.Thread(target=process_files, args=(html_path, comprovantes_dir, output_dir)).start()

# Criando a interface gráfica
root = tk.Tk()
root.title("Processador de Comprovantes")

tk.Label(root, text="Arquivo HTML:").grid(row=0, column=0, padx=10, pady=10)
html_entry = tk.Entry(root, width=50)
html_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_html_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Pasta dos Comprovantes:").grid(row=1, column=0, padx=10, pady=10)
comprovantes_entry = tk.Entry(root, width=50)
comprovantes_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_comprovantes_folder).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Pasta de Destino:").grid(row=2, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=2, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_output_folder).grid(row=2, column=2, padx=10, pady=10)

tk.Button(root, text="Iniciar Processamento", command=start_processing).grid(row=3, column=0, columnspan=3, padx=10, pady=20)

# Adicionando a barra de progresso
progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress.grid(row=4, column=0, columnspan=3, padx=10, pady=20)

root.mainloop()
