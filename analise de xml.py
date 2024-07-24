import xml.etree.ElementTree as ET
import pandas as pd
import os
import tkinter as tk
import sys

def procurar_arquivo_xml_e_criar_planilha():
    numeros = entry.get().split()  # Divide a entrada em uma lista de números
    
    # Verifica se pelo menos um número foi fornecido
    if numeros:
        pasta = "C:/"  # Substitua pelo caminho da sua pasta
        arquivos_na_pasta = os.listdir(pasta)
        
        # Procura o arquivo cujo nome contenha pelo menos um dos números informados
        encontrados = []
        for arquivo in arquivos_na_pasta:
            if any(numero in arquivo for numero in numeros):
                encontrados.append(os.path.join(pasta, arquivo))
        
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        item = 0
        
        colunas = ["número_item", "código_produto", "nome_produto", "valor_produto", "numero_lote", "quant_lote", "data_fab", "data_validade"]
        valores = [] 
        # Itera pelos arquivos XML encontrados
        for arquivo in encontrados:
            tree = ET.parse(arquivo)
            root = tree.getroot()
                
            for det in root.findall('.//nfe:det', ns):
                # Extraia informações do XML (substitua pelas tags XML desejadas)
                item +=1
                cProd = det.find('.//nfe:cProd', ns).text
                xProd = det.find('.//nfe:xProd', ns).text
                vProd = det.find('.//nfe:vProd', ns).text
                nLote = det.find('.//nfe:nLote', ns).text
                qLote = det.find('.//nfe:qLote', ns).text
                dFab = det.find('.//nfe:dFab', ns).text
                dVal = det.find('.//nfe:dVal', ns).text
                valores.append([item, cProd, xProd, vProd, nLote, qLote, dFab, dVal])
                
        # Adicione as informações à planilha
        tabela = pd.DataFrame(columns=colunas, data=valores)
            
        # Salvar o DataFrame em um arquivo Excel
        excel_file = "analise_xml.xlsx"
        tabela.to_excel(excel_file, index=False)
            
        # Abre o arquivo e fecha o sistema
        os.system(f'start excel "{excel_file}"')
        sys.exit()
            
    else:
        resultado.config(text='Nenhum arquivo XML foi encontrado com os números no nome.')

# Cria a janela principal
root = tk.Tk()
root.title("Procurar Arquivo XML pelo Nome e Criar Planilha Excel")

# Cria os widgets
label = tk.Label(root, text="Digite os números separados por espaço:")
entry = tk.Entry(root)
procurar_button = tk.Button(root, text="Procurar e Criar Planilha", command=procurar_arquivo_xml_e_criar_planilha)
resultado = tk.Label(root, text="Resultado será exibido aqui")

# Organiza os widgets na janela
label.pack()
entry.pack()
procurar_button.pack()
resultado.pack()

# Inicia o loop principal da GUI
root.mainloop()