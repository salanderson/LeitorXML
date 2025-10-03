import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import os


def extrair_dados_xml(caminho_arquivo):
    
    try:
        
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        arquivos_xml = [f for f in os.listdir('.')if f.lower().endswith('.xml')]

        for caminho_arquivo in arquivos_xml:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()

        emit = root.find('.//nfe:emit', ns)
        if emit is None:
            texto_resultado.insert(tk.END, f'{caminho_arquivo}: ignorado (Sem <cnpj>, <nome> e <fantasia>)')
        
        ide = root.find('.//nfe:ide', ns)
        if ide is None:
            texto_resultado.insert(tk.END, f'{caminho_arquivo}: ignorado  (sem <mod> ou sem <serie>)')
            
        
        mod = root.find('.//nfe:mod', ns)
        serie = root.find('.//nfe:serie',ns)

        if mod is None or serie is None:
            texto_resultado.insert(tk.END, f'{caminho_arquivo}: ignorado  (sem <mod> ou sem <serie>)')
            
        dados = {

            'mod' : ide.findtext('nfe:mod', default='Nao encontrado', namespaces=ns),
            'serie' : ide.findtext('nfe:serie', default='Nao encontrado', namespaces=ns),
            'nNF' : ide.findtext('nfe:nNF', default='Nao encontrado', namespaces=ns),
            'dhEmi' : ide.findtext('nfe:dhEmi', default='Nao encontrado', namespaces=ns),
            'dhSaiEnt': ide.findtext('nfe:dhSaiEnt', default='Nao encontrado', namespaces=ns),

            'cnpj' : emit.findtext('nfe:CNPJ', default='Nao encontrado', namespaces=ns),
            'nome' : emit.findtext('nfe:xNome', default='Nao encontrado', namespaces=ns),
            'fant' : emit.findtext('nfe:xFant', default='Nao encontrado', namespaces=ns),
            
        }

        return dados

    except Exception as e:
        print(f'Erro ao processar {caminho_arquivo}: {e}')
        return None
    
dados_xmls = []

def importar_xmls():
    global dados_xmls
    dados_xmls = []

    pasta = filedialog.askdirectory(title='Selecione a pasta com XMLs')
    if not pasta:
        return
    
    arquivos = [f for f in os.listdir(pasta) if f.endswith('.xml')]
    texto_resultado.delete('1.0', tk.END)

    if not arquivos:
        texto_resultado.insert(tk.END, 'Nenhum arquivo XML encontrado. \n')
        return
    
    total = len(arquivos)
    barra_progresso['maximum'] = total
    
    for i, arquivo in enumerate(arquivos, start=1):
        caminho = os.path.join(pasta, arquivo)
        dados = extrair_dados_xml(caminho)
        
        texto_resultado.insert(tk.END, f'Arquivo: {arquivo}\n')
        if dados is not None:
            dados_xmls.append(dados)
            for chave, valor in dados.items():
                texto_resultado.insert(tk.END, f'{chave}: {valor}\n')
        else:
            texto_resultado.insert(tk.END, 'Erro ao extrair dados do XML.\n')    
        texto_resultado.insert(tk.END, "-" * 60 + "\n")
    
        barra_progresso['value'] = i
        janela.update_idletasks()
        barra_progresso['value'] = 0
    
def listar_dados_simplificados():
    texto_resultado.delete('1.0', tk.END)
    
    if not dados_xmls:
        texto_resultado.insert(tk.END, 'Nenhum dados disponivel, Importe os XMLs primeiro. \n')    
        return
    
    dados_emitente = dados_xmls[0]
    texto_resultado.insert(tk.END, f"CNPJ: {dados_emitente['cnpj']} | Nome: {dados_emitente['nome']} | Fantasia: {dados_emitente['fant']}\n")
    texto_resultado.insert(tk.END, "-" * 80 + "\n")
    
    texto_resultado.insert(tk.END, 'Lista simplificada de Notas Fiscais: \n')
    for i, dados in enumerate(dados_xmls, start=1):
        modelo = dados.get('mod', 'Desconhecido')
        texto_resultado.insert(tk.END, f"{i}. Numero: {dados['nNF']} | Serie: {dados['serie']} | Modelo: {modelo} | Emissao: {dados['dhEmi']}\n")

def obter_ultima_nota():
    
    if not dados_xmls:
        return None
    try:
        notas_validas = [d for d in dados_xmls if d['nNF'].isdigit()]
        if not notas_validas:
            return None
        return max(dados_xmls, key=lambda x: int(x['nNF']))
    except Exception as e:
        print(f'Erro ao buscar ultima nota: {e}')
        return None

def mostrar_ultima_nota():
    texto_resultado.delete('1.0', tk.END)
    
    if not dados_xmls:
        texto_resultado.insert(tk.END, 'Nenhum dados disponivel. Importe os XMLs primeiro. \n')
        return
    
    modelos = {}

    for dados in dados_xmls:
        try:
            numero = int(dados['nNF'])
            modelo = dados.get('mod', 'Desconhecido')

            if modelo not in modelos:
                modelos[modelo] = []

            modelos[modelo].append((numero, dados))

        except ValueError:
            continue
    
    for modelo, lista in modelos.items():
        if not lista:
            continue

        ultima_nota = max(lista, key=lambda x: x[0])
        dados_ultima = ultima_nota[1]

        texto_resultado.insert(tk.END, f"\nUltima nota emitida - Modelo {modelo}:\n")
        texto_resultado.insert(tk.END, f"Numero: {dados_ultima['nNF']}\n")
        texto_resultado.insert(tk.END, f"Serie: {dados_ultima.get('serie', 'Nao Informado')}\n")
        texto_resultado.insert(tk.END, f"Data: {dados_ultima.get('dhEmi', 'Nao Informado')}\n")
        texto_resultado.insert(tk.END, f"CNPJ: {dados_ultima.get('cnpj', 'Nao Informado')}\n")
        texto_resultado.insert(tk.END, f"Nome: {dados_ultima.get('nome', 'Nao Informado')}\n")
        texto_resultado.insert(tk.END, f"Fantasia: {dados_ultima.get('fant', 'Nao Informado')}\n")
    
def verificar_sequencia_numerica():
    texto_resultado.delete('1.0', tk.END)
    if not dados_xmls:
        texto_resultado.insert(tk.END, 'Nenhum dados disponivel. Importe os XMLs primeiro. \n ')
        return
    
    agrupados = {}

    for dados in dados_xmls:
        try:
            numero = int(dados['nNF'])
            serie = dados.get('serie', 'Sem serie')
            modelo = dados.get('mod', 'Desconhecido')
            chave = (modelo, serie)

            if chave not in agrupados:
                agrupados[chave] = {
                    'numeros' : [],
                    'cnpj' : dados.get('cnpj', 'Nao encontrado'),
                    'nome' : dados.get('nome', 'Nao encontrado'),
                    'fant' : dados.get('fant', 'Nao encontrado')
                }

            agrupados[chave]['numeros'].append(numero)

        except ValueError:
            texto_resultado.insert(tk.END, f"Numero invalido encontrado: {dados['nNF']}\n")

    if not agrupados:
        texto_resultado.insert(tk.END, 'Nenhum numero de nota valido encontrado.\n')
        return
    
    for (modelo, serie), info in agrupados.items():
        numeros_nf = info['numeros']
        texto_resultado.insert(tk.END, f"\nModelo: {modelo} | Serie: {serie}\n")
        texto_resultado.insert(tk.END, f"CNPJ: {info['cnpj']} | Nome: {info['nome']} | Fantasia: {info['fant']}\n")

        if not numeros_nf:
            texto_resultado.insert(tk.END, 'Nenhum numero encontrado para este grupo. \n')
            continue


        numeros_nf.sort()
        faltantes = [i for i in range(numeros_nf[0], numeros_nf[-1]) if i not in numeros_nf]

        if faltantes:
            texto_resultado.insert(tk.END, 'Notas faltando na sequencia: \n')
            for numero in faltantes:
                texto_resultado.insert(tk.END, f'{numero}\n')
        
        else:
            texto_resultado.insert(tk.END, 'Sequencia completa. Nenhum numero faltando. \n')
    


janela = tk.Tk()
janela.title('Leitor de XML')
janela.geometry("700x500")

frame_principal = tk.Frame(janela)
frame_principal.pack(fill='both', expand=True)

frame_botoes = tk.Frame(janela)
frame_botoes.pack(side='top', pady=10)

frame_texto = tk.Frame(janela)
frame_texto.pack(fill='both', expand=True)

frame_rodape = tk.Frame(janela)
frame_rodape.pack(side='bottom', fill='x')

botao = tk.Button(frame_botoes, text='Importar XML', command= importar_xmls)
botao.pack(side=tk.LEFT, padx=5)

botao_listar = tk.Button(frame_botoes,text='Listar XML', command= listar_dados_simplificados)
botao_listar.pack(side=tk.LEFT, padx=5)

botao_verificar = tk.Button(frame_botoes, text='Notas Faltantes', command= verificar_sequencia_numerica)
botao_verificar.pack(side=tk.LEFT, padx=5)

botao_ultima_nota = tk.Button(frame_botoes,text='Ultima Nota Emitida', command= mostrar_ultima_nota)
botao_ultima_nota.pack(side=tk.LEFT, padx=5)

texto_resultado = tk.Text(janela, wrap='word', height=20, width=90)
scrollbar = tk.Scrollbar(janela, command=texto_resultado.yview)
texto_resultado.configure(yscrollcommand=scrollbar.set)

texto_resultado = ScrolledText(janela, width=100, height=30)
texto_resultado.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

texto_resultado.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

barra_progresso = ttk.Progressbar(frame_rodape, orient='horizontal', length=650, mode='determinate')
barra_progresso.pack(side='top', pady=10)

label_desenvolvedor = tk.Label(frame_rodape, text='Desenvolvido por Anderson Saldanha', font=('Arial', 10), anchor='center')
label_desenvolvedor.pack(side='bottom', fill='x', pady=(0, 5))



janela.mainloop()