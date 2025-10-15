# Importação de bibliotecas necessárias
import xml.etree.ElementTree as ET  # Manipulação de arquivos XML
import tkinter as tk  # Interface gráfica
from tkinter import filedialog  # Diálogo para seleção de arquivos
from tkinter import ttk  # Componentes visuais como barra de progresso
from tkinter.scrolledtext import ScrolledText  # Área de texto com rolagem
import os  # Operações com o sistema de arquivos




# Função para extrair dados de um arquivo XML
def extrair_dados_xml(caminho_arquivo):
    try:
        # Namespace utilizado nos arquivos XML da NFe
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()

        # Busca os dados do emitente
        emit = root.find('.//nfe:emit', ns)
        if emit is None:
            texto_resultado.insert(tk.END, f'{caminho_arquivo}: ignorado (Sem <cnpj>, <nome> e <fantasia>)')

        # Busca os dados da identificação da nota
        ide = root.find('.//nfe:ide', ns)
        if ide is None:
            texto_resultado.insert(tk.END, f'{caminho_arquivo}: ignorado  (sem <mod> ou sem <serie>)')

        # Verifica se os campos obrigatórios estão presentes
        mod = root.find('.//nfe:mod', ns)
        serie = root.find('.//nfe:serie', ns)
        if mod is None or serie is None:
            texto_resultado.insert(tk.END, f'{caminho_arquivo}: ignorado  (sem <mod> ou sem <serie>)')

        # Extrai os dados relevantes da nota
        dados = {
            'mod': ide.findtext('nfe:mod', default='Nao encontrado', namespaces=ns),
            'serie': ide.findtext('nfe:serie', default='Nao encontrado', namespaces=ns),
            'nNF': ide.findtext('nfe:nNF', default='Nao encontrado', namespaces=ns),
            'dhEmi': ide.findtext('nfe:dhEmi', default='Nao encontrado', namespaces=ns),
            'dhSaiEnt': ide.findtext('nfe:dhSaiEnt', default='Nao encontrado', namespaces=ns),
            'cnpj': emit.findtext('nfe:CNPJ', default='Nao encontrado', namespaces=ns),
            'nome': emit.findtext('nfe:xNome', default='Nao encontrado', namespaces=ns),
            'fant': emit.findtext('nfe:xFant', default='Nao encontrado', namespaces=ns),
        }

        return dados

    except Exception as e:
        print(f'Erro ao processar {caminho_arquivo}: {e}')
        return None

# Lista que armazenará os dados extraídos dos XMLs
dados_xmls = []

def selecionar_varias_pastas():
    pastas = []
    while True:
        pasta = filedialog.askdirectory(title='Selecionar uma pasta com XMLs (Cancelar para encerrar) ')
        if not pasta:
            break
        pastas.append(pasta)
    return pastas
# Função para importar os arquivos XML de uma pasta
def importar_xmls():
    global dados_xmls
    dados_xmls = []

    # Abre diálogo para seleção de pasta
    pastas = selecionar_varias_pastas()
    if not pastas:
        return
    
    texto_resultado.delete('1.0', tk.END)
    arquivos = []

    for pasta in pastas:
        arquivos += [os.path.join(pasta, f) for f in os.listdir(pasta) if f.lower().endswith('.xml')]

    if not arquivos:
        texto_resultado.insert(tk.END, 'Nenhum arquivo XML encontrado. \n')
        return

    total = len(arquivos)
    barra_progresso['maximum'] = total

    # Processa cada arquivo XML
    for i, arquivo in enumerate(arquivos, start=1):
        caminho = os.path.basename(arquivo)
        dados = extrair_dados_xml(arquivo)

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

# Função para listar os dados simplificados das notas fiscais
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

# Função para obter a última nota fiscal com número válido
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

# Função para exibir a última nota emitida por modelo
def mostrar_ultima_nota():
    texto_resultado.delete('1.0', tk.END)

    if not dados_xmls:
        texto_resultado.insert(tk.END, 'Nenhum dados disponivel. Importe os XMLs primeiro. \n')
        return

    grupos = {}

    # Agrupa notas por modelo
    for dados in dados_xmls:
        try:
            numero = int(dados['nNF'])
            modelo = dados.get('mod', 'Desconhecido')
            serie = dados.get('serie', 'Sem serie')
            chave = (modelo, serie)

            if chave not in grupos:
                grupos[chave] = []

            grupos[chave].append((numero, dados))

        except ValueError:
            continue

    # Exibe a nota com maior número por modelo
    for (modelo, serie), lista in grupos.items():
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

# Função para verificar se há falhas na sequência numérica das notas
def verificar_sequencia_numerica():
    texto_resultado.delete('1.0', tk.END)
    if not dados_xmls:
        texto_resultado.insert(tk.END, 'Nenhum dados disponivel. Importe os XMLs primeiro. \n ')
        return

    agrupados = {}

    # Agrupa notas por modelo e série
    for dados in dados_xmls:
        try:
            numero = int(dados['nNF'])
            serie = dados.get('serie', 'Sem serie')
            modelo = dados.get('mod', 'Desconhecido')
            chave = (modelo, serie)

            if chave not in agrupados:
                agrupados[chave] = {
                    'numeros': [],
                    'cnpj': dados.get('cnpj', 'Nao encontrado'),
                    'nome': dados.get('nome', 'Nao encontrado'),
                    'fant': dados.get('fant', 'Nao encontrado')
                }

            agrupados[chave]['numeros'].append(numero)

        except ValueError:
            texto_resultado.insert(tk.END, f"Numero invalido encontrado: {dados['nNF']}\n")

      # Verifica se houve agrupamento de notas por modelo e série
    if not agrupados:
        texto_resultado.insert(tk.END, 'Nenhum numero de nota valido encontrado.\n')
        return

    # Itera sobre cada grupo de notas fiscais agrupadas por modelo e série
    for (modelo, serie), info in agrupados.items():
        numeros_nf = info['numeros']  # Lista de números de nota fiscal para o grupo atual

        # Exibe informações do grupo atual (modelo, série, emitente)
        texto_resultado.insert(tk.END, f"\nModelo: {modelo} | Serie: {serie}\n")
        texto_resultado.insert(tk.END, f"CNPJ: {info['cnpj']} | Nome: {info['nome']} | Fantasia: {info['fant']}\n")

        # Caso não haja números de nota no grupo, exibe mensagem e continua para o próximo grupo
        if not numeros_nf:
            texto_resultado.insert(tk.END, 'Nenhum numero encontrado para este grupo. \n')
            continue

        # Ordena os números de nota fiscal para facilitar a verificação de sequência
        numeros_nf.sort()

        # Identifica os números faltantes na sequência entre o menor e o maior número
        faltantes = [i for i in range(numeros_nf[0], numeros_nf[-1]) if i not in numeros_nf]

        # Exibe os números faltantes, se houver
        if faltantes:
            texto_resultado.insert(tk.END, 'Notas faltando na sequencia: \n')
            for numero in faltantes:
                texto_resultado.insert(tk.END, f'{numero}\n')
        else:
            # Caso não haja faltantes, informa que a sequência está completa
            texto_resultado.insert(tk.END, 'Sequencia completa. Nenhum numero faltando. \n')



# Criação da janela principal da interface gráfica
janela = tk.Tk()
janela.title('Leitor de XML')  # Título da janela
janela.geometry("700x500")  # Tamanho da janela

# Frame principal que agrupa os demais componentes
frame_principal = tk.Frame(janela)
frame_principal.pack(fill='both', expand=True)

# Frame que contém os botões superiores
frame_botoes = tk.Frame(janela)
frame_botoes.pack(side='top', pady=10)

# Frame que contém a área de texto
frame_texto = tk.Frame(janela)
frame_texto.pack(fill='both', expand=True)

# Frame inferior que contém a barra de progresso e o rodapé
frame_rodape = tk.Frame(janela)
frame_rodape.pack(side='bottom', fill='x')

# Botão para importar os arquivos XML
botao = tk.Button(frame_botoes, text='Importar XML', command= importar_xmls)
botao.pack(side=tk.LEFT, padx=5)

# Botão para listar os dados simplificados dos XMLs
botao_listar = tk.Button(frame_botoes,text='Listar XML', command= listar_dados_simplificados)
botao_listar.pack(side=tk.LEFT, padx=5)

# Botão para verificar se há notas faltantes na sequência
botao_verificar = tk.Button(frame_botoes, text='Notas Faltantes', command= verificar_sequencia_numerica)
botao_verificar.pack(side=tk.LEFT, padx=5)

# Botão para mostrar a última nota emitida
botao_ultima_nota = tk.Button(frame_botoes,text='Ultima Nota Emitida', command= mostrar_ultima_nota)
botao_ultima_nota.pack(side=tk.LEFT, padx=5)

# Área de texto para exibir os resultados, com rolagem vertical
texto_resultado = tk.Text(janela, wrap='word', height=20, width=90)
scrollbar = tk.Scrollbar(janela, command=texto_resultado.yview)
texto_resultado.configure(yscrollcommand=scrollbar.set)

# Substitui a área de texto por uma versão com rolagem embutida
texto_resultado = ScrolledText(janela, width=100, height=30)
texto_resultado.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Posiciona a área de texto e a barra de rolagem
texto_resultado.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Barra de progresso para indicar o andamento da importação
barra_progresso = ttk.Progressbar(frame_rodape, orient='horizontal', length=650, mode='determinate')
barra_progresso.pack(side='top', pady=10)

# Label com crédito do desenvolvedor
label_desenvolvedor = tk.Label(frame_rodape, text='Desenvolvido por Anderson Saldanha', font=('Arial', 10), anchor='center')
label_desenvolvedor.pack(side='bottom', fill='x', pady=(0, 5))

# Inicia o loop principal da interface gráfica
janela.mainloop()
