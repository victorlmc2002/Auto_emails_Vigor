from openpyxl import load_workbook
from devedores import Devedores
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
import re
from itertools import groupby
import tkinter as tk
from tkinter import ttk, messagebox
from email.mime.image import MIMEImage

# Configurações
MES_ANTERIOR = "Junho"
MES = "07 - Julho"
EMAIL_FROM = 'vigor.imob@gmail.com'
EMAIL_PASSWORD = 'uryt lswm ptkb kuwq'
EMAIL_TO_PROP = 'victorlmc2002@gmail.com'
EMAIL_TO_INQUILINO = 'suporte@vigornegocios.com.br'

QTD = 57  # Quantidade de imóveis a serem processados
ENDERECOS = [
    "CONSTANTE RAMOS 29 APTO 1003",
    "",
    "BARÃO DO FLAMENGO 50, APARTAMENTO 401",
    "AV OSWALDO CRUZ 149 APTO 906",
    "AV OSWALDO CRUZ 149 APTO 1206",
    "VISCONDE DE CARAVELAS 98 APTO 1004",
    "AV OSWALDO CRUZ 149 APTO 2306",
    "AV EPITACIO PESSOA 2030 APTO 202",
    "",
    "AV OSWALDO CRUZ 149 APTO 105",
    "AV DULCIDIO CARDOSO 1200 APTO 305",
    "AV OSWALDO CRUZ 149 APTO 2006",
    "AV OSWALDO CRUZ 149 APTO 305",
    "AV OSWALDO CRUZ 149 APTO 2206",
    "",
    "AV OSWALDO CRUZ 99 APTO 305",
    "RUA DJALMA ULRICH 229 APTO 813",
    "RUA RODOLFO DANTAS 40 APTO 302",
    "RUA PAISSANDU 25 APTO 702",
    "AV RAINHA ELIZABETH 376 APTO 202",
    "RUA DEPUTADO SOARES FILHO 310 APTO 402",
    "RUA RIBEIRO DE ALMEIDA 21 APTO 402",
    "RUA SÁ FERREIRA 23 APTO 402",
    "AV OSWALDO CRUZ 149 APTO 1403",
    "RUA MARQUES DE ABRANTES 127 APTO 904",
    "RUA MARQUES DE ABRANTES 127 APTO 204",
    "AV CLAUDIO BESSERMAN VIANNA 3 BL1 APTO 205",
    "AV OSWALDOCRUZ 149 APTO 106",
    "RUA GETULIO DAS NEVES 31 APTO 103",
    "PRAIA DE BOTAFOGO 48 APTO 30",
    "AV OSWALDO CRUZ 67 APTO 1006",
    "RUA SENADOR VERGUEIRO 192 APTO 802",
    "RUA DAS LARANJEIRAS 29 APTO 1306",
    "RUA MARQUES DE ABRANRES 142 APTO 701",
    "AV OSWALDOCRUZ 149 APTO 1601",
    "RUA ARISTIDES ESPINOLA 20 APTO 404",
    "RUA HERMENEGILDO DE BARROS 9 APTO 1007",
    "RUA SENADOR VERGUEIRO 93 COBERTURA",
    "RUA ALMTE PEREIRA GUIMARÃES 79 APTO 301",
    "RUA REALGRANDEZA 38 APTO 302",
    "RUA MAESTRO FRANCISCO BRAGA 235 APTO 203",
    "RUA JORNALISTA ORLANDO DANTAS 12 COB 02",
    "RUA CRUZ LIMA 33 APTO 101",
    "PRAIA DO FLAMENGO 364 APTO 901",
    "RUA MARQUES DE ABRANYES 192 BL01 APTO 803",
    "RUA BELIZARIO TAVORA 77 APTO 405",
    "RUA TIMOTEO DA COSTA 1100 APTO 303",
    "RUA HUMAITA 234 BL2 APTO 704",
    "RUA VISCONDE DE PIRAJÁ 444 APTO 801",
    "RUA GUSTAVO SAMPAIO 244 BL1 APTO 1001",
    "RUA RAIMUNDO CORREA 68 LJ C",
    "RUA PRES. CARLOS DE CAMPOS 115 BL1 APTO 302",
    "RUA SENADOR VERGUEIRO 114 APTO 1102",
    "AVENIDA TIM MAIA 7585 BL3 APTO 207",
    "RUA FONTE DA SAUDADE 129 APTO 801",
    "RUA SENADOR VERGUEIRO 107 APTO 303",
    "RUA DONA MARIANA 182 APTO 103"
]
# Lista de nomes
NOMES = [
    "JESSICA FONTENELLE FREITAS",
    "",
    "LUCIANA CHARDELLI NUNES",
    "ELIANA LORENTZ CHAVES",
    "MARIANA TERK CAMPOS",
    "LUISA MACIEL CAMILLO",
    "BERNARDO DE OLIVEIRA NUNES",
    "JULIANA LEITE DE ARAUJO",
    "",
    "SUELY DA CONCEIÇÃO ALVES",
    "HENRIQUE DUTRA FRANKLIN MARTIN",
    "ITAMAR KOZNIAK",
    "GILBERTO KARPILOVSKY",
    "HERNANI AQUINIFERNADES CHAVES",
    "",
    "AIME DO CARMO RODRIGUES LOPES",
    "LORENZO COPOLLA",
    "MARCIA MARIA RAMOS DE MONCADA",
    "FERNANDO DA ROCHA VAZ BANDEIRA",
    "MONTREAL INFORMATICA AS",
    "JULIO CESAR GIOMO",
    "NIVEA MUNIZ FELIX",
    "PAULO MARCUS MOURA DE ROCHA",
    "LEONARDO ARIEL PARDO",
    "LIVIA FRANCISQUINI DE SIQUEIRA",
    "FERNADO BORTOLO DE REZENDE",
    "BIANCA DE CARVALHO TESTA ACAMPORA",
    "YURI CASTELLO BRANCO GOMES",
    "JESSICA PEREIRA SA VIANNA",
    "JULIO CEZAR PADRÃO DE OLIVEIRA",
    "LEONARDO BARROS SILVEIRA",
    "VERONICA DA COSTA DALCANAL",
    "FLAVIA CASTELLAN BRAGA",
    "KARL GEORGES MEIRELLES GALLAO",
    "MARIA ELIZABETH ALMEIDA MARQUES",
    "MARCELO FALCÃO JORDÃO RAMOS",
    "GULIHERNE OLIVEIRA SILVA",
    "MAURO JACOB LOUSADA",
    "ALEXANDRE KACELNIK",
    "THAIS PINTO COELHO DE ANDRADE",
    "ALEANDRA PEREIRA FLORIDO",
    "ANDRESSA PUHL PETRAZZINI",
    "IURI JOSÉ DE MORAES FERREIRA",
    "AQUILES POLLETI MOREIRA",
    "FARES FERREIRA PESSOA",
    "JULIA PAGY GABRIEL",
    "ANDRE GUEDES DE QUEIROZ PEREIRA",
    "NATALIA MOURA BRASIL",
    "JOÃO PAULO VERGUEIRO DE MOURA",
    "CILENE BARBOSA",
    "ERICO MARCELO CERQUEIRA ALVES",
    "VICTOR IRRMAN",
    "RODRIGO DE CARVALHO ROCHA",
    "PEDRO HENRIQUE VIEIRA GALVÃO DE LIMA",
    "ROSA MARIA GONÇALVES DE CANHA",
    "NECYLIO BEZERRA DE ARAUJO NETO",
    "IURI CARNEVALE DE CARVALHO"
]
# Lista de emails
EMAILS = [
    "fontenelle.jessica@gmail.com",
    "",
    "luciana@cn-advogados.com",
    "elianalochaves@gmail.com",
    "m-terk@uol.com.br",
    "luisa.macielc@gmail.com",
    "bernardo.nunes@gmail.com",
    "jully_uff@yahoo.com.br",
    "",
    "suelycascoutinho@hotmail.com",
    "henriquemdutra@hotmail.com",
    "itamarkozniak@gmail.com",
    "gilbertocarpi@terra.com.br",
    "anacristinabfc@gmail.com",
    "",
    "aime_lopes@hotmail.com",
    "lorenzocoppola17@gmail.com",
    "marciamoncada@yahoo.com.br",
    "fnando.demelo@pm.me",
    "contasapagar@montreal.com.br",
    "julio-cesargiomo@hotmail.com",
    "niveafelix75@gmail.com",
    "paulomarcusferreira@hotmail.com",
    "pardo.leonardo@gmail.com",
    "liviafsiqueira15@gmail.com",
    "fernandorezzende@gmail.com",
    "matheuszugepaz@hotmail.com.br",
    "yuricbgomes@hotmail.com",
    "jessica.carine@hotmail.com",
    "jcpdeol@gmail.com",
    "leopacheco@poli.ufrj.br",
    "vdalcanal@yahoo.com.br",
    "fcbraga@hotmail.com",
    "karlmuhs@hotmail.com",
    "zazadelbosco@gmail.com",
    "marcelofalcaodf@gmail.com",
    "guilhermeos91@gmail.com",
    "maurojlosada@hotmail.com",
    "kacelnik17@gmail.com",
    "thaispcandrade@gmail.com",
    "florido.aleandra@gmail.com",
    "andressa_petrazzini@hotmail.com",
    "iuriucm@hotmail.com",
    "aqpomo@gmail.com",
    "farespessoa@gmail.com",
    "juliapagy@gmail.com",
    "andre.gqp@gmail.com",
    "natalia.mbrazil@gmail.com",
    "jpaulocampos@hotmail.com",
    "cilene.cbrj@gmail.com",
    "emcarj@gmail.com",
    "victor.irrmann@gmail.com",
    "rodrangra@hotmail.com",
    "pedrohglima1@gmail.com",
    "rosacanha@uol.com.br",
    "necylioneto@icloud.com",
    "roscarnevale01@gmail.com",
]
# Caminhos das pastas
BASE_PATH = Path('C:/pasta marcelo/Administradora de Imóveis/Boletos 2025')
PASTA_BOLETOS = BASE_PATH / MES / 'Boletos'
print(PASTA_BOLETOS)
PASTA_CONDOMINIO = BASE_PATH / MES / 'Taxa de Condomínio'
PASTA_REPASSES = BASE_PATH / MES / 'Repasses'
ARQ_EXCEL = BASE_PATH / 'Planilha nova 2025.xlsx'

class EmailInterface:
    def __init__(self, devedores):
        self.devedores = devedores
        self.selected_items = set()
        self.root = tk.Tk()
        self.root.title("Envio de Emails - Sistema de Boletos")
        self.root.geometry("800x600")
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main_frame, text="Selecione os imóveis para envio de emails", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Frame de seleção
        selection_frame = ttk.Frame(main_frame)
        selection_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Lista de imóveis com scrollbar
        self.tree = ttk.Treeview(selection_frame, columns=('Nome', 'Valor'), show='headings')
        self.tree.heading('Nome', text='Proprietário/Inquilino')
        self.tree.heading('Valor', text='Endereço')
        
        vsb = ttk.Scrollbar(selection_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Preencher a lista
        for i, devedor in enumerate(self.devedores, 1):
            self.tree.insert("", "end", values=(f"{i} - {devedor._nome_inquilino}", f"{devedor._endereco}"), tags=(str(i),))
        
        # Configurar tags para seleção
        self.tree.tag_configure('selected', background='lightblue')
        
        # Opções de envio
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=10)
        
        self.recipient_var = tk.StringVar(value="inquilino")
        ttk.Radiobutton(options_frame, text="Enviar para Inquilino", variable=self.recipient_var, 
                       value="inquilino").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(options_frame, text="Enviar para Proprietário", variable=self.recipient_var, 
                       value="proprietario").pack(side=tk.LEFT, padx=10)
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Selecionar Todos", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Desmarcar Todos", command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Enviar Selecionados", command=self.send_selected).pack(side=tk.RIGHT, padx=5)
        
        # Configurar seleção múltipla
        self.tree.bind("<Button-1>", self.on_click)
    
    def on_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            tags = self.tree.item(item, "tags")
            if tags:
                idx = int(tags[0])
                if idx in self.selected_items:
                    self.selected_items.remove(idx)
                    self.tree.item(item, tags=(tags[0],))
                else:
                    self.selected_items.add(idx)
                    self.tree.item(item, tags=(tags[0], 'selected'))
    
    def select_all(self):
        self.selected_items = set(range(1, len(self.devedores)+1))
        for item in self.tree.get_children():
            tags = self.tree.item(item, "tags")
            if tags:
                self.tree.item(item, tags=(tags[0], 'selected'))
    
    def deselect_all(self):
        self.selected_items = set()
        for item in self.tree.get_children():
            tags = self.tree.item(item, "tags")
            if tags:
                self.tree.item(item, tags=(tags[0],))
    
    def send_selected(self):
        if not self.selected_items:
            messagebox.showwarning("Nenhum selecionado", "Por favor, selecione pelo menos um imóvel.")
            return
        
        recipient_type = self.recipient_var.get()
        confirm = messagebox.askyesno(
            "Confirmar envio",
            f"Você está prestes a enviar emails para {len(self.selected_items)} imóveis ({recipient_type}s).\nDeseja continuar?"
        )
        
        if not confirm:
            return
        
        self.root.config(cursor="watch")
        self.root.update()
        
        try:
            for idx in self.selected_items:
                devedor = self.devedores[idx-1]
                if recipient_type == "inquilino":
                    enviar_email_inquilino(devedor)
                else:
                    enviar_email_proprietario(devedor)
            
            messagebox.showinfo("Sucesso", "Emails enviados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao enviar os emails:\n{str(e)}")
        finally:
            self.root.config(cursor="")
    
    def run(self):
        self.root.mainloop()

def extrair_numero(nome_arquivo):
    """Extrai número do nome do arquivo para ordenação."""
    match = re.search(r'\d+', nome_arquivo.stem)
    return int(match.group()) if match else float('inf')

def processar_pasta(pasta_path):
    """Processa arquivos PDF em uma pasta, agrupando por número."""
    arquivos = sorted(pasta_path.glob('*.pdf'), key=extrair_numero)
    
    grupos = []
    for key, group in groupby(arquivos, key=extrair_numero):
        grupo = list(group)
        grupos.append([str(p) for p in grupo] if len(grupo) > 1 else str(grupo[0]))
    
    return grupos

def ler_planilha(excel, n):
    """Lê dados da planilha e cria lista de devedores."""
    wb = load_workbook(excel, data_only=True)
    planilha = wb[MES]
    
    devedores = []
    imovel = 0
    x = 5
    pdfs = processar_pasta(PASTA_BOLETOS)
    conds = processar_pasta(PASTA_CONDOMINIO)
    repasses = processar_pasta(PASTA_REPASSES)
    while imovel < n:
        for y in range(2, 17, 5):
            if imovel >= n:
                break
            #print(x, y)
            #print(imovel)
            dados = extrair_dados_linha(planilha, x, y)
            
            devedor = criar_devedor(
                ENDERECOS[imovel], NOMES[imovel], EMAILS[imovel], dados, 
                pdfs[imovel] if imovel < len(pdfs) else None,
                conds[imovel] if imovel < len(conds) else None,
                repasses[imovel] if imovel < len(repasses) else None
            )
            print(devedor)
            devedores.append(devedor)
            imovel += 1
        x += 17
    # for _ in range(19):
    #     temp = x
    #     for y in range(2, 17, 5):
    #         dados = extrair_dados_linha(planilha, x, y)
    #         pdfs = processar_pasta(PASTA_BOLETOS)
    #         conds = processar_pasta(PASTA_CONDOMINIO)
    #         repasses = processar_pasta(PASTA_REPASSES)
            
    #         devedor = criar_devedor(
    #             ENDERECOS[imovel], NOMES[imovel], EMAILS[imovel], dados, 
    #             pdfs[imovel] if imovel < len(pdfs) else None,
    #             conds[imovel] if imovel < len(conds) else None,
    #             repasses[imovel] if imovel < len(repasses) else None
    #         )
    #         devedores.append(devedor)
    #         imovel += 1
    #         x = temp
    #     x += 17
    
    return devedores

def extrair_dados_linha(planilha, x, y):
    """Extrai dados de uma linha específica da planilha."""
    dados = {
        'aluguel': None,
        'iptu': None,
        'cond': None,
        'valor': None,
        'cotas_extras': [],
        'taxa': None,
        'tarifa': None,
        'repasse': None
    }
    index = 0
    
    while planilha.cell(row=x, column=y).value != "Receita":
        celula = planilha.cell(row=x, column=y)
        valor = planilha.cell(row=x, column=y+1).value
        #print(celula.value, valor)
        if celula.value == "Aluguel":
            dados['aluguel'] = valor
        elif celula.value == "Iptu cota":
            dados['iptu'] = valor
        elif celula.value == "Taxa de Condomínio":
            dados['cond'] = valor
        elif celula.value == "Valor Boleto":
            dados['valor'] = valor
        elif celula.value is None:
            index = 1
        elif celula.value is not None and index == 0 and (float(valor) > 0):
            dados['cotas_extras'].append({celula.value: valor})
        
        x += 1
    
    dados['taxa'] = planilha.cell(row=x+2, column=y+1).value
    dados['tarifa'] = planilha.cell(row=x+2, column=y+3).value
    dados['repasse'] = planilha.cell(row=x+3, column=y+1).value
    
    return dados

def criar_devedor(endereco, nome, email, dados, pdf, cond, repasse):
    """Cria um objeto Devedor com os dados fornecidos."""
    return Devedores(
        endereco, nome, email, dados['valor'], dados['aluguel'], dados['iptu'], dados['cond'], 
        dados['cotas_extras'], dados['repasse'], dados['taxa'], dados['tarifa'],
        pdf, None if dados['cond'] is not None else cond, repasse
    )

def formatar_texto_inquilino(devedor):
    """Formata texto do e-mail para inquilinos."""
    partes = [f"Aluguel: R$ {devedor._aluguel:.2f}"]
    
    if not isinstance(devedor._iptu, str) and devedor._iptu is not None and devedor._iptu > 0:
        partes.append(f"IPTU: R$ {devedor._iptu:.2f}")
    if not isinstance(devedor._cond, str) and devedor._cond is not None and devedor._cond > 0:
        partes.append(f"Taxa de condomínio: R$ {devedor._cond:.2f}")
    if devedor._cotas_extras:
        partes.append("<br><b>Descontos/Reembolsos:<br></b>")
        for cota in devedor._cotas_extras:
            for chave, valor in cota.items():
                if valor > 0:
                    partes.append(f"{chave}: R$ {valor:.2f}")
    
    partes.append(f"<br><b>Valor do Boleto: R$ {devedor._valor:.2f}<br><br>Atenciosamente,<br>Vigor Negócios<br></b>")
    return "<br>".join(partes)

def formatar_texto_proprietario(devedor):
    """Formata texto do e-mail para proprietários."""
    partes = [f"Receita: R$ {devedor._valor:.2f}", "<br><b>Reembolso de Despesas:</b><br>"]
    
    if not isinstance(devedor._iptu, str) and devedor._iptu is not None and devedor._iptu > 0:
        partes.append(f"IPTU: R$ {devedor._iptu:.2f}")
    if not isinstance(devedor._cond, str) and devedor._cond is not None and devedor._cond > 0:
        partes.append(f"Taxa de condomínio: R$ {devedor._cond:.2f}")
    if not isinstance(devedor._taxa, str) and devedor._cond is not None and devedor._cond > 0:
        partes.append(f"Taxa de administração: R$ {devedor._taxa:.2f}")
    if not isinstance(devedor._tarifa, str) and devedor._cond is not None and devedor._cond > 0:
        partes.append(f"Tarifa de transferência: R$ {devedor._tarifa:.2f}")    
    partes.append(f"<br><b>Valor do Repasse: R$ {devedor._repasse:.2f}</b>")
    return "<br>".join(partes)

def extrair_mensagem_assunto(caminho_arquivo):
    """Extrai mensagem para o assunto do e-mail a partir do nome do arquivo."""
    if isinstance(caminho_arquivo, list):
        caminho_arquivo = caminho_arquivo[0]
    
    nome_arquivo = Path(caminho_arquivo).stem
    partes = nome_arquivo.split(" - ", 1)[1]
    return partes.rsplit(" - ", 1)[0] if " - " in partes else partes

def criar_email(devedor, texto_formatado, anexos, imagem="lg.png"):
    """Cria e configura mensagem de e-mail, com opção de imagem embutida."""
    print(devedor._email_inquilino)
    if not devedor._endereco or not devedor._nome_inquilino or not devedor._email_inquilino:
        raise ValueError("Imóvel inválido: endereço, nome ou email não fornecidos.")
    msg = MIMEMultipart('related')
    msg['Subject'] = extrair_mensagem_assunto(anexos[0])
    msg['From'] = EMAIL_FROM
    msg['To'] = devedor._email_inquilino

    # Corpo do email (com imagem embutida se fornecida)
    if devedor._cond is None:
        corpo_email = f"""
        <p></p>
        Boa tarde, {devedor._nome_inquilino.split()[0].capitalize()}<br>
        Segue o boleto referente ao  aluguel do mês de {MES_ANTERIOR} e taxa de condomínio referente ao mês vigente em anexo<br>
        <p><b>Valores:</b><br></p>
        {texto_formatado}  
        """
    else:
        corpo_email = f"""
        <p></p>
        Boa tarde, {devedor._nome_inquilino.split()[0].capitalize()}<br>
        Segue o boleto referente ao mês de {MES_ANTERIOR} em anexo<br>
        <p><b>Valores:</b><br></p>
        {texto_formatado}  
        """

    if imagem:
        # Adiciona a tag <img> ao corpo do email
        corpo_email += '<br><img src="cid:imagem1"><br>'

    msg_alt = MIMEMultipart('alternative')
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText(corpo_email, 'html'))

    # Anexa imagem embutida, se fornecida
    if imagem:
        with open(imagem, 'rb') as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header('Content-ID', '<imagem1>')
            mime_img.add_header('Content-Disposition', 'inline', filename=Path(imagem).name)
            msg.attach(mime_img)

    # Anexa arquivos PDF
    for anexo in anexos:
        if anexo:
            with open(anexo, 'rb') as f:
                part = MIMEApplication(f.read(), _subtype='pdf')
                part.add_header('Content-Disposition', 'attachment', filename=Path(anexo).name)
                msg.attach(part)

    return msg

def enviar_email(msg):
    """Envia e-mail configurado."""
    with smtplib.SMTP('smtp.gmail.com', 587) as s:
        s.starttls()
        s.login(msg['From'], EMAIL_PASSWORD)
        s.send_message(msg)
    print(f'Email enviado para {msg["To"]}')

def enviar_email_proprietario(devedor):
    """Envia e-mail para proprietário."""
    if not devedor._pdfrepasse:
        return
    
    msg = criar_email(
        devedor,
        formatar_texto_proprietario(devedor),
        [devedor._pdfrepasse] if isinstance(devedor._pdfrepasse, str) else devedor._pdfrepasse
    )
    enviar_email(msg)

def enviar_email_inquilino(devedor):
    """Envia e-mail para inquilino."""
    if not devedor._pdfboleto:
        return
    
    anexos = [devedor._pdfboleto]
    if devedor._pdfcond:
        anexos.append(devedor._pdfcond)
    
    msg = criar_email(
        devedor,
        formatar_texto_inquilino(devedor),
        anexos
    )
    enviar_email(msg)

def main():
    """Função principal."""
    devedores = ler_planilha(ARQ_EXCEL, QTD)

    app = EmailInterface(devedores)
    app.run()

if __name__ == "__main__":
    main()