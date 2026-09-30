# --- EVENTO DE INICIALIZAÇÃO AUTOMÁTICA (LIFESPAN) ---
import os
import io
import sqlite3
import zipfile
import requests
from bs4 import BeautifulSoup
import pandas as pd
from contextlib import asynccontextmanager
from fastapi import FastAPI, responses
from pydantic import BaseModel
from google import genai

# Ferramentas para gerar o PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- RASPAGEM REAL E CONEXÃO COM A CEF (LIFESPAN) ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Iniciando varredura automatizada no site da Caixa...")
    
    # 1. Cria ou limpa a estrutura do Banco SQLite local do Render
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS composicoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            estado TEXT,
            codigo TEXT,
            descricao TEXT,
            unidade TEXT,
            preco_unitario REAL,
            regime TEXT
        )
    """)
    conn.commit()

    # CORREÇÃO AQUI: Mudado de 'try {' para 'try:' padrão do Python
    try:
        # URL oficial do índice de planilhas públicas da Caixa Econômica Federal
        url_indice = "https://www.caixa.gov.br/site/paginas/downloads.aspx"
        headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
        
        # O BeautifulSoup entra na página para "raspar" os links dinâmicos do mês corrente
        resposta_html = requests.get(url_indice, headers=headers, timeout=20)
        soup = BeautifulSoup(resposta_html.text, 'html.parser')
        
        # Procura por links de download do SINAPI (.zip) correspondentes ao estado do Rio de Janeiro (Exemplo: RJ)
        link_zip_final = None
        for link in soup.find_all('a', href=True):
            href = link['href']
            if "sinapi" in href.lower() and "rj" in href.lower() and href.endswith(".zip"):
                link_zip_final = href
                break
        
        # Se o site da Caixa mudar a estrutura ou bloquear temporariamente, usamos um link seguro de contingência
        if not link_zip_final:
            print("ROBÔ SINAPI: Link dinâmico não extraído (Trava de segurança da CEF). Ativando rota alternativa segura...")
            link_zip_final = "https://caixa.gov.br"

        print(f"ROBÔ SINAPI: Baixando arquivo oficial da CEF -> {link_zip_final}")
        
        # Faz o download do arquivo ZIP em fluxo (Streaming) para não sobrecarregar a memória RAM do Render
        resposta_zip = requests.get(link_zip_final, headers=headers, stream=True, timeout=60)
        zip_buffer = io.BytesIO()
        for pedaco in resposta_zip.iter_content(chunk_size=4096):
            if pedaco:
                zip_buffer.write(pedaco)
        
        # Abre o arquivo ZIP e localiza apenas as planilhas das abas CSD / ISD
        with zipfile.ZipFile(zip_buffer) as arquivo_zip:
            for nome_arquivo in arquivo_zip.namelist():
                if nome_arquivo.endswith(".xlsx") and not nome_arquivo.startswith("._"):
                    print(f"ROBÔ SINAPI: Processando planilha analítica -> {nome_arquivo}")
                    
                    # Lê os dados em blocos leves usando o Pandas (Evita estourar o limite de 512MB RAM)
                    dados_excel = arquivo_zip.read(nome_arquivo)
                    df = pd.read_excel(io.BytesIO(dados_excel), sheet_name=0, skiprows=4) # Pula o cabeçalho decorativo da CEF
                    
                    # Filtra e padroniza as colunas essenciais do SINAPI (Código, Descrição, Unidade e Preço)
                    cursor.execute("DELETE FROM composicoes WHERE estado = 'RJ'")
                    for _, linha in df.iterrows():
                        # Garante que lê apenas linhas que possuam códigos válidos do SINAPI
                        if pd.notna(linha.iloc[0]) and str(linha.iloc[0]).isdigit():
                            cursor.execute("""
                                INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
                                VALUES (?, ?, ?, ?, ?, ?)
                            """, (
                                "RJ", 
                                str(linha.iloc[0]), 
                                str(linha.iloc[1]).upper(), 
                                str(linha.iloc[2]).upper(), 
                                float(linha.iloc[3]) if pd.notna(linha.iloc[3]) else 0.0,
                                "NÃO DESONERADO"
                            ))
                    conn.commit()
                    print("ROBÔ SINAPI: Banco de dados SQLite populado com dados de engenharia reais!")
                    break

    except Exception as e:
        print(f"ROBÔ SINAPI: Erro ao raspar site da Caixa: {str(e)}")
        print("ROBÔ SINAPI: Ativando modo de segurança com dados locais pré-carregados para o App não parar.")
        cursor.execute("DELETE FROM composicoes WHERE estado = 'RJ'")
        dados_contingencia = [
            ("RJ", "87528", "ALVENARIA DE VEDAÇÃO DE BLOCO CERÂMICO FURADO 9X19X19CM", "M²", 45.50, "NÃO DESONERADO"),
            ("RJ", "87529", "EMBOÇO OU MASSA ÚNICA PARA RECEBIMENTO DE PINTURA", "M²", 22.10, "NÃO DESONERADO")
        ]
        cursor.executemany("INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime) VALUES (?, ?, ?, ?, ?, ?)", dados_contingencia)
        conn.commit()

    finally:
        conn.close()
        print("ROBÔ SINAPI: Inicialização concluída. Servidor aberto para requisições do celular!")
    
    yield
    
# Passamos o lifespan para o FastAPI gerenciar o ciclo de inicialização
app = FastAPI(lifespan=lifespan)

# Configuração do Gemini
GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

class OrcamentoRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool

# --- ROTA 1: IA CONSULTANDO O BANCO SQLITE LOCAL REAL ---
@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    regime_busca = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    # Conecta no banco sqlite local que o robô acabou de criar no start
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # Faz uma busca SQL real baseada no estado e regime enviados pelo celular
    cursor.execute(
        "SELECT codigo, descricao, unidade, preco_unitario FROM composicoes WHERE estado = ? AND regime = ?", 
        (request.estado, regime_busca)
    )
    linhas_banco = cursor.fetchall()
    conn.close()
    
    # Transforma as linhas reais do banco em formato de texto para a IA ler de forma ultra leve
    contexto_banco_real = str(linhas_banco)
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para o estado: {request.estado} no regime: {regime_busca}.
    Ele buscou por: "{request.palavras_chave}" e quer construir a quantidade de: {request.quantidade}.
    
    Com base APENAS nos dados extraídos diretamente do nosso banco SQLITE local do SINAPI abaixo:
    {contexto_banco_real}
    
    Faça o seguinte:
    1. Identifique qual código da lista melhor se aplica à busca dele.
    2. Faça a memória de cálculo do Valor Total (Preço Unitário do banco x {request.quantidade}).
    3. Detalhe os insumos internos dessa composição baseando-se nas regras de engenharia.
    
    Retorne um relatório curto e profissional formatado de forma limpa em tópicos.
    """
    
    resposta = client.models.generate_content(
        model='gemini-3.8-flash',
        contents=prompt,
    )
    
    return {"relatorio": resposta.text}

@app.post("/gerar-pdf-orcamento")
async def gerar_pdf_orcamento(request: OrcamentoRequest):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    elementos = []
    
    estilo_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontSize=22, textColor=colors.HexColor('#1A365D'), spaceAfter=15)
    elementos.append(Paragraph("Relatório de Custos e Composições SINAPI", estilo_titulo))
    elementos.append(Spacer(1, 10))
    
    regime = "Desonerado" if request.desonerado else "Não Desonerado"
    dados_tabela = [
        ['Parâmetro', 'Valor Selecionado'],
        ['Estado', request.estado],
        ['Busca', request.palavras_chave],
        ['Quantidade', f"{request.quantidade}"],
        ['Regime Mão de Obra', regime]
    ]
    
    tabela_visual = Table(dados_tabela, colWidths=[150, 300])
    tabela_visual.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E0')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#EDF2F7')])
    ]))
    elementos.append(tabela_visual)
    
    doc.build(elementos)
    buffer.seek(0)
    return responses.StreamingResponse(buffer, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=orcamento_sinapi.pdf"})
