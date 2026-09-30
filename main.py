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

# --- CARREGAMENTO DO ARQUIVO CONSOLIDADO NACIONAL (LIFESPAN COBERTURA NACIONAL) ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Inicializando leitura da planilha nacional unificada...")
    
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

    # Nome exato do arquivo consolidado que você mencionou
    nome_arquivo_local = "sinapi_2026_08.xlsx"

    try:
        if os.path.exists(nome_arquivo_local):
            print(f"ROBÔ SINAPI: Localizado arquivo unificado {nome_arquivo_local}. Iniciando indexação...")
            
            # Carrega o arquivo Excel para inspecionar todas as abas (Estados/Regimes)
            excel_file = pd.ExcelFile(nome_arquivo_local, engine="openpyxl")
            cursor.execute("DELETE FROM composicoes") # Limpa o banco anterior
            
            # Varre cada aba da planilha nacional automaticamente
            for nome_aba in excel_file.sheet_names:
                nome_aba_upper = nome_aba.upper().strip()
                
                # CORREÇÃO 1: Ignora abas decorativas que não são estados
                if nome_aba_upper in ["MENU", "BUSCA", "LEIAME", "INSTRUÇÕES"]:
                    print(f"ROBÔ SINAPI: Ignorando aba decorativa [{nome_aba}].")
                    continue
                
                if len(nome_aba_upper) >= 2:
                    estado_aba = nome_aba_upper[:2]
                    regime_aba = "DESONERADO" if "DES" in nome_aba_upper and "NDES" not in nome_aba_upper else "NÃO DESONERADO"
                    
                    print(f"ROBÔ SINAPI: Indexando dados de {estado_aba} ({regime_aba}) da aba [{nome_aba}]...")
                    
                    df = pd.read_excel(excel_file, sheet_name=nome_aba, skiprows=4)
                    
                    for _, linha in df.iterrows():
                        # CORREÇÃO 2: Garante de forma segura que a linha possui colunas (tamanho) antes de ler o .iloc
                        if len(linha) > 0 and pd.notna(linha.iloc) and str(linha.iloc).strip().isdigit():
                            
                            # Define os preços de forma segura, garantindo que a coluna existe
                            preco = 0.0
                            if len(linha) > 3 and pd.notna(linha.iloc):
                                try:
                                    preco = float(linha.iloc)
                                except:
                                    preco = 0.0
                                    
                            cursor.execute("""
                                INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
                                VALUES (?, ?, ?, ?, ?, ?)
                            """, (
                                estado_aba,
                                str(linha.iloc).strip(),
                                str(linha.iloc).upper().strip(),
                                str(linha.iloc).upper().strip() if len(linha) > 2 and pd.notna(linha.iloc) else "-",
                                preco,
                                regime_aba
                            ))
            conn.commit()
            print("ROBÔ SINAPI: Sucesso Absoluto! Base Nacional SQLite populado com dados de engenharia reais!")
        else:
            print(f"ROBÔ SINAPI: Arquivo {nome_arquivo_local} não localizado na raiz.")
            raise FileNotFoundError()

    except Exception as e:
        print(f"ROBÔ SINAPI: Erro ao processar Planilha Nacional: {str(e)}")
        # Contingência básica caso o arquivo falhe ao abrir
        cursor.execute("DELETE FROM composicoes")
        cursor.execute("""
            INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
            VALUES ('RJ', '87528', 'ALVENARIA DE VEDAÇÃO DE BLOCO CERÂMICO FURADO', 'M²', 45.50, 'NÃO DESONERADO')
        """)
        conn.commit()
    finally:
        conn.close()
        print("ROBÔ SINAPI: Inicialização e carregamento concluídos!")
    
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

# --- ATUALIZAÇÃO DA ROTA DE CONSULTA DA IA ---
@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    regime_busca = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    estado_busca = request.estado.upper().strip()
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # Busca cirúrgica filtrando dinamicamente por Estado e Regime selecionados no Celular!
    cursor.execute(
        "SELECT codigo, descricao, unidade, preco_unitario FROM composicoes WHERE estado = ? AND regime = ?", 
        (estado_busca, regime_busca)
    )
    linhas_banco = cursor.fetchall()
    conn.close()
    
    contexto_banco_real = str(linhas_banco)
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para o estado: {estado_busca} no regime: {regime_busca}.
    Ele buscou por: "{request.palavras_chave}" e quer construir a quantidade de: {request.quantidade}.
    
    Com base APENAS nos dados extraídos diretamente do nosso banco SQLITE local do SINAPI abaixo:
    {contexto_banco_real}
    
    Faça o seguinte:
    1. Identifique qual código da lista melhor se aplica à busca dele.
    2. Faça a memória de cálculo do Valor Total (Preço Unitário do banco x {request.quantidade}).
    3. Detalhe os insumos internos dessa composição baseando-se nas regras de engenharia.
    
    Retorne um relatório curto e profissional formatado de forma limpa em tópicos.
    """
    
    resposta = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
    return {"relatorio": resposta.text}

# --- ROTA DE GERAR PDF ---
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
