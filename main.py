# --- EVENTO DE INICIALIZAÇÃO AUTOMÁTICA (LIFESPAN) ---
import os
import io
import sqlite3
import zipfile
import requests
import json
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

# --- MODELO DE DADOS PARA RECEBER O PDF CALCULADO ---
class PDFRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool
    relatorio_texto: str

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

# --- ROTA 1 ATUALIZADA: RETORNA TEXTO + TABELA ESTRUTURADA ---
@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    termo_busca = f"%{request.palavras_chave.upper().strip()}%"
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT codigo, descricao, unidade, preco_unitario 
        FROM composicoes 
        WHERE descricao LIKE ? OR codigo = ?
        LIMIT 20
    """, (termo_busca, request.palavras_chave.strip()))
    linhas_banco = cursor.fetchall()
    conn.close()
    
    contexto_banco_real = str(linhas_banco) if linhas_banco else "Nenhum item correspondente exato foi localizado."
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para: "{request.palavras_chave}" (Qtd: {request.quantidade}) no estado: {request.estado.upper()} ({regime_texto}).
    
    Dados extraídos do banco SQLite:
    {contexto_banco_real}
    
    Responda RIGOROSAMENTE no formato JSON abaixo, contendo o relatório em Markdown e uma lista com até 4 insumos/materiais principais gerados analiticamente. Não adicione nenhuma palavra fora do JSON.
    
    Modelo de Resposta Esperada:
    {{
        "relatorio": "Texto do relatório aqui...",
        "insumos": [
            {{"nome": "Nome do Material 1", "qtd": 1.5, "unidade": "KG", "total": "R$ 15,00"}},
            {{"nome": "Nome do Profissional 2", "qtd": 2.0, "unidade": "H", "total": "R$ 50,00"}}
        ]
    }}
    """
    
    resposta = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
    
    # Limpa possíveis formatações de bloco de código que a IA coloca (```json ... ```)
    texto_limpo = resposta.text.strip().removeprefix("```json").removesuffix("```").strip()
    return responses.PlainTextResponse(texto_limpo, media_type="application/json")

# --- ROTA 2 CORRIGIDA: MOTOR DE GERAÇÃO DO PDF SEM ERROS ---
@app.post("/gerar-pdf-orcamento")
async def gerar_pdf_orcamento(request: PDFRequest):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    elementos = []
    
    estilo_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontSize=22, textColor=colors.HexColor('#1A365D'), spaceAfter=15)
    elementos.append(Paragraph("Relatório de Custos e Composições SINAPI", estilo_titulo))
    elementos.append(Spacer(1, 10))
    
    regime = "Desonerado" if request.desonerado else "Não Desonerado"
    
    # Criando os parágrafos de texto do relatório com quebra de linha correta para o PDF
    estilo_corpo = styles['Normal']
    elementos.append(Paragraph(f"<b>Estado:</b> {request.estado} | <b>Regime:</b> {regime}", estilo_corpo))
    elementos.append(Paragraph(f"<b>Serviço Solicitado:</b> {request.palavras_chave} (Quantidade: {request.quantidade})", estilo_corpo))
    elementos.append(Spacer(1, 15))
    
    # Adiciona o resumo textual estruturado enviado pelo app
    elementos.append(Paragraph("<b>Detalhamento do Orçamento:</b>", styles['Heading3']))
    elementos.append(Spacer(1, 5))
    elementos.append(Paragraph(request.relatorio_texto.replace("\n", "<br/>"), estilo_corpo))
    
    doc.build(elementos)
    buffer.seek(0)
    return responses.StreamingResponse(buffer, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=orcamento_sinapi.pdf"})