import os
import io
import sqlite3
import pandas as pd
import json
from contextlib import asynccontextmanager
from fastapi import FastAPI, responses
from pydantic import BaseModel
from google import genai
from google.genai import types

# Ferramentas para gerar o PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- MODELO DE DADOS ENVIADO PELO APP ---
class OrcamentoRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool

class PDFRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool
    relatorio_texto: str

# --- INICIALIZAÇÃO E LEITURA DA PLANILHA NACIONAL PROTEGIDA (LIFESPAN) ---
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

    nome_arquivo_local = "sinapi_2026_08.xlsx"

    try:
        if os.path.exists(nome_arquivo_local):
            print(f"ROBÔ SINAPI: Localizado arquivo unificado {nome_arquivo_local}. Iniciando indexação...")
            excel_file = pd.ExcelFile(nome_arquivo_local, engine="openpyxl")
            cursor.execute("DELETE FROM composicoes")
            
            for nome_aba in excel_file.sheet_names:
                nome_aba_upper = nome_aba.upper().strip()
                
                if nome_aba_upper in ["MENU", "BUSCA", "LEIAME", "INSTRUÇÕES"]:
                    print(f"ROBÔ SINAPI: Ignorando aba decorativa [{nome_aba}].")
                    continue
                
                if len(nome_aba_upper) >= 2:
                    estado_aba = nome_aba_upper[:2]
                    regime_aba = "DESONERADO" if "DES" in nome_aba_upper and "NDES" not in nome_aba_upper else "NÃO DESONERADO"
                    
                    print(f"ROBÔ SINAPI: Indexando dados de {estado_aba} ({regime_aba}) da aba [{nome_aba}]...")
                    df = pd.read_excel(excel_file, sheet_name=nome_aba, skiprows=4)
                    
                    for _, linha in df.iterrows():
                        # LÓGICA DE PROTEÇÃO DE ÍNDICES CONTRA CÉLULAS VAZIAS (EVITA ERRO ASGI) [1.2]
                        if len(linha) > 0 and pd.notna(linha.iloc[0]) and str(linha.iloc[0]).strip().isdigit():
                            preco = 0.0
                            if len(linha) > 7 and pd.notna(linha.iloc[7]):
                                try:
                                    preco = float(linha.iloc[7])
                                except:
                                    preco = 0.0
                                    
                            cursor.execute("""
                                INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
                                VALUES (?, ?, ?, ?, ?, ?)
                            """, (
                                estado_aba,
                                str(linha.iloc[0]).strip(),
                                str(linha.iloc[1]).upper().strip(),
                                str(linha.iloc[2]).upper().strip() if len(linha) > 2 and pd.notna(linha.iloc[2]) else "-",
                                preco,
                                regime_aba
                            ))
            conn.commit()
            print("ROBÔ SINAPI: Base Nacional SQLite populada com dados de engenharia reais!")
        else:
            print(f"ROBÔ SINAPI: Arquivo {nome_arquivo_local} não localizado na raiz.")
            raise FileNotFoundError()

    except Exception as e:
        print(f"ROBÔ SINAPI: Erro ao processar Planilha Nacional: {str(e)}")
        cursor.execute("DELETE FROM composicoes")
        cursor.execute("""
            INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
            VALUES ('RJ', '87528', 'ALVENARIA DE VEDAÇÃO DE BLOCO CERÂMICO FURADO', 'M²', 45.50, 'NÃO DESONERADO')
        """)
        conn.commit()
    finally:
        conn.close()
        print("ROBÔ SINAPI: Inicialização concluída!")
    
    yield

app = FastAPI(lifespan=lifespan)

GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

# --- ROTA 1: RETORNA TEXTO DO CARD + LISTA DA TABELA EM JSON ---
@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    termo_busca = f"%{request.palavras_chave.upper().strip()}%"
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # Filtra por Estado, Regime e Palavra-Chave na tabela nacional unificada
    cursor.execute("""
        SELECT codigo, descricao, unidade, preco_unitario 
        FROM composicoes 
        WHERE (descricao LIKE ? OR codigo = ?) AND estado = ? AND regime = ?
        LIMIT 20
    """, (termo_busca, request.palavras_chave.strip(), request.estado.upper().strip(), "DESONERADO" if request.desonerado else "NÃO DESONERADO"))
    linhas_banco = cursor.fetchall()
    conn.close()
    
    contexto_banco_real = str(linhas_banco) if linhas_banco else "Nenhum item correspondente localizado."
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para: "{request.palavras_chave}" (Qtd: {request.quantidade}) no estado: {request.estado.upper()} ({regime_texto}).
    
    Dados reais extraídos do nosso banco SQLite:
    {contexto_banco_real}
    
    Com base nesses dados, preencha os campos obrigatórios do relatório descritivo e da lista analítica de insumos associados.
    """

    # 🚨 A MÁGICA DA BLINDAGEM: Força o Gemini a obedecer a estrutura exata do JSON
    config_ia = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema={
            "type": "OBJECT",
            "properties": {
                "relatorio": {"type": "STRING"},
                "insumos": {
                    "type": "ARRAY",
                    "items": {
                        "type": "OBJECT",
                        "properties": {
                            "nome": {"type": "STRING"},
                            "qtd": {"type": "NUMBER"},
                            "unidade": {"type": "STRING"},
                            "total": {"type": "STRING"}
                        },
                        "required": ["nome", "qtd", "unidade", "total"]
                    }
                }
            },
            "required": ["relatorio", "insumos"]
        }
    )
    
    # Executa a chamada passando a configuração de segurança rígida
    resposta = client.models.generate_content(
        model='gemini-3.8-flash',
        contents=prompt,
        config=config_ia
    )
    
    # Como o Google garante o envio do JSON puro, devolvemos direto para o celular sem risco de erro ASGI!
    return responses.PlainTextResponse(resposta.text, media_type="application/json")

# --- ROTA 2: GERAÇÃO DO ARQUIVO PDF CORPORATIVO ---
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
    estilo_corpo = styles['Normal']
    elementos.append(Paragraph(f"<b>Estado:</b> {request.estado} | <b>Regime:</b> {regime}", estilo_corpo))
    elementos.append(Paragraph(f"<b>Busca:</b> {request.palavras_chave} (Quantidade: {request.quantidade})", estilo_corpo))
    elementos.append(Spacer(1, 15))
    
    elementos.append(Paragraph("<b>Detalhamento Completo do Orçamento:</b>", styles['Heading3']))
    elementos.append(Spacer(1, 5))
    elementos.append(Paragraph(request.relatorio_texto.replace("\n", "<br/>"), estilo_corpo))
    
    doc.build(elementos)
    buffer.seek(0)
    return responses.StreamingResponse(buffer, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=orcamento_sinapi.pdf"})
