import os
import io
import json
import sqlite3
import pandas as pd
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

# --- MODELOS DE DADOS ---
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

# --- 1. SEU ROBÔ NACIONAL DE INICIALIZAÇÃO (PREENCHE O BANCO) ---
# --- 1. SEU ROBÔ NACIONAL DE INICIALIZAÇÃO (COM TRATAMENTO DE TEXTO BRUTO) ---
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
                    continue
                
                if len(nome_aba_upper) >= 2:
                    estado_aba = nome_aba_upper[:2]
                    regime_aba = "DESONERADO" if "DES" in nome_aba_upper and "NDES" not in nome_aba_upper else "NÃO DESONERADO"
                    
                    print(f"ROBÔ SINAPI: Lendo a aba [{nome_aba}]...")
                    df = pd.read_excel(excel_file, sheet_name=nome_aba, skiprows=4)
                    
                    for _, linha in df.iterrows():
                        # TRATAMENTO DE DADOS COLETADOS DO EXCEL
                        if len(linha) > 0:
                            # Tenta localizar o código removendo decimais extras (.0) comuns do Pandas
                            val_codigo = str(linha.iloc).strip() if pd.notna(linha.iloc) else ""
                            if val_codigo.endswith(".0"):
                                val_codigo = val_codigo[:-2]

                            # Só insere se achou um número de código válido
                            if val_codigo.isdigit():
                                desc_val = str(linha.iloc).upper().strip() if len(linha) > 1 and pd.notna(linha.iloc) else ""
                                unid_val = str(linha.iloc).upper().strip() if len(linha) > 2 and pd.notna(linha.iloc) else "-"
                                
                                preco = 0.0
                                if len(linha) > 7 and pd.notna(linha.iloc):
                                    try:
                                        preco = float(linha.iloc)
                                    except:
                                        preco = 0.0
                                        
                                cursor.execute("""
                                    INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """, (estado_aba, val_codigo, desc_val, unid_val, preco, regime_aba))
            conn.commit()
            print("ROBÔ SINAPI: Base Nacional SQLite populada com sucesso!")
        else:
            raise FileNotFoundError()
    except Exception as e:
        print(f"ROBÔ SINAPI: Fallback de segurança ativo: {str(e)}")
        cursor.execute("DELETE FROM composicoes")
        cursor.execute("INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime) VALUES ('RJ', '87528', 'ALVENARIA DE VEDAÇÃO DE BLOCO CERÂMICO FURADO', 'M²', 45.50, 'NÃO DESONERADO')")
        conn.commit()
    finally:
        conn.close()
    yield

app = FastAPI(lifespan=lifespan)

GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

# --- ROTA 1 DE PRODUÇÃO: PROTEGIDA CONTRA INSTABILIDADES DO GOOGLE ---
# --- 2. SUA ROTA DE CONSULTA DA IA COM BUSCA AMPLA POR APROXIMAÇÃO ---
@app.post("/calcular-orcamento")
def calcular_orcamento(request: OrcamentoRequest):
    # Trata a entrada removendo espaços e forçando maiúsculas
    termo_limpo = request.palavras_chave.strip()
    termo_busca_like = f"%{termo_limpo.upper()}%"
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # BUSCA BLINDADA: Tenta achar correspondência exata do código OU aproximação do texto
    cursor.execute("""
        SELECT codigo, descricao, unidade, preco_unitario 
        FROM composicoes 
        WHERE codigo = ? OR descricao LIKE ?
        LIMIT 15
    """, (termo_limpo, termo_busca_like))
    
    linhas_banco = cursor.fetchall()
    conn.close()
    
    dados_estruturados_reais = []
    if lines_banco := linhas_banco:
        for linha in lines_banco:
            if len(linha) >= 4:
                dados_estruturados_reais.append(
                    f"Código: {str(linha)} | Descrição: {str(linha)} | Unidade: {str(linha)} | Preço Unitário: R$ {float(linha):.2f}"
                )
    
    contexto_banco_real = "\n".join(dados_estruturados_reais) if dados_estruturados_reais else "Nenhum item correspondente localizado no SINAPI."
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa.
    O usuário quer um orçamento para: "{request.palavras_chave}" (Quantidade: {request.quantidade}) no estado: {request.estado.upper()} ({regime_texto}).
    Dados reais extraídos do SQLite:
    {contexto_banco_real}
    """

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
    
    try:
        resposta = client.models.generate_content(model='gemini-3.8-flash', contents=prompt, config=config_ia)
        dados_resposta = json.loads(resposta.text)
        return {
            "relatorio": dados_resposta.get("relatorio", "Orçamento calculado."),
            "insumos": dados_resposta.get("insumos", [])
        }
    except Exception as e:
        return {
            "relatorio": "⚠️ O servidor gratuito do Google está temporariamente congestionado. Tente novamente em alguns segundos.",
            "insumos": [{"nome": "Servidor Ocupado", "qtd": 0.0, "unidade": "Erro", "total": "Repita"}]
        }

# --- 3. SUA ROTA DE GERAÇÃO DO PDF CORPORATIVO ---
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
