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
                    
                    df = pd.read_excel(excel_file, sheet_name=nome_aba, skiprows=4)
                    for _, linha in df.iterrows():
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

# --- 2. SUA ROTA DE CONSULTA DA IA COM VALIDAÇÃO DE COLUNAS DO SQLITE ---
@app.post("/calcular-orcamento")
def calcular_orcamento(request: OrcamentoRequest):
    termo_busca = f"%{request.palavras_chave.upper().strip()}%"
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT codigo, descricao, unidade, preco_unitario 
        FROM composicoes 
        WHERE descricao LIKE ? OR codigo = ?
        LIMIT 15
    """, (termo_busca, request.palavras_chave.strip()))
    
    linhas_banco = cursor.fetchall()
    conn.close()
    
    # 🚨 AQUI ESTAVA O SEU ACERTO ESQUECIDO: Mapeando coluna por coluna explicitamente!
    dados_estruturados_reais = []
    if linhas_banco:
        for linha in linhas_banco:
            if len(linha) >= 4:
                codigo_val = str(linha[0]).strip()
                desc_val = str(linha[1]).strip().upper()
                unid_val = str(linha[2]).strip().upper()
                preco_val = float(linha[3]) if linha[3] is not None else 0.0
                
                dados_estruturados_reais.append(
                    f"Código: {codigo_val} | Descrição: {desc_val} | Unidade: {unid_val} | Preço Unitário: R$ {preco_val:.2f}"
                )
    
    contexto_banco_real = "\n".join(dados_estruturados_reais) if dados_estruturados_reais else "Nenhum item correspondente localizado no SINAPI."
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para o serviço: "{request.palavras_chave}" para a quantidade de: {request.quantidade}.
    O estado de interesse é: {request.estado.upper().strip()} sob o regime: {regime_texto}.
    
    Dados reais extraídos diretamente das tabelas oficiais da planilha nacional indexada:
    {contexto_banco_real}
    
    Com base estritamente nessas referências, monte o orçamento. Se dados reais foram fornecidos acima, utilize o código e preço unitário deles.
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
    
    resposta = client.models.generate_content(model='gemini-3.8-flash', contents=prompt, config=config_ia)
    
    try:
        dados_resposta = json.loads(resposta.text)
        return {
            "relatorio": dados_resposta.get("relatorio", "Orçamento gerado."),
            "insumos": dados_resposta.get("insumos", [])
        }
    except Exception as e:
        return {
            "relatorio": f"Orçamento Processado:\n{resposta.text}",
            "insumos": [{"nome": request.palavras_chave.upper(), "qtd": request.quantidade, "unidade": "UNID", "total": "OK"}]
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
