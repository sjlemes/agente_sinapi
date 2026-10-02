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

# --- 1. SEU ROBÔ NACIONAL DE INICIALIZAÇÃO (OLAHNDO NOME DAS COLUNAS) ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Inicializando banco nacional unificado por cabeçalhos...")
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS composicoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, estado TEXT, codigo TEXT,
            descricao TEXT, unidade TEXT, preco_unitario REAL, regime TEXT
        )
    """)
    conn.commit()

    nome_arquivo_local = "sinapi_2026_08.xlsx"

    try:
        if os.path.exists(nome_arquivo_local):
            print(f"ROBÔ SINAPI: Indexando arquivo {nome_arquivo_local}...")
            excel_file = pd.ExcelFile(nome_arquivo_local, engine="openpyxl")
            cursor.execute("DELETE FROM composicoes")
            
            for nome_aba in excel_file.sheet_names:
                nome_aba_upper = nome_aba.upper().strip()
                if nome_aba_upper in ["MENU", "BUSCA", "LEIAME", "INSTRUÇÕES"]:
                    continue
                
                regime_aba = "DESONERADO" if "CD" in nome_aba_upper or "IC" in nome_aba_upper else "NÃO DESONERADO"
                print(f"ROBÔ SINAPI: Mapeando cabeçalhos analíticos da aba [{nome_aba}]...")
                
                # Lê a planilha mantendo a linha 4 como o cabeçalho real (nomes das colunas)
                df = pd.read_excel(excel_file, sheet_name=nome_aba, skiprows=4, engine="openpyxl")
                
                # Padroniza o nome de todas as colunas para letras maiúsculas e sem espaços (Evita erros de digitação da CEF)
                df.columns = [str(c).upper().strip() for c in df.columns]
                
                # Procura de forma dinâmica em qual coluna o Código e o Preço estão escondidos
                col_codigo = next((c for c in df.columns if "CODIGO" in c or "CÓDIGO" in c), None)
                col_descricao = next((c for c in df.columns if "DESCRICAO" in c or "DESCRIÇÃO" in c), None)
                col_unidade = next((c for c in df.columns if "UNIDADE" in c), None)
                col_preco = next((c for c in df.columns if "PRECO" in c or "PREÇO" in c or "CUSTO" in c or "VALOR" in c), None)
                col_estado = next((c for c in df.columns if "UF" in c or "ESTADO" in c), None)

                # Se o robô localizou a estrutura mínima de colunas, varre as linhas por nome
                if col_codigo and col_descricao:
                    print(f"ROBÔ SINAPI: Colunas identificadas -> Código: [{col_codigo}], Preço: [{col_preco}]")
                    
                    for _, linha in df.iterrows():
                        # Captura e limpa o código independente da coluna onde ele esteja
                        val_codigo = str(linha[col_codigo]).strip() if pd.notna(linha[col_codigo]) else ""
                        if val_codigo.endswith(".0"):
                            val_codigo = val_codigo[:-2]
                        
                        # Se encontrou o código numérico válido, salva a linha com os dados mapeados
                        if val_codigo.isdigit():
                            estado_registro = str(linha[col_estado]).strip().upper() if col_estado and pd.notna(linha[col_estado]) else "NACIONAL"
                            desc_val = str(linha[col_descricao]).upper().strip() if pd.notna(linha[col_descricao]) else ""
                            unid_val = str(linha[col_unidade]).upper().strip() if col_unidade and pd.notna(linha[col_unidade]) else "-"
                            
                            preco = 0.0
                            if col_preco and pd.notna(linha[col_preco]):
                                try:
                                    preco = float(linha[col_preco])
                                except:
                                    preco = 0.0
                                    
                            cursor.execute("""
                                INSERT INTO composicoes (estado, codigo, descricao, unidade, preco_unitario, regime)
                                VALUES (?, ?, ?, ?, ?, ?)
                            """, (estado_registro, val_codigo, desc_val, unid_val, preco, regime_aba))
            
            conn.commit()
            print("ROBÔ SINAPI: Base Nacional SQLite populada com sucesso usando Mapeamento por Nomes!")
        else:
            raise FileNotFoundError()
    except Exception as e:
        print(f"ROBÔ SINAPI: Erro ao indexar arquivo: {str(e)}")
    finally:
        conn.close()
    yield

app = FastAPI(lifespan=lifespan)

GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

# --- ROTA 1 DE PRODUÇÃO: PROTEGIDA CONTRA INSTABILIDADES DO GOOGLE ---
# --- 2. SUA ROTA DE CONSULTA DA IA COM BUSCA AMPLA POR APROXIMAÇÃO ---
# --- 2. SUA ROTA DE CONSULTA DA IA ATUALIZADA (BUSCA AMPLA GLOBAL) ---
@app.post("/calcular-orcamento")
def calcular_orcamento(request: OrcamentoRequest):
    termo_limpo = request.palavras_chave.strip()
    termo_busca_like = f"%{termo_limpo.upper()}%"
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # Busca focada puramente na identificação do código ou da palavra-chave textual
    cursor.execute("""
        SELECT codigo, descricao, unidade, preco_unitario 
        FROM composicoes 
        WHERE (codigo = ? OR descricao LIKE ?) AND estado = ? AND regime = ?
        LIMIT 25
    """, (termo_limpo, termo_busca_like, request.estado.upper().strip(), "DESONERADO" if request.desonerado else "NÃO DESONERADO"))
    
    linhas_banco = cursor.fetchall()
    conn.close()
    
    dados_estruturados_reais = []
    if linhas_banco:
        for linha in linhas_banco:
            if len(linha) >= 4:
                dados_estruturados_reais.append(
                    f"Código: {str(linha).strip()} | Descrição: {str(linha).strip().upper()} | Unidade: {str(linha).strip().upper()} | Preço Unitário: R$ {float(linha):.2f}"
                )
    
    contexto_banco_real = "\n".join(dados_estruturados_reais) if dados_estruturados_reais else "Nenhum item correspondente localizado no SINAPI."
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para: "{request.palavras_chave}" (Quantidade: {request.quantidade}) no estado: {request.estado.upper()} ({regime_texto}).
    
    Dados reais extraídos diretamente das planilhas oficiais indexadas:
    {contexto_banco_real}
    
    Com base estritamente nessas referências reais fornecidas acima, monte o orçamento. Use o preço unitário e descrição encontrados no banco.
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
