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

# --- CARREGAMENTO RELACIONAL, MULTIINDEX E GUARDIÃO DE LAYOUT (LIFESPAN) ---
# --- COPIE E SUBSTITUA APENAS O BLOCO DO LIFESPAN NO SEU MAIN.PY ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Inicializando Banco Relacional SQLite definitivo...")
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS composicoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT UNIQUE, descricao TEXT, unidade TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS insumos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, estado TEXT, codigo TEXT, descricao TEXT, unidade TEXT, preco_unitario REAL, regime TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS analitico (
            id INTEGER PRIMARY KEY AUTOINCREMENT, codigo_composicao TEXT, codigo_insumo TEXT, descricao_insumo TEXT, unidade_insumo TEXT, coeficiente REAL
        )
    """)
    conn.commit()

    nome_arquivo_local = "sinapi_2026_08.xlsx"

    try:
        if os.path.exists(nome_arquivo_local):
            print(f"ROBÔ SINAPI: Localizado arquivo {nome_arquivo_local}. Iniciando varredura plana...")
            excel_file = pd.ExcelFile(nome_arquivo_local, engine="openpyxl")
            
            cursor.execute("DELETE FROM composicoes")
            cursor.execute("DELETE FROM insumos")
            cursor.execute("DELETE FROM analitico")
            
            for nome_aba in excel_file.sheet_names:
                nome_aba_upper = nome_aba.upper().strip()
                if nome_aba_upper in ["MENU", "BUSCA", "LEIAME", "INSTRUÇÕES"]:
                    continue

                # 🧱 1. LEITURA PLANA DE PREÇOS (PULA A LINHA 9 DE AVISOS MESCLADOS)
                if nome_aba_upper in ["CSD", "CCD", "ISD", "ICD"]:
                    regime_aba = "DESONERADO" if "CD" in nome_aba_upper or "IC" in nome_aba_upper else "NÃO DESONERADO"
                    print(f"ROBÔ SINAPI: Lendo aba [{nome_aba}] como tabela simples plana...")
                    
                    df = pd.read_excel(
                        nome_arquivo_local, 
                        sheet_name=nome_aba, 
                        skiprows=9, 
                        engine="openpyxl",
                        engine_kwargs={"data_only": True, "read_only": True}
                    )
                    
                    # Limpa quebras de linha e espaços nos nomes das colunas da linha 10
                    df.columns = [str(c).replace("\n", " ").replace("\r", " ").strip().upper() for c in df.columns]
                    
                    # Ajuste fino de filtros de colunas textuais na linha plana
                    col_codigo = next((c for c in df.columns if "CÓDIGO" in c or "CODIGO" in c), None)
                    col_descricao = next((c for c in df.columns if "DESCRIÇÃO" in c or "DESCRICAO" in c), None)
                    col_unidade = next((c for c in df.columns if "UNIDADE" in c or "UNID" in c), None)

                    if col_codigo and col_descricao:
                        for _, linha in df.iterrows():
                            celula_bruta = str(linha[col_codigo]).strip()
                            if celula_bruta.endswith(".0"):
                                celula_bruta = celula_bruta[:-2]
                                
                            val_codigo = "".join(filter(str.isdigit, celula_bruta))
                            
                            if val_codigo and val_codigo.isdigit() and len(val_codigo) >= 4:
                                desc_val = str(linha[col_descricao]).upper().strip() if pd.notna(linha[col_descricao]) else ""
                                unid_val = str(linha[col_unidade]).upper().strip() if col_unidade and pd.notna(linha[col_unidade]) else "-"
                                
                                # Popula a tabela de composições globais para habilitar o JOIN
                                if "CS" in nome_aba_upper or "CC" in nome_aba_upper:
                                    cursor.execute("""
                                        INSERT OR IGNORE INTO composicoes (codigo, descricao, unidade)
                                        VALUES (?, ?, ?)
                                    """, (val_codigo, desc_val, unid_val))
                                
                                # Indexa os preços por coluna associada
                                for col_nome in df.columns:
                                    if any(k in col_nome for k in ["CUSTO", "PRECO", "PREÇO", "VALOR"]):
                                        preco = 0.0
                                        try:
                                            if pd.notna(linha[col_nome]):
                                                preco = float(linha[col_nome])
                                        except:
                                            preco = 0.0
                                        
                                        cursor.execute("""
                                            INSERT INTO insumos (estado, codigo, descricao, unidade, preco_unitario, regime)
                                            VALUES (?, ?, ?, ?, ?, ?)
                                        """, ("NACIONAL", val_codigo, desc_val, unid_val, preco, regime_aba))

                # 🧱 2. PROCESSAMENTO INTELIGENTE DA ABA ANALÍTICO (BASEADO NOS NOMES REAIS DA LINHA 10)
                elif nome_aba_upper in ["ANALÍTICO", "ANALITICO"]:
                    print("ROBÔ SINAPI: Mapeando estrutura analítica interna de forma plana...")
                    df_ana = pd.read_excel(nome_arquivo_local, sheet_name=nome_aba, skiprows=9, engine="openpyxl")
                    
                    # Padroniza as colunas da aba analítica
                    df_ana.columns = [str(c).replace("\n", " ").replace("\r", " ").strip().upper() for c in df_ana.columns]
                    
                    # Detecta dinamicamente as colunas B, D, E, F e G pelos nomes da linha 10
                    col_comp = next((c for c in df_ana.columns if "COMPOSIÇÃO" in c or "COMPOSICAO" in c), None)
                    col_item = next((c for c in df_ana.columns if "ITEM" in c or "INSUMO" in c or "CÓDIGO" in c), None)
                    col_desc_item = next((c for c in df_ana.columns if "DESCRIÇÃO" in c or "DESCRICAO" in c), None)
                    col_unid_item = next((c for c in df_ana.columns if "UNIDADE" in c), None)
                    col_coef = next((c for c in df_ana.columns if "COEFICIENTE" in c or "QUANTIDADE" in c), None)
                    
                    if col_comp and col_item:
                        for _, linha in df_ana.iterrows():
                            # Limpa e isola o código da composição pai
                            raw_comp = str(linha[col_comp]).strip()
                            if raw_comp.endswith(".0"): raw_comp = raw_comp[:-2]
                            val_comp = "".join(filter(str.isdigit, raw_comp))
                            
                            # Limpa e isola o código do insumo filho
                            raw_item = str(linha[col_item]).strip()
                            if raw_item.endswith(".0"): raw_item = raw_item[:-2]
                            val_item = "".join(filter(str.isdigit, raw_item))
                            
                            if val_comp and val_item and len(val_comp) >= 4 and len(val_item) >= 4:
                                desc_ins = str(linha[col_desc_item]).upper().strip() if col_desc_item and pd.notna(linha[col_desc_item]) else ""
                                unid_ins = str(linha[col_unid_item]).upper().strip() if col_unid_item and pd.notna(linha[col_unid_item]) else "-"
                                
                                coef = 0.0
                                try:
                                    if col_coef and pd.notna(linha[col_coef]):
                                        coef = float(linha[col_coef])
                                except:
                                    coef = 0.0
                                    
                                cursor.execute("""
                                    INSERT INTO analitico (codigo_composicao, codigo_insumo, descricao_insumo, unidade_insumo, coeficiente)
                                    VALUES (?, ?, ?, ?, ?)
                                """, (val_comp, val_item, desc_ins, unid_ins, coef))
            
            conn.commit()
            print("ROBÔ SINAPI: Base Relacional Nacional SQLite populada com sucesso de forma plana!")






                       # --- COPIE E ENCAIXE ESTE BLOCO DE RAIO-X REVISADO E LEVE ---
            print("----------------------------------------------------------------")
            print("🔎 ROBÔ SINAPI: INICIANDO RAIO-X AUDITORIA DO BANCO DE DADOS...")
            
            cursor.execute("SELECT COUNT(*) FROM composicoes")
            print(f"📊 TOTAL DE COMPOSIÇÕES INDEXADAS: {cursor.fetchone()[0]} linhas.")
            
            cursor.execute("SELECT COUNT(*) FROM insumos")
            print(f"📊 TOTAL DE INSUMOS INDEXADOS: {cursor.fetchone()[0]} linhas.")
            
            cursor.execute("SELECT COUNT(*) FROM analitico")
            print(f"📊 TOTAL DE LINHAS NO ANALÍTICO: {cursor.fetchone()[0]} linhas.")
            
            print("\n📋 MOSTRANDO AS 3 PRIMEIRAS COMPOSIÇÕES SALVAS:")
            cursor.execute("SELECT codigo, unidade, descricao FROM composicoes LIMIT 3")
            for idx, am in enumerate(cursor.fetchall()):
                print(f"   Comp {idx+1} -> Código: [{am[0]}] | Unidade: [{am[1]}] | Descrição: {am[2][:40]}...")

            print("\n📋 TESTANDO LOG DE CONSULTA REAL (SIMULAÇÃO DE BUSCA POR 'ALVENARIA'):")
            cursor.execute("""
                SELECT c.codigo, i.preco_unitario, c.descricao 
                FROM composicoes c
                JOIN insumos i ON c.codigo = i.codigo
                WHERE c.descricao LIKE '%ALVENARIA%'
                LIMIT 2
            """)
            testes_join = cursor.fetchall()
            if testes_join:
                for idx, tj in enumerate(testes_join):
                    print(f"   Match {idx+1} -> Código: [{tj[0]}] | Preço: [R$ {tj[1]:.2f}] | {tj[2][:40]}...")
            else:
                print("   ⚠️ O CRUZAMENTO (JOIN) NÃO RETORNOU NENHUM RESULTADO PARA 'ALVENARIA'!")
            print("----------------------------------------------------------------")






        else:
            raise FileNotFoundError()
    except Exception as e:
        print(f"⚠️ [ERRO CRÍTICO NO LIFESPAN]: Falha no mapeamento plano: {str(e)}")
    finally:
        conn.close()
    yield

app = FastAPI(lifespan=lifespan)

GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

# --- ROTA 1 DE PRODUÇÃO: PROTEGIDA CONTRA INSTABILIDADES DO GOOGLE ---
@app.post("/calcular-orcamento")
def calcular_orcamento(request: OrcamentoRequest):
    termo_limpo = request.palavras_chave.strip()
    termo_busca_like = f"%{termo_limpo.upper()}%"
    
    # Prepara o estado e o regime para aceitarem buscas aproximadas por texto
    estado_busca = f"%{request.estado.upper().strip()}%"
    regime_busca = f"%{('DESONERADO' if request.desonerado else 'NÃO DESONERADO')}%"
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # 🚨 BUSCA FLEXÍVEL BLINDADA: Troca o '=' rígido por 'LIKE' no Estado e no Regime!
    # Isso garante que se o banco salvou 'RJ ' ou 'CS_RJ', o Python captura mesmo assim!
    cursor.execute("""
        SELECT c.codigo, c.descricao, c.unidade, i.preco_unitario 
        FROM composicoes c
        JOIN insumos i ON c.codigo = i.codigo
        WHERE (c.codigo = ? OR c.descricao LIKE ?) 
          AND i.estado LIKE ? 
          AND i.regime LIKE ?
        LIMIT 15
    """, (termo_limpo, termo_busca_like, estado_busca, regime_busca))
    
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
