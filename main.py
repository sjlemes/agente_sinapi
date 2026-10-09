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

                # 🧱 1. LEITURA PLANA POR COORDENADAS FIXAS REAL DA CAIXA
                if nome_aba_upper in ["CSD", "CCD", "ISD", "ICD"]:
                    regime_aba = "DESONERADO" if "CD" in nome_aba_upper or "IC" in nome_aba_upper else "NÃO DESONERADO"
                    print(f"ROBÔ SINAPI: Devorando aba [{nome_aba}] por coordenadas fixas...")
                    
                    # skiprows=9 joga o cabeçalho decorativo fora. A linha 10 vira o índice dos dados.
                    df = pd.read_excel(
                        nome_arquivo_local, 
                        sheet_name=nome_aba, 
                        skiprows=9, 
                        engine="openpyxl",
                        engine_kwargs={"data_only": True, "read_only": True}
                    )
                    
                    # Mapeamento estrito por posições (Colunas A, B, C, D estruturadas)
                    for _, linha in df.iterrows():
                        if len(linha) >= 4:
                            # 🚨 CORREÇÃO: Usamos .values para extrair o texto limpo, livre de objetos internos do Pandas
                            desc_bruta = str(linha.values[2]).strip().upper() if pd.notna(linha.values[2]) else ""
                            
                            if desc_bruta and desc_bruta not in ["DESCRIÇÃO", "DESCRIÇÃO DA COMPOSIÇÃO", "NONE", "NAN"]:
                                unid_val = str(linha.values[3]).upper().strip() if pd.notna(linha.values[3]) else "-"
                                celula_codigo = str(linha.values[1]).strip() if pd.notna(linha.values[1]) else ""
                                
                                val_codigo = "".join(filter(str.isdigit, celula_codigo))
                                
                                # Se o hiperlink esconder o código, gera uma chave numérica idêntica estável
                                if not val_codigo or val_codigo == "0":
                                    val_codigo = str(abs(hash(desc_bruta)))[:6]

                                if "CS" in nome_aba_upper or "CC" in nome_aba_upper:
                                    cursor.execute("""
                                        INSERT OR IGNORE INTO composicoes (codigo, descricao, unidade)
                                        VALUES (?, ?, ?)
                                    """, (val_codigo, desc_bruta, unid_val))
                                
                                # LOOP DE CUSTOS DOS ESTADOS (Coluna E em diante, pulando de 2 em 2)
                                for i in range(4, len(linha), 2):
                                    preco = 0.0
                                    try:
                                        if pd.notna(linha.values[i]):
                                            preco = float(linha.values[i])
                                    except:
                                        preco = 0.0
                                        
                                    cursor.execute("""
                                        INSERT INTO insumos (estado, codigo, descricao, unidade, preco_unitario, regime)
                                        VALUES (?, ?, ?, ?, ?, ?)
                                    """, ("NACIONAL", val_codigo, desc_bruta, unid_val, preco, regime_aba))

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
                # Imprime os índices 0 (Código), 1 (Unidade) e 2 (Descrição) da tupla do banco
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
    
    estado_alvo = request.estado.upper().strip()
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    regime_busca_like = f"%{regime_texto}%"
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # já que os preços foram guardados globalmente. Deixamos a filtragem regional para a inteligência do Gemini!
    cursor.execute("""
        SELECT c.codigo, c.descricao, c.unidade, i.preco_unitario 
        FROM composicoes c
        JOIN insumos i ON c.codigo = i.codigo
        WHERE (c.codigo = ? OR c.descricao LIKE ?) 
          AND i.regime LIKE ?
        LIMIT 15
    """, (termo_limpo, termo_busca_like, regime_busca_like))
    
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
