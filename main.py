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
# --- SUBSTITUA APENAS O BLOCO DO LIFESPAN NO SEU MAIN.PY ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Inicializando Banco Relacional SQLite...")
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
            print(f"ROBÔ SINAPI: Localizado arquivo {nome_arquivo_local}. Iniciando indexação rígida...")
            cursor.execute("DELETE FROM composicoes")
            cursor.execute("DELETE FROM insumos")
            cursor.execute("DELETE FROM analitico")
            
            # Abrimos o arquivo inspecionando as abas
            excel_file = pd.ExcelFile(nome_arquivo_local, engine="openpyxl")
            
            for nome_aba in excel_file.sheet_names:
                nome_aba_upper = nome_aba.upper().strip()
                if nome_aba_upper in ["MENU", "BUSCA", "LEIAME", "INSTRUÇÕES"]:
                    continue

                regime_aba = "DESONERADO" if "CD" in nome_aba_upper or "IC" in nome_aba_upper else "NÃO DESONERADO"
                
                # 🧱 1. PROCESSAMENTO DE COMPOSIÇÕES E INSUMOS (CSD, CCD, ISD, ICD)
                if nome_aba_upper in ["CSD", "CCD", "ISD", "ICD"]:
                    print(f"ROBÔ SINAPI: Lendo e desmesclando aba de preços [{nome_aba}]...")
                    
                    # 🚨 AJUSTE 1: Lemos o Excel ativando o data_only direto no read_excel para abrir os valores das fórmulas
                    df = pd.read_excel(
                        nome_arquivo_local, 
                        sheet_name=nome_aba, 
                        header=[8, 9], # Junta as linhas 9 e 10 mescladas do Excel
                        engine="openpyxl",
                        engine_kwargs={"data_only": True, "read_only": True} # Envelopado corretamente aqui!
                    )

                    
                    # Achata o cabeçalho duplo em nomes lineares
                    df.columns = [f"{str(c[0]).strip().upper()}_{str(c[1]).strip().upper()}" for c in df.columns]
                    
                    # Forçamos a leitura pelas posições físicas exatas para o Pandas não se perder nas colunas
                    for _, linha in df.iterrows():
                        if len(linha) >= 8:
                            # Layout Nacional Unificado: Coluna 0 (A) = Estado | Coluna 2 (C) = Código SINAPI | Coluna 3 (D) = Descrição | Coluna 4 (E) = Unidade
                            estado_registro = str(linha.iloc[0]).strip().upper() if pd.notna(linha.iloc[0]) else "NACIONAL"
                            
                            # Limpa e isola o código da composição/insumo
                            cod_bruto = str(linha.iloc[2]).strip().split(".")[0] if pd.notna(linha.iloc[2]) else ""
                            val_codigo = "".join(filter(str.isdigit, cod_bruto))
                            
                            if val_codigo and len(val_codigo) >= 4:
                                desc_val = str(linha.iloc[3]).upper().strip()
                                unid_val = str(linha.iloc[4]).upper().strip()
                                
                                # Se for aba de composição, alimenta o índice de serviços globais
                                if "CS" in nome_aba_upper or "CC" in nome_aba_upper:
                                    cursor.execute("""
                                        INSERT OR IGNORE INTO composicoes (codigo, descricao, unidade)
                                        VALUES (?, ?, ?)
                                    """, (val_codigo, desc_val, unid_val))
                                
                                # Varre as colunas mapeando os custos associados a cada estado
                                for col_nome in df.columns:
                                    if "CUSTO" in col_nome or "PREÇO" in col_nome or "PRECO" in col_nome:
                                        partes = col_nome.split("_")
                                        if len(partes) >= 2:
                                            estado_sigla = partes[0].upper().strip()
                                            
                                            preco = 0.0
                                            try:
                                                if pd.notna(linha[col_nome]):
                                                    preco = float(linha[col_nome])
                                            except:
                                                preco = 0.0
                                                
                                            cursor.execute("""
                                                INSERT INTO insumos (estado, codigo, descricao, unidade, preco_unitario, regime)
                                                VALUES (?, ?, ?, ?, ?, ?)
                                            """, (estado_sigla, val_codigo, desc_val, unid_val, preco, regime_aba))

                # 🧱 2. PROCESSAMENTO DA ABA ANALÍTICO (COLUNAS B, D, E, F, G DA LINHA 10)
                elif nome_aba_upper in ["ANALÍTICO", "ANALITICO"]:
                    print("ROBÔ SINAPI: Mapeando estrutura analítica interna...")
                    df_ana = pd.read_excel(nome_arquivo_local, sheet_name=nome_aba, skiprows=9, engine="openpyxl")
                    
                    for _, linha in df_ana.iterrows():
                        if len(linha) >= 7:
                            cod_comp_bruto = str(linha.iloc[1]).strip().split(".")[0] if pd.notna(linha.iloc[1]) else ""
                            cod_comp = "".join(filter(str.isdigit, cod_comp_bruto))
                            
                            cod_ins_bruto = str(linha.iloc[3]).strip().split(".")[0] if pd.notna(linha.iloc[3]) else ""
                            cod_ins = "".join(filter(str.isdigit, cod_ins_bruto))
                            
                            if cod_comp and cod_ins:
                                desc_ins = str(linha.iloc[4]).upper().strip() if pd.notna(linha.iloc[4]) else ""
                                unid_ins = str(linha.iloc[5]).upper().strip() if pd.notna(linha.iloc[5]) else "-"
                                
                                coef = 0.0
                                try:
                                    if pd.notna(linha.iloc[6]):
                                        coef = float(linha.iloc[6])
                                except:
                                    coef = 0.0
                                    
                                cursor.execute("""
                                    INSERT INTO analitico (codigo_composicao, codigo_insumo, descricao_insumo, unidade_insumo, coeficiente)
                                    VALUES (?, ?, ?, ?, ?)
                                """, (cod_comp, cod_ins, desc_ins, unid_ins, coef))
            
            conn.commit()
            print("ROBÔ SINAPI: Base Relacional Nacional SQLite populada com sucesso!")
            
            # --- SEU RAIO-X DE VALIDAÇÃO DO LOG ---
            print("----------------------------------------------------------------")
            print("🔎 ROBÔ SINAPI: INICIANDO RAIO-X AUDITORIA DO BANCO DE DADOS...")
            cursor.execute("SELECT COUNT(*) FROM composicoes")
            print(f"📊 TOTAL DE COMPOSIÇÕES INDEXADAS: {cursor.fetchone()[0]} linhas.")
            cursor.execute("SELECT COUNT(*) FROM insumos")
            print(f"📊 TOTAL DE INSUMOS INDEXADOS: {cursor.fetchone()[0]} linhas.")
            
            cursor.execute("SELECT estado, codigo, descricao, preco_unitario FROM insumos WHERE codigo != '0' LIMIT 5")
            amostras = cursor.fetchall()
            print("📋 MOSTRANDO AS 5 PRIMEIRAS LINHAS DA TABELA DE INSUMOS:")
            for index, am in enumerate(amostras):
                print(f"   Amostra {index+1} -> Estado: [{am[0]}] | Código: [{am[1]}] | Preço: [R$ {am[3]:.2f}] | Descrição: {am[2][:40]}...")
            print("----------------------------------------------------------------")
            
        else:
            raise FileNotFoundError()
    except Exception as e:
        print(f"⚠️ [ERRO CRÍTICO NO LIFESPAN]: Falha no mapeamento: {str(e)}")
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
