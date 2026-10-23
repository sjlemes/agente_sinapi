import os
import io
import json
import sqlite3
import pandas as pd
import re
from contextlib import asynccontextmanager
from fastapi import FastAPI, responses, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from google import genai
from google.genai import types
from typing import List, Optional
from openpyxl.utils import get_column_letter

# Ferramentas para gerar planilha Excel
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Inicialização do client com a nova SDK (busca automaticamente GEMINI_API_KEY no ambiente)
client = genai.Client()

# Definição do modelo de entrada vindo do Ktor/Android
class OrcamentoRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool  # True para Desonerado, False para Não Desonerado

import re
import pandas as pd
import sqlite3
import os
from contextlib import asynccontextmanager

def extrair_codigo_limpo(celula_valor) -> str:
    """Extrai o código oficial limpo de uma célula do Excel,
    removendo fórmulas de HIPERLINK e padronizando estritamente para dígitos.
    """
    if celula_valor is None:
        return ""
    
    val_str = str(celula_valor).strip()
    
    # Se a célula contém uma fórmula de hiperlink (ex: =HIPERLINK("...", "12345"))
    if val_str.startswith("="):
        match_hiperlink = re.search(r'HIPERLINK\s*\([^,;]+[,;]\s*["\']?([^"\']+)["\']?\)', val_str, re.IGNORECASE)
        if match_hiperlink:
            val_str = match_hiperlink.group(1).strip()
        else:
            # Fallback: pega o último trecho entre aspas da fórmula
            textos_aspas = re.findall(r'"([^"]+)"', val_str)
            if textos_aspas:
                val_str = textos_aspas[-1].strip()

    # Remove ponto zero caso tenha vindo como float (ex: "12345.0")
    if val_str.endswith(".0"):
        val_str = val_str[:-2]

    # Isola estritamente os dígitos do código oficial da Caixa
    val_codigo = "".join(filter(str.isdigit, val_str))
    return val_codigo


@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Inicializando Banco Relacional SQLite definitivo com Regex...")
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
            print(f"ROBÔ SINAPI: Localizado arquivo {nome_arquivo_local}. Iniciando varredura limpa...")
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
                    print(f"ROBÔ SINAPI: Devorando aba [{nome_aba}] com limpeza de hiperlinks...")
                    
                    df = pd.read_excel(
                        nome_arquivo_local, 
                        sheet_name=nome_aba, 
                        skiprows=9, 
                        engine="openpyxl",
                        engine_kwargs={"data_only": True, "read_only": True}
                    )
                    
                    for _, linha in df.iterrows():
                        if len(linha) >= 4:
                            desc_bruta = str(linha.values[2]).strip().upper() if pd.notna(linha.values[2]) else ""
                            
                            if desc_bruta and desc_bruta not in ["DESCRIÇÃO", "DESCRIÇÃO DA COMPOSIÇÃO", "NONE", "NAN"]:
                                unid_val = str(linha.values[3]).upper().strip() if pd.notna(linha.values[3]) else "-"
                                
                                # 🚨 APLICANDO A FUNÇÃO DE LIMPEZA ROBUSTA DE CÓDIGOS
                                celula_codigo = linha.values[1]
                                val_codigo = extrair_codigo_limpo(celula_codigo)
                                
                                if not val_codigo or len(val_codigo) < 3:
                                    val_codigo = str(abs(hash(desc_bruta)))[:6]

                                if "CS" in nome_aba_upper or "CC" in nome_aba_upper:
                                    cursor.execute("""
                                        INSERT OR IGNORE INTO composicoes (codigo, descricao, unidade)
                                        VALUES (?, ?, ?)
                                    """, (val_codigo, desc_bruta, unid_val))
                                
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

                # 🧱 2. PROCESSAMENTO INTELIGENTE DA ABA ANALÍTICO COM LIMPEZA DE CÓDIGOS
                elif nome_aba_upper in ["ANALÍTICO", "ANALITICO"]:
                    print("ROBÔ SINAPI: Mapeando estrutura analítica interna com limpeza de códigos...")
                    df_ana = pd.read_excel(nome_arquivo_local, sheet_name=nome_aba, skiprows=9, engine="openpyxl")
                    
                    df_ana.columns = [str(c).replace("\n", " ").replace("\r", " ").strip().upper() for c in df_ana.columns]
                    
                    col_comp = next((c for c in df_ana.columns if "COMPOSIÇÃO" in c or "COMPOSICAO" in c), None)
                    col_item = next((c for c in df_ana.columns if "ITEM" in c or "INSUMO" in c or "CÓDIGO" in c), None)
                    col_desc_item = next((c for c in df_ana.columns if "DESCRIÇÃO" in c or "DESCRICAO" in c), None)
                    col_unid_item = next((c for c in df_ana.columns if "UNIDADE" in c), None)
                    col_coef = next((c for c in df_ana.columns if "COEFICIENTE" in c or "QUANTIDADE" in c), None)
                    
                    if col_comp and col_item:
                        for _, linha in df_ana.iterrows():
                            # 🚨 APLICANDO A FUNÇÃO DE LIMPEZA NA ABA ANALÍTICA
                            val_comp = extrair_codigo_limpo(linha[col_comp])
                            val_item = extrair_codigo_limpo(linha[col_item])
                            
                            if val_comp and val_item:
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
            print("ROBÔ SINAPI: Base Relacional Nacional populada com códigos oficiais limpos!")

        else:
            raise FileNotFoundError(f"Arquivo {nome_arquivo_local} não encontrado.")
    except Exception as e:
        print(f"⚠️ [ERRO CRÍTICO NO LIFESPAN]: {str(e)}")
    finally:
        conn.close()
    yield

app = FastAPI(lifespan=lifespan)

GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

# --- ROTA 1 DE PRODUÇÃO: PRINCIPAL (VERSÃO HIERÁRQUICA EM ÁRVORE) ---
@app.post("/calcular-orcamento")
async def calcular_orcamento(dados: OrcamentoRequest):
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # 1. Tratamento seguro das palavras-chave para busca flexível por múltiplos termos
    termos = dados.palavras_chave.strip().split()
    if not termos:
        conn.close()
        raise HTTPException(status_code=400, detail="Informe ao menos uma palavra-chave para a busca.")
    
    # 💡 CORREÇÃO CRASH: Usamos OR em vez de AND para permitir buscas com 2 ou mais palavras sem zerar o resultado
    condicoes = []
    parametros = []
    for termo in termos:
        condicoes.append("c.descricao LIKE ?")
        parametros.append(f"%{termo}%")
    
    where_clause = " OR ".join(condicoes)
    
    # Conversão elegante do booleano recebido para o padrão do banco
    regime_filtro = "DESONERADO" if dados.desonerado else "NÃO DESONERADO"
    
    # 2. Consulta SQL Cruzada: Composições + Preços Regionais/Estado + Insumos da Tabela Analítica
    query = f"""
        SELECT 
            c.codigo AS codigo_comp, 
            c.descricao AS desc_comp, 
            c.unidade AS unid_comp,
            i.preco_unitario,
            i.regime,
            a.codigo_insumo,
            a.descricao_insumo,
            a.unidade_insumo,
            a.coeficiente
        FROM composicoes c
        JOIN insumos i ON c.codigo = i.codigo
        LEFT JOIN analitico a ON c.codigo = a.codigo_composicao
        WHERE ({where_clause}) 
          AND i.regime = ?
        LIMIT 30
    """
    
    parametros_completos = parametros + [regime_filtro]
    cursor.execute(query, parametros_completos)
    linhas = cursor.fetchall()
    conn.close()
    
    if not linhas:
        return {
            "relatorio": "⚠️ Nenhuma composição encontrada para os termos informados no banco de dados.",
            "composicoes": []
        }

    # 3. Organização e empacotamento hierárquico dos dados reais para o Gemini e Compose
    composicoes_processadas = {}
    codigos_encontrados = []
    
    for row in linhas:
        cod_comp, desc_comp, unid_comp, preco_unit, regime = row[0], row[1], row[2], row[3], row[4]
        
        if cod_comp not in composicoes_processadas:
            preco_base = preco_unit or 0.0
            total_comp = preco_base * dados.quantidade
            
            composicoes_processadas[cod_comp] = {
                "codigo": str(cod_comp),
                "descricao": desc_comp,
                "unidade": unid_comp,
                "total": f"R$ {total_comp:,.2f}",
                "insumos": []
            }
            codigos_encontrados.append(str(cod_comp))

    # 🔍 BUSCA COMPLEMENTAR SEGURA DOS INSUMOS FILHOS NA TABELA ANALITICO
    if codigos_encontrados:
        conn_ana = sqlite3.connect("sinapi.db")
        cursor_ana = conn_ana.cursor()
        
        placeholders = ",".join(["?"] * len(codigos_encontrados))
        query_ana = f"""
            SELECT codigo_composicao, codigo_insumo, descricao_insumo, unidade_insumo, coeficiente
            FROM analitico
            WHERE codigo_composicao IN ({placeholders})
        """
        cursor_ana.execute(query_ana, codigos_encontrados)
        linhas_analiticas = cursor_ana.fetchall()
        conn_ana.close()
        
        for ana_row in linhas_analiticas:
            c_comp, c_ins, d_ins, u_ins, coef = ana_row
            if c_comp in composicoes_processadas:
                coef_val = coef or 0.0
                qtd_insumo_total = coef_val * dados.quantidade
                composicoes_processadas[c_comp]["insumos"].append({
                    "nome": f"[{c_ins}] {d_ins}" if d_ins else f"[{c_ins}] Insumo Vinculado",
                    "qtd": qtd_insumo_total,
                    "unidade": u_ins or "-",
                    "total": f"Coef: {coef_val:.4f}"
                })

    dados_estruturados = list(composicoes_processadas.values())

    print("----------------------------------------------------------------")
    print("🔎 ROBÔ SINAPI: RAIO-X HIERÁRQUICO ENVIADO PARA IA...")
    print(f"Total de composições mapeadas: {len(dados_estruturados)}")
    for item in dados_estruturados:
        print(f"   -> Comp: {item['codigo']} | Insumos filhos encontrados: {len(item['insumos'])}")
    print("----------------------------------------------------------------")

    # 4. Geração do Parecer Técnico com a nova SDK do Gemini (`client.models.generate_content`)
    prompt_ia = f"""
    Você é um Engenheiro Civil Sênior especialista em orçamentação de obras públicas e tabelas SINAPI.
    Abaixo estão os dados reais estruturados extraídos diretamente do SQLite para a busca '{dados.palavras_chave}' (Regime: {regime_filtro}):

    {dados_estruturados}

    Quantidade desejada na obra: {dados.quantidade}
    
    Com base estritamente nestes dados factuais, elabore um parecer analítico descritivo, técnico e contextualizado sobre a viabilidade, composição de custos e observações de execução para este cenário. Retorne exatamente o JSON exigido no schema, preenchendo o relatório descritivo e repassando ou ajustando as composições de forma correta.
    """

    # 🛠️ ESQUEMA RECONFIGURADO EM ÁRVORE HIERÁRQUICA PARA O JETPACK COMPOSE
    config_ia = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema={
            "type": "OBJECT",
            "properties": {
                "relatorio": {"type": "STRING"},
                "composicoes": {
                    "type": "ARRAY",
                    "items": {
                        "type": "OBJECT",
                        "properties": {
                            "codigo": {"type": "STRING"},
                            "descricao": {"type": "STRING"},
                            "unidade": {"type": "STRING"},
                            "total": {"type": "STRING"},
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
                        "required": ["codigo", "descricao", "unidade", "total", "insumos"]
                    }
                }
            },
            "required": ["relatorio", "composicoes"]
        }
    )    

    try:
        resposta = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt_ia,
            config=config_ia
        )
        dados_resposta = json.loads(resposta.text)
        return {
            "relatorio": dados_resposta.get("relatorio", "Orçamento calculado com sucesso."),
            "composicoes": dados_resposta.get("composicoes", [])
        }
    except Exception as e:
        print(f"ROBÔ SINAPI: Instabilidade na API de IA: {str(e)}")
        return {
            "relatorio": "⚠️ O servidor gratuito do Google está temporariamente congestionado devido à alta demanda global neste minuto.\n\nA nossa infraestrutura local está 100% operacional. Por favor, clique no botão novamente em alguns segundos para reprocessar a consulta.",
            "composicoes": [
                {
                    "codigo": "ERRO",
                    "descricao": "Servidor do Google Ocupado",
                    "unidade": "Erro",
                    "total": "Tente de Novo",
                    "insumos": []
                }
            ]
        }