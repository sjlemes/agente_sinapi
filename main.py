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

@asynccontextmanager
async def lifespan(app: FastAPI):
    print("ROBÔ SINAPI: Inicializando Banco Relacional SQLite definitivo (Versão Limpa)...")
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

                # 🧱 1. LEITURA PLANA POR COORDENADAS FIXAS
                if nome_aba_upper in ["CSD", "CCD", "ISD", "ICD"]:
                    regime_aba = "DESONERADO" if "CD" in nome_aba_upper or "IC" in nome_aba_upper else "NÃO DESONERADO"
                    print(f"ROBÔ SINAPI: Devorando aba [{nome_aba}]...")
                    
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
                                
                                # Extração direta e limpa do código oficial da planilha
                                celula_codigo = str(linha.values[1]).strip() if pd.notna(linha.values[1]) else ""
                                if celula_codigo.endswith(".0"):
                                    celula_codigo = celula_codigo[:-2]
                                val_codigo = "".join(filter(str.isdigit, celula_codigo))
                                
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

                # 🧱 2. PROCESSAMENTO DA ABA ANALÍTICO (CORRIGIDO POR ÍNDICES FIXOS)
                elif nome_aba_upper in ["ANALÍTICO", "ANALITICO"]:
                    print("ROBÔ SINAPI: Mapeando estrutura analítica interna...")
                    df_ana = pd.read_excel(
                        nome_arquivo_local, 
                        sheet_name=nome_aba, 
                        skiprows=9, 
                        engine="openpyxl",
                        engine_kwargs={"data_only": True, "read_only": True}
                    )
                    
                    for _, linha in df_ana.iterrows():
                        if len(linha) >= 7:
                            # Coluna de Composição (geralmente índice 1) e Coluna de Item/Insumo (geralmente índice 3)
                            raw_comp = str(linha.values[1]).strip() if pd.notna(linha.values[1]) else ""
                            if raw_comp.endswith(".0"): raw_comp = raw_comp[:-2]
                            val_comp = "".join(filter(str.isdigit, raw_comp))
                            
                            raw_item = str(linha.values[3]).strip() if pd.notna(linha.values[3]) else ""
                            if raw_item.endswith(".0"): raw_item = raw_item[:-2]
                            val_item = "".join(filter(str.isdigit, raw_item))
                            
                            if val_comp and val_item:
                                desc_ins = str(linha.values[4]).upper().strip() if pd.notna(linha.values[4]) else ""
                                unid_ins = str(linha.values[5]).upper().strip() if pd.notna(linha.values[5]) else "-"
                                
                                coef = 0.0
                                try:
                                    if pd.notna(linha.values[6]):
                                        # Trata caso o coeficiente venha como string com vírgula (ex: "1,2790")
                                        val_coef_str = str(linha.values[6]).strip().replace(",", ".")
                                        coef = float(val_coef_str)
                                except:
                                    coef = 0.0
                                    
                                cursor.execute("""
                                    INSERT INTO analitico (codigo_composicao, codigo_insumo, descricao_insumo, unidade_insumo, coeficiente)
                                    VALUES (?, ?, ?, ?, ?)
                                """, (val_comp, val_item, desc_ins, unid_ins, coef))
            conn.commit()
            print("ROBÔ SINAPI: Base Relacional Nacional populada com sucesso (Sem redundâncias)!")

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
    
    palavras = [p.upper().strip() for p in termos if p.strip()]
    
    # Montagem dinâmica dos filtros LIKE
    condicoes = [f"c.descricao LIKE ?" for _ in palavras]
    where_clause = " OR ".join(condicoes)
    
    # 🎯 SISTEMA DE PONTUAÇÃO: Conta quantas palavras da busca a descrição possui
    score_parts = [f"CASE WHEN c.descricao LIKE ? THEN 1 ELSE 0 END" for _ in palavras]
    score_expression = " + ".join(score_parts)
    
    regime_filtro = "DESONERADO" if dados.desonerado else "NÃO DESONERADO"
    
    # 2. Consulta SQL Inteligente: Filtra preços válidos e ordena pelos mais relevantes primeiro
    query = f"""
        SELECT 
            c.codigo AS codigo_comp, 
            c.descricao AS desc_comp, 
            c.unidade AS unid_comp,
            i.preco_unitario,
            i.regime,
            ({score_expression}) AS relevancia
        FROM composicoes c
        JOIN insumos i ON c.codigo = i.codigo
        WHERE ({where_clause}) 
          AND i.regime = ?
          AND i.preco_unitario > 0
        GROUP BY c.codigo
        ORDER BY relevancia DESC, c.descricao ASC
        LIMIT 6
    """
    
    # Organização estrita dos parâmetros para a Query (WHERE + SCORE + REGIME)
    parametros_where = [f"%{p}%" for p in palavras]
    parametros_score = [f"%{p}%" for p in palavras]
    parametros_completos = parametros_where + parametros_score + [regime_filtro]
    
    cursor.execute(query, parametros_completos)
    linhas = cursor.fetchall()
    conn.close()
    
    if not linhas:
        return {
            "relatorio": "⚠️ Nenhuma composição encontrada para os termos informados no banco de dados.",
            "composicoes": []
        }
    
    # 3. Organização limpa das composições encontradas
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

    # 🔍 BUSCA COMPLEMENTAR SEGURA DOS INSUMOS FILHOS NA TABELA ANALITICO PARA TODAS AS COMPOSIÇÕES
    if codigos_encontrados:
        conn_ana = sqlite3.connect("sinapi.db")
        cursor_ana = conn_ana.cursor()
        
        placeholders = ",".join(["?"] * len(codigos_encontrados))
        codigos_str = [str(c).strip() for c in codigos_encontrados]
        
        query_ana = f"""
            SELECT 
                TRIM(CAST(codigo_composicao AS TEXT)) AS c_comp, 
                TRIM(CAST(codigo_insumo AS TEXT)) AS c_ins, 
                descricao_insumo, 
                unidade_insumo, 
                coeficiente
            FROM analitico
            WHERE TRIM(CAST(codigo_composicao AS TEXT)) IN ({placeholders})
              AND codigo_insumo IS NOT NULL 
              AND TRIM(CAST(codigo_insumo AS TEXT)) != ''
        """
        
        cursor_ana.execute(query_ana, codigos_str)
        linhas_analiticas = cursor_ana.fetchall()
        
        print(f"🔎 Total de linhas brutas na analitico: {len(linhas_analiticas)}")
        
        for idx, ana_row in enumerate(linhas_analiticas):
            c_comp, c_ins, d_ins, u_ins, coef = ana_row
            c_comp = str(c_comp).strip()
            c_ins = str(c_ins).strip()
            
            # Verificação de segurança
            if not c_ins or c_ins == "None" or c_ins == c_comp:
                print(f"[{idx}] ⏭️ Ignorado (vazio ou igual ao pai): comp={c_comp}, ins={c_ins}")
                continue
                
            if c_comp in composicoes_processadas:
                # Tratamento seguro do coeficiente (seja string com vírgula ou float)
                if isinstance(coef, str):
                    coef_val = float(coef.replace(',', '.'))
                else:
                    coef_val = float(coef or 0.0)
                
                # Busca o preço unitário
                cursor_ana.execute(
                    "SELECT preco_unitario FROM insumos WHERE TRIM(CAST(codigo AS TEXT)) = ? AND regime = ? LIMIT 1",
                    (c_ins, regime_filtro)
                )
                res_preco = cursor_ana.fetchone()
                preco_val = float(res_preco[0]) if res_preco and res_preco[0] else 0.0
                
                custo_total_componente = coef_val * preco_val * dados.quantidade
                qtd_insumo_total = coef_val * dados.quantidade
                
                composicoes_processadas[c_comp]["insumos"].append({
                    "nome": f"[{c_ins}] {d_ins}" if d_ins else f"[{c_ins}] Insumo Vinculado",
                    "qtd": qtd_insumo_total,
                    "unidade": u_ins or "-",
                    "total": formatar_brl(custo_total_componente)
                })
                print(f"[{idx}] ✅ Sucesso! Insumo {c_ins} adicionado à composição {c_comp}")
            else:
                print(f"[{idx}] ⚠️ A chave '{c_comp}' não foi encontrada em composicoes_processadas.")
                
        conn_ana.close()

    dados_estruturados = list(composicoes_processadas.values())

    print("----------------------------------------------------------------")
    print("🔎 ROBÔ SINAPI: RAIO-X HIERÁRQUICO ENVIADO PARA IA...")
    print(f"Total de composições mapeadas: {len(dados_estruturados)}")
    for item in dados_estruturados:
        print(f"   -> Comp: {item['codigo']} | Insumos filhos encontrados: {len(item['insumos'])}")
    print(dados_estruturados)   
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

    # 4. Geração do Parecer Técnico com tratamento blindado de timeout/instabilidade da IA
    try:
        # Definimos um timeout curto ou deixamos a chamada protegida
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
        print(f"⚠️ ROBÔ SINAPI: IA Indisponível ({str(e)}). Entrando em modo de fallback local instantâneo.")
        
        # 🛡️ FALLBACK INTELIGENTE: Se a IA falhar ou der 503, montamos a estrutura localmente 
        # para garantir que o app receba os dados em menos de 1 segundo, sem timeout!
        composicoes_fallback = []
        for item in dados_estruturados:
            composicoes_fallback.append({
                "codigo": item["codigo"],
                "descricao": item["descricao"],
                "unidade": item["unidade"],
                "total": item["total"],
                "insumos": [
                    {
                        "nome": ins["nome"],
                        "qtd": ins["qtd"],
                        "unidade": ins["unidade"],
                        "total": ins["total"]
                    } for ins in item["insumos"]
                ]
            })

        return {
            "relatorio": "⚠️ O servidor de IA do Google está temporariamente sobrecarregado (Erro 503). Os dados técnicos e custos abaixo foram extraídos diretamente do banco oficial SINAPI com sucesso.",
            "composicoes": composicoes_fallback
        }

# --- FORMATAÇÃO DE QTD E CUSTOS PARA FORMATO BR ---
def formatar_brl(valor: float) -> str:
        """Formata valor para a moeda brasileira (ex: R$ 1.234,56)"""
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_numero_br(valor: float) -> str:
    """Formata quantidades numéricas para o padrão brasileiro (ex: 1.234,56)"""
    return f"{valor:,.4f}".replace(",", "X").replace(".", ",").replace("X", ".").rstrip('0').rstrip(',')