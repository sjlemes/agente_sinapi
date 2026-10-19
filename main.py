import os
import io
import json
import sqlite3
import pandas as pd
from contextlib import asynccontextmanager
from fastapi import FastAPI, responses, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from google import genai
from google.genai import types
from typing import List, Optional

# Ferramentas para gerar planilha Excel
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Definição do modelo de entrada vindo do Ktor/Android
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

            # --- CHECANDO SE PLANILHA FOI DEPURADA E SQLITE POPULADO---
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

# --- ROTA 1 DE PRODUÇÃO: PRINCIPAL (VERSÃO HIERÁRQUICA EM ÁRVORE) ---
@app.post("/calcular-orcamento")
def calcular_orcamento(request: OrcamentoRequest):
    termo_bruto = request.palavras_chave.strip()
    
    # 🚨 DECLARAÇÃO DAS VARIÁVEIS CRUCIAIS DE ESCOPO (ZERA O NAMEERROR)
    estado_alvo = request.estado.upper().strip()
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    regime_busca_like = f"%{regime_texto}%"
    
    # Quebra o texto digitado pelo celular em palavras limpas
    palavras = [p.upper().strip() for p in termo_bruto.split() if p.strip()]
    
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    # Montagem da busca usando Listas estáveis [] (SUA LÓGICA ORIGINAL PRESERVADA)
    if termo_bruto.isdigit() or not palavras:
        sql = """
            SELECT c.codigo, c.descricao, c.unidade, i.preco_unitario 
            FROM composicoes c
            JOIN insumos i ON c.codigo = i.codigo
            WHERE (c.codigo = ? OR c.descricao LIKE ?) AND i.regime LIKE ?
            GROUP BY c.codigo LIMIT 15
        """
        parametros = [termo_bruto, f"%{termo_bruto}%", regime_busca_like]
    else:
        # Monta um filtro 'LIKE' dinâmico para cada palavra isolada
        condicoes_likes = " AND ".join(["c.descricao LIKE ?" for _ in palavras])
        
        sql = f"""
            SELECT c.codigo, c.descricao, c.unidade, i.preco_unitario 
            FROM composicoes c
            JOIN insumos i ON c.codigo = i.codigo
            WHERE ({condicoes_likes}) AND i.regime LIKE ?
            GROUP BY c.codigo LIMIT 15
        """
        parametros = [f"%{p}%" for p in palavras]
        parametros.append(regime_busca_like)
        
    cursor.execute(sql, parametros)
    linhas_banco = cursor.fetchall()
    conn.close()
    
    # Mapeia coluna por coluna explicitamente para estruturar os dados para a IA
    dados_estruturados_reais = []
    if linhas_banco:
        for static_linha in linhas_banco:
            if len(static_linha) >= 4:
                codigo_val = str(static_linha[0]).strip()
                desc_val = str(static_linha[1]).strip().upper()
                unid_val = str(static_linha[2]).strip().upper()
                preco_val = float(static_linha[3]) if static_linha[3] is not None else 0.0
                
                dados_estruturados_reais.append(
                    f"Código SINAPI: {codigo_val} | Descrição: {desc_val} | Unidade: {unid_val} | Preço de Referência: R$ {preco_val:.2f}"
                )
    
    contexto_banco_real = "\n".join(dados_estruturados_reais) if dados_estruturados_reais else "Nenhum item correspondente exato localizado nas tabelas oficiais do SINAPI."
    
    prompt = f"""
    Você é um Engenheiro de Custos especialista na tabela SINAPI da Caixa Econômica Federal.
    O usuário quer um orçamento para o serviço: "{request.palavras_chave}" na quantidade de: {request.quantidade}.
    O estado de interesse dele é: {estado_alvo} sob o regime de encargos: {regime_texto}.
    
    Aqui estão os dados reais extraídos diretamente das tabelas oficiais da planilha nacional unificada:
    {contexto_banco_real}
    
    Com base estritamente nessas referências, monte o orçamento estruturado em composições e seus respectivos insumos e materiais. Use o código, descrição e preço unitário encontrados acima.
    """

    # 🛠️ NOVO ESQUEMA RECONFIGURADO EM ÁRVORE HIERÁRQUICA PARA O JETPACK COMPOSE
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
            contents=prompt,
            config=config_ia
        )
        dados_resposta = json.loads(resposta.text)
        return {
            "relatorio": dados_resposta.get("relatorio", "Orçamento calculado com sucesso."),
            "composicoes": dados_resposta.get("composicoes", [])
        }
    except Exception as e:
        print(f"ROBÔ SINAPI: Instabilidade na API de IA: {str(e)}")
        # Tratamento de erro original preservado e adaptado para a nova estrutura de árvore
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

# --- NOVA ROTA 2: EXPORTAÇÃO EXCEL DA COMPOSIÇÃO CONSULTADA ---
@app.post("/exportar-excel")
def exportar_excel(request: OrcamentoRequest):
    termo_bruto = request.palavras_chave.strip()
    estado_alvo = request.estado.upper().strip()
    regime_texto = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    regime_busca_like = f"%{regime_texto}%"
    palavras = [p.upper().strip() for p in termo_bruto.split() if p.strip()]
    
    # 1. EXECUTA A MESMA BUSCA SQL QUE VOCÊ JÁ HOMOLOGOU
    conn = sqlite3.connect("sinapi.db")
    cursor = conn.cursor()
    
    if termo_bruto.isdigit() or not palavras:
        sql = """
            SELECT c.codigo, c.descricao, c.unidade, i.preco_unitario 
            FROM composicoes c
            JOIN insumos i ON c.codigo = i.codigo
            WHERE (c.codigo = ? OR c.descricao LIKE ?) AND i.regime LIKE ?
            GROUP BY c.codigo LIMIT 15
        """
        parametros = [termo_bruto, f"%{termo_bruto}%", regime_busca_like]
    else:
        condicoes_likes = " AND ".join(["c.descricao LIKE ?" for _ in palavras])
        sql = f"""
            SELECT c.codigo, c.descricao, c.unidade, i.preco_unitario 
            FROM composicoes c
            JOIN insumos i ON c.codigo = i.codigo
            WHERE ({condicoes_likes}) AND i.regime LIKE ?
            GROUP BY c.codigo LIMIT 15
        """
        parametros = [f"%{p}%" for p in palavras]
        parametros.append(regime_busca_like)
        
    cursor.execute(sql, parametros)
    linhas_banco = cursor.fetchall()
    
    if not linhas_banco:
        conn.close()
        raise HTTPException(status_code=404, detail="Nenhum dado encontrado para exportar.")

    # 2. SE DETECTOU DADOS, BUSCA OS FILHOS (ANALÍTICO) DA PRIMEIRA COMPOSIÇÃO ENCONTRADA
    # Para o Excel solto, vamos detalhar analiticamente a primeira composição que casou com a busca
    primeira_comp = linhas_banco[0]
    codigo_comp = str(primeira_comp[0]).strip()
    desc_comp = str(primeira_comp[1]).strip().upper()
    unid_comp = str(primeira_comp[2]).strip().upper()
    
    query_insumos = """
        SELECT i.descricao, a.coeficiente, i.unidade, i.preco_unitario
        FROM analitico a
        JOIN insumos i ON a.codigo_insumo = i.codigo
        WHERE a.codigo_composicao = ? AND i.regime LIKE ?
    """
    cursor.execute(query_insumos, (codigo_comp, regime_busca_like))
    linhas_insumos = cursor.fetchall()
    conn.close()

    # 3. CONSTRUÇÃO DA PLANILHA EXCEL INTELIGENTE COM OPENPYXL
    wb = Workbook()
    ws = wb.active
    ws.title = "Orçamento SINAPI"
    
    # Ativa as linhas de grade para visualização no Excel
    ws.views.sheetView[0].showGridLines = True
    
    # Estilizações Elegantes (Padrão Engenharia)
    fonte_titulo = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    fonte_cabecalho = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    fonte_normal = Font(name="Arial", size=10)
    fonte_bold = Font(name="Arial", size=10, bold=True)
    
    fill_azul_escuro = PatternFill(start_color="0D47A1", end_color="0D47A1", fill_type="solid")
    fill_cinza_claro = PatternFill(start_color="F1F3F5", end_color="F1F3F5", fill_type="solid")
    
    borda_fina = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    # Título do Relatório
    ws.merge_cells("A1:E1")
    ws["A1"] = f"RELATÓRIO DE CUSTOS SINAPI - {estado_alvo} ({regime_texto})"
    ws["A1"].font = fonte_titulo
    ws["A1"].fill = fill_azul_escuro
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # Informações da Composição Mãe
    ws["A3"] = "Composição Principal:"
    ws["A3"].font = fonte_bold
    ws["B3"] = f"{codigo_comp} - {desc_comp}"
    ws["B3"].font = fonte_normal
    
    ws["A4"] = "Quantidade Pedida:"
    ws["A4"].font = fonte_bold
    ws["B4"] = request.quantidade
    ws["B4"].font = fonte_normal
    ws["C4"] = unid_comp
    ws["C4"].font = fonte_normal

    # Cabeçalho da Tabela de Materiais
    cabecalhos = ["Cód/Descrição Insumo", "Unidade", "Coeficiente", "Preço Unitário (R$)", "Total Calculado (R$)"]
    for col_num, text in enumerate(cabecalhos, 1):
        cell = ws.cell(row=6, column=col_num)
        cell.value = text
        cell.font = fonte_cabecalho
        cell.fill = fill_azul_escuro
        cell.alignment = Alignment(horizontal="center" if col_num in [2,3] else "left" if col_num==1 else "right")
    ws.row_dimensions[6].height = 25

    # Corpo da Tabela (Insumos Filhos)
    linha_atual = 7
    soma_total_geral = 0.0
    
    for idx, ins in enumerate(linhas_insumos, 1):
        desc_ins = str(ins[0]).strip().upper()
        coef = float(ins[1]) if ins[1] else 0.0
        unid_ins = str(ins[2]).strip().upper()
        preco_u = float(ins[3]) if ins[3] else 0.0
        
        qtd_total_insumo = request.quantidade * coef
        custo_total_insumo = qtd_total_insumo * preco_u
        soma_total_geral += custo_total_insumo

        # Inserção das células
        c1 = ws.cell(row=linha_atual, column=1, value=desc_ins)
        c2 = ws.cell(row=linha_atual, column=2, value=unid_ins)
        c3 = ws.cell(row=linha_atual, column=3, value=coef)
        c4 = ws.cell(row=linha_atual, column=4, value=preco_u)
        c5 = ws.cell(row=linha_atual, column=5, value=custo_total_insumo)

        # Formatações Numéricas e Alinhamentos
        c1.alignment = Alignment(horizontal="left")
        c2.alignment = Alignment(horizontal="center")
        c3.alignment = Alignment(horizontal="right")
        c4.alignment = Alignment(horizontal="right")
        c5.alignment = Alignment(horizontal="right")
        
        # Máscara de Moeda do Excel nativa (facilita somas posteriores pelo usuário)
        c4.number_format = '"R$"#,##0.00'
        c5.number_format = '"R$"#,##0.00'
        c3.number_format = '#,##0.000'

        # Aplica fontes e bordas
        for c in [c1, c2, c3, c4, c5]:
            c.font = fonte_normal
            c.border = borda_fina
            if idx % 2 == 0:
                c.fill = fill_cinza_claro # Efeito zebra na planilha!
                
        linha_atual += 1

    # Linha de Total Geral da Composição
    ws.cell(row=linha_atual, column=4, value="TOTAL DA ETAPA:").font = fonte_bold
    ws.cell(row=linha_atual, column=4).alignment = Alignment(horizontal="right")
    
    total_cell = ws.cell(row=linha_atual, column=5, value=soma_total_geral)
    total_cell.font = fonte_bold
    total_cell.number_format = '"R$"#,##0.00'
    total_cell.alignment = Alignment(horizontal="right")

    # Ajuste Automático da Largura das Colunas para não cortar o texto no Excel
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # 4. SALVA O ARQUIVO EM MEMÓRIA E ENVIA VIA STREAM
    conteudo_excel = io.BytesIO()
    wb.save(conteudo_excel)
    conteudo_excel.seek(0)
    
    nome_arquivo = f"Orcamento_{codigo_comp}_{estado_alvo}.xlsx"
    return StreamingResponse(
        conteudo_excel,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nome_arquivo}"}
    )