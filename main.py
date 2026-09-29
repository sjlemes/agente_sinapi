import os
import io
from fastapi import FastAPI, responses
from pydantic import BaseModel
from google import genai

# Bibliotecas necessárias para a geração do PDF profissional
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# 1. Configura o cliente do Gemini de forma segura
GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)

app = FastAPI()

# Modelo de dados que o App Android envia
class OrcamentoRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool

# --- ROTA 1: GERA O RELATÓRIO EM TEXTO COM IA ---
@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    regime = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    prompt = f"""
    Você é um Engenheiro de Custos sênior especialista em auditoria de orçamentos e na base de dados oficial do SINAPI da Caixa Econômica Federal.
    O usuário precisa de um orçamento para o estado: {request.estado} sob o regime de encargos: {regime}.
    He buscou pelo serviço: "{request.palavras_chave}" para executar uma quantidade de: {request.quantidade}.
    
    Com base no seu conhecimento atualizado das referências analíticas oficiais do SINAPI, faça:
    1. Identifique o código numérico oficial do SINAPI da composição mais adequada.
    2. Apresente os dados da COMPOSIÇÃO PRINCIPAL: Código, Descrição Completa, Unidade de medida e Valor Unitário.
    3. Apresente a memória de cálculo do VALOR TOTAL (Valor Unitário x {request.quantidade}).
    4. Liste detalhadamente a composição analítica por dentro (todos os INSUMOS e MÃO DE OBRA utilizados):
       - Nome do Insumo/Profissional (Ex: Cimento, Servente, Pedreiro)
       - Quantidade calculada proporcional para atender {request.quantidade} unidades do serviço.
       - Unidade de medida do insumo (KG, H, M³, etc).
       - Preço unitário e Preço Total daquele insumo na estrutura.
       
    Formate o relatório final em tópicos simples e profissionais de forma que o aplicativo consiga exibir de maneira limpa.
    """
    
    resposta = client.models.generate_content(
        model='gemini-3.8-flash',
        contents=prompt,
    )
    
    return {"relatorio": resposta.text}

# --- ROTA 2: GERA O ARQUIVO PDF PARA DOWNLOAD ---
@app.post("/gerar-pdf-orcamento")
async def gerar_pdf_orcamento(request: OrcamentoRequest):
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40,
        title="Relatorio_Orcamento_SINAPI"
    )
    
    styles = getSampleStyleSheet()
    elementos = []
    
    estilo_titulo = ParagraphStyle(
        'TituloPDF',
        parent=styles['Heading1'],
        fontSize=22,
        textColor=colors.HexColor('#1A365D'),
        spaceAfter=15
    )
    
    elementos.append(Paragraph("Relatório de Custos e Composições SINAPI", estilo_titulo))
    elementos.append(Spacer(1, 10))
    
    # Montando uma tabela estruturada no PDF com os dados enviados pelo celular
    regime = "Desonerado" if request.desonerado else "Não Desonerado"
    dados_tabela = [
        ['Parâmetro', 'Valor Selecionado'],
        ['Estado', request.estado],
        ['Busca', request.palavras_chave],
        ['Quantidade', f"{request.quantidade}"],
        ['Regime Mão de Obra', regime]
    ]
    
    # colWidths define a largura das colunas (Ex: 150 pontos para a primeira, 300 para a segunda)
    tabela_visual = Table(dados_tabela, colWidths=[150, 300])
    
    tabela_visual.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F7FAFC')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E0')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#EDF2F7')])
    ]))
    
    elementos.append(tabela_visual)
    
    elementos.append(Spacer(1, 40))
    elementos.append(Paragraph(
        "<font size=8 color='#718096'>Este relatório foi gerado por um agente de IA baseado nas referências públicas do SINAPI. Verifique sempre as fontes oficiais antes de executar a obra.</font>", 
        styles['Italic']
    ))
    
    doc.build(elementos)
    buffer.seek(0)
    
    return responses.StreamingResponse(buffer, media_type="application/pdf", headers={
        "Content-Disposition": "attachment; filename=orcamento_sinapi.pdf"
    })