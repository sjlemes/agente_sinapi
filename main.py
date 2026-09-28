import os
from fastapi import FastAPI
from pydantic import BaseModel
from google import genai

# 1. Configura o novo cliente do Gemini
GOOGLE_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GOOGLE_API_KEY)


app = FastAPI()

# Modelo de dados que o App Android vai enviar
class OrcamentoRequest(BaseModel):
    estado: str
    palavras_chave: str
    quantidade: float
    desonerado: bool    

@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    regime = "DESONERADO" if request.desonerado else "NÃO DESONERADO"
    
    # Nova instrução cirúrgica para o Gemini agir como banco de dados analítico e engenheiro
    prompt = f"""
    Você é um Engenheiro de Custos sênior especialista em auditoria de orçamentos e na base de dados oficial do SINAPI da Caixa Econômica Federal.
    O usuário precisa de um orçamento para o estado: {request.estado} sob o regime de encargos: {regime}.
    Ele buscou pelo serviço: "{request.palavras_chave}" para executar uma quantidade de: {request.quantidade}.
    
    Com base no seu conhecimento atualizado das referências analíticas oficiais do SINAPI, faça:
    1. Identifique o código numérico oficial do SINAPI da composição mais adequada.
    2. Apresente os dados da COMPOSIÇÃO PRINCIPAL: Código, Descrição Completa, Unidade de medida e Valor Unitário.
    3. Apresente a memória de cálculo do VALOR TOTAL (Valor Unitário x {request.quantidade}).
    4. Liste detalhadamente a composição analítica por dentro (todos os INSUMOS e MÃO DE OBRA utilizados):
       - Nome do Insumo/Profissional (Ex: Cimento, Servente, Pedreiro)
       - Quantidade calculada proporcional para atender {request.quantidade} unidades do serviço.
       - Unidade de medida do insumo (KG, H, M³, etc).
       - Preço unitário e Preço Total daquele insumo na estrutura.
       
    Formate o relatório final em Markdown de forma muito visual, utilizando tabelas limpas para separar a Composição dos Insumos internos.
    """
    
    resposta = client.models.generate_content(
        model='gemini-3.8-flash',
        contents=prompt,
    )
    
    return {"relatorio": resposta.text}

