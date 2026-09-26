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

# Base de dados simulada estruturada corretamente
dados_sinapi_simulados = [
    {
        "codigo": "87528", 
        "descricao": "ALVENARIA DE VEDAÇÃO DE BLOCO CERÂMICO FURADO 9X19X19CM", 
        "unidade": "M²", 
        "preco_unitario": 45.50
    },
    {
        "codigo": "87529", 
        "descricao": "EMBOÇO OU MASSA ÚNICA PARA RECEBIMENTO DE PINTURA", 
        "unidade": "M²", 
        "preco_unitario": 22.10
    },
    {
        "codigo": "88267", 
        "descricao": "CARPINTEIRO DE FORMAS WITH ENCARGOS COMPLEMENTARES", 
        "unidade": "H", 
        "preco_unitario": 25.00
    }
]

@app.post("/calcular-orcamento")
async def calcular_orcamento(request: OrcamentoRequest):
    contexto_sinapi = str(dados_sinapi_simulados)
    
    prompt = f"""
    Você é um engenheiro de custos especialista em construção civil e na tabela SINAPI.
    O usuário quer um orçamento para o estado {request.estado}.
    Ele buscou por: "{request.palavras_chave}" e quer construir uma quantidade de {request.quantidade}.
    
    Com base APENAS nas opções disponíveis na tabela SINAPI abaixo:
    {contexto_sinapi}
    
    Faça o seguinte:
    1. Identifique qual código melhor se aplica ao pedido dele.
    2. Calcule o valor total (Preço Unitário x Quantidade).
    3. Retorne um relatório curto e profissional formatado em tópicos simples explicando o cálculo.
    """
    
    resposta = client.models.generate_content(
        model='gemini-3.8-flash',
        contents=prompt,
    )
    
    return {"relatorio": resposta.text}

