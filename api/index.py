from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List
import base64
import shutil
import os

# Importa a nossa função de processamento principal
from .excel_processor import process_excel_files_from_paths

# Adiciona a informação do servidor de produção para o esquema OpenAPI
servers = [
    {
        "url": "https://processador-excel-api-candido-adv.vercel.app",
        "description": "Servidor de Produção"
    }
]

# Define o modelo de dados para o pedido
class FilePayload(BaseModel):
    filename: str
    file_base64: str

class ProcessRequest(BaseModel):
    files: List[FilePayload]

app = FastAPI(servers=servers)

@app.get("/api/health")
def health_check():
    """Verifica se a API está no ar."""
    return {"status": "ok"}

@app.post("/api/process")
async def process_excel_files(request: ProcessRequest):
    """
    Recebe uma lista de ficheiros Excel codificados em Base64, processa-os e retorna o JSON unificado.
    """
    if not request.files:
        raise HTTPException(status_code=400, detail="Nenhum ficheiro enviado.")

    temp_dir = "temp_files"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    file_paths = []
    try:
        for file_payload in request.files:
            file_path = os.path.join(temp_dir, file_payload.filename)
            
            # Decodifica a string Base64 de volta para bytes
            try:
                file_bytes = base64.b64decode(file_payload.file_base64)
            except Exception:
                raise HTTPException(status_code=400, detail=f"Erro ao decodificar o ficheiro Base64: {file_payload.filename}")

            # Salva os bytes decodificados num ficheiro temporário
            with open(file_path, "wb") as buffer:
                buffer.write(file_bytes)
            file_paths.append(file_path)

        # Chama a nossa lógica de processamento já existente
        unified_data = process_excel_files_from_paths(file_paths)
        return unified_data

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ocorreu um erro ao processar os ficheiros: {str(e)}")
    finally:
        # Limpa a diretoria temporária
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
