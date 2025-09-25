from fastapi import FastAPI, UploadFile, File, HTTPException
from typing import List, Annotated
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

app = FastAPI(servers=servers)

@app.get("/api/health")
def health_check():
    """Verifica se a API está no ar."""
    return {"status": "ok"}

@app.post("/api/process")
async def process_excel_files(files: Annotated[List[UploadFile], File()]):
    """
    Recebe uma lista de ficheiros Excel, processa-os e retorna o JSON unificado.
    """
    if not files:
        raise HTTPException(status_code=400, detail="Nenhum ficheiro enviado.")

    # Cria uma diretoria temporária para guardar os ficheiros
    temp_dir = "temp_files"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    file_paths = []
    for file in files:
        file_path = os.path.join(temp_dir, file.filename)
        file_paths.append(file_path)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

    try:
        # Chama a função de processamento com a lista de caminhos dos ficheiros guardados
        unified_data = process_excel_files_from_paths(file_paths)
        
        return unified_data
    except Exception as e:
        # Em caso de erro no processamento, retornamos um erro 500
        raise HTTPException(status_code=500, detail=f"Ocorreu um erro ao processar os ficheiros: {str(e)}")
    finally:
        # Limpa a diretoria temporária
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
