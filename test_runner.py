import json
from api.excel_processor import process_excel_files_from_paths
from api.utils import DateTimeEncoder

if __name__ == "__main__":
    # Caminhos para os ficheiros Excel originais
    files_to_process = [
        r'c:\Users\Paulo Martins\Desktop\Candido ADV - Projeto prompt GPT BUILDER\24.09.2025\LEGAL ONE -Relatório Publicações.xlsx',
        r'c:\Users\Paulo Martins\Desktop\Candido ADV - Projeto prompt GPT BUILDER\24.09.2025\ADIVISE -candido-dortas-sociedade-de-advogadas-e-advogados-638943060441979375..xlsx'
    ]
    output_file = 'output.json'

    print("A processar ficheiros para gerar 'output.json'...")

    # Chama a nossa nova função de processamento modular
    final_data = process_excel_files_from_paths(files_to_process)

    # Converte a lista final para JSON e salva no arquivo
    final_json_output = json.dumps(final_data, indent=4, ensure_ascii=False, cls=DateTimeEncoder)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_json_output)
        
    print(f"Arquivo '{output_file}' gerado com sucesso com as novas padronizações.")
