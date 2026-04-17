
import glob
import os
from datetime import datetime
import pathlib
import shutil

class AutomacaoArquivos:
    """Classe para automatizar buscas e movimentação de arquivos"""
    
    def __init__(self):
        """Inicializa a classe com arquivo armazenado como None"""
        self.arquivo_armazenado = None
        self.caminho_arquivo = None
    
    def buscar_ultimo_arquivo(self, nome_arquivo, diretorio_busca='*'):
        """
        Busca o arquivo mais recente com o nome especificado no sistema,
        incluindo variantes com numeração como (1), (2), etc.
        
        Exemplo: Busca por 'planilha.xlsx' encontra:
          - planilha.xlsx
          - planilha (1).xlsx
          - planilha (2).xlsx
        
        Args:
            nome_arquivo (str): Nome do arquivo a buscar (ex: "relatorio.xlsx")
            diretorio_busca (str): Diretório onde buscar (padrão: '*' = todo o sistema)
        
        Returns:
            str: Caminho do arquivo encontrado ou None se não encontrar
        """
        try:
            # Separa nome base e extensão
            nome_base, extensao = os.path.splitext(nome_arquivo)
            
            # Cria padrão para buscar o arquivo e suas variantes
            # Busca: "nome_arquivo", "nome_arquivo (1)", "nome_arquivo (2)", etc
            padrao_arquivo = f"{nome_base}*{extensao}"
            
            # Se diretorio_busca for '*', busca em todo o sistema a partir do C:
            if diretorio_busca == '*':
                padrao_busca = f"C:\\**\\{padrao_arquivo}"
            else:
                padrao_busca = f"{diretorio_busca}\\**\\{padrao_arquivo}"
            
            # Encontra todos os arquivos que correspondem ao padrão
            arquivos_encontrados = glob.glob(padrao_busca, recursive=True)
            
            if not arquivos_encontrados:
                print(f"❌ Nenhum arquivo '{nome_arquivo}' (ou variantes) encontrado")
                return None
            
            # Encontra o arquivo mais recente pela data de modificação
            arquivo_recente = max(arquivos_encontrados, key=os.path.getmtime)
            
            # Armazena o arquivo
            self.caminho_arquivo = arquivo_recente
            self.arquivo_armazenado = os.path.basename(arquivo_recente)
            
            print(f"✅ Arquivo encontrado: {self.caminho_arquivo}")
            print(f"   Data de modificação: {datetime.fromtimestamp(os.path.getmtime(arquivo_recente)).strftime('%d/%m/%Y %H:%M:%S')}")
            
            # Se encontrou variante, avisa qual foi
            if os.path.basename(arquivo_recente) != nome_arquivo:
                print(f"   Arquivo com variante: {os.path.basename(arquivo_recente)}")
            
            return arquivo_recente
        
        except Exception as e:
            print(f"❌ Erro ao buscar arquivo: {e}")
            return None
    
    def obter_arquivo_armazenado(self):
        """
        Retorna os dados do arquivo armazenado
        
        Returns:
            dict: Dicionário com informações do arquivo armazenado
        """
        if not self.arquivo_armazenado:
            print("❌ Nenhum arquivo armazenado ainda")
            return None
        
        return {
            'nome': self.arquivo_armazenado,
            'caminho_completo': self.caminho_arquivo,
            'tamanho_mb': round(os.path.getsize(self.caminho_arquivo) / (1024 * 1024), 2)
        }
    
    def mover_para_planilhas(self, diretorio_destino=None):
        """
        Move o arquivo armazenado para a pasta 'planilhas'
        
        Args:
            diretorio_destino (str): Caminho da pasta planilhas (padrão: ./planilhas)
        
        Returns:
            bool: True se movido com sucesso, False caso contrário
        """
        if not self.caminho_arquivo:
            print("❌ Nenhum arquivo armazenado para mover")
            return False
        
        try:
            # Define o diretório de destino
            if diretorio_destino is None:
                # Usa a pasta 'planilhas' no diretório current do projeto
                diretorio_destino = pathlib.Path(__file__).parent.parent / 'planilhas'
            else:
                diretorio_destino = pathlib.Path(diretorio_destino)
            
            # Cria a pasta se não existir
            diretorio_destino.mkdir(parents=True, exist_ok=True)
            
            # Define o caminho de destino do arquivo
            arquivo_destino = diretorio_destino / self.arquivo_armazenado
            
            # Move o arquivo
            shutil.move(self.caminho_arquivo, str(arquivo_destino))
            
            print(f"✅ Arquivo movido com sucesso para: {arquivo_destino}")
            
            # Atualiza o caminho armazenado
            self.caminho_arquivo = str(arquivo_destino)
            
            return True
        
        except Exception as e:
            print(f"❌ Erro ao mover arquivo: {e}")
            return False
    
    def limpar_armazenamento(self):
        """Limpa o arquivo armazenado"""
        self.arquivo_armazenado = None
        self.caminho_arquivo = None
        print("✅ Armazenamento limpo")

