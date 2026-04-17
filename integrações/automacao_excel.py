import pandas as pd

class FrequenciasDf():
    def __init__(self, caminho_db, nome_coluna_nomes, nome_coluna_curso, nome_coluna_frequencia):
        self.nome_coluna_nomes = nome_coluna_nomes
        self.nome_coluna_curso = nome_coluna_curso
        self.nome_coluna_frequencia = nome_coluna_frequencia
        self.caminho_db = caminho_db
        self.df = None
        self.df_geradora = None

    def ler_base(self):
        try:
            dataframe_pwd = f'{self.caminho_db}.xlsx'
            self.df = pd.read_excel(dataframe_pwd)
        except Exception as error:
            print (error)


    def iterar_base(self):
        try:
            self.df_geradora = self.df.iterrows()
            return self.df_geradora
        except Exception as error:
            print (error)

    def aplicar_frequencia(self, indice, frequencia):
        try:
            self.df.at[indice, self.nome_coluna_frequencia] = frequencia
        except Exception as error:
            print (error)

    def salvar_planilha(self):
        try:
            nome_planilha = str(input('Nome da Planilha: '))
            self.df.to_excel(nome_planilha, index=False)
        except Exception as error:
            print (error)

class RelatorioGeral():
    def __init__(self, caminho_db, coluna_nomes, coluna_status, coluna_atraso, coluna_abono, coluna_curso=None):
        self.caminho_db = caminho_db
        self.coluna_nomes = coluna_nomes
        self.coluna_status = coluna_status
        self.coluna_atraso = coluna_atraso
        self.coluna_abono = coluna_abono
        self.coluna_curso = coluna_curso
        self.df = None
        self.df_geradora = None

    def ler_base(self):
        try:
            dataframe_pwd = f'{self.caminho_db}'
            self.df = pd.read_csv(dataframe_pwd, encoding='latin_1', sep=';')
        except Exception as error:
            print (error)

    def iterar_base(self):
        try:
            self.df_geradora = self.df.iterrows()
            return self.df_geradora
        except Exception as error:
            print (error)

class CriarPlanilhaFrequencias():
    def __init__(self, nome_arq, coluna_nomes, coluna_curso, coluna_frequencia):
        self.nome_arq = nome_arq
        self.coluna_nomes = coluna_nomes
        self.coluna_curso = coluna_curso
        self.coluna_frequencia = coluna_frequencia
        self.df = None
        self.indice_linha_atual = 0
    
    def criar_planilha(self):
        """
        Cria uma planilha vazia com as colunas especificadas (tipo texto)
        """
        try:
            # Cria um DataFrame com as colunas vazias do tipo texto
            self.df = pd.DataFrame({
                self.coluna_nomes: pd.Series(dtype='str'),
                self.coluna_curso: pd.Series(dtype='str'),
                self.coluna_frequencia: pd.Series(dtype='str')
            })
            
            print(f"✅ Planilha criada com as colunas: {self.coluna_nomes}, {self.coluna_curso}, {self.coluna_frequencia}")
            return self.df
        
        except Exception as error:
            print(f"❌ Erro ao criar planilha: {error}")
            return None
    
    def adicionar_valor(self, valor, coluna, indice=None):
        """
        Adiciona um valor a uma coluna específica
        
        Args:
            valor: O valor a ser adicionado
            coluna: Nome da coluna ('nome', 'curso' ou 'frequencia')
            indice: Índice da linha (se None, usa indice_linha_atual)
        
        Returns:
            bool: True se adicionado com sucesso, False caso contrário
        """
        try:
            if self.df is None:
                print("❌ Planilha não foi criada. Execute criar_planilha() primeiro")
                return False
            
            # Mapeia nomes simplificados para os nomes das colunas
            mapa_colunas = {
                'nome': self.coluna_nomes,
                'curso': self.coluna_curso,
                'frequencia': self.coluna_frequencia
            }
            
            # Valida se a coluna existe
            if coluna not in mapa_colunas:
                print(f"❌ Coluna '{coluna}' não existe. Use: 'nome', 'curso' ou 'frequencia'")
                return False
            
            coluna_real = mapa_colunas[coluna]
            
            # Se for adicionar na coluna 'nome', cria uma nova linha
            if coluna == 'nome':
                indice = self.indice_linha_atual
                # Cria a linha se ela não existir
                if indice >= len(self.df):
                    self.df.loc[indice] = [None, None, None]
            else:
                # Para outras colunas, usa o índice atual
                indice = self.indice_linha_atual
            
            # Adiciona o valor
            self.df.at[indice, coluna_real] = valor
            
            print(f"✅ Valor '{valor}' adicionado à coluna '{coluna}' (linha {indice})")
            
            # Incrementa o índice ao adicionar frequência (última coluna do registro)
            if coluna == 'frequencia':
                self.indice_linha_atual += 1
            
            return True
        
        except Exception as error:
            print(f"❌ Erro ao adicionar valor: {error}")
            return False
    
    def salvar_planilha(self):
        """
        Salva a planilha em arquivo Excel usando o nome_arq
        """
        nome = input("nome da planilha: ")
        try:
            if self.df is None:
                print("❌ Planilha não foi criada ou está vazia")
                return False
            
            # Se o nome não tiver extensão, adiciona .xlsx
            nome = nome if nome.endswith('.xlsx') else f"{nome}.xlsx"
            
            self.df.to_excel(nome, index=False)
            print(f"✅ Planilha salva como: {nome}")
            return True
        
        except Exception as error:
            print(f"❌ Erro ao salvar planilha: {error}")
            return False
    
    def exibir_planilha(self):
        """
        Exibe a planilha atual
        """
        if self.df is None:
            print("❌ Planilha não foi criada")
            return
        
        print(self.df)