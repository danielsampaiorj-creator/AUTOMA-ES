import os
import traceback
import unicodedata
import re
from dotenv import load_dotenv
from integrações.automacao_navegador import ConectRh
from core.path import AutomacaoArquivos
from integrações.automacao_excel import FrequenciasDf, RelatorioGeral, CriarPlanilhaFrequencias
from core.data import subtrair_dias_uteis_hoje

def normalizar_nome(nome: str) -> str:
    # Remove caracteres de controle e espaços unicode
    nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
    # Substitui qualquer sequência de whitespace por um espaço simples
    nome = re.sub(r'\s+', ' ', nome)
    # Maiúsculo e strip
    return nome.strip().upper()





data = subtrair_dias_uteis_hoje(2, formato='br')
load_dotenv()
email = os.getenv('EMAIL')
senha = os.getenv('SENHA')


def main():
    '''Automação do ConectRh'''
    CIEDS_CONECT = ConectRh()
    CIEDS_CONECT.abrir_navegador()
    CIEDS_CONECT.logar(email, senha)
    try:
        CIEDS_CONECT.pagina_faltas()
        CIEDS_CONECT.baixar_faltas(data=data)
        CIEDS_CONECT.fechar_popup()
        arquivo_planilha_faltas = AutomacaoArquivos()
        arquivo_planilha_faltas = arquivo_planilha_faltas.buscar_ultimo_arquivo(
            nome_arquivo='Relatorio de Faltas e Atrasos.csv',
            diretorio_busca=r'C:\Users\danielsampaio.rj\Desktop\Pasta de Teste de Scripts Python\cieds_bot_v02\planilhas'
            )
        
        planilha_faltas = RelatorioGeral(
            caminho_db=arquivo_planilha_faltas,
            coluna_nomes='Nome',
            coluna_status='Status',
            coluna_atraso='Atraso',
            coluna_abono='AbonoFalta'
            )
        planilha_faltas.ler_base()
        planilha_faltas.iterar_base()

        CIEDS_CONECT.pagina_inicial()

    except Exception as error:
        print (f'Erro crítico: {error} - Verifique o processo manualmente ')

    try: 
        CIEDS_CONECT.pagina_contratos()
        CIEDS_CONECT.baixar_contratos()
        CIEDS_CONECT.fechar_popup()

        arquivo_planilha_contratos = AutomacaoArquivos()
        arquivo_planilha_contratos.buscar_ultimo_arquivo(
            nome_arquivo='Contratos.csv',
            diretorio_busca=r'C:\Users\danielsampaio.rj\Desktop\Pasta de Teste de Scripts Python\cieds_bot_v02'
            )
        
        planilha_contratos = RelatorioGeral(
            caminho_db=arquivo_planilha_contratos.caminho_arquivo,
            coluna_nomes='CandidatoNome',
            coluna_curso='CursoAprendizagem',
            coluna_status=None,
            coluna_atraso=None,
            coluna_abono=None,   
        )
        planilha_contratos.ler_base()
        planilha_contratos.iterar_base()
        CIEDS_CONECT.pagina_inicial()
    except Exception as error:
        print (f'Erro crítico: {error} - Verifique o processo manualmente ')

    
    try:
        data2 = subtrair_dias_uteis_hoje(2)
        path = r'planilhas\Frequencias'
        planilha_frequencias = CriarPlanilhaFrequencias(
            nome_arq = f'{path} - {data2}.xlsx',
            coluna_nomes = 'Nome',
            coluna_curso = 'Curso',
            coluna_frequencia = 'Frequência'
        )
        planilha_frequencias.criar_planilha()
        CIEDS_CONECT.pagina_relatorios()
        for idx_faltas, linha_faltas in planilha_faltas.df.iterrows():
                nome_aluno = str(linha_faltas[planilha_faltas.coluna_nomes]).strip()
                status = str(linha_faltas[planilha_faltas.coluna_status]).strip()
                atraso = str(linha_faltas[planilha_faltas.coluna_atraso]).strip()
                abono = str(linha_faltas[planilha_faltas.coluna_abono]).strip()
    
                if status == 'Ativo' and atraso == '00:00' and abono == 'Não':
                    aluno_encontrado = False
                    for idx_contratos, linha_contratos in planilha_contratos.df.iterrows():
                        nome_contrato = str(linha_contratos[planilha_contratos.coluna_nomes]).strip()
                        curso_contrato = str(linha_contratos[planilha_contratos.coluna_curso]).strip()
                        # if normalizar_nome(nome_aluno) == normalizar_nome(nome_contrato):
                        if (normalizar_nome(nome_aluno) in normalizar_nome(nome_contrato) or
                        normalizar_nome(nome_contrato) in normalizar_nome(nome_aluno)):
                            aluno_encontrado = True
                            valor_frequencia = 'Não encontrado'
                            
                            try:
                                CIEDS_CONECT.frequencia_aprendiz(
                                    nome=nome_contrato,
                                    curso=curso_contrato
                                )
                                tentativas = 0
                                verificacao_sucesso = False
                                while tentativas < 3 and not verificacao_sucesso:
                                    try:
                                        CIEDS_CONECT.pop_up_frequencias()
                                        valor_frequencia = CIEDS_CONECT.pegar_valor_frequencia()
                                        if valor_frequencia is not None:
                                            verificacao_sucesso = True
                                    except Exception:
                                        tentativas += 1
                                        try:
                                            CIEDS_CONECT.fechar_popup()
                                        except:
                                            pass
                                        if tentativas >= 3:
                                            valor_frequencia = 'Não encontrado'

                                try:
                                    CIEDS_CONECT.fechar_popup()
                                except:
                                    pass
                            except Exception as erro:
                                print(f'⚠️ Erro ao processar {nome_contrato}: {erro}')
                                valor_frequencia = 'Não encontrado'
                                try:
                                    CIEDS_CONECT.fechar_popup()
                                except:
                                    pass
                            
                            # Adiciona sempre a linha com o valor (encontrado ou não)
                            planilha_frequencias.adicionar_valor(valor=nome_contrato, coluna='nome')
                            planilha_frequencias.adicionar_valor(valor=curso_contrato, coluna='curso')
                            planilha_frequencias.adicionar_valor(valor=valor_frequencia, coluna='frequencia')
                            
                            if valor_frequencia != 'Não encontrado':
                                print(f'✅ {nome_contrato} - {curso_contrato} - {valor_frequencia}')
                            else:
                                print(f'❌ {nome_contrato} - {curso_contrato} - {valor_frequencia}')
                            
                            break
                    
                    if not aluno_encontrado:
                        print(f'⚠️ {nome_aluno} não encontrado em Contratos')
        planilha_frequencias.salvar_planilha()
    except Exception as error:
        print (f'Erro crítico: {error} - Verifique o processo manualmente ')
        print(f"❌ Tipo de erro: {type(error).__name__}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
