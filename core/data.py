from datetime import datetime, timedelta

def formatar_data_br(data):
    """Formata data no padrão brasileiro (DD/MM/YYYY)"""
    return data.strftime('%d/%m/%Y')


def subtrair_dias_uteis(data, dias_uteis, formato='datetime'):
    """
    Subtrai dias úteis (pula fins de semana)
    
    Args:
        data: Data inicial
        dias_uteis: Quantos dias úteis subtrair
        formato: 'datetime' retorna objeto datetime, 'br' retorna string DD/MM/YYYY
    """
    dias_subtraidos = 0
    data_atual = data
    
    while dias_subtraidos < dias_uteis:
        data_atual -= timedelta(days=1)
        # 0=segunda, 1=terça, ..., 5=sábado, 6=domingo
        if data_atual.weekday() < 5:  # seg-sex = 0-4
            dias_subtraidos += 1
    
    if formato == 'br':
        return formatar_data_br(data_atual)
    return data_atual


def subtrair_dias_uteis_hoje(dias_uteis, formato='datetime'):
    """
    Subtrai dias úteis da data de hoje
    
    Args:
        dias_uteis: Quantos dias úteis subtrair
        formato: 'datetime' retorna objeto datetime, 'br' retorna string DD/MM/YYYY
    """
    return subtrair_dias_uteis(datetime.now(), dias_uteis, formato)
