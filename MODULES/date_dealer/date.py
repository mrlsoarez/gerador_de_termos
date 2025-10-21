import datetime

def pegar_data_hoje_ptbr(marcador):
    date_en_us = datetime.date.today()
    return date_en_us.strftime(rf"%d{marcador}%m{marcador}%y")