from datetime import datetime, timedelta


def hora_menos_um_minuto():
    now = datetime.now()
    hora_modificada = now - timedelta(minutes=1)
    return hora_modificada.strftime("%H.%M.%S")
