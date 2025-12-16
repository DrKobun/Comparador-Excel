valor = input()
valor = int(valor)


segundos = valor % 60
horas = 0
minutos = 0
if segundos > 60:
    minutos = segundos // 60
    if minutos > 60:
        horas = minutos % 60





print(f"{horas}:{minutos}:{segundos}")