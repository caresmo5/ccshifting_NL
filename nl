import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
import random


# INTENTAR:
#   Podría probar a tener en cuenta la productividad media de cada agente (si lo viera viable, por día de la semana y por horas).
#   A lo mejor sería hacer una tabla con las siguientes columnas: País, Día, Hora, Agente, Productividad.
#   Incluso podría sacar las horas cada 10min y, como sé quién está y quien no a ese detalle, puedo sacar los agentes que hay cada 10 min.
#   Como sé los agentes concretos, sé las contestadas concretas cada 10min (productividad por hora / 6). Sumando todo el día, lo tengo.
#   Para esto podría repetir todas las horas cada 10min tantas veces como agentes haya. Luego poner al lado los agentes repetidos.
#   Luego una columna de 1s. Como sé en qué 10min no está cada agente, hago: si es este agente y esta hora (en la que está descansando),
#   pon un 0. Entonces elimino los 0s de la columna. Para cada agente, sustituyo los 1s por su respectiva productividad.
#   OJO, así lo que estoy diciendo es que cada agente está trabajando de 7h a 22h excepto en sus pausas. Tendría que poner 1s solo en las
#   horas en las que trabajan. Esto ya para más adelante.



planning = pd.read_excel('U:\Turnos\Horarios y pausas\Definitivo\Planning Internacional 2 (prueba de Carlos España para Python).xlsx', sheet_name= 'Hoja1')
planning.iloc[14,9:] = planning.iloc[11,9:]
names = planning.iloc[14]
planning = planning.iloc[15:]
planning.columns = names
planning = planning[planning['País'] == 'HOLANDA']
planning = planning[(planning['Status'] == 'ALTA') | (planning['Status'] == 'BAJA TEMPORAL')]


definitivo = {'Línea': [],
         'Nombre': [],
         'Turno': [],
         'Agente': [],
         'Preferencia': [],
         'Pref. Texto': [],
         'Subturno': [],
         'Tipo de pausa': [],
         'Hora': [],
         'Pais': [],
         'Dia': []}

definitivo = pd.DataFrame(definitivo)

prevision = {'Dia': [],
         'Pais': [],
         'Hora': [],
         'Agentes': [],
         'Productividad': [],
         'Contestadas': [],
         'Recibidas': [],
         'Accesibilidad sin limitar': [],
         'Accesibilidad limitada': []}

prevision = pd.DataFrame(prevision)





#################################################################
##### HOLANDA
#################################################################

def pausasNL(dia,planning,definitivo,prevision):
    planning = planning[['Línea', 'Nombre', dia]]
    planning[dia] = planning[dia].astype(str)
    planning = planning.dropna(axis=0, subset=[dia])
    planning = planning[planning[dia] != 'V']
    planning = planning[planning[dia] != 'v']
    planning = planning[planning[dia] != 'B']
    planning = planning[planning[dia] != 'P']
    planning = planning[planning[dia] != 'F']
    planning = planning[planning[dia] != 'Li']
    planning = planning[planning[dia] != 'LI']
    planning = planning[planning[dia] != 'EJ']
    planning = planning[planning[dia] != '0']
    planning = planning.sort_values(by=dia)

    
    preferencias = pd.read_excel('U:\Turnos\Horarios y pausas\Definitivo\PreferenciasNL.xlsx')
    preferencias.rename(columns={'Nombre':'Agente','Agente':'Nombre'}, inplace=True)
    planning = planning.merge(preferencias, on = 'Nombre')
    
    forecast = pd.read_excel('U:\Turnos\Horarios y pausas\Definitivo\ForecastNL.xlsx')
    names = ['Hora']
    names[1:368] = pd.date_range(datetime.strptime('2024-01-01','%Y-%m-%d'), periods=366).tolist()
    forecast.columns = names
    forecast = forecast[dia]
    forecast = forecast.values

    turnos = pd.read_excel('U:\Turnos\Horarios y pausas\Definitivo\TurnosNL.xlsx', sheet_name='Sin pausas')

    horas = turnos.iloc[:,0]

    nombreturnos = turnos.columns
    nombreturnos = nombreturnos[1:]
    frecuencia = [0,0,0,0,0]
    frecuencia[0] = planning[dia].value_counts().get('M8',0) + planning[dia].value_counts().get('M8*',0)
    frecuencia[1] = planning[dia].value_counts().get('M9',0) + planning[dia].value_counts().get('M9*',0)
    frecuencia[2] = planning[dia].value_counts().get('M10',0) + planning[dia].value_counts().get('M10*',0)
    frecuencia[3] = planning[dia].value_counts().get('M11:30',0) + planning[dia].value_counts().get('M11:30*',0)
    frecuencia[4] = planning[dia].value_counts().get('M12',0) + planning[dia].value_counts().get('M12*',0)
    contador = np.array(frecuencia)

    turnos = turnos.drop(turnos.columns[0],axis=1)
    turnos = turnos.values

    agentes = np.dot(turnos,contador)

    productividad = 1
    tope_accesibilidad = 0.97

    contestadas = agentes * productividad
    contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad

    accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
    accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

    acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
    acc.iloc[:,0] = horas
    acc.iloc[:,1] = accesibilidad


    turnoscon = pd.read_excel('U:\Turnos\Horarios y pausas\Definitivo\TurnosNL.xlsx', sheet_name='Con pausas')
    pausas = turnoscon[27:30]



    planning = planning.assign(Subturno = np.zeros((len(planning['Nombre']),1)))
    planning = planning.assign(Lunch = np.zeros((len(planning['Nombre']),1)))

    orden = pd.read_excel('Orden horas.xlsx')


    # Asigno lunch
        
    # M8

    for i in range(len(planning['Nombre'])):
        if planning.iloc[i, 2] == 'M8':
            horas_dispo = pausas.loc[[28],['M8.1','M8.2','M8.3']]
            horas_dispo.index = ['Hora']
            horas_dispo = acc.loc[acc['Hora'].isin(horas_dispo.iloc[0])]
            horas_dispo = horas_dispo.assign(Subturno = ['M8.1','M8.2','M8.3'])

            m81 = planning['Subturno'].value_counts().get('M8.1', 0)
            m82 = planning['Subturno'].value_counts().get('M8.2', 0)
            m83 = planning['Subturno'].value_counts().get('M8.3', 0)

            horas_dispo = horas_dispo.assign(Frecuencia = [m81,m82,m83])

            maximo = horas_dispo['Accesibilidad'].max()
            maximo = horas_dispo[horas_dispo['Accesibilidad'] == maximo].iloc[0]
            planning.iloc[i,6] = maximo.loc['Subturno']
            planning.iloc[i,7] = maximo.loc['Hora']     

            posicion = maximo.loc['Hora']
            posicion = orden[orden['Hora'] == posicion]
            posicion = posicion.iloc[0,1] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde (2 posiciones más tarde)
            agentes[posicion] = agentes[posicion] - 1

            contestadas = agentes * productividad
            contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
            accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
            accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

            acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
            acc.iloc[:,0] = horas
            acc.iloc[:,1] = accesibilidad


    # M9

    for i in range(len(planning['Nombre'])):
        if planning.iloc[i, 2] == 'M9':
            horas_dispo = pausas.loc[[28],['M9.1','M9.2','M9.3','M9.4']]
            horas_dispo.index = ['Hora']
            horas_dispo = acc.loc[acc['Hora'].isin(horas_dispo.iloc[0])]
            horas_dispo = horas_dispo.assign(Subturno = ['M9.1','M9.2','M9.3','M9.4'])

            m91 = planning['Subturno'].value_counts().get('M9.1', 0)
            m92 = planning['Subturno'].value_counts().get('M9.2', 0)
            m93 = planning['Subturno'].value_counts().get('M9.3', 0)
            m94 = planning['Subturno'].value_counts().get('M9.3', 0)

            horas_dispo = horas_dispo.assign(Frecuencia = [m91,m92,m93,m94])

            maximo = horas_dispo['Accesibilidad'].max()
            maximo = horas_dispo[horas_dispo['Accesibilidad'] == maximo].iloc[0]
            planning.iloc[i,6] = maximo.loc['Subturno']
            planning.iloc[i,7] = maximo.loc['Hora']     

            posicion = maximo.loc['Hora']
            posicion = orden[orden['Hora'] == posicion]
            posicion = posicion.iloc[0,1] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde (2 posiciones más tarde)
            agentes[posicion] = agentes[posicion] - 1

            contestadas = agentes * productividad
            contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
            accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
            accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

            acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
            acc.iloc[:,0] = horas
            acc.iloc[:,1] = accesibilidad


    # M10

    for i in range(len(planning['Nombre'])):
        if planning.iloc[i, 2] == 'M10':
            horas_dispo = pausas.loc[[28],['M10.1','M10.2']]
            horas_dispo.index = ['Hora']
            horas_dispo = acc.loc[acc['Hora'].isin(horas_dispo.iloc[0])]
            horas_dispo = horas_dispo.assign(Subturno = ['M10.1','M10.2'])

            m101 = planning['Subturno'].value_counts().get('M10.1', 0)
            m102 = planning['Subturno'].value_counts().get('M10.2', 0)

            horas_dispo = horas_dispo.assign(Frecuencia = [m101,m102])

            maximo = horas_dispo['Accesibilidad'].max()
            maximo = horas_dispo[horas_dispo['Accesibilidad'] == maximo].iloc[0]
            planning.iloc[i,6] = maximo.loc['Subturno']
            planning.iloc[i,7] = maximo.loc['Hora']     

            posicion = maximo.loc['Hora']
            posicion = orden[orden['Hora'] == posicion]
            posicion = posicion.iloc[0,1] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde (2 posiciones más tarde)
            agentes[posicion] = agentes[posicion] - 1

            contestadas = agentes * productividad
            contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
            accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
            accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

            acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
            acc.iloc[:,0] = horas
            acc.iloc[:,1] = accesibilidad


    # M11:30

    for i in range(len(planning['Nombre'])):
        if planning.iloc[i, 2] == 'M11:30':
            horas_dispo = pausas.loc[[28],['M1130.1','M1130.2']]
            horas_dispo.index = ['Hora']
            horas_dispo = acc.loc[acc['Hora'].isin(horas_dispo.iloc[0])]
            horas_dispo = horas_dispo.assign(Subturno = ['M1130.1','M1130.2'])

            m111 = planning['Subturno'].value_counts().get('M1130.1', 0)
            m112 = planning['Subturno'].value_counts().get('M1130.2', 0)

            horas_dispo = horas_dispo.assign(Frecuencia = [m111,m112])

            maximo = horas_dispo['Accesibilidad'].max()
            maximo = horas_dispo[horas_dispo['Accesibilidad'] == maximo].iloc[0]
            planning.iloc[i,6] = maximo.loc['Subturno']
            planning.iloc[i,7] = maximo.loc['Hora']     

            posicion = maximo.loc['Hora']
            posicion = orden[orden['Hora'] == posicion]
            posicion = posicion.iloc[0,1] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde (2 posiciones más tarde)
            agentes[posicion] = agentes[posicion] - 1

            contestadas = agentes * productividad
            contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
            accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
            accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

            acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
            acc.iloc[:,0] = horas
            acc.iloc[:,1] = accesibilidad


    # M12

    for i in range(len(planning['Nombre'])):
        if planning.iloc[i, 2] == 'M12':
            horas_dispo = pausas.loc[[28],['M12']]
            horas_dispo.index = ['Hora']
            horas_dispo = acc.loc[acc['Hora'].isin(horas_dispo.iloc[0])]
            horas_dispo = horas_dispo.assign(Subturno = ['M12'])

            m12 = planning['Subturno'].value_counts().get('M12', 0)

            horas_dispo = horas_dispo.assign(Frecuencia = [m12])

            maximo = horas_dispo['Accesibilidad'].max()
            maximo = horas_dispo[horas_dispo['Accesibilidad'] == maximo].iloc[0]
            planning.iloc[i,6] = maximo.loc['Subturno']
            planning.iloc[i,7] = maximo.loc['Hora']     

            posicion = maximo.loc['Hora']
            posicion = orden[orden['Hora'] == posicion]
            posicion = posicion.iloc[0,1] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde (2 posiciones más tarde)
            agentes[posicion] = agentes[posicion] - 1

            contestadas = agentes * productividad
            contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
            accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
            accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

            acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
            acc.iloc[:,0] = horas
            acc.iloc[:,1] = accesibilidad



    # Creo columnas Lunch2 y Lunch3
        
    planning = planning.assign(Lunch2 = np.zeros((len(planning['Nombre']),1)))
    planning = planning.assign(Lunch3 = np.zeros((len(planning['Nombre']),1)))



    # Breaks

    planning = planning.assign(Break1 = np.zeros((len(planning['Nombre']),1)))
    planning = planning.assign(Break2 = np.zeros((len(planning['Nombre']),1)))

    for i in range(len(planning['Nombre'])):
        subturno = planning.iloc[i,6]
        pausa_sub = pausas.loc[27,subturno]
        planning.iloc[i,10] = pausa_sub
        pausa_sub = pausas.loc[29,subturno]
        planning.iloc[i,11] = pausa_sub


    # Turnos especiales

    especiales = pd.read_excel('U:\Turnos\Horarios y pausas\Definitivo\Turnos especiales.xlsx', sheet_name= 'Hoja1')

        # Carolien Verhoeven (6h al día con pausa de media hora)

    nombre = 'VERHOEVEN, CAROLIEN'
    if nombre in planning['Nombre'].values:
        turno = planning.loc[planning['Nombre'] == nombre, dia].values[0]
        lunch_esp3 = especiales.loc[especiales['Agente'] == nombre]
        lunch_esp3 = lunch_esp3.loc[lunch_esp3['Turno'] == turno, 'Lunch'].values[0]
        lunch_viejo = planning.loc[planning['Nombre'] == nombre, 'Lunch'].values[0]
        posicion = orden.loc[orden['Hora'] == lunch_viejo, 'Numero'].values[0] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde
        agentes[posicion] = agentes[posicion] + 1
        posicion = orden.loc[orden['Hora'] == lunch_esp3, 'Numero'].values[0]
        agentes[posicion] = agentes[posicion] - 1
                
                # Actualizo los arrays contestadas, accesibilidad, accesibilidad_dia y acc

        contestadas = agentes * productividad
        contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
        accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
        accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

        acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
        acc.iloc[:,0] = horas
        acc.iloc[:,1] = accesibilidad

                # Actualizo breaks y lunch
            
        planning.loc[planning['Nombre'] == nombre, 'Lunch'] = lunch_esp3
            
        break1_esp = especiales.loc[especiales['Agente'] == nombre]
        break1_esp = break1_esp.loc[break1_esp['Turno'] == turno, 'Break 1'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break1'] = break1_esp

        break2_esp = especiales.loc[especiales['Agente'] == nombre]
        break2_esp = break2_esp.loc[break2_esp['Turno'] == turno, 'Break 2'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break2'] = break2_esp
                
                # Resto las dos horas que no está cuando tiene un * en el turno
            
        if turno == 'M8':
            agentes[13] -= 1
            agentes[14] -= 1
            agentes[15] -= 1
            agentes[16] -= 1
        elif turno == 'M9':
            agentes[15] -= 1
            agentes[16] -= 1
            agentes[17] -= 1
            agentes[18] -= 1
        elif turno == 'M10':
            agentes[17] -= 1
            agentes[18] -= 1
            agentes[19] -= 1
            agentes[20] -= 1
        elif turno == 'M11:30':
            agentes[20] -= 1
            agentes[21] -= 1
            agentes[22] -= 1
            agentes[23] -= 1
        elif turno == 'M12':
            agentes[8] -= 1
            agentes[9] -= 1
            agentes[10] -= 1
            agentes[11] -= 1

        # Christain Gijsbregts (5h al día con pausa de media hora)

    nombre = 'GIJSBREGTS, CHRISTAIN CONNOR'
    if nombre in planning['Nombre'].values:
        turno = planning.loc[planning['Nombre'] == nombre, dia].values[0]
        lunch_esp3 = especiales.loc[especiales['Agente'] == nombre]
        lunch_esp3 = lunch_esp3.loc[lunch_esp3['Turno'] == turno, 'Lunch'].values[0]
        lunch_viejo = planning.loc[planning['Nombre'] == nombre, 'Lunch'].values[0]
        posicion = orden.loc[orden['Hora'] == lunch_viejo, 'Numero'].values[0] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde
        agentes[posicion] = agentes[posicion] + 1
        posicion = orden.loc[orden['Hora'] == lunch_esp3, 'Numero'].values[0]
        agentes[posicion] = agentes[posicion] - 1
                
                # Actualizo los arrays contestadas, accesibilidad, accesibilidad_dia y acc

        contestadas = agentes * productividad
        contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
        accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
        accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

        acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
        acc.iloc[:,0] = horas
        acc.iloc[:,1] = accesibilidad

                # Actualizo breaks y lunch
            
        planning.loc[planning['Nombre'] == nombre, 'Lunch'] = lunch_esp3
            
        break1_esp = especiales.loc[especiales['Agente'] == nombre]
        break1_esp = break1_esp.loc[break1_esp['Turno'] == turno, 'Break 1'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break1'] = break1_esp

        break2_esp = especiales.loc[especiales['Agente'] == nombre]
        break2_esp = break2_esp.loc[break2_esp['Turno'] == turno, 'Break 2'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break2'] = break2_esp
            
        if turno == 'M8':
            agentes[11] -= 1
            agentes[12] -= 1
            agentes[13] -= 1
            agentes[14] -= 1
            agentes[15] -= 1
            agentes[16] -= 1
        elif turno == 'M9':
            agentes[13] -= 1
            agentes[14] -= 1
            agentes[15] -= 1
            agentes[16] -= 1
            agentes[17] -= 1
            agentes[18] -= 1
        elif turno == 'M10':
            agentes[15] -= 1
            agentes[16] -= 1
            agentes[17] -= 1
            agentes[18] -= 1
            agentes[19] -= 1
            agentes[20] -= 1
        elif turno == 'M11:30':
            agentes[18] -= 1
            agentes[19] -= 1
            agentes[20] -= 1
            agentes[21] -= 1
            agentes[22] -= 1
            agentes[23] -= 1
        elif turno == 'M12':
            agentes[8] -= 1
            agentes[9] -= 1
            agentes[10] -= 1
            agentes[11] -= 1
            agentes[12] -= 1
            agentes[13] -= 1

        # Kylie Lenaerts (los lunes hace 9h)

    nombre = 'LENAERTS, KYLIE'
    if nombre in planning['Nombre'].values and dia.weekday() == 0:
        turno = planning.loc[planning['Nombre'] == nombre, dia].values[0]
        lunch_esp3 = especiales.loc[especiales['Agente'] == nombre]
        lunch_esp3 = lunch_esp3.loc[lunch_esp3['Turno'] == turno, 'Lunch'].values[0]
        lunch_viejo = planning.loc[planning['Nombre'] == nombre, 'Lunch'].values[0]
        posicion = orden.loc[orden['Hora'] == lunch_viejo, 'Numero'].values[0] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde
        agentes[posicion] = agentes[posicion] + 1
        posicion = orden.loc[orden['Hora'] == lunch_esp3, 'Numero'].values[0]
        agentes[posicion] = agentes[posicion] - 1
                
                # Actualizo los arrays contestadas, accesibilidad, accesibilidad_dia y acc

        contestadas = agentes * productividad
        contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
        accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
        accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

        acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
        acc.iloc[:,0] = horas
        acc.iloc[:,1] = accesibilidad

                # Actualizo breaks y lunch
            
        planning.loc[planning['Nombre'] == nombre, 'Lunch'] = lunch_esp3
            
        break1_esp = especiales.loc[especiales['Agente'] == nombre]
        break1_esp = break1_esp.loc[break1_esp['Turno'] == turno, 'Break 1'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break1'] = break1_esp

        break2_esp = especiales.loc[especiales['Agente'] == nombre]
        break2_esp = break2_esp.loc[break2_esp['Turno'] == turno, 'Break 2'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break2'] = break2_esp
            
        if turno == 'M8':
            agentes[17] += 1
            agentes[18] += 1
        elif turno == 'M9':
            agentes[19] += 1
            agentes[20] += 1
        elif turno == 'M10':
            agentes[21] += 1
            agentes[22] += 1
        elif turno == 'M11:30':
            agentes[5] += 1
            agentes[6] += 1
        elif turno == 'M12':
            agentes[6] += 1
            agentes[7] += 1

        # Miriam Bakker (7h el lunes, 6h martes, miércoles y jueves, 8h el sábado)

    nombre = 'BAKKER, MIRIAM BEATRICE'
    if nombre in planning['Nombre'].values and dia.weekday() != 5:
        turno = planning.loc[planning['Nombre'] == nombre, dia].values[0]
        lunch_esp3 = especiales.loc[especiales['Agente'] == nombre]
        lunch_esp3 = lunch_esp3.loc[lunch_esp3['Turno'] == turno, 'Lunch'].values[0]
        lunch_viejo = planning.loc[planning['Nombre'] == nombre, 'Lunch'].values[0]
        posicion = orden.loc[orden['Hora'] == lunch_viejo, 'Numero'].values[0] - 2 # Resto 2 porque el horario de Holanda empieza 1h más tarde
        agentes[posicion] = agentes[posicion] + 1
        posicion = orden.loc[orden['Hora'] == lunch_esp3, 'Numero'].values[0]
        agentes[posicion] = agentes[posicion] - 1
                
                # Actualizo los arrays contestadas, accesibilidad, accesibilidad_dia y acc

        contestadas = agentes * productividad
        contestadas = np.minimum(contestadas,forecast) * tope_accesibilidad
        accesibilidad = np.nan_to_num(np.divide(contestadas, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
        accesibilidad_dia = np.sum(contestadas) / np.sum(forecast) * 100

        acc = pd.DataFrame(np.random.randn(24,2), columns=['Hora','Accesibilidad'])
        acc.iloc[:,0] = horas
        acc.iloc[:,1] = accesibilidad

                # Actualizo breaks y lunch
            
        planning.loc[planning['Nombre'] == nombre, 'Lunch'] = lunch_esp3
            
        break1_esp = especiales.loc[especiales['Agente'] == nombre]
        break1_esp = break1_esp.loc[break1_esp['Turno'] == turno, 'Break 1'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break1'] = break1_esp

        break2_esp = especiales.loc[especiales['Agente'] == nombre]
        break2_esp = break2_esp.loc[break2_esp['Turno'] == turno, 'Break 2'].values[0]
        planning.loc[planning['Nombre'] == nombre, 'Break2'] = break2_esp
                
                # Resto las dos horas que no está cuando tiene un * en el turno
        if dia.weekday() == 0:    
            if turno == 'M8':
                agentes[15] -= 1
                agentes[16] -= 1
            elif turno == 'M9':
                agentes[17] -= 1
                agentes[18] -= 1
            elif turno == 'M10':
                agentes[19] -= 1
                agentes[20] -= 1
            elif turno == 'M11:30':
                agentes[22] -= 1
                agentes[23] -= 1
            elif turno == 'M12':
                agentes[8] -= 1
                agentes[9] -= 1

        if dia.weekday() == 1 | dia.weekday() == 2 | dia.weekday() == 3:    
            if turno == 'M8':
                agentes[13] -= 1
                agentes[14] -= 1
                agentes[15] -= 1
                agentes[16] -= 1
            elif turno == 'M9':
                agentes[15] -= 1
                agentes[16] -= 1
                agentes[17] -= 1
                agentes[18] -= 1
            elif turno == 'M10':
                agentes[17] -= 1
                agentes[18] -= 1
                agentes[19] -= 1
                agentes[20] -= 1
            elif turno == 'M11:30':
                agentes[20] -= 1
                agentes[21] -= 1
                agentes[22] -= 1
                agentes[23] -= 1
            elif turno == 'M12':
                agentes[8] -= 1
                agentes[9] -= 1
                agentes[10] -= 1
                agentes[11] -= 1



    # Línea inglesa

    linea_inglesa = ['AVILA RODRIGO, CINTIA DEL VALLE','ANTUNANO LLANA, PAULA','WARREN GUILLOT, ANTHONY JOHN','RAVELO CABALLERO, ALEJANDRO']



    # Actualizo columnas Lunch2 y Lunch3
        
    planning = planning.assign(Lunch2 = np.zeros((len(planning['Nombre']),1)))
    planning = planning.assign(Lunch3 = np.zeros((len(planning['Nombre']),1)))

    for i in planning.index:
        planning.loc[i, 'Lunch2'] = (datetime.combine(datetime.today(), planning.loc[i, 'Lunch']) + timedelta(minutes=10)).time()
        planning.loc[i, 'Lunch3'] = (datetime.combine(datetime.today(), planning.loc[i, 'Lunch']) + timedelta(minutes=20)).time()





    planning2 = pd.melt(planning, id_vars=['Línea', 'Nombre',dia,'Agente', 'Preferencia', 'Pref. Texto', 'Subturno'], var_name='Tipo de pausa', value_name='Hora')
    planning2['Pais'] = ['HOLANDA'] * len(planning2['Nombre'])
    planning2['Dia'] = [dia] * len(planning2['Nombre'])

    definitivo = pd.concat([definitivo, planning2], ignore_index=True)

    auxprevision = {'Dia': [],
         'Pais': [],
         'Hora': [],
         'Agentes': [],
         'Productividad': [],
         'Contestadas': [],
         'Recibidas': [],
         'Accesibilidad sin limitar': [],
         'Accesibilidad limitada': []}

    auxprevision = pd.DataFrame(auxprevision)

    auxprevision['Hora'] = acc['Hora']
    auxprevision['Accesibilidad limitada'] = acc['Accesibilidad']
    auxprevision['Dia'] = [dia] * len(auxprevision['Hora'])
    auxprevision['Pais'] = ['HOLANDA'] * len(auxprevision['Hora'])
    auxprevision['Accesibilidad sin limitar'] = np.nan_to_num(np.divide(agentes * productividad, forecast, out=np.zeros_like(contestadas), where=(forecast != 0)) * 100)
    auxprevision['Agentes'] = agentes
    auxprevision['Productividad'] = [productividad] * len(auxprevision['Hora'])
    auxprevision['Contestadas'] = contestadas
    auxprevision['Recibidas'] = forecast

    prevision = pd.concat([prevision,auxprevision], ignore_index = True)

    return definitivo, prevision






inicio = input('¿Desde qué día? (aaaa-mm-dd) ')
inicio = datetime.strptime(inicio,'%Y-%m-%d')
fin = input('¿Hasta qué día? (aaaa-mm-dd) ')
fin = datetime.strptime(fin,'%Y-%m-%d')

diferencia = fin - inicio
diferencia = diferencia.days + 1

for i in range(diferencia):
    definitivo, prevision = pausasNL(inicio + timedelta(days=i),planning,definitivo,prevision)

columnas_a_concatenar = definitivo.iloc[:, 11:]

definitivo['Turnos'] = columnas_a_concatenar.apply(lambda row: ' '.join(str(cell) for cell in row if pd.notna(cell)), axis=1)
definitivo.drop(definitivo.columns[11:-1],axis=1,inplace=True)
definitivo.drop(definitivo.columns[2],axis=1,inplace=True)


definitivo.to_excel('Pausas_limpioNL.xlsx',index=False)
prevision.to_excel('Prevision_limpioNL.xlsx', index=False)
