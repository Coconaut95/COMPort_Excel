"""
Requerimientos:
- Mostrar en forma de tabla el error por fase para cada punto de corriente.
ej, FAse 1: 70 mA- error; 5A- error
- Con la tabla anterior hacer graficos del error por fase en funcion de la corriente.

LP:   L1-- 220V 5A
LP:   L1-- 220V 0.04A
LP:   L-2- 220V 5A
LP:   L-2- 220V 0.04A
LP:   L--3 220V 5A
LP:   L--3 220V 0.04A
"""
import re
import os
import keyboard
from dic_to_list import dic_to_list
from write_excel import write_excel_scva, write_excel_norma

key = True

while key:

    # ---------------------------Inicia comunicacion serie-------------------------------------------------------------------
    from com_serie import read_serial_port
    read_serial_port()
    # _______________________________________________________________________________________________________________________

    file_load = [] #contiene cada linea del archivo de datos.

    while True:
        id_medidor = input('Ingrese tipo y nº de serie completo del medidor: ').upper()
        if len(id_medidor) == 12:
            # TIPO[nº de serie: ########]
            id_num = int(id_medidor[4:])
            tipo = id_medidor[:4]
            break
        else:
            print('Intente nuevamente por favor.')

    def open_file():
        global file_load
        abs_path = os.path.dirname(__file__)
        relative_path = "/DATA"
        full_path = abs_path + relative_path
        filename = full_path + '/data.txt' # archivo por default: data.txt
        with open(filename, 'r') as file:
            file_load = [lines.strip() for lines in file.readlines()]
        return file_load

    def data_from_txt():
        dic_list = []
        i = 0

        for info in file_load:
            expression = 'Mid:   ([+-0-9]+[.0-9]+)'
            matches = re.search(expression, info)
            expression1 = 'LP:   [A-Z]([0-9]+--|-[0-9]-|--[0-9]|[0-9]+)+ ([0-9]+)V+ ([0-9\.0-9]+)A+ phi=([+-0-9\.0-9]+)'
            matches1 = re.search(expression1, info)

            if matches1:
                fase = str(matches1.group(1))
                tension = float(matches1.group(2))
                corriente = float(matches1.group(3))
                angulo = str(matches1.group(4))

                dic = {'fase': fase, 'tension': tension, 'phi': angulo, 'corriente': corriente, 'error': None}
                dic_list.append(dic)

            if matches:
                value_error = float(matches.group(1))
                dic_list[i]['error'] = value_error
                i += 1

        return dic_list

    #Abre el archivo de datos que contiene la info del MTE.
    open_file()

    print("Ingrese:" + '\n' + "1) Para excel SCVA:" + '\n' + "2) Para excel Ensayo por norma:")
    input_num = input()

    if input_num == '1':
        dic_to_list(data_from_txt())
        write_excel_scva(id_num)

    elif input_num == '2':
        write_excel_norma(data_from_txt(), id_num, tipo)

    ##################################################################
    # Opciones:
    ##################################################################
    print('##############################################')
    print('Si desea cargar nuevos datos, presione ENTER.')
    print('En caso contrario, presione cualquier tecla.')

    if keyboard.read_key() == 'enter':
        continue
    else:
        key = False
