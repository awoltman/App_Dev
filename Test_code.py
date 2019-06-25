from pymodbus.client.sync import ModbusTcpClient as ModbusClient
import matplotlib.pyplot as plt
import xlwt
from datetime import datetime
import sqlite3
import time

conn = sqlite3.connect('test_data.db')
client = ModbusClient(mesthod = 'tcp', host = '10.81.7.195', port = 8899)
UNIT = 0x01
'''
## Setting up styles for Excel ##
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook()
ws = wb.add_sheet('Tempurature Data')
ws.write(0, 1, 'T1', style0)
ws.write(0, 2, 'T2', style0)
ws.write(0, 3, 'T3', style0)
ws.write(0, 4, 'T4', style0)
ws.write(0, 4, 'Time', style0)
'''
select()

def recore_temps:
    try:
        c = conn.cursor()
        c.execute('''CREATE TABLE stocks (TEMPS, FREEZE_TIMES)''')
        c.execute('''INSERT INTO TEMPS values ('Time', 'T1', 'T2', 'T3', 'T4')''')
        c.execute('''INSERT INTO FREEZE_TIMES values ('Time', 'Freeze Time 1', 'Freeze Time 2', 'Freeze Time 3', 'Freeze Time 4', 'Freeze Time 5', 'Freeze Time 6', 'Freeze Time 7','Freeze Time 8', 'Freeze Time 9', 'Freeze Time 10',
         'Freeze Time 11', 'Freeze Time 12', 'Freeze Time 13', 'Freeze Time 14', 'Freeze Time 15', 'Freeze Time 16', 'Freeze Time 17','Freeze Time 18', 'Freeze Time 19', 'Freeze Time 20')''')
        while True:
        named_tuple = time.localtime() # get struct_time
        time_string = time.strftime("%m/%d/%Y %H:%M.%S")

        Temps_store = client.read_holding_registers(6,4,UNIT)
        Freezetime_temp = client.read_holding_registers(574,20,unit = UNIT)
        time_temp = (time_string,Freezetime_temp.registers[0],Freezetime_temp.registers[1],Freezetime_temp.registers[2],Freezetime_temp.registers[3],Freezetime_temp.registers[4],Freezetime_temp.registers[5]\
            ,Freezetime_temp.registers[6],Freezetime_temp.registers[7],Freezetime_temp.registers[8],Freezetime_temp.registers[9],Freezetime_temp.registers[10],Freezetime_temp.registers[11]\
                ,Freezetime_temp.registers[12],Freezetime_temp.registers[13],Freezetime_temp.registers[14],Freezetime_temp.registers[15],Freezetime_temp.registers[16],Freezetime_temp.registers[17]\
                    ,Freezetime_temp.registers[18],Freezetime_temp.registers[19],Freezetime_temp.registers[20])
        temp_temp = (time_string, Temps_store.registers[0],Temps_store.registers[1],Temps_store.registers[2],Temps_store.registers[3])
        c.execute("INSERT INTO FREEZE_TIMES values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", time_temp)
        c.execute("INSERT INTO TEMPS values (?,?,?,?,?)")
        conn.commit()
'''     
        ##This section is for writing to Excel##

        ws.write(ex, 0, time_string, style1)
        ws.write(ex, 1, Temps_temp.registers[0], style0)
        ws.write(ex, 2, Temps_temp.registers[1], style0)
        ws.write(ex, 3, Temps_temp.registers[2], style0)
        ws.write(ex, 4, Temps_temp.registers[3], style0)
'''
    except KeyboardInterrupt:
        '''
        ## used to save EXCEL file once done collecting data ##
        wb.save('temp.xls')
        '''
        conn.close()
        pass

def select:
    print('C for collect')
    print('D for done')
    g = input('Enter what you would like to do:')

    if(g == 'C'):
        record_temps()
    elif(g == 'D'):
        client.close()
    else:
        select()
