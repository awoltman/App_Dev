from pymodbus.client.sync import ModbusTcpClient as ModbusClient
import matplotlib.pyplot as plt
import xlwt
from datetime import datetime
import sqlite3
import time
def init():
    conn = sqlite3.connect('test_data.db')
    client = ModbusClient(mesthod = 'tcp', host = '10.81.7.195', port = 8899)
    UNIT = 0x01

    c = conn.cursor()
    time_temp = ()
'''
## Initialized the database(.db) file ##
c.execute("CREATE TABLE TEMPS ('Time', 'T1', 'T2', 'T3', 'T4')")
c.execute("CREATE TABLE FREEZE_TIMES ('Time', 'Freeze_Time_1', 'Freeze_Time 2', 'Freeze Time 3', 'Freeze Time 4', 'Freeze Time 5', 'Freeze Time 6', 'Freeze Time 7','Freeze Time 8', 'Freeze Time 9', 'Freeze Time 10',\
'Freeze Time 11', 'Freeze Time 12', 'Freeze Time 13', 'Freeze Time 14', 'Freeze Time 15', 'Freeze Time 16', 'Freeze Time 17','Freeze Time 18', 'Freeze Time 19', 'Freeze Time 20')")
'''
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
def record_temps():
    try:
        while True:
            #named_tuple = time.localtime() # get struct_time
            time_string = time.strftime("%m/%d/%Y %H:%M.%S")
            Freezetime_temp = client.read_holding_registers(574,20,unit = UNIT)
            f_one = Freezetime_temp.registers[0]
            f_two = Freezetime_temp.registers[1]
            f_three = Freezetime_temp.registers[2]
            f_four = Freezetime_temp.registers[3]
            f_five = Freezetime_temp.registers[4]
            f_six = Freezetime_temp.registers[5]
            f_seven = Freezetime_temp.registers[6]
            f_eight = Freezetime_temp.registers[7]
            f_nine = Freezetime_temp.registers[8]
            f_ten = Freezetime_temp.registers[9]
            f_eleven = Freezetime_temp.registers[10]
            f_twelve = Freezetime_temp.registers[11]
            f_thirteen = Freezetime_temp.registers[12]
            f_fourteen = Freezetime_temp.registers[13]
            f_fifteen = Freezetime_temp.registers[14]
            f_sixteen = Freezetime_temp.registers[15]
            f_seventeen = Freezetime_temp.registers[16]
            f_eighteen = Freezetime_temp.registers[17]
            f_nineteen = Freezetime_temp.registers[18]
            f_twenty = Freezetime_temp.registers[19]

            time_temp = (time_string,f_one,f_two,f_three,f_four,f_five,f_six,f_seven,f_eight,f_nine,f_ten,f_eleven,f_twelve,f_thirteen,f_fourteen,f_fifteen,f_sixteen,f_seventeen,f_eighteen,f_nineteen,f_twenty)
            c.execute("INSERT INTO FREEZE_TIMES values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", time_temp)
            Temps_store = client.read_holding_registers(6,4,unit =UNIT)
            temp_temp = (time_string, Temps_store.registers[0],Temps_store.registers[1],Temps_store.registers[2],Temps_store.registers[3])
            c.execute("INSERT INTO TEMPS values (?,?,?,?,?)", temp_temp)
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
        select()
def reset_default():
    client.write_registers(451,1,unit =UNIT)
    pass

def select():
    print('C for collect')
    print('D for done')
    print('R for reset defaults')
    g = input('Enter what you would like to do:')

    if(g == 'C'):
        record_temps()
    elif(g == 'D'):
        client.close()
    elif(g == 'R'):
        reset_default()
    else:
        select()
