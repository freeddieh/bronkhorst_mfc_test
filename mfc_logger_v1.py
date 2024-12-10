import os
import time
import serial
import propar as pp
import datetime as dt
from airpy import *
from serial.tools import list_ports

#######################################################################
###----------Most commonly used Bronkhorst MFC DDE numbers----------###
#######################################################################

# {
# 8: 'Measure',               # Measurement from 0-32000 (32000 = 100%)
# 9: 'Setpoint',              # Setpoint from 0-3200
# 129: 'Capacity Unit',       # Readout unit of the MFC at capacity
# 205: 'fMeasure',            # Actual flow in capacity unit
# 206: 'fSetpoint',           # Set flow in capacity unit 
# 253: 'Standard Mass FLow'   # In units ln/min
# }

#######################################################################
###-----------------Bronkhorst MFC Controller setup-----------------###
#######################################################################


def read_bh_flow(bh_mfc: BronkhorstMFC) -> float:
    if 'mln' in bh_mfc.readout_unit:
        bh_mfc.data_unit = 'mLn/min'
        return float(bh_mfc.read_bronkhorst(205)[205])
    elif 'mln' not in bh_mfc.readout_unit and 'ln' in bh_mfc.readout_unit:
        bh_mfc.data_unit = 'mLn/min'
        return (float(bh_mfc.read_bronkhorst(205)[205]) * 1000)


def data_logging(ser: serial.Serial, 
                 headers: str, 
                 log_name: str, 
                 bronkhorst_mfc: list[BronkhorstMFC]|None = None) -> None:
    '''
    
    Reads data from a serial print from an Arduino, adds a 
    timestamp, and logs the data in 5 second intervals, and 
    stores the data in daily .csv files with customizable names

    '''
    
    # Defines all constants for use in the loop
    today = dt.date.today()
    filename = f'test/logs/{today}_{log_name}.csv'
    error_log_name = f'test/errorlogs/{today}_data_errorlog.csv'
    file = open(filename, 'a')
    line = 0

    # Use a try except statement to close the serial port and file
    # when stopping the data logging from the terminal
    print(f'Data logging for {today} started at {time.strftime("%H:%M:%S")}.')
    while True:
        bh_data = []
        for mfc in bronkhorst_mfc:
            bh_data.append(read_bh_flow(mfc))
        get_data = str(ser.readline().decode())
        data = get_data[0:][:-2]
        ### Flow error check ###

        #try:
        #    big_flow = float(data.split(',')[1])
        #except (IndexError, ValueError):
        #    pass

        # Exit the loop when the day switches, and starts a new loop for the next day
        if today != dt.date.today():
            file.close()
            print(f'Data collection for {today} complete! New file created.')
            data_logging(ser, headers, log_name)
            break
        
        # If at the start of the document write headers
        if line == 0 and os.stat(filename).st_size==0:
            file.write(headers + '\n')
            line += 1

        # Check that the data fetched complete and the data sturcture 
        # is not corrupted as it is read from the Arduino
        if ',' not in data or len(data.split(',')) != 2:
            error_message = f'Error: Data corrupted ({data}).'
            error_log(error_message, error_log_name)
            continue
        try:
            for value in data.split(','):
                float(value)
        except ValueError:
            error_message = f'Error: Data corrupted ({data}).'
            error_log(error_message, error_log_name)
            continue

        #######################################################
        ### SECTION FOR MFC ERROR CHECK IMPLEMENT AS NEEDED ###
        #######################################################

        # Check the big MFC output for data corruption that 
        # affects the measurement but not the overall data structure
        #if prev_big_flow is not None and big_flow is not None:
        #    if not prev_big_flow-0.5 < big_flow < prev_big_flow+0.5:
        #        if big_flow_check >= 12:
        #            prev_big_flow = big_flow
        #            big_flow_check = 0
        #            continue
        #        else:
        #            error_message = f'Error: Data corrupted ({data}).'
        #            error_log(error_message, error_log_name)
        #            big_flow_check += 1
        #            continue
        #    else:
        #        prev_big_flow = big_flow
        #elif big_flow is not None and prev_big_flow is None:
        #    prev_big_flow = big_flow

        # Check if the small flow is correct
        #if float(data.split(',')[0]) < 140:
        #    error_message = f'Error: Data corrupted ({data}).'
        #    error_log(error_message, error_log_name)
        #    continue

        # Add a timestamp and write the data to the .csv file
        data = f'{time.strftime("%Y-%m-%d %H:%M:%S")},{data},{bh_data[0]},{bh_data[1]}'
        file.write(data + '\n')
        file.flush()

        time.sleep(5)


def error_log(error: str, log_name: str) -> None:
    # Function for logging errors in the data when they occur
    error_log = open(log_name, 'a')
    error = f'{time.strftime("%Y-%m-%d %H:%M:%S")}, {error}'
    error_log.write(error + '\n')
    error_log.close()


def main_logger(bronkhorsts) -> None:
    '''
    
    Function for execution of the data logger.
    
    '''
    measure_MFC = Arduino() 
    serial_read = measure_MFC.find_arduino_port()

    # The header is defined manually
    # Please ensure that the header order and the order of the Bronkhorst MFC list is identical
    headers = 'Date,Brooks 100SCCM [mLn/min],Brooks 2.5SLM [mLn/min],Bronkhorst 100SCCM [mLn/min], Bronkhorst 2.5SLM[mLn/min]'
    log_name = 'flow'

    data_logging(serial_read, headers, log_name, bronkhorsts)


# Only run the script from this document
if __name__ == '__main__':  
    main_logger()