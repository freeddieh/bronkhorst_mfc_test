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
# 9: 'Setpoint',              # Setpoint from 0-32000
# 129: 'Capacity Unit',       # Readout unit of the MFC at capacity
# 205: 'fMeasure',            # Actual flow in capacity unit
# 206: 'fSetpoint',           # Set flow in capacity unit 
# 253: 'Standard Mass FLow'   # In units ln/min
# }

#######################################################################
###-----------------Bronkhorst MFC Controller setup-----------------###
#######################################################################


def read_bh_flow(bh_mfc: BronkhorstMFC) -> float:
    """
    
    Reads the setpoint of a Bronkhorst MFC
    
    :param bh_mfc: BronkhorstMFC object for the MFC to read the setpoint of
    
    :return: Float of the current flow in mLn/min
    
    """

    # Find the capacity unit of the MFC to standardise output to mLn/min
    # 205 is the DDE number of the actual flow of a Bronkhorst MFC
    if 'mln' in bh_mfc.readout_unit:
        bh_mfc.data_unit = 'mLn/min'
        return float(bh_mfc.read_bronkhorst(205)[205])
    elif 'mln' not in bh_mfc.readout_unit and 'ln' in bh_mfc.readout_unit:
        bh_mfc.data_unit = 'mLn/min'
        return (float(bh_mfc.read_bronkhorst(205)[205]) * 1000)
    
    
def read_bh_set(bh_mfc: BronkhorstMFC) -> float:
    """
    
    Reads the setpoint of a Bronkhorst MFC
    
    :param bh_mfc: BronkhorstMFC object for the MFC to read the setpoint of
    
    :return: Float of the current flow in mLn/min
    
    """

    # Find the capacity unit of the MFC to standardise output to mLn/min
    # 205 is the DDE number of the actual flow of a Bronkhorst MFC
    if 'mln' in bh_mfc.readout_unit:
        bh_mfc.data_unit = 'mLn/min'
        return float(bh_mfc.read_bronkhorst(206)[206])
    elif 'mln' not in bh_mfc.readout_unit and 'ln' in bh_mfc.readout_unit:
        bh_mfc.data_unit = 'mLn/min'
        return (float(bh_mfc.read_bronkhorst(206)[206]) * 1000)


def data_logging(ser: serial.Serial, 
                 headers: str, 
                 log_name: str, 
                 bronkhorst_mfc: list[BronkhorstMFC]|None = None) -> None:
    '''
    
    Reads data from a serial print from an Arduino, adds a 
    timestamp, and logs the data in 5 second intervals, and 
    stores the data in daily .csv files with customizable names

    :param ser: Serial object for the port of the Arduino
    :param headers: Header of the log to be made
    :param log_name: Identification name of the log
    :param bronkhorst_mfc: List of BronkhorstMFC objects to 
                           read and include in the log
    
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
        data = f'{time.strftime("%Y-%m-%d %H:%M:%S")},{bh_data[0]},{bh_data[1]}'
        file.write(data + '\n')
        file.flush()

        time.sleep(5)


def error_log(error: str, log_name: str) -> None:
    """
    
    Error log writer for data logging.

    :param error: Error message for the specific error.
    :param log_name: Name of the error log file
    
    """
    # Function for logging errors in the data when they occur
    error_log = open(log_name, 'a')
    error = f'{time.strftime("%Y-%m-%d %H:%M:%S")}, {error}'
    error_log.write(error + '\n')
    error_log.close()


def main_logger(bronkhorsts) -> None:
    '''
    
    Function for execution of the data logger.

    :param bronkhorsts: List of Bronkhorst MFC objects to be controlled.
    
    '''

    # The header is defined manually
    # Please ensure that the header order and the order of the Bronkhorst MFC list is identical
    headers = 'Date,Bronkhorst 100SCCM [mLn/min], Bronkhorst 2.5SLM[mLn/min]'
    log_name = 'flow'

    data_logging(headers, log_name, bronkhorsts)


# Only run the script from this document
if __name__ == '__main__':  
    main_logger()