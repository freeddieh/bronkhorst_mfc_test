import serial
import warnings
import propar as pp
import tkinter as tk
from tkinter import ttk
from serial.tools import list_ports
from typing import Dict, Tuple, Any


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


class BronkhorstMFC:
    def __init__(self, port: str) -> None:
        """
        
        Initializes the BronkhorstMFC class by reading the port of a 
        connected Bronkhorst MFC
        
        """

        self.port = port
        self.communication = None
    
    def initialize_communication(self) -> None:
        self.communication = pp.instrument(self.port)

    def read_bronkhorst(self, dde_numbers: list[int]) -> dict:
        """
        
        :param dde_numbers: Parameter numbers of parameters to read
        
        :returns: Dictionary with read parameter values and the parameter number
        """

        if self.communication is None:
            self.initialize_communication()
        parameters = {dde: self.communication.readParameter(dde) for dde in dde_numbers}
        return parameters
    
    def write_bronkhorst(self, dde_number: int, value: Any) -> None:
        """

        :param dde_number: Parameter number for the parameter to be written
        :param value: Value to write to the parameter number
        
        """
        if self.communication is None:
            self.initialize_communication()

        self.communication.writeParameter(dde_number, value)

class Arduino:
    def __init__(self) -> None:
        pass

    def find_arduino_port(self) -> serial.Serial:
        '''
        
        Finds the port for all connected Arduino boards

        :returns: serial.Serial object for the connected Arduino
        
        '''
    
        arduino_ports = [p.device for p in list_ports.comports() if p.manufacturer
                         is not None and 'Arduino' in p.manufacturer]

        if not arduino_ports:
            raise IOError('No connected Arduino found. Please check connections.')
        if len(arduino_ports) > 1:
            warnings.warn(f'Multiple connected Arduinos found using first found at port: {arduino_ports[0]}.')

            # Ready for implementation of chosing from multiple Arduino boards 
            # but would not work with cron automatization 
            
            #for idx, port in enumerate(arduino_ports):
            #    print(f'{idx+1}) Arduino found connected at port: {port}')
            #arduino_to_use = input('Please specify line number of Arduino to use:')
            #port_error = 'Please specify the portnumber as an integer within the printed list range.'
            #try:
            #    arduino_to_use = int(arduino_to_use)
            #except ValueError as e:
            #    raise e(port_error)
            #if len(arduino_ports)+1 < arduino_to_use or arduino_to_use <= 0:
            #    raise ValueError(port_error)
            
        return serial.Serial(arduino_ports[0], 9600)
    

class DropdownMenu:
    def __init__(self, root: tk.Tk, options: list[str]) -> None:
        """
        
        Initializes dropdown menu from the root, and with the input options
        
        """

        self.root = root
        self.selected_value = None  # To store the selected value
        
        # Create a label for the window
        self.label = tk.Label(root, text='Vælg Program')
        self.label.pack(pady=10)

        # Create the combobox (dropdown)
        self.options = options
        self.combobox = ttk.Combobox(root, values=self.options)
        self.combobox.pack(pady=10)

        # Set a default value
        self.combobox.set('')

        # Create a button to confirm selection
        self.button = tk.Button(root, text='Bekræft', command=self.on_button_click)
        self.button.pack(pady=10)
    
        # Create a label to display feedback message
        self.feedback_label = tk.Label(root, text="", fg="green")  # Default message is empty
        self.feedback_label.pack(pady=10)

    def on_button_click(self) -> None:
        """
        
        Defines what happens when the confirm button is clickes
        
        """
        # Get the selected value from the combobox
        self.selected_value = str(self.combobox.get())

        if not self.selected_value:
            # Display error message if no value is selected
            self.feedback_label.config(text='Fejl! Vælg Program fra dropdown menuen', fg="red")
        else:
            # Display success message if a value is selected
            self.feedback_label.config(text=f"Program valgt: {self.selected_value}", fg="green")
            # Optionally, destroy the window after feedback is displayed, with a delay of 600 ms
            self.root.after(600, self.root.destroy)
        

    def get_selected_value(self) -> str:
        """
        
        Returns the selected value in a string format

        :returns: String of selected value 

        """
        return self.selected_value


def find_bronkhorst_ports() -> dict:
    """

    Finds all relevant serial ports that might have MFC connections

    :return: Dictionary of connected Bronkhorst MFC's with 
             associated flow rates as a dictionary
    
    """
    
    # Connect to the local instrument, when no settings provided
    # defaults to locally connected instrument (address=0x80, baudrate=38400)
    mfcs = {}
    mfc_ports = [p.device for p in list_ports.comports() if 'FTDI' 
                in p.manufacturer and p.manufacturer is not None]
    for port in mfc_ports:
        mfc = pp.instrument(port)
        if mfc.readParameter(8) is not None:
            _ = mfc.master.get_nodes()[0]
            flowrate = str(mfc.readParameter(253))[:3]
            mfcs[f'{flowrate} {mfc.readParameter(129).strip()}'] = mfc
            print(f'MFC with flowrate {flowrate} {mfc.readParameter(129).strip()} found at port {port}')
        else:
            print(f'MFC not found at port {port}')
            pass
    return mfcs


def start_program():
    """
    
    WIP...
    
    """

    pass


if __name__ =='__main__':
    mfc_port_connections = find_bronkhorst_ports()
    print(mfc_port_connections)
    #mfc_connections = {f'MFC_{capacity}': bh.Bronkhorst(connection) for 
    #                   capacity, connection in enumerate(mfc_port_connections.items())}