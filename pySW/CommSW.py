"""
# *****************************************************************************
# *                                                                           *
# *    Copyright (C) 2020  Kalyan Inamdar, kalyaninamdar@protonmail.com       *
# *                                                                           *
# *    This library is free software; you can redistribute it and/or          *
# *    modify it under the terms of the GNU Lesser General Public             *
# *    License as published by the Free Software Foundation; either           *
# *    version 2 of the License, or (at your option) any later version        *
# *                                                                           *
# *    This library is distributed in the hope that it will be useful,        *
# *    but WITHOUT ANY WARRANTY; without even the implied warranty of         *
# *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          *
# *    GNU Lesser General Public License for more details.                    *
# *                                                                           *
# *    You should have received a copy of the GNU General Public License      *
# *    along with this program.  If not, see http://www.gnu.org/licenses/.    *
# *                                                                           *
# *****************************************************************************
"""
import subprocess as sb
import win32com.client
import pythoncom
import os


#
class CommSW:
    def __init__(self):
        self.part_name_inn = ""
        self.unequal_lengths_error_message = "If a list of multiple variables is given, then lists of equal \n\
        lengths should be given for 'modified_val' and 'unit' inputs."

    #
    @staticmethod
    def start_sw(*args):
        #                                                                     #
        # Function to start Solidworks from Python.                           #
        #                                                                     #
        # Accepts an optional argument: the year of version of Solidworks.    #
        #                                                                     #
        # If you have only one version of Solidworks on your computer, you do #
        # not need to provide this input.                                     #
        #                                                                     #
        # Example: If you have Solidworks 2019 and Solidworks 2020 on your    #
        # system and you want to start Solidworks 2020 the function you call  #
        # should look like this: startSW(2020)                                #
        #                                                                     #
        if not args:
            sw_process_name = r'C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/SLDWORKS.exe'
            sb.Popen(sw_process_name)
        else:
            year = int(args[0][-1])
            sw_process_name = "SldWorks.Application.%d" % (20 + (year - 2))
            win32com.client.Dispatch(sw_process_name)

    #
    @staticmethod
    def shut_sw():
        #                                                                     #
        # Function to close Solidworks from Python.                           #
        # Does not accept any input.                                          #
        #                                                                     #
        sb.call('Taskkill /IM SLDWORKS.exe /F')

    #
    @staticmethod
    def connect_to_sw():
        #                                                                     #
        # Function to establish a connection to Solidworks from Python.       #
        # Does not accept any input.                                          #
        #                                                                     #
        global swcom
        swcom = win32com.client.Dispatch("SLDWORKS.Application")

    #
    def open_assembly(self, part_name_input):
        #                                                                     #
        # Function to open an assembly document in Solidworks from Python.    #
        #                                                                     #
        # Accepts one input as the filename with the path if the working      #
        # directory of your script and the directory in which the assembly    #
        # file is saved are different.                                        #
        #                                                                     #
        self.part_name_inn = part_name_input
        self.part_name_inn = self.part_name_inn.replace('\\', '/')
        #
        if os.path.basename(self.part_name_inn).split('.')[-1].lower() == 'sldasm':
            print("Opening Assembly: " + self.part_name_inn)
        else:
            self.part_name_inn = self.part_name_inn + '.SLDASM'
        #
        open_doc = swcom.OpenDoc6
        arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, self.part_name_inn)
        arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 2)
        arg3 = win32com.client.VARIANT(pythoncom.VT_I4, 0)
        arg5 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
        arg6 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
        #
        open_doc(arg1, arg2, arg3, "", arg5, arg6)

    #
    def open_part(self, part_name_input):
        #                                                                     #
        # Function to open an part document in Solidworks from Python.        #
        #                                                                     #
        # Accepts one input as the filename with the path if the working      #
        # directory of your script and the directory in which the part file   #
        # is saved are different.                                             #
        #                                                                     #
        self.part_name_inn = part_name_input
        self.part_name_inn = self.part_name_inn.replace('\\', '/')
        #
        open_doc = swcom.OpenDoc6
        arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, self.part_name_inn)
        arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 1)
        arg3 = win32com.client.VARIANT(pythoncom.VT_I4, 1)
        arg5 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
        arg6 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
        #
        open_doc(arg1, arg2, arg3, "", arg5, arg6)

    #
    @staticmethod
    def update_part():
        if 'model' in globals():
            pass
        else:
            global model
            model = swcom.ActiveDoc
        model.EditRebuild3

    #
    def close_part(self):
        swcom.CloseDoc(os.path.basename(self.prtName))

    #
    @staticmethod
    def save_assembly(directory, file_name, file_extension):
        if 'model' in globals():
            pass
        else:
            global model
            model = swcom.ActiveDoc
        directory = directory.replace('\\', '/')
        com_file_name = directory + '/' + file_name + '.' + file_extension
        arg = win32com.client.VARIANT(pythoncom.VT_BSTR, com_file_name)
        model.SaveAs3(arg, 0, 0)

    #
    @staticmethod
    def get_global_variables():
        #                                                                     #
        # Function to extract a set of global variables in a Solidworks       #
        # part/assembly file. The part/assembly is then automatically updated.#
        #                                                                     #
        # Does not accept any input.                                          #
        # Provides output in the form of a dictionary with Global Variables as#
        # the keys and values of the variables as values in the dictionary.   #
        #                                                                     #
        if 'model' in globals():
            print("Using 'model' from globals...")
        else:
            global model
            model = swcom.ActiveDoc
        #
        if 'eqMgr' in globals():
            print("Using 'eqMgr' from globals...")
        else:
            global eqMgr
            eqMgr = model.GetEquationMgr
        #
        n = eqMgr.getCount
        #
        data = {}
        #
        for i in range(n):
            if eqMgr.GlobalVariable(i):
                data[eqMgr.Equation(i).split('"')[1]] = i
            #
        #
        if len(data.keys()) == 0:
            raise KeyError("There are not any 'Global Variables' present in the currently active Solidworks document.")
        else:
            return data

    #
    def modify_global_var(self, variable, modified_val, unit):
        #                                                                     #
        # Function to modify a global variable or a set of global variables   #
        # in a Solidworks part/assembly file. The part/assembly is then       #
        # automatically updated.                                              #
        #                                                                     #
        # Accepts three inputs: variable name, modified value, and the unit   #
        # of the variable. The inputs can be string, integer and string       #
        # respectively or a list of variables, list of modified values and a  #
        # list of units of respective variables.                              #
        #                                                                     #
        # Note: In case you need to modify multiple dimensions using lists    #
        # the length of the lists must strictly be equal.                     #
        #                                                                     #
        if 'model' in globals():
            print("Using 'model' from globals...")
        else:
            global model
            model = swcom.ActiveDoc
        #
        if 'eqMgr' in globals():
            print("Using 'eqMgr' from globals...")
        else:
            global eqMgr
            eqMgr = model.GetEquationMgr
        #
        data = self.get_global_variables()
        #
        if isinstance(variable, str):
            eqMgr.Equation(data[variable], "\"" + variable + "\" = " + str(modified_val) + unit + "")
        elif isinstance(variable, list):
            if isinstance(modified_val, list):
                if isinstance(unit, list):
                    for i in range(len(variable)):
                        eqMgr.Equation(data[variable[i]],
                                       "\"" + variable[i] + "\" = " + str(modified_val[i]) + unit[i] + "")
                else:
                    raise TypeError(self.unequal_lengths_error_message)
            else:
                raise TypeError(self.unequal_lengths_error_message)
        else:
            raise TypeError("Incorrect input for the variables. \n\
            Inputs can either be string, \n\
            integer and string or lists containing variables, values and units.")
        #
        self.update_part()

    #
    def modify_linked_var(self, variable, modified_val, unit, *args):
        #                                                                     #
        # Function to modify a global variable/dimension or a set of          #
        # dimensions in a linked 'equations' file. The part/assembly is then  #
        # automatically updated.                                              #
        #                                                                     #
        # Accepts three inputs: variable name, modified value, and the unit   #
        # of the variable. The inputs can be string, integer and string       #
        # respectively or a list of variables, list of modified values and a  #
        # list of units of respective variables. Additionally the function    #
        # accepts one more optional argument, which is the complete path of   #
        # the equations file. If a path to the equations file is not provided #
        # then the function searches for a file named 'equations.txt' in the  #
        # working directory of the code.                                     #
        #                                                                     #
        # Note: In case you need to modify multiple dimensions using lists    #
        # the length of the lists must strictly be equal.                     #
        #                                                                     #
        #
        # Check the filename
        if len(args) == 0:
            file = 'equations.txt'
        else:
            file = args[0]
        #
        # READ FILE WITH ORIGINAL DIMENSIONS
        try:
            reader = open(file, 'r')
        except IOError:
            raise IOError
        finally:
            data = {}
            num_lines = len(reader.readlines())
            reader.close()
            reader = open(file)
            lines = reader.readlines()
            reader.close()
            for i in range(num_lines):
                dim = lines[i].split('"')[1]
                temp_val = lines[i].split(' ')[1]
                #
                val = temp_val.replace(unit, '').replace('= ', '').replace('\n', '')
                data[dim] = val
        #
        # MODIFY DIMENSIONS
        if isinstance(variable, list):
            if isinstance(modified_val, list):
                if isinstance(unit, list):
                    for z in range(len(variable)):
                        data[variable[i]] = modified_val[i]
                else:
                    raise TypeError(self.unequal_lengths_error_message)
            else:
                raise TypeError(self.unequal_lengths_error_message)
        elif isinstance(variable, str):
            data[variable] = modified_val
        else:
            raise TypeError("The inputs types given.")
        #
        # WRITE FILE WITH MODIFIED DIMENSIONS
        writer = open(file, 'w')
        for key, value in data.items():
            writer.write('"' + key + '"= ' + str(value) + unit)
        writer.close()
        #
        self.update_part()
    #
