#***********************
#
# Copyright (c) 2021-2023, Operation Technology, Inc.
#
# THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO OPERATION TECHNOLOGY, INC. 
# ANY USE OF THIS PROGRAM IS SUBJECT TO THE PROGRAM SOFTWARE LICENSE AGREEMENT, 
# EXCEPT THAT THE USER MAY MODIFY THE PROGRAM FOR ITS OWN USE. 
# HOWEVER, THE PROGRAM MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS 
# WITHOUT THE PRIOR WRITTEN CONSENT OF OPERATION TECHNOLOGY, INC.
#
#***********************

import sqlite3 
from datetime import datetime 


def connect_to_project(project_path):
    # ....Connecting to SQLite DB....
    project = sqlite3.connect(project_path)
    cursor = project.cursor()

    # ....Get all the table names from report Database....
    all_tables_name = project.execute("SELECT name FROM sqlite_master WHERE type = 'table';")
    table_names_list = []
    for table_names in all_tables_name:
        table_names_list.append(table_names[0])

    return cursor, table_names_list


def get_updated_coordinates(coordinates):
    second_emp_index = [i for i, ltr in enumerate(coordinates) if ltr == '$'][1]
    int_part = int(coordinates[second_emp_index + 1:]) + 1
    coordinates = coordinates[:second_emp_index + 1] + str(int_part)
    return coordinates


def get_actual_day(int_val):
    int_to_day = {0: "Mon.", 1: "Tue.", 2: "Wed.", 3: "Thurs.", 4: "Fri.", 5: "Sat.", 6: "Sun."}
    value_day = int_to_day[int_val]
    return value_day

def increase_column_by_one(coordinate):
     first_emp_index = [i for i, ltr in enumerate(coordinate) if ltr == '$'][0]
     second_emp_index = [i for i, ltr in enumerate(coordinate) if ltr == '$'][1]
     column_alphabet = coordinate[first_emp_index + 1:second_emp_index]
     if len(column_alphabet) == 1:
         if column_alphabet == 'Z':
             start_writing_column = 'AA'
         else:
             start_writing_column = chr(ord(column_alphabet)+1)

     elif len(column_alphabet) == 2:
         if column_alphabet == 'ZZ':
             start_writing_column = 'AAA'
         else:
             first_char_of_column_alphabet = column_alphabet[:len(column_alphabet)-1]
             second_char_of_column_alphabet = column_alphabet[len(column_alphabet)-1:]
             start_writing_column = first_char_of_column_alphabet + chr(ord(second_char_of_column_alphabet)+1)

     start_writing_coordinate = coordinate[:first_emp_index+1] + start_writing_column + coordinate[second_emp_index:]
     return start_writing_coordinate

def date_and_time_from_timeStamp(time_stamp):
     date_value = time_stamp[:time_stamp.find(' ')]
     time_value = time_stamp[time_stamp.find(' ')+1:]
     date_final = datetime.strptime(date_value, '%m-%d-%Y').date()
     return date_final, time_value

