import openpyxl

def open_excel_sheet(filename, sheet_name):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.get_sheet_by_name(sheet_name)
    return sheet

def find_word_in_row(sheet, column_name, word):
    for row in range(1, sheet.max_row +1):
        if(sheet[column_name + str(row)].value == word):
            print('Inicio de la lista detectado en la fila {} del archivo'.format(row))
            list_detected=row
            break
    return list_detected

def find_word_in_column(sheet, row_index, word):
    for col in sheet.iter_cols(min_row=row_index, max_row =row_index):
        for cell in col:
            if(cell.value == word):
                print('Inicio de la lista detectado en la columna {} del archivo'.format(cell.coordinate))
                list_detected=cell.coordinate
                break
    return list_detected

def extract_list_by_index(sheet, row_start, column):
    students=[]
    for row in range (row_start+1, sheet.max_row+1):
        if(sheet[column + str(row)].value!= None):
            students.append(sheet[column + str(row)].value)
    return students



if __name__ == "__main__":
    list_file='servos.xlsx'
    list_sheet_name='Sheet1'
    list_column='B'
    list_word='Apellidos y Nombres'

    attendance_file='Meet.xlsx'
    attendance_sheet_name='Attendance'

    list_sheet = open_excel_sheet(list_file,list_sheet_name)
    list_row_index = find_word_in_row(list_sheet, list_column, list_word)
    list_students = extract_list_by_index(list_sheet, list_row_index, list_column)

    attendance_sheet=open_excel_sheet(attendance_file,attendance_sheet_name)
    attendance_column_index = find_word_in_column(attendance_sheet,1,'servos')
    attendance_students = extract_list_by_index(attendance_sheet, 2, attendance_column_index[0:-1])

    print('Lista de estudiantes')
    for item in attendance_students:
        print(item)
    print('Lista de estudiantes')
    for item in list_students:
        print(item)