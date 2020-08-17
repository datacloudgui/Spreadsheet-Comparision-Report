import openpyxl
from openpyxl import Workbook

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

def attended_split(list_students, attendance_students):
    #Use the oficial list to split attendance list in two list:
    #attended in list (attended_students) and attended out of list (attended_not_in_list)

    attended_students = []
    attended_not_in_list = attendance_students

    for student in list_students:
        #Split the name in official list
        words_name = student.lower().split(' ')
        word_find_count = 0
        for student_attended in attendance_students:
            #Iterate over attendance list searching the splited name
            for word in words_name:
                #Search in the current attendance item for words in the name splited
                if(student_attended.lower().find(word) != -1):
                    word_find_count += 1
                    if(word_find_count == 2):
                        #Store attended official students if at least 2 words are founded
                        # replace 2 with len(words_name) to verify all the words in the official list
                        # at the end of the process attended_not_in_list will contain the remained names 
                        #for people that attend the class but aren't in the official list.
                        attended_students.append(student)
                        attended_not_in_list.remove(student_attended)
                        break
    return attended_students, attended_not_in_list

def extract_absence_students(list_students, attended_students):
    absence_students = list_students

    for item in attended_students:
        absence_students.remove(item)
    return absence_students

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

    attended_students, attended_not_in_list = attended_split(list_students, attendance_students)
    absence_students = extract_absence_students(list_students, attended_students)
    #print('Lista de estudiantes')
    #for item in attendance_students:
    #    print(item)
    #print('Lista de estudiantes')
    #for item in list_students:
    #    print(item)

    print('Asistentes en lista:')
    for item in attended_students:
        print(item)
    print('----------------------------------')

    print('Asistentes NO en lista')
    for item in attended_not_in_list:        
        print(item)

    print('----------------------------------')

    print('Ausentes')
    for item in absence_students:        
        print(item)

    book = Workbook()
    sheet = book.active
    sheet.title = 'servos'
    column = 'A'
    sheet['A1']='Asistentes en lista'
    for row in range (0, len(attended_students)):
        sheet[column + str(row+2)] = attended_students[row]

    book.save("report.xlsx")