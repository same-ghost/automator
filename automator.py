# IMPORT ----------------------------------------------------------------------
import openpyxl
import time


# GLOBAL ----------------------------------------------------------------------
FILE_LOCATION = '/mnt/c/Users/talbert/Downloads/linux/'


# FUNCTION --------------------------------------------------------------------
def source_sheet():
    '''
        Open course listings workbook and return the active sheet.
    '''
    wb = openpyxl.load_workbook(FILE_LOCATION + 'datasource.xlsx')
    return wb['Degree Checklist - Course List ']
# END source_sheet()


def target_sheet():
    '''
        Open student listings workbook and return the active sheet.
    '''
    wb = openpyxl.load_workbook(FILE_LOCATION + 'skeleton.xlsx')
    return wb['Automator Student List']
# END target_sheet()


def validator(grade):
    '''
        Check a single course grade against the list of valid grades.
    '''
    valid_grades = ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 
    'D+', 'D', 'D-', None]
    return grade in valid_grades
# END validator()


def format_term(course_term):
    '''
        Returns the term a course was taken in the format 
        FA ##/SP ##/SU ##.
    '''
    terms = {'D1': 'FA', 'D2': 'SP', 'D4': 'SU', '  ': 'EQUIV'}
    term = terms[course_term[4:6]] + ' ' + course_term[2:4]
    return term
# END format_term()


# MAIN ------------------------------------------------------------------------
source_sheet = source_sheet()
target_book = openpyxl.load_workbook(FILE_LOCATION + 'skeleton.xlsx')
target_sheet = target_book['Automator Student List']

print("Test begins: ")
t0 = time.process_time()

# Iterate through all students in Automator Student List
for student in target_sheet.iter_rows(min_row=2):
    target_id = student[0].value

    # Iterate through all courses in Automator Course List
    for course in source_sheet.iter_rows(min_row=2):
        
        course_grade = course[5].value
        if validator(course_grade):

            course_id = course[0].value
            course_code = course[2].value
            course_term = course[1].value
            course_ger = course[6].value

            if target_id == course_id:
                if course_ger == 'FYW':
                    student[10].value = course_code
                    student[11].value = course_term
                    student[12].value = course_grade

                if course_ger == 'WR':
                    student[13].value = course_code
                    student[14].value = course_term
                    student[15].value = course_grade

                if course_ger == 'NWL':
                    student[16].value = course_code
                    student[17].value = course_term
                    student[18].value = course_grade

                if course_ger == 'HB':
                    student[22].value = course_code
                    student[23].value = course_term
                    student[24].value = course_grade

                if course_ger == 'TA':
                    student[28].value = course_code
                    student[29].value = course_term
                    student[30].value = course_grade

                if course_ger == 'HA':
                    student[31].value = course_code
                    student[32].value = course_term
                    student[33].value = course_grade

                if course_ger == 'VP':
                    student[34].value = course_code
                    student[35].value = course_term
                    student[36].value = course_grade

                if course_ger == 'MR':
                    student[37].value = course_code
                    student[38].value = course_term
                    student[39].value = course_grade

                if course_ger == 'FL':
                    student[40].value = course_code
                    student[41].value = course_term
                    student[42].value = course_grade

                if course_ger == 'UQ':
                    student[43].value = course_code
                    student[44].value = course_term
                    student[45].value = course_grade

                if course_ger == 'MB':
                    student[46].value = course_code
                    student[47].value = course_term
                    student[48].value = course_grade

                if course_ger == 'NE':
                    student[49].value = course_code
                    student[50].value = course_term
                    student[51].value = course_grade

                if course_ger == 'WC':
                    student[52].value = course_code
                    student[53].value = course_term
                    student[54].value = course_grade

# Print test results
t1 = time.process_time()
print(f"Test completed in {t1-t0} seconds.")

# Save workbook
target_book.save(FILE_LOCATION + 'targeted.xlsx')