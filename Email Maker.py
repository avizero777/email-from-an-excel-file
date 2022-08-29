def email(excel_file):
    from openpyxl import load_workbook
    book = load_workbook(excel_file)
    sh = book.active
    class person:
        def __init__(self, name, surname, gender):
            self.name = name
            self.surname = surname
            self.gender = gender
    for i in range(2, sh.max_row+1):
        print("\n")
        for j in range(1, sh.max_column):
            cell_obj = sh.cell(row=i, column=j)
            if j == 1 :
                first_name = cell_obj.value
            elif j == 2 :
                nick_name = cell_obj.value
            else:
                gnd = cell_obj.value
        data = person(first_name, nick_name, gnd)
        emails_file = open('emails', 'a')
        if gnd == 'female':
            emails_file.write('\n' + 'Miss.' + data.name +'.'+ data.surname + '@email.com')
        elif gnd == 'male':
            emails_file.write('\n' + 'Mr.' + data.name +'.'+ data.surname + '@email.com')
        emails_file.close
    

try:
    email(r'   .xlsx') # <===  """"enter your excel file with the absolute path """"
except:
    print('FileNotFoundError')