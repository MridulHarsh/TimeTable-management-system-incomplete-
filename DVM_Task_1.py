from openpyxl import Workbook, load_workbook
workbook = load_workbook(filename = "database.xlsx")
sheet = workbook.active

baka = 1
sudo = False
loop = False
chemlec = [] #done
eeelec = [] #done
thermolec = [] #done
phylec = [] #done
mathlec = [] #done
eglec = [] #done
chemtut = [] #done
eeetut = [] #done
thermotut = [] #done
phytut = [] #done
mathtut = [] #done
administrator = 'baka'
lect = ("CHEM (F111)", "EEE (F111)", "THERMO (BITS F111)", "PHY (F111)", "MATH (F111)", "EG (BITS F110)")
tut = ("CHEM (F111)", "EEE (F111)", "THERMO (BITS F111)", "PHY (F111)", "MATH (F111)")
lab = ("CHEM (F110)", "PHY (F110)", "EG (BITS F110)")
def display_menu2():
    print("A. Select the Lecture you want to enroll")
    for i in range(6):
        print(f"{i + 1}. {lect[i]}")
    print("\n")
    print("B. Select the Tutorial you want to enroll")
    for j in range(5):
        print(f"{j + 7}. {tut[j]}")
    print("\n")
    print("C. Select the Lab you want to enroll")
    for k in range(3):
        print(f"{k + 12}. {lab[k]}")
    print("\n")
    print("Type 0 to go to the previous menu")

def display_menu1():
    print("===== Command Menu =====")
    print("1. Enroll Subjects")
    print("2. Check Clashes")
    print("3. Export Time Table")
    print("4. Exit")
def execute_command1():
    global baka
    global choice
    if choice == '1':
        baka = 0
        display_menu2()
        choice = ''
    elif choice  == '2':
        pass
    elif choice == '3':
        print("Exporting Time Table...")

class Course():

    def __init__(self, choice):
        self.choice = choice
    def get_all_sections(self):
        if self.choice == '1':
            print("Here are the available Lectures")
            for row in sheet:
                if row[1].value != None:
                    chemlec.append(row[1].value)
            print(chemlec)
        elif self.choice == '2':
            print("Here are the available Lectures")
            for row in sheet:
                if row[5].value != None:
                    eeelec.append(row[5].value)
            print(eeelec)
        elif self.choice == '3':
            print("Here are the available Lectures")
            for row in sheet:
                if row[9].value != None:
                    thermolec.append(row[9].value)
            print(thermolec)
        elif self.choice == '4':
            print("Here are the available Lectures")
            for row in sheet:
                if row[13].value != None:
                    phylec.append(row[13].value)
            print(phylec)
        elif self.choice == '5':
            print("Here are the available Lectures")
            for row in sheet:
                if row[17].value != None:
                    mathlec.append(row[17].value)
            print(mathlec)
        elif self.choice == '6':
            print("Here are the available Lectures")
            for row in sheet:
                if row[21].value != None:
                    eglec.append(row[21].value)
            print(eglec)
        elif self.choice == '7':
            print("Here are the available Lectures")
            for row in sheet:
                if row[25].value != None:
                    chemtut.append(row[25].value)
            print(chemtut)
        elif self.choice == '8':
            print("Here are the available Lectures")
            for row in sheet:
                if row[29].value != None:
                    eeetut.append(row[29].value)
            print(eeetut)
        elif self.choice == '9':
            print("Here are the available Lectures")
            for row in sheet:
                if row[33].value != None:
                    thermotut.append(row[33].value)
            print(thermotut)
        elif self.choice == '10':
            print("Here are the available Lectures")
            for row in sheet:
                if row[37].value != None:
                    phytut.append(row[37].value)
            print(phytut)
        elif self.choice == '11':
            print("Here are the available Lectures")
            for row in sheet:
                if row[41].value != None:
                    mathtut.append(row[41].value)
            print(mathtut)
    def __str__(self):
        pass
    def populate_section(self):
        pass
class Timetable():
    def __init__(self):
        pass
    def enroll_subject(self):
        pass
    def check_clashes(self):
        pass
    def export_to_csv(self):
        pass

if __name__ == "__main__":
    zz = int(input("Welcome to Time Table Generator. Please select 1 option \n 1. Student \n 2. Administrator \n Select the desired option: "))
    if zz == 1:
        sudo = False
        loop = True
        display_menu1()
    elif zz == 2:
        password = str(input("Enter your Administrator Password: "))
        if password == administrator:
            loop = True
            sudo = True
            display_menu1()
        elif password != administrator:
            print("Not Authorised!")
    while loop:
        choice = str(input("Select the desired option: "))
        if baka == 1:
            if choice == '4':
                break
            elif choice == '1':
                execute_command1()
            else:
                display_menu1()
                execute_command1()
        if baka == 0:
            if choice == '0':
                baka = 1
                display_menu1()
            else:
                course = Course(choice)
                course.get_all_sections()