import gspread
from google.oauth2 import service_account
import regex as re
import time
from rich.console import Console
import pickle

console = Console()
tasks = [f"task {n}" for n in range(1, 3)]


""" Setting up an API and authorizing """
scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_file(
    "credentials.json")
scoped_credentials = credentials.with_scopes(scope)
client = gspread.authorize(scoped_credentials)

""" Setting up a worksheet """
sheet = client.open_by_url(
    'https://docs.google.com/spreadsheets/d'
    '/1pNHiTmrgEBjON2kh_KEzIsDIqPTKIkaQ6H0lr29lhYM/edit#gid=1647235932')  #
# Opening a spreadsheet by url
worksheetName = '3 курс Менеджмент'
worksheet = sheet.worksheet(
    worksheetName)  # Selecting a worksheet by its name

""" Main part of the code """
regex_numbers = re.compile(
    '([0-9][0-9]-апр.|[0-9]-апр.)')  # Creating a rule for finder
dates = worksheet.findall(
    regex_numbers)  # Listing all dates in the worksheet using our rule
regex_time = re.compile(
    '[0-9][0-9].[0-9][0-9]-[0-9][0-9].[0-9][0-9]')  # Creating a second rule
# for finder
lessonTimeList = worksheet.findall(
    regex_time)  # Listing all lesson times in the worksheet using our rule
startDate = dates[0].value[:-5]  # First day number
endDate = dates[-1].value[:-5]  # Last day number

firstDay = dates[0].row  # Getting a first day row number
secondDay = dates[1].row  # Getting a second day row number
thirdDay = dates[2].row  # Getting a third day row number
fourthDay = dates[3].row  # Getting a fourth day row number
fifthDay = dates[4].row  # Getting a fifth day row number
sixthDay = dates[5].row  # Getting a sixth day row number

group_number = gn = "МБ{}".format(int(input("Введите номер вашей группы: ")))
print("Ваша группа — {}".format(gn))
secondLang_dict = {
    "Луканина": "Луканина Екатерина Викторовна",
    "Кошкина": "Кошкина Марина Олеговна",
    "Березовская": "Березовская Татьяна Евгеньевна",
    "Абрамова": "Абрамова Надежда Владимировна",
    "Романенко": "Романенко Екатерина Евгеньевна",
    "Воинова": "Воинова Ольга Зиновьевна",
    "Кондакова": "Кондакова Лейла Васильевна",
    "Немировская": "Немировская Александра Дмитриевна",
    "Круглова": "Круглова Анастасия Владимировна",
    "Волкевич": "Волкевич Анастасия Юрьевна",
}
secondLangName_input = sln = input(
    "Введите фамилию преподавателя по второму иностранному языку: ")
secondLangName = sln = secondLang_dict[sln]
print(f"Имя вашего преподавателя — {sln}")

""" First Day """
row = firstDay
firstDay_list = []  # List of info about lessons on Monday
firstDay_lessons = []  # Lessons on first day
firstDay_lessons.clear()
while row < secondDay - 1:
    row_value = worksheet.row_values(row)
    if row_value[3] == gn:
        firstDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(firstDay_lessons):
            del firstDay_lessons[i][0:2]
            filtered = filter(lambda nonempty: nonempty != ''
                              and nonempty != ' '
                              and nonempty != gn
                              and nonempty != '1',
                              firstDay_lessons[t])
            t = t + 1
            i = i + 1
        row = row + 1
        firstDay_list.append(list(filtered))
    else:
        row = row + 1

# Checking if any lessons exist
if len(firstDay_list) > 0:
    firstDay_list.insert(0, dates[0].value)
else:
    firstDay_list.insert(0, "None")
print(firstDay_list)

""" Second Day """
with console.status("[bold green]API timeout...") as status:
    task = tasks.pop(0)
    time.sleep(10)
    console.log("50 seconds left")
    time.sleep(10)
    console.log("40 seconds left")
    time.sleep(10)
    console.log("30 seconds left")
    time.sleep(10)
    console.log("20 seconds left")
    time.sleep(10)
    console.log("10 seconds left")
    time.sleep(10)
    console.log("Done!")
row = secondDay
secondDay_list = []  # List of info about lessons on second day
secondDay_lessons = []  # Lessons on second day
secondDay_lessons.clear()
while row < thirdDay - 1:
    row_value = worksheet.row_values(row)
    if row_value[3] == gn or row_value[3] == "1":
        secondDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(secondDay_lessons):
            del secondDay_lessons[i][0:2]
            filtered = filter(lambda nonempty:
                              nonempty != ''
                              and nonempty != ' '
                              and nonempty != gn
                              and nonempty != '1', secondDay_lessons[t])
            t = t + 1
            i = i + 1
        row = row + 1
        secondDay_list.append(list(filtered))
    else:
        row = row + 1

# Checking if any lessons exist
if len(secondDay_list) > 0:
    secondDay_list.insert(0, dates[1].value)
else:
    secondDay_list.insert(0, "None")
print(secondDay_list)

""" Third Day """
# time.sleep(60)  # Reads per minute timeout for Sheets API
with console.status("[bold green]API timeout...") as status:
    task = tasks.pop(0)
    time.sleep(10)
    console.log("50 seconds left")
    time.sleep(10)
    console.log("40 seconds left")
    time.sleep(10)
    console.log("30 seconds left")
    time.sleep(10)
    console.log("20 seconds left")
    time.sleep(10)
    console.log("10 seconds left")
    time.sleep(10)
    console.log("Done!")

row = thirdDay
thirdDay_list = []  # List of info about lessons on third day
thirdDay_lessons = []  # Lessons on third day
thirdDay_lessons.clear()
while row < fourthDay - 1:
    row_value = worksheet.row_values(row)
    if row_value[3] == gn or row_value[3] == "1" \
            or row_value[4] == "Физическая культура и спорт ":
        thirdDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(thirdDay_lessons):
            del thirdDay_lessons[i][0:2]
            filtered = filter(lambda nonempty:
                              nonempty != ''
                              and nonempty != ' '
                              and nonempty != gn
                              and nonempty != '1',
                              thirdDay_lessons[t])
            t = t + 1
            i = i + 1
        row = row + 1
        thirdDay_list.append(list(filtered))
    elif row_value[5] == "Луканина Екатерина Викторовна":
        spanish_time = row_value[2]
        row = row + 1
    elif row_value[5] == sln:
        thirdDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(thirdDay_lessons):
            del thirdDay_lessons[i][0:2]
            filtered = filter(lambda nonempty:
                              nonempty != ''
                              and nonempty != ' ', thirdDay_lessons[t])
            thirdDay_lessons[t].insert(0, spanish_time)
            t = t + 1
            i = i + 1
        row = row + 1
        thirdDay_list.append(list(filtered))
    else:
        row = row + 1
# Checking if any lessons exist
if len(thirdDay_list) > 0:
    thirdDay_list.insert(0, dates[2].value)
else:
    thirdDay_list.insert(0, "None")
print(thirdDay_list)

""" Fourth Day """
with console.status("[bold green]API timeout...") as status:
    task = tasks.pop(0)
    time.sleep(10)
    console.log("50 seconds left")
    time.sleep(10)
    console.log("40 seconds left")
    time.sleep(10)
    console.log("30 seconds left")
    time.sleep(10)
    console.log("20 seconds left")
    time.sleep(10)
    console.log("10 seconds left")
    time.sleep(10)
    console.log("Done!")
row = fourthDay
fourthDay_list = []  # List of info about lessons on fourth day
fourthDay_lessons = []  # Lessons on fourth day
fourthDay_lessons.clear()
while row < fifthDay - 1:
    row_value = worksheet.row_values(row)
if row_value[3] == gn or row_value[3] == "1":
    fourthDay_lessons.append(row_value)
    i = 0
    t = 0
    while i < len(fourthDay_lessons):
        del fourthDay_lessons[i][0:2]
        filtered = filter(lambda nonempty:
                          nonempty != ''
                          and nonempty != ' '
                          and nonempty != gn
                          and nonempty != '1', fourthDay_lessons[t])
        t = t + 1
        i = i + 1
    row = row + 1
    fourthDay_list.append(list(filtered))
else:
    row = row + 1

# Checking if any lessons exist
if len(fourthDay_list) > 0:
    fourthDay_list.insert(0, dates[3].value)
else:
    fourthDay_list.insert(0, "None")
print(fourthDay_list)

""" Fifth Day """
# time.sleep(60)  # Reads per minute timeout for Sheets API
with console.status("[bold green]API timeout...") as status:
    task = tasks.pop(0)
    time.sleep(10)
    console.log("50 seconds left")
    time.sleep(10)
    console.log("40 seconds left")
    time.sleep(10)
    console.log("30 seconds left")
    time.sleep(10)
    console.log("20 seconds left")
    time.sleep(10)
    console.log("10 seconds left")
    time.sleep(10)
    console.log("Done!")

row = fifthDay
fifthDay_list = []  # List of info about lessons on fifth day
fifthDay_lessons = []  # Lessons on fifth day
fifthDay_lessons.clear()
while row < sixthDay - 1:
    row_value = worksheet.row_values(row)
    if row_value[3] == gn or row_value[3] == "1" \
            or row_value[4] == "Физическая культура и спорт ":
        fifthDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(fifthDay_lessons):
            del fifthDay_lessons[i][0:2]
            filtered = filter(lambda nonempty:
                              nonempty != ''
                              and nonempty != ' '
                              and nonempty != gn
                              and nonempty != '1', fifthDay_lessons[t])
            t = t + 1
            i = i + 1
        row = row + 1
        fifthDay_list.append(list(filtered))
    elif row_value[5] == "Луканина Екатерина Викторовна":
        spanish_time = row_value[2]
        row = row + 1
    elif row_value[5] == sln:
        fifthDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(fifthDay_lessons):
            del fifthDay_lessons[i][0:2]
            filtered = filter(lambda nonempty:
                              nonempty != '' and
                              nonempty != ' ', fifthDay_lessons[t])
            fifthDay_lessons[t].insert(0, spanish_time)
            t = t + 1
            i = i + 1
        row = row + 1
        fifthDay_list.append(list(filtered))
    else:
        row = row + 1

# Checking if any lessons exist
if len(fifthDay_list) > 0:
    fifthDay_list.insert(0, dates[4].value)
else:
    fifthDay_list.insert(0, "None")
print(fifthDay_list)

""" Sixth Day """
row = sixthDay
sixthDay_list = []  # List of info about lessons on sixth day
sixthDay_lessons = []  # Lessons on sixth day
sixthDay_lessons.clear()
while row <= lessonTimeList[-1].row:
    row_value = worksheet.row_values(row)
    if row_value[3] == gn or row_value[3] == "1" \
            or row_value[4] == "Физическая культура и спорт ":
        sixthDay_lessons.append(row_value)
        i = 0
        t = 0
        while i < len(sixthDay_lessons):
            del sixthDay_lessons[i][0:2]
            filtered = filter(lambda nonempty:
                              nonempty != '' and nonempty != ' '
                              and nonempty != gn
                              and nonempty != '1', sixthDay_lessons[t])
            t = t + 1
            i = i + 1
        row = row + 1
        sixthDay_list.append(list(filtered))
    else:
        row = row + 1

# Checking if any lessons exist
if len(sixthDay_list) > 0:
    sixthDay_list.insert(0, dates[5].value)
else:
    sixthDay_list.insert(0, "None")
print(sixthDay_list)

""" Pickle """
data = [firstDay_list, secondDay_list, thirdDay_list, fourthDay_list,
        fifthDay_list, sixthDay_list]

with open('data.pkl', 'wb') as f:
    pickle.dump(data, f)
