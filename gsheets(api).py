import gspread
from google.oauth2 import service_account
import regex as re
import time
from rich.console import Console
import pickle
import logging

################################################################################
# DEBUG: Detailed information, typically of interest only when diagnosing
# problems.
#
# INFO: Confirmation that things are working as expected.
#
# WARNING: An indication that something unexpected happened, or indicative of
# some issue in the near future (e.g. ‘disk space low’). The software is
# still working as expected.
#
# ERROR: Due to a more serious issue, the software has not been able to
# perform some function.
#
# CRITICAL: A serious error, indicating that the program itself may be unable
# to continue running.
################################################################################

""" API Timeout """
console = Console()


def api_timeout(sec):
    seconds = [f"{n}" for n in range(1, int(sec) // 10)]
    with console.status("[bold blue]API timeout[blink]..."):
        console.log(f"[bold blue]API timeout {sec} seconds...")
        while seconds:
            s = seconds.pop(0)
            time.sleep(10)
            console.log(
                f"[blue]{int(sec) - int(s) * 10} [black]seconds "
                f"left!")
        console.log("[bold blue]Done!")


def api_group_timeout(sec):
    seconds = [f"{n}" for n in range(1, int(sec) // 10)]
    console.log(f"[bold blue]API timeout {sec} seconds...")
    while seconds:
        s = seconds.pop(0)
        time.sleep(10)
        console.log(
            f"[blue]{int(sec) - int(s) * 10} [black]seconds "
            f"left!")
    console.log("[bold blue]Done!")


""" Setting up a logging module """
logging.basicConfig(filename='log.log', filemode='w', level=logging.DEBUG,
                    format='%(asctime)s:%(levelname)s:%(message)s')
logging.basicConfig(filename='errors.log', filemode='w', level=logging.ERROR,
                    format='%(asctime)s:%(levelname)s:%(message)s')
debug = logging.debug
info = logging.info

""" Setting up an API and authorizing """
scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_file(
    "credentials.json")  # Service credentials for your Google account API
scoped_credentials = credentials.with_scopes(scope)
client = gspread.authorize(scoped_credentials)

""" Setting up a worksheet """
sheet = client.open_by_url(
    'https://docs.google.com/spreadsheets/d/1_N-zgQqld-hZZ-ZybK-owxTLpXXfjwZHsw7TY8BtJjo/edit#gid=1954731353')  #
# Opening a spreadsheet by url
worksheetName = '4 курс менеджмент'
worksheet = sheet.worksheet(
    worksheetName)  # Selecting a worksheet by its name

""" Main part of the code """
regex_numbers = re.compile(
    f'([0-9][0-9]-сент.|[0-9]-сент.|[0-9][0-9]-окт.|[0-9]-окт.)')  # Creating a rule for finder
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
info("Ваша группа — {}".format(gn))
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
sln = secondLang_dict[sln]
info(f"Имя вашего преподавателя — {sln}")

""" First Day """
row = firstDay
firstDay_list = []  # List of info about lessons on Monday
firstDay_lessons = []  # Lessons on first day
firstDay_lessons.clear()
with console.status("[bold green]Понедельник",
                    spinner="aesthetic"):
    while row < secondDay - 1:
        row_value = worksheet.row_values(row)
        if row_value[3] == gn or any("лек." in s for s in row_value) and \
                row_value[3] == ' ':
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
debug(firstDay_list)
info("Понедельник посчитан")

""" Second Day """
api_timeout(60)
row = secondDay
secondDay_list = []  # List of info about lessons on second day
secondDay_lessons = []  # Lessons on second day
secondDay_lessons.clear()
with console.status("[bold green]Вторник",
                    spinner="aesthetic"):
    while row < thirdDay - 1:
        row_value = worksheet.row_values(row)
        if row_value[3] == gn or "Губанова Елена Евгеньевна" in row_value or any("лек." in s for s in row_value) and \
                row_value[3] == ' ':
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
debug(secondDay_list)
info("Вторник посчитан")

""" Third Day """
api_timeout(60)
row = thirdDay
thirdDay_list = []  # List of info about lessons on third day
thirdDay_lessons = []  # Lessons on third day
thirdDay_lessons.clear()
with console.status("[bold green]Среда",
                    spinner="aesthetic"):
    while row < fourthDay - 1:
        row_value = worksheet.row_values(row)
        if row_value[3] == gn or row_value[3] == "1" or any("лек." in s for s in row_value) and \
                row_value[3] == ' ':
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
            api_group_timeout(60)
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
debug(thirdDay_list)
info("Среда посчитана")

""" Fourth Day """
api_timeout(60)
row = fourthDay
fourthDay_list = []  # List of info about lessons on fourth day
fourthDay_lessons = []  # Lessons on fourth day
fourthDay_lessons.clear()
with console.status("[bold green]Четверг",
                    spinner="aesthetic"):
    while row < fifthDay - 1:
        row_value = worksheet.row_values(row)
        if row_value[3] == gn or "Губанова Елена Евгеньевна" in row_value or any("лек." in s for s in row_value) and \
                row_value[3] == ' ':
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
debug(fourthDay_list)
info("Четверг посчитан")

""" Fifth Day """
api_timeout(60)
row = fifthDay
fifthDay_list = []  # List of info about lessons on fifth day
fifthDay_lessons = []  # Lessons on fifth day
fifthDay_lessons.clear()
with console.status("[bold green]Пятница",
                    spinner="aesthetic"):
    while row < sixthDay - 1:
        row_value = worksheet.row_values(row)
        if row_value[3] == gn or any("лек." in s for s in row_value) and \
                row_value[3] == ' ':
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
            api_group_timeout(60)
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
debug(fifthDay_list)
info("Пятница посчитана")

""" Sixth Day """
api_timeout(60)
row = sixthDay
sixthDay_list = []  # List of info about lessons on sixth day
sixthDay_lessons = []  # Lessons on sixth day
sixthDay_lessons.clear()
with console.status("[bold green]Суббота",
                    spinner="aesthetic"):
    while row <= lessonTimeList[-1].row:
        row_value = worksheet.row_values(row)
        if row_value[3] == gn or "Губанова Елена Евгеньевна" in row_value or any("лек." in s for s in row_value) and \
                row_value[3] == ' ':
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
debug(sixthDay_list)
info("Суббота посчитана")

""" Pickle """
data = [firstDay_list, secondDay_list, thirdDay_list, fourthDay_list,
        fifthDay_list, sixthDay_list]

with open('data.pkl', 'wb') as f:
    pickle.dump(data, f)

console.log("[bold green] Process finished.")

if __name__ == '__main__':
    exec(open("gcalendar(api).py").read())
