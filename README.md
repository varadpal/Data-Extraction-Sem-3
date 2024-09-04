# Data-Extraction-Sem-3

    import pandas as pd
    
    excel = pd.ExcelFile('Results.xlsx')
    sheet1 = excel.parse('Results').to_dict()
    
    for i in sheet1:
        for j in range(0, 3):
            del sheet1[i][j]
    
    for i in sheet1['Unnamed: 0']:
        if pd.isna(sheet1['Unnamed: 0'][i]):
            continue
        else:
            name = sheet1['Unnamed: 0'][i]
            registrationNo = sheet1['Unnamed: 1'][i]
            program = sheet1['Unnamed: 2'][i]
            date = sheet1['Unnamed: 3'][i].strftime("%d-%b-%Y")
            courses = {
                'Statistics & Probability': 'Pass',
                'Microsoft Excel': 'Pass',
                'VBA': 'Pass',
                'SQL': 'Pass',
                'Power BI': 'Pass',
            }
            if sheet1['Unnamed: 5'][i] == 'Fail':
                courses['Statistics & Probability'] = 'Fail'
            if sheet1['Unnamed: 7'][i] == "Fail":
                courses['Microsoft Excel'] = 'Fail'
            if sheet1['Unnamed: 9'][i] == "Fail":
                courses['VBA'] = 'Fail'
            if sheet1['Unnamed: 11'][i] == "Fail":
                courses['SQL'] = 'Fail'
            if sheet1['Unnamed: 13'][i] == "Fail":
                courses['Power BI'] = 'Fail'

        grandTotal = sheet1['Unnamed: 14'][i]

        file = open('Each Student Record//{}.txt'.format(name), 'w')
        passedCourses = []
        failedCourses = []
        for key in courses.keys():
            if courses[key] != 'Fail':
                passedCourses.append(key)
            else:
                failedCourses.append(key)

        if failedCourses == []:
            file.write(f"{name} bearing registration number \n{registrationNo},"
            f"enrolled in the program {program}\nhas appeared in final exams on {date}."
            f"Student\nhas successfully qualified the following courses:\n")
            for x in range(len(passedCourses)):
                file.write(f"{passedCourses[x]}, ")
            file.write(f"and have\nscored a grand total of {grandTotal} marks.")

        else:
            file.write(f"{name} bearing registration number \n{registrationNo},"
            f"enrolled in the program {program}\nhas appeared in final exams on {date}."
            f"Student\nhas successfully qualified the following courses:\n")
            for x in range(len(passedCourses)):
                file.write(f"{passedCourses[x]}, ")
            file.write(f"and have\nscored a grand total of {grandTotal} marks.\nBut he failed the ")
            for y in range(len(failedCourses)):
                file.write(f"{failedCourses[y]} ")
            file.write("course.\nHe may have to re-appear in this exam.")
