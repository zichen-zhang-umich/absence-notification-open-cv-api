# cmdType.value: 1 stands for left click once, 2 stands for left click twice, 3 stands for right click once
# 4 stands for type, 5 stands for wait, 6 stands for scroll

# ctype: 0 stands for empty, 1 stands for string, 2 stands for number, 3 stands for date,
# 4 stands for boolean, 5 stands for error

def check(sheet1):
    checkCmd = True
    if sheet1.nrows<2:
        print("Warning: No Data Present")
        checkCmd = False
    i = 1
    while i < sheet1.nrows:
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('Error on row #',i+1," col #1")
            checkCmd = False
        cmdValue = sheet1.row(i)[1]
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('Error on row #',i+1," col #2")
                checkCmd = False
        i += 1
    return checkCmd