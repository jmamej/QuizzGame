import xlrd
import random

score = 0
num_of_q = 0
question_num = 0


def quiz_start():
    print(f"Welcome to the great quiz!")
    while 1:
        play_y_n = input(f"Do you want to play? (y/n) ")
        if play_y_n.lower() == 'y':
            print("Let's start!")
            play_start = True
            break
        elif play_y_n.lower() == 'n' or play_y_n.lower() == 'quit':
            return 0
        else:
            print(f"I don't understand :/")


def quiz_game():
    while 1:
        global score
        global num_of_q
        global question_num
        if num_of_q >= 10:
            print(f"Your final score is {score}/10")
            return 0
        q_and_a = xlrd.open_workbook('QandA.xls')
        wb = q_and_a.sheet_by_name('Sheet1')
        question_num += 1
        question = wb.cell(question_num, 0).value
        ans_index = [0, 1, 2, 3]
        random.shuffle(ans_index)
        ans_list = ['', '', '', '']
        a = 0
        for i in ans_index:
            ans = str(wb.cell(question_num, i + 1).value)    #pulling ans from excel
            ans_list[a] = ans   #pushing ans to list
            a += 1
        correct_ans_index = ans_index.index(3)
        print(f"\nQuestion {question_num}: {question}\n"
              f"A: {ans_list[0]}\n"
              f"B: {ans_list[1]}\n"
              f"C: {ans_list[2]}\n"
              f"D: {ans_list[3]}")
        ans_dict = {
            "A": 0,
            "B": 1,
            "C": 2,
            "D": 3
        }
        while 1:
            input_ans = input("> ")
            if input_ans.upper() in ans_dict.keys():
                break
            else:
                print("Incorrect input try again.\n> ")
        if ans_dict[input_ans.upper()] == correct_ans_index:
            score += 1
            print(f"Correct answer! Score {score}/10")
        else:
            print(f"Wrong answer. Final score {score}/10")
            return 0


quiz_start()
quiz_game()
