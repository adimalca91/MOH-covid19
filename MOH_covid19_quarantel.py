import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
import matplotlib.pyplot as plt
import seaborn as sns

sns.set()

# "C:/Users/dan.panorama2/Desktop/Adi/VisualStudio/MOH/"
WORKING_DIR = "C:/Users/עדי/Desktop/Adi/MOH/"
EXCEL_NAME = "book20.xlsx"
FILE_NAME = "show_data20.txt"

wb = load_workbook(filename=WORKING_DIR + EXCEL_NAME)
# print(wb)

sheet = wb.active
room_num_col = sheet['N']
group_num_col = sheet['D']
id_num_col = sheet['A']
age_col = sheet['O']


def create_age_arr():
    age_arr = []
    for i in range(len(age_col)):
        # print(first_col[i].value)
        age_arr.append(age_col[i].value)
    # print(len(age_arr[1:]))
    # print(age_arr[1:])
    return (age_arr[1:])


def room_analysis(room_col):
    rooms_arr = []
    for i in range(len(room_col)):
        # print(first_col[i].value)
        rooms_arr.append(room_num_col[i].value)
    used_room_arr = set(rooms_arr[1:])
    # print(used_room_arr)
    # Write to a file the Number of USED rooms:
    room_msg = f"Total Number of People: {len(rooms_arr[1:])} \nNumber of used rooms: {str(len(used_room_arr))} \n "
    return room_msg, used_room_arr, rooms_arr


def create_file(msg):
    with open(WORKING_DIR + FILE_NAME, "w") as f:
        f.write(msg)
        f.close()


def ppl_in_room(rooms_arr, used_room_arr):
    '''Check How many people are currently in the same room (ADD TO THE SAME FILE)'''
    for room in used_room_arr:
        num_of_ppl_in_room = rooms_arr[1:].count(room)
        count_msg = f"Room {room} - {num_of_ppl_in_room} people"
        with open(WORKING_DIR + FILE_NAME, "a") as f:
            f.write("\n" + count_msg)


def add_enter():
    # ADD ENTER at the end of this section in the File
    with open(WORKING_DIR + FILE_NAME, "a") as f:
        f.write("\n")
    f.close()


def group_analysis(group_col):
    group_num_arr = []
    for i in range(len(group_col)):
        # print(first_col[i].value)
        group_num_arr.append(group_col[i].value)
    group_num_no_duplicates = set(group_num_arr[1:])
    return group_num_arr, group_num_no_duplicates


def count_ppl_in_group(group_num_arr, group_num_no_duplicates):
    num_ppl_in_group_arr = []
    # Count the amount of people in the same Group Number
    for group in group_num_no_duplicates:
        num_of_ppl_in_group = group_num_arr[1:].count(group)
        num_ppl_in_group_arr.append(num_of_ppl_in_group)
        if (group == ''):
            group = "None"
        count_group_msg = f"Group {group} - {num_of_ppl_in_group} people"
        with open(WORKING_DIR + FILE_NAME, "a") as f:
            f.write("\n" + count_group_msg)
    return num_ppl_in_group_arr


# Adding labels to the top of each bar
def add_labels(x, y):
    for i in range(len(x)):
        plt.text(i, y[i], y[i], ha='center')


def graph_info_bar(group_num_no_duplicates, num_ppl_in_group_arr):
    group_num_no_dups_arr = list(group_num_no_duplicates)  # convert set to arr for plotting!
    # bar_colors = ["yellow", "blue", "red", "orange", "green", "purple"]
    # print(group_num_no_dups_arr)
    # fig = plt.figure()
    plt.bar(group_num_no_dups_arr, num_ppl_in_group_arr)  # TODO: added edge color to each bar

    plt.xlabel("Group Number")
    plt.ylabel("Amount of People in each Group")
    plt.title("Amount of People in Group Number")

    # Make y-axis interval steps of 1 (OPTIONAL)
    y_ticks = np.arange(0, max(num_ppl_in_group_arr) + 2, 5)  # TODO: changed the max height to +2 instead of +1
    plt.yticks(y_ticks)
    add_labels(group_num_no_dups_arr, num_ppl_in_group_arr)
    plt.show()


def graph_info_pie(group_num_no_duplicates, num_ppl_in_group_arr):
    # ax1 = plt.subplots()
    plt.pie(num_ppl_in_group_arr, labels=group_num_no_duplicates,
            autopct='%1.1f%%')  # shows distribution in percentage!
    plt.legend(title="Group Numbers:")
    plt.show()


def count_ages(age_arr):
    no_dups_age = set(age_arr)
    age_count_dict = {}
    for age in no_dups_age:
        age_count_dict[age] = age_arr.count(age)

    return (age_count_dict)


def plot_age_analysis(age_count_dict):
    ages = list(age_count_dict.keys())
    amounts = list(age_count_dict.values())
    plt.bar(ages, amounts)
    plt.xlabel("Age")
    plt.ylabel("Amount of People in each age group")
    plt.title("Age Analysis")
    # add_labels(ages, amounts)
    plt.show()


if __name__ == "__main__":
    room_msg, used_room_arr, rooms_arr = room_analysis(room_num_col)
    create_file(room_msg)
    ppl_in_room(rooms_arr, used_room_arr)
    add_enter()
    group_num_arr, group_num_no_duplicates = group_analysis(group_num_col)
    num_ppl_in_group_arr = count_ppl_in_group(group_num_arr, group_num_no_duplicates)
    add_enter()
    graph_info_bar(group_num_no_duplicates, num_ppl_in_group_arr)
    graph_info_pie(group_num_no_duplicates, num_ppl_in_group_arr)
    plot_age_analysis(count_ages(create_age_arr()))  # TODO: added age analysis!

