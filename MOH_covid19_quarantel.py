import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
import matplotlib.pyplot as plt

# "C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/"
WORKING_DIR = "C:/Users/עדי/Desktop/Adi/MOH/"
EXCEL_NAME = "book12.xlsx"
FILE_NAME = "show_data12.txt"

wb = load_workbook(filename=WORKING_DIR + EXCEL_NAME)
# print(wb)

sheet = wb.active
room_num_col = sheet['N']
group_num_col = sheet['D']


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


# room_msg, used_room_arr, rooms_arr = room_analysis(room_num_col)

def create_file(msg):
    with open(WORKING_DIR + FILE_NAME, "w") as f:
        f.write(msg)
        f.close()


# create_file(room_msg)

def ppl_in_room(rooms_arr, used_room_arr):
    '''Check How many people are currently in the same room (ADD TO THE SAME FILE)'''
    for room in used_room_arr:
        num_of_ppl_in_room = rooms_arr[1:].count(room)
        count_msg = f"Room {room} - {num_of_ppl_in_room} people"
        with open(WORKING_DIR + FILE_NAME, "a") as f:
            f.write("\n" + count_msg)


# ppl_in_room(rooms_arr, used_room_arr)


def add_enter():
    # ADD ENTER at the end of this section in the File
    with open(WORKING_DIR + FILE_NAME, "a") as f:
        f.write("\n")
    f.close()


# add_enter()


def group_analysis(group_col):
    group_num_arr = []
    for i in range(len(group_col)):
        # print(first_col[i].value)
        group_num_arr.append(group_col[i].value)
    group_num_no_duplicates = set(group_num_arr[1:])
    return group_num_arr, group_num_no_duplicates


# group_num_arr, group_num_no_duplicates = group_analysis(group_num_col)

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


# num_ppl_in_group_arr = count_ppl_in_group(group_num_arr, group_num_no_duplicates)

# add_enter()


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


# graph_info(group_num_no_duplicates, num_ppl_in_group_arr)

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

########################################################################################################
########################################### WORKING CODE START #########################################
########################################################################################################

# rooms_arr = []
# for i in range(len(room_num_col)):
#     # print(first_col[i].value)
#     rooms_arr.append(room_num_col[i].value)

# # print(names_arr[1:]) #Get rid of the header - usually in Hebrew
# print(rooms_arr)

# used_room_arr = set(rooms_arr[1:])
# print(used_room_arr)
# print("Number Rooms with duplicates: ", len(rooms_arr[1:]))
# print("Number of used rooms: ", len(used_room_arr))   #This is the total number of USED rooms! No duplicates!

# # Write to a file the Number of USED rooms:
# conclusion = f"Total Number of People: {len(rooms_arr[1:])} \nNumber of used rooms: {str(len(used_room_arr))} \n "

# with open("C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/show_data.txt", "w") as f:
#     f.write(conclusion)
#     f.close()

# # Check How many people are currently in the same room (ADD TO THE SAME FILE)
# for room in used_room_arr:
#     num_of_ppl_in_room = rooms_arr[1:].count(room)
#     count_msg = f"Room {room} - {num_of_ppl_in_room} people"
#     with open("C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/show_data.txt", "a") as f:
#         f.write("\n" + count_msg)

# # ADD ENTER at the end of this section in the File
# with open("C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/show_data.txt", "a") as f:
#     f.write("\n")
# f.close()

# # Group Number -
# group_num_col = sheet['D']
# print(len(group_num_col))
# group_num_arr = []
# for i in range(len(group_num_col)):
#     # print(first_col[i].value)
#     group_num_arr.append(group_num_col[i].value)

# print(group_num_arr)

# group_num_no_duplicates = set(group_num_arr[1:])
# print(group_num_no_duplicates)
# print("Group Number with duplicates ", len(group_num_arr[1:]))
# print("Group Number NO duplicates ", len(group_num_no_duplicates))   #This is the total number of USED rooms! No duplicates!

# num_ppl_in_group_arr = []
# # Count the amount of people in the same Group Number
# for group in group_num_no_duplicates:
#     num_of_ppl_in_group = group_num_arr[1:].count(group)
#     num_ppl_in_group_arr.append(num_of_ppl_in_group)
#     if (group == ''):
#         group = "None"
#     count_group_msg = f"Group {group} - {num_of_ppl_in_group} people"
#     with open("C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/show_data.txt", "a") as f:
#         f.write("\n" + count_group_msg)


# # ADD ENTER at the end of this section in the File
# with open("C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/show_data.txt", "a") as f:
#     f.write("\n")
# f.close()


# group_num_no_dups_arr = list(group_num_no_duplicates) #convert set to arr for plotting!
# #bar_colors = ["yellow", "blue", "red", "orange", "green", "purple"]
# # print(group_num_no_dups_arr)
# fig = plt.figure()
# plt.bar(group_num_no_dups_arr, num_ppl_in_group_arr)

# plt.xlabel("Group Number")
# plt.ylabel("Amount of People in each Group")
# plt.title("Amount of People in Group Number")

# #Make y-axis interval steps of 1
# y_ticks = np.arange(0, max(num_ppl_in_group_arr)+1, 1)
# plt.yticks(y_ticks)
# plt.show()


########################################################################################################
########################################### WORKING CODE END ###########################################
########################################################################################################


# for i in x1:
#     print(i + " : " + x1.value)


#################################################################################

# chart = BarChart() #create a Barchart object
# chart.title = "Hotel Information"
# chart.y_axis.title = "Amount of people in hotel"
# chart.x_axis.title = "Positive / Isolated"

# # sheet_values = wb['Sheet']

# data = Reference(sheet, min_row=2, max_row=5, min_col=1, max_col=3)
# chart.add_data(data, from_rows=True, titles_from_data=True)

# sheet.add_chart(chart, "A10")
# wb.save('Book1.xlsx')


###############################################################################

# with open("C:/Users/dan.panorama2/Desktop/VisualStudio/MOH/moh_files_test1.txt", "w") as f:
#     f.write("HEY THIS IS FIRST PROGRAM CREATING A FILE! SHOW IN CORRECT DIRECTORY")
#     f.close()

# with open("moh_files_test.txt", "r") as f:
#     print(type(f))