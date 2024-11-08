import json

import pandas as pd
from datetime import datetime

def get_group_list(year, timetable):
    group_list = []
    if year == 'М':
        for group_name in timetable.keys():
            if group_name.startswith('М') or group_name.startswith('ВСО'):
                group_list.append(group_name)
    else:
        for group_name in timetable.keys():
            if group_name.startswith(str(year)):
                group_list.append(group_name)
    return group_list

def are_times_close_enough(time1, time2, delta):
    # the function gets two time moments and checks if time2 - time1 is less than delta
    # delta is represented in minutes
    # if it is less then the function returns True
    date1 = datetime.combine(datetime.today().date(), time1)
    date2 = datetime.combine(datetime.today().date(), time2)
    delta_sec = delta * 60
    if abs((date1 - date2).total_seconds()) <= delta_sec:
        return True
    return False

def str_to_time(time_string):
    # the function gets a time string and converts it to time object
    # accepted formats are 'HH:MM' and 'HH:MM:SS'
    if len(time_string) == 5:
        return datetime.strptime(time_string, '%H:%M').time()
    if len(time_string) == 8:
        return datetime.strptime(time_string, '%H:%M:%S').time()
    return time_string

timetable = pd.read_json("timetable.json")
years = [1,2,3,4,5,6,'М']
# groups_size = pd.read_json("groups_size.json")
with open('groups_size.json', 'r') as f:
    groups_size = json.load(f)

office_work_hours = {}

for year in years:
    # group_list variable has
    group_list = get_group_list(year, timetable)
    students = {
        'понедельник': {},
        'вторник': {},
        'среда': {},
        'четверг': {},
        'пятница': {}
    }

    for group_name in group_list:
        for day in timetable[group_name].keys():
            if day in {'суббота', 'воскресенье'}:
                continue
            for student_time in timetable[group_name][day].keys():
                if student_time != 'None':
                    if student_time not in students[day]:
                        students[day][student_time] = 0
                    students[day][student_time] += groups_size[group_name]

    # search for the least busy day
    office_business = {
        'понедельник': {},
        'вторник': {},
        'среда': {},
        'четверг': {},
        'пятница': {}
    }
    for day in students.keys():
        for i in range(0,5):
            office_start_time_sec = (10 + i) * 3600
            business_now = 0
            for student_time in students[day].keys():
                student_time_sec = int(student_time[:2])*3600 + int(student_time[3:5])*60
                # case where office starts after subject of timetable
                if student_time_sec < office_start_time_sec < student_time_sec + 95*60:
                    business_now += students[day][student_time] * (office_start_time_sec - student_time_sec)
                # case where office starts before subject of timetable
                if office_start_time_sec < student_time_sec < office_start_time_sec + 3*3600:
                    # case where office ends after subject of timetable
                    if student_time_sec + 95*60 < office_start_time_sec + 3*3600:
                        business_now += students[day][student_time] * 95*60
                    # case where office ends before subject of timetable
                    if student_time_sec + 95*60 > office_start_time_sec + 3*3600:
                        business_now += students[day][student_time] * (office_start_time_sec - student_time_sec + 3*3600)
                office_business[day]["1" + str(i) + ":00:00"] = business_now
    min_business = office_business['понедельник']["10:00:00"]
    min_day = 'понедельник'
    min_time = "10:00:00"
    for day in office_business.keys():
        for time in office_business[day].keys():
            if office_business[day][time] < min_business:
                min_business = office_business[day][time]
                min_day = day
                min_time = time
    end_time = str(int(min_time[:2]) + 3) + min_time[2:]
    office_work_hours[year] = [min_day, min_time, end_time]

with open('office_work_hours.json', 'w') as f:
    json.dump(office_work_hours, f, indent=4, ensure_ascii=False)