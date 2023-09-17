
import os, re
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt

rootDir = 'path/to/lessonTable/folder'
filenames = os.listdir(rootDir)
filenames = list(filter(lambda x: x.find('xls') != -1, filenames))

all_lessons = np.zeros((25, 7, 6), dtype=np.int32)
pattern = r'([\d,\- ]+)[\(\[]+周'
reg = re.compile(pattern)
regDate1 = re.compile(r'([0-9]+)-([0-9]+)')
regDate2 = re.compile(r'[^-]([0-9]+),')
regDate3 = re.compile(r',([0-9]+)[^-]')

def fill_lessons(lesson, day, time):
    #print(lesson)
    found = reg.findall(lesson)
    for k in found:
        #print(k)
        if len(k) <= 2:
            #print(int(k))
            all_lessons[int(k)][day][time] += 1
            continue
        duration, single = [], set()
        #print(dur, sing)
        periods = k.split(',')
        for p in periods:
            if p.find('-') != -1:
                duration.append(p.split('-'))
            else:
                single.add(p)
        for a in duration:
            for week in range(int(a[0]), int(a[1]) + 1):
                #if week == 4 and day == 1 and time == 0:
                #    print(lesson)
                #    print(k, duration, single)
                all_lessons[week][day][time] += 1
        for a in single:
            all_lessons[int(a)][day][time] += 1

lesson_strs = []

def process_person(frame):
    for day in range(1, 8):
        day_lessons = data[data.columns[day]]
        #print(day_lessons)
        tempList = []
        for time in range(2, 8):
            lesson = day_lessons.loc[time]
            result = ''.join((str(lesson).split('\n')))
            if len(result) <= 10:
                tempList.append('none')
            else:
                tempList.append(result)
            if (type(lesson) == type(.1)):
                continue
            if (len(lesson) == 1):
                continue
            fill_lessons(lesson, day - 1, time - 2)
        lesson_strs.append(tempList)
lesson_str=[]

for f in filenames:
    #print(f)
    ff = os.path.join(rootDir, f)
    data = pd.read_excel(ff)
    process_person(data)
    
for week in range(1, 3):
    bg = pd.DataFrame(columns=['周一', '周二', '周三', '周四', '周五'])
    bg.loc["8:00-10:00"] = [all_lessons[week][i][1] for i in range(1, 6)]
    bg.loc["10:00-12:00"] = [all_lessons[week][i][2] for i in range(1, 6)]
    bg.loc["2:00-4:00"] = [all_lessons[week][i][3] for i in range(1, 6)]
    bg.loc["4:00-6:00"] = [all_lessons[week][i][4] for i in range(1, 6)]
    bg.loc["7:00-9:00"] = [all_lessons[week][i][5] for i in range(1, 6)]
    print(bg)
    
def plotH(week, all_lessons, save=False):
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    times = ['8:00-9:50', '10:10-12:00', '2:00-3:50', '4:10-6:00', '7:00-8:50', '9:10-10:00']
    
    plt.figure()
    plt.imshow(all_lessons[week])
    plt.xticks(np.arange(len(days)), labels=days)
    plt.yticks(np.arange(len(times)), labels=times, ha='right')
    plt.title('Week ' + str(week) + ' Time Table')
    plt.colorbar()

    for day in range(len(days)):
        for time in range(len(times)):
            txt = 'No. ' + str(time + 3) + '\n' + str(all_lessons[week, time, day]) + ' Stu'
            plt.text(day, time, txt, ha='center', va='center', color='red')

    plt.tight_layout()
    plt.savefig(r'C:\Users\a\Desktop\课表2023春\figures\week' + str(week) + ".jpg", dpi=450)
    plt.show()
    
all_lessons = all_lessons.swapaxes(1, 2)

for week in range(8, 9):
    plotH(week, all_lessons, False)
    
all_lessons = all_lessons.swapaxes(1, 2)

all_lessons = all_lessons.transpose(1, 2, 0)
accu_str = np.sum(all_lessons, axis=-1)
print(accu_str)