# Author: Ng Wei Jian 
# Title: Automating Extraction of BGS borehole logs 
# The purpose of this script is to allow extraction of borehole logs from BGS
import re 
import urllib.request
import xlsxwriter
import os

def check_dir(file_name):
    directory = os.path.dirname(file_name)
    if not os.path.exists(directory):
        os.makedirs(directory)

check_dir(os.path.join('Borehole_logs',''))

BGS = open("GeoIndexData.txt")
headings = BGS.readline().split()
# Create Excel workbook to store data
excel = xlsxwriter.Workbook('Data.xlsx')
worksheet = excel.add_worksheet('My_data')
column_counter = 0
for k in headings:
    worksheet.write(0, column_counter, k)
    column_counter += 1
    

# headings contain the headings of the data
# [1]: Record; [2]: Reference; [3]: Name; [4]: Borehole Length; [5]: Year known; [6]: Eastings; [7]: Northings
data = []
for line in BGS:
    data.append(line.split("\t"))

# cleaning up the link column using a function 
def clean(lt):
    string = lt[0]
    m = re.match("<a href='(.*)\' class", string)
    new_lt = [m.group(1)+'/images/'] + lt[1::]
    return new_lt
data = list(map(clean, data))
data1 = []
# filtering out the confidential data -> This is inefficient
for i in data:
   if "shop" not in i[0]:
       data1.append(i)

# In cases where the name of the BH references has special character; this step will remove it
def remove_special(lt):
    lt_new = lt
    lt_new[1] = re.sub('/','',lt[1])
    return lt_new

data1 = list(map(remove_special,data1))

# Assigning agent for opener to avoid 403 error 
opener = urllib.request.build_opener()
opener.addheaders = [('User-Agent', 'Mozilla/5.0')]
urllib.request.install_opener(opener)

# Function to clean the url list contained in data1[0] / images folder
def clean2 (my_string):
    my_string = list(my_string)
    my_string.pop(0) # pop the b
    my_string.pop(0)
    for i in range(5):
        my_string.pop(-1)
    return "".join(my_string)+'.png'

# Function to download / retrive images
def download_image(url, file_path, file_name):
    full_path = file_path + file_name + '.png'
    urllib.request.urlretrieve(url, full_path)

counter = 0 # this is for tracking progress
new_length = len(data1)
error_log = open("error_log.txt","w+") # open error log file to write
row_counter = 1
column_counter = 0 # reset back to 0
count1 = 0 # this is to check amount of file not downloaded
count_error = 0 # this is to check amount of file not downloaded
for i in data1:
    url = list(urllib.request.urlopen(i[0]))
    url = list(map(str, url))
    url = list(map(clean2, url))
    count = 0 # this is to increment the file
    for j in url:
        count1 += 1
        try:
            download_image(j, 'Borehole_logs/', i[1]+'_'+ str(count))
            count += 1
        except urllib.error.HTTPError as e:
            if e.code in (..., 403, ...):
                error_log.write('The image file from this link: ' + j + ' has not been downloaded due to 403 error \n')
                count_error += 1
                # appending the error link to error list for checking
                continue
    for k in i:
        worksheet.write(row_counter, column_counter, k)
        column_counter += 1
    column_counter = 0
    row_counter += 1
    count = 0
    counter += 1
    print(i[1] +' file download completed. Progress = ', round(100*counter/new_length,1), '%')

error_log.close()  # closing error log file
excel.close() # closing excel file

print('Completed')
print(150*'=')
print('Summary Report')
print(150*'=')
print('Total number of logs in area:', len(data)-1)
print('Number of logs that needs to be purchased: %d ' % (int((len(data)-len(data1)))))
print('Percentage images not downloaded due to 403 error: ', round(100*count_error/count1, 1), '%')

    

