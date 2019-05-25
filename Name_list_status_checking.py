import csv
import xlrd
import operator
import jellyfish
import sys
import itertools
import os

#Sample cmd: python Application_Name.py "ESMS Names 06212016.csv" "new OFA.csv" "Both ESMS and OFA.csv" "ESMS not in OFA.csv" "OFA not in ESMS.csv" "All names.csv" 

def main():
   
    if len(sys.argv) != 7:
        print ("Error: Please type in 6 different file names as arguments. First twos are the source file paths and the rests are the output file path and name")
    print("Process start.")
    print("Extracting the same names in both files")
    both = extract_same_names(sys.argv[1],sys.argv[2],sys.argv[3])
    print ("The names both appears in source A and source B has been extracted")
    print("Extracting the names in source A but not in source B")
    names_in_a_not_in_b =extract_names_in_source1_not_in_source2(sys.argv[1],sys.argv[2],sys.argv[4])
    print ("The names appears in source A but not in source B has been extracted")
    print("Extracting the names in source B but not in source A")
    names_in_b_not_in_a = extract_names_in_source2_not_in_source1(sys.argv[1],sys.argv[2],sys.argv[5])
    print ("The names appears in source B but not in source A has been extracted") 
    print ("Extracting all names from source A and source B")
    all_names = extract_all_names(sys.argv[3],sys.argv[4],sys.argv[5],sys.argv[6])
    print ("All the names appears in source and source A has been extracted")
    print ("Start to looking for duplicate data.")
    d1 = find_duplicates(sys.argv[1])
    d2 = find_duplicates(sys.argv[2])
    d3 = find_duplicates(sys.argv[3])
    d4 = find_duplicates(sys.argv[4])
    d5 = find_duplicates(sys.argv[5])
    d6 = find_duplicates(sys.argv[6])
    print ("Process finished") 
    print ("")
    print ("Summary Report:")
    num_in_source1 = row_number_checking(sys.argv[1])
    num_in_source2 = row_number_checking(sys.argv[2])
    print("Total number of names from source A: " + str(num_in_source1))
    print("Total number of names from source B: " + str(num_in_source2))
    print("Total number of names from source A and B (no duplicate names): "+str(all_names))
    print("Total number of names that identical: "+str(both))
    print("Total number of names not in file OFA but in ESMS: "+ str(names_in_a_not_in_b))
    print("Total number of names not in file ESMS but in OFA: "+ str(names_in_b_not_in_a))
    
    print("Total number of duplicates is found in file "+ sys.argv[1]+": "+str(d1))
    print("Total number of duplicates is found in file "+ sys.argv[2]+": "+str(d2))
    print("Total number of duplicates is found in file "+ sys.argv[3]+": "+str(d3))
    print("Total number of duplicates is found in file "+ sys.argv[4]+": "+str(d4))
    print("Total number of duplicates is found in file "+ sys.argv[5]+": "+str(d5))
    print("Total number of duplicates is found in file "+ sys.argv[6]+": "+str(d6))

    print("Identical names in ESMS: " + str("{:.0%}".format(both/num_in_source1)))
    print("Identical names in OFA: " + str("{:.0%}".format(both/num_in_source2)))
    print("Percentage of the names appears in ESMS not in OFA source files: " + str("{:.0%}".format(names_in_a_not_in_b/num_in_source1)))
    print("Percentage of the names appears in OFA not in ESMS source files: " + str("{:.0%}".format(names_in_b_not_in_a/num_in_source2)))
    

def extract_name_from_xls (source, output):
    f = open('new OFA.csv','w', newline='')
    c = csv.writer(f)
    workbook = xlrd.open_workbook('OFA Names 06212016.xls')
    sheet = workbook.sheet_by_index(0);
    for i in range(1, sheet.nrows):
        cell = sheet.cell(i, 0).value.split(", ")
        if len(cell) > 2:
            if cell[1] == 'JR' or cell[1] == 'SR.' or cell[1] == 'JR.':
                tmp = []
                cell[0]+=", JR"
                tmp.append(cell[0])
                tmp.append(cell[2])
                c.writerow(tmp)
                del tmp[:]
            else:
                print (cell)
        else:
            tmp = []
            tmp.append(cell[0])
            tmp.append(cell[1])
            c.writerow(tmp)
            del tmp[:]
    f.close
def row_number_checking (source):
    f1 = open(source,'r')
    c1 = csv.reader(f1)
    result = 0
    for eachrow in c1:
        result = result +1
    f1.close()
    return result
def extract_same_names(source_1, source_2, output):
    f1 = open(source_1,'r')
    f2 = open(source_2,'r')
    f3 = open(output,'w', newline='')

    c1 = csv.reader(f1)
    c2 = csv.reader(f2)
    c3 = csv.writer(f3)
    count = 0
    masterlist = list(c2)

    for hosts_row in c1:
        for master_row in masterlist: 
            if hosts_row[0] == master_row[0] and hosts_row[1] == master_row[1]: 
                c3.writerow(master_row)
                count = count + 1
                break
    f1.close()
    f2.close()
    f3.close()
    return count
def extract_names_in_source1_not_in_source2(source_1, source_2, output):

    f1 = open(source_1,'r')
    f2 = open(source_2,'r')
    f3 = open(output,'w', newline='')

    c1 = csv.reader(f1)
    c2 = csv.reader(f2)
    c3 = csv.writer(f3)
    count = 0
    masterlist = list(c2)

    for hosts_row in c1:
        flag = False
        for master_row in masterlist: 
            if hosts_row[0] == master_row[0] and hosts_row[1] == master_row[1]: 
                flag = True
                break
        if(flag == False):
             count = count+1
             c3.writerow(hosts_row)

    f1.close()
    f2.close()
    f3.close()
    return count
def extract_names_in_source2_not_in_source1(source_1, source_2, output):

    f1 = open(source_2,'r')
    f2 = open(source_1,'r')
    f3 = open(output,'w', newline='')  

    c1 = csv.reader(f1)
    c2 = csv.reader(f2)
    c3 = csv.writer(f3)
    count = 0
    masterlist = list(c2)

    for hosts_row in c1:
        flag = False
        for master_row in masterlist: 
            if hosts_row[0] == master_row[0] and hosts_row[1] == master_row[1]: 
                flag = True
                break
        if(flag == False):
             count = count+1
             c3.writerow(hosts_row)

    f1.close()
    f2.close()
    f3.close()
    return count

def extract_all_names(source_1,source_2,source_3,output):
    f1 = open(source_1,'r')
    f2 = open(source_2,'r')
    f3 = open(source_3,'r')
    f4 = open(output,'w', newline='')
    c1 = csv.reader(f1)
    c2 = csv.reader(f2)
    c3 = csv.reader(f3)
    c4 = csv.writer(f4)
    count = 0
    for eachline in c1:
        c4.writerow(eachline)
    for eachline in c2:
        c4.writerow(eachline)
    for eachline in c3:
        c4.writerow(eachline)
    f1.close
    f2.close
    f3.close
    f4.close
    f1 = open(output,'r')
    c1 = csv.reader(f1)
    sorted_result = sorted(c1, key=operator.itemgetter(0))
    f1.close()
    f1 = open(output,'w', newline='')
    c1 = csv.writer(f1)
    for eachline in sorted_result:
        c1.writerow(eachline)
        count = count+1
    f1.close
    return count
def find_duplicates(source):
    newpath = os.getcwd()+"\duplicate"
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    f1 = open(source,'r')
    f2 = open(source,'r')
    tmp = source.split(".csv")
    output = newpath + "\\" + tmp[0] + "_duplicate.csv"
    f3 = open(output,'w', newline='')
    c1 = csv.reader(f1)
    c2 = csv.reader(f2)
    c3 = csv.writer(f3)
    c3.writerow(["Last Name","First Name","Line Number 1", "Line Number 2"])
    count1 = 0
    current1 = 0
    current2 = 0
    checked_list = []
    result_list = []
    masterlist = list(c2)
    for row1 in c1:
        for row2 in masterlist:
            if row1[0] == row2[0] and row1[1] == row2[1] and current1 != current2 and (current1 in checked_list) == False:
                count1 = count1 + 1
                result_list = result_list + row1
                num1 = map(int, str(current1+1).split(','))
                result_list = result_list + list(num1) 
                num2 = map(int, str(current2+1).split(','))
                result_list = result_list + list(num2) 

                checked_list.append(current2)
                c3.writerow(result_list)
                del result_list[:]
                break;
            current2 = current2 + 1
        current2 = 0
        current1 = current1 + 1
    f1.close
    f2.close
    f3.close

    return count1
if __name__ == "__main__":
    main()
