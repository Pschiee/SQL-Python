def get_queries(file):
    query = ''
    statements=[]
    labels=[]
    
    #Goes through file line by line. Removes \n first. If the line is a comment, it adds it to the label list.
    #It will continue adding the lines until it finds the ; at which point is assumes end of query.
    for line in file:
        line_stripped = line.strip('\n')
        label_text=''
        if '--' in line_stripped:
            for i in line_stripped:
                if i != '-':
                    label_text += i
            labels.append(label_text)
        elif ';' in line_stripped:
            query+= line_stripped
            statements.append(query)
            query=''
        else:
            query = query  + line_stripped + ' '   
    return statements,labels

def execute_queries(queries,cur):
    data=[]
    #Goes through queries and executes them.
    for i in range(0,len(queries)):
        cur.execute(queries[i])
        temp = cur.fetchall()
        data.append(temp[0][0])
    return data

def write_to_excel(workbook_name,data,labels):
    ### Create a new excel document. If it already exists it will overwrite it so be careful.
    workbook = xls.Workbook(workbook_name)
    ### Create a worksheet variable. This is so the code knows which worksheet to store stuff in.
    worksheet = workbook.add_worksheet()
    #Close the excel sheet when you are done.   

    for i in range(0,1):
        ### Creates row variable so we know which row to start on.
        row=0

###The start of the for loop. It will run through every result in the "data" list which is all our data.
    for i in data:
        #For loops to determine which column we are writing to. Column 1 = number. Column 2 = Label and column 3 = Result
        for column in range(0,3):
            if column ==0:
                worksheet.write(row,column,row+1) 
            elif column == 1:
                worksheet.write(row,column,labels[row])
            else:
                worksheet.write(row,column,i)
    #Once we have finished that row then add 1 to row to go to the next row
        row+=1
    workbook.close()    
    
################################################## Code starts here ##############################
import psycopg2 as pg2
import xlsxwriter as xls

       
workbook_name = 'Example.xlsx'
queries_name = 'Example_Queries.txt'

### Creates connection to the database
conn = pg2.connect(database='dvdrental',user='postgres',password='[PASSWORD GOES HERE]',host='localhost')
### Creates cursor so we know where to execute code.
cur=conn.cursor()
### Opens the file in read only mode.
file = open(queries_name,"r")
queries,labels = get_queries(file) 
file.close()

data= execute_queries(queries,cur)

write_to_excel(workbook_name,data,labels)
