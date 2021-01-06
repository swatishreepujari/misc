import os
import getopt
import sys
import datetime
import pandas as pd

def help():
    os.system("cls")
    help_str=f'''
************************************************************

Run the script with  date option . You can enter present or past date . 

{sys.argv[0]}  -d ddmmyyyy 

************************************************************
'''
    print(help_str)
    sys.exit(0)

def getdate(dat):
    format = "%d%m%Y"
    try:
        datetime.datetime.strptime(dat, format)
    except ValueError:
        print("This is the incorrect date string format. It should be DDMMYYYY")
        sys.exit(2)

def get_curr_date():
    cur_dat=datetime.date.today()
    return cur_dat.strftime("%d%m%Y")


def getinput():
    try:
        arguments, values = getopt.getopt(sys.argv[1:], "hd:", ["help", "date="])
    except getopt.error as err:
    # Output error, and return with an error code
        print (str(err))
        sys.exit(2)
    if len(arguments) == 0 :
        arg=[]
        arg.append('-d')
        arg.append(get_curr_date())
        arguments.append(tuple(arg))
    for current_argument, current_value in arguments:
        if current_argument in ("-h", "--help"):
            help()
        elif current_argument in ("-d", "--date"):
            getdate(current_value)
            return current_value

def get_csv_name(v_date):
    csv_url="https://www1.nseindia.com/archives/nsccl/volt/CMVOLT_"+v_date+".CSV"
    return csv_url

def dataf(csv_name):
    return pd.read_csv(csv_name)

def excel_name(dat):
    path="d:\\test\stock_analysis-"
    return path+dat+"-"+str(datetime.datetime.now().day)+"-"+str(datetime.datetime.now().month)+".xlsx"


def main():
    input_date=getinput()
    print(get_csv_name(input_date))
    nifty100_df=dataf("d:\\test\ind_nifty100list.csv")
    df=dataf(get_csv_name(input_date))
    df.rename(columns = {'Previous Day Underlying Volatility (D)':'Previous','Current Day Underlying Daily Volatility (E) = Sqrt(0.995*D*D + 0.005*C*C)':'Current','Underlying Annualised Volatility (F) = E*Sqrt(365)':'Yearly'}, inplace = True)
    df1 = pd.DataFrame(df, columns=['Symbol','Previous','Current','Yearly'])
    merge_df=pd.merge(df1,nifty100_df,on="Symbol")
    merge_df['Previous']= pd.to_numeric(merge_df["Previous"],errors="ignore") * 100
    merge_df['Current']= pd.to_numeric(merge_df["Current"],errors="ignore") * 100
    merge_df['Yearly']= pd.to_numeric(merge_df["Yearly"],errors="ignore") * 100
    final_df=merge_df.sort_values('Current',ascending=False)
    final_df.to_excel(excel_name(input_date))


if (__name__ == "__main__"):
    os.system("cls")
    main()
