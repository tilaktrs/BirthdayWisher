import pandas as pd 
import datetime
import smtplib
import os
os.chdir(r"E:\Visual Studio Project\BirthdayWish project")

GMAIL_ID='raotilakrajsingh@gmail.com'
GMAIL_PSWD = '19studdisssstudd98'

def sendEmail(to,sub,msg):
    print(f"Email to {to} sent with  subject:{sub} and messsage {msg}")# it is a way to format your string that is more readable and fast. The f or F in front of strings tell Python to look at the values inside {} and substitute them with the variables values if exists.
    #creating session
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)

    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()



if __name__ == "__main__":
     #sendEmail(GMAIL_ID,"HAPPY BIRTHDAY","HEY , happy birthday! stay safe!stay blessed!")
     df = pd.read_excel("Book1.xlsx")
     #print(df)
     today = datetime.datetime.now().strftime("%d-%m")
     yearNow = datetime.datetime.now().strftime("%Y")#to keep track of year to avoid sending mail again and again
    #print(today)
     #to iterate data frame to get index and its corresponding column values
     writeInd = []
     for index, item in df.iterrows():
        #print(index, item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")#to bring it in same format as today date to compare
        #print(bday)
        if(today == bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'],'HAPPY BIRTHDAY',item['Dialogue'])
            writeInd.append(index)

#print(writeInd)
for i in writeInd:
    yr = df.loc[i,'Year']#loc is label-based, which means that we have to specify the name of the rows and columns that we need to filter out.
    df.loc[i,'Year'] = str(yr)+','+str(yearNow)
    print(df.loc[i,'Year'])

#print(df)#after wish in a particular year we will lock that year to avoid sending again and again
df.to_excel('Book1.xlsx', index=False)
#print(df)




