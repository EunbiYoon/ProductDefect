import pandas as pd
import numpy as np
import xlrd
import smtplib
from PIL import Image
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date,timedelta
import msoffcrypto
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
from matplotlib.ticker import MaxNLocator
import datetime
from PyPDF2 import PdfFileMerger

server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']='eunbi1.yoon@lge.com'
#msg['To']='remoun.abdo@lge.com, eunbi1.yoon@lge.com, soonan.park@lge.com, isaac.milad@lge.com, aaron1.garcia@lge.com, russell.wilson@lge.com, iggeun.kwon@lge.com, dharmin.mistry@lge.com, matthew.sohn@lge.com, peter.meleek@lge.com, min1.park@lge.com, david.heo@lge.com, jiyoon1.heo@lge.com, seungjae.cho@lge.com, nungseo.park@lge.com'

#Subject 꾸미기
today=date.today()
today=today.strftime('%m/%d')
week=datetime.datetime.now()
week=week.isocalendar()
week=week[1]
msg['Subject']='[W'+str(week)+' '+today+'] Q-Bank Daily Report by R&D Team'




#Page 3
#전체 틀 만들기
fig=plt.figure(figsize=(23,14))
ax=fig.subplots(nrows=6,ncols=3,gridspec_kw={'height_ratios': [1,100,100,1,100,100]})

######################################## [i,j]
Title=[[0,0,0],
    ['Cabinet cover Gap','Noise test – no issue','UE Error retest – no issue'],
       ['Control panel Gap','Rotor noise test – no issue','Bellows leakage retest- no issue'],
          [0,0,0],
       ['Motor noise retest – no issue','Top cover Gap','PCB Touch button'],
          ['Bad Spin Inner tub',0,0]]
Title=pd.DataFrame(Title)

ax[0,0].set_axis_off()
ax[0,1].set_axis_off()
ax[0,2].set_axis_off()
ax[3,0].set_axis_off()
ax[3,1].set_axis_off()
ax[3,2].set_axis_off()
ax[5,1].set_axis_off()
ax[5,2].set_axis_off()

ax[0,0].annotate('1. Daily LQC Defect History\n<Front Loader>',xy=(0,0),fontsize=13)
ax[3,0].annotate("<Top Loader>",xy=(0,0),fontsize=11)


# Top Loader
for i in range(6):
    if i==0 or i==3:
        print("")
    else: 
        for j in range(3):
            if i==5 and i*j>=5:
                print("")
            else:
                data=pd.read_excel('//us-so11-na08765/R&D Secrets/Q-bank/Daily Report/PGM File.xlsx',sheet_name=str(i)+str(j))
                data=data.T
                print(data)
                data.columns=['W39','W40','W41','W42','W43','W44','W45','W46']
                print(data)
                data=data.drop(['NAME'],axis=0)
                print(data)
                data=data.apply(pd.to_numeric)
                print(data)

                data.plot(kind='bar',ax=ax[i,j])
                ax[i,j].set_xticklabels(labels=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],rotation=0,fontsize=9)
                ax[i,j].set_ylim(0,data.max().max()+2)
                ax[i,j].set_ylabel('EA',color='gray',fontsize=9)
                ax[i,j].set_xlabel('Day of Week',color='gray',fontsize=9)
                ax[i,j].legend(loc='upper right')
                ax[i,j].set_title(Title.at[i,j],fontsize=10)

                data=data.reset_index(drop=True)
                data=data.T
                data=data.fillna(0)
                data=data.reset_index(drop=True)
                data=data.T

                for t in range(len(data.index)):
                    for k in range(len(data.columns)):
                        if int(data.at[t,k])>0:
                            ax[i,j].annotate(int(data.at[t,k]),xy=(t-0.09*(3.1-k),data.at[t,k]+0.12),ha='left',va='bottom',fontsize=9)

plt.tight_layout()
plt.savefig('Q-Bank1.pdf')
plt.savefig('Q-Bank1.png')

plt.show()


# Page 2
fig = plt.figure(constrained_layout = True, figsize=(16,8))
gs = fig.add_gridspec(8,16)
ax1 = fig.add_subplot(gs[0:3,3:13])
ax2 = fig.add_subplot(gs[5:8,0:16])


# Table 1
ax1.set_axis_off()

data=pd.read_excel('//us-so11-na08765/R&D Secrets/Q-bank/Daily Report/PGM File.xlsx',sheet_name='PPM')
data=data.set_index('Product')
prod=pd.read_excel('//us-so11-na08765/R&D Secrets/Q-bank/Daily Report/PGM File.xlsx',sheet_name='Prod')
prod=prod.set_index('Prod')

table=ax1.table(cellText=data.values, rowLabels=data.index, colLabels=data.columns, loc='center', cellLoc='center',
               rowColours=['#D4C9EC','#D4C9EC','#D4C9EC','#D4C9EC','#D4C9EC','#D4C9EC','#EFDEBD','#EFDEBD','#EFDEBD','#EFDEBD'],
               colColours=['#AFB6BD','#AFB6BD','#B0CFEF','#B0CFEF','#B0CFEF','#B0CFEF','#AFB6BD','#AFB6BD','#AFB6BD','#AFB6BD','#AFB6BD','#AFB6BD','#AFB6BD','#AFB6BD'])
table.auto_set_font_size(False)
table.set_fontsize(9)
table.auto_set_column_width(col=list(range(len(data.columns))))
ax1.set_title('2. Q- Bank Item Daily PPM Monitoring',x=-0.1, pad=10,fontsize=11)


# Table 2
data=pd.read_excel('//us-so11-na08765/R&D Secrets/Q-bank/Daily Report/PGM File.xlsx',sheet_name='Team')
data=data.set_index('TIC')
data=data.T
graph2=data['Q Bank Items'].plot(kind='bar',ax=ax2,color=['#D6D6D6','#D6D6D6','#D6D6D6','#D6D6D6','#D6D6D6','#D6D6D6','#D6D6D6','#D6D6D6','black','#D6D6D6'])
ax2.axes.xaxis.set_visible(False)
data=data.T

table2=ax2.table(cellText=data.values, rowLabels=data.index, colLabels=data.columns, loc='bottom', cellLoc='center')
table2.auto_set_font_size(False)
table2.set_fontsize(9)

ax2.set_title('3. Q-Bank 2 : Registration Item (Total 116 Items)',x=0.15, pad=15,fontsize=11)

data=data.T
x=np.arange(len(data))

y=data['Q Bank Items'].values
for i in range(len(data)-2):
    ax2.annotate(y[i],xy=(x[i]-0.05,y[i]),va='bottom',color='#ADACAC',fontsize=10)
    
ax2.annotate(y[8],xy=(8-0.05,y[8]),va='bottom',color='black',fontsize=10)
ax2.annotate(y[9],xy=(9-0.05,y[9]),va='bottom',color='#ADACAC',fontsize=10)


# 첫번째 줄
the_cell = table[1,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[1,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')


# 두번째 줄
the_cell = table[2,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[2,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')


# 세번째 줄
the_cell = table[3,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[3,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[3,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[3,9]
the_text = the_cell.get_text()
the_text.set_color('black')
the_cell.set_facecolor('yellow')

the_cell = table[3,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[3,11]
the_text = the_cell.get_text()
the_text.set_color('black')
the_cell.set_facecolor('yellow')

the_cell = table[3,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[3,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')


# 네번째 줄
the_cell = table[4,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[4,7]
the_text = the_cell.get_text()
the_text.set_color('black')
the_cell.set_facecolor('yellow')

the_cell = table[4,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[4,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[4,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[4,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[4,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[4,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')


# 다섯번째 줄
the_cell = table[5,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[5,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[5,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[5,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[5,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[5,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[5,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[5,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')


# 여섯번째 줄
the_cell = table[6,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[6,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[6,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[6,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[6,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[6,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[6,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[6,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')


# 일곱번째 줄
the_cell = table[7,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[7,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')


# 여덟번째 줄
the_cell = table[8,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[8,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[8,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[8,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[8,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[8,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[8,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[8,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')


# 아홉번째 줄
the_cell = table[9,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[9,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[9,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[9,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[9,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[9,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[9,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')

the_cell = table[9,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('#00B050')


# 열번째 줄
the_cell = table[10,6]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,7]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,8]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,9]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,10]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,11]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,12]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')

the_cell = table[10,13]
the_text = the_cell.get_text()
the_text.set_color('white')
the_cell.set_facecolor('red')


plt.savefig('Q-Bank2.pdf')
plt.savefig('Q-Bank2.png')



# Page 3
fig = plt.figure(constrained_layout = True, figsize=(16,8))
gs = fig.add_gridspec(8,16)
ax1 = fig.add_subplot(gs[0:3,3:13])
ax2 = fig.add_subplot(gs[5:8,0:16])



# Table 1
ax1.set_axis_off()

data=pd.read_excel('//us-so11-na08765/R&D Secrets/Q-bank/Daily Report/PGM File.xlsx',sheet_name='List')
data=data.set_index('NO')

table1=ax1.table(cellText=data.values, rowLabels=data.index, colLabels=data.columns, loc='center', cellLoc='center',
               rowColours=['#D4C9EC','#D4C9EC','#D4C9EC','#D4C9EC','#D4C9EC','#D4C9EC','#EFDEBD','#EFDEBD','#EFDEBD','#EFDEBD'],
                 colColours=np.full(len(data.columns),'#C9E6EC'))
table1.auto_set_font_size(False)
table1.set_fontsize(9)
table1.auto_set_column_width(col=list(range(len(data.columns))))
ax1.set_title('4. Q-Bank items detailed review by R&D Team',x=-0.2, pad=0,fontsize=11)



# Table 2
data=pd.read_excel('//us-so11-na08765/R&D Secrets/Q-bank/Daily Report/PGM File.xlsx',sheet_name='Plan')
data=data.set_index('Week')
data=data.T

axF1=data['Applied Plan'].astype(int).plot(kind='bar',ax=ax2,color='#BBBABA')
axF2 = axF1.twinx()
axF1=data['Completed Rate'].plot(kind='line',marker='o', markersize=6,color='black')
ax2.axes.xaxis.set_visible(False)

ax2.set_ylim(0,10)
axF2.set_ylim(0,110)


data=data.T
table2=ax2.table(cellText=data.values, rowLabels=data.index, colLabels=data.columns, loc='bottom', cellLoc='center')
table2.auto_set_font_size(False)
table2.set_fontsize(9)

ax2.set_title('5. Q-Bank kick off by R&D Team',x=0.02, pad=15,fontsize=11)

ax2.set_ylabel('EA',color='black',fontsize=9)
axF2.set_ylabel('%',color='black',fontsize=9)

data=data.T
x=np.arange(len(data))
y1=data['Applied Plan'].values
y2=data['Completed Rate'].values
for i in range(len(data)):
    if y1[i]!=0:
        ax2.annotate(y1[i],xy=(x[i]-0.05,y1[i]),va='bottom',color='#959494',fontsize=10)
    axF2.annotate(y2[i],xy=(x[i]-0.05,y2[i]),va='bottom',color='black',fontsize=10)
   

ax2.axes.xaxis.set_visible(False)
plt.savefig('Q-Bank3.pdf')
plt.savefig('Q-Bank3.png')




# pdf 파일로 저장하기
today=date.today()
today=today.strftime('%m/%d')
file_name='Q-Bank Daily Report by R&D Team_'+today+'.pdf'


pdfs = ['Q-Bank1.pdf','Q-Bank2.pdf','Q-Bank3.pdf']

merger = PdfFileMerger()

for pdf in pdfs:
    merger.append(pdf)

merger.write('Q-Bank Daily Report by R&D Team_1122.pdf')
merger.close()




#Body 꾸미기
text0='This is DX activities from LGEUS R&D Team\nPerson in charge: LGEUS R&D Team Eunbi Yoon\n\n\n'
msg.attach(MIMEText(text0,'plain'))
text1='Dear All,\n\nThis is the report ofthe Daily Q Bank Activities Progress by R&D Team\n* It is based on a day before reporting.\nPlease refer to attachment'
msg.attach(MIMEText(text1,'plain'))

#첨부 파일1
with open('Q-Bank1.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('Q-Bank1.png'))
msg.attach(image)

#첨부 파일1
with open('Q-Bank2.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('Q-Bank2.png'))
msg.attach(image)

#첨부 파일1
with open('Q-Bank3.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('Q-Bank3.png'))
msg.attach(image)

#첨부 파일1
with open('sign.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('sign.png'))
msg.attach(image)



#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")

