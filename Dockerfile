FROM python:3.8.2

# WORKDIR C:\Users\kamsharine\Documents\Flotech\2023\January\alarm - consolidated email\src
# directory of python file

#have to install eveyrhing , installed and used in the python dfile
COPY requirements.txt requirements.txt
RUN pip3 install -r requirements.txt

# RUN pip install pywin32
# RUN pip install pypiwin32

COPY recipients.xlsx recipients.xlsx

COPY /excel/List_of_Alarms.xls /excel/List_of_Alarms.xls

COPY app.py ./

CMD ["python3","./app.py"]