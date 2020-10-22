# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 15:59:48 2020

@author: michal-zolyniak
http://PLWARELJJJQ2G2.eu.mrshmc.com:5000/
"""

from flask import Flask, render_template, request, redirect
import pyodbc

app = Flask(__name__)
app.config['SECRET_KEY'] = 'mysecretkey'
constr = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};' \
         r'DBQ=c:\Enron\_Mercer\Py\FeedbackForm2\.idea\Macro_Database.accdb;'
switcher = {
    'very satisfied': 5,
    'satisfied': 4,
    'somewhat dissatisfied': 3,
    'dissatisfied': 2,
    'very dissatisfied': 1,
}

def funQsTrans(i):
    return switcher.get(i, "Invalid value")

def funFeedbackExist(id):
    conn = pyodbc.connect(constr)
    cursor = conn.cursor()
    sql_select = "select * from tblFeedback WHERE MacroID= %s" % (id)
    cursor.execute(sql_select)
    records = cursor.fetchall()

    if records:
        for row in records:
            feedAdded = row.UpdatedFeedbackQuestions
    else:
        feedAdded = 'miss'

    cursor.close()
    conn.close()
    return feedAdded


def funAddFeedback(id, q1, q2, q3, q4, q5, q6):
    conn = pyodbc.connect(constr)
    cursor = conn.cursor()
    sql_update = "UPDATE tblFeedback SET q1=%s,q2=%s,q3='%s',q4=%s,q5=%s,q6=%s,UpdatedFeedbackQuestions=%s WHERE MacroID= %s" % (
    q1, q2, q3, q4, q5, q6, 1, id)
    cursor.execute(sql_update)
    conn.commit()
    cursor.close()
    conn.close()

@app.route('/', methods=['GET'])
def root():
    return redirect('/feedback')

@app.route('/feedback', methods=['GET', 'POST'])
def feedback():

    idMacro = request.args.get('id')
    if idMacro:
        if request.method == 'POST':
            q1 = funQsTrans(request.form['q1'])
            q2 = funQsTrans(request.form['q2'])
            q3 = request.form['AddText']
            q4 = funQsTrans(request.form['q3'])
            q5 = funQsTrans(request.form['q4'])
            q6 = funQsTrans(request.form['q5'])
            funAddFeedback(idMacro, q1, q2, q3, q4, q5, q6)
            return render_template('msg.html', msg1='Thank you for your feedback.', ico='img/success.svg')
        else:
            state = funFeedbackExist(idMacro)
            if state==True:
                return render_template('msg.html', msg1='Sorry but feedback for this request has been already added.',
                                       ico='img/warning.svg')
            elif state==False:
                return render_template('Feedback.html', dic=switcher)
            else:
                return render_template('msg.html', msg1='Sorry but ID for this request is not valid.',
                                       ico='img/error.svg')
    else:
        return render_template('msg.html', msg1='Error.', ico='img/error.svg')

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)
