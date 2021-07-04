import os, pandas as pd
from flask.helpers import send_from_directory
from flask import render_template, url_for, flash, redirect, request, abort, send_file
from app import app
# from sending_emails_app.app_functionality import sndMail

@app.route("/") 
@app.route("/home",methods=['GET','POST'])
def home(): #route function
    if request.method == 'POST':   
        master_file = request.files['master_file']
        data_file = request.files['data_file']

        master_data = pd.read_excel(master_file)
        data = pd.read_excel(data_file)

        master_rolls = master_data['University RollNo.']
        emails = data["email"]
        balance = master_data["Balance"]

        locs = []
        for i in range(len(emails)):
            for j in range(len(master_rolls)):
                if str(emails[i]).split("@")[0] == str(master_rolls[j]):
                    if balance[j] < 50000:
                        locs.append(j)
        output = master_data.iloc[locs]

        output = pd.DataFrame(output)
        try:
            os.remove(os.path.join(os.getcwd(),'uploads\output.xlsx'))
            print("delete")
        except:
            pass
        output = output.to_excel("uploads/output.xlsx",index=0)
        # return render_template("home.html")  
        return redirect("/download/output.xlsx")
    return render_template('home.html') 

@app.route('/download/<path:filename>',methods = ['GET','POST'])
def downloadFile (filename):
    path = os.getcwd()
    path = os.path.join(path,"uploads")
    #For windows you need to use drive name [ex: F:/Example.pdf]
    return send_from_directory(path,"output.xlsx", as_attachment=True)

@app.route("/about") # route for about webpage 
def about(): # about route function
    return render_template('about.html') 