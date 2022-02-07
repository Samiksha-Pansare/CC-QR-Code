from flask import Flask, render_template, request, send_file, redirect
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.sqlite3'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class Users(db.Model):
    id = db.Column('id', db.Integer, primary_key = True)
    name = db.Column('name', db.String(100))
    mobno = db.Column('mobno', db.String(100))
    
    def __init__(self,name,mobno):
        self.name = name
        self.mobno = mobno

class Admin(db.Model):
    id = db.Column('id', db.Integer, primary_key = True)
    name = db.Column('name', db.String(100))
    password = db.Column('password', db.String(100))

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/adddata', methods = ['POST'])
def adddata():
    name = request.form.get('name')
    mobno = request.form.get('mobno')
    new_user = Users(name = name, mobno = mobno)
    try:
        db.session.add(new_user)
        db.session.commit()
        return redirect('/')
    except:
        return "There was an error in adding your details"    


@app.route("/getdata")
def getdata():
    if request.method == "POST":
        name  = request.form.get('name')
        password = request.form.get('password')
        user = Users.query.filter_by(name=name, password = password).first()
        if not user:
            return redirect("NO SUCH USER FOUND")
        if user:
            return redirect("/downloaddata")
    return render_template("getdata.html")

@app.route('/downloaddata', methods = ['POST'])
def downloaddata():
    wb = Workbook()
    ws = wb.active
    ws.title = "CC Users Data"
    headings = ["ID","Name","Mobile Number"]
    ws.append(headings)
    users = Users.query.all()
    for user in users:
        record = [user.id, user.name, user.mobno]
        ws.append(record)
    wb.save(filename='CC_Users_Data.xlsx')
    return send_file('CC_Users_Data.xlsx', as_attachment=True, download_name='CC_Users_Data.xlsx')   

if __name__ == '__main__':
    db.create_all()
    app.run(debug = True)