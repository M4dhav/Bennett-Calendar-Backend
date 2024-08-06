from xls2xlsx import XLS2XLSX
from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename
from flask_cors import CORS
from io import BytesIO
from openpyxl import Workbook
import os
import json

app = Flask(__name__) 
CORS(app)


ALLOWED_EXTENSIONS = set(['xls', ])
UPLOAD_FOLDER = os.path.abspath(os.path.join(os.path.dirname(__file__), 'Downloads'))
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 500 * 1000 * 1000  # 500 MB
app.config['CORS_HEADER'] = 'application/json'

def allowedFile(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
           

splcourses = {"AI":["CSET225", "Intelligent Model Design using AI"], "Cloud Computing":["CSET232", "Design of Cloud Architectural Solutions"] }
maincourses = {"CSET208":"Ethics for Engineers, Patents, Copyrights and IPR","CSET302":"Automata Theory and Computability",\
    "CSET305":"High Performance Computing","CSET303":"Seminar on Special Topics in Emerging Areas", "CSET304":"Competitive Programming",}
splelectives = {"Natural Language Processing":"CSET346", }
electives = {"Soft Computing": "CSET326"}
      
           
def parse(specialisation,elective, wb):
    elecCourse = electives[elective]
    splcourse = splcourses[specialisation]
    ttwb = wb
    tt = ttwb.active
    coursenames = [[],[],[],[],[]]
    rooms = [[],[],[],[],[]]
    c = 0
    for i in range(2,7):
        for j in range(5,14):
            value = tt.cell(row = j, column = i).value
            if value == None or value == "":
                coursenames[c].append("Free")
                rooms[c].append("Free")
                continue
            
            if any(k in value for k in list(maincourses.keys())):
                i1 = value.index("{")
                i2 = value.index("}")
                room = value[i1+1:i2]
                coursenames[c].append(value)
                rooms[c].append(room)
            elif splcourse[0] in value:
                i1 = value.index(splcourse[0])
                value = value[i1:]
                i2 = value.index("}")
                value = value[:i2+1]
                print("This is value",value)
                i3 = value.index("{")
                room = value[i3+1:i2]
                rooms[c].append(room)
                coursenames[c].append(value)
            elif elecCourse in value:
                i1 = value.index(elecCourse)
                value = value[i1:]
                i2 = value.index("}")
                value = value[:i2+1]
                i3 = value.index("{")
                room = value[i3+1:i2]
                rooms[c].append(room)
                coursenames[c].append(value)
            else:
                coursenames[c].append("Free")
                rooms[c].append("Free")
        c += 1
    return coursenames, rooms
           
           
           
           
@app.route('/upload', methods=['POST', 'GET'])
def fileUpload():
    if request.method == 'POST':
        file = request.files.getlist('files')
        spl = request.form['spl']
        spl_elective = request.form['spl_elective']
        elective = request.form['elective']
        print(spl, elective, spl_elective)
        filename = ""
        print(request.files, "....")
        for f in file:
            print(f.filename)
            filename = secure_filename(f.filename)
            print(allowedFile(filename))
            if allowedFile(filename):
                f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                x2x = XLS2XLSX(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                wb = x2x.to_xlsx()
                coursenames, rooms = parse("AI", "Soft Computing", wb)
                
            else:
                return jsonify({'message': 'File type not allowed'}), 400
            

        return jsonify(json.dumps({'coursenames': coursenames,'rooms': rooms}))
        # return jsonify({"name": filename, "status": "success"})
    else:
        return jsonify({"status": "Upload API GET Request Running"})



# on the terminal type: curl http://127.0.0.1:5000/ 
# returns hello world when we use GET. 
# returns the data that we send when we use POST. 
@app.route('/', methods = ['GET', 'POST']) 
def home(): 
	if(request.method == 'GET'): 
		data = "hello world"
		return jsonify({'data': data}) 


if __name__ == '__main__': 
	app.run(debug = True) 