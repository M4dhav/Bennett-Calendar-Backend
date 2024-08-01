from xls2xlsx import XLS2XLSX
from flask import Flask, jsonify, request
x2x = XLS2XLSX("spreadsheet.xls")
x2x.to_xlsx("spreadsheet.xlsx")



app = Flask(__name__) 


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