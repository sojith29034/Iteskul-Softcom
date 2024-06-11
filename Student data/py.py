from flask import Flask, request, redirect, render_template_string

app = Flask(__name__)

# Replace these with your own credentials
USERNAME = 'user'
PASSWORD = 'pass'

@app.route('/')
def index():
    return render_template_string(open('index.html').read())

@app.route('/login', methods=['POST'])
def login():
    username = request.form['username']
    password = request.form['password']
    
    if username == USERNAME and password == PASSWORD:
        return redirect('https://mathsmp2324.streamlit.app/')
    else:
        return "Invalid credentials, please try again."

if __name__ == '__main__':
    app.run(debug=True)
