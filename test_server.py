from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('test.html')

if __name__ == '__main__':
    print("Starting test Flask server...")
    app.run(debug=True, host='127.0.0.1', port=5050)
