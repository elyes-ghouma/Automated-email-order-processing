from flask import Flask, Response, render_template_string, send_file
from functions import *

app = Flask(__name__)


@app.route("/")
def start():
    return Response(web_main(), mimetype="text/html")


@app.route("/download")
def end():
    return download()
    
    


if __name__ == "__main__":
    # Runs the server locally on port 5000
    app.run(host="0.0.0.0", port=5000, debug=True )

