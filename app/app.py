from flask import Flask
from flask import send_file
from excel import makeexcel
import os

app = Flask(__name__,)

@app.route("/<user_mail>:<kw>:<jahr>")
def kimaiexport(user_mail, kw, jahr):

    fileDownload = makeexcel(user_mail, kw, jahr)

    fileDownload = fileDownload.replace("tmp/", "")

    currentDir = os.getcwd()

    return send_file(currentDir + '/tmp/' + fileDownload, download_name=fileDownload)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port = 5000, threaded = True, debug = False)