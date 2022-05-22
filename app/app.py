from flask import Flask
from flask import send_file
from excel import makeexcel
import os

app = Flask(__name__,)

@app.route("/<user>:<kw>:<jahr>")
def kimaiexport(user, kw, jahr):

    fileDownload = makeexcel(user, kw, jahr)

    fileDownload = fileDownload.replace("tmp/", "")

    currentDir = os.getcwd()

    return send_file(currentDir + '/tmp/' + fileDownload, download_name=fileDownload)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port = 5000, threaded = True, debug = False)