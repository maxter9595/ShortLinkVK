import os
from typing import Union

import requests
import pandas as pd
from dotenv import load_dotenv
from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename


load_dotenv()
VK_ACCESS_TOKEN = os.getenv(key='VK_ACCESS_TOKEN')
OUTPUT_FILE_NAME = os.getenv(key='OUTPUT_FILE_NAME')
NEW_COL_NAME = os.getenv(key='NEW_COL_NAME')


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def get_method_url(method_name: str) -> str:
    return "https://api.vk.com/method" + f"/{method_name}/"


def shorten_link(link: str) -> Union[str, None]:
    params = {
        'url': link,
        'access_token': VK_ACCESS_TOKEN,
        'v': '5.131',
        'private': '0'
    }
    response = requests.get(
        get_method_url("utils.getShortLink"),
        params=params
    )
    if response.status_code == 200 and response.json().get('response'):
        return response.json().get('response').get('short_url')
    return None


def check_file_extension(filename: str) -> bool:
    allowed_ext = app.config['ALLOWED_EXTENSIONS']
    allowed_ext_check = filename.rsplit('.', 1)[1].lower() in allowed_ext
    return '.' in filename and allowed_ext_check


@app.route('/')
def show_index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400

    file = request.files['file']
    if file.filename == '' or not check_file_extension(file.filename):
        return 'No selected file or invalid file type', 400

    filename = secure_filename(
        file.filename
    )
    filepath = os.path.join(
        app.config['UPLOAD_FOLDER'],
        filename
    )

    file.save(filepath)
    df = pd.read_excel(filepath)

    results = []
    for link in df.iloc[:, 0].dropna():
        short_link = shorten_link(link)
        if short_link is None:
            print(f"Failed to shorten link: {link}")
        results.append(short_link)

    df[NEW_COL_NAME] = results
    output_path = os.path.join(
        app.config['UPLOAD_FOLDER'],
        OUTPUT_FILE_NAME
    )

    df.to_excel(
        output_path,
        index=False
    )
    return send_file(
        output_path,
        as_attachment=True
    )


if __name__ == '__main__':
    app.run()
