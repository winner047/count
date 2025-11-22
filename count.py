from flask import Flask, request, render_template_string
import pandas as pd
inp

app = Flask(__name__)

# HTML template for file upload
upload_form = '''
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <title>Upload Excel File</title>
  </head>
  <body>
    <div class="container mt-5">
      <h2>Upload an XLSX File</h2>
      <form method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx"><br><br>
        <button type="submit" class="btn btn-primary">Upload</button>
      </form>
      {% if data %}
        <h3>Data from Excel:</h3>
        <table class="table table-striped">
          {{ data.to_html(classes='data', header="true", index=False) | safe }}
        </table>
      {% endif %}
    </div>
  </body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    data = None
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return "No file part"
        file = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename.
        if file.filename == '':
            return "No selected file"
        if file and file.filename.endswith('.xlsx'):
            try:
                data = pd.read_excel(file)
            except Exception as e:
                return f"Error reading file: {str(e)}"
    return render_template_string(upload_form, data=data)

if __name__ == '__main__':
    app.run(debug=True)


