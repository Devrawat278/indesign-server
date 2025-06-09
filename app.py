from flask import Flask, jsonify, request, render_template

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/export_pdf', methods=['POST'])
def export_pdf():
    data = request.json
    filename = data.get('filename', 'default.pdf')
    return jsonify({"status": "success", "message": f"PDF exported as {filename}"})

@app.route('/find_xml_tags', methods=['GET'])
def find_xml_tags():
    tags = ['ce:para', 'ce:title', 'ce:section', 'ce:list']
    return jsonify({"tags": tags})

@app.route('/clean_document', methods=['POST'])
def clean_document():
    return jsonify({"status": "success", "message": "Document cleaned"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
