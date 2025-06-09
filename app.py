from flask import Flask, jsonify, request

app = Flask(__name__)

@app.route('/')
def home():
    return "Hello, this is your InDesign server!"

# Example API: Export PDF
@app.route('/export_pdf', methods=['POST'])
def export_pdf():
    data = request.json
    filename = data.get('filename', 'default.pdf')
    # Here you would add code to trigger InDesign PDF export
    return jsonify({"status": "success", "message": f"PDF exported as {filename}"})

# Example API: Find XML tags
@app.route('/find_xml_tags', methods=['GET'])
def find_xml_tags():
    # Pretend we read tags from an InDesign file
    tags = ['ce:para', 'ce:title', 'ce:section', 'ce:list']
    return jsonify({"tags": tags})

# Example API: Clean document
@app.route('/clean_document', methods=['POST'])
def clean_document():
    # Pretend to clean doc, real logic goes here
    return jsonify({"status": "success", "message": "Document cleaned"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
