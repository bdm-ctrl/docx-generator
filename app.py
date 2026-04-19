from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import Pt, Cm
import io
import os
import re

app = Flask(__name__)

def replace_placeholders(doc, fields):
    """Replace all placeholders in document with actual values"""
    
    # Маппинг алиасов для совместимости с шаблонами
    aliases = {
        'doc_date_day': 'contract_date_day',
        'doc_date_month': 'contract_date_month',
        'doc_number': 'contract_number',
        'doc_date': 'contract_date',
    }
    
    # Расширяем fields алиасами
    extended_fields = dict(fields)
    for original, alias in aliases.items():
        if original in fields:
            extended_fields[alias] = fields[original]
    
    def replace_in_paragraph(paragraph, fields):
        for key, value in fields.items():
            placeholder = '{{' + key.upper() + '}}'
            if placeholder in paragraph.text:
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value or ''))
    
    def replace_in_table(table, fields):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, fields)
                for nested_table in cell.tables:
                    replace_in_table(nested_table, fields)
    
    # Replace in all paragraphs
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, extended_fields)
    
    # Replace in all tables
    for table in doc.tables:
        replace_in_table(table, extended_fields)
    
    return doc

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        doc_type = data.get('doc_type')
        fields = data.get('fields', {})
        
        if not doc_type:
            return jsonify({'error': 'doc_type is required'}), 400
        
        # Map doc_type to template file
        template_map = {
            'doc_contract': 'templates/contract_template.docx',
            'doc_appendix': 'templates/appendix_template.docx',
            'doc_invoice_ru': 'templates/invoice_ru_template.docx',
            'doc_act': 'templates/act_template.docx',
        }
        
        template_path = template_map.get(doc_type)
        if not template_path:
            return jsonify({'error': f'Unknown doc_type: {doc_type}'}), 400
        
        if not os.path.exists(template_path):
            return jsonify({'error': f'Template not found: {template_path}'}), 404
        
        # Load template and replace placeholders
        doc = Document(template_path)
        doc = replace_placeholders(doc, fields)
        
        # Save to buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # Generate filename
        doc_number = fields.get('doc_number', 'draft')
        filename = f"{doc_type}_{doc_number}.docx"
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
