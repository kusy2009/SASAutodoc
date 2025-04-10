import os
import logging
import re
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from openai import OpenAI
import json
import tempfile
from document_generator import create_manual, create_presentation

# Set up logging
logging.basicConfig(level=logging.DEBUG, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "a secret key")

# Initialize OpenAI client
client = OpenAI(api_key="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")

def generate_intelligent_comments(sas_code):
    """Generate intelligent comments for SAS code using OpenAI API"""
    try:
        messages = [{
            "role": "system",
            "content": """You are a SAS programming expert. Analyze the provided SAS code and return JSON containing comments.
            Format your response in JSON as:
            {
                "code": "The original code with /* comments */ added above relevant lines"
            }"""
        }, {
            "role": "user",
            "content": f"Please add only critical logical comments precisely to this SAS code and return as JSON:\n{sas_code}"
        }]

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.2,
            response_format={"type": "json_object"}
        )

        result = json.loads(response.choices[0].message.content)
        return result.get('code', sas_code)
    except Exception as e:
        logger.error(f"Error generating intelligent comments: {str(e)}")
        return sas_code
    
def find_macro_calls(sas_code):
    """Extract macro calls from SAS code"""
    try:
        macro_calls = re.findall(r'%(\w+)\s*\(', sas_code, re.IGNORECASE)
        unique_calls = list(set(macro_calls))
        if unique_calls:
            return ", ".join([f"%{macro}" for macro in unique_calls])
        return ""
    except Exception as e:
        logger.error(f"Error finding macro calls: {str(e)}")
        return ""

def generate_parameter_description(param_name, macro_code):
    """Generate an AI description for a macro parameter"""
    try:
        messages = [{
            "role": "system",
            "content": """You are a SAS macro expert. Generate a very brief description for the given macro parameter.
            Consider the parameter name and the macro code context.
            Return response in JSON format as: {"description": "parameter description"}"""
        }, {
            "role": "user",
            "content": f"Generate a brief description for the parameter '{param_name}' in this SAS macro:\n{macro_code}"
        }]
        response = client.chat.completions.create(
            model="gpt-4o", 
            messages=messages,
            temperature=0.2,
            max_tokens=60,
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        return result.get('description', f"Parameter {param_name} for the macro")
    except Exception as e:
        logger.error(f"Error generating parameter description: {str(e)}")
        return f"Parameter {param_name} for the macro"

def extract_macro_parameters(sas_code):
    """Extract macro parameters with AI-generated descriptions"""
    try:
        macro_def = re.search(r'%macro\s+\w+\((.*?)\);', sas_code, re.IGNORECASE | re.MULTILINE)
        if not macro_def:
            return []

        params_str = macro_def.group(1)
        params = [p.strip() for p in params_str.split(',') if p.strip()]
        
        param_list = []
        for param in params:
            param_parts = param.split('=')
            param_name = param_parts[0].strip()
            default_value = param_parts[1].strip() if len(param_parts) > 1 else ''
            
            # Generate AI-powered description
            description = generate_parameter_description(param_name, sas_code)
            
            param_list.append([
                param_name,
                default_value or "None",
                description
            ])

        return param_list
    except Exception as e:
        logger.error(f"Parameter extraction error: {str(e)}")
        return []

def generate_doc_content(sas_code, feedback=None):
    """Generate documentation content using OpenAI API"""
    try:
        structure = analyze_macro_structure(sas_code)
        macro_match = re.search(r'%macro\s+(\w+)', sas_code, re.IGNORECASE)
        if not macro_match:
            raise ValueError("Could not find macro name in the code")

        # Extract actual parameters from the macro
        macro_params = extract_macro_parameters(sas_code)

        base_content = {
            "Overview": "string describing the macro's purpose",
            "Syntax": "string showing the macro call syntax",
            "Parameters": {
                "table_headers": ["Parameter", "Default", "Description"],
                "table_rows": macro_params  # Use actual extracted parameters
            },
            "Key Features and Functionalities": {
                "main_section": "A concise overview of the macro's main features",
                "subsections": []
            },
            "Usage Examples": ["Example usage of the macro will be provided here"],
            "Return Values": "",
            "Error Handling": "",
            "Summary": ""
        }

        messages = [{
            "role": "system",
            "content": f"""You are a SAS macro expert. Analyze the provided SAS macro code and return a JSON document with comprehensive documentation.
            The macro contains {structure['data_steps']} DATA steps and {structure['proc_steps']} PROC steps.
            IMPORTANT: Include at least two practical usage examples showing different parameter combinations.
            Return a valid JSON object following this exact structure:
            {json.dumps(base_content, indent=2)}"""
        }, {
            "role": "user",
            "content": f"Please analyze this SAS macro code and provide documentation in JSON format:\n{sas_code}"
        }]

        if feedback:
            messages.append({
                "role": "user",
                "content": f"Please update the JSON documentation with these changes: {feedback}"
            })

        logger.debug("Sending request to OpenAI for documentation content")
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.2,
            response_format={"type": "json_object"}
        )

        content = response.choices[0].message.content
        logger.debug(f"OpenAI documentation response: {content}")
        return json.loads(content)

    except Exception as e:
        logger.error(f"Error generating documentation content: {str(e)}")
        raise

def generate_program_header(program_name, programmer_name, project_name, sas_code, program_specs):
    """Generate a standardized program header based on the template"""
    try:
        current_date = datetime.now().strftime("%Y-%m-%d")
        macro_params = extract_macro_parameters(sas_code)
        macro_calls = find_macro_calls(sas_code)

        messages = [{
            "role": "system",
            "content": """You are a SAS macro expert. Analyze the provided SAS macro code and return a JSON response containing the macro's purpose and example usage.
            Response must be in this exact JSON format:
            {
                "purpose": "A clear, concise purpose statement",
                "example": "A practical example of calling the macro"
            }"""
        }, {
            "role": "user",
            "content": f"Please analyze this SAS macro code and provide a JSON response with purpose and example:\n{sas_code}"
        }]

        logger.debug("Sending request to OpenAI for header content")
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.2,
            response_format={"type": "json_object"}
        )

        content = response.choices[0].message.content
        logger.debug(f"OpenAI header response: {content}")
        ai_response = json.loads(content)

        purpose = ai_response.get('purpose', 'Purpose not available')
        macro_example = ai_response.get('example', 'Example not available')
        params_section = "\n".join(f"[{param[0]}/{param[1]}/{param[2]}]" for param in macro_params) if macro_params else "No parameters"

        header = f"""
/****************************************************************************************
**                    Copyright: Bristol-Myers Squibb Company
*****************************************************************************************/        
/****************************************************************************************
@Start

@Program Name:  {program_name} 
@Programmer:    {programmer_name} 
@Date:          {current_date}
@Project:       {project_name}
@Purpose:       {purpose}

{{Macro Called Example:
{macro_example}
}}

{{Program Specifications:   
#Type/Level/Category/Heritage Company
[{program_specs['type']}/{program_specs['level']}/{program_specs['category']}/{program_specs['heritage']}]
}}

{{Functional Specifications:
@Input files:  
@Output Produced: 
@Macros called by: 
@Macros Used: {macro_calls}
}}

{{Macro Parameters:                
#Parameter Name/Default/Description
{params_section}
}}

{{Modification History:  
#Date/Version/Programmer/Description  
[{current_date}/1/{programmer_name}/Initial Version] 
}}

@End
****************************************************************************************/
"""
        logger.debug(f"Generated header:\n{header}")
        return header

    except Exception as e:
        logger.error(f"Error generating header: {str(e)}")
        raise

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview-documentation', methods=['POST'])
def preview_documentation():
    """Generate preview of the documentation for a single macro"""
    try:
        logger.debug("Received preview documentation request")
        data = request.json
        sas_code = data.get('code', '').strip()
        generate_header = data.get('generate_header', False)
        programmer_name = data.get('programmer_name', '')
        project_name = data.get('project_name', '')
        show_comments = data.get('show_comments', False)
        program_specs = data.get('program_specs', {
            'type': 'macro',
            'level': 'global',
            'category': 'Utility',
            'heritage': 'NewCo'
        })

        if not sas_code:
            return jsonify({'error': 'SAS code is required'}), 400

        logger.debug(f"Processing SAS code with header: {generate_header}, comments: {show_comments}")

        # Extract macro and verify
        macro_match = re.search(r'%macro\s+(\w+)', sas_code, re.IGNORECASE)
        if not macro_match:
            return jsonify({'error': 'No valid SAS macro found in the code'}), 400

        macro_name = macro_match.group(1)
        logger.info(f"Processing macro: {macro_name}")

        # Process the code
        processed_code = sas_code

        # Generate header if requested
        header = None
        if generate_header:
            try:
                header = generate_program_header(macro_name, programmer_name, project_name, processed_code, program_specs)
                if header:
                    processed_code = header + "\n" + processed_code
                    logger.debug("Successfully generated and added header")
            except Exception as e:
                logger.error(f"Error generating header: {str(e)}")
                return jsonify({'error': f'Failed to generate header: {str(e)}'}), 500

        # Add intelligent comments if requested
        if show_comments:
            try:
                processed_code = generate_intelligent_comments(processed_code)
                logger.debug("Successfully added intelligent comments")
            except Exception as e:
                logger.error(f"Error generating comments: {str(e)}")
                return jsonify({'error': f'Failed to generate comments: {str(e)}'}), 500

        # Generate documentation content
        try:
            doc_content = generate_doc_content(processed_code)
            preview_html = create_preview_content(doc_content)

            # Always show code section if either header or comments are requested
            should_show_code = generate_header or show_comments

            response_data = {
                'preview_html': preview_html,
                'macro_name': macro_name,
                'content': doc_content,
                'code': processed_code,
                'show_code': should_show_code,
                'header': header
            }
            logger.debug("Successfully generated preview")
            return jsonify([response_data])

        except Exception as e:
            logger.error(f"Error generating documentation: {str(e)}")
            return jsonify({'error': f'Failed to generate documentation: {str(e)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error in preview generation: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/download-documentation', methods=['POST'])
def download_documentation():
    """Handle documentation download for a single macro"""
    try:
        logger.debug("Received download documentation request")
        data = request.json
        if not data or not isinstance(data.get('content'), list):
            return jsonify({'error': 'Invalid request format'}), 400

        doc_data = data['content'][0]
        if not doc_data:
            return jsonify({'error': 'Invalid documentation content'}), 400

        macro_name = doc_data.get('macro_name')
        content = doc_data.get('content')

        if not macro_name or not content:
            return jsonify({'error': 'Missing required documentation content'}), 400

        output_format = data.get('outputFormat', 'rtf').lower()
        temp_dir = tempfile.mkdtemp()
        filename_base = f"{macro_name}_User_Manual"

        try:
            logger.debug(f"Generating documentation in {output_format} format")
            if output_format == 'pptx':
                output_filename = f'{filename_base}.pptx'
                doc_path = os.path.join(temp_dir, output_filename)
                create_presentation(content, doc_path)
            else:
                output_filename = f'{filename_base}.{output_format}'
                doc_path = os.path.join(temp_dir, output_filename)
                logger.debug(f"Creating manual with format: {output_format}")

                # Special handling for RTF format
                if output_format == 'rtf':
                    logger.debug("Using RTF-specific generation process")
                    create_manual(content, doc_path, output_format, macro_name, rtf_mode=True)
                else:
                    create_manual(content, doc_path, output_format, macro_name)

            if not os.path.exists(doc_path):
                logger.error(f"Document file was not created at {doc_path}")
                raise Exception("Document generation failed")

            logger.debug(f"Document generated successfully at {doc_path}")
            logger.debug(f"File size: {os.path.getsize(doc_path)} bytes")

            # Updated MIME types with specific RTF handling
            mime_types = {
                'pdf': 'application/pdf',
                'rtf': 'application/rtf',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'html': 'text/html'
            }

            # Ensure the file is readable
            if not os.access(doc_path, os.R_OK):
                logger.error(f"File {doc_path} is not readable")
                raise Exception("Generated document is not accessible")

            return send_file(
                doc_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype=mime_types.get(output_format, 'application/octet-stream')
            )

        except Exception as e:
            logger.error(f"Error generating document: {str(e)}")
            logger.error(f"Output format: {output_format}")
            logger.error(f"Document path: {doc_path if 'doc_path' in locals() else 'not created'}")

            if 'doc_path' in locals() and os.path.exists(doc_path):
                logger.error(f"File exists but may be corrupted. Size: {os.path.getsize(doc_path)} bytes")

            return jsonify({'error': f'Failed to generate {output_format.upper()} document: {str(e)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error in download: {str(e)}")
        return jsonify({'error': str(e)}), 500

def create_preview_content(doc_content):
    """Generate HTML preview content for documentation"""
    try:
        if not isinstance(doc_content, dict):
            raise ValueError("Invalid documentation content format")

        preview_html = "<div class='documentation-preview'>"

        # Required sections only
        sections = [
            ('Overview', lambda x: f"<p>{x}</p>"),
            ('Syntax', lambda x: f"<pre class='example-code'>{x}</pre>"),
            ('Parameters', lambda x: create_parameters_table(x)),
            ('Key Features and Functionalities', lambda x: create_features_section(x)),
            ('Usage Examples', lambda x: create_examples_section(x)),
            ('Return Values', lambda x: f"<p>{x}</p>"),
            ('Error Handling', lambda x: f"<p>{x}</p>"),
            ('Summary', lambda x: f"<p>{x}</p>")
        ]

        for section_name, formatter in sections:
            content = doc_content.get(section_name, '')
            if content:
                preview_html += f"<h2>{section_name}</h2>{formatter(content)}"

        preview_html += "</div>"
        return preview_html

    except Exception as e:
        logger.error(f"Error creating preview content: {str(e)}")
        return "<p>Error generating preview content. Please try again.</p>"

def create_parameters_table(parameters):
    """Create HTML table for parameters section"""
    if not isinstance(parameters, dict):
        return "<p>No parameters specified</p>"

    headers = parameters.get('table_headers', [])
    rows = parameters.get('table_rows', [])

    if not headers or not rows:
        return "<p>No parameters specified</p>"

    html = """<table class="table table-bordered">
        <thead><tr>"""
    html += ''.join(f"<th>{header}</th>" for header in headers)
    html += "</tr></thead><tbody>"
    html += ''.join(
        f"<tr>{''.join(f'<td>{cell}</td>' for cell in row)}</tr>"
        for row in rows
    )
    html += "</tbody></table>"
    return html

def create_features_section(features):
    """Create HTML for features section"""
    if not isinstance(features, dict):
        return "<p>No features specified</p>"

    html = f"<p>{features.get('main_section', '')}</p>"
    for subsection in features.get('subsections', []):
        if isinstance(subsection, dict):
            title = subsection.get('title', '')
            description = subsection.get('description', '')
            if title and description:
                html += f"<h3>{title}</h3><p>{description}</p>"
    return html

def create_examples_section(examples):
    """Create HTML for examples section"""
    if not examples:
        return "<p>No examples available</p>"

    return ''.join(f"<pre class='example-code'>{example}</pre>" for example in examples)

def analyze_macro_structure(macro_code):
    """Analyze macro content to determine if it contains DATA/PROC steps"""
    try:
        lowercase_code = macro_code.lower()
        data_count = len(re.findall(r'^\s*data\s+(?!_null_)', lowercase_code, re.MULTILINE))
        proc_count = len(re.findall(r'^\s*proc\s+\w+', lowercase_code, re.MULTILINE))

        return {
            'has_data_proc': bool(data_count + proc_count > 0),
            'data_steps': data_count,
            'proc_steps': proc_count
        }
    except Exception as e:
        logger.error(f"Error analyzing macro structure: {str(e)}")
        return {'has_data_proc': False, 'data_steps': 0, 'proc_steps': 0}

def extract_macros(sas_code):
    """Extract all macros from the SAS code"""
    try:
        logger.debug("Attempting to extract macro from code")
        current_macro = ""
        in_macro = False
        macro_level = 0
        macros = []

        for line in sas_code.split('\n'):
            stripped_line = line.strip().lower()

            if '%macro' in stripped_line:
                macro_level += 1
                if not in_macro:
                    in_macro = True
                    current_macro = line + '\n'
                    continue

            if '%mend' in stripped_line:
                macro_level -= 1
                if in_macro:
                    current_macro += line + '\n'
                    if macro_level == 0:
                        macros.append(current_macro.strip())
                        in_macro = False
                        current_macro = ""
                continue

            if in_macro:
                current_macro += line + '\n'

        if in_macro and current_macro:
            macros.append(current_macro.strip())

        return macros if macros else []

    except Exception as e:
        logger.error(f"Error in extract_macros: {str(e)}")
        return []


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)