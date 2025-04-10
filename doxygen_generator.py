import re
import logging
from datetime import datetime
from openai import OpenAI
import json
import os

logger = logging.getLogger(__name__)
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

def generate_doxygen_header(sas_code, programmer_name):
    """Generate a Doxygen-style header for a SAS macro"""
    try:
        # Extract macro name
        macro_match = re.search(r'%macro\s+(\w+)', sas_code, re.IGNORECASE)
        if not macro_match:
            raise ValueError("Could not find macro name in the code")
        macro_name = macro_match.group(1)

        # Extract parameters
        macro_params = extract_macro_parameters(sas_code)
        
        # Find macro dependencies
        macro_calls = find_macro_calls(sas_code)
        dependent_macros = macro_calls.split(", ") if macro_calls else []

        # Find dataset references
        input_datasets = re.findall(r'set\s+(\w+\.\w+)|from\s+(\w+\.\w+)', sas_code, re.IGNORECASE)
        output_datasets = re.findall(r'data\s+(\w+\.\w+)', sas_code, re.IGNORECASE)

        # Flatten and clean dataset lists
        input_datasets = list(set([ds for tuple_ds in input_datasets for ds in tuple_ds if ds]))
        output_datasets = list(set(output_datasets))

        messages = [{
            "role": "system",
            "content": """You are a SAS macro expert. Analyze the provided SAS macro code and return a JSON response containing key information for a Doxygen header.
            Response must be in this exact JSON format:
            {
                "brief": "One-sentence functional description",
                "details": "Extended markdown-formatted explanation of purpose, key functionalities, and usage context",
                "return": "Explanation of return value/output",
                "warning": "Critical usage constraints if any",
                "note": "Important implementation details if any",
                "todo": "Suggested future enhancements if any"
            }"""
        }, {
            "role": "user",
            "content": f"Please analyze this SAS macro code and provide a JSON response:\n{sas_code}"
        }]

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.2,
            response_format={"type": "json_object"}
        )

        content = json.loads(response.choices[0].message.content)

        # Build the Doxygen header
        header = f"""/**
@file {macro_name}.sas
@brief {content.get('brief', 'Brief description not available')}

@details
{content.get('details', 'Detailed description not available')}

"""
        # Add parameters section
        if macro_params:
            for param in macro_params:
                header += f"@param {param[0]} {param[2]}\n"

        header += f"\n@return {content.get('return', 'Return value description not available')}\n"
        header += f"@version 1.0\n"
        header += f"@author {programmer_name}\n\n"

        # Add data inputs section if found
        if input_datasets:
            header += "<h4>Data Inputs</h4>\n"
            for dataset in input_datasets:
                header += f"@li {dataset}\n"
            header += "\n"

        # Add data outputs section if found
        if output_datasets:
            header += "<h4>Data Outputs</h4>\n"
            for dataset in output_datasets:
                header += f"@li {dataset}\n"
            header += "\n"

        # Add SAS Macros section if dependencies found
        if dependent_macros:
            header += "<h4>SAS Macros</h4>\n"
            for macro in dependent_macros:
                header += f"@li {macro.strip('%')}.sas\n"
            header += "\n"

        # Add warning if available
        if content.get('warning'):
            header += f"@warning {content.get('warning')}\n"

        # Add note if available
        if content.get('note'):
            header += f"@note {content.get('note')}\n"

        # Add todo if available
        if content.get('todo'):
            header += f"@todo {content.get('todo')}\n"

        # Add maintenance section
        current_date = datetime.now().strftime("%Y-%m-%d")
        header += f"""
@maintenance
- {current_date}: Initial implementation [{programmer_name}]
*/"""

        return header

    except Exception as e:
        logger.error(f"Error generating Doxygen header: {str(e)}")
        raise