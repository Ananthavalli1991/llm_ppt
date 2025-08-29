import os
import json
import io
import mimetypes
from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import openai  # Using OpenAI as a concrete example

# Create a Flask application instance
app = Flask(__name__)
# Configure a secure upload folder for temporary files
app.config['UPLOAD_FOLDER'] = 'temp_uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
# Define allowed file extensions for templates
ALLOWED_EXTENSIONS = {'pptx', 'potx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# A helper function to extract key styles and images from the template.
# This is a simplified implementation. A production-grade version would be much more robust.
def extract_template_info(template_path):
    """
    Extracts essential style information from the PowerPoint template.
    """
    try:
        prs = Presentation(template_path)
        styles = {}

        # Attempt to get a title slide and content slide layout for style sampling
        try:
            title_layout = prs.slide_layouts[0]
            # Extracting font name from the title placeholder
            title_placeholder = title_layout.placeholders[0]
            if title_placeholder.has_text_frame:
                styles['title_font'] = title_placeholder.text_frame.paragraphs[0].font.name
        except (IndexError, AttributeError):
            styles['title_font'] = 'Calibri'

        try:
            content_layout = prs.slide_layouts[1]
            # Extracting font name from the content placeholder
            content_placeholder = content_layout.placeholders[1]
            if content_placeholder.has_text_frame:
                styles['content_font'] = content_placeholder.text_frame.paragraphs[0].font.name
        except (IndexError, AttributeError):
            styles['content_font'] = 'Calibri'

        # This is where a more advanced implementation would also extract colors,
        # and images from master slides and layouts. For this example, we'll
        # just use the extracted fonts.

        return styles
    except Exception as e:
        print(f"Error extracting template info: {e}")
        return {}

# A helper function to interact with the LLM and get structured slide content.
def generate_slide_content(text_input, llm_guidance, api_key):
    """
    Sends a request to the LLM API to structure the input text into a presentation.
    The LLM is prompted to return a specific JSON format.
    """
    # Define a clear system prompt to guide the LLM's behavior and persona
    system_prompt = "You are an expert presentation designer and content strategist. Your task is to transform raw text into a well-structured and logical presentation outline. Your output MUST be in a valid JSON format as described. Do not include any additional text or explanations outside of the JSON."

    # The user prompt contains the content and the guidance
    user_prompt = f"""
    Transform the following text into a presentation outline. Each slide should have a concise title and a list of key points.

    Optional Guidance: "{llm_guidance}"

    Input Text:
    {text_input}

    Strict JSON Output Format:
    {{
        "presentation_title": "Presentation Title",
        "slides": [
            {{
                "slide_title": "Slide Title 1",
                "content_points": ["Point 1", "Point 2", "Point 3"]
            }},
            {{
                "slide_title": "Slide Title 2",
                "content_points": ["Point 1", "Point 2"]
            }}
        ]
    }}
    """
    
    try:
        # Use a `try...except` block for robust error handling
        client = openai.OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model="gpt-4-turbo-preview",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.4
        )
        # Parse the JSON response from the LLM
        slide_data = json.loads(response.choices[0].message.content)
        return slide_data
    except Exception as e:
        print(f"Error calling LLM API: {e}")
        return None
from flask import send_from_directory

@app.route('/')
def serve_index():
    return send_from_directory(os.getcwd(), 'frontend/index.html')
# The main API endpoint to handle the presentation generation request.
@app.route('/generate_pptx', methods=['POST'])
def generate_pptx():
    """
    Main endpoint that orchestrates the entire process:
    1. Receives form data and uploaded file.
    2. Calls the LLM for content structure.
    3. Creates a new presentation using the provided template.
    4. Populates the presentation with LLM-generated content.
    5. Sends the final file back to the user.
    """
    if 'template' not in request.files:
        return jsonify({'error': 'No template file provided'}), 400

    template_file = request.files['template']
    text_input = request.form.get('text_input')
    llm_guidance = request.form.get('llm_guidance', '')
    api_key = request.form.get('api_key')

    # Basic input validation
    if not all([text_input, api_key, template_file]):
        return jsonify({'error': 'Missing required fields (text, API key, or template)'}), 400

    if template_file and allowed_file(template_file.filename):
        # Save the uploaded template file securely and temporarily
        filename = secure_filename(template_file.filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        template_file.save(template_path)
    else:
        return jsonify({'error': 'Invalid file type. Please upload a .pptx or .potx file.'}), 400

    try:
        # Step 1: Generate slide content using the LLM
        slide_data = generate_slide_content(text_input, llm_guidance, api_key)
        if not slide_data:
            return jsonify({'error': 'Failed to generate content from LLM. Please check your API key and input text.'}), 500

        # Step 2: Create the presentation based on the template
        prs = Presentation(template_path)
        template_styles = extract_template_info(template_path)

        # Add a title slide
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title_shape = slide.shapes.title
        title_shape.text = slide_data['presentation_title']
        
        # Apply the font from the template if available
        if 'title_font' in template_styles:
            title_shape.text_frame.paragraphs[0].font.name = template_styles['title_font']

        # Add content slides
        content_slide_layout = prs.slide_layouts[1] # A typical 'Title and Content' layout
        for slide_info in slide_data['slides']:
            slide = prs.slides.add_slide(content_slide_layout)
            title_shape = slide.shapes.title
            body_shape = slide.placeholders[1]

            title_shape.text = slide_info['slide_title']
            
            # Apply the font from the template if available
            if 'title_font' in template_styles:
                title_shape.text_frame.paragraphs[0].font.name = template_styles['title_font']
                
            # Populate the content placeholder with bullet points
            tf = body_shape.text_frame
            tf.clear()
            for point in slide_info['content_points']:
                p = tf.add_paragraph()
                p.text = point
                p.level = 0
                if 'content_font' in template_styles:
                    p.font.name = template_styles['content_font']
            
    except Exception as e:
        print(f"Error creating PowerPoint: {e}")
        return jsonify({'error': f'Failed to create presentation file: {e}'}), 500
    finally:
        # Clean up the uploaded template file to save disk space
        os.remove(template_path)

    # Step 3: Save the generated presentation to a buffer and serve it
    try:
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        
        # Return the file for download
        return send_file(
            output,
            as_attachment=True,
            download_name="generated_presentation.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        print(f"Error serving file: {e}")
        return jsonify({'error': f'Failed to serve the generated file: {e}'}), 500

if __name__ == '__main__':
    # Start the Flask application in debug mode
    app.run(debug=True)
