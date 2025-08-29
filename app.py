import os
import io
import json
from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import openai
import anthropic # Added for Anthropic API support
import google.generativeai as genai # Added for Gemini API support

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Helper function to generate slide content using LLM
def generate_slide_content(text_input, llm_guidance, api_key, provider):
    prompt = f"""
    You are an expert presentation designer. I will provide a large chunk of text and you will break it down into a logical presentation structure.
    
    Guidance for presentation: "{llm_guidance}"
    
    Please structure your response in a clear, parsable format. Use JSON with the following structure:
    {{
        "title": "Presentation Title",
        "slides": [
            {{
                "title": "Slide 1 Title",
                "content": ["Point 1", "Point 2", "Point 3"],
                "notes": "Speaker notes for slide 1."
            }},
            {{
                "title": "Slide 2 Title",
                "content": "A single paragraph of text.",
                "notes": "Speaker notes for slide 2."
            }}
        ]
    }}
    
    Input Text:
    {text_input}
    """
    
    try:
        slide_data = None
        print(f"Calling {provider} API...")
        if provider == 'openai':
            client = openai.OpenAI(api_key=api_key, timeout=120.0)
            response = client.chat.completions.create(
                model="gpt-4-turbo-preview",
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt}
                ]
            )
            slide_data = json.loads(response.choices[0].message.content)
        elif provider == 'anthropic':
            client = anthropic.Anthropic(api_key=api_key, timeout=120.0)
            response = client.messages.create(
                model="claude-3-5-sonnet-20240620",
                max_tokens=2000,
                system="You are a helpful assistant that responds in JSON.",
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            slide_data = json.loads(response.content[0].text)
        elif provider == 'gemini':
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            response = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
            slide_data = json.loads(response.text)
        
        print(f"Successfully received response from {provider} API.")
        return slide_data

    except Exception as e:
        print(f"Error calling LLM API ({provider}): {e}")
        return None

# Route to serve the HTML file
@app.route('/')
def serve_index():
    return send_from_directory(os.getcwd(), 'frontend/index.html')

# Main API endpoint to generate the presentation
@app.route('/generate_pptx', methods=['POST'])
def generate_pptx():
    # Print received form data for debugging
    print("Received form data:")
    for key, value in request.form.items():
        if key == 'api_key':
            print(f"  {key}: '[REDACTED]'")
        else:
            print(f"  {key}: '{value}'")
    
    if 'template' not in request.files:
        return jsonify({'error': 'No template file provided'}), 400

    template_file = request.files['template']
    text_input = request.form.get('text_input')
    llm_guidance = request.form.get('llm_guidance', '')
    api_key = request.form.get('api_key')
    provider = request.form.get('provider')

    if not text_input or not api_key or not provider:
        return jsonify({'error': 'Missing required fields. Please fill out all text fields and select a provider.'}), 400

    print(f"Saving uploaded template: {template_file.filename}")
    filename = secure_filename(template_file.filename)
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    template_file.save(template_path)

    slide_data = generate_slide_content(text_input, llm_guidance, api_key, provider)
    if not slide_data:
        return jsonify({'error': 'Failed to generate content from LLM'}), 500

    try:
        print("Starting PowerPoint generation...")
        prs = Presentation(template_path)
        
        # Add a title slide
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        title.text = slide_data['title']

        # Add content slides
        content_slide_layout = prs.slide_layouts[1]
        for slide_info in slide_data['slides']:
            slide = prs.slides.add_slide(content_slide_layout)
            title = slide.shapes.title
            body = slide.placeholders[1]
            title.text = slide_info['title']
            
            if isinstance(slide_info['content'], list):
                tf = body.text_frame
                tf.clear()
                for item in slide_info['content']:
                    p = tf.add_paragraph()
                    p.text = item
            else:
                body.text = slide_info['content']

            if 'notes' in slide_info:
                slide.notes_slide.notes_text_frame.text = slide_info['notes']
        
        print("PowerPoint generation complete.")
    except Exception as e:
        print(f"Error creating PowerPoint: {e}")
        return jsonify({'error': 'Failed to create presentation file'}), 500
    finally:
        os.remove(template_path)
        print(f"Removed temporary template file: {template_path}")

    output_path = 'generated_presentation.pptx'
    prs.save(output_path)
    
    return send_file(output_path, as_attachment=True, download_name="generated_presentation.pptx", mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')

if __name__ == '__main__':
    app.run(debug=True)
