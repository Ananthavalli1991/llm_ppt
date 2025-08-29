#**Your Text, Your Style: LLM-Powered Presentation Generator**
This is a full-stack web application that leverages Large Language Models (LLMs) to automatically generate professional PowerPoint presentations from a simple text or Markdown input. The application allows users to control the output's style and theme by uploading a custom PowerPoint template.

##**Features**
Content Generation: Paste a long text or Markdown document, and the LLM will automatically parse, summarize, and structure it into a coherent presentation.

##**Custom Styling:** Upload an existing .pptx or .potx template. The application uses the template's master slides and theme to ensure the generated presentation matches your desired visual style, including fonts, colors, and layout.

##**Flexible LLM Integration:** The app supports a "bring your own key" model for popular LLM providers: OpenAI, Anthropic, and Gemini. API keys are processed in memory on the back-end and are never stored.

##**Intelligent Layouts:** The LLM-driven process automatically selects the most appropriate slide layouts and determines the number of slides based on the content and a user-provided guidance string (e.g., "investor pitch deck", "technical summary").

##**Optional Speaker Notes:** The LLM can be prompted to generate speaker notes for each slide, helping you prepare for your presentation.

#**Technical Stack**
##Front-End: A simple and responsive web interface built with pure HTML, CSS (Tailwind CSS for utility classes), and JavaScript.

##Back-End: A Python server using the Flask framework, responsible for handling all server-side logic, including:

Securely handling file uploads.

Making API calls to the selected LLM provider.

Dynamically generating PowerPoint files using the python-pptx library.

##LLM Providers: Integration with OpenAI, Anthropic, and Gemini via their official Python client libraries.

##Setup and Usage
Follow these steps to set up and run the application locally.

##Clone the Repository:

git clone [https://github.com/your-username/your-repo-name.git](https://github.com/your-username/your-repo-name.git)
cd your-repo-name

##Create a Virtual Environment: (Recommended)

python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

##Install Dependencies:

pip install -r requirements.txt

##Run the Flask Server:

python app.py

## the App: Navigate to http://127.0.0.1:5000/ in your web browser.

