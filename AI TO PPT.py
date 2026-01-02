import streamlit as st
import base64
import pptx
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt, Inches
import os
import requests
from dotenv import load_dotenv
from openai import OpenAI
import re  # Added for Regex

# Load environment variables
load_dotenv()
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY")

# Initialize OpenAI client (only for text generation)
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"), base_url="https://openrouter.ai/api/v1")

# -------------------- Helper Functions --------------------
def clean_font_selection(selection, label):
    """Handles font family selection and 'Other' input"""
    if selection == "Other":
        custom_font = st.text_input(f"Enter custom font family for {label}:").strip()
        if not custom_font or custom_font in ["Calibri", "Arial", "Times New Roman", "Verdana", "Tahoma", "Other"]:
            st.warning(f"⚠️ Please enter a valid custom font for {label}. Defaulting to Calibri.")
            return "Calibri"
        return custom_font
    return selection

def clean_text_and_tags(text):
    """
    Removes HTML tags and cleans up whitespace/markdown artifacts.
    Smartly replaces structural tags with newlines to preserve bullet points.
    """
    if not text:
        return ""
    
    # 1. Replace structural closing tags with newlines to preserve formatting
    # This prevents <li>Item 1</li><li>Item 2</li> from becoming "Item 1Item 2"
    text = text.replace('</li>', '\n')
    text = text.replace('</p>', '\n')
    text = text.replace('<br>', '\n')
    text = text.replace('<br/>', '\n')
    text = text.replace('</div>', '\n')

    # 2. Remove all remaining HTML tags (<...>)
    text = re.sub(r'<[^>]+>', '', text)

    # 3. Remove common Markdown artifacts (bold/italic stars)
    text = text.replace('**', '').replace('__', '')

    # 4. Collapse multiple newlines into one and strip whitespace
    text = re.sub(r'\n\s*\n', '\n', text)
    
    return text.strip()

# ------------------ AI FUNCTIONS ------------------ #
def generate_slide_titles(topic, num_slides=5):
    """Generate slide titles based on topic and number of slides."""
    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Generate {num_slides} concise slide titles for a PPT on '{topic}'."}
        ],
        temperature=0.7,
        max_tokens=150,
    )
    raw_content = clean_text_and_tags(response.choices[0].message.content) # Clean response immediately
    raw_titles = raw_content.split("\n")
    titles = [t.lstrip("0123456789. -") for t in raw_titles if t.strip()]
    return titles[:num_slides]


def generate_slide_content(slide_title, style="bullets"):
    """Generate content for each slide based on style."""
    if style == "bullets":
        prompt = f"Generate 3-5 concise bullet points for a PowerPoint slide titled: '{slide_title}'."
    else:
        prompt = f"Generate a concise paragraph (4–5 sentences) for a slide titled: '{slide_title}'."

    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[{"role": "system", "content": "You are a helpful assistant."},
                  {"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=200,
    )
    # Clean the content immediately after generation
    return clean_text_and_tags(response.choices[0].message.content)


def fetch_pexels_image(query, save_path):
    """Fetch an image from Pexels API based on query."""
    headers = {"Authorization": PEXELS_API_KEY}
    params = {"query": query, "per_page": 1, "orientation": "landscape"}
    url = "https://api.pexels.com/v1/search"

    try:
        response = requests.get(url, headers=headers, params=params)
        data = response.json()
        if "photos" in data and len(data["photos"]) > 0:
            image_url = data["photos"][0]["src"]["large"]
            img_data = requests.get(image_url).content
            with open(save_path, "wb") as f:
                f.write(img_data)
            return save_path
        else:
            # print("No Pexels image found for:", query)
            return None
    except Exception as e:
        # print("Pexels API error:", e)
        return None

# ------------------ PPT CREATION ------------------ #
def create_presentation(topic, slide_titles, slide_contents, style,
                        first_page_title_size, global_title_size, content_font_size, font_title, font_content, add_images=False):
    """Create a PPT with generated titles, content, and optional images."""
    prs = pptx.Presentation()
    slide_layout = prs.slide_layouts[1]

    # Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_shape = title_slide.shapes.title
    # Final safety clean for topic
    title_shape.text = clean_text_and_tags(topic)

    text_frame = title_shape.text_frame
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.size = Pt(first_page_title_size)
            run.font.bold = True
            run.font.name = font_title

    # Content Slides
    for idx, (slide_title, slide_content) in enumerate(zip(slide_titles, slide_contents), start=1):
        slide = prs.slides.add_slide(slide_layout)

        # Slide Title
        title_shape = slide.shapes.title
        title_shape.text = slide_title # Already cleaned in generation step
        for p in title_shape.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(global_title_size)
                run.font.name = font_title

        # Content Area
        text_frame = slide.shapes.placeholders[1].text_frame
        text_frame.clear()

        # The content is already cleaned by generate_slide_content, 
        # so we just need to split it for bullets.
        if style == "bullets":
            for line in slide_content.split("\n"):
                line = line.strip("-• \t")
                if line:
                    p = text_frame.add_paragraph()
                    p.text = line
                    p.level = 0
                    p.font.size = Pt(content_font_size)
                    p.font.name = font_content
        else:  # Paragraph
            p = text_frame.paragraphs[0]
            p.text = slide_content
            p.font.size = Pt(content_font_size)

        # Optional: Add Pexels image
        if add_images and PEXELS_API_KEY:
            os.makedirs("generated_ppt/images", exist_ok=True)
            image_path = f"generated_ppt/images/slide_{idx}.jpg"
            if fetch_pexels_image(slide_title, image_path):
                try:
                    slide.shapes.add_picture(image_path, Inches(5), Inches(2), Inches(3), Inches(3))
                except Exception as e:
                    pass
                    # print("Image placement failed:", e)

    # Save PPT
    os.makedirs("generated_ppt", exist_ok=True)
    ppt_filename = f"generated_ppt/{topic.replace(' ', '_')}_presentation.pptx"
    prs.save(ppt_filename)
    return ppt_filename


# ------------------ STREAMLIT APP ------------------ #
def get_ppt_download_link(ppt_filename):
    """Generate download link for the PPT."""
    with open(ppt_filename, "rb") as file:
        ppt_contents = file.read()
    b64_ppt = base64.b64encode(ppt_contents).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64_ppt}" download="{os.path.basename(ppt_filename)}">Download Presentation</a>'


def main():
    st.title("📊 AI-Powered PPT Maker")
    st.subheader("Generate PowerPoint Presentations with Pexels Images")

    # User inputs
    topic = st.text_input("Enter the topic for your presentation:")
    num_slides = st.slider("Number of slides", min_value=2, max_value=10, value=3, step=1)
    first_page_title_size = st.slider("Title font size (pt):", min_value=10, max_value=60, value=32, step=2)
    global_title_size = st.slider("Slide titles font size:", 20, 60, 32, 2)
    content_font_size = st.slider("Content font size (pt):", min_value=8, max_value=40, value=16, step=1)
    st.write("Font Settings")
    same_font = st.radio("Use the same font for Title & Content?", ["Yes", "No"], index=0)

    if same_font == "Yes":
        font = st.selectbox("Choose font family:", ["Calibri", "Arial", "Times New Roman", "Verdana", "Tahoma", "Other"])
        font_title = font_content = clean_font_selection(font, "Slides")
    else:
        font_title = st.selectbox("Font family for Titles:", ["Calibri", "Arial", "Times New Roman", "Verdana", "Tahoma", "Other"])
        font_title = clean_font_selection(font_title, "Titles")

        font_content = st.selectbox("Font family for Content:", ["Calibri", "Arial", "Times New Roman", "Verdana", "Tahoma", "Other"])
        font_content = clean_font_selection(font_content, "Content")

    style = st.radio("Content style:", ["bullets", "paragraph"])
    add_images = st.checkbox("Include Pexels images in slides")
    generate_button = st.button("🚀 Generate Presentation")

    if generate_button and topic:
        st.info("🔍 Generating slide titles...")
        slide_titles = generate_slide_titles(topic, num_slides)
        st.write("✅ Slide titles:", slide_titles)

        st.info("📝 Generating slide contents...")
        slide_contents = [generate_slide_content(title, style) for title in slide_titles]

        st.info("📂 Creating presentation...")
        ppt_filename = create_presentation(topic, slide_titles, slide_contents, style,first_page_title_size,global_title_size, content_font_size, font_title, font_content, add_images)

        st.success("🎉 Presentation generated successfully!")
        st.markdown(get_ppt_download_link(ppt_filename), unsafe_allow_html=True)

     # ✅ Custom dark theme + footer
    st.markdown("""
        <style>
            /* Set dark theme */
            body {
                background-color: rgb(1, 1, 1);
                color: white;
                font-family: 'Arial', sans-serif;
            }
            /* Fixed footer */
            .fixed-bottom {
                position: fixed;
                bottom: 10px;
                left: 50%;
                transform: translateX(-50%);
                font-size: 12px;
                color: white;
                background-color: rgb(0, 0, 0);
                padding: 5px 10px;
                border-radius: 5px;
                z-index: 1000;
            }    
        </style>
        <div class="fixed-bottom">made with ❤️ from Sohan </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
