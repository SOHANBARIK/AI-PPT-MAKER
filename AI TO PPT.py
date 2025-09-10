import streamlit as st
import base64
import openai
import pptx
from pptx.util import Inches, Pt
import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

st.markdown(
    """
        <style>
        /* Set dark theme */
        body {
            background-color:rgb(0, 0, 0);
            color: white;
            font-family: 'Arial', sans-serif;
        }

        /* Fixed "Made with ❤️ for Sohan" at the bottom-center */
        .fixed-bottom {
            position: fixed;
            bottom: 10px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 12px;
            color: white;
            background-color:rgb(0, 0, 0);
            padding: 5px 10px;
            border-radius: 5px;
            z-index: 1000;
        }

    </style>
    <div class="fixed-bottom">made with ❤️ from Sohan </div>
    """,
    unsafe_allow_html=True
)
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"), base_url="https://openrouter.ai/api/v1")

TITLE_FONT_SIZE = Pt(30)
SLIDE_FONT_SIZE = Pt(16)

def generate_slide_titles(topic, num_slides):
    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Generate {num_slides} slide titles for a PPT on {topic}."}
        ],
        max_tokens=100,
        temperature=0.7,
    )
    return response.choices[0].message.content.strip().split("\n")

def generate_slide_content(slide_title, style="bullets"):
    """Generates content for a specific slide title"""
    if style == "bullets":
        user_prompt = (
            f"Generate 3-5 concise bullet points for a PowerPoint slide titled: '{slide_title}'. "
            f"Keep each bullet point short and clear."
        )
    else:  # paragraph
        user_prompt = (
            f"Generate a concise paragraph (4–5 sentences) for a PowerPoint slide titled: '{slide_title}'. "
            f"Keep it clear and professional."
        )

    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.7,
    )
    return response.choices[0].message.content.strip()

def create_presentation(topic, slide_titles, slide_contents, style="bullets", font_size=16):
    prs = pptx.Presentation()
    slide_layout = prs.slide_layouts[1]

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic

    # Add slides with bullet points
    for slide_title, slide_content in zip(slide_titles, slide_contents):
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_title

        # Put bullets in placeholder (content area)
        text_frame = slide.shapes.placeholders[1].text_frame
        text_frame.clear()
 # Put bullets in placeholder (content area)
        if style == "bullets":
            for line in slide_content.split("\n"):
                line = line.strip("-• \t")
                if line:
                    p = text_frame.add_paragraph()
                    p.text = line
                    p.level = 0
                    p.font.size = Pt(font_size)
        else:  # paragraph style
            p = text_frame.paragraphs[0]
            p.text = slide_content
            p.font.size = Pt(font_size)
        # Customize font size

    # Ensure output directory exists
    os.makedirs("generated_ppt", exist_ok=True)
    prs.save(f"generated_ppt/{topic}_presentation.pptx")

def get_ppt_download_link(topic):
    ppt_filename = f"generated_ppt/{topic}_presentation.pptx"
    with open(ppt_filename, "rb") as file:
        ppt_contents = file.read()
    b64_ppt = base64.b64encode(ppt_contents).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64_ppt}" download="{topic}_presentation.pptx">Download Presentation</a>'

def main():
    st.title("AI-Powered PPT Maker")
    st.subheader("Text to PPT Generation using LLM")
    topic = st.text_input("Enter the topic you want to generate the PPT presentation on:")
    num_slides= st.slider("Select number of slides", min_value=2, max_value=10, value=3,step=1)
    font_size=st.slider("Select font size:", min_value=10, max_value=40, value=16, step=1)
    style=st.radio("Select content style:",["bullets", "paragraph"])
    generate_button = st.button("Generate Presentation")

    if generate_button and topic:
        st.info("Generating slide titles...")
        slide_titles = generate_slide_titles(topic, num_slides)
        filtered_slide_titles = [item.strip() for item in slide_titles if item.strip() != ""]
        st.write("Slide titles generated:", filtered_slide_titles)

        st.info("Generating slide contents...")
        slide_contents = [generate_slide_content(title) for title in filtered_slide_titles]

        st.info("Creating presentation...")
        create_presentation(topic, filtered_slide_titles, slide_contents, style, font_size)
        st.success("Presentation generated successfully!")
        st.markdown(get_ppt_download_link(topic), unsafe_allow_html=True)

if __name__ == "__main__":
    main()










