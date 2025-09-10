import streamlit as st
import base64
import openai
import pptx
from pptx.util import Inches, Pt
import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"), base_url="https://openrouter.ai/api/v1")

TITLE_FONT_SIZE = Pt(30)
SLIDE_FONT_SIZE = Pt(16)

def generate_slide_titles(topic):
    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Generate 5 slide titles for a PPT on {topic}."}
        ],
        max_tokens=100,
        temperature=0.7,
    )
    return response.choices[0].message.content.strip().split("\n")

def generate_slide_content(slide_title):
    """Generates content for a specific slide title"""
    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Generate content for a slide titled: '{slide_title}'."}
        ],
        temperature=0.7,
    )
    return response.choices[0].message.content.strip()

def create_presentation(topic, slide_titles, slide_contents):
    prs = pptx.Presentation()
    slide_layout = prs.slide_layouts[1]

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic

    # Add slides
    for slide_title, slide_content in zip(slide_titles, slide_contents):
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_title
        slide.shapes.placeholders[1].text = slide_content

        # Customize font size
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    paragraph.font.size = SLIDE_FONT_SIZE

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
    generate_button = st.button("Generate Presentation")

    if generate_button and topic:
        st.info("Generating slide titles...")
        slide_titles = generate_slide_titles(topic)
        filtered_slide_titles = [item.strip() for item in slide_titles if item.strip() != ""]
        st.write("Slide titles generated:", filtered_slide_titles)

        st.info("Generating slide contents...")
        slide_contents = [generate_slide_content(title) for title in filtered_slide_titles]

        st.info("Creating presentation...")
        create_presentation(topic, filtered_slide_titles, slide_contents)
        st.success("Presentation generated successfully!")
        st.markdown(get_ppt_download_link(topic), unsafe_allow_html=True)

if __name__ == "__main__":
    main()

