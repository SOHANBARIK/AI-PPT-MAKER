import streamlit as st
import base64
import pptx
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt
import os
from dotenv import load_dotenv
from openai import OpenAI

# Load API key
load_dotenv()
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


def generate_slide_titles(slide_title, num_slides=5):
    """Generates titles/bullets for slides"""
    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[{
            "role": "system", "content": "You are a helpful assistant."
        }, {
            "role": "user",
            "content": f"Generate {num_slides} rigid titles for a PowerPoint slide titled: '{slide_title}'. "
                       f"Be very specific and titles must be the important ones."
        }],
        temperature=0.7,
    )
    raw_titles = response.choices[0].message.content.strip().split("\n")
    return [t.lstrip("0123456789. -") for t in raw_titles if t.strip()][:num_slides]


def generate_slide_content(slide_title, style="bullets"):
    """Generates content for a given slide"""
    if style == "bullets":
        user_prompt = f"Generate 3-5 concise bullet points for a PowerPoint slide titled: '{slide_title}'."
    else:
        user_prompt = f"Generate a concise paragraph (4–5 sentences) for a slide titled: '{slide_title}'."

    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct",
        messages=[{"role": "system", "content": "You are a helpful assistant."},
                  {"role": "user", "content": user_prompt}],
        temperature=0.7,
    )
    return response.choices[0].message.content.strip()


def create_presentation(topic, slide_titles, slide_contents, style,
                        first_page_title_size, global_title_size, content_font_size,
                        font_title, font_content):
    """Builds and saves a PowerPoint presentation"""
    prs = pptx.Presentation()

    # ----- Title Slide -----
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_shape = title_slide.shapes.title
    title_shape.text = topic

    text_frame = title_shape.text_frame
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.size = Pt(first_page_title_size)
            run.font.bold = True
            run.font.name = font_title

    # ----- Content Slides -----
    for slide_title, slide_content in zip(slide_titles, slide_contents):
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Title
        title_shape = slide.shapes.title
        title_shape.text = slide_title
        for p in title_shape.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(global_title_size)
                run.font.name = font_title

        # Content
        text_frame = slide.shapes.placeholders[1].text_frame
        text_frame.clear()
        if style == "bullets":
            for line in slide_content.split("\n"):
                line = line.strip("-• \t")
                if line:
                    p = text_frame.add_paragraph()
                    p.text = line
                    p.font.size = Pt(content_font_size)
                    p.font.name = font_content
        else:
            p = text_frame.paragraphs[0]
            p.text = slide_content
            p.font.size = Pt(content_font_size)
            p.font.name = font_content

    os.makedirs("generated_ppt", exist_ok=True)
    ppt_path = f"generated_ppt/{topic}_presentation.pptx"
    prs.save(ppt_path)
    return ppt_path


def get_ppt_download_link(file_path, topic):
    """Generates download link for PPT"""
    with open(file_path, "rb") as file:
        ppt_contents = file.read()
    b64_ppt = base64.b64encode(ppt_contents).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64_ppt}" download="{topic}_presentation.pptx">Download Presentation</a>'


# -------------------- Streamlit App --------------------

def main():
    st.title("AI-Powered PPT Maker")
    st.subheader("Text to PPT Generation using LLM")

    topic = st.text_input("Enter the topic for the PPT presentation exact as you want to generate:")
    num_slides = st.slider("Number of slides:", 2, 10, 3, 1)
    first_page_title_size = st.slider("First page title font size:", 20, 80, 44, 2)
    global_title_size = st.slider("Slide titles font size:", 20, 60, 32, 2)
    content_font_size = st.slider("Slide content font size:", 10, 40, 16, 2)

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
    generate_button = st.button("Generate Presentation")

    if generate_button and topic:
        st.info("Generating slide titles...")
        slide_titles = generate_slide_titles(topic, num_slides)

        st.info("Generating slide contents...")
        slide_contents = [generate_slide_content(title, style) for title in slide_titles]

        st.info("Creating presentation...")
        ppt_path = create_presentation(
            topic, slide_titles, slide_contents, style,
            first_page_title_size, global_title_size, content_font_size,
            font_title, font_content
        )

        st.success("Presentation generated successfully!")
        st.markdown(get_ppt_download_link(ppt_path, topic), unsafe_allow_html=True)


if __name__ == "__main__":
    main()
