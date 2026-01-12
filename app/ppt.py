from pptx import Presentation
import pandas as pd

def generate_ppt_from_csv(csv_path, out_pptx):
    df = pd.read_csv(csv_path)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = "InsightDeck Summary"

    body = slide.placeholders[1].text_frame
    body.text = f"Rows: {len(df)}"

    for col in df.columns[:5]:
        body.add_paragraph().text = f"- {col}"

    prs.save(out_pptx)
