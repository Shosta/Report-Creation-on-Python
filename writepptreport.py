"""
Write the eIoT report in a Powerpoint file that has one slide.\n

"""


def write_ppt(ppt_file_name):
    from pptx import Presentation

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "Hello, World!"
    subtitle.text = "python-pptx was here!"

    prs.save(ppt_file_name)