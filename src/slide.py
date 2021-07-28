from typing import List

from slide_image import SlideImage


class Slide:
    def __init__(self, slide_number: int, title: str, images: List[SlideImage]):
        self.slide_number = slide_number
        self.title = title
        self.images = images
