from typing import List


class SlideImage:
    STATUS_OK = "ok"
    STATUS_EXTENSION_NOT_SUPPORTED = "extension_not_supported"
    STATUS_CONVERSION_FAILED = "conversion_failed"
    STATUS_RESIZING_FAILED = "resizing_failed"
    STATUS_API_REQUEST_FAILED = "api_request_failed"
    STATUS_NO_SOURCES_FOUND = "no_sources_found"

    def __init__(self, internal_img_pah: str, website_sources: List[str], image_sources: List[str], status):
        self.internal_img_path = internal_img_pah
        self.website_sources = website_sources
        self.image_sources = image_sources
        self.status = status
