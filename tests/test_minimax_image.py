import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import requests
from PIL import Image

from app.pipeline import PPTImagePipeline
from app.settings import Settings


class H3SlideTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.output = Path(temp.name) / "slide.png"
        self.settings = Settings(
            app={"output_root": temp.name},
            models={
                "text": {"api_key": "text-key"},
                "image": {
                    "provider": "minimax",
                    "model": "MiniMax-H3-Image",
                    "base_url": "https://api.minimax.io/v1/images/generations",
                    "timeout": 30,
                    "retries": 2,
                },
            },
        )
        self.pipeline = PPTImagePipeline(self.settings)

    def generate(self):
        config = self.pipeline.build_runtime_config(None, None, None, "image-key", None, None)
        self.pipeline._generate_slide_image(config, "Render a slide", self.output)

    @patch("app.minimax_image.requests.get")
    @patch("app.minimax_image.requests.post")
    def test_partial_success_downloads_the_first_usable_slide(self, post, get):
        self.settings.models.image.variants = 2
        post.return_value = Mock(status_code=200)
        post.return_value.json.return_value = {
            "data": [
                {"error": {"code": "1027", "message": "Image withheld"}},
                {"url": "https://example.com/slide.png", "size": "1600x900"},
            ],
            "request_id": "request-one",
        }
        png = io.BytesIO()
        Image.new("RGB", (16, 9)).save(png, format="PNG")
        get.return_value.content = png.getvalue()
        self.generate()
        self.assertEqual(self.output.read_bytes(), png.getvalue())
        post.assert_called_once()
        self.assertEqual(post.call_args.args, (self.settings.image_api_url,))
        request = post.call_args.kwargs
        self.assertEqual(request["headers"], {"Authorization": "Bearer image-key"})
        self.assertEqual(request["timeout"], 600)
        self.assertFalse(request["allow_redirects"])
        self.assertEqual(request["json"], {
            "model": "MiniMax-H3-Image",
            "prompt": self.pipeline._augment_prompt("Render a slide"),
            "n": 2, "area": "4K", "aspect_ratio": "16:9",
            "output_format": "png", "response_format": "url",
        })
        get.assert_called_once_with("https://example.com/slide.png", timeout=600)

    @patch("app.minimax_image.requests.post")
    def test_request_error_reports_service_message_without_retry(self, post):
        image = self.settings.models.image
        image.base_url = "https://api.minimax.cn/v1/images/generations"
        image.model = "MiniMax-H3-Image-Highspeed"
        image.size = "2720x1536"
        image.timeout = 900
        post.return_value = Mock(status_code=422)
        post.return_value.json.return_value = {"error": {"message": "All images withheld"}}
        with self.assertRaisesRegex(ValueError, "HTTP 422: All images withheld"):
            self.generate()
        post.assert_called_once()
        self.assertEqual(post.call_args.args, (image.base_url,))
        self.assertEqual(post.call_args.kwargs["timeout"], 900)
        payload = post.call_args.kwargs["json"]
        self.assertEqual(payload["model"], image.model)
        self.assertEqual(payload["size"], image.size)
        self.assertNotIn("area", payload)
        self.assertNotIn("aspect_ratio", payload)
        self.assertFalse(self.output.exists())

    @patch("app.minimax_image.requests.post")
    def test_generation_timeout_does_not_resubmit(self, post):
        post.side_effect = requests.Timeout("generation timed out")
        with self.assertRaisesRegex(ValueError, "generation timed out"):
            self.generate()
        post.assert_called_once()

    def test_image_key_is_required(self):
        with self.assertRaisesRegex(ValueError, "Image model api_key"):
            self.pipeline.build_runtime_config(None, None, None, None, None, None)

    def test_default_image_provider_keeps_its_existing_configuration(self):
        settings = Settings(app={"output_root": str(self.output.parent)}, models={"text": {"api_key": "text-key"}})
        pipeline = PPTImagePipeline(settings)
        config = pipeline.build_runtime_config(None, None, None, None, None, None)
        self.assertEqual(config.image_api_key, "text-key")
        with patch.object(pipeline, "_generate_http_slide_image") as generate:
            pipeline._generate_slide_image(config, "Render a slide", self.output)
        generate.assert_called_once()
        self.assertEqual(generate.call_args.kwargs["image_api_url"], settings.image_api_url)
