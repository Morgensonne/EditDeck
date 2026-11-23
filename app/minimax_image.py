from typing import Any

import requests


def generate_h3_image(
    *, endpoint: str, api_key: str, model: str, prompt: str,
    size: str, variants: int, timeout: int,
) -> tuple[bytes, dict[str, Any]]:
    if endpoint.rstrip("/") not in {
        "https://api.minimax.io/v1/images/generations",
        "https://api.minimax.cn/v1/images/generations",
    }:
        raise ValueError("Set models.image.base_url to the H3 generation endpoint for your API key.")

    payload = {
        "model": model,
        "prompt": prompt,
        "n": variants,
        "output_format": "png",
        "response_format": "url",
    }
    if size in {"1K", "2K", "4K"}:
        payload.update(aspect_ratio="16:9", area=size)
    elif size and size != "auto":
        payload["size"] = size
    else:
        payload["aspect_ratio"] = "16:9"

    response = requests.post(
        endpoint.rstrip("/"),
        headers={"Authorization": f"Bearer {api_key}"},
        json=payload,
        timeout=timeout,
        allow_redirects=False,
    )
    if response.status_code != 200:
        try:
            message = response.json()["error"]["message"]
        except (ValueError, KeyError, TypeError):
            message = response.reason
        raise ValueError(f"H3 image generation returned HTTP {response.status_code}: {message}")

    result = response.json()
    errors = []
    for image in result.get("data", []):
        if image.get("error"):
            errors.append(image["error"].get("message", "Image generation failed"))
        elif image.get("url"):
            download = requests.get(image["url"], timeout=timeout)
            download.raise_for_status()
            if not download.content:
                raise ValueError("H3 image download was empty.")
            return download.content, {"request_id": result.get("request_id"), "size": image.get("size")}

    raise ValueError("H3 returned no slide image: " + (", ".join(errors) or "missing image URL"))
