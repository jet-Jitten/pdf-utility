from PIL import Image
from pathlib import Path

def images_to_pdf(output_pdf, image_files):
    images = []

    for img_file in image_files:
        img = Image.open(img_file)

        # Convert images with transparency to RGB
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        images.append(img)

    if not images:
        raise ValueError("No images provided")

    output_path = Path(output_pdf)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    first_image = images[0]
    rest_images = images[1:]

    first_image.save(
        output_path,
        save_all=True,
        append_images=rest_images
    )
