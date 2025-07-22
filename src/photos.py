import os

PHOTOS_DIR_PATH = "../assets"
all_files = os.listdir(PHOTOS_DIR_PATH)
all_images = [image for image in all_files if image.lower().endswith(('.jpg', '.jpeg', '.png'))]
list_of_images_path = [os.path.join(PHOTOS_DIR_PATH, nome) for nome in all_images]
