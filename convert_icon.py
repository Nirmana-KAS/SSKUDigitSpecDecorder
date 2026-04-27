from PIL import Image

SIZES = [(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)]

img = Image.open('icon.png')
img.save('icon.ico', format='ICO', sizes=SIZES)
print(f'icon.ico written with sizes {SIZES}')
