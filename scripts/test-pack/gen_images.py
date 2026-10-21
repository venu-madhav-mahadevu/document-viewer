# JPEG and TIFF test files, plus shared pictures used inside the Word/PowerPoint/PDF tests.
import os
from tp_common import OUT, ASSETS, FONTS, Registry
from PIL import Image, ImageDraw, ImageFont

def font(size, bold=False):
    return ImageFont.truetype(os.path.join(FONTS, 'arialbd.ttf' if bold else 'arial.ttf'), size)

def test_chart(w, h, title, lines, accent=(31, 78, 121)):
    """A labeled test chart: title, notes, color bars, gray ramp, fine lines, corner markers."""
    img = Image.new('RGB', (w, h), 'white')
    d = ImageDraw.Draw(img)
    s = w / 1600
    d.rectangle([0, 0, w, int(110 * s)], fill=accent)
    d.text((int(30 * s), int(25 * s)), title, font=font(int(52 * s), True), fill='white')
    y = int(140 * s)
    for line in lines:
        d.text((int(30 * s), y), line, font=font(int(34 * s)), fill=(32, 33, 36))
        y += int(48 * s)
    # color bars: names written on them
    bars = [('RED', (255, 0, 0)), ('GREEN', (0, 176, 80)), ('BLUE', (0, 112, 192)), ('YELLOW', (255, 255, 0)),
            ('MAGENTA', (255, 0, 255)), ('CYAN', (0, 255, 255)), ('BLACK', (0, 0, 0)), ('WHITE', (255, 255, 255))]
    top = y + int(20 * s)
    bw = (w - int(60 * s)) // len(bars)
    for i, (name, col) in enumerate(bars):
        x0 = int(30 * s) + i * bw
        d.rectangle([x0, top, x0 + bw - 4, top + int(160 * s)], fill=col, outline=(128, 128, 128))
        d.text((x0 + 8, top + int(165 * s)), name, font=font(int(24 * s)), fill=(32, 33, 36))
    # gray ramp
    rt = top + int(220 * s)
    for i in range(32):
        g = int(i * 255 / 31)
        x0 = int(30 * s) + i * (w - int(60 * s)) // 32
        d.rectangle([x0, rt, x0 + (w - int(60 * s)) // 32, rt + int(70 * s)], fill=(g, g, g))
    # fine lines (sharpness)
    lt = rt + int(100 * s)
    for i in range(0, int(400 * s), 4):
        d.line([int(30 * s) + i, lt, int(30 * s) + i, lt + int(80 * s)], fill='black', width=1)
    d.text((int(30 * s) + int(420 * s), lt + int(20 * s)), '<- fine lines must look sharp, not blurry', font=font(int(28 * s)), fill=(32, 33, 36))
    # corner markers: orientation check
    m = int(70 * s)
    for (x, y0, label) in [(0, h - m, 'BOTTOM-LEFT'), (w - int(330 * s), h - m, 'BOTTOM-RIGHT')]:
        d.rectangle([x, y0, x + int(330 * s), h], fill=accent)
        d.text((x + 10, y0 + 12), label, font=font(int(32 * s), True), fill='white')
    return img

def main():
    R = Registry('images')

    # pictures reused by other generators
    pic = test_chart(1200, 800, 'Embedded picture', ['This picture is embedded in the document.', 'Colors and text must look like this.'])
    pic.save(os.path.join(ASSETS, 'picture.png'))
    pic.convert('RGB').save(os.path.join(ASSETS, 'picture.jpg'), quality=90)
    small = test_chart(600, 400, 'Picture 2', ['Second picture'], accent=(112, 48, 160))
    small.save(os.path.join(ASSETS, 'picture2.png'))

    # ---------- JPG / JPEG
    f = 'test-image.jpg'
    img = test_chart(1600, 1000, 'JPG-01..05  Test chart (.jpg)', [
        R.add('JPG-01', f, 'Whole image', 'Image opens', 'Image shows fully, not cut off'),
        R.add('JPG-02', f, 'Color bars', 'Colors', 'Each bar matches the name written under it'),
        R.add('JPG-03', f, 'Gray ramp', 'Gray levels', '32 steps from black to white, no color tint'),
        R.add('JPG-04', f, 'Fine lines', 'Sharpness', 'Thin black lines are sharp, not blurred'),
        R.add('JPG-05', f, 'Bottom corners', 'Orientation', 'BOTTOM-LEFT is bottom left, BOTTOM-RIGHT is bottom right'),
    ])
    img.save(os.path.join(OUT, f), quality=92)

    f = 'test-image.jpeg'
    test_chart(1600, 1000, 'JPEG-01  Same test chart, .jpeg extension', [
        R.add('JPEG-01', f, 'Whole image', '.jpeg extension', 'Opens exactly like test-image.jpg'),
        R.add('JPEG-02', f, 'Zoom', 'Zoom in and out', 'Image zooms smoothly; fine lines stay sharp when zoomed in'),
    ], accent=(0, 128, 138)).save(os.path.join(OUT, f), quality=92)

    f = 'test-image-progressive.jpeg'
    test_chart(1600, 1000, 'JPEG-03  Progressive JPEG', [
        R.add('JPEG-03', f, 'Whole image', 'Progressive JPEG (web-optimised)', 'Opens the same as a normal JPEG'),
    ], accent=(84, 130, 53)).save(os.path.join(OUT, f), quality=90, progressive=True, optimize=True)

    f = 'test-image-rotated-exif.jpg'
    up = test_chart(1600, 1000, 'JPG-06  Phone photo rotation', [
        R.add('JPG-06', f, 'Whole image', 'Rotation from camera data (EXIF orientation 6)', 'This text reads normally (not sideways). BOTTOM corners are at the bottom'),
        'The pixels are stored sideways; the file says "rotate 90°".',
    ], accent=(192, 0, 0))
    stored = up.rotate(90, expand=True)  # pixels stored rotated; EXIF tells viewers to turn them back
    exif = Image.Exif(); exif[0x0112] = 6
    stored.save(os.path.join(OUT, f), quality=90, exif=exif)

    f = 'test-image-cmyk.jpg'
    test_chart(1600, 1000, 'JPG-07  CMYK JPEG (print / scanner)', [
        R.add('JPG-07', f, 'Whole image', 'CMYK color JPEG', 'Colors look normal (red is red, not inverted or dark)', '⚠️'),
    ], accent=(191, 144, 0)).convert('CMYK').save(os.path.join(OUT, f), quality=90)

    f = 'test-image-large.jpg'
    test_chart(8000, 5000, 'JPG-08  Very large photo (8000 x 5000)', [
        R.add('JPG-08', f, 'Whole image', 'Very large image (40 megapixels)', 'Opens within a few seconds; zoom works; page stays responsive'),
    ]).save(os.path.join(OUT, f), quality=85)

    # ---------- TIFF
    def tiff_page(n, total, title, lines, mode='RGB'):
        img = test_chart(1700, 2200, title, lines + [f'Page {n} of {total}'])
        d = ImageDraw.Draw(img)
        d.text((60, 2000), f'PAGE {n} / {total}', font=font(90, True), fill=(31, 78, 121))
        return img.convert(mode)

    f = 'test-scan-multipage.tif'
    pages = []
    for n in range(1, 6):
        lines = []
        if n == 1:
            lines = [R.add('TIF-01', f, 'All pages', 'Multi-page TIFF (5 pages)', 'Page bar shows 5 pages; each page says PAGE n / 5'),
                     R.add('TIF-02', f, 'Thumbnails', 'Page thumbnails', '5 thumbnails, in order 1 to 5')]
        if n == 3:
            lines = [R.add('TIF-03', f, 'Page 3', 'Go to page 3 from the page bar', 'Page 3 is shown')]
        pages.append(tiff_page(n, 5, f'TIF  Multi-page scan  (page {n})', lines))
    pages[0].save(os.path.join(OUT, f), save_all=True, append_images=pages[1:], compression='tiff_lzw')

    variants = [
        ('test-tiff-uncompressed.tiff', 'TIFF-01', 'No compression', 'RGB', 'raw', {}),
        ('test-tiff-lzw.tiff', 'TIFF-02', 'LZW compression', 'RGB', 'tiff_lzw', {}),
        ('test-tiff-packbits.tiff', 'TIFF-03', 'PackBits compression', 'RGB', 'packbits', {}),
        ('test-tiff-deflate.tiff', 'TIFF-04', 'Deflate (ZIP) compression', 'RGB', 'tiff_adobe_deflate', {}),
        ('test-tiff-jpeg.tif', 'TIF-04', 'JPEG compression (color scan)', 'RGB', 'jpeg', {'quality': 90}),
        ('test-tiff-fax-g4.tif', 'TIF-05', 'Black and white fax (CCITT Group 4)', '1', 'group4', {}),
        ('test-tiff-fax-g3.tif', 'TIF-06', 'Black and white fax (CCITT Group 3)', '1', 'group3', {}),
        ('test-tiff-grayscale.tiff', 'TIFF-05', 'Grayscale 8-bit', 'L', 'tiff_lzw', {}),
        ('test-tiff-cmyk.tiff', 'TIFF-06', 'CMYK color', 'CMYK', 'tiff_lzw', {}),
    ]
    for file, tid, what, mode, comp, extra in variants:
        expected = 'Text is readable; black and white only' if mode == '1' else ('Gray levels, no color' if mode == 'L' else 'Colors match the bar names')
        status = '⚠️' if mode == 'CMYK' else ''
        line = R.add(tid, file, 'Whole page', what, f'Opens; {expected}', status)
        tiff_page(1, 1, f'{tid}  {what}', [line], mode).save(os.path.join(OUT, file), compression=comp, **extra)

    # 16-bit grayscale
    f = 'test-tiff-16bit.tif'
    base = tiff_page(1, 1, 'TIF-07  16-bit grayscale', [R.add('TIF-07', f, 'Whole page', '16-bit grayscale (medical / archive scans)', 'Opens; gray image is readable', '⚠️')], 'L')
    base.convert('I;16').save(os.path.join(OUT, f))
    R.save()

if __name__ == '__main__':
    main()
