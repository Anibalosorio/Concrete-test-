import urllib.request
import re

url = 'https://ctp-llp.com/'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
try:
    response = urllib.request.urlopen(req, timeout=10)
    html = response.read().decode('utf-8')

    # Look for logo images
    logo_patterns = re.findall(r'(https?://[^\"\'>]+(?:logo|Logo|LOGO)[^\"\'>]*\.(?:png|jpg|jpeg|svg|webp))', html)

    # Look for images in header/nav
    img_patterns = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', html)

    print('Logo URLs found:')
    for logo in logo_patterns[:5]:
        print(f'  {logo}')

    print('\nFirst image URLs:')
    for img in img_patterns[:10]:
        print(f'  {img}')

except Exception as e:
    print(f'Error: {e}')
