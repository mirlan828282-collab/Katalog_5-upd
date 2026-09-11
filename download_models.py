from pathlib import Path
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent

MODELS = {
    'face_detection_yunet_2023mar.onnx': 'https://github.com/opencv/opencv_zoo/raw/main/models/face_detection_yunet/face_detection_yunet_2023mar.onnx',
    'face_recognition_sface_2021dec.onnx': 'https://github.com/opencv/opencv_zoo/raw/main/models/face_recognition_sface/face_recognition_sface_2021dec.onnx',
}

ASSETS = {
    'flag_kg.png': 'https://thumb.wikimedia.org/wikipedia/commons/thumb/c/c7/Flag_of_Kyrgyzstan.svg/960px-Flag_of_Kyrgyzstan.svg.png',
    'emblem_kg.png': 'https://thumb.wikimedia.org/wikipedia/commons/thumb/f/f1/Emblem_of_Kyrgyzstan.svg/960px-Emblem_of_Kyrgyzstan.svg.png',
}


def download(url, target, min_size=1000):
    if target.exists() and target.stat().st_size >= min_size:
        print('OK', target.name)
        return
    print('Downloading', target.name)
    req = Request(url, headers={'User-Agent': 'PhotoArchiveCatalog/1.0'})
    with urlopen(req, timeout=120) as r, open(target, 'wb') as f:
        while True:
            chunk = r.read(1024 * 1024)
            if not chunk:
                break
            f.write(chunk)
    if target.stat().st_size < min_size:
        raise RuntimeError(f'Downloaded file is too small: {target}')
    print('Saved', target, target.stat().st_size)

models_dir = ROOT / 'models'
models_dir.mkdir(exist_ok=True)
for name, url in MODELS.items():
    download(url, models_dir / name, 10000)

assets_dir = ROOT / 'assets'
assets_dir.mkdir(exist_ok=True)
for name, url in ASSETS.items():
    download(url, assets_dir / name, 10000)
