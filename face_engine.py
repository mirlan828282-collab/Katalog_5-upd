import sys
from pathlib import Path
try:
    import cv2
    import numpy as np
except Exception:
    cv2=None; np=None

def resource_dir():
    if getattr(sys,'frozen',False): return Path(getattr(sys,'_MEIPASS',Path(sys.executable).parent))
    return Path(__file__).resolve().parent
MODEL_DIR=resource_dir()/'models'
DETECTOR_MODEL=MODEL_DIR/'face_detection_yunet_2023mar.onnx'
RECOGNIZER_MODEL=MODEL_DIR/'face_recognition_sface_2021dec.onnx'

class FaceEngine:
    def __init__(self):
        self.available=False;self.error='';self.detector=None;self.recognizer=None
        if cv2 is None or np is None:self.error='OpenCV не установлен.';return
        if not DETECTOR_MODEL.exists() or not RECOGNIZER_MODEL.exists():self.error='Модели YuNet/SFace не найдены в папке models.';return
        try:
            self.detector=cv2.FaceDetectorYN_create(str(DETECTOR_MODEL),'',(320,320),0.60,0.30,5000)
            self.recognizer=cv2.FaceRecognizerSF_create(str(RECOGNIZER_MODEL),'');self.available=True
        except Exception as exc:self.error=str(exc)
    def _load(self,image_path):
        # np.fromfile + imdecode reliably opens paths with Cyrillic names on Windows.
        data=np.fromfile(str(image_path),dtype=np.uint8);img=cv2.imdecode(data,cv2.IMREAD_COLOR)
        if img is None:raise ValueError('Не удалось открыть изображение.')
        return img
    def embeddings(self,image_path):
        if not self.available:raise RuntimeError(self.error or 'Модуль поиска по лицу недоступен.')
        img=self._load(image_path);h,w=img.shape[:2];self.detector.setInputSize((w,h));_,faces=self.detector.detect(img)
        if faces is None or len(faces)==0:raise ValueError('Лицо на фотографии не найдено.')
        faces=sorted(faces,key=lambda f:float(f[2]*f[3]),reverse=True);out=[]
        for face in faces:
            try:
                aligned=self.recognizer.alignCrop(img,face);feat=self.recognizer.feature(aligned).flatten().astype('float32');n=float(np.linalg.norm(feat));
                if n>0:feat/=n
                out.append(feat)
            except Exception:pass
        if not out:raise ValueError('Не удалось извлечь признаки лица.')
        return out
    def embedding(self,image_path):return self.embeddings(image_path)[0]
    @staticmethod
    def cosine(a,b):
        if np is None:return 0.0
        a=np.asarray(a,dtype='float32');b=np.asarray(b,dtype='float32');den=float(np.linalg.norm(a)*np.linalg.norm(b));return float(np.dot(a,b)/den) if den else 0.0
