# inference.py: praproses + prediksi dengan model TFLite (LiteRT). Dipakai oleh app.py dan notebook.
import json
import os

import numpy as np
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
_RESAMPLE = getattr(Image, "Resampling", Image)   # kompatibel Pillow lama dan baru


def _interpreter_cls():
    try:
        from ai_edge_litert.interpreter import Interpreter  # runtime ringan (tanpa TensorFlow penuh)
        return Interpreter
    except ImportError:
        import tensorflow as tf
        return tf.lite.Interpreter


class SignClassifier:
    def __init__(self, config_path=None):
        with open(config_path or os.path.join(HERE, "config.json")) as f:
            self.cfg = json.load(f)
        self.labels = self.cfg["labels"]
        self.threshold = float(self.cfg["threshold"])
        self.itp = _interpreter_cls()(model_path=os.path.join(HERE, self.cfg["model_file"]))
        self.itp.allocate_tensors()
        self.inp = self.itp.get_input_details()[0]
        self.out = self.itp.get_output_details()[0]

    @staticmethod
    def preprocess(image):
        # RGB/RGBA/grayscale (numpy atau PIL) -> grayscale -> crop tengah persegi -> 28x28 -> [0,1]
        img = image if isinstance(image, Image.Image) else Image.fromarray(np.asarray(image).astype(np.uint8))
        img = img.convert("L")
        w, h = img.size
        s = min(w, h)
        img = img.crop(((w - s) // 2, (h - s) // 2, (w - s) // 2 + s, (h - s) // 2 + s))
        if img.size != (28, 28):
            img = img.resize((28, 28), _RESAMPLE.BOX)
        x = np.asarray(img, dtype=np.float32) / 255.0
        return x.reshape(1, 28, 28, 1)

    def predict(self, image, threshold=None):
        x = self.preprocess(image)
        self.itp.set_tensor(self.inp["index"], x)
        self.itp.invoke()
        p = self.itp.get_tensor(self.out["index"])[0]
        tau = self.threshold if threshold is None else float(threshold)
        top = np.argsort(-p)[:3]
        return {"label": self.labels[int(top[0])], "confidence": float(p[top[0]]),
                "accepted": bool(p[top[0]] >= tau), "threshold": tau,
                "top3": [(self.labels[int(i)], float(p[i])) for i in top],
                "probs": {self.labels[i]: float(p[i]) for i in range(len(p))},
                "input28": (x[0, ..., 0] * 255).astype(np.uint8)}
