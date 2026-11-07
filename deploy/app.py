# app.py: antarmuka Gradio untuk Hugging Face Spaces
import glob
import os

import gradio as gr

from inference import HERE, SignClassifier

clf = SignClassifier()
CONTOH = sorted(glob.glob(os.path.join(HERE, "examples", "*.png")))


def prediksi(image, threshold):
    if image is None:
        return {}, "Silakan unggah foto atau ambil gambar tangan terlebih dahulu.", None
    r = clf.predict(image, threshold)
    if r["accepted"]:
        status = (f"### ✅ Huruf terdeteksi: **{r['label']}**\n"
                  f"Confidence {r['confidence']*100:.1f}% ≥ ambang {r['threshold']*100:.0f}%.")
    else:
        status = (f"### ⚠️ Tidak yakin, prediksi ditolak\n"
                  f"Confidence {r['confidence']*100:.1f}% < ambang {r['threshold']*100:.0f}%. "
                  f"Input kemungkinan bukan isyarat huruf ASL statis atau kualitas fotonya kurang baik "
                  f"(tebakan teratas: {r['label']}).")
    return r["probs"], status, r["input28"]


with gr.Blocks(title="Pengenal Huruf Bahasa Isyarat ASL") as demo:
    gr.Markdown(
        "# 🤟 Pengenal Huruf Bahasa Isyarat ASL (Sign Language MNIST)\n"
        f"CNN 3 blok konvolusi, dikompresi ke TensorFlow Lite ({clf.cfg['variant']}, {clf.cfg['size_kb']} KB, "
        f"akurasi uji {clf.cfg['test_acc']*100:.2f}%). Mengenali 24 huruf statis (A–Y tanpa J dan Z).\n\n"
        "**Fitur unik:** (1) ambang confidence untuk menolak input yang meragukan atau di luar domain; "
        "(2) model TFLite terkompresi yang ringan.")
    with gr.Row():
        with gr.Column():
            inp = gr.Image(type="numpy", label="Citra tangan", sources=["upload", "webcam", "clipboard"])
            thr = gr.Slider(0.0, 1.0, value=clf.threshold, step=0.01, label="Ambang confidence τ")
            btn = gr.Button("Prediksi", variant="primary")
        with gr.Column():
            out_label = gr.Label(num_top_classes=3, label="Top-3 prediksi")
            out_status = gr.Markdown()
            out_prev = gr.Image(label="Input setelah praproses (28×28)", image_mode="L")
    if CONTOH:
        gr.Examples(examples=[[p, clf.threshold] for p in CONTOH], inputs=[inp, thr])
    btn.click(prediksi, inputs=[inp, thr], outputs=[out_label, out_status, out_prev])

if __name__ == "__main__":
    demo.launch()
