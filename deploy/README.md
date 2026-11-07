---
title: Pengenal Huruf Bahasa Isyarat ASL
emoji: 🤟
colorFrom: blue
colorTo: indigo
sdk: gradio
sdk_version: 5.9.1
python_version: "3.10"
app_file: app.py
pinned: false
---

# Pengenal Huruf Bahasa Isyarat ASL (Sign Language MNIST)

Aplikasi demo proyek mata kuliah IF2121 Pembelajaran Mendalam, Universitas Mikroskil.
Model CNN dilatih pada dataset Sign Language MNIST (24 huruf statis), lalu dikompresi ke TensorFlow Lite.

- **Fitur unik 1**: ambang confidence τ. Prediksi dengan confidence < τ ditolak sebagai "tidak yakin / di luar domain".
- **Fitur unik 2**: model TensorFlow Lite terkompresi yang dijalankan dengan runtime LiteRT, tanpa TensorFlow penuh.

Cara memakai: unggah foto tangan (latar polos, tangan di tengah), atau pilih salah satu contoh, lalu klik **Prediksi**.
