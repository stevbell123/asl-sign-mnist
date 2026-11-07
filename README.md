# Klasifikasi Huruf Bahasa Isyarat Amerika (ASL) pada Dataset Sign Language MNIST Menggunakan Deep Neural Network dan Convolutional Neural Network

Proyek **IF2121 – Pembelajaran Mendalam**, Universitas Mikroskil, Semester Ganjil TA 2026-2027. Kelompok **Kelompok1** (IF-A-PAGI).

Model DNN (MLP) dan CNN untuk mengklasifikasikan 24 huruf statis American Sign Language dari citra 28×28 (dataset [Sign Language MNIST](https://www.kaggle.com/datasets/datamunge/sign-language-mnist)).

- Aplikasi: (isi link Hugging Face Space)
- Video demo: (isi link video)

## Anggota dan peran
| NIM | Nama | Peran |
|---|---|---|
| 231111408 | Steven David H Malau | Ketua Tim & Model Engineer (CNN) |
| 231111488 | Wan Saka Nasa | Data Engineer (EDA, Preprocessing, tf.data) |
| 231111207 | Edwin Valentino Tarigan | DNN & Optimasi (Perceptron, MLP, Optimizer) |
| 231110581 | Eric Chandra | Evaluasi & Fitur Unik (Threshold/OOD, TFLite, Grad-CAM) |
| 231111650 | Dhaifan Al Ghani | Deployment & Dokumentasi (Gradio, HF Spaces, GitHub, Video) |

## Hasil (data uji resmi, 7.172 citra)
| ID | Model | Parameter | Akurasi | Macro-F1 |
|---|---|---|---|---|
| E1 | MLP Baseline | 236,952 | 68.00% | 65.68% |
| E2 | MLP + BatchNorm + Dropout + Adam | 571,928 | 81.40% | 80.03% |
| E3 | CNN Baseline | 423,448 | 91.58% | 90.53% |
| E4 | CNN + BatchNorm + Dropout | 441,912 | 99.75% | 99.72% |
| E5 | CNN + BatchNorm + Dropout + Augmentasi (model akhir) | 441,912 | 100.00% | 100.00% |

**Fitur unik 1 (ambang confidence τ = 0.50)**: akurasi prediksi yang diterima 100.00% dengan coverage 100.00%.

**Fitur unik 2 (TFLite Dynamic range INT8)**: 446 KB (Keras 5274 KB), akurasi 100.00%.

## Cara menjalankan
### Training dan evaluasi (Google Colab)
1. Buka notebook `.ipynb` di Google Colab, lalu pilih Runtime → T4 GPU.
2. Isi sel KONFIGURASI, lalu Runtime → Run all.

### Aplikasi secara lokal
```bash
cd deploy
pip install -r requirements.txt gradio
python app.py
```

## Struktur
```
├── <notebook>.ipynb      # source code lengkap
├── deploy/                # aplikasi Gradio (HF Spaces)
├── hasil/                 # grafik, results.json, ringkasan eksperimen
└── requirements.txt
```
