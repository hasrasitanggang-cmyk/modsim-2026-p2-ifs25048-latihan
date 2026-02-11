import pandas as pd
import sys

# ==============================
# BACA DATA
# ==============================
try:
    df = pd.read_excel("data_kuesioner.xlsx")
except Exception:
    try:
        df = pd.read_csv("data_kuesioner.csv")
    except Exception:
        print("File data_kuesioner.xlsx atau data_kuesioner.csv tidak ditemukan atau tidak bisa dibaca.")
        sys.exit(1)

pertanyaan = list(df.columns[1:])   # Q1 – Q17

# ==============================
# HITUNG SEMUA JAWABAN
# ==============================
all_answers = df[pertanyaan].values.flatten()
all_answers = all_answers[pd.notna(all_answers)]
total_respon = len(all_answers)

counts = pd.Series(all_answers).value_counts()
percent = (counts / total_respon * 100) if total_respon > 0 else counts * 0

skor_map = {
    "SS": 6,
    "S": 5,
    "CS": 4,
    "CTS": 3,
    "TS": 2,
    "STS": 1
}

# ==============================
# INPUT
# ==============================
if len(sys.argv) > 1:
    target_question = sys.argv[1].strip().lower()
else:
    target_question = input().strip().lower()

# ==============================
# Q1
# ==============================
if target_question == "q1":
    if not counts.empty:
        skala = counts.idxmax()
        print(f"{skala}|{int(counts[skala])}|{percent[skala]:.1f}")
    else:
        print("Tidak ada data untuk Q1")

# ==============================
# Q2
# ==============================
elif target_question == "q2":
    if not counts.empty:
        skala = counts.idxmin()
        print(f"{skala}|{int(counts[skala])}|{percent[skala]:.1f}")
    else:
        print("Tidak ada data untuk Q2")

# ==============================
# Q3–Q8
# ==============================
elif target_question in ["q3","q4","q5","q6","q7","q8"]:
    skala_map = {
        "q3": "SS",
        "q4": "S",
        "q5": "CS",
        "q6": "CTS",
        "q7": "TS",
        "q8": "STS"
    }

    skala = skala_map[target_question]
    hasil = {}

    for q in pertanyaan:
        hasil[q] = (df[q] == skala).sum()

    if hasil:
        q_max = max(hasil, key=hasil.get)
        count_max = hasil[q_max]
        persen = (count_max / len(df) * 100) if len(df) > 0 else 0
        print(f"{q_max}|{int(count_max)}|{persen:.1f}")
    else:
        print("Tidak ada data untuk pertanyaan ini")

# ==============================
# Q9
# ==============================
elif target_question == "q9":
    output = []
    for q in pertanyaan:
        jumlah = (df[q] == "STS").sum()
        total_valid = len(df[q].dropna())
        if jumlah > 0 and total_valid > 0:
            persen = jumlah / total_valid * 100
            output.append(f"{q}:{persen:.1f}")
    if output:
        print("|".join(output))
    else:
        print("Tidak ada responden dengan jawaban STS")

# ==============================
# Q10 (2 angka koma)
# ==============================
elif target_question == "q10":
    total_skor = 0
    for skala in skor_map:
        total_skor += counts.get(skala, 0) * skor_map[skala]

    rata2 = (total_skor / total_respon) if total_respon > 0 else 0
    print(f"{rata2:.2f}")

# ==============================
# Q11 (2 angka koma)
# ==============================
elif target_question == "q11":
    skor_q = {}
    for q in pertanyaan:
        skor_values = df[q].map(skor_map).dropna()
        if not skor_values.empty:
            skor_q[q] = skor_values.mean()

    if skor_q:
        q_target = max(skor_q, key=skor_q.get)
        print(f"{q_target}:{skor_q[q_target]:.2f}")
    else:
        print("Tidak ada data untuk pertanyaan ini")

# ==============================
# Q12 (2 angka koma)
# ==============================
elif target_question == "q12":
    skor_q = {}
    for q in pertanyaan:
        skor_values = df[q].map(skor_map).dropna()
        if not skor_values.empty:
            skor_q[q] = skor_values.mean()

    if skor_q:
        q_target = min(skor_q, key=skor_q.get)
        print(f"{q_target}:{skor_q[q_target]:.2f}")
    else:
        print("Tidak ada data untuk pertanyaan ini")

# ==============================
# Q13 (1 angka koma)
# ==============================
elif target_question == "q13":
    positif = int(counts.get("SS", 0)) + int(counts.get("S", 0))
    netral = int(counts.get("CS", 0))
    negatif = int(counts.get("CTS", 0)) + int(counts.get("TS", 0)) + int(counts.get("STS", 0))

    if total_respon > 0:
        positif_pct = positif / total_respon * 100
        netral_pct = netral / total_respon * 100
        negatif_pct = negatif / total_respon * 100
    else:
        positif_pct = netral_pct = negatif_pct = 0

    print(
        f"positif={positif}:{positif_pct:.1f}|"
        f"netral={netral}:{netral_pct:.1f}|"
        f"negatif={negatif}:{negatif_pct:.1f}"
    )

else:
    print("Input tidak valid. Gunakan q1 sampai q13.")