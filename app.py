import streamlit as st
import pandas as pd
import requests
import time
from PIL import Image
import matplotlib.pyplot as plt

st.set_page_config(
      page_title="PO Anyar Group",
      page_icon=Image.open("image.png"),
      layout="centered"
    )

st.title("📦 Purchase Order Anyar Group Automation")

# === STEP 1: FORM INPUT USER ===
st.header("1. Upload File Excel")
uploaded_file = st.file_uploader("📂 Upload file Excel yang berisi 'No. Dokumen'", type=["xlsx"])

st.header("2. Masukkan Token")
csrf_token = st.text_input("🔐 CSRF Token", type="password")
session_id = st.text_input("🔑 Session ID", type="password")

submit_btn = st.button("🚀 Mulai Ambil Data")

if submit_btn and uploaded_file and csrf_token and session_id:
    try:
        df_input = pd.read_excel(uploaded_file)

        if "No. Dokumen" not in df_input.columns:
            st.error("❌ Kolom 'No. Dokumen' tidak ditemukan pada file Excel.")
        else:
            dokumen_list = df_input["No. Dokumen"].astype(str).tolist()

            # === STEP 2: HEADER DENGAN COOKIE DARI BROWSER ===
            headers = {
                "Accept": "*/*",
                "Accept-Encoding": "gzip, deflate, br, zstd",
                "Accept-Language": "en-US,en;q=0.9",
                "Connection": "keep-alive",
                "Cookie": f"csrftoken={csrf_token}; sessionid={session_id}",
                "Host": "vis.anyargroup.co.id",
                "Referer": "https://vis.anyargroup.co.id/purchase_order",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "User-Agent": "Mozilla/5.0",
                "X-Requested-With": "XMLHttpRequest",
                "sec-ch-ua": '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
            }

            results = []
            progress = st.progress(0, text="Sedang mengambil data...")

            for idx, no_dokumen in enumerate(dokumen_list):
                url = f"https://vis.anyargroup.co.id/detailpo?u=VL0000446&d={no_dokumen}&s=O&o=22"
                try:
                    response = requests.get(url, headers=headers)
                    response.raise_for_status()

                    try:
                        data = response.json()
                    except ValueError:
                        st.warning(f"[!] Gagal decode JSON dari {no_dokumen}. Cek login/token.")
                        continue

                    if "data" in data:
                        for dc_name, items in data["data"].items():
                            for item in items:
                                result = {
                                    "Tanggal PO": item[6],
                                    "No PO": no_dokumen,
                                    "Distribution Center": item[0],
                                    "Product Name": item[18],
                                    "Qty": item[19]
                                }
                                results.append(result)
                    else:
                        st.warning(f"[!] Tidak ada data pada dokumen {no_dokumen}")

                    time.sleep(0.5)
                    progress.progress((idx + 1) / len(dokumen_list))

                except Exception as e:
                    st.error(f"[!] Error saat ambil dokumen {no_dokumen}: {e}")

            # === STEP 4: TAMPILKAN DAN DOWNLOAD HASIL ===
            if results:
                df_result = pd.DataFrame(results)
                st.success("✅ Data berhasil diambil.")
                st.dataframe(df_result)

                # === VISUALISASI DENGAN MATPLOTLIB ===
                st.subheader("📊 Visualisasi Data PO")

                # Total Qty per Product Name (Top 10)
                st.markdown("#### Top 10 Product by Total Qty")
                top_products = df_result.groupby("Product Name")["Qty"].sum().sort_values(ascending=False).head(10)
                fig1, ax1 = plt.subplots()
                top_products.plot(kind='barh', color='skyblue', ax=ax1)
                ax1.set_xlabel("Total Qty")
                ax1.set_ylabel("Product Name")
                ax1.set_title("Top 10 Product by Qty")
                ax1.invert_yaxis()
                st.pyplot(fig1)

                # Jumlah PO per Distribution Center
                st.markdown("#### Jumlah PO per Distribution Center")
                po_per_dc = df_result.groupby("Distribution Center")["No PO"].nunique().sort_values(ascending=False)
                fig2, ax2 = plt.subplots()
                po_per_dc.plot(kind='bar', color='lightgreen', ax=ax2)
                ax2.set_ylabel("Jumlah PO")
                ax2.set_title("PO Count per DC")
                st.pyplot(fig2)
            else:
                st.warning("⚠️ Tidak ada data berhasil diambil.")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")

elif submit_btn:
    st.warning("⚠️ Mohon upload file dan isi CSRF Token serta session ID terlebih dahulu.")
