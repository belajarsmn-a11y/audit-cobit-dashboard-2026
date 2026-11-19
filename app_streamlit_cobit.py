import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="Audit COBIT 2019", page_icon="📊", layout="wide")

st.title("Dashboard Analisis Audit IT Governance COBIT 2019")
st.caption("Studi kasus simulasi — Domain EDM dan APO")

st.info(
    "Unggah file Dataset_Audit_COBIT2019_EDM_APO.xlsx. "
    "Pastikan file memiliki sheet 'Dataset_Audit' dan kolom sesuai kebutuhan."
)

uploaded_file = st.file_uploader("Pilih dataset Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Dataset_Audit")

        required = [
            "Kode Proses", "Nama Proses", "Domain",
            "PA1.1", "PA2.1", "PA2.2", "PA3.1", "PA3.2", "Target"
        ]
        missing = [col for col in required if col not in df.columns]
        if missing:
            st.error(f"Kolom wajib tidak ditemukan: {missing}")
            st.stop()

        rating_rank = {"N": 0, "P": 1, "L": 2, "F": 3}
        rating_cols = ["PA1.1", "PA2.1", "PA2.2", "PA3.1", "PA3.2"]

        for col in rating_cols:
            df[col] = df[col].astype(str).str.strip().str.upper()
            invalid = sorted(set(df[col]) - set(rating_rank))
            if invalid:
                st.error(f"Rating tidak valid pada kolom {col}: {invalid}")
                st.stop()

        df["Target"] = pd.to_numeric(df["Target"], errors="coerce")
        if df["Target"].isna().any():
            st.error("Kolom Target harus berisi angka capability level.")
            st.stop()

        def hitung_capability_level(row):
            level = 0
            if rating_rank[row["PA1.1"]] >= rating_rank["L"]:
                level = 1
            if (
                row["PA1.1"] == "F"
                and rating_rank[row["PA2.1"]] >= rating_rank["L"]
                and rating_rank[row["PA2.2"]] >= rating_rank["L"]
            ):
                level = 2
            if (
                row["PA1.1"] == "F"
                and row["PA2.1"] == "F"
                and row["PA2.2"] == "F"
                and rating_rank[row["PA3.1"]] >= rating_rank["L"]
                and rating_rank[row["PA3.2"]] >= rating_rank["L"]
            ):
                level = 3
            return level

        df["As-Is"] = df.apply(hitung_capability_level, axis=1)
        df["Gap"] = df["Target"] - df["As-Is"]

        if (df["Gap"] < 0).any():
            st.error("Ada Target yang lebih rendah daripada As-Is. Periksa dataset.")
            st.stop()

        df["Prioritas"] = df["Gap"].apply(
            lambda gap: "Tinggi" if gap >= 2
            else ("Sedang" if gap == 1 else "Pemeliharaan")
        )

        st.success("Dataset berhasil dibaca dan dianalisis.")

        st.subheader("Ringkasan hasil audit")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Jumlah proses", len(df))
        c2.metric("Rata-rata As-Is", f'{df["As-Is"].mean():.2f}')
        c3.metric("Rata-rata Gap", f'{df["Gap"].mean():.2f}')
        c4.metric("Gap tertinggi", int(df["Gap"].max()))

        st.subheader("Hasil analisis per proses")
        st.dataframe(
            df[["Kode Proses", "Nama Proses", "Domain", "As-Is", "Target", "Gap", "Prioritas"]],
            use_container_width=True,
            hide_index=True
        )

        st.subheader("Prioritas perbaikan")
        prioritas = df.sort_values(
            by=["Gap", "Kode Proses"], ascending=[False, True]
        )
        st.dataframe(
            prioritas[["Kode Proses", "Nama Proses", "Gap", "Prioritas"]],
            use_container_width=True,
            hide_index=True
        )

        st.subheader("Ringkasan per domain")
        ringkasan_domain = df.groupby("Domain").agg(
            Jumlah_Proses=("Kode Proses", "count"),
            Rata_As_Is=("As-Is", "mean"),
            Rata_Gap=("Gap", "mean")
        ).round(2)
        st.dataframe(ringkasan_domain, use_container_width=True)

        st.subheader("Grafik perbandingan As-Is dan Target")
        fig1, ax1 = plt.subplots(figsize=(10, 5))
        df.plot(x="Kode Proses", y=["As-Is", "Target"], kind="bar", ax=ax1)
        ax1.set_title("Analisis Gap Capability Level COBIT 2019")
        ax1.set_xlabel("Proses pada Domain EDM dan APO")
        ax1.set_ylabel("Capability Level")
        ax1.set_ylim(0, 5)
        ax1.legend(["As-Is", "Target / To-Be"])
        plt.xticks(rotation=0)
        fig1.tight_layout()
        st.pyplot(fig1)
        plt.close(fig1)

        st.subheader("Grafik Gap setiap proses")
        fig2, ax2 = plt.subplots(figsize=(10, 4))
        df.plot(x="Kode Proses", y="Gap", kind="bar", legend=False, ax=ax2)
        ax2.set_title("Besarnya Gap per Proses COBIT 2019")
        ax2.set_xlabel("Proses")
        ax2.set_ylabel("Gap (tingkat)")
        plt.xticks(rotation=0)
        fig2.tight_layout()
        st.pyplot(fig2)
        plt.close(fig2)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Hasil_Per_Proses", index=False)
            ringkasan_domain.to_excel(writer, sheet_name="Ringkasan_Domain")
            prioritas.to_excel(writer, sheet_name="Prioritas_Perbaikan", index=False)

        st.download_button(
            label="Download Hasil Analisis Excel",
            data=output.getvalue(),
            file_name="Hasil_Analisis_Audit_COBIT2019.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except ValueError as e:
        st.error(f"Terjadi kesalahan membaca dataset: {e}")
    except Exception as e:
        st.error(f"Program mengalami kendala: {e}")
else:
    st.warning("Silakan unggah dataset Excel untuk memulai analisis.")
